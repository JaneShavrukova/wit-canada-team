#!/usr/bin/env python3
# ============================================================
# WIT Canada — Signature contact-icon generator
# ============================================================
# Regenerates assets/icon-mail.png / icon-phone.png / icon-link.png,
# the three glyphs used in the email signature's contact rows
# (SignatureGenerator.html). Run only when the icons need to change:
#
#     python3 tools/make_signature_icons.py
#
# Repo-only — `.claspignore` does not need to exclude it (clasp only
# pushes .js/.html), and it never runs on Apps Script.
#
# NAVY below must stay in sync with THEME.brand.electric
# (core/Theme.js), which is the source of truth for the color.
#
# Pure stdlib: shapes are signed-distance fields sampled 4x4 per pixel
# for anti-aliasing, then written as RGBA PNG via zlib. No Pillow, no
# SVG rasterizer (macOS has neither reliably).
# ============================================================

import math
import struct
import zlib

NAVY = (0x02, 0x0b, 0xa6)  # == THEME.brand.electric
SIZE = 128                 # drawn at 128px, displayed at 18px
SS = 4                     # supersampling factor per axis


# ── signed distance helpers (negative = inside) ──────────────

def sd_round_rect(px, py, cx, cy, hw, hh, r):
    dx = abs(px - cx) - (hw - r)
    dy = abs(py - cy) - (hh - r)
    outside = math.hypot(max(dx, 0.0), max(dy, 0.0))
    inside = min(max(dx, dy), 0.0)
    return outside + inside - r


def sd_circle(px, py, cx, cy, r):
    return math.hypot(px - cx, py - cy) - r


def sd_segment(px, py, ax, ay, bx, by):
    vx, vy = bx - ax, by - ay
    wx, wy = px - ax, py - ay
    denom = vx * vx + vy * vy
    t = 0.0 if denom == 0 else max(0.0, min(1.0, (wx * vx + wy * vy) / denom))
    return math.hypot(wx - t * vx, wy - t * vy)


def ring(d, stroke):
    """Outline of a shape whose SDF is `d`, centred on the path."""
    return abs(d) - stroke / 2.0


def rotate(px, py, cx, cy, deg):
    a = math.radians(deg)
    ca, sa = math.cos(a), math.sin(a)
    dx, dy = px - cx, py - cy
    return cx + dx * ca - dy * sa, cy + dx * sa + dy * ca


# ── icon definitions: each returns True if (x, y) is inked ───

def mail(x, y):
    # envelope body: rounded-rect outline
    if ring(sd_round_rect(x, y, 64, 64, 56, 41, 7), 12) <= 0:
        return True
    # flap: a V from the inner top corners down to the apex
    if min(sd_segment(x, y, 19, 27, 64, 73),
           sd_segment(x, y, 109, 27, 64, 73)) - 6 <= 0:
        return True
    return False


def phone(x, y):
    # handset outline
    if ring(sd_round_rect(x, y, 64, 64, 28, 56, 9), 11) <= 0:
        return True
    # filled screen
    if sd_round_rect(x, y, 64, 53, 15, 32, 2) <= 0:
        return True
    # home button
    if sd_circle(x, y, 64, 101, 7.5) <= 0:
        return True
    return False


def link(x, y):
    # two interlocking stadium rings, the pair tilted to the diagonal
    lx, ly = rotate(x, y, 64, 64, 45)
    # Centres are spaced so the strokes cross but each ring's hole stays
    # open — too much overlap and the pair reads as one blob.
    for cy in (40, 88):
        if ring(sd_round_rect(lx, ly, 64, cy, 20, 30, 20), 11) <= 0:
            return True
    return False


# ── raster + PNG encode ──────────────────────────────────────

def render(fn):
    px = bytearray(SIZE * SIZE * 4)
    step = 1.0 / SS
    offset = step / 2.0
    for y in range(SIZE):
        for x in range(SIZE):
            hits = 0
            for sy in range(SS):
                py = y + offset + sy * step
                for sx in range(SS):
                    if fn(x + offset + sx * step, py):
                        hits += 1
            if hits:
                i = (y * SIZE + x) * 4
                px[i:i + 3] = bytes(NAVY)
                px[i + 3] = round(255 * hits / (SS * SS))
    return px


def write_png(path, px):
    raw = b''.join(
        b'\x00' + bytes(px[y * SIZE * 4:(y + 1) * SIZE * 4]) for y in range(SIZE)
    )

    def chunk(tag, data):
        return (struct.pack('>I', len(data)) + tag + data
                + struct.pack('>I', zlib.crc32(tag + data) & 0xffffffff))

    header = struct.pack('>IIBBBBB', SIZE, SIZE, 8, 6, 0, 0, 0)
    with open(path, 'wb') as f:
        f.write(b'\x89PNG\r\n\x1a\n'
                + chunk(b'IHDR', header)
                + chunk(b'IDAT', zlib.compress(raw, 9))
                + chunk(b'IEND', b''))
    print(f'wrote {path}')


if __name__ == '__main__':
    import os
    out = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'assets')
    for name, fn in (('icon-mail', mail), ('icon-phone', phone), ('icon-link', link)):
        write_png(os.path.join(out, f'{name}.png'), render(fn))
