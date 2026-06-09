// ============================================================
// WIT Canada — Theme (design tokens, single source of truth)
// ============================================================
// One palette for EVERY surface in the system:
//   • Report sheets  → JS tokens via .setBackground()/.setFontColor()
//   • Sidebars / web → CSS variables injected by themeCss()
//   • HTML emails    → inline `${THEME.x}` (email clients don't
//                      support CSS variables, so values are baked in)
//
// Add or change a color HERE; never hard-code hex in feature files.
//
// Load order: this file is evaluated after core/Config.js (alphabetical
// within core/), so referencing CONFIG at top level below is safe. Do
// not reference THEME from another file's top level — only inside
// function bodies (call time).
// ============================================================

const THEME = {
  // ── Brand ────────────────────────────────────────────────
  primary:     '#1b4f8a',
  primaryDark: '#16407a', // hover / pressed
  primaryTint: '#ebf2fa', // light fill behind primary text
  onPrimary:   '#ffffff',

  // ── Text ─────────────────────────────────────────────────
  textStrong: '#2c3e50', // headings
  textBody:   '#5d6d7e', // body copy
  textMuted:  '#7f8c8d', // captions / secondary
  textFaint:  '#adb5bd', // footers / disclaimers

  // ── Surfaces & lines ─────────────────────────────────────
  surface:      '#ffffff',
  surfaceAlt:   '#f4f6fb', // page background
  surfaceMuted: '#f8f9fa', // subtle strips (e.g. report timestamp row)
  border:       '#e8ecf0',

  // ── Status / intent ──────────────────────────────────────
  success: '#0f6e56',
  warning: '#e8b84b',
  danger:  '#9c1c1c',
};

// ── Report-sheet chrome (derived from tokens) ──────────────
// Consumed by core/ReportBuilder.js and the report builders.
THEME.report = {
  TITLE_BG:    THEME.primary,
  TITLE_FG:    THEME.onPrimary,
  SUBTITLE_BG: THEME.surfaceMuted,
  SUBTITLE_FG: THEME.textMuted,
  SECTION_BG:  THEME.primaryTint,
  SECTION_FG:  THEME.primary,
};

// ── Email-Status pill / cell colors ────────────────────────
// `bg` styles sheet cell backgrounds; `bg`+`fg` style HTML badges.
// Keyed by the Email-Status enum value (see CONFIG.STATUS).
THEME.status = {
  [CONFIG.STATUS.REQUEST]:       { bg: '#fff3cd', fg: '#856404' },
  [CONFIG.STATUS.SENT]:          { bg: '#cfe2ff', fg: '#0a4a90' },
  [CONFIG.STATUS.CREATED]:       { bg: '#d1e7dd', fg: '#0a5933' },
  [CONFIG.STATUS.NOT_ACTIVATED]: { bg: '#fde8e8', fg: '#9c1c1c' },
};

/**
 * Returns a <style> block exposing the tokens as CSS custom
 * properties, for inclusion in HtmlService templates:
 *
 *   <?!= themeCss() ?>
 *   ... color: var(--color-primary);
 *
 * (Use for sidebars / web-app pages — NOT emails, which need
 * inline values from THEME directly.)
 *
 * @returns {string}
 */
function themeCss() {
  const vars = {
    'color-primary':       THEME.primary,
    'color-primary-dark':  THEME.primaryDark,
    'color-primary-tint':  THEME.primaryTint,
    'color-on-primary':    THEME.onPrimary,
    'color-text-strong':   THEME.textStrong,
    'color-text-body':     THEME.textBody,
    'color-text-muted':    THEME.textMuted,
    'color-surface':       THEME.surface,
    'color-surface-alt':   THEME.surfaceAlt,
    'color-surface-muted': THEME.surfaceMuted,
    'color-border':        THEME.border,
    'color-success':       THEME.success,
    'color-warning':       THEME.warning,
    'color-danger':        THEME.danger,
  };

  const body = Object.entries(vars)
    .map(([name, value]) => `--${name}:${value};`)
    .join('');

  return `<style>:root{${body}}</style>`;
}
