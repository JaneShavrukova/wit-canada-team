// ============================================================
// WIT Canada — Sidebar
// ============================================================

function showFileGuide() {
  showSidebar('FileGuide', 'About this file', 320);
}

function showNewMemberGuideSidebar() {
  showSidebar('OnboardingGuide', 'Member onboarding guide', 280);
}

function openNewMemberGuideExternal() {
  const url = CONFIG.URLS.MEMBER_GUIDE;
  const html = HtmlService.createHtmlOutput(`<!DOCTYPE html>
<html>
<head>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: 'Google Sans', Arial, sans-serif;
    background: #f8f9fa;
    display: flex;
    align-items: flex-start;
    justify-content: center;
    padding: 16px;
  }
  .card {
    background: #fff;
    border-radius: 12px;
    box-shadow: 0 2px 12px rgba(27,79,138,0.10);
    padding: 28px 28px 24px;
    max-width: 420px;
    width: 100%;
  }
  .header {
    background: #1b4f8a;
    border-radius: 8px;
    padding: 16px 20px;
    margin-bottom: 20px;
    display: flex;
    align-items: center;
    gap: 12px;
  }
  .header .emoji { font-size: 24px; line-height: 1; }
  .header h1 {
    color: #fff;
    font-size: 16px;
    font-weight: 600;
    line-height: 1.3;
  }
  .description {
    color: #444;
    font-size: 13px;
    line-height: 1.6;
    margin-bottom: 16px;
  }
  .covers-label {
    color: #1b4f8a;
    font-size: 12px;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    margin-bottom: 8px;
  }
  ul {
    list-style: none;
    margin-bottom: 24px;
  }
  ul li {
    color: #444;
    font-size: 13px;
    padding: 4px 0 4px 20px;
    position: relative;
    line-height: 1.5;
  }
  ul li::before {
    content: '✓';
    position: absolute;
    left: 0;
    color: #1b4f8a;
    font-weight: 700;
    font-size: 12px;
  }
  .buttons {
    display: flex;
    gap: 10px;
  }
  .btn {
    flex: 1;
    padding: 10px 16px;
    border-radius: 6px;
    font-size: 13px;
    font-weight: 600;
    cursor: pointer;
    border: none;
    transition: background 0.15s, color 0.15s;
    font-family: inherit;
  }
  .btn-primary {
    background: #1b4f8a;
    color: #fff;
    text-decoration: none;
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 6px;
  }
  .btn-primary:hover { background: #16407a; }
  .btn-secondary {
    background: #ebf2fa;
    color: #1b4f8a;
  }
  .btn-secondary:hover { background: #ddeaf7; }
  .btn-secondary.copied {
    background: #e6f4ea;
    color: #1e7e34;
  }
</style>
</head>
<body>
<div class="card">
  <div class="header">
    <span class="emoji">🎉</span>
    <h1>New Member Onboarding Guide</h1>
  </div>
  <p class="description">
    A step-by-step guide for new WIT Canada members to get set up and connected with the team.
  </p>
  <p class="covers-label">What it covers</p>
  <ul>
    <li>Setting up their WIT email</li>
    <li>Joining Google Chat channels</li>
    <li>Accessing the Shared Drive</li>
    <li>Completing their profile</li>
  </ul>
  <div class="buttons">
    <a class="btn btn-primary" href="${url}" target="_blank" onclick="google.script.host.close()">
      Open Guide ↗
    </a>
    <button class="btn btn-secondary" id="copyBtn" onclick="copyLink()">
      Copy link
    </button>
  </div>
</div>
<script>
  function copyLink() {
    navigator.clipboard.writeText('${url}').then(function() {
      var btn = document.getElementById('copyBtn');
      btn.textContent = '✓ Copied!';
      btn.classList.add('copied');
      setTimeout(function() {
        btn.textContent = 'Copy link';
        btn.classList.remove('copied');
      }, 2000);
    });
  }
</script>
</body>
</html>`);
  html.setWidth(480).setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(html, 'New member guide');
}

function showSidebar(templateName, title, width) {
  const html = HtmlService
    .createTemplateFromFile(templateName)
    .evaluate()
    .setTitle(title)
    .setWidth(width);

  SpreadsheetApp.getUi().showSidebar(html);
}