=== Zimunim Extension - Policy Deployment ===

Extension ID: dhlleojgaiclmkkipcooihgffajlochd
Version     : 1.0.2
Install path: C:\ProgramData\Zimunim

----- Easiest install (recommended) -----
1. Right-click install.ps1 -> Run with PowerShell as Administrator.
   (If blocked: open elevated PowerShell, then:
    Set-ExecutionPolicy -Scope Process Bypass; .\install.ps1)
2. Restart Chrome / Brave / Edge.
3. Verify in chrome://extensions - the extension appears with
   "Installed by enterprise policy".

----- Manual install with REG files -----
1. Copy zimunim.crx and update.xml to: C:\ProgramData\Zimunim
2. Double-click install-all.reg (or install-chrome.reg / brave / edge for a
   single browser) and confirm.
3. Restart the browser.

If the CRX or update.xml are placed somewhere other than C:\ProgramData\Zimunim,
edit the file:/// paths inside update.xml and the .reg file accordingly.

----- Uninstall -----
- Run uninstall.ps1 as Administrator, OR
- Double-click uninstall-all.reg.

----- Files in this folder -----
  install.ps1           - smart installer (admin)
  uninstall.ps1         - smart uninstaller (admin)
  update.xml            - update manifest (referenced by the policy)
  install-all.reg       - REG for Chrome + Brave + Edge
  install-chrome.reg    - REG for Chrome only
  install-brave.reg     - REG for Brave only
  install-edge.reg      - REG for Edge only
  uninstall-all.reg     - removes policy from all three
  uninstall-<browser>.reg - removes policy from a single browser

The zimunim.crx file lives one folder up (dist\zimunim.crx). Keep it
together with this folder when distributing.
