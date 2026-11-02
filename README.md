# AgentGQ Bridge Test v1.1

This is a static, **client-side-only** Outlook add-in diagnostic. It can read only the currently opened message when the user clicks a button. It does not send email data anywhere, does not read the whole inbox, and does not run in the background. The Office.js library is loaded from Microsoft's official CDN.

## Publish without owning a server

1. Create a **public** GitHub repository named `agentgq-bridge-test` (public because GitHub Pages on some plans requires it). Do not include private email, credentials, internal URLs, or secrets.
2. Upload `index.html`, `configure.html`, `taskpane.html`, `taskpane.js`, `icon.png` to the repository root. Do not upload any emails or API tokens.
3. In repository Settings > Pages, select **Deploy from a branch**, branch `main`, folder `/ (root)`. Wait until GitHub shows the public HTTPS URL.
4. Open `<your-pages-url>/configure.html`, paste the URL, and download `manifest.xml`.
5. In Outlook Web, uninstall the earlier broken AgentGQ Bridge Test (if needed), then choose **Add from file** and select the downloaded `manifest.xml`.
6. Open a **non-sensitive test email**, click AgentGQ Bridge Test, and click **Прочитать открытое письмо**.

If Outlook reports an error, check whether the corporate network can reach both GitHub Pages and `https://appsforoffice.microsoft.com/lib/1/hosted/office.js`, and whether third-party add-ins are permitted.

**Important:** Publishing the code on GitHub Pages makes the code publicly accessible. Email content remains in the Outlook add-in's browser runtime in this test, but any future change that transmits data must be separately reviewed and authorized. Do not deploy this prototype to production or use it with sensitive corporate correspondence without approval.

This test **does not** provide ChatGPT scheduled automations with mailbox access. That requires a separate approved backend and supported authentication/integration.
