# Respice for Outlook

*Respice*, Latin for "look back". An Outlook add-in that makes you look once more before a risky email goes out.

When you press **Send**, it checks To, Cc and Bcc. If any enabled rule fires, Outlook shows a dialog with **Don't Send** / **Send Anyway**.

| # | Rule | Fires when | Settings |
|---|------|------------|----------|
| 1 | External recipient | any recipient is outside your internal domains | on/off, internal domains (default: your own address's domain, subdomains included) |
| 2 | Mailing list | any recipient matches a list pattern, or Outlook marks it as a distribution list | on/off, list addresses with `*` / `?` wildcards, DL auto-detection on/off |
| 3 | Too many recipients | unique To + Cc + Bcc recipients exceed X | on/off, X (default 15) |
| 4 | Reply to more than the sender | the message is a reply (compose type or `RE:`-style subject) and has more than one recipient | on/off |

Example dialog:

```
Check before sending:
• External recipient: john@customer.com · outside example.com
• Mailing list: all@example.com
• Reply to more than the sender: this reply goes to 7 recipients. Did you mean Reply all?
```

## Privacy

- Recipient addresses are checked locally inside Outlook. Nothing is sent anywhere.
- Settings are stored in your own mailbox (Office `roamingSettings`) and follow you across Outlook clients.
- The hosted files are static HTML/JS. No server code, no analytics, no tracking.

## Supported clients

Outlook clients that support Smart Alerts (Mailbox requirement set 1.12+):

- New Outlook for Mac
- Outlook on the web
- New Outlook for Windows
- Classic Outlook for Windows (Microsoft 365)

Not supported: legacy Outlook for Mac, Outlook mobile.

## Install (for yourself)

### 1. Host it on GitHub Pages

```bash
git init && git add . && git commit -m "Respice for Outlook 1.0.0"
git branch -M main
git remote add origin https://github.com/rokvintar/respice-for-outlook.git
git push -u origin main
```

On GitHub: **Settings → Pages → Build and deployment → Deploy from a branch → `main` / root**.
Wait until `https://rokvintar.github.io/respice-for-outlook/` loads. On a free GitHub plan the repo must be public.

### 2. Point the manifest at your Pages URL

```bash
node scripts/build-manifest.mjs https://rokvintar.github.io/respice-for-outlook
git commit -am "Manifest for GitHub Pages" && git push
```

### 3. Validate (optional, needs internet)

```bash
npm run validate
```

### 4. Sideload into Outlook

1. Open <https://aka.ms/olksideload> in a browser. This opens Outlook on the web with the add-in dialog.
2. **My add-ins → Custom Addins → Add a custom add-in → Add from File…** and select `manifest.xml`.
3. The add-in is attached to your mailbox and appears in Outlook for Mac, Windows and web within a few minutes. Restart Outlook if it doesn't.

If the "Custom Addins" option is missing, your organisation has disabled user-installed add-ins. Ask IT to deploy `manifest.xml` through the Microsoft 365 admin center (Integrated Apps).

### 5. Configure

Open any message (or a draft) → ribbon / **Apps** menu → **Respice settings**.
Fill in your mailing-list addresses, adjust X, toggle rules, then **Save**.
While composing, **Check now** runs the rules against the current draft without sending.

## Limitations

- **Reply vs. Reply all:** Outlook's add-in API only reports `newMail` / `reply` / `forward`, and at send time Outlook on Mac and on the web report `newMail` even for replies. Rule 4 therefore treats a message as a reply when Outlook says `reply` **or** the subject starts with a reply prefix (`RE:`, `AW:`, `SV:`, `Odg:`, `VS:`, `RIF:`, `Antw:`, `答复:` …), and fires when it has more than one recipient. A plain reply where you added people yourself also triggers it; a reply whose `RE:` you deleted does not.
- **Distribution lists aren't expanded.** A list counts as one recipient for rule 3, and external members hidden inside a list aren't detected by rule 1.
- **Fail-open:** if the add-in can't load (offline, GitHub Pages down), Outlook lets the mail go (`SendMode="PromptUser"`).
- The dialog text is capped at about 500 characters. Long address lists are shortened to "+N more".

## Development

No build step and no dependencies. Plain HTML/JS, served as-is.

```
src/launchevent.js   rule logic + OnMessageSend handler (self-contained: classic Windows loads it without a DOM)
src/commands.html    event runtime page for web / Mac / new Outlook
src/taskpane.*       settings UI
manifest.template.xml → scripts/build-manifest.mjs → manifest.xml
test/                Node unit tests
```

```bash
npm test
```

Bump `<Version>` in `manifest.template.xml` and rebuild the manifest whenever the manifest changes. Code-only changes go live as soon as GitHub Pages redeploys.

## Microsoft Marketplace

This is possible later. Smart Alerts add-ins qualify for Marketplace only with `PromptUser` or `SoftBlock` send mode (this add-in uses `PromptUser`). A listing additionally needs a Partner Center account, a privacy policy URL and a support URL.

## License

MIT
