# Respice for Outlook

*Respice*, Latin for "look back". An Outlook add-in that makes you look once more before a risky email goes out.

It works at two moments:

- **While you write:** when you open a new mail, reply or forward, and whenever you change To, Cc or Bcc, a one-line info bar above the message shows what Respice would warn about, e.g. `𝗘𝘅𝘁𝗲𝗿𝗻𝗮𝗹: john@customer.com · 𝗟𝗶𝘀𝘁: all@example.com · 𝗥𝗲𝗽𝗹𝘆 𝗮𝗹𝗹: 7 people` (labels use Unicode bold letters, since the bar is plain text). The bar disappears when the issue is gone.
- **When you press Send:** it checks To, Cc and Bcc again. If any enabled rule fires, Outlook shows a dialog with **Don't Send** / **Send Anyway**.

Each rule has a mode in the settings: **Bar + check at Send** (default: also a dialog when you press Send), **Bar only** (warning while writing, Send goes through), **Check at Send only** (no bar, only the dialog at Send; useful when Outlook's own External recipients MailTip already covers the bar), or **Off**.

| # | Rule | Fires when | Settings |
|---|------|------------|----------|
| 1 | External recipient | any recipient is outside your internal domains | mode, internal domains (default: your own address's domain, subdomains included) |
| 2 | Mailing list | any recipient matches a list pattern, or Outlook marks it as a distribution list | mode, list addresses with `*` / `?` wildcards, DL auto-detection on/off |
| 3 | Too many recipients | unique To + Cc + Bcc recipients exceed X | mode, X (default 15) |
| 4 | Reply to more than the sender | the message is a reply (compose type, `inReplyTo`, or `RE:`-style subject) and has more than one recipient | mode |

Example dialog:

```
Check before sending:
𝗘𝘅𝘁𝗲𝗿𝗻𝗮𝗹: john@customer.com
𝗟𝗶𝘀𝘁: all@example.com
𝗥𝗲𝗽𝗹𝘆 𝗮𝗹𝗹: goes to 7 people
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

**Outlook on iOS and Android** (version 4.2425 or newer): the warning bar only. Outlook mobile has no Smart Alerts, so there is no check at Send, and it shows at most one bar per message. Settings are edited on desktop and sync through the mailbox.

Not supported: legacy Outlook for Mac.

## Install (2 minutes)

You only need one file, the manifest. The add-in's code is served from GitHub Pages.

1. **Download the manifest:** <https://rokvintar.github.io/respice-for-outlook/manifest.xml> (right-click → *Save Link As…*, keep the name `manifest.xml`).
2. **Add it to Outlook:** open <https://aka.ms/olksideload> → **My add-ins** → **Custom Addins** → **Add a custom add-in** → **Add from File…** → pick `manifest.xml` → confirm.
3. **Restart Outlook.** The add-in is attached to your mailbox, so it appears in Outlook for Mac, Windows, web and mobile.

Then open any message → **Apps** (or the ribbon) → **Respice settings** to add your mailing-list addresses and choose each rule's mode.

If **Custom Addins** is missing, your organisation has disabled user-installed add-ins. Ask IT to deploy the manifest through the Microsoft 365 admin center (**Settings → Integrated apps → Upload custom apps**), which also installs it for a whole team.

**Updating:** code changes reach you automatically after an Outlook restart. Only when the manifest version changes (see the release notes) do you need to remove the add-in and add the new `manifest.xml` again; copy your list addresses first.

## Host your own copy (developers)

1. Fork the repo and enable **Settings → Pages → Deploy from a branch → `main` / root** (a free GitHub plan needs a public repo).
2. Point the manifest at your Pages URL and commit it:
   ```bash
   node scripts/build-manifest.mjs https://<you>.github.io/respice-for-outlook
   ```
3. Optional: `npm run validate` (Microsoft's online manifest validator).
4. Install your `manifest.xml` as above.

## Limitations

- **Reply vs. Reply all:** Outlook's add-in API doesn't distinguish Reply from Reply all, and at send time Outlook on Mac and on the web report the compose type as `newMail` even for replies. Rule 4 therefore treats a message as a reply when any of these hold: Outlook reports `reply`; the draft has an `inReplyTo` ID (the message it answers, Mailbox 1.14+); or the subject starts with a reply prefix (`RE:`, `AW:`, `SV:`, `Odg:`, `VS:`, `RIF:`, `Antw:`, `答复:` …). Forwards (`FW:`, `Fwd:`, `WG:`, `TR:` …) are never treated as replies. It fires when the reply has more than one recipient, so a plain reply where you added people yourself also triggers it.
- **Distribution lists aren't expanded.** A list counts as one recipient for rule 3, and external members hidden inside a list aren't detected by rule 1.
- **Fail-open:** if the add-in can't load (offline, GitHub Pages down), Outlook lets the mail go (`SendMode="PromptUser"`).
- The dialog text is capped at about 500 characters. Long address lists are shortened to "+N more".
- **The warning bar is informational:** it never blocks; the Send dialog is the gate. Its text is capped at 150 characters (Outlook's limit). It uses Outlook's info style (info icon; the Respice icon appears only in classic Outlook on Windows); you can dismiss it, and it comes back on the next recipient change if the issue remains.
- **Reply in the reading pane on Mac:** when the Send dialog has to appear, Outlook first opens the reply in its own window. That's Outlook's behaviour for add-in dialogs and can't be changed by the add-in; the warning bar lets you fix things before pressing Send.

## Development

No build step and no runtime dependencies. Plain HTML/JS, served as-is. Tests use Node's built-in runner plus `acorn` (dev only).

```
src/launchevent.js   rule logic + OnMessageSend handler (self-contained: classic Windows loads it without a DOM)
src/commands.html    event runtime page for web / Mac / new Outlook
src/taskpane.*       settings UI
manifest.template.xml → scripts/build-manifest.mjs → manifest.xml
test/                Node unit tests
```

```bash
npm install   # once, dev dependencies for the syntax check
npm test
```

`src/launchevent.js` must stay ECMAScript 2016 with no `? :` operators, because classic Outlook on Windows before Version 2403 won't load it otherwise. `test/syntax.test.js` enforces this.

The version lives only in `package.json`. `scripts/build-manifest.mjs` writes it into the manifest's `<Version>`, the description shown in Outlook's add-in details, and the settings-button tooltip. Bump it and rebuild the manifest whenever the manifest changes; also update the `?v=` cache-busting tags in `src/*.html`. After a manifest change, remove the add-in in Outlook and add the new `manifest.xml` again (copy your list addresses first; settings may be reset). Code-only changes go live as soon as GitHub Pages redeploys.

## Microsoft Marketplace

This is possible later. Smart Alerts add-ins qualify for Marketplace only with `PromptUser` or `SoftBlock` send mode (this add-in uses `PromptUser`). A listing additionally needs a Partner Center account, a privacy policy URL and a support URL.

## License

MIT
