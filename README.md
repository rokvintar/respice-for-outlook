# <img src="assets/icon-64.png" width="32" height="32" alt=""> Respice for Outlook

*Respice*, Latin for "look back". An Outlook add-in that makes you look once more before a risky email goes out.

## 📥 Install

1. Download [manifest.xml](https://rokvintar.github.io/respice-for-outlook/manifest.xml).
2. Go to [aka.ms/olksideload](https://aka.ms/olksideload) → **My add-ins** → **Add a custom add-in** → **Add from File** → select the file.
3. Restart Outlook.

**Settings:** open any message → **Apps** → **Respice settings**.

**No "Add a custom add-in" option?** Your IT blocks user add-ins. Ask them to deploy the manifest.

**Updates** install automatically. If a release says *new manifest*, remove the add-in and repeat the steps.

## 🔒 Privacy

> [!NOTE]
> Everything runs inside your Outlook. No data leaves it: no server, no analytics, no tracking. Settings are saved in your own mailbox.

## ⚙️ How it works

Respice warns in two ways:

- ***Bar*** – a one-line note above the message while you write. It updates whenever you change To, Cc or Bcc, e.g. `𝗘𝘅𝘁𝗲𝗿𝗻𝗮𝗹: john@customer.com · 𝗟𝗶𝘀𝘁: all@example.com · 𝗥𝗲𝗰𝗶𝗽𝗶𝗲𝗻𝘁𝘀: 23 (limit 15) · 𝗥𝗲𝗽𝗹𝘆 𝗮𝗹𝗹: 23 people`.
- ***Dialog*** – a pop-up when you press Send, with **Don't Send** / **Send Anyway**.

**Rules**

| Rule | Warns when |
|------|------------|
| External recipient | someone is outside your organisation (your email domain, or domains you list) |
| Mailing list | a recipient is one of your list addresses (`*` wildcards allowed) or an Outlook distribution list |
| Too many recipients | To + Cc + Bcc exceed your limit (default 15) |
| Reply to more than the sender | a reply goes to more than one person |

**Modes**

For each rule a mode can be set. Possible modes:

| Mode | While writing | At Send |
|------|---------------|---------|
| Bar + check at Send (default) | ***Bar*** | ***Dialog*** |
| Bar only | ***Bar*** | – |
| Check at Send only | – | ***Dialog*** |
| Off | – | – |

## 💻 Supported clients

| Outlook | Bar | Check at Send |
|---------|-----|---------------|
| Mac (new Outlook) | ✅ | ✅ |
| Web | ✅ | ✅ |
| Windows (new and classic, Microsoft 365) | ✅ | ✅ |
| iOS / Android | ✅ | – |
| Legacy Outlook for Mac | – | – |

## ⚠️ Known limitations

*As of 2026-11-10 (v1.3.1).*

- **Reply vs. Reply all:** Outlook's add-in API doesn't distinguish Reply from Reply all, and at send time Outlook on Mac and on the web report the compose type as `newMail` even for replies. Rule 4 therefore treats a message as a reply when any of these hold: Outlook reports `reply`; the draft has an `inReplyTo` ID (the message it answers, Mailbox 1.14+); or the subject starts with a reply prefix (`RE:`, `AW:`, `SV:`, `Odg:`, `VS:`, `RIF:`, `Antw:`, `答复:` …). Forwards (`FW:`, `Fwd:`, `WG:`, `TR:` …) are never treated as replies. It fires when the reply has more than one recipient, so a plain reply where you added people yourself also triggers it.
- **Distribution lists aren't expanded.** A list counts as one recipient for rule 3, and external members hidden inside a list aren't detected by rule 1.
- **Fail-open:** if the add-in can't load (offline, GitHub Pages down), Outlook lets the mail go (`SendMode="PromptUser"`).
- The dialog text is capped at about 500 characters. Long address lists are shortened to "+N more".
- **The warning bar is informational:** it never blocks; the Send dialog is the gate. Its text is capped at 150 characters (Outlook's limit). It uses Outlook's info style (info icon; the Respice icon appears only in classic Outlook on Windows); you can dismiss it, and it comes back on the next recipient change if the issue remains.
- **Reply in the reading pane on Mac:** when the Send dialog has to appear, Outlook first opens the reply in its own window. That's Outlook's behaviour for add-in dialogs and can't be changed by the add-in; the warning bar lets you fix things before pressing Send.

## 📄 License

MIT
