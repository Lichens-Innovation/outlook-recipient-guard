# Recipient Guard — Outlook add-in

Intercepts sending (Smart Alerts, `OnMessageSend` event) and shows a
**"Send anyway / Don't send"** alert if a recipient (To, Cc, Bcc) matches a rule.
Protects against autocomplete mistakes.

## Files

| File | Purpose |
|---|---|
| `manifest.xml` | Load in Outlook via *Add a custom add-in → Add from File* |
| `launchevent.js` | Check logic (`onMessageSendHandler` handler) |
| `commands.html` | HTML runtime (Outlook on the web, new Outlook, Mac) |
| `taskpane.html` | Configuration task pane ("Destinataires surveillés" button while composing) |
| `assets/` | Icons |

The rule list **is not in the code**: it is stored in your mailbox's
`roamingSettings` (private, follows your account across all clients).

## Installation

1. **Host the files over HTTPS.** The manifest only contains URLs; Outlook loads the JS
   from the web. This repo deploys itself to GitHub Pages on every push to `main`
   (`.github/workflows/pages.yml`, Settings → Pages → Source: *GitHub Actions*). The workflow
   fails if a file referenced by `manifest.xml` is missing, if the Pages URL differs from the
   manifest's base URL, or if any manifest URL does not return 200.
   The default expected URL is `https://lichens-innovation.github.io/outlook-recipient-guard/`.
   Different host? `./set-base-url.sh https://my-host/path` updates the manifest.
2. Check that `https://…/commands.html` and `https://…/launchevent.js` respond.
3. Outlook → *Add-ins* → *My add-ins* → **Add a custom add-in → Add from File** →
   `manifest.xml`.
4. Open a new email → ribbon → **Destinataires surveillés** → enter rules →
   *Enregistrer* (Save). The *Tester ce brouillon* (Test this draft) button checks the current recipients.

## Rule syntax

```
# comment
jean.tremblay@gmail.com | Jean's personal address   ← exact + reason shown
@ancien-client.com                                  ← whole domain
andre*@gmail.com                                    ← * wildcards
```

Option: *Also alert for any external recipient* + list of internal domains
(e.g. `lichens.ai`) — subdomains are considered internal.

## Behavior

- `SendMode="PromptUser"`: the alert offers **Don't send** (back to the draft)
  or **Send anyway**. To block with no way to override, change it to
  `SoftBlock` in the manifest.
- If the script errors, sending is never blocked.
- Requires Mailbox 1.12: Outlook on the web, new Outlook (Windows/Mac), recent classic
  Outlook for Windows. Not supported on mobile (mail is sent without checking).

## Updating

- JS/HTML changes → push to the host (Outlook cache: clear it or wait a few minutes).
- Manifest changes → increment `<Version>`, remove then re-add the add-in.
