# contact-sankhacooray-com

A tiny, free, serverless **Google Apps Script** backend for the contact form on
[v3.sankhacooray.com](https://v3.sankhacooray.com). Each inquiry:

1. is appended as a row to a private Google Sheet (your inbox of leads),
2. triggers a branded HTML auto-reply to the sender ("got your message"),
3. (optionally) emails **you** an internal alert with reply-to set to the sender.

No servers, no database, no transactional-email account. Runs entirely inside
your own Google account.

```
src/
├── appsscript.json   # manifest: scopes + web-app config (Anyone, execute as me)
└── Code.js           # doGet / doPost / submitInquiry / setup / email helpers
```

## Setup (paste into your own Google account)

1. Go to <https://script.google.com> → **New project**.
2. Replace the default `Code.gs` contents with [`src/Code.js`](src/Code.js).
3. Show the manifest: **Project Settings → ☑ Show "appsscript.json"**, then open
   `appsscript.json` in the editor and paste [`src/appsscript.json`](src/appsscript.json).
4. **Run → `setup`.** Authorize the scopes when prompted. The execution log prints
   the URL of a new Google Sheet ("sankhacooray.com — Contact inquiries") created
   in your Drive. This is where inquiries land.
5. **Deploy → New deployment → Web app:**
   - Description: `v1.0`
   - Execute as: **Me**
   - Who has access: **Anyone**
   - Deploy → copy the **Web app URL** (ends in `/exec`).
6. Paste that URL into the website: `v3-sankhacooray-com/js/contact.js` →
   `APPS_SCRIPT_URL`.

## Optional Script Properties

Project Settings → **Script Properties** → Add:

| Key            | Value                              | Effect |
|----------------|------------------------------------|--------|
| `NOTIFY_EMAIL` | your inbox, e.g. `you@gmail.com`   | Sends you an alert per inquiry (reply-to = sender). Unset = no alert. |
| `FROM_EMAIL`   | e.g. `hello@sankhacooray.com`      | Send the auto-reply *from* this alias. Must be a verified "Send mail as" alias in your Gmail. Unset/unverified → falls back to your primary address. |
| `FROM_NAME`    | e.g. `Sankha Cooray`               | Sender display name. Defaults to `Sankha Cooray`. |

To use `FROM_EMAIL`: Gmail → Settings → Accounts → **Send mail as** → add the
address, verify with the emailed code, then set the property.

## Updating the code later

After editing `Code.js` in the IDE you must **redeploy** for the live `/exec`
URL to serve the new code: **Deploy → Manage deployments → edit (pencil) →
Version: New version → Deploy**. This keeps the same URL.

## Notes

- **CORS:** the form POSTs with `Content-Type: text/plain` (a CORS "simple
  request", no preflight). The `ANYONE_ANONYMOUS` deployment auto-adds
  `Access-Control-Allow-Origin: *`, so the browser can read the JSON response.
- **Spam:** a hidden honeypot field (`company`) silently drops bot submissions.
- **Quotas:** consumer Gmail allows ~100 emails/day, which is plenty for a
  personal contact form.
- The frontend lives in the `v3-sankhacooray-com` repo; this repo is backend only.
