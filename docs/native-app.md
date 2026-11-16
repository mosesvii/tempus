# Tempus app for Mac (and later iPhone)

The native app is much faster than the web page. It keeps its own copy of your entries and opens
straight away, then updates from the sheet in the background. It talks to `backend/nativeApi.js`, which
runs the same `api_*` functions as the web page, so all the rules stay in one place.

## One-time setup

1. Deploy this code (`clasp push`, or the updater script) so `backend/nativeApi.js` is in the project.
2. In the Apps Script editor, choose `setupNativeApiKey` in the function menu and click **Run**.
   Open **Execution log** and copy the key it printed.
3. **Deploy → New deployment → Web app.** Description "Tempus app", Execute as **Me**,
   Who has access **Anyone**. Copy the web app URL (ends in `/exec`).
   Keep your existing deployment as it is: that's the web page.
4. In the app, open **Settings** (⌘,), paste the URL and the key, and click **Connect**.

The key is what protects the new deployment, so keep it private. Running `setupNativeApiKey` again
makes a new key and the old one stops working. Opening that deployment's URL in a browser without
being signed in shows only a short notice, never the Tempus page.

## Updating

After later code changes, use **Deploy → Manage deployments**, edit the "Tempus app" deployment and
choose **New version**, or the app keeps running the old code.

## What the app can call

`NATIVE_API_ALLOWED` in `backend/nativeApi.js` lists the `api_*` functions the app may call, plus
`bootstrap`, which returns settings, hour types, contracts, entries and public holidays in one go.
