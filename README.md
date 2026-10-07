# my-app

BWMS: a simple warehouse support app that works alongside the main WMS.

## What is in this repository

- `index.html` - the whole app in one file. Open it in a browser or host it on any static host.
- `Code.gs` - the Google Apps Script for the Google Sheet that stores picking progress.

## What the app does

- **Outbound > Picking:** drop in a pick ticket PDF, copy a picker link, and send it to a picker. The picker opens the link, enters their name, and confirms each line on their phone.
- **Outbound > Outbound dashboard:** shows every pick list issued, who is picking it, lines and quantity done, and time against the standard (120 qty per hour).
- **Inbound:** menu placeholders only in this build.

## Setting up the dashboard

1. Create a Google Sheet, open Extensions > Apps Script, and paste in `Code.gs`.
2. Run `setup` once, then deploy as a web app (Execute as: Me, Access: Anyone).
3. Open the app, go to Outbound > Outbound dashboard, and paste the web app URL (it ends in `/exec`).

Without this step the picking page still works, and results can be sent by WhatsApp.
