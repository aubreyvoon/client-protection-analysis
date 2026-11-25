# Client Protection Analysis

A one-page web app for client protection reviews (life, critical illness, PA, medical, retirement, will / trust). All amounts in RM. Works on iPhone, iPad and Mac in Safari. Add it to the Home Screen to use it like an app.

- Client details are kept only in your browser on your own device. Nothing is sent anywhere until you tap **Save to Google Sheet**.
- **Save as PDF** opens the print dialog; choose "Save as PDF" (iPhone / iPad: Share > Print, then pinch out on the preview).
- This repository contains no client data, no sheet link and no keys.

## Connect your Google Sheet (once)

Every save adds a **new tab named after the client** ("Name", then "Name (2)" if saved again).

1. Open your Google Sheet, then **Extensions > Apps Script**. Delete the sample code, paste in `apps-script/Code.gs`, and save.
2. Click **Project Settings** (gear icon) > **Script properties** > **Add script property**. Name: `SECRET_KEY`. Value: a long random password you make up.
3. Click **Deploy > New deployment**. Type: **Web app**. Execute as: **Me**. Who has access: **Anyone**. Authorise when asked, then copy the **Web app URL** (ends with `/exec`).
4. Open the app, expand **Google Sheets connection**, paste the URL and the same secret key, then tap **Test connection**.
5. Fill in a client, then tap **Save to Google Sheet** (tap again to confirm).

If you edit `Code.gs` later: **Deploy > Manage deployments > Edit > New version**.

The secret key stops other people from writing to your sheet through the web app URL. Keep both private, and change the key in Script properties if you think it has leaked.
