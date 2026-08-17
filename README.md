# YED 2026 Audition RSVP

A simple Google Apps Script mail merge and RSVP tracking system for boarding school auditions.

## Features

* Sends personalised audition emails using a Gmail draft
* Mail merges details such as name, date and time
* Generates a unique confirmation link for each recipient
* Records attendance confirmation in Google Sheets

## Deployment
### Initial Deployment
In Apps Script:

1. Go to Deploy → New deployment.
2. Select Web app.
3. Set Execute as to Me.
4. Set access so that all intended recipients can open the web app.
5. Click Deploy and authorise the script if prompted.
6. Copy the /exec web app URL.
7. Paste that URL into WEB_APP_URL in Config.js.
8. Save the project again.

### Subsequent Deployment
If you change server-side RSVP code later, update the deployment through:

**Deploy → Manage deployments → Edit → New version → Deploy**

## Usage

1. Prepare the audition details in Google Sheets.
2. Generate confirmation tokens.
3. Generate personalised confirmation links.
4. Prepare the Gmail draft using the required mail merge placeholders (e.g. `{{Name}}`).
5. Send the audition emails using the **Auditions** menu.
6. Recipients can open their link and click **Confirm Attendance**.
7. Their confirmation status is automatically updated in the spreadsheet.

## Notes

This project was created for internal boarding school audition management.
