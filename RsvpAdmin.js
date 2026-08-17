// ============================================================
// RSVP ADMIN / SHEET MENU
// ============================================================

/**
 * Creates the custom "Auditions" menu.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("Auditions")
    .addItem(
      "1. Generate Confirmation Tokens",
      "generateConfirmationTokens"
    )
    .addItem(
      "2. Generate Confirmation Links",
      "generateConfirmationLinks"
    )
    .addSeparator()
    .addItem(
      "Send Audition Emails",
      "sendEmails"
    )
    .addToUi();
}



// ============================================================
// GENERATE CONFIRMATION TOKENS
// ============================================================

/**
 * Generates a unique token for every row.
 *
 * Runs on the ACTIVE TAB only.
 *
 * Also sets blank statuses to Pending.
 */
function generateConfirmationTokens() {
  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getActiveSheet();

  const data =
    sheet
      .getDataRange()
      .getValues();

  const headers = data[0];

  const tokenCol =
    headers.indexOf(TOKEN_COL);

  const statusCol =
    headers.indexOf(STATUS_COL);

  if (tokenCol === -1) {
    throw new Error(
      `Column "${TOKEN_COL}" not found.`
    );
  }

  if (statusCol === -1) {
    throw new Error(
      `Column "${STATUS_COL}" not found.`
    );
  }

  for (let i = 1; i < data.length; i++) {

    // Only create token if blank
    if (!data[i][tokenCol]) {
      sheet
        .getRange(
          i + 1,
          tokenCol + 1
        )
        .setValue(
          Utilities.getUuid()
        );
    }

    // Set initial status to Pending if blank
    if (!data[i][statusCol]) {
      sheet
        .getRange(
          i + 1,
          statusCol + 1
        )
        .setValue("Pending");
    }
  }
}



// ============================================================
// GENERATE CONFIRMATION LINKS
// ============================================================

/**
 * Generates a personalised confirmation URL for each recipient.
 *
 * Runs on the ACTIVE TAB only.
 *
 * TEST       -> TEST sheet ID
 * Emcee      -> Emcee sheet ID
 * Performers -> Performers sheet ID
 */
function generateConfirmationLinks() {
  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getActiveSheet();

  const data =
    sheet
      .getDataRange()
      .getValues();

  const headers = data[0];

  const tokenCol =
    headers.indexOf(TOKEN_COL);

  const linkCol =
    headers.indexOf(
      CONFIRMATION_LINK_COL
    );

  if (tokenCol === -1) {
    throw new Error(
      `Column "${TOKEN_COL}" not found.`
    );
  }

  if (linkCol === -1) {
    throw new Error(
      `Column "${CONFIRMATION_LINK_COL}" not found.`
    );
  }

  // Unique ID of whichever tab is active right now
  const sheetId =
    sheet.getSheetId();

  for (let i = 1; i < data.length; i++) {
    const token =
      data[i][tokenCol];

    if (token) {
      const confirmationLink =
        WEB_APP_URL +
        "#sheet=" +
        sheetId +
        "&token=" +
        encodeURIComponent(token);

      sheet
        .getRange(
          i + 1,
          linkCol + 1
        )
        .setValue(
          confirmationLink
        );
    }
  }
}
