// ============================================================
// RSVP WEB APP
// ============================================================

/**
 * Displays the RSVP webpage.
 *
 * Opening the page does NOT confirm attendance.
 * The user must press the Confirm Attendance button.
 */
function doGet() {
  return HtmlService
    .createHtmlOutputFromFile("RsvpPage")
    .setTitle("Audition Attendance Confirmation");
}

// ============================================================
// GET AUDITION DETAILS
// ============================================================

/**
 * Called when the webpage first opens.
 *
 * Read-only.
 * Does NOT confirm attendance.
 */
function getAuditionDetails(
  sheetId,
  token
) {

  const record =
    findAuditionRecord_(
      sheetId,
      token
    );


  if (!record) {

    return {
      success: false,
      message:
        "We could not find an audition associated with this link.",
    };

  }


  return {
    success: true,
    name: record.name,
    date: record.date,
    time: record.time,
    status: record.status,
  };
}



// ============================================================
// CONFIRM ATTENDANCE
// ============================================================

/**
 * Only called after the person presses
 * the Confirm Attendance button.
 */
function confirmAttendance(
  sheetId,
  token
) {

  // Prevent multiple confirmation writes from
  // modifying the spreadsheet at the same time.
  const lock =
    LockService.getScriptLock();

  // Wait up to 10 seconds if another confirmation
  // is currently being processed.
  lock.waitLock(10000);


  try {

    const record =
      findAuditionRecord_(
        sheetId,
        token
      );


    if (!record) {

      return {
        success: false,
        message:
          "We could not find an audition associated with this link.",
      };

    }


    // Do not overwrite the original confirmation
    // timestamp if they confirm again later.
    if (
      record.status !== "Confirmed"
    ) {

      record.sheet
        .getRange(
          record.rowNumber,
          record.statusCol + 1
        )
        .setValue(
          "Confirmed"
        );


      record.sheet
        .getRange(
          record.rowNumber,
          record.confirmedAtCol + 1
        )
        .setValue(
          new Date()
        );


      // Make sure spreadsheet changes are committed
      // before another confirmation gets the lock.
      SpreadsheetApp.flush();

    }


    return {
      success: true,
      name: record.name,
      date: record.date,
      time: record.time,
    };


  } finally {

    // Always release the lock,
    // even if an error occurs.
    lock.releaseLock();

  }

}



// ============================================================
// FIND AUDITION RECORD
// ============================================================

/**
 * Finds the correct row using:
 *
 * sheet ID + unique token
 *
 * Uses getDisplayValues() so Date and Time appear
 * exactly as formatted in Google Sheets.
 */
function findAuditionRecord_(
  sheetId,
  token
) {

  // Open overall spreadsheet file
  const spreadsheet =
    SpreadsheetApp.openById(
      SPREADSHEET_ID
    );


  // Find the correct tab
  const sheet =
    spreadsheet.getSheetById(
      Number(sheetId)
    );


  if (!sheet) {
    return null;
  }


  const data =
    sheet
      .getDataRange()
      .getDisplayValues();


  const headers =
    data[0];


  const tokenCol =
    headers.indexOf(
      TOKEN_COL
    );

  const nameCol =
    headers.indexOf(
      NAME_COL
    );

  const dateCol =
    headers.indexOf(
      DATE_COL
    );

  const timeCol =
    headers.indexOf(
      TIME_COL
    );

  const statusCol =
    headers.indexOf(
      STATUS_COL
    );

  const confirmedAtCol =
    headers.indexOf(
      CONFIRMED_AT_COL
    );


  // Ensure required headers exist
  if (
    tokenCol === -1 ||
    nameCol === -1 ||
    dateCol === -1 ||
    timeCol === -1 ||
    statusCol === -1 ||
    confirmedAtCol === -1
  ) {

    throw new Error(
      "One or more required columns are missing. " +
      "Required headers: " +
      NAME_COL + ", " +
      DATE_COL + ", " +
      TIME_COL + ", " +
      STATUS_COL + ", " +
      CONFIRMED_AT_COL + ", " +
      TOKEN_COL + "."
    );

  }


  // Find matching token
  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    if (
      String(
        data[i][tokenCol]
      ) ===
      String(token)
    ) {

      return {
        sheet: sheet,

        rowNumber:
          i + 1,

        name:
          data[i][nameCol],

        date:
          data[i][dateCol],

        time:
          data[i][timeCol],

        status:
          data[i][statusCol],

        statusCol:
          statusCol,

        confirmedAtCol:
          confirmedAtCol,
      };

    }
  }


  return null;
}
