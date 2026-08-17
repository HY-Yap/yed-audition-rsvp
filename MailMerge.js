// ============================================================
// MAIL MERGE
// ============================================================

/**
 * Sends emails from sheet data.
 *
 * This runs on whichever tab is active when you use the menu.
 *
 * @param {string} subjectLine Optional Gmail draft subject line.
 * @param {Sheet} sheet Sheet to read data from.
 */
function sendEmails(
  subjectLine,
  sheet = SpreadsheetApp.getActiveSheet()
) {
  let processedSubjectLine = subjectLine;

  if (!processedSubjectLine) {
    processedSubjectLine = Browser.inputBox(
      "Mail Merge",
      "Type or copy/paste the subject line of the Gmail " +
        "draft message you would like to mail merge with:",
      Browser.Buttons.OK_CANCEL,
    );

    if (
      processedSubjectLine === "cancel" ||
      processedSubjectLine === ""
    ) {
      return;
    }
  }

  // Gets the Gmail draft used as the template
  const emailTemplate =
    getGmailTemplateFromDrafts_(processedSubjectLine);

  // Gets all displayed data from the active sheet
  const dataRange = sheet.getDataRange();
  const data = dataRange.getDisplayValues();

  // First row contains column headings
  const heads = data.shift();

  const emailSentColIdx =
    heads.indexOf(EMAIL_SENT_COL);

  if (emailSentColIdx === -1) {
    throw new Error(
      `Column "${EMAIL_SENT_COL}" not found.`
    );
  }

  // Converts the rows into objects using column headings
  const obj = data.map((r) =>
    heads.reduce((o, k, i) => {
      o[k] = r[i] || "";
      return o;
    }, {}),
  );

  // Stores Email Sent results
  const out = [];

  // Loop through rows
  obj.forEach((row) => {
    // Only send if Email Sent is blank
    if (row[EMAIL_SENT_COL] === "") {
      try {
        const msgObj =
          fillInTemplateFromObject_(
            emailTemplate.message,
            row
          );

        GmailApp.sendEmail(
          row[RECIPIENT_COL],
          msgObj.subject,
          msgObj.text,
          {
            htmlBody: msgObj.html,
            attachments:
              emailTemplate.attachments,
            inlineImages:
              emailTemplate.inlineImages,
          },
        );

        // Successful send
        out.push([new Date()]);

      } catch (e) {
        // Record error in Email Sent column
        out.push([e.message]);
      }

    } else {
      // Preserve existing Email Sent value
      out.push([row[EMAIL_SENT_COL]]);
    }
  });

  if (out.length > 0) {
    sheet
      .getRange(
        2,
        emailSentColIdx + 1,
        out.length
      )
      .setValues(out);
  }
}


// ------------------------------------------------------------
// Find Gmail draft
// ------------------------------------------------------------

function getGmailTemplateFromDrafts_(subjectLine) {
  try {
    const drafts = GmailApp.getDrafts();

    const draft =
      drafts.filter(
        subjectFilter_(subjectLine)
      )[0];

    if (!draft) {
      throw new Error("Draft not found");
    }

    const msg = draft.getMessage();

    // Inline images
    const allInlineImages =
      draft
        .getMessage()
        .getAttachments({
          includeInlineImages: true,
          includeAttachments: false,
        });

    // Normal attachments
    const attachments =
      draft
        .getMessage()
        .getAttachments({
          includeInlineImages: false
        });

    const htmlBody = msg.getBody();

    // Map image names to blobs
    const imgObj =
      allInlineImages.reduce(
        (obj, image) => {
          obj[image.getName()] = image;
          return obj;
        },
        {}
      );

    // Find cid images in Gmail HTML
    const imgexp =
      /<img.*?src="cid:(.*?)".*?alt="(.*?)"[^\>]+>/g;

    const matches =
      [...htmlBody.matchAll(imgexp)];

    const inlineImagesObj = {};

    for (const match of matches) {
      inlineImagesObj[match[1]] =
        imgObj[match[2]];
    }

    return {
      message: {
        subject: subjectLine,
        text: msg.getPlainBody(),
        html: htmlBody,
      },

      attachments: attachments,
      inlineImages: inlineImagesObj,
    };

  } catch (e) {
    throw new Error(
      "Oops - can't find Gmail draft"
    );
  }
}


// ------------------------------------------------------------
// Match Gmail draft subject
// ------------------------------------------------------------

function subjectFilter_(subjectLine) {
  return (element) => {
    return (
      element
        .getMessage()
        .getSubject() === subjectLine
    );
  };
}


// ------------------------------------------------------------
// Replace {{Column}} placeholders
// ------------------------------------------------------------

function fillInTemplateFromObject_(
  template,
  data
) {
  let templateString =
    JSON.stringify(template);

  templateString =
    templateString.replace(
      /{{[^{}]+}}/g,
      (key) => {
        return escapeData_(
          data[
            key.replace(/[{}]+/g, "")
          ] || ""
        );
      }
    );

  return JSON.parse(templateString);
}


// ------------------------------------------------------------
// Make data JSON-safe
// ------------------------------------------------------------

function escapeData_(str) {
  return str
    .replace(/[\\]/g, "\\\\")
    .replace(/[\"]/g, '\\"')
    .replace(/[\/]/g, "\\/")
    .replace(/[\b]/g, "\\b")
    .replace(/[\f]/g, "\\f")
    .replace(/[\n]/g, "\\n")
    .replace(/[\r]/g, "\\r")
    .replace(/[\t]/g, "\\t");
}
