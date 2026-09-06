/**
 * Taste of Madurai — Prebooking backend
 * -------------------------------------------
 * Deploy this as a Google Apps Script Web App (see SETUP.md).
 * It writes every booking to a Google Sheet, saves the payment
 * screenshot to a Drive folder, and serves data back to admin.html.
 */

const SHEET_NAME = "Bookings";
const DRIVE_FOLDER_NAME = "Taste of Madurai - Payment Screenshots";
const PRICES = { mutton: 300, chicken: 250, veg: 200 };

// A shared token that must match in every request. This is NOT a strong
// secret (anyone can view it in your public index.html/admin.html source),
// but it stops random bots that scan the internet for open Apps Script
// URLs and try to spam them with junk data. Change this to your own string
// before deploying, and use the SAME string in index.html and admin.html.
const SHARED_TOKEN = "tom-9f8a2-madurai";

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      "Timestamp", "Booking ID", "Name", "Batch", "Contact", "Email",
      "Mutton Qty", "Chicken Qty", "Amount", "Screenshot URL", "Paid", "Cancelled", "Veg Qty"
    ]);
  } else {
    // Migrate an older sheet that's missing newer columns
    if (sheet.getLastColumn() < 12) sheet.getRange(1, 12).setValue("Cancelled");
    if (sheet.getLastColumn() < 13) sheet.getRange(1, 13).setValue("Veg Qty");
  }
  return sheet;
}

function getFolder_() {
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(DRIVE_FOLDER_NAME);
}

function makeBookingId_(sheet) {
  const lastRow = sheet.getLastRow(); // header is row 1
  const seq = lastRow; // header=1 means first booking gets seq=1
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Asia/Kolkata", "yyMMdd");
  return "TOM" + stamp + "-" + String(seq).padStart(3, "0");
}

function jsonError_(msg) {
  return ContentService.createTextOutput(JSON.stringify({ error: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);

  if (body.token !== SHARED_TOKEN) {
    return jsonError_("Unauthorized");
  }

  if (body.action === "markPaid") {
    return markPaid_(body.bookingId);
  }

  if (body.action === "cancelBooking") {
    return setCancelled_(body.bookingId, true);
  }

  if (body.action === "restoreBooking") {
    return setCancelled_(body.bookingId, false);
  }

  // New booking submission
  const sheet = getSheet_();
  const bookingId = makeBookingId_(sheet);
  const mutton = Math.max(0, Math.min(20, Number(body.mutton) || 0));
  const chicken = Math.max(0, Math.min(20, Number(body.chicken) || 0));
  const veg = Math.max(0, Math.min(20, Number(body.veg) || 0));
  // Recompute the amount from trusted prices rather than trusting the client.
  const amount = mutton * PRICES.mutton + chicken * PRICES.chicken + veg * PRICES.veg;

  let screenshotUrl = "";
  if (body.screenshot) {
    try {
      const parts = body.screenshot.split(",");
      const meta = parts[0]; // e.g. data:image/png;base64
      const b64 = parts[1];
      const mime = meta.substring(meta.indexOf(":") + 1, meta.indexOf(";"));
      const bytes = Utilities.base64Decode(b64);
      const blob = Utilities.newBlob(bytes, mime, bookingId + "_" + (body.screenshotName || "screenshot"));
      const file = getFolder_().createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      screenshotUrl = "https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w1000";
    } catch (err) {
      screenshotUrl = "";
    }
  }

  sheet.appendRow([
    new Date(),
    bookingId,
    String(body.name || "").slice(0, 100),
    String(body.batch || "").slice(0, 20),
    String(body.contact || "").slice(0, 20),
    String(body.email || "").slice(0, 100),
    mutton,
    chicken,
    amount,
    screenshotUrl,
    false,
    false,
    veg
  ]);

  return ContentService.createTextOutput(JSON.stringify({ bookingId: bookingId }))
    .setMimeType(ContentService.MimeType.JSON);
}

function markPaid_(bookingId) {
  const sheet = getSheet_();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === bookingId) {
      sheet.getRange(i + 1, 11).setValue(true); // "Paid" column
      break;
    }
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function setCancelled_(bookingId, cancelled) {
  const sheet = getSheet_();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === bookingId) {
      sheet.getRange(i + 1, 12).setValue(cancelled); // "Cancelled" column
      break;
    }
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  if (!e.parameter.token || e.parameter.token !== SHARED_TOKEN) {
    return jsonError_("Unauthorized");
  }
  const sheet = getSheet_();
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1); // skip header
  const bookings = rows.map(r => ({
    timestamp: r[0],
    bookingId: r[1],
    name: r[2],
    batch: r[3],
    contact: r[4],
    email: r[5],
    mutton: r[6],
    chicken: r[7],
    amount: r[8],
    screenshotUrl: r[9],
    paid: r[10] === true,
    cancelled: r[11] === true,
    veg: r[12] || 0
  }));
  return ContentService.createTextOutput(JSON.stringify({ bookings: bookings }))
    .setMimeType(ContentService.MimeType.JSON);
}
