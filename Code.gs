/**
 * ROOHI – receives website orders and enquiries and writes them to this Google Sheet.
 * Paste this whole file into Extensions > Apps Script (replace everything), then deploy as a Web app.
 */

// 1) Choose any secret word. It must be the same as "sheetKey" in the website's CONFIG.
var SECRET_KEY = "change-this-secret-word";

// 2) Optional: an email address that should be notified about every new order and enquiry. Leave "" for none.
var NOTIFY_EMAIL = "";

var ORDERS_SHEET = "Orders";
var ENQUIRIES_SHEET = "Enquiries";
var STATUSES = ["New", "Confirmed", "Dispatched", "Delivered", "Cancelled"];

var ORDER_HEADERS = ["Order ID", "Date", "Status", "Customer name", "Mobile", "WhatsApp", "Email",
  "Emirate", "Area", "Street / building", "Apartment / unit", "Landmark", "Map link", "Delivery notes",
  "Items", "Items count", "Subtotal", "Delivery fee", "Total to collect", "Payment", "Order note"];
var ENQUIRY_HEADERS = ["Enquiry ID", "Date", "Status", "Name", "Email", "Phone", "Message"];

// Opening the Web app URL in a browser should show this text (a quick way to test the deployment).
function doGet() {
  return ContentService.createTextOutput("ROOHI order endpoint is running.");
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var d = JSON.parse(e.postData.contents);
    if (d.key !== SECRET_KEY) return reply("forbidden");
    if (d.type === "order") saveOrder(d);
    else if (d.type === "enquiry") saveEnquiry(d);
    else return reply("unknown type");
    return reply("ok");
  } catch (err) {
    return reply("error: " + err);
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function reply(text) {
  return ContentService.createTextOutput(text);
}

function getSheet(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    var head = sh.getRange(1, 1, 1, headers.length);
    head.setFontWeight("bold").setBackground("#15100c").setFontColor("#ecdfc8");
    sh.setFrozenRows(1);
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).build();
    sh.getRange(2, 3, 2000, 1).setDataValidation(rule);
  }
  return sh;
}

function alreadySaved(sh, id) {
  if (sh.getLastRow() < 2) return false;
  var found = sh.getRange(2, 1, sh.getLastRow() - 1, 1).createTextFinder(id).matchEntireCell(true).findNext();
  return found !== null;
}

function txt(v) {
  return String(v === null || v === undefined ? "" : v).slice(0, 3000);
}

function stamp(iso) {
  var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var dt = iso ? new Date(iso) : new Date();
  if (isNaN(dt.getTime())) dt = new Date();
  return Utilities.formatDate(dt, tz, "yyyy-MM-dd HH:mm");
}

function saveOrder(d) {
  var sh = getSheet(ORDERS_SHEET, ORDER_HEADERS);
  var id = txt(d.orderId);
  if (!id || alreadySaved(sh, id)) return;            // ignore duplicates (e.g. automatic retries)
  var row = sh.getLastRow() + 1;
  var n = ORDER_HEADERS.length;
  sh.getRange(row, 1, 1, n).setNumberFormat("@");     // store as plain text so nothing is treated as a formula
  sh.getRange(row, 16, 1, 4).setNumberFormat("#,##0");
  sh.getRange(row, 1, 1, n).setValues([[
    id, stamp(d.date), "New", txt(d.name), txt(d.phone), txt(d.whatsapp), txt(d.email),
    txt(d.emirate), txt(d.area), txt(d.street), txt(d.unit), txt(d.landmark), txt(d.mapLink), txt(d.deliveryNotes),
    txt(d.items), Number(d.itemCount) || 0, Number(d.subtotal) || 0, Number(d.delivery) || 0, Number(d.total) || 0,
    txt(d.payment), txt(d.orderNote)
  ]]);
  sh.getRange(row, 15).setWrap(true);
  sh.getRange(row, 1, 1, n).setVerticalAlignment("top");
  notify("New ROOHI order " + id,
    "Customer: " + txt(d.name) + "\nMobile: " + txt(d.phone) + "\nAddress: " + [d.street, d.unit, d.area, d.emirate].join(", ") +
    "\n\nItems:\n" + txt(d.items) + "\n\nTotal to collect: " + txt(d.total) + " (" + txt(d.payment) + ")");
}

function saveEnquiry(d) {
  var sh = getSheet(ENQUIRIES_SHEET, ENQUIRY_HEADERS);
  var id = txt(d.id);
  if (!id || alreadySaved(sh, id)) return;
  var row = sh.getLastRow() + 1;
  sh.getRange(row, 1, 1, ENQUIRY_HEADERS.length).setNumberFormat("@");
  sh.getRange(row, 1, 1, ENQUIRY_HEADERS.length).setValues([[
    id, stamp(d.date), "New", txt(d.name), txt(d.email), txt(d.phone), txt(d.message)
  ]]);
  sh.getRange(row, 7).setWrap(true);
  notify("New ROOHI enquiry from " + txt(d.name), txt(d.message) + "\n\nEmail: " + txt(d.email) + "\nPhone: " + txt(d.phone));
}

function notify(subject, body) {
  if (!NOTIFY_EMAIL) return;
  try { MailApp.sendEmail(NOTIFY_EMAIL, subject, body); } catch (err) {}
}
