/**
 * Partee Animalz RSVP: Google Sheet receiver
 *
 * Paste this whole file into Extensions > Apps Script (from the Google Sheet
 * that should hold the replies), then deploy it as a Web app.
 *
 * - One row per phone number. If a guest replies again, their row is updated.
 * - The web app only accepts replies. It never sends the guest list back out.
 * - A "Summary" tab counts heads attending and replies.
 * - Every reply also emails the account that owns this script.
 */

var SHEET_NAME = 'RSVPs';
var MAX_PARTY = 6;

function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    if (d.site) return out('ok'); // hidden spam-trap field was filled in

    var name = String(d.name || '').trim().slice(0, 80);
    var phone = String(d.phone || '').trim().slice(0, 24);
    var digits = phone.replace(/\D/g, '');
    var attending = d.status !== 'declined';
    var party = attending ? Math.max(1, Math.min(MAX_PARTY, parseInt(d.party, 10) || 1)) : 0;
    if (name.length < 2 || digits.length < 7 || digits.length > 15) return out('invalid');

    var total = 0, replies = 0;
    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      var sh = getSheet_();
      var now = new Date();
      var status = attending ? 'Attending' : 'Not attending';
      var last = sh.getLastRow();
      var row = -1;
      if (last > 1) {
        var phones = sh.getRange(2, 2, last - 1, 1).getValues();
        for (var i = 0; i < phones.length; i++) {
          if (String(phones[i][0]).replace(/\D/g, '') === digits) { row = i + 2; break; }
        }
      }
      if (row > 0) {
        sh.getRange(row, 1, 1, 4).setValues([[name, phone, status, party]]);
        sh.getRange(row, 6).setValue(now);
      } else {
        sh.appendRow([name, phone, status, party, now, now]);
      }

      var n = sh.getLastRow() - 1;
      if (n > 0) {
        var data = sh.getRange(2, 3, n, 2).getValues();
        replies = data.length;
        for (var j = 0; j < data.length; j++) {
          if (data[j][0] === 'Attending') total += Number(data[j][1]) || 0;
        }
      }
    } finally {
      lock.releaseLock();
    }

    // Email the host a note for every reply. A mail failure never blocks the RSVP.
    try {
      var to = Session.getEffectiveUser().getEmail();
      var who = name + ' (' + phone + ')';
      MailApp.sendEmail(
        to,
        'RSVP: ' + name + ' - ' + (attending ? 'Attending (' + party + ')' : 'Not attending'),
        who + '\n' +
        (attending ? 'Attending with ' + party + (party > 1 ? ' heads' : ' head') : 'Can\'t make it') + '\n\n' +
        'Total heads attending so far: ' + total + '\n' +
        'Total replies: ' + replies
      );
    } catch (mailErr) {}

    return out('ok');
  } catch (err) {
    return out('error');
  }
}

function doGet() {
  return out('Partee Animalz RSVP endpoint is running.');
}

function out(text) {
  return ContentService.createTextOutput(text);
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, 6).setValues([['Name', 'Phone', 'Status', 'Heads', 'First reply', 'Last updated']]).setFontWeight('bold');
    sh.getRange('B:B').setNumberFormat('@'); // keep leading zeros in phone numbers
    sh.getRange('E:F').setNumberFormat('yyyy-mm-dd hh:mm');
    sh.setFrozenRows(1);
    sh.setColumnWidths(1, 6, 150);

    var sum = ss.getSheetByName('Summary') || ss.insertSheet('Summary');
    sum.getRange('A1:B3').setValues([
      ['Heads attending', "=SUMIF(RSVPs!C:C,\"Attending\",RSVPs!D:D)"],
      ['Replies', "=MAX(0,COUNTA(RSVPs!A:A)-1)"],
      ['Not attending', "=COUNTIF(RSVPs!C:C,\"Not attending\")"]
    ]);
    sum.getRange('A1:A3').setFontWeight('bold');
    sum.setColumnWidth(1, 160);
  }
  return sh;
}
