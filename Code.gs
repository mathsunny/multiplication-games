/** Google Sheets > 확장 프로그램 > Apps Script에 붙여 넣으세요. */
const SHEET_NAME = '메모리게임기록';
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const r = JSON.parse(e.postData.contents);
    const grade = Number(r.grade), room = Number(r.classroom), no = Number(r.number);
    const sets = Number(r.sets), attempts = Number(r.attempts), seconds = Number(r.seconds);
    if (![grade,room,no,sets,attempts,seconds].every(Number.isInteger) || grade<1 || grade>6 || room<1 || room>15 || no<1 || no>40 || ![3,4,6].includes(sets) || attempts<sets || attempts>10000 || seconds<0 || seconds>86400) throw Error('Invalid result');
    const book = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = book.getSheetByName(SHEET_NAME);
    if (!sheet) {sheet=book.insertSheet(SHEET_NAME);sheet.appendRow(['제출시각','학년','반','번호','세트 수','시도 횟수','소요 시간(초)','완료']);sheet.setFrozenRows(1);}
    sheet.appendRow([new Date(),grade,room,no,sets,attempts,seconds,'완료']);
    return ContentService.createTextOutput('OK');
  } catch (err) {return ContentService.createTextOutput('ERROR');}
  finally {lock.releaseLock();}
}