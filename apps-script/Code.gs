/**
 * ที่เก็บคะแนนโหวตของ newyear-poll
 * วางโค้ดนี้ใน Apps Script ของ Google Sheet แล้ว Deploy เป็น Web app (ดู README)
 *
 * ชีต "votes" มีคอลัมน์: id | name | place | other | date | updatedAt
 * คนหนึ่งคน (id จากเบราว์เซอร์) มีได้หนึ่งแถว โหวตซ้ำ = แก้แถวเดิม
 */
const SHEET_NAME = 'votes';
const HEADERS = ['id', 'name', 'place', 'other', 'date', 'updatedAt'];
const MAX_VOTES = 500;

function doGet() {
  return json_({ ok: true, votes: readVotes_() });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const vote = {
      id: text_(d.id, 64),
      name: text_(d.name, 30),
      place: text_(d.place, 80),
      other: text_(d.other, 60),
      date: text_(d.date, 40),
    };
    if (!vote.id || !vote.name || !vote.place || !vote.date) {
      return json_({ ok: false, error: 'missing_field' });
    }

    const sh = sheet_();
    const last = sh.getLastRow();
    let row = 0;
    if (last > 1) {
      const ids = sh.getRange(2, 1, last - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === vote.id) { row = i + 2; break; }
      }
    }
    if (!row) {
      if (last - 1 >= MAX_VOTES) return json_({ ok: false, error: 'poll_full' });
      row = last + 1;
    }
    const values = [vote.id, vote.name, vote.place, vote.other, vote.date, new Date()];
    // เก็บเป็นข้อความล้วน กันชีตแปลง "24 ธ.ค." เป็นวันที่เอง
    sh.getRange(row, 1, 1, 5).setNumberFormat('@');
    sh.getRange(row, 1, 1, HEADERS.length).setValues([values]);
    return json_({ ok: true, votes: readVotes_() });
  } catch (err) {
    return json_({ ok: false, error: 'server_error' });
  } finally {
    try { lock.releaseLock(); } catch (err) {}
  }
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** ส่งกลับเฉพาะข้อมูลที่แสดงบนเว็บ ไม่ส่ง id เพื่อไม่ให้คนอื่นแก้โหวตแทนกันได้ */
function readVotes_() {
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, HEADERS.length).getValues()
    .filter(function (r) { return r[0]; })
    .map(function (r) {
      return { name: String(r[1]), place: String(r[2]), other: String(r[3]), date: String(r[4]) };
    });
}

/** ตัดความยาว และกันไม่ให้ข้อความถูกอ่านเป็นสูตรในชีต */
function text_(v, max) {
  let s = String(v == null ? '' : v).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
