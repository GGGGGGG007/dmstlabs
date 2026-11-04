/**
 * 견적 기록 클라우드 저장 (index.html의 PRESET_URL Apps Script에 추가)
 *
 * 설치 방법
 * 1) 프리셋·거래처를 저장하는 Apps Script 프로젝트를 연다 (PRESET_URL로 배포된 프로젝트).
 * 2) 이 파일 내용을 새 스크립트 파일(quotes.gs)로 붙여 넣는다.
 * 3) 기존 doPost(e) 함수 맨 앞쪽, JSON을 읽은 바로 다음 줄에 아래 한 줄을 넣는다.
 *      var quoteRes = handleQuoteAction(body); if (quoteRes) return quoteRes;
 *    (body는 JSON.parse(e.postData.contents) 한 값. 변수 이름이 다르면 그 이름으로 바꾼다.)
 * 4) 배포 → 배포 관리 → 기존 배포 수정 → 버전 "새 버전" → 배포. (URL은 그대로 유지됨)
 *
 * 스프레드시트에 "견적" 시트가 자동으로 생기고, 한 줄에 견적 하나씩 저장된다.
 */
var QUOTE_SHEET = '견적';

function handleQuoteAction(body) {
  if (!body || ['getQuotes', 'saveQuote', 'deleteQuotes'].indexOf(body.action) < 0) return null;
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = quoteSheet_();
    var out;
    if (body.action === 'getQuotes') {
      var rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues() : [];
      out = { ok: true, quotes: rows.filter(function (r) { return r[0]; }).map(function (r) {
        try { return JSON.parse(r[4]); } catch (e) { return null; }
      }).filter(Boolean) };
    } else if (body.action === 'saveQuote') {
      var q = body.quote || {};
      if (!body.id) throw new Error('id 없음');
      q.id = body.id;
      var row = [q.id, q.name || '', q.client || '', q.savedAt || new Date().toISOString(), JSON.stringify(q)];
      var at = findQuoteRow_(sh, q.id);
      if (at) sh.getRange(at, 1, 1, 5).setValues([row]); else sh.appendRow(row);
      out = { ok: true, id: q.id };
    } else {
      var ids = body.ids || [];
      for (var i = 0; i < ids.length; i++) {
        var r2 = findQuoteRow_(sh, ids[i]);
        if (r2) sh.deleteRow(r2);
      }
      out = { ok: true };
    }
    return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err && err.message || err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function quoteSheet_() {
  // 스프레드시트에 붙은 스크립트가 아니면 QUOTE_SPREADSHEET_ID 스크립트 속성에 시트 ID를 넣어 둔다
  var ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('QUOTE_SPREADSHEET_ID'));
  var sh = ss.getSheetByName(QUOTE_SHEET);
  if (!sh) {
    sh = ss.insertSheet(QUOTE_SHEET);
    sh.appendRow(['id', '이름', '거래처', '저장시각', '데이터(JSON)']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function findQuoteRow_(sh, id) {
  if (sh.getLastRow() < 2) return 0;
  var hit = sh.getRange(2, 1, sh.getLastRow() - 1, 1).createTextFinder(String(id)).matchEntireCell(true).findNext();
  return hit ? hit.getRow() : 0;
}
