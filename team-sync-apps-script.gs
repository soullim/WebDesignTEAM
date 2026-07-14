/**
 * 웹디자인팀 일정보드 — 팀 공유 백엔드 (Google Apps Script)
 *
 * ── 설정 방법 (최초 1회, 약 5분) ─────────────────────────────
 * 1. https://sheets.new 에서 새 구글 시트 생성 (이름 예: "일정보드 데이터")
 * 2. 시트 메뉴에서 [확장 프로그램] → [Apps Script] 클릭
 * 3. 편집기에 기본으로 있는 코드를 지우고 이 파일 내용 전체를 붙여넣기 → 저장(💾)
 * 4. 우측 상단 [배포] → [새 배포] 클릭
 *    - 유형 선택(⚙ 아이콘): "웹 앱"
 *    - 설명: 아무거나 (예: 일정보드 v1)
 *    - 다음 사용자 인증 정보로 실행: "나"
 *    - 액세스 권한이 있는 사용자: "모든 사용자"  ← 중요!
 * 5. [배포] 클릭 → 권한 승인 창이 뜨면 본인 계정 선택 → "고급" → "이동" → 허용
 * 6. 발급된 "웹 앱 URL" (https://script.google.com/macros/s/...../exec) 을 복사해서
 *    일정보드 담당자(또는 Claude)에게 전달 → index.html의 SB_GAS_URL에 입력하면 완료
 *
 * ── 이후 코드를 수정했을 때 ──────────────────────────────────
 * [배포] → [배포 관리] → 연필(✏) 아이콘 → 버전: "새 버전" → [배포]
 * (새 배포를 만들면 URL이 바뀌니 반드시 "배포 관리"에서 수정할 것)
 * ──────────────────────────────────────────────────────────
 */

var SHEET_NAME = 'board';
var CHUNK = 45000; // 셀당 5만자 제한 대비 분할 저장

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  return sh;
}

function readData_() {
  var sh = getSheet_();
  var last = sh.getLastRow();
  if (last < 1) return '';
  var vals = sh.getRange(1, 1, last, 1).getValues();
  return vals.map(function (r) { return r[0]; }).join('');
}

function writeData_(str) {
  var sh = getSheet_();
  sh.clearContents();
  var rows = [];
  for (var i = 0; i < str.length; i += CHUNK) rows.push([str.slice(i, i + CHUNK)]);
  if (rows.length) sh.getRange(1, 1, rows.length, 1).setValues(rows);
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* 읽기: 일정보드가 페이지를 열 때 호출 */
function doGet() {
  var raw = readData_();
  if (!raw) return out_({});
  return ContentService.createTextOutput(raw).setMimeType(ContentService.MimeType.JSON);
}

/* 쓰기: [팀 공유] 버튼이 호출 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.tryLock(5000);
    var raw = e && e.postData ? e.postData.contents : '';
    var data = JSON.parse(raw);
    if (!data || !data.members || !data.tasks) return out_({ ok: false, error: '일정보드 데이터 형식이 아님' });
    writeData_(raw);
    return out_({ ok: true, updatedAt: data.updatedAt || '' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}
