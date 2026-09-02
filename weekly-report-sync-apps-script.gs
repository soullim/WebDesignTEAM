/**
 * 웹디자인팀 일정보드 — "주간보고 연동" 백엔드 (Google Apps Script)
 *
 * "2026년_IT부문 주간업무 보고" 구글시트에서 웹기획팀 행에 적힌 리뷰일·오픈일을
 * 서버 사이드에서 읽어와 후보 목록(JSON)으로 내보냅니다.
 * 일정보드는 이 주소를 주기적으로 조회해 "주간보고 신규" 카드에 후보를 띄우고,
 * 뚜앵님이 직접 어떤 타임라인에 반영할지 확인한 뒤에만 실제로 반영합니다.
 * (시트 문장이 자유형식이라 자동 매칭은 하지 않고, 사람이 확인하는 단계를 항상 거칩니다.)
 *
 * ── 설정 방법 (최초 1회, 약 5분) ─────────────────────────────
 * 1. https://script.google.com/home/start 에서 새 프로젝트 생성
 *    (또는 "2026년_IT부문 주간업무 보고" 시트 메뉴 → 확장 프로그램 → Apps Script)
 * 2. 편집기에 기본으로 있는 코드를 지우고 이 파일 내용 전체를 붙여넣기 → 저장(💾)
 * 3. 아래 SHEET_ID가 실제 시트 ID와 다르면 바꿔주세요.
 *    (시트 URL의 /d/ 와 /edit 사이 부분이 ID예요. 지금은 이미 맞게 채워져 있어요.)
 * 4. 우측 상단 [배포] → [새 배포] 클릭
 *    - 유형 선택(⚙ 아이콘): "웹 앱"
 *    - 설명: 아무거나 (예: 주간보고 연동 v1)
 *    - 다음 사용자 인증 정보로 실행: "나"
 *    - 액세스 권한이 있는 사용자: "모든 사용자"  ← 중요! (시트 자체는 비공개로 남아요)
 * 5. [배포] 클릭 → 권한 승인 창이 뜨면 본인 계정 선택 → "고급" → "이동" → 허용
 * 6. 발급된 "웹 앱 URL" (https://script.google.com/macros/s/...../exec) 을 복사해서
 *    index.html의 WR_GAS_URL 에 붙여넣으면 완료
 *
 * ── 이후 코드를 수정했을 때 ──────────────────────────────────
 * [배포] → [배포 관리] → 연필(✏) 아이콘 → 버전: "새 버전" → [배포]
 * (새 배포를 만들면 URL이 바뀌니 반드시 "배포 관리"에서 수정할 것)
 *
 * ── 다른 팀도 같이 보고 싶다면 ───────────────────────────────
 * 아래 TEAM_TARGETS 배열에 팀 이름을 추가하면 돼요. (예: '웹디자인팀' 추가)
 * ──────────────────────────────────────────────────────────
 */

var SHEET_ID = '1e_gzTJ3J3bOFQV76CTug6gAjJx20IIIsH61QO5Jn1ww';
var TEAM_TARGETS = ['웹기획팀'];
var REPORT_COL = 3; // 0-index — "유관부서 (담당)" 자리에 실제 상세 리포트가 들어있는 열

function doGet(e) {
  var out;
  try {
    out = { ok: true, updatedAt: new Date().toISOString(), items: collectItems() };
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function collectItems() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheets = ss.getSheets();
  var items = [];
  sheets.forEach(function (sh) {
    var data = sh.getDataRange().getValues();
    var year = null, headerRow = -1;
    for (var r = 0; r < Math.min(data.length, 12); r++) {
      var rowText = data[r].join(' ');
      var ym = rowText.match(/(\d{4})\s*년/);
      if (ym && !year) year = +ym[1];
      if (String(data[r][0]).trim() === '팀') { headerRow = r; break; }
    }
    if (headerRow < 0) return;
    if (!year) year = new Date().getFullYear();
    for (var r2 = headerRow + 1; r2 < data.length; r2++) {
      var team = String(data[r2][0] || '').trim();
      if (TEAM_TARGETS.indexOf(team) < 0) continue;
      var body = String(data[r2][REPORT_COL] || '');
      if (!body.trim()) continue;
      extractItemsFromText(body, year, sh.getName(), team).forEach(function (it) { items.push(it); });
    }
  });
  return items;
}

function pad2(n) { return (n < 10 ? '0' : '') + n; }

/* 한 팀의 상세 리포트 문단에서 "날짜가 포함된 줄"을 찾아 리뷰/오픈 후보로 변환 */
function extractItemsFromText(text, year, week, team) {
  var items = [];
  var lines = text.split('\n');
  lines.forEach(function (raw) {
    var line = raw.trim();
    if (!line) return;
    var dateRe = /(\d{1,2})\s*[\/월]\s*(\d{1,2})/g;
    var found = [], m;
    while ((m = dateRe.exec(line))) {
      var mo = +m[1], da = +m[2];
      if (mo >= 1 && mo <= 12 && da >= 1 && da <= 31) found.push(mo + '-' + da);
    }
    if (!found.length) return;
    var kind = /리뷰/.test(line) ? '리뷰' : (/오픈|배포|게시|정식\s*서비스/.test(line) ? '오픈' : '기타');
    var status = /완료/.test(line) ? '완료' : (/예정/.test(line) ? '예정' : '');
    var title = line
      .replace(/^[-*·•ㄴ]\s*/, '')
      .replace(/\(\s*\d{1,2}\s*[\/월]\s*\d{1,2}[^)]*\)/g, '')
      .replace(/\s{2,}/g, ' ')
      .trim();
    // 같은 줄에 날짜가 여러 개면(예: "완료(8/28), 9/1 리뷰 예정") 각각 후보로 제출 — 확인은 사람이 함
    var seen = {};
    found.forEach(function (md) {
      if (seen[md]) return; seen[md] = 1;
      var p = md.split('-');
      items.push({
        team: team,
        week: week,
        date: year + '-' + pad2(+p[0]) + '-' + pad2(+p[1]),
        kind: kind,
        status: status,
        title: title,
        raw: line
      });
    });
  });
  return items;
}
