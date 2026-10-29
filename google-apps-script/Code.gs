// 키핀터치 캠페인 신청서 → 구글시트 저장
// 가이드 페이지(a~g-type.html)의 신청 폼이 이 웹 앱으로 JSON을 POST합니다.
// 모든 타입·고료의 신청을 한 탭("신청 목록")에 모아 저장하고, 타입·고료는 열로 구분합니다.

const SPREADSHEET_ID = '1WduulbFxVqsjjc5hMV4CIuHd0vk9YqvHop85DeiRBR0';
const SHEET_NAME = '신청 목록';

const HEADERS = [
  '제출일시', '타입', '고료', '배정 제품', '이름', '인스타그램',
  '휴대폰', '이메일', '우편번호', '주소 (상세주소 포함)', '요청사항'
];

// 앞자리 0이 숫자 변환으로 사라지지 않도록 텍스트로 고정 (' 접두사는 셀에 표시되지 않음)
function asText_(value) {
  const v = String(value == null ? '' : value).trim();
  return v ? "'" + v : '';
}

// 배송지 주소 + 상세 주소를 한 칸에 합침
function joinAddress_(data) {
  const road = String(data.roadAddress || '').trim();
  const detail = String(data.detailAddress || '').trim();
  const joined = [road, detail].filter(Boolean).join(' ');
  return joined || String(data.address || '').trim();
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getSheet_();

    sheet.appendRow([
      new Date(),
      data.tier || '',
      data.price || '',
      data.productLine || '',
      data.name || '',
      data.instagram || '',
      asText_(data.phone),
      data.email || '',
      asText_(data.zipcode),
      joinAddress_(data),
      data.note || ''
    ]);

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function getSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME, 0);
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    // 휴대폰(G)·우편번호(I) 열 전체를 텍스트 서식으로
    sheet.getRange('G:G').setNumberFormat('@');
    sheet.getRange('I:I').setNumberFormat('@');
  }
  return sheet;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
