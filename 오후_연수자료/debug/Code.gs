// 버그 즐겨찾기 (구글 시트 + Apps Script)
//
// - 사이트 추가: 화면의 ＋ 버튼, 또는 시트에 직접 한 줄 적기
// - 사이트 삭제: 나만 (비밀번호 필요)
// ★ 아래 비밀번호를 나만 아는 것으로 꼭 바꾸세요.
const ADMIN_PASSWORD = '1234';

// 사이트가 저장될 시트 이름 (없으면 자동으로 만들어져요)
// 시트 구성: A 카테고리 | B 사이트 이름 | C 주소 | D 메모
const SHEET_NAME = '즐겨찾기';

// ---------------------------------------------------------------
// 웹 앱 열기
// ---------------------------------------------------------------
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('나의 즐겨찾기')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ---------------------------------------------------------------
// 누구나: 사이트 추가 (화면의 ＋ 버튼)
// ---------------------------------------------------------------
function addSite(category, name, url, memo) {
  category = String(category || '').trim().slice(0, 20) || '기타'; // 비우면 '기타'
  name = clean_(name, 40, '사이트 이름');
  url = clean_(url, 500, '주소');
  memo = String(memo || '').trim().slice(0, 100); // 메모는 비워도 돼요
  if (/\s/.test(url)) throw new Error('주소에는 띄어쓰기를 넣을 수 없어요.');
  if (!/^https?:\/\//i.test(url)) url = 'https://' + url; // naver.com → https://naver.com

  const lock = LockService.getScriptLock();
  lock.waitLock(10000); // 여러 명이 동시에 저장해도 겹치지 않게 한 명씩 처리
  try {
    const sheet = getSheet_();
    const row = sheet.getLastRow() + 1;
    const range = sheet.getRange(row, 1, 1, 4);
    range.setNumberFormat('@'); // 글자 그대로 저장 (=, + 로 시작해도 수식이 되지 않게)
    range.setValues([[category, name, url, memo]]);
  } finally {
    lock.releaseLock();
  }
  return true;
}

// ---------------------------------------------------------------
// 누구나: 사이트 목록 읽기 (시트에 적힌 순서 그대로)
// 시트에 직접 적은 줄도 똑같이 읽어요.
// ---------------------------------------------------------------
function getSites() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, 4).getDisplayValues();
  const sites = [];
  values.forEach(function (r, i) {
    const url = r[2].trim();
    if (!url) return; // 주소가 없는 줄은 건너뛰기
    sites.push({
      row: i + 2, // 시트의 몇 번째 줄인지 (삭제할 때 써요)
      category: r[0].trim() || '기타',
      name: r[1].trim() || url,
      url: url,
      memo: r[3].trim()
    });
  });
  return sites;
}

// ---------------------------------------------------------------
// 나만: 사이트 삭제
// ---------------------------------------------------------------
function deleteSite(password, row, url) {
  if (password !== ADMIN_PASSWORD) {
    throw new Error('비밀번호가 맞지 않아요.');
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error('사이트를 찾을 수 없어요.');
    const urls = sheet.getRange(2, 3, lastRow - 1, 1).getDisplayValues();

    // 화면을 연 뒤에 시트 줄이 바뀌었을 수 있어서, 주소가 같은지 한 번 더 확인해요
    let target = Number(row) - 2;
    if (!urls[target] || urls[target][0].trim() !== url) {
      target = -1;
      for (let i = 0; i < urls.length; i++) {
        if (urls[i][0].trim() === url) { target = i; break; }
      }
    }
    if (target < 0) throw new Error('사이트를 찾을 수 없어요. 이미 삭제되었을 수 있어요.');
    sheet.deleteRow(target + 2);
    return true;
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------
// 도우미 함수 (이름이 _ 로 끝나면 화면에서 직접 부를 수 없어요)
// ---------------------------------------------------------------

// 즐겨찾기 시트 가져오기 (없으면 제목 줄과 함께 새로 만들기)
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['카테고리', '사이트 이름', '주소', '메모']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// 입력값 정리: 앞뒤 빈칸 없애고, 비었거나 너무 길면 알려주기
function clean_(value, max, label) {
  const text = String(value || '').trim();
  if (!text) throw new Error(label + '을(를) 입력해 주세요.');
  if (text.length > max) throw new Error(label + '은(는) ' + max + '자까지 쓸 수 있어요.');
  return text;
}
