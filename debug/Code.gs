// 버그 즐겨찾기 (구글 시트 + Apps Script)
//
// - 사이트 추가: 화면의 ＋ 버튼(카테고리는 드롭다운에서 고르기), 또는 시트에 직접 한 줄 적기
// - 관리 모드(비밀번호 필요): 카테고리 추가·이름 바꾸기·삭제, 사이트 삭제
// ★ 아래 비밀번호를 나만 아는 것으로 꼭 바꾸세요.
const ADMIN_PASSWORD = '1234';

// 시트 탭 이름 (없으면 자동으로 만들어져요)
// '즐겨찾기' 탭: A 카테고리 | B 사이트 이름 | C 주소 | D 메모
// '카테고리' 탭: A 카테고리 이름 (적힌 순서대로 화면에 나와요)
const SITE_SHEET = '즐겨찾기';
const CATEGORY_SHEET = '카테고리';
const DEFAULT_CATEGORIES = ['업무', '수업', '연수']; // 카테고리 탭을 처음 만들 때 넣어 둘 이름

// ---------------------------------------------------------------
// 웹 앱 열기
// ---------------------------------------------------------------
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('나의 즐겨찾기')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ---------------------------------------------------------------
// 누구나: 카테고리와 사이트 목록 읽기
// ---------------------------------------------------------------
function getData() {
  const categories = readCategories_();
  const sites = readSites_();
  // 시트에 직접 적은, 아직 등록하지 않은 카테고리도 화면에 보이게
  sites.forEach(function (s) {
    if (categories.indexOf(s.category) < 0) categories.push(s.category);
  });
  return { categories: categories, sites: sites };
}

// ---------------------------------------------------------------
// 누구나: 사이트 추가 (화면의 ＋ 버튼)
// ---------------------------------------------------------------
function addSite(category, name, url, memo) {
  category = String(category || '').trim().slice(0, 20) || '기타';
  name = clean_(name, 40, '사이트 이름');
  url = clean_(url, 500, '주소');
  memo = String(memo || '').trim().slice(0, 100); // 메모는 비워도 돼요
  if (/\s/.test(url)) throw new Error('주소에는 띄어쓰기를 넣을 수 없어요.');
  if (!/^https?:\/\//i.test(url)) url = 'https://' + url; // naver.com → https://naver.com

  withLock_(function () {
    appendTextRow_(getSheet_(SITE_SHEET), [category, name, url, memo]);
  });
  return true;
}

// ---------------------------------------------------------------
// 관리 모드: 비밀번호 확인
// ---------------------------------------------------------------
function checkAdmin(password) {
  checkPassword_(password);
  return true;
}

// ---------------------------------------------------------------
// 관리 모드: 사이트 삭제
// ---------------------------------------------------------------
function deleteSite(password, row, url) {
  checkPassword_(password);
  withLock_(function () {
    const sheet = getSheet_(SITE_SHEET);
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
  });
  return true;
}

// ---------------------------------------------------------------
// 관리 모드: 카테고리 추가
// ---------------------------------------------------------------
function addCategory(password, name) {
  checkPassword_(password);
  name = clean_(name, 20, '카테고리 이름');
  withLock_(function () {
    if (getData().categories.indexOf(name) >= 0) throw new Error('이미 있는 카테고리예요.');
    appendTextRow_(getSheet_(CATEGORY_SHEET), [name]);
  });
  return true;
}

// ---------------------------------------------------------------
// 관리 모드: 카테고리 이름 바꾸기 (그 카테고리의 사이트도 함께 바뀌어요)
// ---------------------------------------------------------------
function renameCategory(password, oldName, newName) {
  checkPassword_(password);
  newName = clean_(newName, 20, '카테고리 이름');
  if (newName === oldName) return true;
  withLock_(function () {
    if (getData().categories.indexOf(newName) >= 0) throw new Error('이미 있는 카테고리예요.');
    const changed = replaceInColumnA_(getSheet_(CATEGORY_SHEET), oldName, newName, '');
    if (!changed) appendTextRow_(getSheet_(CATEGORY_SHEET), [newName]); // 시트에만 적혀 있던 카테고리면 새로 등록
    replaceInColumnA_(getSheet_(SITE_SHEET), oldName, newName, '기타');
  });
  return true;
}

// ---------------------------------------------------------------
// 관리 모드: 카테고리 삭제 (그 카테고리의 사이트는 '기타'로 옮겨져요)
// ---------------------------------------------------------------
function deleteCategory(password, name) {
  checkPassword_(password);
  withLock_(function () {
    const inUse = readSites_().some(function (s) { return s.category === name; });
    if (name === '기타' && inUse) throw new Error("'기타'에 사이트가 남아 있어서 지울 수 없어요.");

    const sheet = getSheet_(CATEGORY_SHEET);
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const names = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
      for (let i = names.length - 1; i >= 0; i--) { // 아래 줄부터 지워야 줄 번호가 안 밀려요
        if (names[i][0].trim() === name) sheet.deleteRow(i + 2);
      }
    }
    if (inUse) replaceInColumnA_(getSheet_(SITE_SHEET), name, '기타', '기타');
  });
  return true;
}

// ---------------------------------------------------------------
// 도우미 함수 (이름이 _ 로 끝나면 화면에서 직접 부를 수 없어요)
// ---------------------------------------------------------------

// 시트 탭 가져오기 (없으면 제목 줄과 함께 새로 만들기)
function getSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (name === SITE_SHEET) {
      sheet.appendRow(['카테고리', '사이트 이름', '주소', '메모']);
    } else {
      sheet.appendRow(['카테고리']);
      DEFAULT_CATEGORIES.forEach(function (c) { sheet.appendRow([c]); });
    }
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// '카테고리' 탭의 이름들 (빈 칸과 겹치는 이름은 빼고)
function readCategories_() {
  const sheet = getSheet_(CATEGORY_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const names = [];
  sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues().forEach(function (r) {
    const name = r[0].trim();
    if (name && names.indexOf(name) < 0) names.push(name);
  });
  return names;
}

// '즐겨찾기' 탭의 사이트들 (시트에 직접 적은 줄도 똑같이 읽어요)
function readSites_() {
  const sheet = getSheet_(SITE_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const sites = [];
  sheet.getRange(2, 1, lastRow - 1, 4).getDisplayValues().forEach(function (r, i) {
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

// A열에서 from 을 모두 to 로 바꾸기 (빈 칸은 blankAs 로 보고 비교), 바꾼 개수 돌려주기
function replaceInColumnA_(sheet, from, to, blankAs) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const range = sheet.getRange(2, 1, lastRow - 1, 1);
  const values = range.getDisplayValues();
  let count = 0;
  values.forEach(function (r) {
    if ((r[0].trim() || blankAs) === from) { r[0] = to; count++; }
  });
  if (count) {
    range.setNumberFormat('@');
    range.setValues(values);
  }
  return count;
}

// 맨 아래에 한 줄 쓰기 (=, + 로 시작해도 수식이 되지 않게 글자 그대로)
function appendTextRow_(sheet, values) {
  const range = sheet.getRange(sheet.getLastRow() + 1, 1, 1, values.length);
  range.setNumberFormat('@');
  range.setValues([values]);
}

// 여러 명이 동시에 고쳐도 겹치지 않게 한 명씩 처리
function withLock_(work) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    return work();
  } finally {
    lock.releaseLock();
  }
}

function checkPassword_(password) {
  if (password !== ADMIN_PASSWORD) throw new Error('비밀번호가 맞지 않아요.');
}

// 입력값 정리: 앞뒤 빈칸 없애고, 비었거나 너무 길면 알려주기
function clean_(value, max, label) {
  const text = String(value || '').trim();
  if (!text) throw new Error(label + '을(를) 입력해 주세요.');
  if (text.length > max) throw new Error(label + '은(는) ' + max + '자까지 쓸 수 있어요.');
  return text;
}
