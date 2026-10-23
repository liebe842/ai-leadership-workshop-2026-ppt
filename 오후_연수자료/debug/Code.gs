// 버그 게시판 (구글 시트 + Apps Script)
//
// - 글쓰기, 글 읽기: 누구나
// - 글 삭제: 선생님만 (비밀번호 필요)
// ★ 아래 비밀번호를 나만 아는 것으로 꼭 바꾸세요.
const TEACHER_PASSWORD = '1234';

// 글이 저장될 시트 이름 (없으면 자동으로 만들어져요)
const SHEET_NAME = '게시판';

// ---------------------------------------------------------------
// 웹 앱 열기
// ---------------------------------------------------------------
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('우리들의 게시판')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ---------------------------------------------------------------
// 누구나: 글쓰기
// ---------------------------------------------------------------
function addPost(name, title, content) {
  name = clean_(name, 20, '이름');
  title = clean_(title, 50, '제목');
  content = clean_(content, 1000, '내용');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000); // 여러 명이 동시에 써도 겹치지 않게 한 명씩 처리
  try {
    const sheet = getSheet_();
    const id = String(Date.now()); // 글마다 다른 번호
    const time = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
    const row = sheet.getLastRow() + 1;
    const range = sheet.getRange(row, 1, 1, 5);
    range.setNumberFormat('@'); // 글자 그대로 저장 (=, + 로 시작해도 수식이 되지 않게)
    range.setValues([[id, time, name, title, content]]);
  } finally {
    lock.releaseLock();
  }
  return true;
}

// ---------------------------------------------------------------
// 누구나: 글 목록 읽기 (최신 글이 위로)
// ---------------------------------------------------------------
function getPosts() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  return values
    .map(function (r) {
      return {
        id: String(r[0]),
        time: String(r[1]),
        name: String(r[2]),
        title: String(r[3]),
        content: String(r[4])
      };
    })
    .reverse();
}

// ---------------------------------------------------------------
// 선생님만: 글 삭제
// ---------------------------------------------------------------
function deletePost(password, id) {
  if (password !== TEACHER_PASSWORD) {
    throw new Error('비밀번호가 맞지 않아요.');
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error('글을 찾을 수 없어요.');
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(id)) {
        sheet.deleteRow(i + 2);
        return true;
      }
    }
    throw new Error('글을 찾을 수 없어요. 이미 삭제되었을 수 있어요.');
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------
// 도우미 함수 (이름이 _ 로 끝나면 화면에서 직접 부를 수 없어요)
// ---------------------------------------------------------------

// 게시판 시트 가져오기 (없으면 제목 줄과 함께 새로 만들기)
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['번호', '작성 시간', '이름', '제목', '내용']);
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
