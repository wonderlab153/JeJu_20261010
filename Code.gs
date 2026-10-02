/**
 * EXPLORE 제주 · 관찰 기록 저장소 (Google Apps Script)
 * ─────────────────────────────────────────────
 * 하는 일
 *  - 학생이 보낸 관찰 기록을 이 구글 시트의 'records' 탭에 한 줄씩 저장(같은 지점·같은 이름이면 덮어쓰기)
 *  - 사진은 내 구글 드라이브의 'EXPLORE제주_사진' 폴더에 비공개로 저장
 *  - 마스터 페이지가 비밀번호로 기록 목록·사진을 읽고, 기록을 삭제
 *
 * 설치 (처음 한 번)
 *  1) 구글 드라이브에서 새 구글 시트를 만들고 이름을 'EXPLORE제주 기록'으로 바꿉니다.
 *  2) 시트 메뉴 [확장 프로그램] → [Apps Script]를 열고, 기본 코드를 지운 뒤 이 파일 전체를 붙여 넣고 저장합니다.
 *  3) 위쪽 함수 선택에서 setup 을 고르고 [실행] → 권한 허용(내 계정). 'records' 탭과 사진 폴더가 생깁니다.
 *  4) [배포] → [새 배포] → 유형 '웹 앱'
 *       - 실행 계정: 나
 *       - 액세스 권한: 모든 사용자(Anyone)
 *     [배포] 후 나오는 '웹 앱 URL'(…/exec)을 복사해 index.html의 SCRIPT_URL에 붙여 넣습니다.
 *  5) 코드를 고친 뒤에는 [배포] → [배포 관리] → 연필 → 버전 '새 버전'으로 다시 배포해야 반영됩니다(URL은 그대로).
 *
 * 보안 메모
 *  - 마스터 비밀번호는 아래 MASTER_PW 한 곳에만 있습니다(HTML에는 없음). 필요하면 바꾸고 다시 배포하세요.
 *  - 사진 파일은 공유 링크를 만들지 않아 내 계정만 볼 수 있습니다.
 */

const MASTER_PW   = '3004';
const SHEET_NAME  = 'records';
const FOLDER_NAME = 'EXPLORE제주_사진';
const SITES = ['G1','G2','G3','G4','G5','G6','G7','G8','B1','B2'];
const HEAD  = ['recordId','site','siteName','name','see','q1','q2','q3','why','faith','photoIds','updatedAt','receivedAt'];

/** 처음 한 번 실행: 시트 탭과 사진 폴더 준비 */
function setup() { getSheet_(); getFolder_(); }

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) { sh = ss.insertSheet(SHEET_NAME); sh.appendRow(HEAD); sh.setFrozenRows(1); }
  return sh;
}
function getFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder(FOLDER_NAME);
  props.setProperty('FOLDER_ID', f.getId());
  return f;
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function str_(v, max) { let s = String(v == null ? '' : v).slice(0, max); if (/^[=+\-@]/.test(s)) s = "'" + s; return s; } // 수식 주입 방지
function findRow_(sh, recordId) {
  const last = sh.getLastRow(); if (last < 2) return -1;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (ids[i][0] === recordId) return i + 2;
  return -1;
}

/** 읽기: ?action=list&pw=… / ?action=photo&id=…&pw=… */
function doGet(e) {
  const p = e.parameter || {};
  if (p.action === 'ping') return json_({ ok: true });
  if (p.pw !== MASTER_PW) return json_({ ok: false, error: 'password' });
  try {
    if (p.action === 'list') {
      const v = getSheet_().getDataRange().getValues(); const h = v.shift();
      const records = v.filter(r => r[0]).map(r => { const o = {}; h.forEach((k, i) => o[k] = r[i] instanceof Date ? r[i].toISOString() : r[i]); return o; });
      return json_({ ok: true, records });
    }
    if (p.action === 'photo') {
      const f = DriveApp.getFileById(p.id);
      const folderId = getFolder_().getId(); let inFolder = false; const ps = f.getParents();
      while (ps.hasNext()) if (ps.next().getId() === folderId) inFolder = true;
      if (!inFolder) return json_({ ok: false, error: 'notfound' });
      const b = f.getBlob();
      return json_({ ok: true, data: 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes()) });
    }
    return json_({ ok: false, error: 'action' });
  } catch (err) { return json_({ ok: false, error: String(err) }); }
}

/** 쓰기: {action:'save', record, photos, photosMode} / {action:'delete', recordId, pw} */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const b = JSON.parse(e.postData.contents);
    if (b.action === 'save')   return json_(save_(b));
    if (b.action === 'delete') { if (b.pw !== MASTER_PW) return json_({ ok: false, error: 'password' }); return json_(del_(b.recordId)); }
    return json_({ ok: false, error: 'action' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally { try { lock.releaseLock(); } catch (e) {} }
}

function save_(b) {
  const r = b.record || {};
  const site = String(r.site || ''); const name = String(r.name || '').trim().slice(0, 20);
  if (SITES.indexOf(site) < 0 || !name) return { ok: false, error: 'invalid' };
  const recordId = site + '|' + name;
  const sh = getSheet_(); const row = findRow_(sh, recordId);
  let photoIds = row > 0 ? String(sh.getRange(row, 11).getValue() || '') : '';

  if (b.photosMode === 'replace') {
    // 이전 사진은 휴지통으로, 새 사진 저장(최대 3장, JPEG/PNG/WEBP, 장당 3MB 이하)
    photoIds.split(',').filter(String).forEach(id => { try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {} });
    const folder = getFolder_(); const ids = [];
    (Array.isArray(b.photos) ? b.photos : []).slice(0, 3).forEach((d, i) => {
      const m = /^data:(image\/(?:jpeg|png|webp));base64,(.+)$/.exec(String(d)); if (!m) return;
      const bytes = Utilities.base64Decode(m[2]); if (bytes.length > 3 * 1024 * 1024) return;
      const ext = m[1].split('/')[1].replace('jpeg', 'jpg');
      const file = folder.createFile(Utilities.newBlob(bytes, m[1], `${site}_${name}_${i + 1}.${ext}`));
      ids.push(file.getId());
    });
    photoIds = ids.join(',');
  }

  const values = [recordId, site, str_(r.siteName, 60), str_(name, 20), str_(r.see, 1000), str_(r.q1, 200), str_(r.q2, 200), str_(r.q3, 200),
    str_(r.why, 1000), str_(r.faith, 1000), photoIds, str_(r.updatedAt, 40), new Date()];
  if (row > 0) sh.getRange(row, 1, 1, HEAD.length).setValues([values]); else sh.appendRow(values);
  return { ok: true, recordId, photos: photoIds ? photoIds.split(',').length : 0 };
}

function del_(recordId) {
  const sh = getSheet_(); const row = findRow_(sh, String(recordId || ''));
  if (row < 0) return { ok: true, deleted: false };
  String(sh.getRange(row, 11).getValue() || '').split(',').filter(String).forEach(id => { try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {} });
  sh.deleteRow(row);
  return { ok: true, deleted: true };
}
