/**
 * 파주 유적지도 PWA - Google Sheets 신규 구역 + 공유 사진 연동용 Apps Script v8.16
 *
 * 핵심 기능:
 * - 승인여부 기능 없음
 * - 앱에서 전달한 Google Sheets 문서 ID의 첫 번째 시트에 저장
 * - 소재지 / 메모 분리 저장
 * - 신규 구역 사진을 Google Drive 폴더에 저장
 * - Google Sheets에는 사진 URL/파일명/JSON 정보 저장
 * - 다른 사용자는 앱에서 "공유 신규 구역 불러오기"로 사진을 확인하고 다운로드 가능
 *
 * 배포:
 * 1. 이 코드 전체를 기존 Apps Script 프로젝트에 붙여넣기
 * 2. 저장
 * 3. 배포 → 배포 관리 → 기존 웹앱 배포 수정 → 버전: 새 버전
 * 4. 실행 사용자: 나 / 액세스 권한: 모든 사용자
 */

const SHEET_NAME = ''; // 비워두면 첫 번째 시트 탭에 씁니다.
const PHOTO_FOLDER_NAME = '파주_유적지도_신규구역_공유사진';
const HEADERS = [
  'id', 'updatedAt', 'createdAt', '등록자', '번호', '이름', '법정동', '대상구분',
  '소재지', '메모', '지도검색주소', 'lat', 'lng', '세부내용', 'sourceApp',
  'photoJson', 'photoUrls', 'photoNames', 'photoUpdatedAt'
];

function doGet(e) {
  const p = e.parameter || {};
  const callback = safeCallback_(p.callback || 'callback');
  try {
    const action = p.action || 'list';
    if (action === 'ping') return jsonp_(callback, {ok:true, message:'pong', version:'v8.16'});
    if (action === 'upsert') return jsonp_(callback, {ok:true, item:upsertItem_(p), spreadsheetUrl:getSpreadsheet_(p).getUrl()});
    if (action === 'list') return jsonp_(callback, {ok:true, items:listItems_(p), spreadsheetUrl:getSpreadsheet_(p).getUrl()});
    if (action === 'deletePhoto') return jsonp_(callback, deletePhoto_(p));
    if (action === 'downloadPhotoZipDirect') return createPhotoZipDownloadHtml_(p);
    if (action === 'createPhotoZip') return jsonp_(callback, {ok:true, zip:createPhotoZip_(p)});
    return jsonp_(callback, {ok:false, error:'Unknown action: ' + action});
  } catch (err) {
    return jsonp_(callback, {ok:false, error:String(err && err.message ? err.message : err)});
  }
}

function doPost(e) {
  const p = e.parameter || {};
  const token = safePostToken_(p.callbackToken || '');
  try {
    const action = p.action || '';
    if (action === 'uploadPhoto') {
      const result = uploadPhoto_(p);
      return postMessageHtml_(token, {ok:true, photo:result.photo, item:result.item, spreadsheetUrl:getSpreadsheet_(p).getUrl()});
    }
    return postMessageHtml_(token, {ok:false, error:'Unknown POST action: ' + action});
  } catch (err) {
    return postMessageHtml_(token, {ok:false, error:String(err && err.message ? err.message : err)});
  }
}

function safeCallback_(name) {
  name = String(name || 'callback');
  if (!/^[A-Za-z0-9_.$]+$/.test(name)) return 'callback';
  return name;
}
function safePostToken_(token) {
  token = String(token || '');
  return token.replace(/[^A-Za-z0-9_.-]/g, '').slice(0, 120);
}
function jsonp_(callback, obj) {
  return ContentService.createTextOutput(callback + '(' + JSON.stringify(obj) + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
}
function postMessageHtml_(token, obj) {
  const payload = JSON.stringify({pajuUploadToken:token, result:obj}).replace(/</g, '\\u003c');
  const html = '<!doctype html><meta charset="utf-8"><script>window.parent.postMessage(' + payload + ', "*");</script>';
  return HtmlService.createHtmlOutput(html);
}
function cleanSheetId_(value) {
  value = String(value || '').trim();
  const m = value.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (m) return m[1];
  return value;
}
function getSpreadsheet_(p) {
  const id = cleanSheetId_((p && p.sheetId) || '');
  if (id) return SpreadsheetApp.openById(id);
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  throw new Error('Google Sheets 문서 ID가 없습니다. 앱의 Google Sheets 문서 URL/ID 칸에 시트 URL을 입력해 주세요.');
}
function getSheet_(p) {
  const ss = getSpreadsheet_(p);
  let sheet;
  const name = String(SHEET_NAME || '').trim();
  if (name) sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  else sheet = ss.getSheets()[0] || ss.insertSheet('신규구역');

  const lastCol = Math.max(sheet.getLastColumn(), HEADERS.length);
  const first = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const isEmptyHeader = first.join('') === '';
  const hasWrongHeader = first[0] !== 'id';
  const headerChanged = HEADERS.some((h, i) => first[i] !== h) || first.length < HEADERS.length;

  if (isEmptyHeader || hasWrongHeader) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  } else if (headerChanged) {
    const lastRow = Math.max(sheet.getLastRow(), 1);
    const oldHeaders = first;
    const oldValues = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
    const newValues = oldValues.map(row => HEADERS.map(h => {
      const idx = oldHeaders.indexOf(h);
      if (idx >= 0) return row[idx];
      return '';
    }));
    sheet.clearContents();
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    if (newValues.length) sheet.getRange(2, 1, newValues.length, HEADERS.length).setValues(newValues);
    sheet.setFrozenRows(1);
  }
  return sheet;
}
function headerIndexMap_() {
  const m = {};
  HEADERS.forEach((h, i) => m[h] = i + 1);
  return m;
}
function readExistingRowObj_(sheet, targetRow) {
  if (!targetRow || targetRow <= 1) return {};
  const values = sheet.getRange(targetRow, 1, 1, HEADERS.length).getValues()[0];
  const obj = {};
  HEADERS.forEach((h, i) => obj[h] = values[i]);
  return obj;
}
function findRowById_(sheet, id) {
  id = String(id || '');
  if (!id) return -1;
  const lastRow = Math.max(sheet.getLastRow(), 1);
  if (lastRow <= 1) return -1;
  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === id) return i + 2;
  }
  return -1;
}
function rowObjFromParams_(p, existing) {
  const now = new Date();
  existing = existing || {};
  const lat = Number(p.lat !== undefined ? p.lat : existing.lat);
  const lng = Number(p.lng !== undefined ? p.lng : existing.lng);
  const obj = {
    id: String(p.id || existing.id || ('zone-' + now.getTime())),
    updatedAt: now,
    createdAt: existing.createdAt || now,
    '등록자': p.user !== undefined ? p.user : (existing['등록자'] || ''),
    '번호': p.num !== undefined ? p.num : (existing['번호'] || ''),
    '이름': p.name !== undefined ? p.name : (existing['이름'] || '공유 신규 구역'),
    '법정동': p.town !== undefined ? p.town : (existing['법정동'] || ''),
    '대상구분': p.target !== undefined ? p.target : (existing['대상구분'] || p.town || ''),
    '소재지': p.address !== undefined ? p.address : (existing['소재지'] || ''),
    '메모': p.memo !== undefined ? p.memo : (existing['메모'] || ''),
    '지도검색주소': p.search !== undefined ? p.search : (existing['지도검색주소'] || ''),
    lat: lat,
    lng: lng,
    '세부내용': p.detail !== undefined ? p.detail : (p.note !== undefined ? p.note : (existing['세부내용'] || '')),
    sourceApp: p.sourceApp || existing.sourceApp || 'paju-pwa-v8.11',
    photoJson: p.photoJson !== undefined ? p.photoJson : (existing.photoJson || ''),
    photoUrls: p.photoUrls !== undefined ? p.photoUrls : (existing.photoUrls || ''),
    photoNames: p.photoNames !== undefined ? p.photoNames : (existing.photoNames || ''),
    photoUpdatedAt: p.photoUpdatedAt !== undefined ? p.photoUpdatedAt : (existing.photoUpdatedAt || '')
  };
  if (!isFinite(obj.lat) || !isFinite(obj.lng)) throw new Error('lat/lng 값이 올바르지 않습니다.');
  return obj;
}
function upsertItem_(p) {
  const sheet = getSheet_(p);
  const id = String(p.id || ('zone-' + Date.now()));
  const targetRow = findRowById_(sheet, id);
  const existing = readExistingRowObj_(sheet, targetRow);
  const rowObj = rowObjFromParams_(Object.assign({}, p, {id:id}), existing);
  if (targetRow > 0) sheet.getRange(targetRow, 1, 1, HEADERS.length).setValues([HEADERS.map(h => rowObj[h])]);
  else sheet.appendRow(HEADERS.map(h => rowObj[h]));
  SpreadsheetApp.flush();
  return rowObj;
}
function listItems_(p) {
  const sheet = getSheet_(p);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  const values = sheet.getRange(1, 1, lastRow, HEADERS.length).getValues();
  const headers = values[0].map(String);
  const items = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    if (!obj.id || obj.lat === '' || obj.lng === '') continue;
    items.push({
      id: obj.id,
      updatedAt: obj.updatedAt,
      createdAt: obj.createdAt,
      '등록자': obj['등록자'],
      '번호': obj['번호'],
      '이름': obj['이름'],
      '법정동': obj['법정동'],
      '대상구분': obj['대상구분'],
      '소재지': obj['소재지'],
      '메모': obj['메모'],
      '지도검색주소': obj['지도검색주소'],
      lat: obj.lat,
      lng: obj.lng,
      '세부내용': obj['세부내용'],
      photoJson: obj.photoJson,
      photoUrls: obj.photoUrls,
      photoNames: obj.photoNames,
      photoUpdatedAt: obj.photoUpdatedAt
    });
  }
  return items;
}
function getPhotoFolder_() {
  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty('PAJU_PHOTO_FOLDER_ID');
  if (savedId) {
    try { return DriveApp.getFolderById(savedId); } catch(e) {}
  }
  const it = DriveApp.getFoldersByName(PHOTO_FOLDER_NAME);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(PHOTO_FOLDER_NAME);
  props.setProperty('PAJU_PHOTO_FOLDER_ID', folder.getId());
  return folder;
}
function safeFileName_(name) {
  name = String(name || 'photo.jpg').replace(/[\\/:*?"<>|]/g, '_').replace(/\s+/g, '_');
  if (!/\.(jpg|jpeg|png|webp)$/i.test(name)) name += '.jpg';
  return name.slice(0, 120);
}
function parsePhotoJson_(text) {
  if (!text) return [];
  try {
    const arr = JSON.parse(String(text));
    return Array.isArray(arr) ? arr : [];
  } catch(e) { return []; }
}
function uploadPhoto_(p) {
  const sheet = getSheet_(p);
  const id = String(p.id || '').trim();
  if (!id) throw new Error('사진을 연결할 신규 구역 id가 없습니다.');
  const dataUrl = String(p.dataUrl || '');
  const m = dataUrl.match(/^data:([^;]+);base64,(.+)$/);
  if (!m) throw new Error('사진 데이터 형식이 올바르지 않습니다.');
  const mimeType = String(p.mimeType || m[1] || 'image/jpeg');
  if (!/^image\//.test(mimeType)) throw new Error('사진 파일만 업로드할 수 있습니다.');
  const bytes = Utilities.base64Decode(m[2]);
  const now = new Date();
  const fileName = Utilities.formatDate(now, Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyyMMdd_HHmmss') + '_' + id.replace(/[^A-Za-z0-9가-힣_-]/g, '_').slice(0,40) + '_' + safeFileName_(p.fileName || 'photo.jpg');
  const folder = getPhotoFolder_();
  const blob = Utilities.newBlob(bytes, mimeType, fileName);
  const file = folder.createFile(blob);
  let shared = true;
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }
  catch(e) { shared = false; }
  const fileId = file.getId();
  const photo = {
    fileId: fileId,
    name: safeFileName_(p.fileName || fileName),
    mimeType: mimeType,
    size: bytes.length,
    uploadedAt: now,
    uploader: p.user || '',
    itemId: id,
    shared: shared,
    thumbUrl: 'https://drive.google.com/thumbnail?id=' + encodeURIComponent(fileId) + '&sz=w1000',
    url: 'https://drive.google.com/thumbnail?id=' + encodeURIComponent(fileId) + '&sz=w1600',
    downloadUrl: 'https://drive.google.com/uc?export=download&id=' + encodeURIComponent(fileId),
    viewUrl: file.getUrl()
  };
  let targetRow = findRowById_(sheet, id);
  let existing = readExistingRowObj_(sheet, targetRow);
  let rowObj;
  if (targetRow > 0) rowObj = rowObjFromParams_(p, existing);
  else {
    if (p.lat === undefined || p.lng === undefined) throw new Error('신규 행 생성을 위해 lat/lng 값이 필요합니다.');
    rowObj = rowObjFromParams_(p, {});
  }
  const photos = parsePhotoJson_(rowObj.photoJson);
  photos.push(photo);
  rowObj.photoJson = JSON.stringify(photos);
  rowObj.photoUrls = photos.map(function(x){ return x.thumbUrl || x.url || x.downloadUrl || ''; }).filter(String).join('|');
  rowObj.photoNames = photos.map(function(x){ return x.name || ''; }).join('|');
  rowObj.photoUpdatedAt = now;
  rowObj.updatedAt = now;
  if (targetRow > 0) sheet.getRange(targetRow, 1, 1, HEADERS.length).setValues([HEADERS.map(h => rowObj[h])]);
  else sheet.appendRow(HEADERS.map(h => rowObj[h]));
  SpreadsheetApp.flush();
  return {photo:photo, item:rowObj};
}


/**
 * v8.11 보강: 빈값 덮어쓰기 방지
 * 기존 행에 값이 있는데 앱에서 빈 문자열을 보내면 기존 값을 유지합니다.
 * 예전 앱 캐시가 소재지/메모/사진URL 등을 빈값으로 다시 저장하는 문제를 서버 쪽에서도 방지합니다.
 */
function isBlankV811_(v) {
  return v === undefined || v === null || String(v).trim() === '';
}
function keepExistingIfIncomingBlankV811_(incoming, existing) {
  return isBlankV811_(incoming) && !isBlankV811_(existing) ? existing : incoming;
}
function rowObjFromParams_(p, existing) {
  const now = new Date();
  existing = existing || {};
  const latValue = !isBlankV811_(p.lat) ? p.lat : existing.lat;
  const lngValue = !isBlankV811_(p.lng) ? p.lng : existing.lng;
  const lat = Number(latValue);
  const lng = Number(lngValue);

  const obj = {
    id: String(!isBlankV811_(p.id) ? p.id : (existing.id || ('zone-' + now.getTime()))),
    updatedAt: now,
    createdAt: existing.createdAt || now,
    '등록자': keepExistingIfIncomingBlankV811_(p.user, existing['등록자'] || ''),
    '번호': keepExistingIfIncomingBlankV811_(p.num, existing['번호'] || ''),
    '이름': keepExistingIfIncomingBlankV811_(p.name, existing['이름'] || '공유 신규 구역'),
    '법정동': keepExistingIfIncomingBlankV811_(p.town, existing['법정동'] || ''),
    '대상구분': keepExistingIfIncomingBlankV811_(p.target, existing['대상구분'] || p.town || ''),
    '소재지': keepExistingIfIncomingBlankV811_(p.address, existing['소재지'] || ''),
    '메모': keepExistingIfIncomingBlankV811_(p.memo, existing['메모'] || ''),
    '지도검색주소': keepExistingIfIncomingBlankV811_(p.search, existing['지도검색주소'] || ''),
    lat: lat,
    lng: lng,
    '세부내용': keepExistingIfIncomingBlankV811_(!isBlankV811_(p.detail) ? p.detail : p.note, existing['세부내용'] || ''),
    sourceApp: keepExistingIfIncomingBlankV811_(p.sourceApp, existing.sourceApp || 'paju-pwa-v8.16'),
    photoJson: keepExistingIfIncomingBlankV811_(p.photoJson, existing.photoJson || ''),
    photoUrls: keepExistingIfIncomingBlankV811_(p.photoUrls, existing.photoUrls || ''),
    photoNames: keepExistingIfIncomingBlankV811_(p.photoNames, existing.photoNames || ''),
    photoUpdatedAt: keepExistingIfIncomingBlankV811_(p.photoUpdatedAt, existing.photoUpdatedAt || '')
  };
  if (!isFinite(obj.lat) || !isFinite(obj.lng)) throw new Error('lat/lng 값이 올바르지 않습니다.');
  return obj;
}

function createPajuPhotoFolderNow() {
  const folder = getPhotoFolder_();
  let shared = true;
  try { folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }
  catch (e) { shared = false; }
  Logger.log('사진 폴더명: ' + folder.getName());
  Logger.log('사진 폴더 URL: ' + folder.getUrl());
  Logger.log('공유 설정 성공 여부: ' + shared);
  SpreadsheetApp.getUi().alert(
    '사진 폴더 확인/생성 완료\n\n' +
    '폴더명: ' + folder.getName() + '\n' +
    '공유 설정: ' + (shared ? '링크가 있는 사용자 보기 가능' : '자동 공유 설정 실패 - 수동 확인 필요') + '\n\n' +
    'Apps Script 실행 로그에서 폴더 URL을 확인할 수 있습니다.'
  );
}


/**
 * v8.16 보강: 여러 공유 사진을 Google Drive에서 ZIP으로 묶어 다운로드합니다.
 * 앱에서 action=createPhotoZip&fileIds=id1|id2 형태로 호출합니다.
 */
function safeZipFileNameV813_(name) {
  name = String(name || '파주_공유사진.zip').replace(/[\\/:*?"<>|]/g, '_').replace(/\s+/g, '_');
  if (!/\.zip$/i.test(name)) name += '.zip';
  return name.slice(0, 150);
}
function parseFileIdsV813_(value) {
  const seen = {};
  return String(value || '').split(/[|,\n\r\t ]+/).map(function(x){ return String(x || '').trim(); }).filter(function(x){
    if (!/^[A-Za-z0-9_-]{15,}$/.test(x)) return false;
    if (seen[x]) return false;
    seen[x] = true;
    return true;
  });
}

function buildDirectZipDownloadUrlV816_(p) {
  const serviceUrl = ScriptApp.getService().getUrl();
  const params = {
    action:'downloadPhotoZipDirect',
    fileIds:String(p.fileIds || p.ids || ''),
    zipName:safeZipFileNameV813_(p.zipName || '파주_공유사진.zip'),
    sheetId:p.sheetId || '',
    cacheBust:String(new Date().getTime())
  };
  const qs = Object.keys(params).map(function(k){ return encodeURIComponent(k) + '=' + encodeURIComponent(params[k]); }).join('&');
  return serviceUrl + '?' + qs;
}

function createPhotoZip_(p) {
  // v8.16: 더 이상 Google Drive 폴더에 ZIP 파일을 만들지 않습니다.
  // 구버전 앱이 action=createPhotoZip을 호출해도 임시 다운로드 페이지 URL만 반환합니다.
  const ids = parseFileIdsV813_(p.fileIds || p.ids || '');
  if (!ids.length) throw new Error('ZIP으로 묶을 사진 fileId가 없습니다.');
  return {
    direct:true,
    name:safeZipFileNameV813_(p.zipName || '파주_공유사진.zip'),
    photoCount:ids.length,
    downloadUrl:buildDirectZipDownloadUrlV816_(Object.assign({}, p, {fileIds:ids.join('|')})),
    viewUrl:buildDirectZipDownloadUrlV816_(Object.assign({}, p, {fileIds:ids.join('|')})),
    createdAt:new Date(),
    note:'v8.16부터 Google Drive에 ZIP 파일을 남기지 않고 임시 HTML 다운로드로 처리합니다.'
  };
}

function createPhotoZipDownloadHtml_(p) {
  try {
    const ids = parseFileIdsV813_(p.fileIds || p.ids || '');
    if (!ids.length) throw new Error('ZIP으로 묶을 사진 fileId가 없습니다.');
    if (ids.length > 150) throw new Error('한 번에 ZIP으로 묶을 수 있는 사진은 최대 150장입니다. 나눠서 다운로드해 주세요.');
    const blobs = [];
    const errors = [];
    ids.forEach(function(id, i) {
      try {
        const file = DriveApp.getFileById(id);
        let blob = file.getBlob();
        let name = file.getName() || ('photo_' + (i + 1));
        name = String(i + 1).padStart(3, '0') + '_' + name.replace(/[\\/:*?"<>|]/g, '_');
        blob = blob.setName(name);
        blobs.push(blob);
      } catch(e) { errors.push(id); }
    });
    if (!blobs.length) throw new Error('접근 가능한 사진 파일을 찾지 못했습니다. Drive 공유 권한을 확인해 주세요.');
    const now = new Date();
    const defaultName = '파주_공유사진_' + Utilities.formatDate(now, Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyyMMdd_HHmmss') + '.zip';
    const zipName = safeZipFileNameV813_(p.zipName || defaultName);
    const zipBlob = Utilities.zip(blobs, zipName);
    const base64 = Utilities.base64Encode(zipBlob.getBytes());
    const escapedName = JSON.stringify(zipName);
    const escapedBase64 = JSON.stringify(base64);
    const html = '<!doctype html><html><head><meta charset="utf-8"><title>사진 ZIP 다운로드</title>' +
      '<style>body{font-family:system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;padding:24px;line-height:1.6;color:#0f172a}button,a{display:inline-block;margin:8px 8px 0 0;padding:12px 14px;border-radius:10px;border:1px solid #bfdbfe;background:#eff6ff;color:#1d4ed8;font-weight:800;text-decoration:none;cursor:pointer}.ok{color:#15803d}.small{color:#64748b;font-size:13px}</style></head><body>' +
      '<h2>사진 ZIP 다운로드</h2><p class="ok">Google Drive에 ZIP 파일을 저장하지 않고, 임시로 만든 ZIP을 바로 다운로드합니다.</p>' +
      '<p>사진 ' + blobs.length + '장' + (errors.length ? ' / 실패 ' + errors.length + '장' : '') + '</p>' +
      '<button id="dl">다운로드 다시 시도</button><p class="small">다운로드가 자동으로 시작되지 않으면 위 버튼을 누르세요. 이 창을 닫아도 Google Drive에는 ZIP 파일이 남지 않습니다.</p>' +
      '<script>const zipName=' + escapedName + ';const b64=' + escapedBase64 + ';function b64ToBlob(b64,type){const bin=atob(b64);const len=bin.length;const arr=new Uint8Array(len);for(let i=0;i<len;i++)arr[i]=bin.charCodeAt(i);return new Blob([arr],{type:type});}function download(){const blob=b64ToBlob(b64,"application/zip");const url=URL.createObjectURL(blob);const a=document.createElement("a");a.href=url;a.download=zipName;document.body.appendChild(a);a.click();setTimeout(()=>{URL.revokeObjectURL(url);a.remove();},1500);}document.getElementById("dl").onclick=download;setTimeout(download,300);</script>' +
      '</body></html>';
    return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch(err) {
    const msg = String(err && err.message ? err.message : err).replace(/</g, '&lt;');
    return HtmlService.createHtmlOutput('<!doctype html><meta charset="utf-8"><body style="font-family:sans-serif;padding:24px"><h2>사진 ZIP 생성 실패</h2><p>' + msg + '</p><p>Apps Script 권한, Drive 파일 권한, 사진 개수를 확인해 주세요.</p></body>');
  }
}

function extractDriveFileIdV816_(value) {
  value = String(value || '');
  let m = value.match(/[?&]id=([A-Za-z0-9_-]{15,})/);
  if (m) return m[1];
  m = value.match(/\/file\/d\/([A-Za-z0-9_-]{15,})/);
  if (m) return m[1];
  m = value.match(/\/d\/([A-Za-z0-9_-]{15,})/);
  if (m) return m[1];
  return '';
}
function normalizePhotoForSheetV816_(p, idx) {
  p = p || {};
  const rawUrl = String(p.url || p.photoUrl || p.thumbUrl || p.thumbnailUrl || p.downloadUrl || p.viewUrl || '').trim();
  const fileId = String(p.fileId || p.id || extractDriveFileIdV816_(rawUrl) || '').trim();
  const thumbUrl = p.thumbUrl || p.thumbnailUrl || (fileId ? 'https://drive.google.com/thumbnail?id=' + encodeURIComponent(fileId) + '&sz=w1000' : rawUrl);
  return {
    fileId:fileId,
    name:p.name || p.fileName || p.photoName || ('공유 사진 ' + (Number(idx || 0) + 1)),
    mimeType:p.mimeType || '',
    thumbUrl:thumbUrl || rawUrl,
    url:p.url || (fileId ? 'https://drive.google.com/thumbnail?id=' + encodeURIComponent(fileId) + '&sz=w1600' : rawUrl),
    downloadUrl:p.downloadUrl || (fileId ? 'https://drive.google.com/uc?export=download&id=' + encodeURIComponent(fileId) : rawUrl),
    viewUrl:p.viewUrl || (fileId ? 'https://drive.google.com/file/d/' + encodeURIComponent(fileId) + '/view' : rawUrl),
    uploadedAt:p.uploadedAt || '',
    shared:true
  };
}
function photoMatchesDeleteV816_(photo, p) {
  const targetId = String(p.fileId || '').trim();
  const targetUrl = String(p.photoUrl || p.url || '').trim();
  const targetName = String(p.photoName || p.name || '').trim();
  const n = normalizePhotoForSheetV816_(photo, 0);
  if (targetId && n.fileId === targetId) return true;
  if (targetUrl) {
    const urls = [n.thumbUrl, n.url, n.downloadUrl, n.viewUrl].map(String);
    if (urls.indexOf(targetUrl) >= 0) return true;
    const targetUrlId = extractDriveFileIdV816_(targetUrl);
    if (targetUrlId && n.fileId === targetUrlId) return true;
  }
  if (targetName && String(n.name || '') === targetName) return true;
  return false;
}
function deletePhoto_(p) {
  const sheet = getSheet_(p);
  const id = String(p.id || '').replace(/^(gs-|shared-|imported-|edit-)/, '');
  if (!id) throw new Error('사진을 삭제할 신규 구역 id가 없습니다.');
  const targetRow = findRowById_(sheet, id);
  if (targetRow <= 0) throw new Error('Google Sheets에서 해당 신규 구역 행을 찾지 못했습니다: ' + id);
  const rowObj = readExistingRowObj_(sheet, targetRow);
  const before = parsePhotoJson_(rowObj.photoJson).map(normalizePhotoForSheetV816_);
  const after = before.filter(function(photo){ return !photoMatchesDeleteV816_(photo, p); });
  if (after.length === before.length) throw new Error('삭제할 사진을 photoJson에서 찾지 못했습니다. 공유 신규 구역 불러오기를 다시 한 뒤 시도해 주세요.');
  const removed = before.filter(function(photo){ return photoMatchesDeleteV816_(photo, p); });
  rowObj.photoJson = JSON.stringify(after);
  rowObj.photoUrls = after.map(function(x){ return x.thumbUrl || x.url || x.downloadUrl || ''; }).filter(String).join('|');
  rowObj.photoNames = after.map(function(x){ return x.name || ''; }).join('|');
  rowObj.photoUpdatedAt = new Date();
  rowObj.updatedAt = new Date();
  sheet.getRange(targetRow, 1, 1, HEADERS.length).setValues([HEADERS.map(function(h){ return rowObj[h]; })]);
  SpreadsheetApp.flush();
  let trashed = false;
  if (String(p.trashFile || 'true') !== 'false') {
    const ids = {};
    removed.forEach(function(photo){ if (photo.fileId) ids[photo.fileId] = true; });
    if (p.fileId) ids[String(p.fileId)] = true;
    Object.keys(ids).forEach(function(fid){
      try { DriveApp.getFileById(fid).setTrashed(true); trashed = true; } catch(e) {}
    });
  }
  return {ok:true, deletedCount:removed.length, trashed:trashed, item:rowObj};
}
