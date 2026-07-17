/**
 * ==== 画像・プロンプト共有ボード バックエンド (Stable Diffusion設定対応版) ====
 *
 * 【セットアップ手順】
 * 1. 新しいGoogleスプレッドシートを作成する
 * 2. 「拡張機能」→「Apps Script」を開き、このファイルの中身を丸ごと貼り付ける
 * 3. 下の FOLDER_ID に、画像を保存したいGoogle Driveフォルダの ID を入れる
 *    (フォルダを作成→開いたURLの https://drive.google.com/drive/folders/【ここ】 の部分)
 * 4. 上部の関数選択プルダウンで setup を選び、実行(▶)する
 *    → 初回は権限承認画面が出るので許可する。これで posts シートが自動生成される
 * 5. 右上「デプロイ」→「新しいデプロイ」
 *    - 種類: ウェブアプリ
 *    - 次のユーザーとして実行: 自分
 *    - アクセスできるユーザー: 全員
 *    → デプロイ後に表示される「ウェブアプリのURL」をコピーする
 * 6. そのURLを index.html の GAS_URL に貼り付ける
 *
 * 【重要】コードを書き換えたあとは、毎回「デプロイ」→「デプロイを管理」→
 * 鉛筆アイコン→バージョン「新バージョン」→「デプロイ」で更新しないと反映されません。
 * (「新しいデプロイ」を作るとURLが変わるので、既存デプロイの更新を推奨)
 */

const FOLDER_ID = '1JcqYdFAf2GXQurEF7a0kEZut1KabGVd0'; // 画像保存先フォルダ
const SHEET_NAME = 'posts';
const ADMIN_SHEET_NAME = 'admin';       // B1=管理パスワード / B2=オプション設定JSON / A5以降=連絡メッセージ
const COMMENTS_SHEET_NAME = 'comments';  // 作品へのコメント

// オプション機能の既定値(すべて表示)
const DEFAULT_OPTIONS = {
  like: true,          // いいねボタン
  likeCount: true,     // いいね数表示
  downloadCount: true, // ダウンロード数表示
  comment: true,       // コメント機能
  author: true         // 投稿者名表示(offで全員匿名)
};

// シートの列構成 (この順番でスプレッドシートに書き込まれます)
const COLUMNS = [
  'id', 'timestamp', 'imageUrl',
  'title', 'author', 'prompt', 'negativePrompt',
  'settingsJson', 'tags', 'memo',
  'shareId', 'downloadCount', 'likeCount',
  'type', 'basePromptId', 'parentId'
];

// 初回に1回だけ実行する
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(COLUMNS);
  } else {
    sheet.getRange(1, 1, 1, COLUMNS.length).setValues([COLUMNS]);
  }

  let adminSheet = ss.getSheetByName(ADMIN_SHEET_NAME);
  if (!adminSheet) {
    adminSheet = ss.insertSheet(ADMIN_SHEET_NAME);
  }
  if (!adminSheet.getRange('A1').getValue()) adminSheet.getRange('A1').setValue('password');
  if (!adminSheet.getRange('C1').getValue()) adminSheet.getRange('C1').setValue('← B1に管理画面用パスワードを入力(空欄の間はログイン不可)');
  if (!adminSheet.getRange('A2').getValue()) adminSheet.getRange('A2').setValue('options');
  if (!adminSheet.getRange('C2').getValue()) adminSheet.getRange('C2').setValue('← B2はオプション機能の設定(管理画面から自動保存されます。手動編集不要)');
  if (!adminSheet.getRange('A4').getValue()) adminSheet.getRange('A4').setValue('▼ 以下に利用者からの連絡メッセージが追記されます ▼');

  let commentsSheet = ss.getSheetByName(COMMENTS_SHEET_NAME);
  if (!commentsSheet) {
    commentsSheet = ss.insertSheet(COMMENTS_SHEET_NAME);
    commentsSheet.appendRow(['id', 'postId', 'timestamp', 'name', 'body']);
  } else if (commentsSheet.getLastRow() === 0) {
    commentsSheet.appendRow(['id', 'postId', 'timestamp', 'name', 'body']);
  }
}

function doGet(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const values = sheet.getDataRange().getValues();
  const rows = values.slice(1); // ヘッダー除く

  const posts = rows
    .filter(function (row) { return row[0]; }) // 空行除外
    .map(function (row) {
      let settings = {};
      try { settings = JSON.parse(row[7] || '{}'); } catch (e) { settings = {}; }

      return {
        id: row[0],
        timestamp: new Date(row[1]).toISOString(),
        imageUrl: row[2],
        title: row[3] || '',
        author: row[4] || '',
        prompt: row[5] || '',
        negativePrompt: row[6] || '',
        settings: settings,
        tags: row[8]
          ? String(row[8]).split(',').map(function (t) { return t.trim(); }).filter(Boolean)
          : [],
        memo: row[9] || '',
        shareId: row[10] || '',
        downloadCount: Number(row[11]) || 0,
        likeCount: Number(row[12]) || 0,
        type: row[13] || 'image',
        basePromptId: row[14] || '',
        parentId: row[15] || ''
      };
    });

  const json = JSON.stringify({ posts: posts, comments: getAllComments(), options: getOptions() });

  // JSONP(callbackパラメータ付き)の場合はJavaScriptとして返す。
  // GitHub Pages(クロスオリジン)からのfetchはCORSでブロックされるため、
  // <script>タグ読み込みで回避するための対応。
  if (e && e.parameter && e.parameter.callback) {
    return ContentService
      .createTextOutput(e.parameter.callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonOutput(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const raw = (e.parameter && e.parameter.payload) ? e.parameter.payload : e.postData.contents;
    const body = JSON.parse(raw);
    const action = body.action || 'post';

    if (action === 'like') return jsonOutput(handleLike(body));
    if (action === 'download') return jsonOutput(handleDownload(body));
    if (action === 'contact') return jsonOutput(handleContact(body));
    if (action === 'comment') return jsonOutput(handleComment(body));
    if (action === 'adminAuth') return jsonOutput(handleAdminAuth(body));
    if (action === 'adminUpdate') return jsonOutput(handleAdminUpdate(body));
    if (action === 'adminDelete') return jsonOutput(handleAdminDelete(body));
    if (action === 'adminSetOptions') return jsonOutput(handleAdminSetOptions(body));
    if (action === 'adminDeleteComment') return jsonOutput(handleAdminDeleteComment(body));
    return jsonOutput(handleNewPost(body));

  } catch (err) {
    console.error('doPost failed: ' + err);
    return jsonOutput({ success: false, error: String(err) });
  }
}

function uploadImageToDrive(body) {
  const base64 = body.imageBase64.split(',').pop();
  const blob = Utilities.newBlob(
    Utilities.base64Decode(base64),
    body.imageMime || 'image/png',
    body.imageName || ('upload_' + new Date().getTime())
  );
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const file = folder.createFile(blob);

  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (sharingErr) {
    console.error('setSharing failed: ' + sharingErr);
  }

  return 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1000';
}

function driveIdFromUrl(url) {
  const m = String(url || '').match(/[?&]id=([^&]+)/);
  return m ? m[1] : '';
}

function handleNewPost(body) {
  const type = body.type === 'prompt' ? 'prompt' : 'image';
  let imageUrl = '';

  if (type === 'image') {
    imageUrl = uploadImageToDrive(body);
  }

  const id = Utilities.getUuid();
  const shareId = Utilities.getUuid().replace(/-/g, '').slice(0, 6);
  const timestamp = new Date();
  const tags = (body.tags || []).join(',');
  const settingsJson = JSON.stringify(body.settings || {});

  const sheet = getPostsSheet();
  sheet.appendRow([
    id, timestamp, imageUrl,
    body.title || '', body.author || '', body.prompt || '', body.negativePrompt || '',
    settingsJson, tags, body.memo || '',
    shareId, 0, 0,
    type, body.basePromptId || '', body.parentId || ''
  ]);

  return { success: true, id: id, shareId: shareId, imageUrl: imageUrl, type: type };
}

function handleDownload(body) {
  return { success: true, downloadCount: bumpCount(body.id, 11, 1) };
}

function handleLike(body) {
  const delta = Number(body.delta) === -1 ? -1 : 1;
  return { success: true, likeCount: bumpCount(body.id, 12, delta) };
}

function bumpCount(id, colIndex, delta) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getPostsSheet();
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        let n = Number(data[i][colIndex]) || 0;
        n = Math.max(0, n + delta);
        sheet.getRange(i + 1, colIndex + 1).setValue(n);
        return n;
      }
    }
    throw new Error('対象の投稿が見つかりません: ' + id);
  } finally {
    lock.releaseLock();
  }
}

// ===== ADMIN =====
function getAdminPassword() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ADMIN_SHEET_NAME);
  if (!sheet) return '';
  return String(sheet.getRange('B1').getValue() || '').trim();
}

function checkAdminPassword(password) {
  const stored = getAdminPassword();
  return stored !== '' && String(password || '') === stored;
}

function handleAdminAuth(body) {
  if (!checkAdminPassword(body.password)) return { success: false, error: 'auth' };
  return { success: true };
}

function handleAdminUpdate(body) {
  if (!checkAdminPassword(body.password)) return { success: false, error: 'auth' };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getPostsSheet();
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(body.id)) {
        const row = i + 1;

        if (body.imageBase64) {
          const newUrl = uploadImageToDrive(body);
          const oldId = driveIdFromUrl(data[i][2]);
          if (oldId) {
            try { DriveApp.getFileById(oldId).setTrashed(true); } catch (e) { console.error('trash failed: ' + e); }
          }
          sheet.getRange(row, 3).setValue(newUrl);
        }

        sheet.getRange(row, 4).setValue(body.title || '');
        sheet.getRange(row, 5).setValue(body.author || '');
        sheet.getRange(row, 6).setValue(body.prompt || '');
        sheet.getRange(row, 7).setValue(body.negativePrompt || '');
        sheet.getRange(row, 8).setValue(JSON.stringify(body.settings || {}));
        sheet.getRange(row, 9).setValue((body.tags || []).join(','));
        sheet.getRange(row, 10).setValue(body.memo || '');
        sheet.getRange(row, 15).setValue(body.basePromptId || '');
        sheet.getRange(row, 16).setValue(body.parentId || '');
        return { success: true, id: body.id };
      }
    }
    return { success: false, error: '対象の投稿が見つかりません: ' + body.id };
  } finally {
    lock.releaseLock();
  }
}

function handleAdminDelete(body) {
  if (!checkAdminPassword(body.password)) return { success: false, error: 'auth' };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getPostsSheet();
    const data = sheet.getDataRange().getValues();
    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(body.id)) { targetRow = i + 1; break; }
    }
    if (targetRow === -1) return { success: false, error: '対象の投稿が見つかりません: ' + body.id };

    sheet.deleteRow(targetRow);

    const remaining = sheet.getDataRange().getValues();
    for (let i = 1; i < remaining.length; i++) {
      const row = i + 1;
      if (String(remaining[i][15]) === String(body.id)) {
        sheet.getRange(row, 16).setValue('');
      }
      const baseIds = String(remaining[i][14] || '').split(',').map(function (s) { return s.trim(); }).filter(Boolean);
      if (baseIds.indexOf(String(body.id)) !== -1) {
        sheet.getRange(row, 15).setValue(baseIds.filter(function (id) { return id !== String(body.id); }).join(','));
      }
    }

    // 紐づくコメントも削除
    try {
      const cSheet = getCommentsSheet();
      const cData = cSheet.getDataRange().getValues();
      for (let i = cData.length - 1; i >= 1; i--) {
        if (String(cData[i][1]) === String(body.id)) cSheet.deleteRow(i + 1);
      }
    } catch (e) { console.error('comment cleanup failed: ' + e); }

    return { success: true, id: body.id };
  } finally {
    lock.releaseLock();
  }
}

// ===== OPTIONS(オプション機能フラグ) =====
function getAdminSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ADMIN_SHEET_NAME);
}

function getOptions() {
  const sheet = getAdminSheet();
  const opts = {};
  Object.keys(DEFAULT_OPTIONS).forEach(function (k) { opts[k] = DEFAULT_OPTIONS[k]; });
  if (!sheet) return opts;
  const raw = String(sheet.getRange('B2').getValue() || '').trim();
  if (raw) {
    try {
      const parsed = JSON.parse(raw);
      Object.keys(DEFAULT_OPTIONS).forEach(function (k) {
        if (typeof parsed[k] === 'boolean') opts[k] = parsed[k];
      });
    } catch (e) { /* 壊れていたら既定値 */ }
  }
  return opts;
}

function handleAdminSetOptions(body) {
  if (!checkAdminPassword(body.password)) return { success: false, error: 'auth' };
  const sheet = getAdminSheet();
  if (!sheet) return { success: false, error: 'adminシートがありません(setupを実行してください)' };
  const opts = {};
  Object.keys(DEFAULT_OPTIONS).forEach(function (k) {
    opts[k] = (body.options && typeof body.options[k] === 'boolean') ? body.options[k] : DEFAULT_OPTIONS[k];
  });
  sheet.getRange('B2').setValue(JSON.stringify(opts));
  return { success: true, options: opts };
}

// ===== CONTACT(管理人への連絡) =====
function handleContact(body) {
  const msg = String(body.message || '').trim();
  if (!msg) return { success: false, error: 'メッセージが空です' };
  const sheet = getAdminSheet();
  if (!sheet) return { success: false, error: 'adminシートがありません(setupを実行してください)' };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    const name = String(body.name || '').trim() || '(匿名)';
    sheet.getRange(sheet.getLastRow() + 1, 1).setValue('[' + stamp + '] ' + name + ': ' + msg);
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

// ===== COMMENTS(作品コメント) =====
function getCommentsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(COMMENTS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(COMMENTS_SHEET_NAME);
    sheet.appendRow(['id', 'postId', 'timestamp', 'name', 'body']);
  }
  return sheet;
}

function getAllComments() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(COMMENTS_SHEET_NAME);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  return values.slice(1)
    .filter(function (row) { return row[0] && row[1]; })
    .map(function (row) {
      return {
        id: row[0],
        postId: String(row[1]),
        timestamp: new Date(row[2]).toISOString(),
        name: row[3] || '',
        body: row[4] || ''
      };
    });
}

function handleComment(body) {
  const text = String(body.body || '').trim();
  if (!body.postId) return { success: false, error: '投稿IDがありません' };
  if (!text) return { success: false, error: 'コメントが空です' };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getCommentsSheet();
    const id = Utilities.getUuid();
    const timestamp = new Date();
    const name = String(body.name || '').trim();
    sheet.appendRow([id, String(body.postId), timestamp, name, text]);
    return { success: true, comment: { id: id, postId: String(body.postId), timestamp: timestamp.toISOString(), name: name, body: text } };
  } finally {
    lock.releaseLock();
  }
}

function handleAdminDeleteComment(body) {
  if (!checkAdminPassword(body.password)) return { success: false, error: 'auth' };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getCommentsSheet();
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(body.id)) {
        sheet.deleteRow(i + 1);
        return { success: true, id: body.id };
      }
    }
    return { success: false, error: 'コメントが見つかりません' };
  } finally {
    lock.releaseLock();
  }
}

function getPostsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(COLUMNS);
  }
  return sheet;
}
