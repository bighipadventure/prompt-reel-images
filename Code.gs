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

// シートの列構成 (この順番でスプレッドシートに書き込まれます)
const COLUMNS = [
  'id', 'timestamp', 'imageUrl',
  'title', 'author', 'prompt', 'negativePrompt',
  'settingsJson', 'tags', 'memo',
  'shareId', 'downloadCount', 'likeCount'
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
        likeCount: Number(row[12]) || 0
      };
    });

  const json = JSON.stringify({ posts: posts });

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
    return jsonOutput(handleNewPost(body));

  } catch (err) {
    console.error('doPost failed: ' + err);
    return jsonOutput({ success: false, error: String(err) });
  }
}

function handleNewPost(body) {
  // 画像をDriveに保存
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

  const imageUrl = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1000';

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
    shareId, 0, 0
  ]);

  return { success: true, id: id, shareId: shareId, imageUrl: imageUrl };
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

function getPostsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(COLUMNS);
  }
  return sheet;
}
