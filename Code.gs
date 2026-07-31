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
const REF_FOLDER_ID = '19wZly5-DRErlCJl-2ldRDF_xrPx2onC5'; // リクエストの参考資料 保存先フォルダ
const SHEET_NAME = 'posts';
const ADMIN_SHEET_NAME = 'admin';       // B1=管理パスワード / B2=オプション設定JSON / A5以降=連絡メッセージ
const COMMENTS_SHEET_NAME = 'comments';  // 作品へのコメント
const BOARD_SHEET_NAME = 'board';        // 掲示板メッセージ
const DICT_SHEET_NAME = 'dictionary';    // プロンプト内部辞書(en / ja / keywords / genre)
const DICT_COLUMNS = ['en', 'ja', 'keywords', 'genre'];
const DICT_SYNC_LIMIT = 50;              // 1回の同期で翻訳する新規語の上限

// オプション機能の既定値(すべて表示)
const DEFAULT_OPTIONS = {
  like: true,          // いいねボタン
  likeCount: true,     // いいね数表示
  downloadCount: true, // ダウンロード数表示
  comment: true,       // コメント機能
  author: true,        // 投稿者名表示(offで全員匿名)
  board: true,         // 掲示板タブ
  boardVote: true,     // 掲示板のグッド/バッド
  fixRequest: false,   // 修正依頼機能(タブ・投稿モード・紐づけ・「修正依頼作品も表示」)。既定オフ
  dislike: false       // よくないねボタン + しきい値到達で強制非表示。既定オフ
};
const DEFAULT_DISLIKE_THRESHOLD = 10; // よくないねが何回でその投稿を強制非表示にするか

// シートの列構成 (この順番でスプレッドシートに書き込まれます)
const COLUMNS = [
  'id', 'timestamp', 'imageUrl',
  'title', 'author', 'prompt', 'negativePrompt',
  'settingsJson', 'tags', 'memo',
  'shareId', 'downloadCount', 'likeCount',
  'type', 'basePromptId', 'parentId',
  'referenceImages', 'requestId',
  'fixContent', 'fixRequestId', 'dislikeCount'
];

const BOARD_COLUMNS = ['id', 'timestamp', 'name', 'body', 'goodCount', 'badCount'];

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

  let boardSheet = ss.getSheetByName(BOARD_SHEET_NAME);
  if (!boardSheet) {
    boardSheet = ss.insertSheet(BOARD_SHEET_NAME);
    boardSheet.appendRow(BOARD_COLUMNS);
  } else if (boardSheet.getLastRow() === 0) {
    boardSheet.appendRow(BOARD_COLUMNS);
  } else {
    boardSheet.getRange(1, 1, 1, BOARD_COLUMNS.length).setValues([BOARD_COLUMNS]);
  }

  let dictSheet = ss.getSheetByName(DICT_SHEET_NAME);
  if (!dictSheet) {
    dictSheet = ss.insertSheet(DICT_SHEET_NAME);
  }
  if (dictSheet.getLastRow() === 0) {
    dictSheet.appendRow(DICT_COLUMNS);
  } else {
    dictSheet.getRange(1, 1, 1, DICT_COLUMNS.length).setValues([DICT_COLUMNS]);
  }
  if (dictSheet.getLastRow() <= 1) {
    const seedRows = buildSeedRows();
    if (seedRows.length) dictSheet.getRange(2, 1, seedRows.length, DICT_COLUMNS.length).setValues(seedRows);
  }

  installDictSyncTrigger();          // 定期同期トリガー(重複設置しない)
  syncDictionary();                  // 既存投稿分の初回バックフィル(LanguageApp権限承認も兼ねる)
}

// 投稿シートの1行を投稿オブジェクトへ変換する共通処理(doGetの各stageから共通で使う)
function parsePostRow(row) {
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
    parentId: row[15] || '',
    referenceImages: parseReferenceImages(row[16]),
    requestId: row[17] || '',
    fixContent: row[18] || '',
    fixRequestId: row[19] || '',
    dislikeCount: Number(row[20]) || 0
  };
}

function parseAllPosts() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const values = sheet.getDataRange().getValues();
  return values.slice(1)
    .filter(function (row) { return row[0]; }) // 空行除外
    .map(parsePostRow);
}

// 「親の画像投稿」= type=image かつ parentId無し。段階読み込みの第1段階(stage=parents)で最初に返す対象。
// 一覧に最初に表示される投稿のみを軽量に返すことで、体感速度とタイムアウト耐性を上げる。
function isParentImagePost(p) {
  return (p.type || 'image') === 'image' && !p.parentId;
}

function basePromptIdsOfGas(p) {
  return String((p && p.basePromptId) || '').split(',').map(function (s) { return s.trim(); }).filter(Boolean);
}
function fixRequestIdsOfGas(p) {
  return String((p && p.fixRequestId) || '').split(',').map(function (s) { return s.trim(); }).filter(Boolean);
}

function commentsForId(postId) {
  return getAllComments().filter(function (c) { return String(c.postId) === String(postId); });
}

// フロントは段階的に読み込む:
//   stage=parents : 親の画像投稿のみ(最初に一覧表示する分だけ最小限のデータで返す)
//   stage=rest    : それ以外(子画像・プロンプト例・リクエスト・修正依頼・コメント・掲示板・辞書)を背後で取得
//   detailId=     : 特定の投稿1件+直接関連する投稿(ベース/親/リクエスト/修正依頼、逆参照含む)+コメントを
//                   優先取得する。段階読み込みの完了を待たずに個別の投稿を正しく開くために使う。
// いずれのパラメータも無ければ従来どおり全件を1回で返す(後方互換。旧フロント・他の連携先向け)。
function doGet(e) {
  const params = (e && e.parameter) || {};
  let payload;

  if (params.detailId) {
    payload = buildDetailPayload(params.detailId);
  } else if (params.stage === 'parents') {
    payload = buildParentsPayload();
  } else if (params.stage === 'rest') {
    payload = buildRestPayload();
  } else {
    payload = buildFullPayload();
  }

  return outputPayload(e, payload);
}

function outputPayload(e, obj) {
  const json = JSON.stringify(obj);

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

function buildParentsPayload() {
  const all = parseAllPosts();
  const parents = all.filter(isParentImagePost);
  return { posts: parents, options: getOptions(), stage: 'parents' };
}

function buildRestPayload() {
  const all = parseAllPosts();
  const rest = all.filter(function (p) { return !isParentImagePost(p); });
  return {
    posts: rest,
    comments: getAllComments(),
    board: getAllBoard(),
    dictionary: getDictionary(),
    options: getOptions(),
    stage: 'rest'
  };
}

function buildFullPayload() {
  const all = parseAllPosts();
  return { posts: all, comments: getAllComments(), board: getAllBoard(), options: getOptions(), dictionary: getDictionary() };
}

// 指定した投稿1件と、それに直接紐づく投稿(ベースプロンプト/親/応えたリクエスト/対応した修正依頼、
// および逆方向にこの投稿を参照している投稿=子・回答作品など)・コメントをまとめて返す。
// stage=rest の完了を待たずに、個別に開いた投稿の詳細モーダルを正しく(紐づけ・コメント込みで)表示するための優先取得用。
function buildDetailPayload(idOrShareId) {
  const all = parseAllPosts();
  const post = all.find(function (p) { return String(p.id) === String(idOrShareId) || String(p.shareId) === String(idOrShareId); });
  if (!post) return { post: null };

  const relatedIds = {};
  basePromptIdsOfGas(post).forEach(function (id) { relatedIds[id] = true; });
  if (post.parentId) relatedIds[post.parentId] = true;
  if (post.requestId) relatedIds[post.requestId] = true;
  fixRequestIdsOfGas(post).forEach(function (id) { relatedIds[id] = true; });

  all.forEach(function (p) {
    if (String(p.id) === String(post.id)) return;
    if (String(p.parentId) === String(post.id)) relatedIds[p.id] = true;
    if (String(p.requestId) === String(post.id)) relatedIds[p.id] = true;
    if (fixRequestIdsOfGas(p).indexOf(String(post.id)) !== -1) relatedIds[p.id] = true;
    if (basePromptIdsOfGas(p).indexOf(String(post.id)) !== -1) relatedIds[p.id] = true;
  });

  const related = all.filter(function (p) { return relatedIds[p.id]; });

  return { post: post, related: related, comments: commentsForId(post.id), options: getOptions() };
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
    if (action === 'dislike') return jsonOutput(handleDislike(body));
    if (action === 'download') return jsonOutput(handleDownload(body));
    if (action === 'contact') return jsonOutput(handleContact(body));
    if (action === 'comment') return jsonOutput(handleComment(body));
    if (action === 'adminAuth') return jsonOutput(handleAdminAuth(body));
    if (action === 'adminUpdate') return jsonOutput(handleAdminUpdate(body));
    if (action === 'adminDelete') return jsonOutput(handleAdminDelete(body));
    if (action === 'adminSetOptions') return jsonOutput(handleAdminSetOptions(body));
    if (action === 'adminDeleteComment') return jsonOutput(handleAdminDeleteComment(body));
    if (action === 'board') return jsonOutput(handleBoardPost(body));
    if (action === 'boardVote') return jsonOutput(handleBoardVote(body));
    if (action === 'adminDeleteBoard') return jsonOutput(handleAdminDeleteBoard(body));
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

function parseReferenceImages(raw) {
  if (!raw) return [];
  try {
    const arr = JSON.parse(raw);
    return Array.isArray(arr) ? arr : [];
  } catch (e) {
    return [];
  }
}

function uploadReferenceImages(dataUrls) {
  if (!dataUrls || !dataUrls.length) return [];
  const folder = DriveApp.getFolderById(REF_FOLDER_ID);
  return dataUrls.map(function (d) {
    const base64 = String(d).split(',').pop();
    const name = 'ref_' + new Date().getTime() + '_' + Math.random().toString(36).slice(2, 7) + '.jpg';
    const blob = Utilities.newBlob(Utilities.base64Decode(base64), 'image/jpeg', name);
    const file = folder.createFile(blob);
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (sharingErr) {
      console.error('ref setSharing failed: ' + sharingErr);
    }
    return 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1000';
  });
}

function trashReferenceImages(urls) {
  (urls || []).forEach(function (u) {
    const id = driveIdFromUrl(u);
    if (id) {
      try { DriveApp.getFileById(id).setTrashed(true); } catch (e) { console.error('ref trash failed: ' + e); }
    }
  });
}

function handleNewPost(body) {
  const type = body.type === 'prompt' ? 'prompt'
    : (body.type === 'request' ? 'request'
    : (body.type === 'fix' ? 'fix' : 'image'));
  let imageUrl = '';
  let referenceImages = [];
  let requestId = '';

  if (type === 'image' || type === 'fix') {
    imageUrl = uploadImageToDrive(body);
  }
  if (type === 'image') {
    requestId = body.requestId || '';
  }
  if (type === 'request') {
    referenceImages = uploadReferenceImages(body.referenceImages);
  }

  const id = Utilities.getUuid();
  const shareId = Utilities.getUuid().replace(/-/g, '').slice(0, 6);
  const timestamp = new Date();
  const tags = (body.tags || []).join(',');
  const settingsJson = JSON.stringify(body.settings || {});
  const fixContent = type === 'fix' ? (body.fixContent || '') : '';
  const fixRequestId = type === 'image' ? (body.fixRequestId || '') : '';

  const sheet = getPostsSheet();
  sheet.appendRow([
    id, timestamp, imageUrl,
    body.title || '', body.author || '', body.prompt || '', body.negativePrompt || '',
    settingsJson, tags, body.memo || '',
    shareId, 0, 0,
    type, body.basePromptId || '', body.parentId || '',
    JSON.stringify(referenceImages), requestId,
    fixContent, fixRequestId, 0
  ]);

  return { success: true, id: id, shareId: shareId, imageUrl: imageUrl, type: type, referenceImages: referenceImages };
}

function handleDownload(body) {
  return { success: true, downloadCount: bumpCount(body.id, 11, 1) };
}

function handleLike(body) {
  const delta = Number(body.delta) === -1 ? -1 : 1;
  return { success: true, likeCount: bumpCount(body.id, 12, delta) };
}

function handleDislike(body) {
  const delta = Number(body.delta) === -1 ? -1 : 1;
  return { success: true, dislikeCount: bumpCount(body.id, 20, delta) };
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
        sheet.getRange(row, 18).setValue(body.requestId || '');
        sheet.getRange(row, 19).setValue(body.fixContent || '');
        sheet.getRange(row, 20).setValue(body.fixRequestId || '');

        const currentType = data[i][13] || 'image';
        if (currentType === 'request') {
          const current = parseReferenceImages(data[i][16]);
          const keep = Array.isArray(body.keepReferenceImages) ? body.keepReferenceImages : current;
          const removed = current.filter(function (u) { return keep.indexOf(u) === -1; });
          trashReferenceImages(removed);
          const added = uploadReferenceImages(body.newReferenceImages);
          sheet.getRange(row, 17).setValue(JSON.stringify(keep.concat(added)));
        }
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

    trashReferenceImages(parseReferenceImages(data[targetRow - 1][16]));

    sheet.deleteRow(targetRow);

    const remaining = sheet.getDataRange().getValues();
    for (let i = 1; i < remaining.length; i++) {
      const row = i + 1;
      if (String(remaining[i][15]) === String(body.id)) {
        sheet.getRange(row, 16).setValue('');
      }
      if (String(remaining[i][17]) === String(body.id)) {
        sheet.getRange(row, 18).setValue('');
      }
      const baseIds = String(remaining[i][14] || '').split(',').map(function (s) { return s.trim(); }).filter(Boolean);
      if (baseIds.indexOf(String(body.id)) !== -1) {
        sheet.getRange(row, 15).setValue(baseIds.filter(function (id) { return id !== String(body.id); }).join(','));
      }
      const fixIds = String(remaining[i][19] || '').split(',').map(function (s) { return s.trim(); }).filter(Boolean);
      if (fixIds.indexOf(String(body.id)) !== -1) {
        sheet.getRange(row, 20).setValue(fixIds.filter(function (id) { return id !== String(body.id); }).join(','));
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
  opts.dislikeThreshold = DEFAULT_DISLIKE_THRESHOLD;
  if (!sheet) return opts;
  const raw = String(sheet.getRange('B2').getValue() || '').trim();
  if (raw) {
    try {
      const parsed = JSON.parse(raw);
      Object.keys(DEFAULT_OPTIONS).forEach(function (k) {
        if (typeof parsed[k] === 'boolean') opts[k] = parsed[k];
      });
      if (typeof parsed.dislikeThreshold === 'number' && parsed.dislikeThreshold > 0) {
        opts.dislikeThreshold = Math.floor(parsed.dislikeThreshold);
      }
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
  const th = body.options && Number(body.options.dislikeThreshold);
  opts.dislikeThreshold = (th && th > 0) ? Math.floor(th) : DEFAULT_DISLIKE_THRESHOLD;
  sheet.getRange('B2').setValue(JSON.stringify(opts));
  return { success: true, options: opts };
}

// ===== CONTACT(管理者への連絡) =====
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

// ===== BOARD(掲示板) =====
function getBoardSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(BOARD_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(BOARD_SHEET_NAME);
    sheet.appendRow(BOARD_COLUMNS);
  }
  return sheet;
}

function getAllBoard() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BOARD_SHEET_NAME);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  return values.slice(1)
    .filter(function (row) { return row[0]; })
    .map(function (row) {
      return {
        id: row[0],
        timestamp: new Date(row[1]).toISOString(),
        name: row[2] || '',
        body: row[3] || '',
        goodCount: Number(row[4]) || 0,
        badCount: Number(row[5]) || 0
      };
    });
}

function handleBoardPost(body) {
  const text = String(body.body || '').trim();
  if (!text) return { success: false, error: 'メッセージが空です' };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getBoardSheet();
    const id = Utilities.getUuid();
    const timestamp = new Date();
    const name = String(body.name || '').trim();
    sheet.appendRow([id, timestamp, name, text, 0, 0]);
    return { success: true, message: { id: id, timestamp: timestamp.toISOString(), name: name, body: text, goodCount: 0, badCount: 0 } };
  } finally {
    lock.releaseLock();
  }
}

function handleBoardVote(body) {
  const goodDelta = Number(body.goodDelta) || 0;
  const badDelta = Number(body.badDelta) || 0;
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getBoardSheet();
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(body.id)) {
        const good = Math.max(0, (Number(data[i][4]) || 0) + goodDelta);
        const bad = Math.max(0, (Number(data[i][5]) || 0) + badDelta);
        sheet.getRange(i + 1, 5, 1, 2).setValues([[good, bad]]);
        return { success: true, goodCount: good, badCount: bad };
      }
    }
    return { success: false, error: 'メッセージが見つかりません' };
  } finally {
    lock.releaseLock();
  }
}

function handleAdminDeleteBoard(body) {
  if (!checkAdminPassword(body.password)) return { success: false, error: 'auth' };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getBoardSheet();
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(body.id)) {
        sheet.deleteRow(i + 1);
        return { success: true, id: body.id };
      }
    }
    return { success: false, error: 'メッセージが見つかりません' };
  } finally {
    lock.releaseLock();
  }
}

// ===== プロンプト内部辞書(スプレッドシートで一元管理) =====
// 辞書はGAS側の定期同期(syncDictionary)でのみ更新する。
// 投稿処理とは独立して動くため、投稿・画面表示を遅くしない。

function getDictionary() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DICT_SHEET_NAME);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  return values.slice(1)
    .filter(function (row) { return String(row[0] || '').trim(); })
    .map(function (row) {
      return {
        en: String(row[0]).trim(),
        ja: String(row[1] || '').trim(),
        keywords: String(row[2] || '').trim(),
        genre: String(row[3] || '').trim()
      };
    });
}

// プロンプトのトークン分解(フロントの処理と対応させる)
function dictCleanToken(raw) {
  let t = String(raw).replace(/[\r\n]+/g, ' ');
  t = t.replace(/\\([()\[\]])/g, '$1');
  t = t.replace(/[(){}\[\]]/g, ' ');
  t = t.replace(/:\s*-?\d+(?:\.\d+)?/g, ' ');
  t = t.replace(/\s+/g, ' ').trim();
  return t;
}
function dictSplitTokens(text) {
  return String(text || '').split(',').map(dictCleanToken).filter(function (x) { return x; });
}
function dictKey(token) {
  return token.toLowerCase().replace(/_/g, ' ').replace(/\s+/g, ' ').trim();
}
function translateEnToJa(text) {
  try {
    const ja = LanguageApp.translate(String(text), 'en', 'ja');
    return ja ? String(ja).trim() : '';
  } catch (err) {
    console.error('translate failed for "' + text + '": ' + err);
    return '';
  }
}

// 定期トリガーを1つだけ設置する(setupから呼ぶ。重複設置しない)
function installDictSyncTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'syncDictionary') return;
  }
  ScriptApp.newTrigger('syncDictionary').timeBased().everyMinutes(5).create();
}

// 投稿に含まれるプロンプト語のうち、辞書に無いものだけを翻訳して追記する。
// また en があって ja が空の行(手動追加など)も翻訳して埋める。
// 1回あたり DICT_SYNC_LIMIT 件までに抑え、実行時間と翻訳クォータを節約する。
// トリガー(5分毎)で自動実行されるため、投稿処理を待たせない。
function syncDictionary() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return; // 別の同期が動作中ならスキップ
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dictSheet = ss.getSheetByName(DICT_SHEET_NAME);
    if (!dictSheet) return;

    const dictValues = dictSheet.getDataRange().getValues();
    const existing = {};
    const emptyJaRows = [];
    for (let i = 1; i < dictValues.length; i++) {
      const en = String(dictValues[i][0] || '').trim();
      if (!en) continue;
      const key = dictKey(en);
      existing[key] = true;
      if (!String(dictValues[i][1] || '').trim()) emptyJaRows.push({ row: i + 1, en: en });
    }

    // 投稿(画像・プロンプト例)から辞書未登録の語を集める
    const postSheet = ss.getSheetByName(SHEET_NAME);
    const newTokens = [];
    const seenNew = {};
    if (postSheet) {
      const pv = postSheet.getDataRange().getValues();
      for (let r = 1; r < pv.length; r++) {
        if (!pv[r][0]) continue;
        [pv[r][5], pv[r][6]].forEach(function (text) { // prompt / negativePrompt
          dictSplitTokens(text).forEach(function (tok) {
            const key = dictKey(tok);
            if (!key || existing[key] || seenNew[key]) return;
            seenNew[key] = true;
            newTokens.push(key);
          });
        });
      }
    }

    let budget = DICT_SYNC_LIMIT;

    // 1) en有・ja無 の行を翻訳して埋める
    for (let e = 0; e < emptyJaRows.length && budget > 0; e++) {
      const ja = translateEnToJa(emptyJaRows[e].en);
      if (ja) { dictSheet.getRange(emptyJaRows[e].row, 2).setValue(ja); budget--; }
    }

    // 2) 新規語を翻訳して追記(keywords/genreは空欄。管理者が育てる)
    const appendRows = [];
    for (let n = 0; n < newTokens.length && budget > 0; n++) {
      appendRows.push([newTokens[n], translateEnToJa(newTokens[n]), '', '']);
      budget--;
    }
    if (appendRows.length) {
      dictSheet.getRange(dictSheet.getLastRow() + 1, 1, appendRows.length, DICT_COLUMNS.length).setValues(appendRows);
    }
  } finally {
    lock.releaseLock();
  }
}

// setup() 初回のみ dictionary シートへ書き込む内部辞書のシード(en / ja / keywords / genre)
// keywords列は空(管理者が検索用の別名・関連語を追記していく想定)。
function buildSeedRows() {
  const rows = [];
  Object.keys(SEED_DICT_BY_GENRE).forEach(function (genre) {
    const group = SEED_DICT_BY_GENRE[genre];
    Object.keys(group).forEach(function (en) {
      rows.push([en, group[en], '', genre]);
    });
  });
  return rows;
}

const SEED_DICT_BY_GENRE = {
  '品質': {
    'masterpiece': '傑作(品質強調)', 'best quality': '最高品質', 'high quality': '高品質', 'ultra quality': '超高品質',
    'high resolution': '高解像度', 'highres': '高解像度', 'ultra detailed': '超緻密', 'extremely detailed': '極めて緻密',
    'detailed': '緻密', 'very detailed': 'とても緻密', 'intricate details': '精緻なディテール', 'absurdres': '超高解像度',
    '8k': '8K画質', '4k': '4K画質', 'ultra high res': '超高解像度', 'official art': '公式絵', 'sharp focus': 'ピント明瞭',
    'professional': 'プロ品質', 'award winning': '受賞級', 'cinematic': '映画的'
  },
  '画風': {
    'photorealistic': '写実的', 'realistic': 'リアル', 'hyperrealistic': '超写実的', 'semi realistic': '半写実',
    'anime': 'アニメ調', 'anime style': 'アニメ調', 'manga': '漫画調', 'illustration': 'イラスト', 'digital art': 'デジタルアート',
    'digital painting': 'デジタル絵画', 'concept art': 'コンセプトアート', 'sketch': 'スケッチ', 'lineart': '線画',
    'flat color': 'フラットカラー', 'watercolor': '水彩', 'oil painting': '油絵', 'painterly': '絵画的',
    'cel shading': 'セルシェード', 'chibi': 'ちびキャラ', 'realism': '写実主義'
  },
  '人数・人物': {
    '1girl': '女の子1人', '2girls': '女の子2人', '3girls': '女の子3人', 'multiple girls': '複数の女の子',
    '1boy': '男の子1人', '2boys': '男の子2人', 'multiple boys': '複数の男の子', '1other': 'その他1人',
    'solo': '単独', 'solo focus': '一人に焦点', 'couple': '男女ペア', 'group': 'グループ',
    'girl': '女の子', 'boy': '男の子', 'woman': '女性', 'man': '男性', 'female': '女性', 'male': '男性',
    'adult': '大人', 'mature female': '成人女性', 'child': '子供', 'teenage': '10代'
  },
  '髪': {
    'long hair': '長い髪', 'short hair': '短い髪', 'medium hair': 'セミロング', 'very long hair': 'とても長い髪',
    'twintails': 'ツインテール', 'ponytail': 'ポニーテール', 'side ponytail': 'サイドポニー', 'braid': '三つ編み',
    'twin braids': '二つ結びの三つ編み', 'bun': 'お団子ヘア', 'double bun': '二つのお団子', 'bob cut': 'ボブ',
    'bangs': '前髪', 'blunt bangs': 'ぱっつん前髪', 'swept bangs': '流し前髪', 'hair between eyes': '目にかかる前髪',
    'curly hair': '巻き毛', 'wavy hair': 'ウェーブヘア', 'straight hair': 'ストレートヘア', 'messy hair': '乱れ髪',
    'black hair': '黒髪', 'brown hair': '茶髪', 'blonde hair': '金髪', 'blond hair': '金髪', 'white hair': '白髪',
    'silver hair': '銀髪', 'grey hair': '灰髪', 'gray hair': '灰髪', 'red hair': '赤髪', 'pink hair': 'ピンク髪',
    'blue hair': '青髪', 'purple hair': '紫髪', 'green hair': '緑髪', 'orange hair': 'オレンジ髪', 'aqua hair': '水色髪',
    'multicolored hair': '多色の髪', 'gradient hair': 'グラデ髪', 'two tone hair': '二色の髪', 'ahoge': 'アホ毛',
    'floating hair': 'なびく髪', 'wind lift': '風になびく'
  },
  '目': {
    'blue eyes': '青い目', 'red eyes': '赤い目', 'green eyes': '緑の目', 'brown eyes': '茶色の目',
    'black eyes': '黒い目', 'purple eyes': '紫の目', 'yellow eyes': '黄色い目', 'golden eyes': '金色の目',
    'pink eyes': 'ピンクの目', 'heterochromia': 'オッドアイ', 'aqua eyes': '水色の目', 'grey eyes': '灰色の目',
    'closed eyes': '閉じた目', 'half closed eyes': '半目', 'wide eyes': '見開いた目', 'glowing eyes': '光る目',
    'eyelashes': 'まつげ', 'detailed eyes': '緻密な目', 'beautiful detailed eyes': '美しく緻密な目'
  },
  '表情': {
    'smile': '笑顔', 'smiling': '微笑み', 'grin': 'にっこり', 'open mouth': '開いた口', 'closed mouth': '閉じた口',
    'light smile': '薄い笑み', 'blush': '赤面', 'crying': '泣いている', 'tears': '涙', 'angry': '怒り顔',
    'sad': '悲しい表情', 'surprised': '驚き顔', 'expressionless': '無表情', 'serious': '真剣な表情',
    'pout': 'ふくれっ面', 'embarrassed': '照れ顔', 'happy': '幸せそう', 'looking at viewer': 'こちらを見る',
    'looking away': '視線をそらす', 'looking back': '振り返る', 'looking up': '見上げる', 'looking down': '見下ろす'
  },
  '顔・肌': {
    'beautiful': '美しい', 'cute': 'かわいい', 'pretty': '綺麗', 'gorgeous': '華やか', 'elegant': '上品',
    'detailed face': '緻密な顔', 'perfect face': '整った顔', 'symmetrical face': '左右対称の顔',
    'pale skin': '色白の肌', 'dark skin': '褐色肌', 'tan': '日焼け肌', 'freckles': 'そばかす', 'mole': 'ほくろ',
    'makeup': '化粧', 'lipstick': '口紅', 'eyeshadow': 'アイシャドウ'
  },
  '体型': {
    'slim': 'スリム', 'petite': '小柄', 'tall': '長身', 'large breasts': '大きな胸', 'medium breasts': '普通の胸',
    'small breasts': '小さな胸', 'wide hips': '広い腰', 'thighs': '太もも', 'thick thighs': '太い太もも',
    'collarbone': '鎖骨', 'navel': 'へそ', 'bare shoulders': '肩出し'
  },
  '服装': {
    'school uniform': '制服', 'serafuku': 'セーラー服', 'sailor collar': 'セーラー襟', 'blazer': 'ブレザー',
    'dress': 'ドレス', 'long dress': 'ロングドレス', 'sundress': 'サンドレス', 'wedding dress': 'ウェディングドレス',
    'evening gown': 'イブニングドレス', 'shirt': 'シャツ', 'white shirt': '白シャツ', 't-shirt': 'Tシャツ',
    'blouse': 'ブラウス', 'sweater': 'セーター', 'hoodie': 'パーカー', 'jacket': 'ジャケット', 'coat': 'コート',
    'skirt': 'スカート', 'pleated skirt': 'プリーツスカート', 'miniskirt': 'ミニスカート', 'long skirt': 'ロングスカート',
    'pants': 'ズボン', 'jeans': 'ジーンズ', 'shorts': 'ショートパンツ', 'kimono': '着物', 'yukata': '浴衣',
    'hakama': '袴', 'chinese dress': 'チャイナドレス', 'maid': 'メイド', 'maid dress': 'メイド服', 'apron': 'エプロン',
    'swimsuit': '水着', 'bikini': 'ビキニ', 'one-piece swimsuit': 'ワンピース水着', 'school swimsuit': 'スクール水着',
    'nightgown': 'ネグリジェ', 'pajamas': 'パジャマ', 'bodysuit': 'ボディスーツ', 'armor': '鎧', 'gothic lolita': 'ゴスロリ',
    'lolita fashion': 'ロリータ服', 'casual': 'カジュアル服', 'suit': 'スーツ', 'formal': 'フォーマル',
    'bra': 'ブラ', 'panties': '下着', 'underwear': '下着', 'lingerie': 'ランジェリー'
  },
  '靴下・靴': {
    'thighhighs': 'ニーソックス', 'thigh highs': 'ニーソックス', 'stockings': 'ストッキング', 'pantyhose': 'タイツ',
    'socks': '靴下', 'kneehighs': 'ハイソックス', 'garter belt': 'ガーターベルト',
    'shoes': '靴', 'boots': 'ブーツ', 'high heels': 'ハイヒール', 'sneakers': 'スニーカー', 'sandals': 'サンダル'
  },
  '装飾・小物': {
    'gloves': '手袋', 'fingerless gloves': '指なし手袋', 'hat': '帽子', 'beret': 'ベレー帽', 'cap': 'キャップ',
    'witch hat': '魔女帽子', 'ribbon': 'リボン', 'hair ribbon': '髪リボン', 'bow': 'リボン結び', 'hairband': 'カチューシャ',
    'hair ornament': '髪飾り', 'hairclip': 'ヘアピン', 'headband': 'ヘアバンド', 'glasses': 'メガネ',
    'necktie': 'ネクタイ', 'bowtie': '蝶ネクタイ', 'scarf': 'マフラー', 'necklace': 'ネックレス', 'earrings': 'イヤリング',
    'choker': 'チョーカー', 'bracelet': 'ブレスレット', 'cape': 'マント', 'wings': '翼', 'angel wings': '天使の翼',
    'animal ears': '獣耳', 'cat ears': '猫耳', 'fox ears': '狐耳', 'tail': '尻尾', 'halo': '天使の輪', 'horns': '角'
  },
  'ポーズ': {
    'standing': '立ち姿', 'sitting': '座り', 'kneeling': 'ひざまずき', 'lying': '横たわり', 'lying down': '寝そべり',
    'on back': '仰向け', 'on stomach': 'うつ伏せ', 'walking': '歩く', 'running': '走る', 'jumping': '跳ぶ',
    'crossed arms': '腕組み', 'crossed legs': '足を組む', 'arms up': '両手を上げる', 'hand on hip': '腰に手',
    'hands on hips': '両手を腰に', 'waving': '手を振る', 'peace sign': 'ピースサイン', 'stretching': '伸び',
    'from behind': '後ろから', 'looking over shoulder': '肩越しに見る', 'outstretched arm': '腕を伸ばす',
    'holding': '持っている', 'hand up': '手を上げる', 'own hands together': '手を合わせる', 'dynamic pose': '躍動的なポーズ'
  },
  '構図': {
    'full body': '全身', 'upper body': '上半身', 'cowboy shot': '腰から上', 'portrait': '肖像', 'close-up': 'アップ',
    'face focus': '顔アップ', 'wide shot': '引きの構図', 'from above': '俯瞰', 'from below': 'あおり',
    'from side': '横から', 'dutch angle': '斜め構図', 'depth of field': '被写界深度', 'bokeh': 'ボケ',
    'blurry background': '背景ぼかし', 'motion blur': '動きのブレ', 'wide angle': '広角', 'fisheye': '魚眼'
  },
  '背景': {
    'simple background': 'シンプルな背景', 'white background': '白背景', 'black background': '黒背景',
    'gradient background': 'グラデ背景', 'transparent background': '透過背景', 'detailed background': '緻密な背景',
    'outdoors': '屋外', 'indoors': '屋内', 'nature': '自然', 'forest': '森', 'sky': '空', 'blue sky': '青空',
    'clouds': '雲', 'cloudy sky': '曇り空', 'night': '夜', 'night sky': '夜空', 'starry sky': '星空', 'stars': '星',
    'moon': '月', 'full moon': '満月', 'sunset': '夕焼け', 'sunrise': '日の出', 'ocean': '海', 'beach': '砂浜',
    'water': '水', 'underwater': '水中', 'rain': '雨', 'snow': '雪', 'cherry blossoms': '桜', 'flowers': '花',
    'field': '草原', 'flower field': '花畑', 'mountain': '山', 'city': '街', 'cityscape': '街並み', 'street': '通り',
    'building': '建物', 'room': '部屋', 'bedroom': '寝室', 'classroom': '教室', 'library': '図書館', 'cafe': 'カフェ',
    'garden': '庭', 'window': '窓', 'sunlight': '日差し', 'fantasy': 'ファンタジー', 'scenery': '風景'
  },
  '光・色': {
    'lighting': '照明', 'soft lighting': '柔らかな光', 'dramatic lighting': '劇的な光', 'rim light': 'リムライト',
    'backlighting': '逆光', 'volumetric lighting': '光芒', 'god rays': '天使のはしご', 'glowing': '発光',
    'light particles': '光の粒子', 'lens flare': 'レンズフレア', 'sunbeam': '陽光', 'ambient light': '環境光',
    'colorful': 'カラフル', 'vibrant colors': '鮮やかな色', 'pastel colors': 'パステルカラー', 'monochrome': 'モノクロ',
    'sepia': 'セピア', 'vivid': '鮮烈', 'muted colors': '落ち着いた色', 'high contrast': '高コントラスト',
    'shiny': '光沢', 'reflection': '反射', 'shadow': '影', 'glow': '輝き'
  },
  'ネガティブ': {
    'worst quality': '最低品質', 'low quality': '低品質', 'normal quality': '並品質', 'bad quality': '粗悪品質',
    'jpeg artifacts': '圧縮ノイズ', 'blurry': 'ぼやけ', 'lowres': '低解像度', 'pixelated': 'ドット荒れ',
    'bad anatomy': '崩れた人体', 'bad hands': '崩れた手', 'bad proportions': '崩れた比率', 'bad feet': '崩れた足',
    'extra fingers': '指が多い', 'missing fingers': '指が欠損', 'fused fingers': '指の癒着', 'extra digits': '指過多',
    'extra limbs': '手足が多い', 'missing limbs': '手足の欠損', 'extra arms': '腕が多い', 'extra legs': '脚が多い',
    'malformed limbs': '奇形の手足', 'mutated hands': '変形した手', 'deformed': '奇形', 'disfigured': '崩れた造形',
    'mutation': '突然変異', 'ugly': '醜い', 'poorly drawn face': '雑な顔', 'poorly drawn hands': '雑な手',
    'long neck': '長すぎる首', 'cross-eyed': '寄り目', 'watermark': '透かし', 'signature': '署名',
    'username': 'ユーザー名', 'text': '文字', 'logo': 'ロゴ', 'artist name': '作者名', 'cropped': '見切れ',
    'out of frame': '枠外', 'error': 'エラー', 'duplicate': '重複', 'nsfw': '成人向け要素', 'easynegative': '汎用ネガ(embedding)',
    'ng_deepnegative_v1_75t': '汎用ネガ(embedding)', 'badhandv4': '手崩れ対策(embedding)', 'verybadimagenegative': '低品質対策(embedding)'
  }
};
