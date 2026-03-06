/**
 * 野尻湖発掘 化石・遺物入力アプリ 共有保存API（Google Apps Script）
 *
 * 使い方:
 * 1. Googleドライブで新しいApps Scriptプロジェクトを作成
 * 2. このコードを貼り付けて保存
 * 3. 必要なら TARGET_FOLDER_ID に共有ドライブ/共有フォルダIDを設定
 * 4. 「デプロイ > 新しいデプロイ > ウェブアプリ」
 *    - 実行ユーザー: 自分
 *    - アクセスできるユーザー: リンクを知っている全員
 * 5. 発行されたWebアプリURLを本アプリの「共有保存を設定」に入力
 */

const STORAGE_FILE_NAME = "nojiri_kaseki_shared_state.json";
const STORAGE_FILE_ID_KEY = "NOJIRI_STORAGE_FILE_ID";
const TARGET_FOLDER_ID = ""; // 共有ドライブ/共有フォルダIDを使う場合だけ設定

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || "load";
  if (action === "load") {
    return jsonResponse_(loadPayload_());
  }
  if (action === "ping") {
    return jsonResponse_({ ok: true, service: "nojiri-kaseki-cloud" });
  }
  return jsonResponse_({ ok: false, error: "unsupported action" });
}

function doPost(e) {
  const action = (e && e.parameter && e.parameter.action) || "";
  const payloadText = (e && e.parameter && e.parameter.payload) || "{}";
  let payload;
  try {
    payload = JSON.parse(payloadText);
  } catch (_error) {
    return jsonResponse_({ ok: false, error: "invalid payload" });
  }

  if (action === "save") {
    if (!payload || typeof payload !== "object" || !payload.state) {
      return jsonResponse_({ ok: false, error: "state is required" });
    }
    const saved = savePayload_(payload.state, payload.clientId || "");
    return jsonResponse_(saved);
  }
  if (action === "load") {
    return jsonResponse_(loadPayload_());
  }
  return jsonResponse_({ ok: false, error: "unsupported action" });
}

function loadPayload_() {
  const file = getStorageFile_();
  const text = file.getBlob().getDataAsString("UTF-8") || "";
  if (!text) {
    return {
      ok: true,
      state: createInitialState_(),
      updatedAt: "",
    };
  }
  try {
    const parsed = JSON.parse(text);
    const state = parsed && parsed.state ? parsed.state : createInitialState_();
    return {
      ok: true,
      state,
      updatedAt: parsed.updatedAt || "",
      savedBy: parsed.savedBy || "",
    };
  } catch (_error) {
    return {
      ok: true,
      state: createInitialState_(),
      updatedAt: "",
    };
  }
}

function savePayload_(state, clientId) {
  const file = getStorageFile_();
  const payload = {
    state: state,
    updatedAt: new Date().toISOString(),
    savedBy: String(clientId || ""),
  };
  file.setContent(JSON.stringify(payload));
  return {
    ok: true,
    updatedAt: payload.updatedAt,
  };
}

function getStorageFile_() {
  const props = PropertiesService.getScriptProperties();
  const cachedId = props.getProperty(STORAGE_FILE_ID_KEY);
  if (cachedId) {
    try {
      return DriveApp.getFileById(cachedId);
    } catch (_error) {
      // fallback to recreate
    }
  }

  const folder = getTargetFolder_();
  const existing = folder.getFilesByName(STORAGE_FILE_NAME);
  if (existing.hasNext()) {
    const file = existing.next();
    props.setProperty(STORAGE_FILE_ID_KEY, file.getId());
    return file;
  }

  const file = folder.createFile(STORAGE_FILE_NAME, JSON.stringify({
    state: createInitialState_(),
    updatedAt: "",
    savedBy: "",
  }), MimeType.PLAIN_TEXT);
  props.setProperty(STORAGE_FILE_ID_KEY, file.getId());
  return file;
}

function getTargetFolder_() {
  const folderId = String(TARGET_FOLDER_ID || "").trim();
  if (folderId) {
    return DriveApp.getFolderById(folderId);
  }
  return DriveApp.getRootFolder();
}

function createInitialState_() {
  return {
    site: {
      kuwaku: "24-Ⅰ--",
      kuwakuBlock: "",
      kuwakuNo: "",
      levelHeight: "",
      date: "",
      team: "",
      teamOther: "",
      teamLead: "",
      recorder: "",
      updatedAt: "",
    },
    records: [],
  };
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
