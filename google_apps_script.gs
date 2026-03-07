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
  const lock = LockService.getScriptLock();
  let isLocked = false;
  try {
    lock.waitLock(30000);
    isLocked = true;
    const file = getStorageFile_();
    const current = readStoragePayload_(file);
    const mergedState = mergeStates_(current.state, state);
    const payload = {
      state: mergedState,
      updatedAt: new Date().toISOString(),
      savedBy: String(clientId || ""),
    };
    file.setContent(JSON.stringify(payload));
    return {
      ok: true,
      updatedAt: payload.updatedAt,
    };
  } catch (_error) {
    return {
      ok: false,
      error: "save lock timeout",
    };
  } finally {
    if (isLocked) {
      lock.releaseLock();
    }
  }
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
      kuwakuHeadA: "24",
      kuwakuHeadB: "Ⅰ",
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

function readStoragePayload_(file) {
  const text = file.getBlob().getDataAsString("UTF-8") || "";
  if (!text) {
    return {
      state: createInitialState_(),
      updatedAt: "",
      savedBy: "",
    };
  }
  try {
    const parsed = JSON.parse(text);
    return {
      state: parsed && parsed.state ? parsed.state : createInitialState_(),
      updatedAt: str_(parsed && parsed.updatedAt),
      savedBy: str_(parsed && parsed.savedBy),
    };
  } catch (_error) {
    return {
      state: createInitialState_(),
      updatedAt: "",
      savedBy: "",
    };
  }
}

function mergeStates_(baseStateRaw, incomingStateRaw) {
  const baseState = normalizeStateForMerge_(baseStateRaw);
  const incomingState = normalizeStateForMerge_(incomingStateRaw);
  const mergedSite = mergeSiteForCloud_(baseState.site, incomingState.site);
  const mergedRecords = mergeRecordsById_(baseState.records, incomingState.records);
  return {
    site: mergedSite,
    records: mergedRecords,
  };
}

function normalizeStateForMerge_(stateRaw) {
  const base = createInitialState_();
  if (!stateRaw || typeof stateRaw !== "object") {
    return base;
  }
  const siteRaw = stateRaw.site && typeof stateRaw.site === "object" ? stateRaw.site : {};
  const recordsRaw = Array.isArray(stateRaw.records) ? stateRaw.records : [];
  return {
    site: {
      kuwaku: str_(siteRaw.kuwaku) || base.site.kuwaku,
      kuwakuHeadA: str_(siteRaw.kuwakuHeadA) || base.site.kuwakuHeadA,
      kuwakuHeadB: str_(siteRaw.kuwakuHeadB) || base.site.kuwakuHeadB,
      kuwakuBlock: str_(siteRaw.kuwakuBlock),
      kuwakuNo: str_(siteRaw.kuwakuNo),
      levelHeight: str_(siteRaw.levelHeight),
      date: str_(siteRaw.date),
      team: str_(siteRaw.team),
      teamOther: str_(siteRaw.teamOther),
      teamLead: str_(siteRaw.teamLead),
      recorder: str_(siteRaw.recorder),
      updatedAt: str_(siteRaw.updatedAt),
    },
    records: recordsRaw
      .filter((item) => item && typeof item === "object")
      .map((item) => item),
  };
}

function mergeSiteForCloud_(baseSiteRaw, incomingSiteRaw) {
  const baseSite = baseSiteRaw && typeof baseSiteRaw === "object" ? baseSiteRaw : {};
  const incomingSite = incomingSiteRaw && typeof incomingSiteRaw === "object" ? incomingSiteRaw : {};
  const preferIncoming = isIsoTimestampGreaterOrEqual_(incomingSite.updatedAt, baseSite.updatedAt);
  const primary = preferIncoming ? incomingSite : baseSite;
  const secondary = preferIncoming ? baseSite : incomingSite;
  return {
    kuwaku: str_(primary.kuwaku) || str_(secondary.kuwaku) || "24-Ⅰ--",
    kuwakuHeadA: str_(primary.kuwakuHeadA) || str_(secondary.kuwakuHeadA) || "24",
    kuwakuHeadB: str_(primary.kuwakuHeadB) || str_(secondary.kuwakuHeadB) || "Ⅰ",
    kuwakuBlock: str_(primary.kuwakuBlock) || str_(secondary.kuwakuBlock),
    kuwakuNo: str_(primary.kuwakuNo) || str_(secondary.kuwakuNo),
    levelHeight: str_(primary.levelHeight) || str_(secondary.levelHeight),
    date: str_(primary.date) || str_(secondary.date),
    team: str_(primary.team) || str_(secondary.team),
    teamOther: str_(primary.teamOther) || str_(secondary.teamOther),
    teamLead: str_(primary.teamLead) || str_(secondary.teamLead),
    recorder: str_(primary.recorder) || str_(secondary.recorder),
    updatedAt: pickLatestIsoTimestamp_(baseSite.updatedAt, incomingSite.updatedAt),
  };
}

function mergeRecordsById_(baseRecordsRaw, incomingRecordsRaw) {
  const merged = new Map();
  const baseRecords = Array.isArray(baseRecordsRaw) ? baseRecordsRaw : [];
  const incomingRecords = Array.isArray(incomingRecordsRaw) ? incomingRecordsRaw : [];

  baseRecords.forEach((record) => upsertRecordById_(merged, record, false));
  incomingRecords.forEach((record) => upsertRecordById_(merged, record, true));

  return Array.from(merged.values());
}

function upsertRecordById_(map, candidateRaw, preferCandidateOnTie) {
  const candidate = candidateRaw && typeof candidateRaw === "object" ? candidateRaw : null;
  const recordId = str_(candidate && candidate.id);
  if (!candidate || !recordId) {
    return;
  }
  const existing = map.get(recordId);
  if (!existing) {
    map.set(recordId, candidate);
    return;
  }
  map.set(recordId, chooseNewerRecord_(existing, candidate, preferCandidateOnTie));
}

function chooseNewerRecord_(existingRaw, incomingRaw, preferIncomingOnTie) {
  const existing = existingRaw && typeof existingRaw === "object" ? existingRaw : {};
  const incoming = incomingRaw && typeof incomingRaw === "object" ? incomingRaw : {};
  const existingMs = getRecordUpdatedAtMs_(existing);
  const incomingMs = getRecordUpdatedAtMs_(incoming);
  let winner = existing;
  let loser = incoming;

  if (isFinite_(incomingMs) && isFinite_(existingMs)) {
    if (incomingMs > existingMs || (incomingMs === existingMs && preferIncomingOnTie)) {
      winner = incoming;
      loser = existing;
    }
  } else if (isFinite_(incomingMs) && !isFinite_(existingMs)) {
    winner = incoming;
    loser = existing;
  } else if (!isFinite_(incomingMs) && !isFinite_(existingMs) && preferIncomingOnTie) {
    winner = incoming;
    loser = existing;
  }

  const mergedHistory = mergeRecordHistory_(existing.history, incoming.history);
  return Object.assign({}, loser, winner, mergedHistory.length ? { history: mergedHistory } : {});
}

function mergeRecordHistory_(historyARaw, historyBRaw) {
  const map = new Map();
  const listA = Array.isArray(historyARaw) ? historyARaw : [];
  const listB = Array.isArray(historyBRaw) ? historyBRaw : [];

  const upsert = (entryRaw) => {
    const entry = entryRaw && typeof entryRaw === "object" ? entryRaw : null;
    if (!entry) {
      return;
    }
    const key = str_(entry.id) || `${str_(entry.at)}::${str_(entry.action)}::${str_(entry.content)}`;
    if (!key) {
      return;
    }
    const existing = map.get(key);
    if (!existing) {
      map.set(key, entry);
      return;
    }
    const existingMs = parseIsoMs_(existing.at);
    const incomingMs = parseIsoMs_(entry.at);
    if (isFinite_(incomingMs) && (!isFinite_(existingMs) || incomingMs >= existingMs)) {
      map.set(key, entry);
    }
  };

  listA.forEach(upsert);
  listB.forEach(upsert);

  return Array.from(map.values())
    .sort((a, b) => {
      const aMs = parseIsoMs_(a && a.at);
      const bMs = parseIsoMs_(b && b.at);
      if (isFinite_(aMs) && isFinite_(bMs) && aMs !== bMs) {
        return aMs - bMs;
      }
      return str_(a && a.id).localeCompare(str_(b && b.id));
    })
    .slice(-50);
}

function getRecordUpdatedAtMs_(recordRaw) {
  const record = recordRaw && typeof recordRaw === "object" ? recordRaw : {};
  const updatedMs = parseIsoMs_(record.updatedAt);
  if (isFinite_(updatedMs)) {
    return updatedMs;
  }
  const createdMs = parseIsoMs_(record.createdAt);
  return isFinite_(createdMs) ? createdMs : NaN;
}

function isIsoTimestampGreaterOrEqual_(tsA, tsB) {
  const aMs = parseIsoMs_(tsA);
  const bMs = parseIsoMs_(tsB);
  if (isFinite_(aMs) && isFinite_(bMs)) {
    return aMs >= bMs;
  }
  if (isFinite_(aMs)) {
    return true;
  }
  if (isFinite_(bMs)) {
    return false;
  }
  return true;
}

function pickLatestIsoTimestamp_(tsA, tsB) {
  const aMs = parseIsoMs_(tsA);
  const bMs = parseIsoMs_(tsB);
  if (isFinite_(aMs) && isFinite_(bMs)) {
    return new Date(Math.max(aMs, bMs)).toISOString();
  }
  if (isFinite_(aMs)) {
    return new Date(aMs).toISOString();
  }
  if (isFinite_(bMs)) {
    return new Date(bMs).toISOString();
  }
  return "";
}

function parseIsoMs_(valueRaw) {
  const ms = Date.parse(str_(valueRaw));
  return isFinite_(ms) ? ms : NaN;
}

function isFinite_(valueRaw) {
  return Number.isFinite(Number(valueRaw));
}

function str_(valueRaw) {
  if (valueRaw === null || valueRaw === undefined) {
    return "";
  }
  return String(valueRaw).trim();
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
