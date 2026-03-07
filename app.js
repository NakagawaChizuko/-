const STORAGE_KEY = "nojiri-kaseki-mobile-v1";
const CLOUD_ENDPOINT_KEY = "nojiri-kaseki-cloud-endpoint-v1";
const CLOUD_CLIENT_ID_KEY = "nojiri-kaseki-cloud-client-id-v1";
const CLOUD_PULL_INTERVAL_MS = 20000;
const CLOUD_SAVE_DEBOUNCE_MS = 900;
const CLOUD_AUTO_PULL_ENABLED = false;
const DEFAULT_SPECIMEN_PREFIX = "m";
const SPECIMEN_CATEGORY_MAP = {
  m: "哺乳類",
  a: "分析用資料",
  b: "植物",
  l: "生痕",
  s: "貝類",
  i: "昆虫",
  g: "人類考古",
  h: "その他",
};
const VALID_SPECIMEN_PREFIXES = new Set(Object.keys(SPECIMEN_CATEGORY_MAP));
const ANALYSIS_TYPE_MAP = {
  A: "火山灰",
  C: "14C",
  M: "古地磁気",
  F: "フィッショントラック",
  P: "花粉",
  B: "植物",
  I: "昆虫",
  D: "珪藻",
  R: "粒度",
  S: "貝類",
  H: "その他",
  MG: "はぎとり資料",
};
const HISTORY_SNAPSHOT_FIELDS = [
  { key: "specimenNo", label: "標本番号" },
  { key: "nameMemo", label: "名称" },
  { key: "category", label: "分類" },
  { key: "layerName", label: "地層名" },
  { key: "unit", label: "ユニット" },
  { key: "detail", label: "細別" },
  { key: "layerPosition", label: "地層中の位置" },
];
const HISTORY_SNAPSHOT_FIELD_KEYS = new Set(HISTORY_SNAPSHOT_FIELDS.map((field) => field.key));
const PRESET_LAYER_NAMES = [
  "1.芙蓉湖砂シルト部層",
  "2.立が鼻砂部層",
  "3.海端砂シルト部層",
  "4.その他",
];
const OTHER_LAYER_NAME = "4.その他";
const LEGACY_LAYER_NAME_ALIASES = {
  "2.立が花砂部層": "2.立が鼻砂部層",
};
const PRESET_TEAMS = ["1", "2", "3", "4", "その他"];
const OTHER_TEAM_NAME = "その他";
const DEFAULT_KUWAKU_HEAD_A = "24";
const DEFAULT_KUWAKU_HEAD_B = "Ⅰ";
const DEFAULT_KUWAKU = `${DEFAULT_KUWAKU_HEAD_A}-${DEFAULT_KUWAKU_HEAD_B}--`;
const ALL_GRIDS_VALUE = "__KUWAKU_ALL__";
const EMPTY_KUWAKU_VALUE = "__KUWAKU_EMPTY__";
const PLAN_SIZE_CM = 400;
const ALL_UNITS_VALUE = "__UNIT_ALL__";
const EMPTY_UNIT_VALUE = "__UNIT_EMPTY__";
const ALL_DETAILS_VALUE = "__DETAIL_ALL__";
const EMPTY_DETAIL_VALUE = "__DETAIL_EMPTY__";
const SPECIMEN_POINT_COLORS = {
  m: "#d62828",
  a: "#5b21b6",
  b: "#2a9d8f",
  l: "#f4a261",
  s: "#457b9d",
  i: "#6d597a",
  g: "#8f5a2b",
  h: "#6b7280",
};
const PHOTO_COMPRESSION_STEPS = [
  { maxLength: 1280, quality: 0.72 },
  { maxLength: 960, quality: 0.62 },
  { maxLength: 720, quality: 0.54 },
  { maxLength: 560, quality: 0.46 },
];

const createInitialState = () => ({
  site: {
    kuwaku: DEFAULT_KUWAKU,
    kuwakuHeadA: DEFAULT_KUWAKU_HEAD_A,
    kuwakuHeadB: DEFAULT_KUWAKU_HEAD_B,
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
});

let stateNeedsRewriteAfterLoad = false;
let state = loadState();
let editingRecordId = null;
let currentSectionDiagrams = [];
let currentPhotos = [];
let selectedCardRecordId = "";
let selectedOutputKuwaku = ALL_GRIDS_VALUE;
let selectedPlanKuwaku = "";
let selectedPlanUnit = "";
let selectedPlanDetail = ALL_DETAILS_VALUE;
let outputListSortKey = "kuwaku";
let outputListSortDirection = "asc";
let isOverwriteMode = false;
let overwriteOriginalRecord = null;
let toastTimer = null;
let quotaRecoveryInProgress = false;
let cloudEndpoint = loadCloudEndpoint();
let cloudClientId = loadOrCreateCloudClientId();
let cloudSaveTimer = null;
let cloudPullTimer = null;
let cloudPushInProgress = false;
let cloudPullInProgress = false;
let cloudLastSyncedAt = "";
let cloudLastPulledAt = "";
let cloudLastErrorAt = 0;

const tabButtons = document.querySelectorAll(".tab-button");
const tabPanels = document.querySelectorAll(".tab-panel");
const siteForm = document.getElementById("site-form");
const recordForm = document.getElementById("record-form");
const recordFormHost = document.getElementById("record-form-host");
const editRecordFormHost = document.getElementById("edit-record-form-host");
const editHistoryPanel = document.getElementById("edit-history-panel");
const editHistoryList = document.getElementById("edit-history-list");
const editKuwakuHeadAInput = document.getElementById("edit-kuwaku-head-a");
const editKuwakuHeadBInput = document.getElementById("edit-kuwaku-head-b");
const editKuwakuBlockInput = document.getElementById("edit-kuwaku-block");
const editKuwakuNoInput = document.getElementById("edit-kuwaku-no");
const editLevelHeightInput = document.getElementById("edit-level-height");
const editDateInput = document.getElementById("edit-date");
const editTeamInput = document.getElementById("edit-team");
const editTeamOtherInput = document.getElementById("edit-team-other");
const editTeamLeadInput = document.getElementById("edit-team-lead");
const editRecorderInput = document.getElementById("edit-recorder");
const recordIdInput = document.getElementById("record-id-input");
const recordSubmitBtn = document.getElementById("record-submit-btn");
const recordResetBtn = document.getElementById("record-reset-btn");
const recordTableBody = document.getElementById("record-table-body");
const outputListBody = document.getElementById("output-list-body");
const outputListTable = document.getElementById("output-list-table");
const cardOutputList = document.getElementById("card-output-list");
const outputKuwakuSelect = document.getElementById("output-kuwaku-select");
const planKuwakuSelect = document.getElementById("plan-kuwaku-select");
const planUnitSelect = document.getElementById("plan-unit-select");
const planDetailSelect = document.getElementById("plan-detail-select");
const planMapLegend = document.getElementById("plan-map-legend");
const planMapWrap = document.getElementById("plan-map-wrap");
const planKuwakuInfo = document.getElementById("plan-kuwaku-info");

const specimenTabButtons = document.querySelectorAll(".specimen-tab");
const specimenPrefixInput = document.getElementById("specimen-prefix-input");
const specimenSerialInput = document.getElementById("specimen-serial-input");
const specimenNoInput = document.getElementById("specimen-no-input");
const specimenPrefixLabel = document.getElementById("specimen-prefix-label");
const specimenDuplicateWarning = document.getElementById("specimen-duplicate-warning");
const analysisTypeRow = document.getElementById("analysis-type-row");
const analysisTypeSelect = document.getElementById("analysis-type-select");

const dirTabButtons = document.querySelectorAll(".dir-tab");
const nsDirInput = document.getElementById("ns-dir-input");
const ewDirInput = document.getElementById("ew-dir-input");
const layerTabButtons = document.querySelectorAll(".layer-tab");
const layerNameInput = document.getElementById("layer-name-input");
const layerOtherInput = document.getElementById("layer-other-input");
const teamOtherInput = document.getElementById("team-other-input");

const sectionDiagramCameraInput = document.getElementById("section-diagram-camera-input");
const sectionDiagramInput = document.getElementById("section-diagram-input");
const sectionDiagramList = document.getElementById("section-diagram-list");
const photoCameraInput = document.getElementById("photo-camera-input");
const photoInput = document.getElementById("photo-input");
const photoList = document.getElementById("photo-list");

const exportListCsvBtn = document.getElementById("export-list-csv-btn");
const exportCardCsvBtn = document.getElementById("export-card-csv-btn");
const exportJsonBtn = document.getElementById("export-json-btn");
const importJsonInput = document.getElementById("import-json-input");
const cloudEndpointInput = document.getElementById("cloud-endpoint-input");
const cloudConnectBtn = document.getElementById("cloud-connect-btn");
const cloudSyncBtn = document.getElementById("cloud-sync-btn");
const cloudDisableBtn = document.getElementById("cloud-disable-btn");
const cloudStatusEl = document.getElementById("cloud-status");
const toastEl = document.getElementById("toast");

initialize();

function initialize() {
  bindEvents();
  if (stateNeedsRewriteAfterLoad) {
    persist();
    stateNeedsRewriteAfterLoad = false;
  }
  initCloudControls();
  hydrateSiteForm();
  resetRecordForm({ showMessage: false });
  renderRecordTable();
  renderOutputs();
  void bootstrapCloudSync();
}

function bindEvents() {
  tabButtons.forEach((button) => {
    button.addEventListener("click", () => setActiveTab(button.dataset.tab));
  });

  if (cloudConnectBtn) {
    cloudConnectBtn.addEventListener("click", () => {
      void handleCloudConnect();
    });
  }
  if (cloudSyncBtn) {
    cloudSyncBtn.addEventListener("click", () => {
      void handleCloudManualReload();
    });
  }
  if (cloudDisableBtn) {
    cloudDisableBtn.addEventListener("click", () => {
      disableCloudSync({ showToastMessage: true });
    });
  }
  if (cloudEndpointInput) {
    cloudEndpointInput.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        void handleCloudConnect();
      }
    });
  }

  recordForm.addEventListener("input", handleRecordFormFieldEdit);
  recordForm.addEventListener("change", handleRecordFormFieldEdit);

  siteForm.addEventListener("submit", (event) => {
    event.preventDefault();
    const formData = new FormData(siteForm);
    const kuwakuHeadA = normalizeKuwakuHeadA(formData.get("kuwakuHeadA"));
    const kuwakuHeadB = normalizeKuwakuHeadB(formData.get("kuwakuHeadB"));
    const kuwakuBlock = normalizeKuwakuBlock(formData.get("kuwakuBlock"));
    const kuwakuNo = normalizeKuwakuNo(formData.get("kuwakuNo"));
    const teamState = normalizeTeamState(value(formData.get("team")), value(formData.get("teamOther")));
    const nextSiteKuwaku = buildKuwaku(kuwakuHeadA, kuwakuHeadB, kuwakuBlock, kuwakuNo);
    const normalizedKuwaku = parseKuwaku(nextSiteKuwaku);

    state.site = {
      kuwaku: nextSiteKuwaku,
      kuwakuHeadA: normalizedKuwaku.headA,
      kuwakuHeadB: normalizedKuwaku.headB,
      kuwakuBlock: normalizedKuwaku.block,
      kuwakuNo: normalizedKuwaku.no,
      levelHeight: value(formData.get("levelHeight")),
      date: value(formData.get("date")),
      team: teamState.team,
      teamOther: teamState.teamOther,
      teamLead: value(formData.get("teamLead")),
      recorder: value(formData.get("recorder")),
      updatedAt: nowIso(),
    };
    selectedOutputKuwaku = kuwakuValueForSelect(nextSiteKuwaku);
    selectedPlanKuwaku = kuwakuValueForSelect(nextSiteKuwaku);
    persist("区画（グリッド）情報を保存しました");
    renderRecordTable();
    renderOutputs();
  });

  siteForm.elements.team.addEventListener("change", () => {
    syncTeamOtherInput(siteForm.elements.team.value);
  });
  ["kuwakuHeadA", "kuwakuHeadB", "kuwakuBlock", "kuwakuNo"].forEach((name) => {
    const input = siteForm.elements[name];
    if (!(input instanceof Element)) {
      return;
    }
    input.addEventListener("input", updateDuplicateSpecimenWarning);
    input.addEventListener("change", updateDuplicateSpecimenWarning);
  });
  if (editTeamInput) {
    editTeamInput.addEventListener("change", () => {
      syncEditTeamOtherInput(editTeamInput.value);
      editTeamInput.classList.remove("overwrite-updated");
      if (editTeamOtherInput) {
        editTeamOtherInput.classList.remove("overwrite-updated");
      }
    });
  }

  specimenTabButtons.forEach((button) => {
    button.addEventListener("click", () => {
      activateSpecimenPrefix(button.dataset.prefix);
      updateSpecimenNoFromParts();
    });
  });

  specimenSerialInput.addEventListener("input", () => {
    updateSpecimenNoFromParts();
  });
  if (analysisTypeSelect) {
    analysisTypeSelect.addEventListener("change", () => {
      analysisTypeSelect.value = normalizeAnalysisType(analysisTypeSelect.value);
      analysisTypeSelect.classList.remove("overwrite-updated");
    });
  }

  dirTabButtons.forEach((button) => {
    button.addEventListener("click", () => {
      activateDirectionTab(button.dataset.group, button.dataset.value);
    });
  });

  layerTabButtons.forEach((button) => {
    button.addEventListener("click", () => {
      clearLayerSavedTabState();
      layerOtherInput.classList.remove("saved-carry-value");
      activateLayerTab(button.dataset.layer);
    });
  });

  recordForm.addEventListener("submit", (event) => {
    event.preventDefault();

    const isEditTab = getActiveTabId() === "edit-tab";
    let recordKuwaku = "";
    let siteSnapshot = null;
    let editSiteSnapshot = null;

    if (isEditTab) {
      const headA = normalizeKuwakuHeadA(editKuwakuHeadAInput?.value);
      const headB = normalizeKuwakuHeadB(editKuwakuHeadBInput?.value);
      const block = normalizeKuwakuBlock(editKuwakuBlockInput?.value);
      const no = normalizeKuwakuNo(editKuwakuNoInput?.value);
      recordKuwaku = buildKuwaku(headA, headB, block, no);
      const editTeamState = normalizeTeamState(value(editTeamInput?.value), value(editTeamOtherInput?.value));
      editSiteSnapshot = {
        levelHeight: value(editLevelHeightInput?.value),
        date: value(editDateInput?.value),
        team: editTeamState.team,
        teamOther: editTeamState.teamOther,
        teamLead: value(editTeamLeadInput?.value),
        recorder: value(editRecorderInput?.value),
      };
    } else {
      // 入力画面は従来どおり、区画（グリッド）情報フォームの現在値を反映。
      const siteFormData = new FormData(siteForm);
      const siteKuwakuHeadA = normalizeKuwakuHeadA(siteFormData.get("kuwakuHeadA"));
      const siteKuwakuHeadB = normalizeKuwakuHeadB(siteFormData.get("kuwakuHeadB"));
      const siteKuwakuBlock = normalizeKuwakuBlock(siteFormData.get("kuwakuBlock"));
      const siteKuwakuNo = normalizeKuwakuNo(siteFormData.get("kuwakuNo"));
      const siteTeamState = normalizeTeamState(value(siteFormData.get("team")), value(siteFormData.get("teamOther")));
      const nextSiteKuwaku = buildKuwaku(siteKuwakuHeadA, siteKuwakuHeadB, siteKuwakuBlock, siteKuwakuNo);
      const normalizedKuwaku = parseKuwaku(nextSiteKuwaku);
      siteSnapshot = {
        kuwaku: nextSiteKuwaku,
        kuwakuHeadA: normalizedKuwaku.headA,
        kuwakuHeadB: normalizedKuwaku.headB,
        kuwakuBlock: normalizedKuwaku.block,
        kuwakuNo: normalizedKuwaku.no,
        levelHeight: value(siteFormData.get("levelHeight")),
        date: value(siteFormData.get("date")),
        team: siteTeamState.team,
        teamOther: siteTeamState.teamOther,
        teamLead: value(siteFormData.get("teamLead")),
        recorder: value(siteFormData.get("recorder")),
      };
      recordKuwaku = nextSiteKuwaku;
    }

    const formData = new FormData(recordForm);
    if (!isEditTab) {
      const validationMessage = validateInputRequiredFields(siteSnapshot, formData);
      if (validationMessage) {
        showToast(validationMessage);
        return;
      }
      state.site = {
        kuwaku: siteSnapshot.kuwaku,
        kuwakuHeadA: siteSnapshot.kuwakuHeadA,
        kuwakuHeadB: siteSnapshot.kuwakuHeadB,
        kuwakuBlock: siteSnapshot.kuwakuBlock,
        kuwakuNo: siteSnapshot.kuwakuNo,
        levelHeight: siteSnapshot.levelHeight,
        date: siteSnapshot.date,
        team: siteSnapshot.team,
        teamOther: siteSnapshot.teamOther,
        teamLead: siteSnapshot.teamLead,
        recorder: siteSnapshot.recorder,
        updatedAt: nowIso(),
      };
      selectedOutputKuwaku = kuwakuValueForSelect(siteSnapshot.kuwaku);
      selectedPlanKuwaku = kuwakuValueForSelect(siteSnapshot.kuwaku);
    }

    const specimenPrefix = normalizeSpecimenPrefix(value(formData.get("specimenPrefix")));
    const specimenSerial = compactNoSpaceValue(formData.get("specimenSerial"));
    if (!specimenSerial) {
      showToast("標本番号は必須です");
      return;
    }
    if (!/^\d+$/.test(specimenSerial)) {
      showToast("標本番号の数字部分は半角数字で入力してください");
      return;
    }
    const analysisType = specimenPrefix === "a" ? normalizeAnalysisType(value(formData.get("analysisType"))) : "";
    if (specimenPrefix === "a" && !analysisType) {
      showToast("a: 分析用資料を選んだ場合は、区分を選択してください");
      return;
    }

    const nowIsoValue = new Date().toISOString();
    const editingId = editingRecordId || recordIdInput.value;
    const found = findRecord(editingId);
    if (isEditTab && !found) {
      showToast("編集対象が見つかりません。リストから編集を選び直してください");
      return;
    }
    const specimenNo = buildSpecimenNo(specimenPrefix, specimenSerial);
    const saveAnswer = window.confirm(
      isEditTab ? `${specimenNo}の情報を上書き保存しますか？` : `${specimenNo}の情報を保存しますか？`
    );
    if (!saveAnswer) {
      return;
    }
    const hasSpecimenChanged = Boolean(found && found.specimenNo !== specimenNo);
    const recordId = isEditTab ? found?.id || editingId || newId("record") : hasSpecimenChanged ? newId("record") : found?.id || editingId || newId("record");
    const recordSiteSnapshot = isEditTab
      ? {
          levelHeight: value(editSiteSnapshot?.levelHeight),
          date: value(editSiteSnapshot?.date),
          team: value(editSiteSnapshot?.team),
          teamOther: value(editSiteSnapshot?.teamOther),
          teamLead: value(editSiteSnapshot?.teamLead),
          recorder: value(editSiteSnapshot?.recorder),
        }
      : {
          levelHeight: value(siteSnapshot?.levelHeight),
          date: value(siteSnapshot?.date),
          team: value(siteSnapshot?.team),
          teamOther: value(siteSnapshot?.teamOther),
          teamLead: value(siteSnapshot?.teamLead),
          recorder: value(siteSnapshot?.recorder),
        };
    const recordTeamState = normalizeTeamState(recordSiteSnapshot.team, recordSiteSnapshot.teamOther);

    const recordBase = {
      id: recordId,
      kuwaku: recordKuwaku,
      specimenPrefix,
      specimenSerial,
      specimenNo,
      category: categoryFromPrefix(specimenPrefix),
      analysisType,
      levelHeight: recordSiteSnapshot.levelHeight,
      date: recordSiteSnapshot.date,
      team: recordTeamState.team,
      teamOther: recordTeamState.teamOther,
      teamLead: recordSiteSnapshot.teamLead,
      recorder: recordSiteSnapshot.recorder,
      nameMemo: value(formData.get("nameMemo")),
      unit: compactNoSpaceValue(formData.get("unit")),
      discoverer: value(formData.get("discoverer")),
      identifier: value(formData.get("identifier")),
      levelUpperCm: value(formData.get("levelUpperCm")),
      levelLowerCm: value(formData.get("levelLowerCm")),
      occurrenceSection: normalizeNeedFlag(value(formData.get("occurrenceSection"))),
      occurrenceSketch: normalizeNeedFlag(value(formData.get("occurrenceSketch"))),
      nsDir: normalizeNsDir(value(formData.get("nsDir"))),
      nsCm: value(formData.get("nsCm")),
      ewDir: normalizeEwDir(value(formData.get("ewDir"))),
      ewCm: value(formData.get("ewCm")),
      importantFlag: normalizeHasFlag(value(formData.get("importantFlag"))),
      simpleRecordFlag: normalizeCircleDashFlag(value(formData.get("simpleRecordFlag"))),
      layerName: getSelectedLayerName(),
      detail: compactNoSpaceValue(formData.get("detail")),
      detailSub: value(formData.get("detailSub")),
      layerRef: value(formData.get("layerRef")),
      layerFromCm: value(formData.get("layerFromCm")),
      layerRelative: value(formData.get("layerRelative")),
      notes: value(formData.get("notes")),
      sectionDiagrams: clonePhotos(currentSectionDiagrams),
      photos: clonePhotos(currentPhotos),
      createdAt: found?.createdAt || nowIsoValue,
      updatedAt: nowIsoValue,
      deletedAt: "",
    };

    const targetIndex = state.records.findIndex((item) => item.id === recordBase.id);
    const previousRecord = targetIndex >= 0 ? state.records[targetIndex] : null;
    const historyAction = isEditTab ? "上書き保存" : targetIndex >= 0 ? "更新保存" : "新規保存";
    const record = {
      ...recordBase,
      history: buildNextRecordHistory(previousRecord, recordBase, historyAction),
    };

    if (targetIndex >= 0) {
      state.records[targetIndex] = record;
    } else {
      state.records.unshift(record);
    }

    if (isEditTab) {
      persist("上書き保存しました");
      renderRecordTable();
      renderOutputs();
      markOverwriteUpdatedState(found, record, value(found?.kuwaku), recordKuwaku);
      overwriteOriginalRecord = { ...record };
      renderEditHistory(record);
      return;
    }

    const carryForward = {
      layerName: record.layerName,
      unit: record.unit,
      detail: record.detail,
      detailSub: record.detailSub,
      layerRef: record.layerRef,
      layerFromCm: record.layerFromCm,
      layerRelative: record.layerRelative,
    };

    persist("記録を保存しました");
    renderRecordTable();
    renderOutputs();
    resetRecordForm({ showMessage: false });
    applyCarryForwardFields(carryForward);
    markCarryForwardSavedFields(carryForward);
  });

  recordResetBtn.addEventListener("click", () => {
    resetRecordForm({ showMessage: true });
  });

  recordTableBody.addEventListener("click", (event) => {
    const button = event.target.closest("button[data-action]");
    if (!button) {
      return;
    }
    const recordId = button.dataset.id;
    const action = button.dataset.action;
    const record = findRecord(recordId);
    if (!record) {
      showToast("対象データが見つかりません");
      return;
    }

    if (action === "edit") {
      const rowKuwaku = value(button.dataset.kuwaku);
      openRecordForEdit(record.id, rowKuwaku);
      return;
    }

    if (action === "delete") {
      const answer = window.confirm(`標本番号 ${record.specimenNo} を削除しますか？`);
      if (!answer) {
        return;
      }
      state.records = state.records.filter((item) => item.id !== recordId);
      if (editingRecordId === recordId) {
        resetRecordForm({ showMessage: false });
      }
      persist("記録を削除しました");
      renderRecordTable();
      renderOutputs();
    }
  });

  outputListBody.addEventListener("click", (event) => {
    const button = event.target.closest("button[data-action]");
    if (!button) {
      return;
    }
    const action = button.dataset.action;
    const recordId = button.dataset.id;
    const record = findRecord(recordId);
    if (!record) {
      showToast("対象データが見つかりません");
      return;
    }

    if (action === "edit") {
      const rowKuwaku = value(button.dataset.kuwaku);
      openRecordForEdit(recordId, rowKuwaku);
      return;
    }
    if (action === "card") {
      selectedCardRecordId = selectedCardRecordId === recordId ? "" : recordId;
      renderListOutput();
      renderCardOutput();
      return;
    }
    if (action === "delete") {
      const answer = window.confirm(`標本番号 ${record.specimenNo} を削除しますか？`);
      if (!answer) {
        return;
      }
      state.records = state.records.filter((item) => item.id !== recordId);
      if (editingRecordId === recordId) {
        resetRecordForm({ showMessage: false });
      }
      if (selectedCardRecordId === recordId) {
        selectedCardRecordId = "";
      }
      persist("記録を削除しました");
      renderRecordTable();
      renderOutputs();
      return;
    }
  });

  if (outputListTable) {
    const handleSortHeader = (target) => {
      if (!(target instanceof Element)) {
        return;
      }
      const header = target.closest("th[data-sort-key]");
      if (!header || !outputListTable.contains(header)) {
        return;
      }
      const sortKey = value(header.dataset.sortKey);
      if (!sortKey) {
        return;
      }
      if (outputListSortKey === sortKey) {
        outputListSortDirection = outputListSortDirection === "asc" ? "desc" : "asc";
      } else {
        outputListSortKey = sortKey;
        outputListSortDirection = "asc";
      }
      renderListOutput();
    };

    outputListTable.addEventListener("click", (event) => {
      handleSortHeader(event.target);
    });
    outputListTable.addEventListener("keydown", (event) => {
      if (event.key !== "Enter" && event.key !== " ") {
        return;
      }
      event.preventDefault();
      handleSortHeader(event.target);
    });
  }

  if (outputKuwakuSelect) {
    outputKuwakuSelect.addEventListener("change", () => {
      selectedOutputKuwaku = value(outputKuwakuSelect.value) || ALL_GRIDS_VALUE;
      selectedCardRecordId = "";
      renderOutputs();
    });
  }

  if (planKuwakuSelect) {
    planKuwakuSelect.addEventListener("change", () => {
      selectedPlanKuwaku = value(planKuwakuSelect.value);
      renderPlanOutput();
    });
  }

  if (planUnitSelect) {
    planUnitSelect.addEventListener("change", () => {
      selectedPlanUnit = value(planUnitSelect.value);
      renderPlanOutput();
    });
  }
  if (planDetailSelect) {
    planDetailSelect.addEventListener("change", () => {
      selectedPlanDetail = value(planDetailSelect.value) || ALL_DETAILS_VALUE;
      renderPlanOutput();
    });
  }

  if (editKuwakuHeadAInput) {
    editKuwakuHeadAInput.addEventListener("input", () => {
      editKuwakuHeadAInput.classList.remove("overwrite-updated");
      updateDuplicateSpecimenWarning();
    });
  }
  if (editKuwakuHeadBInput) {
    editKuwakuHeadBInput.addEventListener("input", () => {
      editKuwakuHeadBInput.classList.remove("overwrite-updated");
      updateDuplicateSpecimenWarning();
    });
  }
  if (editKuwakuBlockInput) {
    editKuwakuBlockInput.addEventListener("input", () => {
      editKuwakuBlockInput.classList.remove("overwrite-updated");
      updateDuplicateSpecimenWarning();
    });
  }
  if (editKuwakuNoInput) {
    editKuwakuNoInput.addEventListener("input", () => {
      editKuwakuNoInput.classList.remove("overwrite-updated");
      updateDuplicateSpecimenWarning();
    });
  }
  if (editLevelHeightInput) {
    editLevelHeightInput.addEventListener("input", () => {
      editLevelHeightInput.classList.remove("overwrite-updated");
    });
  }
  if (editDateInput) {
    editDateInput.addEventListener("input", () => {
      editDateInput.classList.remove("overwrite-updated");
    });
  }
  if (editTeamOtherInput) {
    editTeamOtherInput.addEventListener("input", () => {
      editTeamOtherInput.classList.remove("overwrite-updated");
    });
  }
  if (editTeamLeadInput) {
    editTeamLeadInput.addEventListener("input", () => {
      editTeamLeadInput.classList.remove("overwrite-updated");
    });
  }
  if (editRecorderInput) {
    editRecorderInput.addEventListener("input", () => {
      editRecorderInput.classList.remove("overwrite-updated");
    });
  }

  const sectionDiagramInputs = [sectionDiagramCameraInput, sectionDiagramInput].filter(Boolean);
  sectionDiagramInputs.forEach((input) => {
    input.addEventListener("change", async (event) => {
      await addSectionDiagramsFromFiles(event.target.files);
      event.target.value = "";
    });
  });

  sectionDiagramList.addEventListener("input", (event) => {
    const input = event.target.closest("input[data-diagram-id]");
    if (!input) {
      return;
    }
    const target = currentSectionDiagrams.find((item) => item.id === input.dataset.diagramId);
    if (!target) {
      return;
    }
    target.caption = input.value;
    persist();
  });

  sectionDiagramList.addEventListener("click", (event) => {
    const button = event.target.closest("button[data-remove-diagram-id]");
    if (!button) {
      return;
    }
    currentSectionDiagrams = currentSectionDiagrams.filter((item) => item.id !== button.dataset.removeDiagramId);
    renderSectionDiagramList();
    persist("断面図を削除しました");
  });

  const photoInputs = [photoCameraInput, photoInput].filter(Boolean);
  photoInputs.forEach((input) => {
    input.addEventListener("change", async (event) => {
      await addPhotosFromFiles(event.target.files);
      event.target.value = "";
    });
  });

  photoList.addEventListener("input", (event) => {
    const input = event.target.closest("input[data-photo-id]");
    if (!input) {
      return;
    }
    const target = currentPhotos.find((photo) => photo.id === input.dataset.photoId);
    if (!target) {
      return;
    }
    target.caption = input.value;
    persist();
  });

  photoList.addEventListener("click", (event) => {
    const button = event.target.closest("button[data-remove-photo-id]");
    if (!button) {
      return;
    }
    currentPhotos = currentPhotos.filter((photo) => photo.id !== button.dataset.removePhotoId);
    renderPhotoList();
    persist("写真を削除しました");
  });

  exportListCsvBtn.addEventListener("click", () => {
    const csv = buildListCsv();
    downloadFile(`nojiri-kaseki-list-${timestamp()}.csv`, csv, "text/csv;charset=utf-8");
    showToast("リストCSVを書き出しました");
  });

  exportCardCsvBtn.addEventListener("click", () => {
    const csv = buildCardCsv();
    downloadFile(`nojiri-kaseki-card-${timestamp()}.csv`, csv, "text/csv;charset=utf-8");
    showToast("カードCSVを書き出しました");
  });

  exportJsonBtn.addEventListener("click", () => {
    const json = JSON.stringify(state, null, 2);
    downloadFile(`nojiri-kaseki-${timestamp()}.json`, json, "application/json");
    showToast("JSONを書き出しました");
  });

  importJsonInput.addEventListener("change", async (event) => {
    const [file] = Array.from(event.target.files || []);
    if (!file) {
      return;
    }
    try {
      const text = await file.text();
      const imported = JSON.parse(text);
      state = normalizeState(imported);
      hydrateSiteForm();
      resetRecordForm({ showMessage: false });
      renderRecordTable();
      renderOutputs();
      persist("JSONを読み込みました");
    } catch (_error) {
      showToast("JSON読み込みに失敗しました");
    } finally {
      importJsonInput.value = "";
    }
  });
}

async function addSectionDiagramsFromFiles(fileList) {
  const files = Array.from(fileList || []);
  if (!files.length) {
    return;
  }

  let added = false;
  for (const file of files) {
    try {
      const dataUrl = await resizeImage(file, 1280, 0.72);
      currentSectionDiagrams.push({
        id: newId("diagram"),
        name: file.name || "diagram.jpg",
        dataUrl,
        caption: "",
        createdAt: new Date().toISOString(),
      });
      added = true;
    } catch (_error) {
      showToast(`断面図追加に失敗: ${file.name}`);
    }
  }

  if (!added) {
    return;
  }
  renderSectionDiagramList();
  persist();
}

async function addPhotosFromFiles(fileList) {
  const files = Array.from(fileList || []);
  if (!files.length) {
    return;
  }

  let added = false;
  for (const file of files) {
    try {
      const dataUrl = await resizeImage(file, 1280, 0.72);
      currentPhotos.push({
        id: newId("photo"),
        name: file.name || "photo.jpg",
        dataUrl,
        caption: "",
        createdAt: new Date().toISOString(),
      });
      added = true;
    } catch (_error) {
      showToast(`写真追加に失敗: ${file.name}`);
    }
  }

  if (!added) {
    return;
  }
  renderPhotoList();
  persist();
}

function setActiveTab(tabId) {
  tabButtons.forEach((button) => {
    button.classList.toggle("active", button.dataset.tab === tabId);
  });
  tabPanels.forEach((panel) => {
    panel.classList.toggle("active", panel.id === tabId);
  });
  syncRecordFormPlacement(tabId);
  syncEditHistoryVisibility(tabId);
  updateDuplicateSpecimenWarning();
  if (CLOUD_AUTO_PULL_ENABLED && cloudEndpoint && (tabId === "output-tab" || tabId === "plan-tab")) {
    void pullStateFromCloud({ force: false, showToastOnSuccess: false, silentOnError: true });
  }
}

function syncRecordFormPlacement(tabId) {
  if (!recordForm || !recordFormHost || !editRecordFormHost) {
    return;
  }

  if (tabId === "edit-tab") {
    if (recordForm.parentElement !== editRecordFormHost) {
      editRecordFormHost.appendChild(recordForm);
    }
    recordSubmitBtn.textContent = "上書き保存";
    return;
  }

  if (recordForm.parentElement !== recordFormHost) {
    recordFormHost.appendChild(recordForm);
  }

  if (tabId === "input-tab") {
    if (isOverwriteMode) {
      isOverwriteMode = false;
      overwriteOriginalRecord = null;
      clearOverwriteUpdatedState();
      resetRecordForm({ showMessage: false });
    }
    recordSubmitBtn.textContent = "記録を保存";
  }
}

function getActiveTabId() {
  const activePanel = document.querySelector(".tab-panel.active");
  return activePanel?.id || "";
}

function hydrateSiteForm() {
  siteForm.elements.kuwakuHeadA.value = state.site.kuwakuHeadA || DEFAULT_KUWAKU_HEAD_A;
  siteForm.elements.kuwakuHeadB.value = state.site.kuwakuHeadB || DEFAULT_KUWAKU_HEAD_B;
  siteForm.elements.kuwakuBlock.value = state.site.kuwakuBlock || "";
  siteForm.elements.kuwakuNo.value = state.site.kuwakuNo || "";
  siteForm.elements.levelHeight.value = state.site.levelHeight || "";
  siteForm.elements.date.value = state.site.date || "";
  const teamState = normalizeTeamState(state.site.team, state.site.teamOther);
  siteForm.elements.team.value = teamState.team || "";
  siteForm.elements.teamOther.value = teamState.teamOther || "";
  syncTeamOtherInput(siteForm.elements.team.value);
  siteForm.elements.teamLead.value = state.site.teamLead || "";
  siteForm.elements.recorder.value = state.site.recorder || "";
}

function activateSpecimenPrefix(prefixRaw) {
  const prefix = normalizeSpecimenPrefix(prefixRaw);
  specimenPrefixInput.value = prefix;
  specimenPrefixLabel.textContent = prefix;
  specimenTabButtons.forEach((button) => {
    button.classList.toggle("active", normalizeSpecimenPrefix(button.dataset.prefix) === prefix);
  });
  syncAnalysisTypeInput(prefix);
}

function updateSpecimenNoFromParts() {
  const prefix = normalizeSpecimenPrefix(specimenPrefixInput.value);
  const serial = value(specimenSerialInput.value);
  specimenPrefixInput.value = prefix;
  specimenNoInput.value = buildSpecimenNo(prefix, serial);
  updateDuplicateSpecimenWarning();
}

function updateDuplicateSpecimenWarning() {
  if (!specimenDuplicateWarning) {
    return;
  }
  if (recordSubmitBtn) {
    recordSubmitBtn.disabled = false;
  }
  const activeTabId = getActiveTabId();
  if (activeTabId !== "input-tab" && activeTabId !== "edit-tab") {
    hideDuplicateSpecimenWarning();
    return;
  }

  const specimenSerial = compactNoSpaceValue(specimenSerialInput?.value);
  const specimenPrefix = normalizeSpecimenPrefix(specimenPrefixInput?.value);
  const specimenNo = buildSpecimenNo(specimenPrefix, specimenSerial);
  if (!specimenNo || !specimenSerial) {
    hideDuplicateSpecimenWarning();
    return;
  }

  const kuwaku = currentKuwakuForDuplicateWarning(activeTabId);
  if (!kuwaku) {
    hideDuplicateSpecimenWarning();
    return;
  }
  const excludeRecordId = activeTabId === "edit-tab" ? value(editingRecordId || recordIdInput?.value) : "";
  const duplicate = findDuplicateRecordByKuwakuAndSpecimen(kuwaku, specimenNo, excludeRecordId);
  if (!duplicate) {
    hideDuplicateSpecimenWarning();
    return;
  }
  specimenDuplicateWarning.textContent = `警告: この区画には ${specimenNo} がすでにあります`;
  specimenDuplicateWarning.classList.remove("hidden");
  if (recordSubmitBtn) {
    recordSubmitBtn.disabled = true;
  }
}

function currentKuwakuForDuplicateWarning(activeTabId = getActiveTabId()) {
  if (activeTabId === "edit-tab") {
    const headA = normalizeKuwakuHeadA(editKuwakuHeadAInput?.value);
    const headB = normalizeKuwakuHeadB(editKuwakuHeadBInput?.value);
    const block = normalizeKuwakuBlock(editKuwakuBlockInput?.value);
    const no = normalizeKuwakuNo(editKuwakuNoInput?.value);
    if (!block || !no) {
      return "";
    }
    return buildKuwaku(headA, headB, block, no);
  }

  const headA = normalizeKuwakuHeadA(siteForm?.elements?.kuwakuHeadA?.value);
  const headB = normalizeKuwakuHeadB(siteForm?.elements?.kuwakuHeadB?.value);
  const block = normalizeKuwakuBlock(siteForm?.elements?.kuwakuBlock?.value);
  const no = normalizeKuwakuNo(siteForm?.elements?.kuwakuNo?.value);
  if (!block || !no) {
    return "";
  }
  return buildKuwaku(headA, headB, block, no);
}

function hideDuplicateSpecimenWarning() {
  if (!specimenDuplicateWarning) {
    return;
  }
  specimenDuplicateWarning.textContent = "";
  specimenDuplicateWarning.classList.add("hidden");
  if (recordSubmitBtn) {
    recordSubmitBtn.disabled = false;
  }
}

function syncAnalysisTypeInput(prefixRaw) {
  if (!analysisTypeRow || !analysisTypeSelect) {
    return;
  }
  const prefix = normalizeSpecimenPrefix(prefixRaw);
  const isAnalysis = prefix === "a";
  analysisTypeRow.classList.toggle("hidden", !isAnalysis);
  if (!isAnalysis) {
    analysisTypeSelect.value = "";
  } else {
    analysisTypeSelect.value = normalizeAnalysisType(analysisTypeSelect.value);
  }
}

function activateDirectionTab(group, valueRaw) {
  if (group === "ns") {
    nsDirInput.value = normalizeNsDir(valueRaw);
  }
  if (group === "ew") {
    ewDirInput.value = normalizeEwDir(valueRaw);
  }
  syncDirectionTabsFromForm();
}

function syncDirectionTabsFromForm() {
  const nsValue = normalizeNsDir(nsDirInput.value);
  const ewValue = normalizeEwDir(ewDirInput.value);
  nsDirInput.value = nsValue;
  ewDirInput.value = ewValue;

  dirTabButtons.forEach((button) => {
    const group = button.dataset.group;
    const selected = group === "ns" ? nsValue : ewValue;
    button.classList.toggle("active", normalizeDirectionValue(group, button.dataset.value) === selected);
  });
}

function activateLayerTab(layerRaw) {
  const layer = PRESET_LAYER_NAMES.includes(value(layerRaw)) ? value(layerRaw) : PRESET_LAYER_NAMES[0];
  layerNameInput.value = layer;
  layerTabButtons.forEach((button) => {
    button.classList.toggle("active", button.dataset.layer === layer);
  });

  const isOther = layer === OTHER_LAYER_NAME;
  layerOtherInput.classList.toggle("hidden", !isOther);
  if (!isOther) {
    layerOtherInput.value = "";
  }
}

function setLayerFromValue(layerRaw) {
  const layerValue = normalizeLayerName(value(layerRaw));
  if (!layerValue) {
    activateLayerTab(PRESET_LAYER_NAMES[0]);
    return;
  }
  if (PRESET_LAYER_NAMES.includes(layerValue) && layerValue !== OTHER_LAYER_NAME) {
    activateLayerTab(layerValue);
    return;
  }

  activateLayerTab(OTHER_LAYER_NAME);
  const otherText = extractOtherLayerText(layerValue);
  layerOtherInput.value = otherText;
}

function getSelectedLayerName() {
  const selected = value(layerNameInput.value) || PRESET_LAYER_NAMES[0];
  if (selected !== OTHER_LAYER_NAME) {
    return selected;
  }

  const otherText = value(layerOtherInput.value);
  return otherText ? `${OTHER_LAYER_NAME}:${otherText}` : OTHER_LAYER_NAME;
}

function applyCarryForwardFields(saved) {
  setLayerFromValue(value(saved?.layerName));
  recordForm.elements.unit.value = value(saved?.unit);
  recordForm.elements.detail.value = value(saved?.detail);
  recordForm.elements.detailSub.value = value(saved?.detailSub);
  recordForm.elements.layerRef.value = value(saved?.layerRef);
  recordForm.elements.layerFromCm.value = value(saved?.layerFromCm);
  recordForm.elements.layerRelative.value = value(saved?.layerRelative);
}

function markCarryForwardSavedFields(saved) {
  clearCarryForwardSavedFields();

  if (value(saved?.unit)) {
    recordForm.elements.unit.classList.add("saved-carry-value");
  }
  if (value(saved?.detail)) {
    recordForm.elements.detail.classList.add("saved-carry-value");
  }
  if (value(saved?.detailSub)) {
    recordForm.elements.detailSub.classList.add("saved-carry-value");
  }
  if (value(saved?.layerRef)) {
    recordForm.elements.layerRef.classList.add("saved-carry-value");
  }
  if (value(saved?.layerFromCm)) {
    recordForm.elements.layerFromCm.classList.add("saved-carry-value");
  }
  if (value(saved?.layerRelative)) {
    recordForm.elements.layerRelative.classList.add("saved-carry-value");
  }

  markLayerSavedTabState();
  if (layerNameInput.value === OTHER_LAYER_NAME && value(layerOtherInput.value)) {
    layerOtherInput.classList.add("saved-carry-value");
  }
}

function clearCarryForwardSavedFields() {
  recordForm.elements.unit.classList.remove("saved-carry-value");
  recordForm.elements.detail.classList.remove("saved-carry-value");
  recordForm.elements.detailSub.classList.remove("saved-carry-value");
  recordForm.elements.layerRef.classList.remove("saved-carry-value");
  recordForm.elements.layerFromCm.classList.remove("saved-carry-value");
  recordForm.elements.layerRelative.classList.remove("saved-carry-value");
  layerOtherInput.classList.remove("saved-carry-value");
  clearLayerSavedTabState();
}

function clearOverwriteUpdatedState() {
  if (!recordForm) {
    return;
  }
  recordForm.querySelectorAll(".overwrite-updated").forEach((element) => {
    element.classList.remove("overwrite-updated");
  });
  if (editKuwakuBlockInput) {
    editKuwakuBlockInput.classList.remove("overwrite-updated");
  }
  if (editKuwakuNoInput) {
    editKuwakuNoInput.classList.remove("overwrite-updated");
  }
  if (editKuwakuHeadAInput) {
    editKuwakuHeadAInput.classList.remove("overwrite-updated");
  }
  if (editKuwakuHeadBInput) {
    editKuwakuHeadBInput.classList.remove("overwrite-updated");
  }
  if (editLevelHeightInput) {
    editLevelHeightInput.classList.remove("overwrite-updated");
  }
  if (editDateInput) {
    editDateInput.classList.remove("overwrite-updated");
  }
  if (editTeamInput) {
    editTeamInput.classList.remove("overwrite-updated");
  }
  if (editTeamOtherInput) {
    editTeamOtherInput.classList.remove("overwrite-updated");
  }
  if (editTeamLeadInput) {
    editTeamLeadInput.classList.remove("overwrite-updated");
  }
  if (editRecorderInput) {
    editRecorderInput.classList.remove("overwrite-updated");
  }
}

function markOverwriteUpdatedState(previousRecord, nextRecord, previousKuwakuRaw, nextKuwakuRaw) {
  clearOverwriteUpdatedState();
  if (!previousRecord || !nextRecord) {
    return;
  }

  const fields = [
    "specimenSerial",
    "nameMemo",
    "unit",
    "discoverer",
    "identifier",
    "levelUpperCm",
    "levelLowerCm",
    "occurrenceSection",
    "occurrenceSketch",
    "importantFlag",
    "simpleRecordFlag",
    "analysisType",
    "nsCm",
    "ewCm",
    "detail",
    "detailSub",
    "layerRef",
    "layerFromCm",
    "layerRelative",
    "notes",
  ];

  fields.forEach((name) => {
    const element = recordForm.elements[name];
    if (!(element instanceof Element)) {
      return;
    }
    const prev = value(previousRecord?.[name]);
    const next = value(nextRecord?.[name]);
    if (prev !== next) {
      element.classList.add("overwrite-updated");
    }
  });

  const prevPrefix = normalizeSpecimenPrefix(previousRecord.specimenPrefix);
  const nextPrefix = normalizeSpecimenPrefix(nextRecord.specimenPrefix);
  if (prevPrefix !== nextPrefix) {
    specimenPrefixLabel.classList.add("overwrite-updated");
    specimenTabButtons.forEach((button) => {
      if (normalizeSpecimenPrefix(button.dataset.prefix) === nextPrefix) {
        button.classList.add("overwrite-updated");
      }
    });
  }

  if (normalizeNsDir(previousRecord.nsDir) !== normalizeNsDir(nextRecord.nsDir)) {
    dirTabButtons.forEach((button) => {
      if (button.dataset.group === "ns") {
        button.classList.add("overwrite-updated");
      }
    });
  }
  if (normalizeEwDir(previousRecord.ewDir) !== normalizeEwDir(nextRecord.ewDir)) {
    dirTabButtons.forEach((button) => {
      if (button.dataset.group === "ew") {
        button.classList.add("overwrite-updated");
      }
    });
  }
  if (value(previousRecord.layerName) !== value(nextRecord.layerName)) {
    layerTabButtons.forEach((button) => {
      if (button.classList.contains("active")) {
        button.classList.add("overwrite-updated");
      }
    });
    layerOtherInput.classList.add("overwrite-updated");
  }

  const previousParts = parseKuwaku(previousKuwakuRaw);
  const nextParts = parseKuwaku(nextKuwakuRaw);
  if (editKuwakuHeadAInput && previousParts.headA !== nextParts.headA) {
    editKuwakuHeadAInput.classList.add("overwrite-updated");
  }
  if (editKuwakuHeadBInput && previousParts.headB !== nextParts.headB) {
    editKuwakuHeadBInput.classList.add("overwrite-updated");
  }
  if (editKuwakuBlockInput && previousParts.block !== nextParts.block) {
    editKuwakuBlockInput.classList.add("overwrite-updated");
  }
  if (editKuwakuNoInput && previousParts.no !== nextParts.no) {
    editKuwakuNoInput.classList.add("overwrite-updated");
  }
  if (editLevelHeightInput && value(previousRecord.levelHeight) !== value(nextRecord.levelHeight)) {
    editLevelHeightInput.classList.add("overwrite-updated");
  }
  if (editDateInput && value(previousRecord.date) !== value(nextRecord.date)) {
    editDateInput.classList.add("overwrite-updated");
  }
  if (
    editTeamInput &&
    (value(previousRecord.team) !== value(nextRecord.team) || value(previousRecord.teamOther) !== value(nextRecord.teamOther))
  ) {
    editTeamInput.classList.add("overwrite-updated");
  }
  if (editTeamOtherInput && value(previousRecord.teamOther) !== value(nextRecord.teamOther)) {
    editTeamOtherInput.classList.add("overwrite-updated");
  }
  if (editTeamLeadInput && value(previousRecord.teamLead) !== value(nextRecord.teamLead)) {
    editTeamLeadInput.classList.add("overwrite-updated");
  }
  if (editRecorderInput && value(previousRecord.recorder) !== value(nextRecord.recorder)) {
    editRecorderInput.classList.add("overwrite-updated");
  }
}

function markLayerSavedTabState() {
  const activeButton = Array.from(layerTabButtons).find((button) => button.classList.contains("active"));
  if (activeButton) {
    activeButton.classList.add("saved-carry-value");
  }
}

function clearLayerSavedTabState() {
  layerTabButtons.forEach((button) => {
    button.classList.remove("saved-carry-value");
  });
}

function handleRecordFormFieldEdit(event) {
  const target = event.target;
  if (!(target instanceof Element)) {
    return;
  }

  target.classList.remove("overwrite-updated");

  const isCarryField =
    target instanceof HTMLInputElement &&
    (target.name === "unit" ||
      target.name === "detail" ||
      target.name === "detailSub" ||
      target.name === "layerRef" ||
      target.name === "layerFromCm" ||
      target.name === "layerRelative");
  if (isCarryField) {
    target.classList.remove("saved-carry-value");
    return;
  }

  if (target === layerOtherInput) {
    layerOtherInput.classList.remove("saved-carry-value");
    clearLayerSavedTabState();
  }
}

function resetRecordForm({ showMessage }) {
  recordForm.reset();
  isOverwriteMode = false;
  overwriteOriginalRecord = null;
  editingRecordId = null;
  recordIdInput.value = "";
  recordSubmitBtn.textContent = "記録を保存";
  clearCarryForwardSavedFields();
  clearOverwriteUpdatedState();

  activateSpecimenPrefix(DEFAULT_SPECIMEN_PREFIX);
  specimenSerialInput.value = "";
  if (analysisTypeSelect) {
    analysisTypeSelect.value = "";
  }
  updateSpecimenNoFromParts();

  recordForm.elements.occurrenceSection.value = "要";
  recordForm.elements.occurrenceSketch.value = "要";
  recordForm.elements.simpleRecordFlag.value = "-";
  setLayerFromValue(PRESET_LAYER_NAMES[0]);

  nsDirInput.value = "北から";
  ewDirInput.value = "東から";
  syncDirectionTabsFromForm();

  currentPhotos = [];
  currentSectionDiagrams = [];
  renderSectionDiagramList();
  renderPhotoList();
  clearEditHistory();

  if (showMessage) {
    showToast("入力をクリアしました");
  }
}

function populateRecordForm(record) {
  const parsedSpecimen = parseSpecimenNo(record.specimenNo, record.specimenPrefix, record.specimenSerial);

  editingRecordId = record.id;
  recordIdInput.value = record.id;
  recordSubmitBtn.textContent = "記録を保存";

  activateSpecimenPrefix(parsedSpecimen.prefix);
  if (analysisTypeSelect) {
    analysisTypeSelect.value = normalizeAnalysisType(record.analysisType);
  }
  specimenSerialInput.value = parsedSpecimen.serial;
  updateSpecimenNoFromParts();

  recordForm.elements.nameMemo.value = record.nameMemo || "";
  recordForm.elements.unit.value = record.unit || "";
  recordForm.elements.discoverer.value = record.discoverer || "";
  recordForm.elements.identifier.value = record.identifier || "";
  recordForm.elements.levelUpperCm.value = record.levelUpperCm || "";
  recordForm.elements.levelLowerCm.value = record.levelLowerCm || "";
  recordForm.elements.occurrenceSection.value = normalizeNeedFlag(record.occurrenceSection);
  recordForm.elements.occurrenceSketch.value = normalizeNeedFlag(record.occurrenceSketch);

  nsDirInput.value = normalizeNsDir(record.nsDir);
  recordForm.elements.nsCm.value = record.nsCm || "";
  ewDirInput.value = normalizeEwDir(record.ewDir);
  recordForm.elements.ewCm.value = record.ewCm || "";
  syncDirectionTabsFromForm();

  recordForm.elements.importantFlag.value = normalizeHasFlag(record.importantFlag);
  recordForm.elements.simpleRecordFlag.value = normalizeCircleDashFlag(record.simpleRecordFlag);
  setLayerFromValue(record.layerName);
  recordForm.elements.detail.value = record.detail || "";
  recordForm.elements.detailSub.value = record.detailSub || "";
  recordForm.elements.layerRef.value = record.layerRef || "";
  recordForm.elements.layerFromCm.value = record.layerFromCm || "";
  recordForm.elements.layerRelative.value = record.layerRelative || "";
  recordForm.elements.notes.value = record.notes || "";
  clearCarryForwardSavedFields();
  clearOverwriteUpdatedState();
  currentSectionDiagrams = clonePhotos(record.sectionDiagrams || []);
  renderSectionDiagramList();
  currentPhotos = clonePhotos(record.photos || []);
  renderPhotoList();
}

function openRecordForEdit(recordId, preferredKuwaku = "") {
  const record = findRecord(recordId);
  if (!record) {
    showToast("対象データが見つかりません");
    return;
  }
  const kuwakuSource = value(preferredKuwaku) || value(record.kuwaku) || getRecordKuwaku(record);
  const kuwakuParts = parseKuwaku(kuwakuSource);
  if (editKuwakuHeadAInput) {
    editKuwakuHeadAInput.value = kuwakuParts.headA;
  }
  if (editKuwakuHeadBInput) {
    editKuwakuHeadBInput.value = kuwakuParts.headB;
  }
  if (editKuwakuBlockInput) {
    editKuwakuBlockInput.value = kuwakuParts.block;
  }
  if (editKuwakuNoInput) {
    editKuwakuNoInput.value = kuwakuParts.no;
  }
  if (editLevelHeightInput) {
    editLevelHeightInput.value = value(record.levelHeight) || value(state.site?.levelHeight);
  }
  if (editDateInput) {
    editDateInput.value = value(record.date) || value(state.site?.date);
  }
  const editTeamState = normalizeTeamState(
    value(record.team) || value(state.site?.team),
    value(record.teamOther) || value(state.site?.teamOther)
  );
  if (editTeamInput) {
    editTeamInput.value = editTeamState.team;
  }
  if (editTeamOtherInput) {
    editTeamOtherInput.value = editTeamState.teamOther;
  }
  syncEditTeamOtherInput(editTeamState.team);
  if (editTeamLeadInput) {
    editTeamLeadInput.value = value(record.teamLead) || value(state.site?.teamLead);
  }
  if (editRecorderInput) {
    editRecorderInput.value = value(record.recorder) || value(state.site?.recorder);
  }
  isOverwriteMode = true;
  overwriteOriginalRecord = { ...record };
  clearOverwriteUpdatedState();
  populateRecordForm(record);
  renderEditHistory(record);
  setActiveTab("edit-tab");
}

function syncEditHistoryVisibility(activeTabId = getActiveTabId()) {
  if (!editHistoryPanel) {
    return;
  }
  const shouldShow = activeTabId === "edit-tab" && Boolean(editingRecordId);
  editHistoryPanel.classList.toggle("hidden", !shouldShow);
}

function clearEditHistory() {
  if (editHistoryList) {
    editHistoryList.innerHTML = "";
  }
  if (editHistoryPanel) {
    editHistoryPanel.classList.add("hidden");
  }
}

function renderEditHistory(record) {
  if (!editHistoryList || !editHistoryPanel) {
    return;
  }
  const history = normalizeRecordHistory(record?.history);
  if (!history.length) {
    editHistoryList.innerHTML = "<p class=\"muted\">履歴はまだありません。</p>";
  } else {
    const displayHistory = history
      .map((entry, index) => ({
        entry,
        prevEntry: index > 0 ? history[index - 1] : null,
      }))
      .reverse();
    editHistoryList.innerHTML = displayHistory
      .map(({ entry, prevEntry }) => {
        const contentHtml = renderHistoryContentHtml(entry, prevEntry);
        return `
          <article class="edit-history-item">
            <p><strong>入力内容:</strong> ${contentHtml}</p>
            <p><strong>年・月日・時間:</strong> ${escapeHtml(formatHistoryDateTime(entry.at))}</p>
          </article>
        `;
      })
      .join("");
  }
  if (getActiveTabId() === "edit-tab") {
    editHistoryPanel.classList.remove("hidden");
  }
}

function renderRecordTable() {
  if (!state.records.length) {
    recordTableBody.innerHTML = "<tr><td colspan=\"8\">まだ入力データがありません。</td></tr>";
    return;
  }

  const recordsForCurrentKuwaku = getInputRecordsForCurrentKuwaku();
  if (!recordsForCurrentKuwaku.length) {
    recordTableBody.innerHTML = "<tr><td colspan=\"8\">現在の区画（グリッド）の入力データがありません。</td></tr>";
    return;
  }

  recordTableBody.innerHTML = recordsForCurrentKuwaku
    .map((record) => {
      return `
      <tr>
        <td>${escapeHtml(getRecordKuwaku(record))}</td>
        <td>${escapeHtml(getRecordTeamValue(record))}</td>
        <td>${escapeHtml(record.specimenNo)}</td>
        <td>${escapeHtml(formatCategoryForRecord(record))}</td>
        <td>${escapeHtml(record.nameMemo || "")}</td>
        <td>${escapeHtml(record.discoverer || "")}</td>
        <td>${escapeHtml(formatLevelRead(record))}</td>
        <td>
          <div class="row-actions">
            <button type="button" data-action="edit" data-id="${record.id}" data-kuwaku="${escapeHtml(
              getRecordKuwaku(record)
            )}">編集</button>
            <button class="danger" type="button" data-action="delete" data-id="${record.id}">削除</button>
          </div>
        </td>
      </tr>
      `;
    })
    .join("");
}

function getInputRecordsForCurrentKuwaku() {
  const currentKuwaku = value(state.site?.kuwaku);
  const sortedRecords = [...state.records].sort(compareRecordsByKuwakuThenSpecimen);
  if (!currentKuwaku || isDefaultKuwaku(currentKuwaku)) {
    return sortedRecords;
  }
  const currentValue = kuwakuValueForSelect(currentKuwaku);
  return sortedRecords.filter((record) => kuwakuValueForSelect(getRecordKuwaku(record)) === currentValue);
}

function renderOutputs() {
  renderCardOutput();
  renderListOutput();
  renderPlanOutput();
}

function renderListOutput() {
  updateOutputListSortHeader();
  if (!state.records.length) {
    syncOutputKuwakuSelect([]);
    outputListBody.innerHTML = "<tr><td colspan=\"16\">出力対象データがありません。</td></tr>";
    return;
  }

  const filteredRecords = getFilteredOutputRecords();
  if (!filteredRecords.length) {
    outputListBody.innerHTML = "<tr><td colspan=\"16\">選択した区画のデータがありません。</td></tr>";
    return;
  }

  outputListBody.innerHTML = sortOutputRecordsForList(filteredRecords)
    .map((record) => {
      const selectedClass = record.id === selectedCardRecordId ? "selected-card-row" : "";
      const cardButtonLabel = record.id === selectedCardRecordId ? "プレビュー表示中" : "カード";
      return `
      <tr class="${selectedClass}">
        <td>${escapeHtml(getRecordKuwaku(record))}</td>
        <td>${escapeHtml(getRecordTeamValue(record))}</td>
        <td>${escapeHtml(record.specimenNo)}</td>
        <td>${escapeHtml(formatCategoryForRecord(record))}</td>
        <td>${escapeHtml(record.nameMemo || "")}</td>
        <td>${escapeHtml(record.importantFlag || "")}</td>
        <td>${escapeHtml(record.unit || "")}</td>
        <td>${escapeHtml(formatDetailForRecord(record))}</td>
        <td>${escapeHtml(record.discoverer || "")}</td>
        <td>${escapeHtml(record.identifier || "")}</td>
        <td>${escapeHtml(formatLevelRead(record))}</td>
        <td>${escapeHtml(record.occurrenceSection || "")}</td>
        <td>${escapeHtml(record.occurrenceSketch || "")}</td>
        <td>${escapeHtml(formatPlanPosition(record))}</td>
        <td>${escapeHtml(record.notes || "")}</td>
        <td>
          <div class="row-actions">
            <button type="button" data-action="card" data-id="${record.id}">${cardButtonLabel}</button>
            <button type="button" data-action="edit" data-id="${record.id}" data-kuwaku="${escapeHtml(
              getRecordKuwaku(record)
            )}">編集</button>
            <button class="danger" type="button" data-action="delete" data-id="${record.id}">削除</button>
          </div>
        </td>
      </tr>
      `;
    })
    .join("");
}

function updateOutputListSortHeader() {
  if (!outputListTable) {
    return;
  }
  const headers = outputListTable.querySelectorAll("th[data-sort-key]");
  headers.forEach((header) => {
    const sortKey = value(header.dataset.sortKey);
    const isActive = sortKey === outputListSortKey;
    header.classList.add("sortable-header");
    header.classList.toggle("sort-asc", isActive && outputListSortDirection === "asc");
    header.classList.toggle("sort-desc", isActive && outputListSortDirection === "desc");
    header.setAttribute("aria-sort", isActive ? (outputListSortDirection === "asc" ? "ascending" : "descending") : "none");
  });
}

function sortOutputRecordsForList(records) {
  const list = [...records];
  list.sort(compareOutputRecordsForList);
  return list;
}

function compareOutputRecordsForList(a, b) {
  let compared = 0;
  switch (outputListSortKey) {
    case "kuwaku":
      compared = compareRecordsByKuwakuThenSpecimen(a, b);
      break;
    case "team":
      compared = compareSortText(getRecordTeamValue(a), getRecordTeamValue(b));
      break;
    case "specimenNo":
      compared = compareRecordsBySpecimenNo(a, b);
      break;
    case "category":
      compared = compareSortText(formatCategoryForRecord(a), formatCategoryForRecord(b));
      break;
    case "nameMemo":
      compared = compareSortText(a?.nameMemo, b?.nameMemo);
      break;
    case "importantFlag":
      compared = compareSortText(a?.importantFlag, b?.importantFlag);
      break;
    case "discoverer":
      compared = compareSortText(a?.discoverer, b?.discoverer);
      break;
    case "identifier":
      compared = compareSortText(a?.identifier, b?.identifier);
      break;
    case "levelRead":
      compared = compareSortText(formatLevelRead(a), formatLevelRead(b));
      break;
    case "occurrenceSection":
      compared = compareSortText(a?.occurrenceSection, b?.occurrenceSection);
      break;
    case "occurrenceSketch":
      compared = compareSortText(a?.occurrenceSketch, b?.occurrenceSketch);
      break;
    case "position":
      compared = compareSortText(formatPlanPosition(a), formatPlanPosition(b));
      break;
    case "unit":
      compared = compareSortText(a?.unit, b?.unit);
      break;
    case "detail":
      compared = compareSortText(formatDetailForRecord(a), formatDetailForRecord(b));
      break;
    case "notes":
      compared = compareSortText(a?.notes, b?.notes);
      break;
    default:
      compared = compareRecordsByKuwakuThenSpecimen(a, b);
      break;
  }
  const fallback = compareRecordsByKuwakuThenSpecimen(a, b);
  const direction = outputListSortDirection === "desc" ? -1 : 1;
  return (compared || fallback) * direction;
}

function compareSortText(a, b) {
  return value(a).localeCompare(value(b), "ja", { numeric: true, sensitivity: "base" });
}

function formatPlanPosition(record) {
  return `${value(record?.nsDir)}${value(record?.nsCm)}cm / ${value(record?.ewDir)}${value(record?.ewCm)}cm`;
}

function renderCardOutput() {
  if (!state.records.length) {
    selectedCardRecordId = "";
    cardOutputList.innerHTML = "";
    return;
  }

  const filteredRecords = getFilteredOutputRecords();
  if (!filteredRecords.length) {
    selectedCardRecordId = "";
    cardOutputList.innerHTML = "";
    return;
  }

  const hasSelected = filteredRecords.some((item) => item.id === selectedCardRecordId);
  if (!hasSelected) {
    selectedCardRecordId = "";
    cardOutputList.innerHTML = "";
    return;
  }
  const selectedRecord = filteredRecords.find((item) => item.id === selectedCardRecordId);
  if (!selectedRecord) {
    cardOutputList.innerHTML = "";
    return;
  }

  const sectionDiagramsHtml = (selectedRecord.sectionDiagrams || []).length
    ? `<div class="card-photo-grid">${selectedRecord.sectionDiagrams
        .map(
          (item) =>
            `<figure><img src="${item.dataUrl}" alt="${escapeHtml(item.name || "diagram")}" /><figcaption>${escapeHtml(
              item.caption || ""
            )}</figcaption></figure>`
        )
        .join("")}</div>`
    : "<p class=\"muted\">断面図なし</p>";

  const photosHtml = (selectedRecord.photos || []).length
    ? `<div class="card-photo-grid">${selectedRecord.photos
        .map(
          (photo) =>
            `<figure><img src="${photo.dataUrl}" alt="${escapeHtml(photo.name || "photo")}" /><figcaption>${escapeHtml(
              photo.caption || ""
            )}</figcaption></figure>`
        )
        .join("")}</div>`
    : "<p class=\"muted\">写真なし</p>";

  cardOutputList.innerHTML = `
    <article class="card-output-item">
      <h3>${escapeHtml(selectedRecord.specimenNo)} / ${escapeHtml(selectedRecord.nameMemo || "")}</h3>
      <div class="kv-grid">
        <div><span>分類</span><strong>${escapeHtml(formatCategoryForRecord(selectedRecord))}</strong></div>
        <div><span>重要品指定</span><strong>${escapeHtml(selectedRecord.importantFlag || "")}</strong></div>
        <div><span>簡易記載</span><strong>${escapeHtml(selectedRecord.simpleRecordFlag || "-")}</strong></div>
        <div><span>地層名</span><strong>${escapeHtml(selectedRecord.layerName || "")}</strong></div>
        <div><span>ユニット</span><strong>${escapeHtml(selectedRecord.unit || "")}</strong></div>
        <div><span>細別</span><strong>${escapeHtml(formatDetailForRecord(selectedRecord))}</strong></div>
        <div><span>地層中の位置</span><strong>${escapeHtml(formatLayerPosition(selectedRecord))}</strong></div>
        <div><span>発見者</span><strong>${escapeHtml(selectedRecord.discoverer || "")}</strong></div>
        <div><span>判定者</span><strong>${escapeHtml(selectedRecord.identifier || "")}</strong></div>
        <div><span>レベル読値(上面/下底)</span><strong>${escapeHtml(formatLevelRead(selectedRecord))}</strong></div>
        <div><span>産出状況断面</span><strong>${escapeHtml(selectedRecord.occurrenceSection || "")}</strong></div>
        <div><span>産状スケッチ</span><strong>${escapeHtml(selectedRecord.occurrenceSketch || "")}</strong></div>
        <div><span>平面位置</span><strong>${escapeHtml(selectedRecord.nsDir || "")} ${escapeHtml(
          selectedRecord.nsCm || ""
        )}cm / ${escapeHtml(selectedRecord.ewDir || "")} ${escapeHtml(selectedRecord.ewCm || "")}cm</strong></div>
      </div>
      <p><strong>備考（観察事項など）:</strong> ${escapeHtml(selectedRecord.notes || "")}</p>
      <p><strong>産出状況断面図:</strong></p>
      ${sectionDiagramsHtml}
      <p><strong>写真:</strong></p>
      ${photosHtml}
    </article>
  `;
}

function getFilteredOutputRecords() {
  const sortedRecords = [...state.records].sort(compareRecordsByKuwakuThenSpecimen);
  syncOutputKuwakuSelect(sortedRecords);
  if (selectedOutputKuwaku === ALL_GRIDS_VALUE) {
    return sortedRecords;
  }
  return sortedRecords.filter((record) => kuwakuValueForSelect(getRecordKuwaku(record)) === selectedOutputKuwaku);
}

function syncOutputKuwakuSelect(records) {
  if (!outputKuwakuSelect) {
    return;
  }
  const options = collectOutputKuwakuOptions(records);
  if (!options.some((item) => item.value === selectedOutputKuwaku)) {
    selectedOutputKuwaku = ALL_GRIDS_VALUE;
  }
  outputKuwakuSelect.innerHTML = options
    .map(
      (item) =>
        `<option value="${escapeHtml(item.value)}" ${item.value === selectedOutputKuwaku ? "selected" : ""}>${escapeHtml(
          item.label
        )}</option>`
    )
    .join("");
}

function collectOutputKuwakuOptions(records) {
  if (!records.length) {
    return [{ value: ALL_GRIDS_VALUE, label: "全グリッド" }];
  }
  const kuwakuSet = new Set(records.map((record) => kuwakuValueForSelect(getRecordKuwaku(record))));
  const options = Array.from(kuwakuSet)
    .sort((a, b) => kuwakuLabelForSelect(a).localeCompare(kuwakuLabelForSelect(b), "ja", { numeric: true, sensitivity: "base" }))
    .map((kuwakuValue) => ({
      value: kuwakuValue,
      label: kuwakuLabelForSelect(kuwakuValue),
    }));
  return [{ value: ALL_GRIDS_VALUE, label: "全グリッド" }, ...options];
}

function renderPlanOutput() {
  if (!planMapWrap || !planMapLegend || !planUnitSelect || !planDetailSelect) {
    return;
  }
  planMapLegend.innerHTML = buildPlanLegendHtml();

  if (!state.records.length) {
    selectedPlanKuwaku = "";
    selectedPlanUnit = "";
    selectedPlanDetail = ALL_DETAILS_VALUE;
    syncPlanKuwakuSelect([]);
    planUnitSelect.innerHTML = "";
    planDetailSelect.innerHTML = "";
    if (planKuwakuInfo) {
      planKuwakuInfo.textContent = "区画: -";
    }
    planMapWrap.innerHTML = "<p class=\"muted\">表示対象データがありません。</p>";
    return;
  }

  const kuwakuFilteredRecords = getFilteredPlanRecords();
  if (planKuwakuInfo) {
    const kuwakuLabel = selectedPlanKuwaku ? kuwakuLabelForSelect(selectedPlanKuwaku) : "-";
    planKuwakuInfo.textContent = `区画: ${kuwakuLabel}`;
  }
  if (!kuwakuFilteredRecords.length) {
    selectedPlanUnit = "";
    selectedPlanDetail = ALL_DETAILS_VALUE;
    planUnitSelect.innerHTML = "";
    planDetailSelect.innerHTML = "";
    planMapWrap.innerHTML = "<p class=\"muted\">この区画（グリッド）には表示対象データがありません。</p>";
    return;
  }

  const units = collectPlanUnits(kuwakuFilteredRecords);
  if (!units.some((unit) => unit.value === selectedPlanUnit)) {
    selectedPlanUnit = units[0].value;
  }
  planUnitSelect.innerHTML = units
    .map(
      (unit) =>
        `<option value="${escapeHtml(unit.value)}" ${unit.value === selectedPlanUnit ? "selected" : ""}>${escapeHtml(
          unit.label
        )}</option>`
    )
    .join("");

  const unitRecords =
    selectedPlanUnit === ALL_UNITS_VALUE
      ? kuwakuFilteredRecords
      : kuwakuFilteredRecords.filter((record) => unitValueForSelect(record.unit) === selectedPlanUnit);

  const details = collectPlanDetails(unitRecords);
  if (!details.some((detail) => detail.value === selectedPlanDetail)) {
    selectedPlanDetail = details[0].value;
  }
  planDetailSelect.innerHTML = details
    .map(
      (detail) =>
        `<option value="${escapeHtml(detail.value)}" ${detail.value === selectedPlanDetail ? "selected" : ""}>${escapeHtml(
          detail.label
        )}</option>`
    )
    .join("");

  const detailRecords =
    selectedPlanDetail === ALL_DETAILS_VALUE
      ? unitRecords
      : unitRecords.filter((record) => detailValueForSelect(record.detail, record.detailSub) === selectedPlanDetail);
  const points = detailRecords.map((record) => buildPlanPoint(record)).filter(Boolean);

  if (!points.length) {
    planMapWrap.innerHTML =
      "<p class=\"muted\">このユニット/細別は、平面位置の数値が未入力のため点を表示できません。</p>";
    return;
  }

  const verticalGrid = [100, 200, 300].map((x) => `<line x1="${x}" y1="0" x2="${x}" y2="${PLAN_SIZE_CM}" />`).join("");
  const horizontalGrid = [100, 200, 300].map((y) => `<line x1="0" y1="${y}" x2="${PLAN_SIZE_CM}" y2="${y}" />`).join("");
  const pointsSvg = points
    .map((point) => {
      const labelX = Math.min(PLAN_SIZE_CM - 2, point.x + 6);
      const labelY = Math.max(8, point.y - 6);
      const ariaLabel = [
        `標本番号 ${point.label || "未設定"}`,
        `化石・遺物名称 ${point.nameMemo || "未設定"}`,
        `ユニット ${point.unit || "未設定"}`,
        `細別 ${point.detail || "未設定"}`,
      ].join(" / ");
      return `
      <g
        class="plan-point-group"
        data-label="${escapeHtml(point.label || "")}"
        data-name-memo="${escapeHtml(point.nameMemo || "")}"
        data-unit="${escapeHtml(point.unit || "")}"
        data-detail="${escapeHtml(point.detail || "")}"
        data-x="${point.x}"
        data-y="${point.y}"
        tabindex="0"
        role="button"
        aria-label="${escapeHtml(ariaLabel)}"
      >
        <circle class="plan-point-hotspot" cx="${point.x}" cy="${point.y}" r="11" fill="transparent" />
        <circle class="plan-point-hit" cx="${point.x}" cy="${point.y}" r="5" fill="${point.color}" />
        <text x="${labelX}" y="${labelY}">${escapeHtml(point.label)}</text>
      </g>
      `;
    })
    .join("");
  const cornerLabels = buildPlanCornerLabels(selectedPlanKuwaku);

  planMapWrap.innerHTML = `
    <div class="plan-map-shell">
      <div class="plan-axis north">北</div>
      <div class="plan-axis east">東</div>
      <div class="plan-axis south">南</div>
      <div class="plan-axis west">西</div>
      <div class="plan-grid-corner top-left">${escapeHtml(cornerLabels.topLeft)}</div>
      <div class="plan-grid-corner top-right">${escapeHtml(cornerLabels.topRight)}</div>
      <div class="plan-grid-corner bottom-left">${escapeHtml(cornerLabels.bottomLeft)}</div>
      <div class="plan-grid-corner bottom-right">${escapeHtml(cornerLabels.bottomRight)}</div>
      <svg class="plan-map-svg" viewBox="0 0 ${PLAN_SIZE_CM} ${PLAN_SIZE_CM}" aria-label="ユニット別平面図">
        <rect x="0" y="0" width="${PLAN_SIZE_CM}" height="${PLAN_SIZE_CM}" />
        ${verticalGrid}
        ${horizontalGrid}
        ${pointsSvg}
      </svg>
      <div class="plan-map-tooltip" hidden></div>
    </div>
  `;
  attachPlanMapTooltips();
}

function getFilteredPlanRecords() {
  const sortedRecords = [...state.records].sort(compareRecordsByKuwakuThenSpecimen);
  syncPlanKuwakuSelect(sortedRecords);
  if (!selectedPlanKuwaku) {
    return [];
  }
  return sortedRecords.filter((record) => kuwakuValueForSelect(getRecordKuwaku(record)) === selectedPlanKuwaku);
}

function syncPlanKuwakuSelect(records) {
  if (!planKuwakuSelect) {
    return;
  }
  const options = collectOutputKuwakuOptions(records).filter((item) => item.value !== ALL_GRIDS_VALUE);
  if (!options.length) {
    selectedPlanKuwaku = "";
    planKuwakuSelect.innerHTML = "";
    return;
  }
  if (!options.some((item) => item.value === selectedPlanKuwaku)) {
    selectedPlanKuwaku = options[0].value;
  }
  planKuwakuSelect.innerHTML = options
    .map(
      (item) =>
        `<option value="${escapeHtml(item.value)}" ${item.value === selectedPlanKuwaku ? "selected" : ""}>${escapeHtml(
          item.label
        )}</option>`
    )
    .join("");
}

function collectPlanUnits(records) {
  const unitSet = new Set(records.map((record) => unitValueForSelect(record.unit)));
  const unitOptions = Array.from(unitSet)
    .sort((a, b) => unitLabelForSelect(a).localeCompare(unitLabelForSelect(b), "ja", { numeric: true, sensitivity: "base" }))
    .map((unitValue) => ({
      value: unitValue,
      label: unitLabelForSelect(unitValue),
    }));
  return [{ value: ALL_UNITS_VALUE, label: "全ユニット" }, ...unitOptions];
}

function unitValueForSelect(unitRaw) {
  const unit = value(unitRaw);
  return unit || EMPTY_UNIT_VALUE;
}

function unitLabelForSelect(unitValue) {
  return unitValue === EMPTY_UNIT_VALUE ? "（未設定）" : unitValue;
}

function collectPlanDetails(records) {
  const detailSet = new Set(records.map((record) => detailValueForSelect(record.detail, record.detailSub)));
  const detailOptions = Array.from(detailSet)
    .sort((a, b) => detailLabelForSelect(a).localeCompare(detailLabelForSelect(b), "ja", { numeric: true, sensitivity: "base" }))
    .map((detailValue) => ({
      value: detailValue,
      label: detailLabelForSelect(detailValue),
    }));
  return [{ value: ALL_DETAILS_VALUE, label: "全細別" }, ...detailOptions];
}

function buildDetailText(detailRaw, detailSubRaw = "") {
  const detail = value(detailRaw);
  const detailSub = value(detailSubRaw);
  if (detail && detailSub) {
    return `${detail} ${detailSub}`;
  }
  return detail || detailSub;
}

function formatDetailForRecord(record) {
  return buildDetailText(record?.detail, record?.detailSub);
}

function detailValueForSelect(detailRaw, detailSubRaw = "") {
  const detail = buildDetailText(detailRaw, detailSubRaw);
  return detail || EMPTY_DETAIL_VALUE;
}

function detailLabelForSelect(detailValue) {
  return detailValue === EMPTY_DETAIL_VALUE ? "（未設定）" : detailValue;
}

function getRecordKuwaku(record) {
  return normalizeKuwakuText(record?.kuwaku);
}

function getRecordTeamValue(record) {
  const teamState = normalizeTeamState(value(record?.team), value(record?.teamOther));
  if (teamState.team) {
    return formatTeamValue(teamState);
  }
  return formatTeamValue(state.site);
}

function getRecordLevelHeight(record) {
  return value(record?.levelHeight) || value(state.site?.levelHeight);
}

function getRecordDate(record) {
  return value(record?.date) || value(state.site?.date);
}

function getRecordTeamLead(record) {
  return value(record?.teamLead) || value(state.site?.teamLead);
}

function getRecordRecorder(record) {
  return value(record?.recorder) || value(state.site?.recorder);
}

function kuwakuValueForSelect(kuwakuRaw) {
  const kuwaku = normalizeKuwakuText(kuwakuRaw);
  return kuwaku || EMPTY_KUWAKU_VALUE;
}

function kuwakuLabelForSelect(kuwakuValue) {
  return kuwakuValue === EMPTY_KUWAKU_VALUE ? "（未設定）" : kuwakuValue;
}

function isDefaultKuwaku(kuwakuRaw) {
  return normalizeKuwakuText(kuwakuRaw) === DEFAULT_KUWAKU;
}

function buildPlanPoint(record) {
  const nsCm = parseDistanceToCm(record.nsCm);
  const ewCm = parseDistanceToCm(record.ewCm);
  if (nsCm == null || ewCm == null) {
    return null;
  }

  const nsDir = normalizeNsDir(record.nsDir);
  const ewDir = normalizeEwDir(record.ewDir);
  const yRaw = nsDir === "北から" ? nsCm : PLAN_SIZE_CM - nsCm;
  const xRaw = ewDir === "西から" ? ewCm : PLAN_SIZE_CM - ewCm;
  const x = clamp(xRaw, 0, PLAN_SIZE_CM);
  const y = clamp(yRaw, 0, PLAN_SIZE_CM);

  const specimen = parseSpecimenNo(record.specimenNo, record.specimenPrefix, record.specimenSerial);
  const prefix = normalizeSpecimenPrefix(specimen.prefix);
  const color = SPECIMEN_POINT_COLORS[prefix] || SPECIMEN_POINT_COLORS.h;
  return {
    x,
    y,
    color,
    label: record.specimenNo || "",
    nameMemo: value(record.nameMemo),
    unit: value(record.unit),
    detail: buildDetailText(record.detail, record.detailSub),
  };
}

function parseDistanceToCm(distanceRaw) {
  const text = value(distanceRaw).replace(",", ".");
  if (!text) {
    return null;
  }
  const matched = text.match(/-?\d+(?:\.\d+)?/);
  if (!matched) {
    return null;
  }
  const num = Number(matched[0]);
  return Number.isFinite(num) ? num : null;
}

function buildPlanCornerLabels(kuwakuRaw) {
  const parts = parseKuwaku(kuwakuLabelForSelect(kuwakuRaw));
  const block = value(parts.block).toUpperCase();
  const no = value(parts.no);
  if (!block || !no) {
    return {
      topLeft: "-",
      topRight: "-",
      bottomLeft: "-",
      bottomRight: "-",
    };
  }
  const rightBlock = incrementGridBlock(block, 1);
  const lowerNo = incrementGridNo(no, 1);
  return {
    topLeft: `${block}-${no}`,
    topRight: `${rightBlock}-${no}`,
    bottomLeft: `${block}-${lowerNo}`,
    bottomRight: `${rightBlock}-${lowerNo}`,
  };
}

function incrementGridBlock(blockRaw, step) {
  const block = value(blockRaw).toUpperCase();
  if (!/^[A-Z]+$/.test(block)) {
    return block;
  }
  if (/^[A-Z]$/.test(block)) {
    const base = block.charCodeAt(0) - 65;
    const next = ((base + step) % 26 + 26) % 26;
    return String.fromCharCode(65 + next);
  }
  let colNumber = 0;
  for (const char of block) {
    colNumber = colNumber * 26 + (char.charCodeAt(0) - 64);
  }
  colNumber += step;
  if (colNumber <= 0) {
    return block;
  }
  let next = "";
  let current = colNumber;
  while (current > 0) {
    const remainder = (current - 1) % 26;
    next = String.fromCharCode(65 + remainder) + next;
    current = Math.floor((current - 1) / 26);
  }
  return next;
}

function incrementGridNo(noRaw, step) {
  const raw = value(noRaw);
  if (!/^-?\d+$/.test(raw)) {
    return raw;
  }
  return String(Number(raw) + step);
}

function buildPlanLegendHtml() {
  const order = ["m", "a", "b", "l", "s", "i", "g", "h"];
  return order
    .map((prefix) => {
      const color = SPECIMEN_POINT_COLORS[prefix] || SPECIMEN_POINT_COLORS.h;
      const label = SPECIMEN_CATEGORY_MAP[prefix] || "";
      return `<span class="plan-legend-item"><span class="plan-legend-dot" style="background:${color}"></span>${prefix}: ${label}</span>`;
    })
    .join("");
}

function attachPlanMapTooltips() {
  if (!planMapWrap) {
    return;
  }
  const shell = planMapWrap.querySelector(".plan-map-shell");
  const svg = shell?.querySelector(".plan-map-svg");
  const tooltip = shell?.querySelector(".plan-map-tooltip");
  if (!shell || !svg || !tooltip) {
    return;
  }

  const hide = () => {
    tooltip.hidden = true;
  };

  const show = (pointEl, mouseEvent = null) => {
    const specimenNo = value(pointEl.dataset.label) || "未設定";
    const nameMemo = value(pointEl.dataset.nameMemo) || "未設定";
    const unit = value(pointEl.dataset.unit) || "未設定";
    const detail = value(pointEl.dataset.detail) || "未設定";
    tooltip.innerHTML = `
      <div><strong>標本番号:</strong> ${escapeHtml(specimenNo)}</div>
      <div><strong>化石・遺物名称:</strong> ${escapeHtml(nameMemo)}</div>
      <div><strong>ユニット:</strong> ${escapeHtml(unit)}</div>
      <div><strong>細別:</strong> ${escapeHtml(detail)}</div>
    `;
    tooltip.hidden = false;
    positionTooltip(pointEl, mouseEvent, shell, svg, tooltip);
  };

  const points = shell.querySelectorAll(".plan-point-group");
  points.forEach((pointEl) => {
    pointEl.addEventListener("mouseenter", (event) => show(pointEl, event));
    pointEl.addEventListener("mousemove", (event) => {
      if (!tooltip.hidden) {
        positionTooltip(pointEl, event, shell, svg, tooltip);
      }
    });
    pointEl.addEventListener("mouseleave", hide);
    pointEl.addEventListener("focus", () => show(pointEl));
    pointEl.addEventListener("blur", hide);
    pointEl.addEventListener("click", (event) => {
      event.stopPropagation();
      show(pointEl, event);
    });
  });

  shell.addEventListener("click", (event) => {
    if (event.target.closest(".plan-point-group")) {
      return;
    }
    hide();
  });
}

function positionTooltip(pointEl, mouseEvent, shell, svg, tooltip) {
  const shellRect = shell.getBoundingClientRect();
  const svgRect = svg.getBoundingClientRect();

  let xLocal = 0;
  let yLocal = 0;
  if (mouseEvent && typeof mouseEvent.clientX === "number" && typeof mouseEvent.clientY === "number") {
    xLocal = mouseEvent.clientX - shellRect.left;
    yLocal = mouseEvent.clientY - shellRect.top;
  } else {
    const x = Number(pointEl.dataset.x);
    const y = Number(pointEl.dataset.y);
    xLocal = svgRect.left - shellRect.left + (x / PLAN_SIZE_CM) * svgRect.width;
    yLocal = svgRect.top - shellRect.top + (y / PLAN_SIZE_CM) * svgRect.height;
  }

  const offset = 14;
  const desiredLeft = xLocal + offset;
  const desiredTop = yLocal + offset;
  const maxLeft = Math.max(8, shellRect.width - tooltip.offsetWidth - 8);
  const maxTop = Math.max(8, shellRect.height - tooltip.offsetHeight - 8);
  tooltip.style.left = `${clamp(desiredLeft, 8, maxLeft)}px`;
  tooltip.style.top = `${clamp(desiredTop, 8, maxTop)}px`;
}

function renderPhotoList() {
  if (!currentPhotos.length) {
    photoList.innerHTML = "<p>写真はまだありません。</p>";
    return;
  }

  photoList.innerHTML = currentPhotos
    .map(
      (photo) => `
      <article class="photo-card">
        <img src="${photo.dataUrl}" alt="${escapeHtml(photo.name || "photo")}" />
        <input
          class="caption"
          data-photo-id="${photo.id}"
          type="text"
          placeholder="写真キャプション"
          value="${escapeHtml(photo.caption || "")}"
        />
        <div class="panel-actions">
          <button class="danger" type="button" data-remove-photo-id="${photo.id}">写真削除</button>
        </div>
      </article>
      `
    )
    .join("");
}

function renderSectionDiagramList() {
  if (!currentSectionDiagrams.length) {
    sectionDiagramList.innerHTML = "<p>断面図はまだありません。</p>";
    return;
  }

  sectionDiagramList.innerHTML = currentSectionDiagrams
    .map(
      (item) => `
      <article class="photo-card">
        <img src="${item.dataUrl}" alt="${escapeHtml(item.name || "diagram")}" />
        <input
          class="caption"
          data-diagram-id="${item.id}"
          type="text"
          placeholder="断面図キャプション"
          value="${escapeHtml(item.caption || "")}"
        />
        <div class="panel-actions">
          <button class="danger" type="button" data-remove-diagram-id="${item.id}">断面図削除</button>
        </div>
      </article>
      `
    )
    .join("");
}

function loadState() {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) {
    stateNeedsRewriteAfterLoad = false;
    return createInitialState();
  }
  try {
    const parsed = JSON.parse(raw);
    const normalized = normalizeState(parsed);
    stateNeedsRewriteAfterLoad = hasSpacingNormalizationDiff(parsed, normalized);
    return normalized;
  } catch (_error) {
    stateNeedsRewriteAfterLoad = false;
    return createInitialState();
  }
}

function buildSpacingNormalizationFingerprint(candidateState) {
  const stateCandidate = candidateState && typeof candidateState === "object" ? candidateState : {};
  const site = stateCandidate.site && typeof stateCandidate.site === "object" ? stateCandidate.site : {};
  const records = Array.isArray(stateCandidate.records) ? stateCandidate.records : [];

  return {
    site: {
      kuwaku: value(site.kuwaku),
      kuwakuHeadA: value(site.kuwakuHeadA),
      kuwakuHeadB: value(site.kuwakuHeadB),
      kuwakuBlock: value(site.kuwakuBlock),
      kuwakuNo: value(site.kuwakuNo),
    },
    records: records.map((record) => {
      return {
        id: value(record?.id),
        kuwaku: value(record?.kuwaku),
        specimenNo: value(record?.specimenNo),
        specimenPrefix: value(record?.specimenPrefix),
        specimenSerial: value(record?.specimenSerial),
        unit: value(record?.unit),
        detail: value(record?.detail),
      };
    }),
  };
}

function hasSpacingNormalizationDiff(beforeState, afterState) {
  try {
    return (
      JSON.stringify(buildSpacingNormalizationFingerprint(beforeState)) !==
      JSON.stringify(buildSpacingNormalizationFingerprint(afterState))
    );
  } catch (_error) {
    return false;
  }
}

function normalizeState(candidate) {
  const safe = createInitialState();
  if (!candidate || typeof candidate !== "object") {
    return safe;
  }

  const kuwakuParts = parseKuwaku(value(candidate.site?.kuwaku));
  const kuwakuHeadA = normalizeKuwakuHeadA(value(candidate.site?.kuwakuHeadA) || kuwakuParts.headA || DEFAULT_KUWAKU_HEAD_A);
  const kuwakuHeadB = normalizeKuwakuHeadB(value(candidate.site?.kuwakuHeadB) || kuwakuParts.headB || DEFAULT_KUWAKU_HEAD_B);
  const kuwakuBlock = normalizeKuwakuBlock(value(candidate.site?.kuwakuBlock) || kuwakuParts.block);
  const kuwakuNo = normalizeKuwakuNo(value(candidate.site?.kuwakuNo) || kuwakuParts.no);
  const teamState = normalizeTeamState(value(candidate.site?.team), value(candidate.site?.teamOther));

  safe.site = {
    kuwaku: buildKuwaku(kuwakuHeadA, kuwakuHeadB, kuwakuBlock, kuwakuNo),
    kuwakuHeadA,
    kuwakuHeadB,
    kuwakuBlock,
    kuwakuNo,
    levelHeight: value(candidate.site?.levelHeight),
    date: value(candidate.site?.date),
    team: teamState.team,
    teamOther: teamState.teamOther,
    teamLead: value(candidate.site?.teamLead),
    recorder: value(candidate.site?.recorder),
  };

  if (Array.isArray(candidate.records)) {
    safe.records = candidate.records.map((item) => normalizeRecord(item, safe.site)).filter(Boolean);
    return safe;
  }

  const artifacts = Array.isArray(candidate.artifacts) ? candidate.artifacts : [];
  const cards = candidate.cards && typeof candidate.cards === "object" ? candidate.cards : {};
  const photos = candidate.photos && typeof candidate.photos === "object" ? candidate.photos : {};

  safe.records = artifacts
    .map((artifact) => {
      if (!artifact || typeof artifact !== "object") {
        return null;
      }
      const id = value(artifact.id);
      if (!id) {
        return null;
      }
      const card = cards[id] && typeof cards[id] === "object" ? cards[id] : {};
      const recordPhotos = Array.isArray(photos[id]) ? photos[id] : [];

      return normalizeRecord({
        id,
        kuwaku: value(artifact.kuwaku) || value(candidate.site?.kuwaku),
        specimenNo: value(artifact.specimenNo),
        specimenPrefix: value(artifact.specimenPrefix),
        specimenSerial: value(artifact.specimenSerial),
        category:
          value(artifact.category) || value(artifact.categories?.[0]) || categoryFromPrefix(value(artifact.specimenPrefix)),
        analysisType: value(artifact.analysisType) || value(card.analysisType) || extractAnalysisTypeFromCategory(value(artifact.category)),
        levelHeight: value(artifact.levelHeight) || value(candidate.site?.levelHeight),
        date: value(artifact.date) || value(candidate.site?.date),
        team: value(artifact.team) || value(candidate.site?.team),
        teamOther: value(artifact.teamOther) || value(candidate.site?.teamOther),
        teamLead: value(artifact.teamLead) || value(candidate.site?.teamLead),
        recorder: value(artifact.recorder) || value(candidate.site?.recorder),
        nameMemo: value(artifact.nameMemo),
        unit: value(artifact.unit),
        discoverer: value(artifact.discoverer),
        identifier: value(artifact.identifier),
        levelUpperCm: value(card.levelUpperCm) || value(artifact.levelUpperCm) || value(artifact.levelRead) || value(artifact.levelError),
        levelLowerCm: value(card.levelLowerCm) || value(artifact.levelLowerCm),
        occurrenceSection: value(artifact.occurrenceSection) || value(artifact.sectionSketch),
        occurrenceSketch: value(artifact.occurrenceSketch) || value(artifact.sectionSketch),
        nsDir: value(artifact.nsDir),
        nsCm: value(artifact.nsCm),
        ewDir: value(artifact.ewDir),
        ewCm: value(artifact.ewCm),
        importantFlag: value(card.isImportant),
        simpleRecordFlag: value(card.simpleRecordFlag),
        layerName: value(card.layerName),
        detail: value(card.detail),
        detailSub: value(card.detailSub),
        layerRef: value(card.layerRef) || value(card.layerPosition),
        layerFromCm: value(card.layerFromCm),
        layerRelative: value(card.layerRelative),
        notes: mergeLegacyNotes({
          notes: value(artifact.notes),
          occurrenceNote: value(card.occurrenceNote),
          sketchNote: value(card.sketchNote),
        }),
        sectionDiagrams: normalizePhotos(card.sectionDiagrams),
        photos: recordPhotos,
        createdAt: value(artifact.createdAt),
        updatedAt: value(artifact.updatedAt),
      }, safe.site);
    })
    .filter(Boolean);

  return safe;
}

function normalizeRecord(item, fallbackSiteRaw = null) {
  if (!item || typeof item !== "object") {
    return null;
  }

  const fallbackSite = fallbackSiteRaw && typeof fallbackSiteRaw === "object" ? fallbackSiteRaw : {};
  const id = value(item.id) || newId("record");
  const parsedSpecimen = parseSpecimenNo(value(item.specimenNo), value(item.specimenPrefix), value(item.specimenSerial));
  const category = normalizeCategory(value(item.category), parsedSpecimen.prefix);
  const analysisType = normalizeAnalysisType(
    value(item.analysisType) || extractAnalysisTypeFromCategory(value(item.category))
  );
  const teamState = normalizeTeamState(
    value(item.team) || value(fallbackSite.team),
    value(item.teamOther) || value(fallbackSite.teamOther)
  );
  const rawKuwaku = normalizeKuwakuText(item.kuwaku);
  const fallbackKuwaku = normalizeKuwakuText(
    value(fallbackSite.kuwaku) || buildKuwaku(fallbackSite.kuwakuHeadA, fallbackSite.kuwakuHeadB, fallbackSite.kuwakuBlock, fallbackSite.kuwakuNo)
  );
  const kuwaku = !rawKuwaku || isDefaultKuwaku(rawKuwaku) ? fallbackKuwaku : rawKuwaku;

  return {
    id,
    kuwaku,
    specimenPrefix: parsedSpecimen.prefix,
    specimenSerial: parsedSpecimen.serial,
    specimenNo: parsedSpecimen.specimenNo,
    category,
    analysisType: parsedSpecimen.prefix === "a" ? analysisType : "",
    levelHeight: value(item.levelHeight) || value(fallbackSite.levelHeight),
    date: value(item.date) || value(fallbackSite.date),
    team: teamState.team,
    teamOther: teamState.teamOther,
    teamLead: value(item.teamLead) || value(fallbackSite.teamLead),
    recorder: value(item.recorder) || value(fallbackSite.recorder),
    nameMemo: value(item.nameMemo),
    unit: compactNoSpaceValue(item.unit),
    discoverer: value(item.discoverer),
    identifier: value(item.identifier),
    levelUpperCm: value(item.levelUpperCm) || value(item.levelRead) || value(item.levelError),
    levelLowerCm: value(item.levelLowerCm),
    occurrenceSection: normalizeNeedFlag(value(item.occurrenceSection) || value(item.sectionSketch)),
    occurrenceSketch: normalizeNeedFlag(value(item.occurrenceSketch) || value(item.sectionSketch)),
    nsDir: normalizeNsDir(value(item.nsDir)),
    nsCm: value(item.nsCm),
    ewDir: normalizeEwDir(value(item.ewDir)),
    ewCm: value(item.ewCm),
    importantFlag: normalizeHasFlag(value(item.importantFlag) || value(item.isImportant)),
    simpleRecordFlag: normalizeCircleDashFlag(value(item.simpleRecordFlag)),
    layerName: normalizeLayerName(value(item.layerName)),
    detail: compactNoSpaceValue(item.detail),
    detailSub: value(item.detailSub),
    layerRef: value(item.layerRef) || value(item.layerPosition),
    layerFromCm: value(item.layerFromCm),
    layerRelative: value(item.layerRelative),
    notes: mergeLegacyNotes({
      notes: value(item.notes),
      occurrenceNote: value(item.occurrenceNote),
      sketchNote: value(item.sketchNote),
    }),
    sectionDiagrams: normalizePhotos(item.sectionDiagrams),
    photos: normalizePhotos(item.photos),
    history: normalizeRecordHistory(item.history),
    createdAt: value(item.createdAt) || new Date().toISOString(),
    updatedAt: value(item.updatedAt) || new Date().toISOString(),
  };
}

function buildNextRecordHistory(previousRecord, nextRecord, actionRaw) {
  const previousHistory = normalizeRecordHistory(previousRecord?.history);
  const previousSnapshot = previousRecord ? createHistorySnapshot(previousRecord) : null;
  const snapshot = createHistorySnapshot(nextRecord);
  const entry = {
    id: newId("history"),
    action: value(actionRaw) || "保存",
    content: buildHistoryContent(nextRecord, snapshot),
    snapshot,
    changedKeys: getHistoryChangedKeys(previousSnapshot, snapshot),
    at: nowIso(),
  };
  return [...previousHistory, entry].slice(-50);
}

function createHistorySnapshot(record) {
  return {
    specimenNo: value(record?.specimenNo),
    nameMemo: value(record?.nameMemo),
    category: formatCategoryForRecord(record),
    layerName: value(record?.layerName),
    unit: value(record?.unit),
    detail: formatDetailForRecord(record),
    layerPosition: formatLayerPosition(record),
  };
}

function buildHistoryContent(record, snapshotRaw = null) {
  const snapshot = snapshotRaw || createHistorySnapshot(record);
  const summaryParts = HISTORY_SNAPSHOT_FIELDS.map((field) => {
    const fieldValue = value(snapshot?.[field.key]) || "-";
    return `${field.label} ${fieldValue}`;
  });
  return summaryParts.join(" / ");
}

function normalizeRecordHistory(historyRaw) {
  if (!Array.isArray(historyRaw)) {
    return [];
  }
  return historyRaw
    .filter((entry) => entry && typeof entry === "object")
    .map((entry) => ({
      id: value(entry.id) || newId("history"),
      action: value(entry.action) || "保存",
      content: value(entry.content),
      snapshot: normalizeHistorySnapshot(entry.snapshot) || extractHistorySnapshotFromContent(value(entry.content)),
      changedKeys: normalizeHistoryChangedKeys(entry.changedKeys),
      at: value(entry.at) || nowIso(),
    }))
    .filter((entry) => entry.content || entry.snapshot);
}

function normalizeHistorySnapshot(snapshotRaw) {
  if (!snapshotRaw || typeof snapshotRaw !== "object") {
    return null;
  }
  const snapshot = {};
  HISTORY_SNAPSHOT_FIELDS.forEach((field) => {
    snapshot[field.key] = value(snapshotRaw[field.key]);
  });
  return snapshot;
}

function normalizeHistoryChangedKeys(changedKeysRaw) {
  if (!Array.isArray(changedKeysRaw)) {
    return [];
  }
  const seen = new Set();
  return changedKeysRaw
    .map((key) => value(key))
    .filter((key) => HISTORY_SNAPSHOT_FIELD_KEYS.has(key))
    .filter((key) => {
      if (seen.has(key)) {
        return false;
      }
      seen.add(key);
      return true;
    });
}

function historyComparableValue(rawValue) {
  return value(rawValue) || "-";
}

function getHistoryChangedKeys(previousSnapshot, currentSnapshot) {
  if (!previousSnapshot) {
    return [];
  }
  return HISTORY_SNAPSHOT_FIELDS
    .map((field) => field.key)
    .filter((key) => historyComparableValue(currentSnapshot?.[key]) !== historyComparableValue(previousSnapshot?.[key]));
}

function extractHistorySnapshotFromContent(contentRaw) {
  const content = value(contentRaw);
  if (!content) {
    return null;
  }

  const parts = content.split(/\s*\/\s*/);
  const snapshot = {};
  HISTORY_SNAPSHOT_FIELDS.forEach((field) => {
    const part = parts.find((item) => value(item).startsWith(field.label));
    if (!part) {
      snapshot[field.key] = "";
      return;
    }
    const stripped = value(part)
      .slice(field.label.length)
      .replace(/^[:：]?\s*/, "");
    snapshot[field.key] = value(stripped);
  });

  const hasAny = HISTORY_SNAPSHOT_FIELDS.some((field) => value(snapshot[field.key]));
  return hasAny ? snapshot : null;
}

function renderHistoryContentHtml(entry, prevEntry) {
  const snapshot = entry?.snapshot || extractHistorySnapshotFromContent(value(entry?.content));
  if (!snapshot) {
    return escapeHtml(entry?.content || "");
  }
  const changedKeys = normalizeHistoryChangedKeys(entry?.changedKeys);
  const changedKeySet = new Set(changedKeys);
  const hasExplicitChangedKeys = changedKeySet.size > 0;
  const prevSnapshot = prevEntry?.snapshot || extractHistorySnapshotFromContent(value(prevEntry?.content));
  return HISTORY_SNAPSHOT_FIELDS.map((field) => {
    const currentValueRaw = value(snapshot[field.key]);
    const currentValue = currentValueRaw || "-";
    const isChanged = hasExplicitChangedKeys
      ? changedKeySet.has(field.key)
      : Boolean(prevSnapshot) &&
        historyComparableValue(currentValueRaw) !== historyComparableValue(prevSnapshot?.[field.key]);
    const className = isChanged ? "edit-history-value changed" : "edit-history-value";
    return `<span class="${className}">${escapeHtml(field.label)}: ${escapeHtml(currentValue)}</span>`;
  }).join(" / ");
}

function formatHistoryDateTime(isoRaw) {
  const iso = value(isoRaw);
  if (!iso) {
    return "-";
  }
  const date = new Date(iso);
  if (Number.isNaN(date.getTime())) {
    return iso;
  }
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  const hh = String(date.getHours()).padStart(2, "0");
  const mm = String(date.getMinutes()).padStart(2, "0");
  return `${y}/${m}/${d} ${hh}:${mm}`;
}

function mergeLegacyNotes({ notes = "", occurrenceNote = "", sketchNote = "" }) {
  const base = value(notes);
  const occ = value(occurrenceNote);
  const sketch = value(sketchNote);
  const merged = [];

  if (base) {
    merged.push(base);
  }
  if (occ && !base.includes(occ)) {
    merged.push(`産出状況メモ: ${occ}`);
  }
  if (sketch && !base.includes(sketch)) {
    merged.push(`スケッチ・観察事項メモ: ${sketch}`);
  }
  return merged.join("\n\n");
}

function normalizePhotos(photosRaw) {
  if (!Array.isArray(photosRaw)) {
    return [];
  }
  return photosRaw
    .filter((photo) => photo && typeof photo === "object")
    .map((photo) => ({
      id: value(photo.id) || newId("photo"),
      name: value(photo.name),
      dataUrl: value(photo.dataUrl),
      caption: value(photo.caption),
      createdAt: value(photo.createdAt) || new Date().toISOString(),
    }))
    .filter((photo) => photo.dataUrl);
}

function findRecord(recordId) {
  return state.records.find((item) => item.id === recordId);
}

function findDuplicateRecordByKuwakuAndSpecimen(kuwakuRaw, specimenNoRaw, excludeRecordIdRaw = "") {
  const kuwaku = normalizeKuwakuText(kuwakuRaw);
  const specimenNo = parseSpecimenNo(specimenNoRaw).specimenNo;
  const excludeRecordId = value(excludeRecordIdRaw);
  if (!kuwaku || !specimenNo) {
    return null;
  }
  return (
    state.records.find((item) => {
      if (!item || value(item.id) === excludeRecordId) {
        return false;
      }
      const itemKuwaku = normalizeKuwakuText(getRecordKuwaku(item));
      if (itemKuwaku !== kuwaku) {
        return false;
      }
      const itemSpecimenNo = parseSpecimenNo(item.specimenNo, item.specimenPrefix, item.specimenSerial).specimenNo;
      return itemSpecimenNo === specimenNo;
    }) || null
  );
}

function persist(successMessage) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    if (successMessage) {
      showToast(successMessage);
    }
    scheduleCloudSave();
  } catch (_error) {
    void recoverFromQuotaError(successMessage);
  }
}

function initCloudControls() {
  if (cloudEndpointInput) {
    cloudEndpointInput.value = cloudEndpoint;
  }
  updateCloudStatus();
}

async function bootstrapCloudSync() {
  if (!cloudEndpoint) {
    updateCloudStatus();
    return;
  }
  try {
    const response = await requestCloud("load");
    const remoteState = normalizeState(response.state);
    const remoteHasData = hasAnyStateData(remoteState);
    const localHasData = hasAnyStateData(state);
    const activeTabId = getActiveTabId();
    const canApplyRemote = activeTabId === "output-tab" || activeTabId === "plan-tab";
    if (remoteHasData) {
      cloudLastPulledAt = value(response.updatedAt) || nowIso();
      if (canApplyRemote) {
        applyStateSnapshot(remoteState);
      }
    } else if (localHasData) {
      await pushStateToCloud({ showToastOnSuccess: false, silentOnError: true });
    }
  } catch (_error) {
    updateCloudStatus("同期エラー");
  }
  startCloudPullTimer();
  updateCloudStatus();
}

async function handleCloudConnect() {
  const nextEndpoint = normalizeCloudEndpoint(value(cloudEndpointInput?.value));
  if (!nextEndpoint) {
    showToast("Google Apps Script のWebアプリURLを入力してください");
    return;
  }

  cloudEndpoint = nextEndpoint;
  saveCloudEndpoint(nextEndpoint);
  if (cloudEndpointInput) {
    cloudEndpointInput.value = nextEndpoint;
  }
  updateCloudStatus("接続確認中");

  try {
    const response = await requestCloud("load");
    const remoteState = normalizeState(response.state);
    const remoteHasData = hasAnyStateData(remoteState);
    const localHasData = hasAnyStateData(state);
    const remoteUpdatedAt = value(response.updatedAt) || getStateUpdatedAt(remoteState);
    const localUpdatedAt = getStateUpdatedAt(state);
    const remoteMs = Number.parseInt(String(Date.parse(remoteUpdatedAt || "")), 10);
    const localMs = Number.parseInt(String(Date.parse(localUpdatedAt || "")), 10);

    if (remoteHasData && localHasData && Number.isFinite(remoteMs) && Number.isFinite(localMs) && localMs > remoteMs) {
      const overwriteCloud = window.confirm(
        "端末側のデータのほうが新しい可能性があります。\nOK: 端末データでクラウドを上書き\nキャンセル: クラウドデータを読み込み"
      );
      if (overwriteCloud) {
        await pushStateToCloud({ showToastOnSuccess: false });
        showToast("共有保存を有効化し、端末データをクラウドへ保存しました");
      } else {
        const applied = applyStateSnapshot(remoteState);
        cloudLastPulledAt = remoteUpdatedAt || nowIso();
        showToast(applied ? "共有保存を有効化し、クラウドデータを読み込みました" : "共有保存を有効化しました");
      }
    } else if (remoteHasData) {
      const applied = applyStateSnapshot(remoteState);
      cloudLastPulledAt = remoteUpdatedAt || nowIso();
      showToast(applied ? "共有保存を有効化し、クラウドデータを読み込みました" : "共有保存を有効化しました");
    } else {
      await pushStateToCloud({ showToastOnSuccess: false });
      showToast("共有保存を有効化しました");
    }
    startCloudPullTimer();
    updateCloudStatus();
  } catch (_error) {
    disableCloudSync({ showToastMessage: false });
    showToast("共有保存の接続に失敗しました。URLと公開設定を確認してください");
  }
}

async function handleCloudManualReload() {
  if (!cloudEndpoint) {
    showToast("共有保存は未設定です");
    return;
  }
  const activeTabId = getActiveTabId();
  if (activeTabId === "input-tab" || activeTabId === "edit-tab") {
    const answer = window.confirm("入力途中の内容は失われる場合があります。クラウドを再読込しますか？");
    if (!answer) {
      return;
    }
  }
  await pullStateFromCloud({ force: true, showToastOnSuccess: true });
}

function disableCloudSync({ showToastMessage } = { showToastMessage: false }) {
  cloudEndpoint = "";
  saveCloudEndpoint("");
  stopCloudPullTimer();
  if (cloudSaveTimer) {
    window.clearTimeout(cloudSaveTimer);
    cloudSaveTimer = null;
  }
  cloudPushInProgress = false;
  cloudPullInProgress = false;
  cloudLastSyncedAt = "";
  cloudLastPulledAt = "";
  if (cloudEndpointInput) {
    cloudEndpointInput.value = "";
  }
  updateCloudStatus();
  if (showToastMessage) {
    showToast("共有保存をOFFにしました（端末保存のみ）");
  }
}

function startCloudPullTimer() {
  stopCloudPullTimer();
  if (!cloudEndpoint || !CLOUD_AUTO_PULL_ENABLED) {
    return;
  }
  cloudPullTimer = window.setInterval(() => {
    const activeTabId = getActiveTabId();
    if (activeTabId === "output-tab" || activeTabId === "plan-tab") {
      void pullStateFromCloud({ force: false, showToastOnSuccess: false, silentOnError: true });
    }
  }, CLOUD_PULL_INTERVAL_MS);
}

function stopCloudPullTimer() {
  if (!cloudPullTimer) {
    return;
  }
  window.clearInterval(cloudPullTimer);
  cloudPullTimer = null;
}

function scheduleCloudSave() {
  if (!cloudEndpoint) {
    return;
  }
  if (cloudSaveTimer) {
    window.clearTimeout(cloudSaveTimer);
  }
  cloudSaveTimer = window.setTimeout(() => {
    cloudSaveTimer = null;
    void pushStateToCloud({ showToastOnSuccess: false, silentOnError: true });
  }, CLOUD_SAVE_DEBOUNCE_MS);
}

async function pullStateFromCloud({ force = false, showToastOnSuccess = false, silentOnError = false } = {}) {
  if (!cloudEndpoint || cloudPullInProgress) {
    return false;
  }
  const tabAtRequest = getActiveTabId();
  if (!force && tabAtRequest !== "output-tab" && tabAtRequest !== "plan-tab") {
    return false;
  }
  cloudPullInProgress = true;
  try {
    const response = await requestCloud("load");
    const remoteState = normalizeState(response.state);
    const remoteHasData = hasAnyStateData(remoteState);
    const localHasData = hasAnyStateData(state);
    const remoteUpdatedAt = value(response.updatedAt) || getStateUpdatedAt(remoteState);
    const localUpdatedAt = getStateUpdatedAt(state);
    const remoteMs = Date.parse(remoteUpdatedAt || "");
    const localMs = Date.parse(localUpdatedAt || "");

    if (!force && !remoteHasData && localHasData) {
      return false;
    }
    if (!force && Number.isFinite(remoteMs) && Number.isFinite(localMs) && remoteMs <= localMs) {
      cloudLastPulledAt = remoteUpdatedAt || cloudLastPulledAt;
      updateCloudStatus();
      return false;
    }
    const tabBeforeApply = getActiveTabId();
    if (!force && (tabBeforeApply === "input-tab" || tabBeforeApply === "edit-tab")) {
      return false;
    }

    const applied = applyStateSnapshot(remoteState, { force });
    if (!applied) {
      return false;
    }
    cloudLastPulledAt = remoteUpdatedAt || nowIso();
    updateCloudStatus();
    if (showToastOnSuccess) {
      showToast("クラウドから最新データを読み込みました");
    }
    return true;
  } catch (_error) {
    if (!silentOnError) {
      notifyCloudError("クラウド読込に失敗しました");
    }
    updateCloudStatus("同期エラー");
    return false;
  } finally {
    cloudPullInProgress = false;
  }
}

async function pushStateToCloud({ showToastOnSuccess = false, silentOnError = false } = {}) {
  if (!cloudEndpoint || cloudPushInProgress) {
    return false;
  }
  cloudPushInProgress = true;
  try {
    const payload = {
      clientId: cloudClientId,
      updatedAt: getStateUpdatedAt(state) || nowIso(),
      state,
    };
    const response = await requestCloud("save", payload);
    cloudLastSyncedAt = value(response.updatedAt) || payload.updatedAt;
    updateCloudStatus();
    if (showToastOnSuccess) {
      showToast("クラウドへ保存しました");
    }
    return true;
  } catch (_error) {
    if (!silentOnError) {
      notifyCloudError("クラウド保存に失敗しました（端末には保存済み）");
    }
    updateCloudStatus("同期エラー");
    return false;
  } finally {
    cloudPushInProgress = false;
  }
}

async function requestCloud(action, payload = null) {
  if (!cloudEndpoint) {
    throw new Error("Cloud endpoint is not configured");
  }

  let response;
  if (action === "load") {
    const separator = cloudEndpoint.includes("?") ? "&" : "?";
    const url = `${cloudEndpoint}${separator}action=load&t=${Date.now()}`;
    response = await fetch(url, {
      method: "GET",
      cache: "no-store",
    });
  } else {
    const form = new URLSearchParams();
    form.set("action", action);
    form.set("payload", JSON.stringify(payload || {}));
    response = await fetch(cloudEndpoint, {
      method: "POST",
      body: form,
    });
  }

  const text = await response.text();
  let body = null;
  try {
    body = JSON.parse(text);
  } catch (_error) {
    throw new Error("Invalid cloud response");
  }

  if (!response.ok || !body || body.ok === false) {
    throw new Error(value(body?.error) || `HTTP ${response.status}`);
  }
  return body;
}

function applyStateSnapshot(nextStateRaw, { force = false } = {}) {
  const activeTabId = getActiveTabId();
  if (!force && (activeTabId === "input-tab" || activeTabId === "edit-tab")) {
    return false;
  }
  state = normalizeState(nextStateRaw);
  hydrateSiteForm();
  resetRecordForm({ showMessage: false });
  renderRecordTable();
  renderOutputs();
  return true;
}

function updateCloudStatus(statusNote = "") {
  if (!cloudStatusEl) {
    return;
  }
  if (!cloudEndpoint) {
    cloudStatusEl.textContent = "保存先: 端末内のみ";
    cloudStatusEl.classList.remove("online");
    return;
  }
  const latest = cloudLastSyncedAt || cloudLastPulledAt;
  const latestText = latest ? ` / 最終同期 ${formatStatusTime(latest)}` : "";
  const note = statusNote ? ` / ${statusNote}` : "";
  cloudStatusEl.textContent = `保存先: 共有Googleドライブ${latestText}${note}`;
  cloudStatusEl.classList.add("online");
}

function notifyCloudError(message) {
  const now = Date.now();
  if (now - cloudLastErrorAt < 5000) {
    return;
  }
  cloudLastErrorAt = now;
  showToast(message);
}

function normalizeCloudEndpoint(urlRaw) {
  const url = value(urlRaw);
  if (!url) {
    return "";
  }
  if (!/^https?:\/\//i.test(url)) {
    return "";
  }
  return url;
}

function loadCloudEndpoint() {
  try {
    return value(localStorage.getItem(CLOUD_ENDPOINT_KEY));
  } catch (_error) {
    return "";
  }
}

function saveCloudEndpoint(url) {
  try {
    const endpoint = value(url);
    if (endpoint) {
      localStorage.setItem(CLOUD_ENDPOINT_KEY, endpoint);
    } else {
      localStorage.removeItem(CLOUD_ENDPOINT_KEY);
    }
  } catch (_error) {
    // ignore
  }
}

function loadOrCreateCloudClientId() {
  try {
    const existing = value(localStorage.getItem(CLOUD_CLIENT_ID_KEY));
    if (existing) {
      return existing;
    }
    const clientId = newId("client");
    localStorage.setItem(CLOUD_CLIENT_ID_KEY, clientId);
    return clientId;
  } catch (_error) {
    return newId("client");
  }
}

function hasAnyStateData(candidateState) {
  const normalized = normalizeState(candidateState);
  if (normalized.records.length) {
    return true;
  }
  const site = normalized.site || {};
  return Boolean(
    value(site.kuwakuHeadA) !== DEFAULT_KUWAKU_HEAD_A ||
      value(site.kuwakuHeadB) !== DEFAULT_KUWAKU_HEAD_B ||
      value(site.kuwakuBlock) ||
      value(site.kuwakuNo) ||
      value(site.levelHeight) ||
      value(site.date) ||
      value(site.team) ||
      value(site.teamLead) ||
      value(site.recorder)
  );
}

function getStateUpdatedAt(candidateState) {
  if (!candidateState || typeof candidateState !== "object") {
    return "";
  }

  let latestMs = Number.NEGATIVE_INFINITY;
  const pushDate = (raw) => {
    const ms = Date.parse(value(raw));
    if (Number.isFinite(ms)) {
      latestMs = Math.max(latestMs, ms);
    }
  };
  pushDate(candidateState.site?.updatedAt);
  if (Array.isArray(candidateState.records)) {
    candidateState.records.forEach((record) => {
      pushDate(record?.updatedAt);
      pushDate(record?.createdAt);
    });
  }
  return Number.isFinite(latestMs) ? new Date(latestMs).toISOString() : "";
}

function formatStatusTime(isoString) {
  const date = new Date(isoString);
  if (!Number.isFinite(date.getTime())) {
    return "-";
  }
  const hh = String(date.getHours()).padStart(2, "0");
  const mm = String(date.getMinutes()).padStart(2, "0");
  const ss = String(date.getSeconds()).padStart(2, "0");
  return `${hh}:${mm}:${ss}`;
}

function buildListCsv() {
  const header = [
    "区画",
    "レベル高(m)",
    "日付",
    "発掘班",
    "班長",
    "記載係",
    "標本番号",
    "分類",
    "名称",
    "ユニット",
    "発見者",
    "判定者",
    "レベル読値_上面(cm)",
    "レベル読値_下底(cm)",
    "産出状況断面",
    "産状スケッチ",
    "平面位置_NS",
    "平面位置_NS_cm",
    "平面位置_EW",
    "平面位置_EW_cm",
    "備考（観察事項など）",
  ];

  const rows = [...state.records].sort(compareRecordsBySpecimenNo).map((record) => [
    getRecordKuwaku(record),
    getRecordLevelHeight(record),
    getRecordDate(record),
    getRecordTeamValue(record),
    getRecordTeamLead(record),
    getRecordRecorder(record),
    record.specimenNo,
    formatCategoryForRecord(record),
    record.nameMemo,
    record.unit,
    record.discoverer,
    record.identifier,
    record.levelUpperCm,
    record.levelLowerCm,
    record.occurrenceSection,
    record.occurrenceSketch,
    record.nsDir,
    record.nsCm,
    record.ewDir,
    record.ewCm,
    record.notes,
  ]);

  return [header, ...rows].map((row) => row.map(csvCell).join(",")).join("\n");
}

function buildCardCsv() {
  const header = [
    "区画",
    "標本番号",
    "分類",
    "名称",
    "重要品指定",
    "簡易記載",
    "地層名",
    "細別",
    "細別（上下など）",
    "層理面もしくは鍵層名",
    "地層中の位置_から(cm)",
    "地層中の位置_上もしくは下",
    "レベル読値_上面(cm)",
    "レベル読値_下底(cm)",
    "平面位置_NS",
    "平面位置_NS_cm",
    "平面位置_EW",
    "平面位置_EW_cm",
    "発見者",
    "判定者",
    "産出状況断面",
    "産状スケッチ",
    "備考（観察事項など）",
    "産出状況断面図枚数",
    "写真枚数",
  ];

  const rows = state.records.map((record) => [
    getRecordKuwaku(record),
    record.specimenNo,
    formatCategoryForRecord(record),
    record.nameMemo,
    record.importantFlag,
    record.simpleRecordFlag,
    record.layerName,
    record.detail,
    record.detailSub,
    record.layerRef,
    record.layerFromCm,
    record.layerRelative,
    record.levelUpperCm,
    record.levelLowerCm,
    record.nsDir,
    record.nsCm,
    record.ewDir,
    record.ewCm,
    record.discoverer,
    record.identifier,
    record.occurrenceSection,
    record.occurrenceSketch,
    record.notes,
    String((record.sectionDiagrams || []).length),
    String((record.photos || []).length),
  ]);

  return [header, ...rows].map((row) => row.map(csvCell).join(",")).join("\n");
}

function csvCell(valueRaw) {
  const valueText = value(valueRaw).replaceAll("\r\n", "\n").replaceAll("\r", "\n");
  const escaped = valueText.replaceAll('"', '""');
  return `"${escaped}"`;
}

function compactNoSpaceValue(inputRaw) {
  return value(inputRaw).replace(/\s+/g, "");
}

function normalizeKuwakuHeadA(headARaw) {
  return compactNoSpaceValue(headARaw);
}

function normalizeKuwakuHeadB(headBRaw) {
  const headB = compactNoSpaceValue(headBRaw);
  const upper = headB.toUpperCase();
  if (upper === "I" || upper === "1" || upper === "Ⅰ") {
    return "Ⅰ";
  }
  return headB;
}

function normalizeKuwakuBlock(blockRaw) {
  return compactNoSpaceValue(blockRaw).toUpperCase();
}

function normalizeKuwakuNo(noRaw) {
  return compactNoSpaceValue(noRaw);
}

function normalizeKuwakuText(kuwakuRaw) {
  const kuwaku = value(kuwakuRaw);
  if (!kuwaku) {
    return "";
  }
  const parts = parseKuwaku(kuwaku);
  return buildKuwaku(parts.headA, parts.headB, parts.block, parts.no);
}

function buildKuwaku(headARaw, headBRaw, blockRaw, noRaw) {
  const headA = normalizeKuwakuHeadA(headARaw) || DEFAULT_KUWAKU_HEAD_A;
  const headB = normalizeKuwakuHeadB(headBRaw) || DEFAULT_KUWAKU_HEAD_B;
  const block = normalizeKuwakuBlock(blockRaw);
  const no = normalizeKuwakuNo(noRaw);
  return `${headA}-${headB}-${block}-${no}`;
}

function parseKuwaku(kuwakuText) {
  const text = value(kuwakuText)
    .replaceAll("－", "-")
    .replaceAll("―", "-")
    .replaceAll("ー", "-")
    .replaceAll("−", "-");
  const parts = text.split("-").map((part) => compactNoSpaceValue(part));
  if (parts.length >= 4) {
    return {
      headA: normalizeKuwakuHeadA(parts[0]) || DEFAULT_KUWAKU_HEAD_A,
      headB: normalizeKuwakuHeadB(parts[1]) || DEFAULT_KUWAKU_HEAD_B,
      block: normalizeKuwakuBlock(parts[2]),
      no: normalizeKuwakuNo(parts[3]),
    };
  }
  return {
    headA: DEFAULT_KUWAKU_HEAD_A,
    headB: DEFAULT_KUWAKU_HEAD_B,
    block: "",
    no: "",
  };
}

function normalizeSpecimenPrefix(prefixRaw) {
  let prefix = compactNoSpaceValue(prefixRaw).toLowerCase();
  if (prefix === "ii") {
    prefix = "i";
  }
  return VALID_SPECIMEN_PREFIXES.has(prefix) ? prefix : DEFAULT_SPECIMEN_PREFIX;
}

function buildSpecimenNo(prefixRaw, serialRaw) {
  const prefix = normalizeSpecimenPrefix(prefixRaw);
  const serial = compactNoSpaceValue(serialRaw);
  return serial ? `${prefix}-${serial}` : "";
}

function parseSpecimenNo(specimenNoRaw, prefixRaw = "", serialRaw = "") {
  const directPrefix = normalizeSpecimenPrefix(compactNoSpaceValue(prefixRaw));
  const directSerial = compactNoSpaceValue(serialRaw);
  if (directSerial) {
    return {
      prefix: directPrefix,
      serial: directSerial,
      specimenNo: buildSpecimenNo(directPrefix, directSerial),
    };
  }

  const specimenNo = compactNoSpaceValue(specimenNoRaw);
  const hyphenMatched = specimenNo.match(/^([A-Za-z]{1,2})-(.+)$/);
  if (hyphenMatched) {
    const prefix = normalizeSpecimenPrefix(hyphenMatched[1]);
    const serial = compactNoSpaceValue(hyphenMatched[2]);
    return {
      prefix,
      serial,
      specimenNo: buildSpecimenNo(prefix, serial),
    };
  }

  const compactMatched = specimenNo.match(/^([A-Za-z]{1,2})(\d.+)$/);
  if (compactMatched) {
    const prefix = normalizeSpecimenPrefix(compactMatched[1]);
    const serial = compactNoSpaceValue(compactMatched[2]);
    return {
      prefix,
      serial,
      specimenNo: buildSpecimenNo(prefix, serial),
    };
  }

  const fallbackPrefix = normalizeSpecimenPrefix(compactNoSpaceValue(prefixRaw));
  return {
    prefix: fallbackPrefix,
    serial: specimenNo,
    specimenNo: buildSpecimenNo(fallbackPrefix, specimenNo),
  };
}

function compareRecordsBySpecimenNo(a, b) {
  const aSpecimen = parseSpecimenNo(a?.specimenNo, a?.specimenPrefix, a?.specimenSerial);
  const bSpecimen = parseSpecimenNo(b?.specimenNo, b?.specimenPrefix, b?.specimenSerial);

  const prefixCompared = aSpecimen.prefix.localeCompare(bSpecimen.prefix, "ja", { sensitivity: "base" });
  if (prefixCompared !== 0) {
    return prefixCompared;
  }

  const aSerial = value(aSpecimen.serial);
  const bSerial = value(bSpecimen.serial);
  const aIsNumber = /^\d+$/.test(aSerial);
  const bIsNumber = /^\d+$/.test(bSerial);
  if (aIsNumber && bIsNumber) {
    const diff = Number(aSerial) - Number(bSerial);
    if (diff !== 0) {
      return diff;
    }
  } else {
    const serialCompared = aSerial.localeCompare(bSerial, "ja", { numeric: true, sensitivity: "base" });
    if (serialCompared !== 0) {
      return serialCompared;
    }
  }

  return value(a?.id).localeCompare(value(b?.id), "ja", { sensitivity: "base" });
}

function compareRecordsByKuwakuThenSpecimen(a, b) {
  const aKuwaku = kuwakuLabelForSelect(kuwakuValueForSelect(getRecordKuwaku(a)));
  const bKuwaku = kuwakuLabelForSelect(kuwakuValueForSelect(getRecordKuwaku(b)));
  const kuwakuCompared = aKuwaku.localeCompare(bKuwaku, "ja", { numeric: true, sensitivity: "base" });
  if (kuwakuCompared !== 0) {
    return kuwakuCompared;
  }
  return compareRecordsBySpecimenNo(a, b);
}

function categoryFromPrefix(prefixRaw) {
  const prefix = normalizeSpecimenPrefix(prefixRaw);
  return `${prefix}: ${SPECIMEN_CATEGORY_MAP[prefix] || ""}`;
}

function normalizeCategory(categoryRaw, prefixRaw) {
  const categoryText = value(categoryRaw);
  const matched = categoryText.match(/^([A-Za-z]{1,2})\s*:\s*(.*)$/);
  if (matched) {
    const prefix = normalizeSpecimenPrefix(matched[1]);
    return `${prefix}: ${SPECIMEN_CATEGORY_MAP[prefix] || value(matched[2])}`;
  }
  if (categoryText) {
    return categoryText;
  }
  return categoryFromPrefix(prefixRaw);
}

function normalizeAnalysisType(typeRaw) {
  const text = value(typeRaw);
  if (!text) {
    return "";
  }
  const matched = text.match(/^([A-Za-z]{1,2})\s*:/);
  const code = (matched ? matched[1] : text).replaceAll(" ", "").toUpperCase();
  if (!ANALYSIS_TYPE_MAP[code]) {
    return "";
  }
  const displayCode = code === "MG" ? "Mg" : code;
  return `${displayCode}: ${ANALYSIS_TYPE_MAP[code]}`;
}

function extractAnalysisTypeFromCategory(categoryRaw) {
  const text = value(categoryRaw);
  if (!text) {
    return "";
  }
  const slashIndex = text.indexOf("/");
  if (slashIndex < 0 && /^a\s*:/i.test(text)) {
    return "";
  }
  const candidate = slashIndex >= 0 ? text.slice(slashIndex + 1) : text;
  const matched = candidate.match(/([A-Za-z]{1,2})\s*:/);
  if (!matched) {
    return "";
  }
  return normalizeAnalysisType(matched[1]);
}

function formatCategoryForRecord(record) {
  const base = normalizeCategory(value(record?.category), value(record?.specimenPrefix));
  const prefix = normalizeSpecimenPrefix(value(record?.specimenPrefix));
  if (prefix !== "a") {
    return base;
  }
  const analysisType = normalizeAnalysisType(value(record?.analysisType));
  if (!analysisType) {
    return base;
  }
  return `${base} / ${analysisType}`;
}

function normalizeLayerName(layerRaw) {
  const layer = value(layerRaw);
  if (!layer) {
    return "";
  }
  if (LEGACY_LAYER_NAME_ALIASES[layer]) {
    return LEGACY_LAYER_NAME_ALIASES[layer];
  }
  const legacyEntries = Object.entries(LEGACY_LAYER_NAME_ALIASES);
  for (const [legacy, latest] of legacyEntries) {
    if (layer.startsWith(`${legacy}:`)) {
      return `${latest}:${layer.slice(`${legacy}:`.length)}`;
    }
    if (layer.startsWith(`${legacy}：`)) {
      return `${latest}：${layer.slice(`${legacy}：`.length)}`;
    }
    if (layer.startsWith(`${legacy}(`) && layer.endsWith(")")) {
      return `${latest}${layer.slice(legacy.length)}`;
    }
  }
  return layer;
}

function extractOtherLayerText(layerRaw) {
  const layer = normalizeLayerName(layerRaw);
  if (!layer) {
    return "";
  }
  if (layer.startsWith(`${OTHER_LAYER_NAME}:`)) {
    return layer.slice(`${OTHER_LAYER_NAME}:`.length).trim();
  }
  if (layer.startsWith(`${OTHER_LAYER_NAME}：`)) {
    return layer.slice(`${OTHER_LAYER_NAME}：`.length).trim();
  }
  if (layer.startsWith(`${OTHER_LAYER_NAME}(`) && layer.endsWith(")")) {
    return layer.slice(OTHER_LAYER_NAME.length + 1, -1).trim();
  }
  if (layer === OTHER_LAYER_NAME) {
    return "";
  }
  return layer;
}

function extractOtherTeamText(teamRaw) {
  const team = value(teamRaw);
  if (!team) {
    return "";
  }
  if (team.startsWith(`${OTHER_TEAM_NAME}:`)) {
    return team.slice(`${OTHER_TEAM_NAME}:`.length).trim();
  }
  if (team.startsWith(`${OTHER_TEAM_NAME}：`)) {
    return team.slice(`${OTHER_TEAM_NAME}：`.length).trim();
  }
  if (team === OTHER_TEAM_NAME) {
    return "";
  }
  return team;
}

function normalizeTeamState(teamRaw, teamOtherRaw = "") {
  const team = value(teamRaw);
  const teamOther = value(teamOtherRaw);

  if (!team) {
    return { team: "", teamOther: "" };
  }
  if (PRESET_TEAMS.includes(team) && team !== OTHER_TEAM_NAME) {
    return { team, teamOther: "" };
  }
  if (team === OTHER_TEAM_NAME) {
    return { team: OTHER_TEAM_NAME, teamOther };
  }
  return { team: OTHER_TEAM_NAME, teamOther: teamOther || extractOtherTeamText(team) };
}

function syncTeamOtherInput(teamRaw) {
  const team = value(teamRaw);
  const isOther = team === OTHER_TEAM_NAME;
  teamOtherInput.classList.toggle("hidden", !isOther);
  if (!isOther) {
    teamOtherInput.value = "";
  }
}

function syncEditTeamOtherInput(teamRaw) {
  if (!editTeamOtherInput) {
    return;
  }
  const team = value(teamRaw);
  const isOther = team === OTHER_TEAM_NAME;
  editTeamOtherInput.classList.toggle("hidden", !isOther);
  if (!isOther) {
    editTeamOtherInput.value = "";
  }
}

function formatTeamValue(site) {
  const teamState = normalizeTeamState(site?.team, site?.teamOther);
  if (teamState.team !== OTHER_TEAM_NAME) {
    return teamState.team;
  }
  return teamState.teamOther ? `${OTHER_TEAM_NAME}:${teamState.teamOther}` : OTHER_TEAM_NAME;
}

function validateInputRequiredFields(siteSnapshot, recordFormData) {
  if (!siteSnapshot) {
    return "入力情報を取得できませんでした";
  }

  const siteRequiredFields = [
    ["区画（グリッド）名の1番目", siteSnapshot.kuwakuHeadA],
    ["区画（グリッド）名の2番目", siteSnapshot.kuwakuHeadB],
    ["区画（グリッド）の英字", siteSnapshot.kuwakuBlock],
    ["区画（グリッド）の番号", siteSnapshot.kuwakuNo],
    ["レベル高", siteSnapshot.levelHeight],
    ["日付", siteSnapshot.date],
    ["発掘班", siteSnapshot.team],
    ["班長", siteSnapshot.teamLead],
    ["記載係", siteSnapshot.recorder],
  ];
  for (const [label, fieldValue] of siteRequiredFields) {
    if (!value(fieldValue)) {
      return `${label}を入力してください`;
    }
  }
  if (siteSnapshot.team === OTHER_TEAM_NAME && !value(siteSnapshot.teamOther)) {
    return "発掘班が「その他」の場合は内容を入力してください";
  }

  const selectedLayerName = getSelectedLayerName();
  const recordRequiredFields = [
    ["標本番号", recordFormData.get("specimenSerial")],
    ["化石・遺物名称", recordFormData.get("nameMemo")],
    ["重要品指定", recordFormData.get("importantFlag")],
    ["簡易記載（専門班の指示による）", recordFormData.get("simpleRecordFlag")],
    ["発見者氏名", recordFormData.get("discoverer")],
    ["判定者氏名", recordFormData.get("identifier")],
    ["レベル読値（上面）", recordFormData.get("levelUpperCm")],
    ["レベル読値（下底）", recordFormData.get("levelLowerCm")],
    ["産出状況断面", recordFormData.get("occurrenceSection")],
    ["産状スケッチ", recordFormData.get("occurrenceSketch")],
    ["平面位置（北から/南から）", recordFormData.get("nsDir")],
    ["平面位置（北から/南からの距離）", recordFormData.get("nsCm")],
    ["平面位置（東から/西から）", recordFormData.get("ewDir")],
    ["平面位置（東から/西からの距離）", recordFormData.get("ewCm")],
    ["地層名", selectedLayerName],
    ["ユニット", recordFormData.get("unit")],
    ["細別", recordFormData.get("detail")],
    ["層理面や鍵層名", recordFormData.get("layerRef")],
    ["地層中の位置（cm）", recordFormData.get("layerFromCm")],
  ];
  for (const [label, fieldValue] of recordRequiredFields) {
    if (!value(fieldValue)) {
      return `${label}を入力してください`;
    }
  }
  if (selectedLayerName === OTHER_LAYER_NAME && !value(layerOtherInput.value)) {
    return "地層名が「4.その他」の場合は内容を入力してください";
  }
  const specimenPrefix = normalizeSpecimenPrefix(value(recordFormData.get("specimenPrefix")));
  if (specimenPrefix === "a" && !normalizeAnalysisType(value(recordFormData.get("analysisType")))) {
    return "a: 分析用資料を選んだ場合は、区分を選択してください";
  }

  return "";
}

function normalizeNeedFlag(valueRaw) {
  return value(valueRaw) === "否" ? "否" : "要";
}

function normalizeHasFlag(valueRaw) {
  const text = value(valueRaw);
  if (text === "有" || text === "無") {
    return text;
  }
  return "";
}

function normalizeCircleDashFlag(valueRaw) {
  const text = value(valueRaw);
  if (text === "○" || text === "◯") {
    return "○";
  }
  return "-";
}

function normalizeNsDir(valueRaw) {
  const dir = value(valueRaw);
  if (dir === "南" || dir === "南から") {
    return "南から";
  }
  return "北から";
}

function normalizeEwDir(valueRaw) {
  const dir = value(valueRaw);
  if (dir === "西" || dir === "西から") {
    return "西から";
  }
  return "東から";
}

function normalizeDirectionValue(group, valueRaw) {
  return group === "ns" ? normalizeNsDir(valueRaw) : normalizeEwDir(valueRaw);
}

function formatLevelRead(record) {
  const upper = value(record?.levelUpperCm);
  const lower = value(record?.levelLowerCm);
  if (!upper && !lower) {
    return "";
  }
  return `${upper || "-"} / ${lower || "-"}`;
}

function formatLayerPosition(record) {
  const ref = value(record?.layerRef);
  const fromCm = value(record?.layerFromCm);
  const relative = value(record?.layerRelative);
  const line1 = ref;
  let line2 = "";
  if (relative && fromCm) {
    line2 = `${relative} に ${fromCm}cm`;
  } else if (relative) {
    line2 = relative;
  } else if (fromCm) {
    line2 = `${fromCm}cm`;
  }

  if (!line1 && !line2) {
    return "";
  }
  if (!line1) {
    return line2;
  }
  if (!line2) {
    return line1;
  }
  return `${line1} / ${line2}`;
}

function clonePhotos(photos) {
  return normalizePhotos(photos).map((photo) => ({ ...photo }));
}

function normalizeAsciiWidth(inputText) {
  return String(inputText)
    .replace(/\u3000/g, " ")
    .replace(/[！-～]/g, (char) => String.fromCharCode(char.charCodeAt(0) - 0xfee0));
}

function value(input) {
  return input == null ? "" : normalizeAsciiWidth(String(input)).trim();
}

function clamp(number, min, max) {
  return Math.min(max, Math.max(min, number));
}

function nowIso() {
  return new Date().toISOString();
}

function newId(prefix) {
  if (window.crypto?.randomUUID) {
    return `${prefix}-${window.crypto.randomUUID()}`;
  }
  return `${prefix}-${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
}

function showToast(message) {
  toastEl.textContent = message;
  toastEl.classList.add("show");
  if (toastTimer) {
    window.clearTimeout(toastTimer);
  }
  toastTimer = window.setTimeout(() => {
    toastEl.classList.remove("show");
  }, 2200);
}

function downloadFile(fileName, text, mimeType) {
  const blob = new Blob([text], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

function escapeHtml(input) {
  return String(input || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function timestamp() {
  const now = new Date();
  const yyyy = String(now.getFullYear());
  const mm = String(now.getMonth() + 1).padStart(2, "0");
  const dd = String(now.getDate()).padStart(2, "0");
  const hh = String(now.getHours()).padStart(2, "0");
  const mi = String(now.getMinutes()).padStart(2, "0");
  return `${yyyy}${mm}${dd}-${hh}${mi}`;
}

async function recoverFromQuotaError(successMessage) {
  if (quotaRecoveryInProgress) {
    showToast("写真データを圧縮中です。数秒待ってから再試行してください");
    return;
  }
  quotaRecoveryInProgress = true;
  showToast("保存容量を超えたため写真を圧縮しています…");
  try {
    for (const step of PHOTO_COMPRESSION_STEPS) {
      const changed = await recompressAllPhotos(step.maxLength, step.quality);
      if (!changed) {
        continue;
      }
      try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
        renderSectionDiagramList();
        renderPhotoList();
        renderOutputs();
        if (successMessage) {
          showToast(`${successMessage}（写真を圧縮して保存）`);
        } else {
          showToast("写真を圧縮して保存しました");
        }
        return;
      } catch (_error) {
        // 次の圧縮段階で再試行
      }
    }
    showToast("保存に失敗しました。写真を一部削除して再試行してください");
  } finally {
    quotaRecoveryInProgress = false;
  }
}

async function recompressAllPhotos(maxLength, quality) {
  let changed = false;
  for (const record of state.records) {
    const sectionResult = await recompressPhotoArray(record.sectionDiagrams, maxLength, quality);
    if (sectionResult.changed) {
      record.sectionDiagrams = sectionResult.photos;
      changed = true;
    }
    const photoResult = await recompressPhotoArray(record.photos, maxLength, quality);
    if (photoResult.changed) {
      record.photos = photoResult.photos;
      changed = true;
    }
  }
  const currentSectionResult = await recompressPhotoArray(currentSectionDiagrams, maxLength, quality);
  if (currentSectionResult.changed) {
    currentSectionDiagrams = currentSectionResult.photos;
    changed = true;
  }
  const currentPhotoResult = await recompressPhotoArray(currentPhotos, maxLength, quality);
  if (currentPhotoResult.changed) {
    currentPhotos = currentPhotoResult.photos;
    changed = true;
  }
  return changed;
}

async function recompressPhotoArray(photosRaw, maxLength, quality) {
  if (!Array.isArray(photosRaw) || !photosRaw.length) {
    return { photos: Array.isArray(photosRaw) ? photosRaw : [], changed: false };
  }
  const nextPhotos = [];
  let changed = false;
  for (const photo of photosRaw) {
    if (!photo || typeof photo !== "object") {
      continue;
    }
    const dataUrl = value(photo.dataUrl);
    if (!dataUrl) {
      continue;
    }
    let nextDataUrl = dataUrl;
    try {
      const recompressed = await resizeDataUrlImage(dataUrl, maxLength, quality);
      if (recompressed && recompressed.length < dataUrl.length) {
        nextDataUrl = recompressed;
        changed = true;
      }
    } catch (_error) {
      // 元画像を維持
    }
    nextPhotos.push({
      ...photo,
      dataUrl: nextDataUrl,
    });
  }
  return { photos: nextPhotos, changed };
}

function resizeDataUrlImage(dataUrl, maxLength, quality = 0.72) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => {
      let width = image.width;
      let height = image.height;
      if (width > maxLength || height > maxLength) {
        const scale = Math.min(maxLength / width, maxLength / height);
        width = Math.max(1, Math.round(width * scale));
        height = Math.max(1, Math.round(height * scale));
      }
      const canvas = document.createElement("canvas");
      canvas.width = width;
      canvas.height = height;
      const context = canvas.getContext("2d");
      if (!context) {
        reject(new Error("Canvas context unavailable"));
        return;
      }
      context.drawImage(image, 0, 0, width, height);
      resolve(canvas.toDataURL("image/jpeg", quality));
    };
    image.onerror = () => reject(new Error("Failed to load image"));
    image.src = dataUrl;
  });
}

function resizeImage(file, maxLength, quality = 0.72) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      resizeDataUrlImage(String(reader.result || ""), maxLength, quality).then(resolve).catch(reject);
    };
    reader.onerror = () => reject(new Error("Failed to read file"));
    reader.readAsDataURL(file);
  });
}
