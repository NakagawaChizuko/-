const STORAGE_KEY = "nojiri-kaseki-mobile-v1";
const CLOUD_ENDPOINT_KEY = "nojiri-kaseki-cloud-endpoint-v1";
const CLOUD_CLIENT_ID_KEY = "nojiri-kaseki-cloud-client-id-v1";
const DEFAULT_CLOUD_ENDPOINT = "https://script.google.com/macros/s/AKfycbw70bPigsRo6HrTzSqUl0--N0Bsno2ybgdfJmtpmmMbQYPqKw-Z9ssOnt5PsGtMT1WSWg/exec";
const CLOUD_PULL_INTERVAL_MS = 20000;
const CLOUD_SAVE_DEBOUNCE_MS = 900;
const CLOUD_AUTO_PULL_ENABLED = false;
const DEFAULT_SPECIMEN_PREFIX = "m";
const SPECIMEN_CATEGORY_MAP = {
  m: "哺乳類",
  b: "植物",
  l: "生痕",
  s: "貝類",
  i: "昆虫",
  g: "人類考古",
  h: "その他",
  a: "分析用試料",
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
const REQUIRED_FIELD_LABELS = {
  kuwakuHeadA: "区画（グリッド）1番目",
  kuwakuHeadB: "区画（グリッド）2番目",
  kuwakuBlock: "区画（グリッド）英字",
  kuwakuNo: "区画（グリッド）番号",
  levelHeight: "レベル高",
  date: "日付",
  team: "発掘班",
  teamOther: "発掘班（その他）",
  teamLead: "班長",
  recorder: "記載係",
  specimenSerial: "標本番号",
  analysisType: "分析用試料の区分",
  nameMemo: "化石・遺物名称",
  importantFlag: "重要品指定",
  simpleRecordFlag: "簡易記載",
  discoverer: "発見者氏名",
  identifier: "判定者氏名",
  levelUpperCm: "レベル読値（上面）",
  levelLowerCm: "レベル読値（下底）",
  occurrenceSection: "産出状況断面",
  occurrenceSketch: "産状スケッチ",
  sectionDiagrams: "産出状況断面図添付",
  sectionDiagramDistanceChecked: "断面図確認: 垂直距離記入",
  sectionDiagramHorizonChecked: "断面図確認: 産出層準記入",
  sectionDiagramLayerFaciesChecked: "断面図確認: 層相記入",
  photoClinometerChecked: "写真確認: クリノメーター",
  photoRulerChecked: "写真確認: 定規",
  nsDir: "平面位置（北から/南から）",
  nsCm: "平面位置（北から/南からの距離）",
  ewDir: "平面位置（東から/西から）",
  ewCm: "平面位置（東から/西からの距離）",
  largeShapeType: "大きなもの形状",
  largeAxisDirection: "長軸・長辺・長半径方向（例:N30W）",
  lineLengthCm: "直線状 長さ",
  rectSide1Cm: "長方形 辺1",
  rectSide2Cm: "長方形 辺2",
  ellipseLongRadiusCm: "楕円 長半径",
  ellipseShortRadiusCm: "楕円 短半径",
  imgP1NsCm: "画像点1 北/南距離",
  imgP1EwCm: "画像点1 東/西距離",
  imgP2NsCm: "画像点2 北/南距離",
  imgP2EwCm: "画像点2 東/西距離",
  imgP3NsCm: "画像点3 北/南距離",
  imgP3EwCm: "画像点3 東/西距離",
  imgP4NsCm: "画像点4 北/南距離",
  imgP4EwCm: "画像点4 東/西距離",
  layerName: "地層名",
  layerOther: "地層名（その他）",
  unit: "ユニット",
  layerRef: "地層中の位置（層理面や鍵層名）",
  layerRelative: "地層中の位置（上/下）",
  layerFromCm: "地層中の位置（cm）",
};
const HISTORY_SNAPSHOT_FIELDS = [
  { key: "specimenNo", label: "標本番号" },
  { key: "nameMemo", label: "名称" },
  { key: "category", label: "分類" },
  { key: "layerName", label: "地層名" },
  { key: "unit", label: "ユニット" },
  { key: "detail", label: "サブユニット" },
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
const ALL_DETAIL_SUBS_VALUE = "__DETAIL_SUB_ALL__";
const EMPTY_DETAIL_SUB_VALUE = "__DETAIL_SUB_EMPTY__";
const EXPORT_PLAN_ALL_UNITS_BUTTON_VALUE = "__EXPORT_PLAN_ALL_UNITS__";
const EXPORT_CATEGORY_ALL_VALUE = "__EXPORT_CATEGORY_ALL__";
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
const IMAGE_SHAPE_CANVAS_DILATE_ITERATIONS = 5;
const LARGE_SHAPE_DIR_PATH = "./shapes";
const LARGE_SHAPE_FILE_LABEL_MAP = {
  palmate_antler: "掌状角",
  incisor: "切歯",
  constricted_shape: "くびれた形",
  rib_curved: "肋骨（湾曲形）",
  triangle: "三角",
  c_shape: "C形",
  diamond_hira: "ひし形",
};
const DEFAULT_LARGE_SHAPE_IMAGE_PATHS = {
  掌状角: `${LARGE_SHAPE_DIR_PATH}/palmate_antler.png`,
  切歯: `${LARGE_SHAPE_DIR_PATH}/incisor.png`,
  くびれた形: `${LARGE_SHAPE_DIR_PATH}/constricted_shape.png`,
  "肋骨（湾曲形）": `${LARGE_SHAPE_DIR_PATH}/rib_curved.png`,
  三角: `${LARGE_SHAPE_DIR_PATH}/triangle.png`,
  C形: `${LARGE_SHAPE_DIR_PATH}/c_shape.png`,
  ひし形: `${LARGE_SHAPE_DIR_PATH}/diamond_hira.png`,
};
const LARGE_SHAPE_IMAGE_FALLBACK_PATHS = {
  掌状角: ["./assets/large-shapes/palmate_antler.png"],
  切歯: ["./assets/large-shapes/incisor.png"],
  三角: ["./assets/large-shapes/triangle.png"],
  C形: ["./assets/large-shapes/c_shape.png"],
  くびれた形: ["./assets/large-shapes/constricted_shape.png"],
  ひし形: ["./assets/large-shapes/diamond_hira.png"],
  "肋骨（湾曲形）": ["./assets/large-shapes/rib_curved.png"],
};
const EXCLUDED_LARGE_SHAPE_LABELS = new Set(["菱形", "肋骨", "肋骨（湾曲型）"]);
const LARGE_SHAPE_MANIFEST_PATH = `${LARGE_SHAPE_DIR_PATH}/manifest.json`;
let largeShapeImagePathMap = new Map(Object.entries(DEFAULT_LARGE_SHAPE_IMAGE_PATHS));
const INLINE_LARGE_SHAPE_DATA_MAP =
  typeof window !== "undefined" &&
  window.__INLINE_LARGE_SHAPE_DATA_MAP__ &&
  typeof window.__INLINE_LARGE_SHAPE_DATA_MAP__ === "object"
    ? window.__INLINE_LARGE_SHAPE_DATA_MAP__
    : {};
const VIEWER_HEAD_SEQUENCE = ["Ⅲ", "Ⅰ", "Ⅱ"];
const VIEWER_HEAD_INDEX_MAP = new Map(VIEWER_HEAD_SEQUENCE.map((head, index) => [head, index]));
const VIEWER_ALTITUDE_BASE_M = 655;
const UNIT_CELL_COLOR_MAP = {
  U1: { background: "hsl(272, 64%, 93%)", border: "hsl(272, 38%, 80%)", color: "#111827" },
  U2: { background: "hsl(286, 62%, 93%)", border: "hsl(286, 36%, 80%)", color: "#111827" },
  U3: { background: "hsl(258, 60%, 93%)", border: "hsl(258, 36%, 80%)", color: "#111827" },
  T1: { background: "hsl(28, 58%, 92%)", border: "hsl(28, 34%, 78%)", color: "#111827" },
  T2: { background: "hsl(34, 56%, 92%)", border: "hsl(34, 34%, 78%)", color: "#111827" },
  T3: { background: "hsl(20, 58%, 92%)", border: "hsl(20, 34%, 78%)", color: "#111827" },
  T4: { background: "hsl(196, 74%, 92%)", border: "hsl(196, 42%, 79%)", color: "#111827" },
  T5: { background: "hsl(52, 84%, 92%)", border: "hsl(52, 46%, 79%)", color: "#111827" },
  T6: { background: "hsl(0, 82%, 93%)", border: "hsl(0, 44%, 80%)", color: "#111827" },
  T7: { background: "hsl(88, 62%, 91%)", border: "hsl(88, 34%, 77%)", color: "#111827" },
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
let selectedPlanDetailSub = ALL_DETAIL_SUBS_VALUE;
let selectedViewerKuwaku = ALL_GRIDS_VALUE;
let selectedViewerUnit = ALL_UNITS_VALUE;
let selectedViewerDetail = ALL_DETAILS_VALUE;
let selectedViewerDetailSub = ALL_DETAIL_SUBS_VALUE;
let selectedViewerPerspective = "top";
let viewerVerticalScale = 1;
let exportListRangeKuwaku = ALL_GRIDS_VALUE;
let exportListRangeCategory = EXPORT_CATEGORY_ALL_VALUE;
let exportListRangeStatus = "all";
let exportListRangeSpecimenFrom = "";
let exportListRangeSpecimenTo = "";
let exportListRangeDateFrom = "";
let exportListRangeDateTo = "";
let exportCardRangeKuwaku = ALL_GRIDS_VALUE;
let exportCardRangeCategory = EXPORT_CATEGORY_ALL_VALUE;
let exportCardRangeStatus = "all";
let exportCardRangeDateFrom = "";
let exportCardRangeDateTo = "";
let exportPlanKuwaku = "";
let exportPlanCategory = EXPORT_CATEGORY_ALL_VALUE;
let exportPlanDateFrom = "";
let exportPlanDateTo = "";
let exportPlanModeUnitEnabled = true;
let exportPlanModeDetailEnabled = false;
let exportPlanModeDetailSubEnabled = false;
let exportPlanModeUnitValues = new Set();
let exportPlanModeDetailUnitValue = "";
let exportPlanModeDetailValues = new Set();
let exportPlanModeDetailSubUnitValue = "";
let exportPlanModeDetailSubDetailValue = "";
let exportPlanModeDetailSubValues = new Set();
let exportPlanModeUnitTouched = false;
let exportPlanModeDetailTouched = false;
let exportPlanModeDetailSubTouched = false;
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
const editTabPanel = document.getElementById("edit-tab");
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
const recordCopyToInputBtn = document.getElementById("record-copy-to-input-btn");
const recordResetBtn = document.getElementById("record-reset-btn");
const recordTableBody = document.getElementById("record-table-body");
const outputListBody = document.getElementById("output-list-body");
const outputListTable = document.getElementById("output-list-table");
const cardOutputList = document.getElementById("card-output-list");
const outputKuwakuSelect = document.getElementById("output-kuwaku-select");
const planKuwakuSelect = document.getElementById("plan-kuwaku-select");
const planUnitSelect = document.getElementById("plan-unit-select");
const planDetailSelect = document.getElementById("plan-detail-select");
const planDetailSubSelect = document.getElementById("plan-detail-sub-select");
const planMapLegend = document.getElementById("plan-map-legend");
const planMapWrap = document.getElementById("plan-map-wrap");
const planKuwakuInfo = document.getElementById("plan-kuwaku-info");
const viewerKuwakuSelect = document.getElementById("viewer-kuwaku-select");
const viewerUnitSelect = document.getElementById("viewer-unit-select");
const viewerDetailSelect = document.getElementById("viewer-detail-select");
const viewerDetailSubSelect = document.getElementById("viewer-detail-sub-select");
const viewerKuwakuInfo = document.getElementById("viewer-kuwaku-info");
const viewerMapLegend = document.getElementById("viewer-map-legend");
const viewerCanvasWrap = document.getElementById("viewer-canvas-wrap");
const viewerTooltip = document.getElementById("viewer-tooltip");
const viewerStatus = document.getElementById("viewer-status");
const viewerViewTopBtn = document.getElementById("viewer-view-top-btn");
const viewerViewIsoBtn = document.getElementById("viewer-view-iso-btn");
const viewerResetBtn = document.getElementById("viewer-reset-btn");
const viewerZScaleInput = document.getElementById("viewer-z-scale-input");
const viewerZScaleValue = document.getElementById("viewer-z-scale-value");
const largeAxisDirectionRow = document.getElementById("large-axis-direction-row");
const exportListRangeKuwakuSelect = document.getElementById("export-range-kuwaku-select");
const exportListRangeCategorySelect = document.getElementById("export-range-category-select");
const exportListRangeStatusSelect = document.getElementById("export-range-status-select");
const exportListRangeSpecimenFromInput = document.getElementById("export-range-specimen-from");
const exportListRangeSpecimenToInput = document.getElementById("export-range-specimen-to");
const exportListRangeDateFromInput = document.getElementById("export-range-date-from");
const exportListRangeDateToInput = document.getElementById("export-range-date-to");
const exportListRangeSummaryEl = document.getElementById("export-range-summary");
const exportCardRangeKuwakuSelect = document.getElementById("export-card-kuwaku-select");
const exportCardRangeCategorySelect = document.getElementById("export-card-category-select");
const exportCardRangeStatusSelect = document.getElementById("export-card-status-select");
const exportCardRangeDateFromInput = document.getElementById("export-card-date-from");
const exportCardRangeDateToInput = document.getElementById("export-card-date-to");
const exportCardRangeSummaryEl = document.getElementById("export-card-summary");
const exportPlanKuwakuSelect = document.getElementById("export-plan-kuwaku-select");
const exportPlanCategorySelect = document.getElementById("export-plan-category-select");
const exportPlanDateFromInput = document.getElementById("export-plan-date-from");
const exportPlanDateToInput = document.getElementById("export-plan-date-to");
const exportPlanModeUnitCheck = document.getElementById("export-plan-mode-unit-check");
const exportPlanModeUnitButtons = document.getElementById("export-plan-mode-unit-buttons");
const exportPlanModeUnitStats = document.getElementById("export-plan-mode-unit-stats");
const exportPlanModeDetailCheck = document.getElementById("export-plan-mode-detail-check");
const exportPlanModeDetailUnitSelect = document.getElementById("export-plan-mode-detail-unit-select");
const exportPlanModeDetailButtons = document.getElementById("export-plan-mode-detail-buttons");
const exportPlanModeDetailStats = document.getElementById("export-plan-mode-detail-stats");
const exportPlanModeDetailSubCheck = document.getElementById("export-plan-mode-detail-sub-check");
const exportPlanModeDetailSubUnitSelect = document.getElementById("export-plan-mode-detail-sub-unit-select");
const exportPlanModeDetailSubDetailSelect = document.getElementById("export-plan-mode-detail-sub-detail-select");
const exportPlanModeDetailSubButtons = document.getElementById("export-plan-mode-detail-sub-buttons");
const exportPlanModeDetailSubStats = document.getElementById("export-plan-mode-detail-sub-stats");
const exportPlanSummaryEl = document.getElementById("export-plan-summary");

const specimenTabButtons = document.querySelectorAll(".specimen-tab");
const specimenPrefixInput = document.getElementById("specimen-prefix-input");
const specimenSerialInput = document.getElementById("specimen-serial-input");
const specimenNoInput = document.getElementById("specimen-no-input");
const specimenPrefixLabel = document.getElementById("specimen-prefix-label");
const specimenDuplicateWarning = document.getElementById("specimen-duplicate-warning");
const analysisTypeRow = document.getElementById("analysis-type-row");
const analysisTypeSelect = document.getElementById("analysis-type-select");

const nsDirInput = document.getElementById("ns-dir-input");
const ewDirInput = document.getElementById("ew-dir-input");
const importantFlagInput = document.getElementById("important-flag-input");
const simpleRecordFlagInput = document.getElementById("simple-record-flag-input");
const occurrenceSectionInput = document.getElementById("occurrence-section-input");
const occurrenceSketchInput = document.getElementById("occurrence-sketch-input");
const layerRelativeInput = document.getElementById("layer-relative-input");
const planSizeModeInput = document.getElementById("plan-size-mode-input");
const largeShapeTypeInput = document.getElementById("large-shape-type-input");
const largeAxisDirectionInput = document.getElementById("large-axis-direction-input");
const largeShapeImageButtons = document.getElementById("large-shape-image-buttons");
const largeShapeImagePreview = document.getElementById("large-shape-image-preview");
const largeShapeImagePreviewTitle = document.getElementById("large-shape-image-preview-title");
const largeShapeImagePreviewImg = document.getElementById("large-shape-image-preview-img");
const line1NsDirInput = document.getElementById("line1-ns-dir-input");
const line1EwDirInput = document.getElementById("line1-ew-dir-input");
const line2NsDirInput = document.getElementById("line2-ns-dir-input");
const line2EwDirInput = document.getElementById("line2-ew-dir-input");
const largeShapeSection = document.getElementById("large-shape-section");
const largeShapePanels = document.querySelectorAll(".large-shape-panel[data-large-shape-panel]");
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
const exportListPdfBtn = document.getElementById("export-list-pdf-btn");
const exportCardPdfBtn = document.getElementById("export-card-pdf-btn");
const exportPlanPdfBtn = document.getElementById("export-plan-pdf-btn");
const exportJsonBtn = document.getElementById("export-json-btn");
const importJsonInput = document.getElementById("import-json-input");
const cloudEndpointInput = document.getElementById("cloud-endpoint-input");
const cloudConnectBtn = document.getElementById("cloud-connect-btn");
const cloudSyncBtn = document.getElementById("cloud-sync-btn");
const cloudDisableBtn = document.getElementById("cloud-disable-btn");
const cloudStatusEl = document.getElementById("cloud-status");
const toastEl = document.getElementById("toast");

const viewer3d = {
  initialized: false,
  available: false,
  scene: null,
  camera: null,
  renderer: null,
  controls: null,
  raycaster: null,
  pointer: null,
  frameHandle: 0,
  resizeObserver: null,
  meshesByRecordId: new Map(),
  pickMeshes: [],
  dataGroup: null,
  labelGroup: null,
  gridGroup: null,
  renderNonce: 0,
};
const planLargeShapeImageCache = new Map();
const planLargeShapeTintedCanvasCache = new Map();
const planLargeShapeTintedDataUrlCache = new Map();

initialize();

function initialize() {
  bindEvents();
  renderLargeShapeImageButtons();
  if (stateNeedsRewriteAfterLoad) {
    persist();
    stateNeedsRewriteAfterLoad = false;
  }
  initCloudControls();
  syncViewerVerticalScaleUi();
  hydrateSiteForm();
  resetRecordForm({ showMessage: false });
  renderRecordTable();
  renderOutputs();
  void loadLargeShapeImageManifest();
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
  if (editTabPanel) {
    editTabPanel.addEventListener("input", () => {
      updateEditMissingRequiredHighlights();
    });
    editTabPanel.addEventListener("change", () => {
      updateEditMissingRequiredHighlights();
    });
    editTabPanel.addEventListener("click", (event) => {
      const target = event.target;
      if (!(target instanceof Element)) {
        return;
      }
      if (!target.closest(".dir-tab, .layer-tab, .specimen-tab, [data-remove-diagram-id], [data-remove-photo-id]")) {
        return;
      }
      updateEditMissingRequiredHighlights();
    });
  }

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

  recordForm.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof Element)) {
      return;
    }
    const button = target.closest(".dir-tab");
    if (!(button instanceof HTMLButtonElement)) {
      return;
    }
    activateDirectionTab(button.dataset.group, button.dataset.value);
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
      showToast("a: 分析用試料を選んだ場合は、区分を選択してください");
      return;
    }
    const attachmentChecklistError = validateAttachmentChecklistForSave(formData);
    if (attachmentChecklistError) {
      showToast(attachmentChecklistError);
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
    const planSizeMode = normalizePlanSizeMode(value(formData.get("planSizeMode")));
    const rawLargeShapeType = value(formData.get("largeShapeType"));
    const largeShapeType =
      planSizeMode === "大きなもの"
        ? normalizeLargeShapeType(rawLargeShapeType) || normalizeLargeShapeLabel(rawLargeShapeType)
        : "";
    const largeAxisDirection = planSizeMode === "大きなもの" ? normalizeLargeAxisDirection(value(formData.get("largeAxisDirection"))) : "";
    const lineLengthCm = value(formData.get("lineLengthCm"));
    const rectSide1Cm = value(formData.get("rectSide1Cm"));
    const rectSide2Cm = value(formData.get("rectSide2Cm"));
    const ellipseLongRadiusCm = value(formData.get("ellipseLongRadiusCm"));
    const ellipseShortRadiusCm = value(formData.get("ellipseShortRadiusCm"));
    const imageCornerFields = extractImageCornerFieldsFromFormData(formData);
    const isLargeImageShape = planSizeMode === "大きなもの" && isLargeShapeImageType(largeShapeType);
    const keepImageCornerFields = planSizeMode === "大きなもの";

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
      sectionDiagramDistanceChecked: normalizeChecklistChecked(formData.get("sectionDiagramDistanceChecked")),
      sectionDiagramHorizonChecked: normalizeChecklistChecked(formData.get("sectionDiagramHorizonChecked")),
      sectionDiagramLayerFaciesChecked: normalizeChecklistChecked(formData.get("sectionDiagramLayerFaciesChecked")),
      photoClinometerChecked: normalizeChecklistChecked(formData.get("photoClinometerChecked")),
      photoRulerChecked: normalizeChecklistChecked(formData.get("photoRulerChecked")),
      nsDir: normalizeNsDir(value(formData.get("nsDir"))),
      nsCm: value(formData.get("nsCm")),
      ewDir: normalizeEwDir(value(formData.get("ewDir"))),
      ewCm: value(formData.get("ewCm")),
      planSizeMode,
      largeShapeType,
      largeAxisDirection: isLargeImageShape ? "" : largeAxisDirection,
      lineLengthCm: planSizeMode === "大きなもの" && largeShapeType === "直線状" ? lineLengthCm : "",
      line1NsDir: "",
      line1NsCm: "",
      line1EwDir: "",
      line1EwCm: "",
      line2NsDir: "",
      line2NsCm: "",
      line2EwDir: "",
      line2EwCm: "",
      rectSide1Cm: planSizeMode === "大きなもの" && largeShapeType === "長方形" ? rectSide1Cm : "",
      rectSide2Cm: planSizeMode === "大きなもの" && largeShapeType === "長方形" ? rectSide2Cm : "",
      ellipseLongRadiusCm: planSizeMode === "大きなもの" && largeShapeType === "楕円" ? ellipseLongRadiusCm : "",
      ellipseShortRadiusCm: planSizeMode === "大きなもの" && largeShapeType === "楕円" ? ellipseShortRadiusCm : "",
      imgP1NsDir: keepImageCornerFields ? imageCornerFields.imgP1NsDir : "",
      imgP1NsCm: keepImageCornerFields ? imageCornerFields.imgP1NsCm : "",
      imgP1EwDir: keepImageCornerFields ? imageCornerFields.imgP1EwDir : "",
      imgP1EwCm: keepImageCornerFields ? imageCornerFields.imgP1EwCm : "",
      imgP2NsDir: keepImageCornerFields ? imageCornerFields.imgP2NsDir : "",
      imgP2NsCm: keepImageCornerFields ? imageCornerFields.imgP2NsCm : "",
      imgP2EwDir: keepImageCornerFields ? imageCornerFields.imgP2EwDir : "",
      imgP2EwCm: keepImageCornerFields ? imageCornerFields.imgP2EwCm : "",
      imgP3NsDir: keepImageCornerFields ? imageCornerFields.imgP3NsDir : "",
      imgP3NsCm: keepImageCornerFields ? imageCornerFields.imgP3NsCm : "",
      imgP3EwDir: keepImageCornerFields ? imageCornerFields.imgP3EwDir : "",
      imgP3EwCm: keepImageCornerFields ? imageCornerFields.imgP3EwCm : "",
      imgP4NsDir: keepImageCornerFields ? imageCornerFields.imgP4NsDir : "",
      imgP4NsCm: keepImageCornerFields ? imageCornerFields.imgP4NsCm : "",
      imgP4EwDir: keepImageCornerFields ? imageCornerFields.imgP4EwDir : "",
      imgP4EwCm: keepImageCornerFields ? imageCornerFields.imgP4EwCm : "",
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
    const missingRequiredKeys = getMissingRequiredKeys(record);
    if (missingRequiredKeys.size > 0) {
      const keepSavingIncomplete = window.confirm("未記入がありますが保存しますか？");
      if (!keepSavingIncomplete) {
        return;
      }
    }

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
      updateEditMissingRequiredHighlights();
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
  if (recordCopyToInputBtn) {
    recordCopyToInputBtn.addEventListener("click", () => {
      copyCurrentEditToInput();
    });
  }

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
    if (action === "copy-to-input") {
      const rowKuwaku = value(button.dataset.kuwaku);
      copySavedRecordToInput(recordId, rowKuwaku);
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
    if (action === "copy-to-input") {
      const rowKuwaku = value(button.dataset.kuwaku);
      copySavedRecordToInput(recordId, rowKuwaku);
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

  if (exportListRangeKuwakuSelect) {
    exportListRangeKuwakuSelect.addEventListener("change", () => {
      exportListRangeKuwaku = value(exportListRangeKuwakuSelect.value) || ALL_GRIDS_VALUE;
      renderExportOutput();
    });
  }
  if (exportListRangeCategorySelect) {
    exportListRangeCategorySelect.addEventListener("change", () => {
      exportListRangeCategory = value(exportListRangeCategorySelect.value) || EXPORT_CATEGORY_ALL_VALUE;
      renderExportOutput();
    });
  }
  if (exportListRangeStatusSelect) {
    exportListRangeStatusSelect.addEventListener("change", () => {
      exportListRangeStatus = value(exportListRangeStatusSelect.value) || "all";
      renderExportOutput();
    });
  }
  if (exportListRangeSpecimenFromInput) {
    exportListRangeSpecimenFromInput.addEventListener("input", () => {
      exportListRangeSpecimenFrom = value(exportListRangeSpecimenFromInput.value);
      renderExportOutput();
    });
  }
  if (exportListRangeSpecimenToInput) {
    exportListRangeSpecimenToInput.addEventListener("input", () => {
      exportListRangeSpecimenTo = value(exportListRangeSpecimenToInput.value);
      renderExportOutput();
    });
  }
  if (exportListRangeDateFromInput) {
    exportListRangeDateFromInput.addEventListener("input", () => {
      exportListRangeDateFrom = value(exportListRangeDateFromInput.value);
      renderExportOutput();
    });
  }
  if (exportListRangeDateToInput) {
    exportListRangeDateToInput.addEventListener("input", () => {
      exportListRangeDateTo = value(exportListRangeDateToInput.value);
      renderExportOutput();
    });
  }
  if (exportCardRangeKuwakuSelect) {
    exportCardRangeKuwakuSelect.addEventListener("change", () => {
      exportCardRangeKuwaku = value(exportCardRangeKuwakuSelect.value) || ALL_GRIDS_VALUE;
      renderExportOutput();
    });
  }
  if (exportCardRangeCategorySelect) {
    exportCardRangeCategorySelect.addEventListener("change", () => {
      exportCardRangeCategory = value(exportCardRangeCategorySelect.value) || EXPORT_CATEGORY_ALL_VALUE;
      renderExportOutput();
    });
  }
  if (exportCardRangeStatusSelect) {
    exportCardRangeStatusSelect.addEventListener("change", () => {
      exportCardRangeStatus = value(exportCardRangeStatusSelect.value) || "all";
      renderExportOutput();
    });
  }
  if (exportCardRangeDateFromInput) {
    exportCardRangeDateFromInput.addEventListener("input", () => {
      exportCardRangeDateFrom = value(exportCardRangeDateFromInput.value);
      renderExportOutput();
    });
  }
  if (exportCardRangeDateToInput) {
    exportCardRangeDateToInput.addEventListener("input", () => {
      exportCardRangeDateTo = value(exportCardRangeDateToInput.value);
      renderExportOutput();
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
  if (planDetailSubSelect) {
    planDetailSubSelect.addEventListener("change", () => {
      selectedPlanDetailSub = value(planDetailSubSelect.value) || ALL_DETAIL_SUBS_VALUE;
      renderPlanOutput();
    });
  }

  if (viewerKuwakuSelect) {
    viewerKuwakuSelect.addEventListener("change", () => {
      selectedViewerKuwaku = value(viewerKuwakuSelect.value) || ALL_GRIDS_VALUE;
      renderViewerOutput();
    });
  }
  if (viewerUnitSelect) {
    viewerUnitSelect.addEventListener("change", () => {
      selectedViewerUnit = value(viewerUnitSelect.value) || ALL_UNITS_VALUE;
      renderViewerOutput();
    });
  }
  if (viewerDetailSelect) {
    viewerDetailSelect.addEventListener("change", () => {
      selectedViewerDetail = value(viewerDetailSelect.value) || ALL_DETAILS_VALUE;
      renderViewerOutput();
    });
  }
  if (viewerDetailSubSelect) {
    viewerDetailSubSelect.addEventListener("change", () => {
      selectedViewerDetailSub = value(viewerDetailSubSelect.value) || ALL_DETAIL_SUBS_VALUE;
      renderViewerOutput();
    });
  }
  if (viewerViewTopBtn) {
    viewerViewTopBtn.addEventListener("click", () => {
      selectedViewerPerspective = "top";
      applyViewerPerspective();
      syncViewerViewButtons();
    });
  }
  if (viewerViewIsoBtn) {
    viewerViewIsoBtn.addEventListener("click", () => {
      selectedViewerPerspective = "iso";
      applyViewerPerspective();
      syncViewerViewButtons();
    });
  }
  if (viewerResetBtn) {
    viewerResetBtn.addEventListener("click", () => {
      renderViewerOutput();
    });
  }
  if (viewerZScaleInput) {
    viewerZScaleInput.addEventListener("input", () => {
      viewerVerticalScale = normalizeViewerVerticalScale(viewerZScaleInput.value);
      syncViewerVerticalScaleUi();
      if (viewer3d.initialized) {
        renderViewerOutput();
      }
    });
  }

  if (exportPlanKuwakuSelect) {
    exportPlanKuwakuSelect.addEventListener("change", () => {
      exportPlanKuwaku = value(exportPlanKuwakuSelect.value);
      renderExportOutput();
    });
  }
  if (exportPlanCategorySelect) {
    exportPlanCategorySelect.addEventListener("change", () => {
      exportPlanCategory = value(exportPlanCategorySelect.value) || EXPORT_CATEGORY_ALL_VALUE;
      renderExportOutput();
    });
  }
  if (exportPlanDateFromInput) {
    exportPlanDateFromInput.addEventListener("input", () => {
      exportPlanDateFrom = value(exportPlanDateFromInput.value);
      renderExportOutput();
    });
  }
  if (exportPlanDateToInput) {
    exportPlanDateToInput.addEventListener("input", () => {
      exportPlanDateTo = value(exportPlanDateToInput.value);
      renderExportOutput();
    });
  }
  if (exportPlanModeUnitCheck) {
    exportPlanModeUnitCheck.addEventListener("change", () => {
      exportPlanModeUnitEnabled = Boolean(exportPlanModeUnitCheck.checked);
      renderExportOutput();
    });
  }
  if (exportPlanModeUnitButtons) {
    exportPlanModeUnitButtons.addEventListener("click", (event) => {
      const button = event.target.closest("button[data-value]");
      if (!(button instanceof HTMLButtonElement)) {
        return;
      }
      const optionValue = value(button.dataset.value);
      exportPlanModeUnitTouched = true;
      if (optionValue === EXPORT_PLAN_ALL_UNITS_BUTTON_VALUE) {
        const unitValues = Array.from(
          collectExportPlanValueOptions(getExportPlanScopedRecords(), (record) => unitValueForSelect(record.unit), unitLabelForSelect)
        )
          .map((item) => value(item.value))
          .filter(Boolean);
        const allSelected =
          unitValues.length > 0 && unitValues.every((unitValue) => exportPlanModeUnitValues.has(unitValue));
        exportPlanModeUnitValues = allSelected ? new Set() : new Set(unitValues);
      } else {
        toggleSelectionInSet(exportPlanModeUnitValues, optionValue);
      }
      renderExportOutput();
    });
  }
  if (exportPlanModeDetailCheck) {
    exportPlanModeDetailCheck.addEventListener("change", () => {
      exportPlanModeDetailEnabled = Boolean(exportPlanModeDetailCheck.checked);
      renderExportOutput();
    });
  }
  if (exportPlanModeDetailUnitSelect) {
    exportPlanModeDetailUnitSelect.addEventListener("change", () => {
      exportPlanModeDetailUnitValue = value(exportPlanModeDetailUnitSelect.value);
      exportPlanModeDetailTouched = false;
      renderExportOutput();
    });
  }
  if (exportPlanModeDetailButtons) {
    exportPlanModeDetailButtons.addEventListener("click", (event) => {
      const button = event.target.closest("button[data-value]");
      if (!(button instanceof HTMLButtonElement)) {
        return;
      }
      exportPlanModeDetailTouched = true;
      toggleSelectionInSet(exportPlanModeDetailValues, value(button.dataset.value));
      renderExportOutput();
    });
  }
  if (exportPlanModeDetailSubCheck) {
    exportPlanModeDetailSubCheck.addEventListener("change", () => {
      exportPlanModeDetailSubEnabled = Boolean(exportPlanModeDetailSubCheck.checked);
      renderExportOutput();
    });
  }
  if (exportPlanModeDetailSubUnitSelect) {
    exportPlanModeDetailSubUnitSelect.addEventListener("change", () => {
      exportPlanModeDetailSubUnitValue = value(exportPlanModeDetailSubUnitSelect.value);
      exportPlanModeDetailSubTouched = false;
      renderExportOutput();
    });
  }
  if (exportPlanModeDetailSubDetailSelect) {
    exportPlanModeDetailSubDetailSelect.addEventListener("change", () => {
      exportPlanModeDetailSubDetailValue = value(exportPlanModeDetailSubDetailSelect.value);
      exportPlanModeDetailSubTouched = false;
      renderExportOutput();
    });
  }
  if (exportPlanModeDetailSubButtons) {
    exportPlanModeDetailSubButtons.addEventListener("click", (event) => {
      const button = event.target.closest("button[data-value]");
      if (!(button instanceof HTMLButtonElement)) {
        return;
      }
      exportPlanModeDetailSubTouched = true;
      toggleSelectionInSet(exportPlanModeDetailSubValues, value(button.dataset.value));
      renderExportOutput();
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
    if (!getListExportRecords().length) {
      showToast("CSV出力対象データがありません");
      return;
    }
    const csv = buildListCsv();
    downloadFile(`nojiri-kaseki-list-${timestamp()}.csv`, csv, "text/csv;charset=utf-8");
    showToast("リストCSVを書き出しました");
  });

  exportCardCsvBtn.addEventListener("click", () => {
    if (!getCardExportRecords().length) {
      showToast("カードCSVの出力対象データがありません");
      return;
    }
    const csv = buildCardCsv();
    downloadFile(`nojiri-kaseki-card-${timestamp()}.csv`, csv, "text/csv;charset=utf-8");
    showToast("カードCSVを書き出しました");
  });

  if (exportListPdfBtn) {
    exportListPdfBtn.addEventListener("click", () => {
      exportListPdf();
    });
  }

  if (exportCardPdfBtn) {
    exportCardPdfBtn.addEventListener("click", () => {
      exportCardPdf();
    });
  }

  if (exportPlanPdfBtn) {
    exportPlanPdfBtn.addEventListener("click", () => {
      exportPlanPdf();
    });
  }

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
  updateEditMissingRequiredHighlights();
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
  if (tabId === "edit-tab") {
    updateEditMissingRequiredHighlights();
  } else {
    clearEditMissingRequiredHighlights();
  }
  if (tabId === "viewer-tab") {
    renderViewerOutput();
    ensureViewerCanvasSize();
  }
  if (CLOUD_AUTO_PULL_ENABLED && cloudEndpoint && (tabId === "output-tab" || tabId === "plan-tab" || tabId === "viewer-tab" || tabId === "export-tab")) {
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
    if (recordCopyToInputBtn) {
      recordCopyToInputBtn.classList.remove("hidden");
    }
    return;
  }

  if (recordForm.parentElement !== recordFormHost) {
    recordFormHost.appendChild(recordForm);
  }
  if (recordCopyToInputBtn) {
    recordCopyToInputBtn.classList.add("hidden");
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
  specimenPrefixLabel.dataset.prefix = prefix;
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
  setDirectionGroupValue(group, valueRaw);
  syncDirectionTabsFromForm();
}

function syncDirectionTabsFromForm() {
  setDirectionGroupValue("ns", nsDirInput?.value);
  setDirectionGroupValue("ew", ewDirInput?.value);
  setDirectionGroupValue("line1Ns", line1NsDirInput?.value);
  setDirectionGroupValue("line1Ew", line1EwDirInput?.value);
  setDirectionGroupValue("line2Ns", line2NsDirInput?.value);
  setDirectionGroupValue("line2Ew", line2EwDirInput?.value);
  setDirectionGroupValue("importantFlag", importantFlagInput?.value);
  setDirectionGroupValue("simpleRecordFlag", simpleRecordFlagInput?.value);
  setDirectionGroupValue("occurrenceSection", occurrenceSectionInput?.value);
  setDirectionGroupValue("occurrenceSketch", occurrenceSketchInput?.value);
  setDirectionGroupValue("layerRelative", layerRelativeInput?.value);
  setDirectionGroupValue("planSizeMode", planSizeModeInput?.value);
  setDirectionGroupValue("largeShapeType", largeShapeTypeInput?.value);
  syncLargeShapeSectionFromForm();

  document.querySelectorAll(".dir-tab").forEach((button) => {
    const group = value(button.dataset.group);
    const selected = getDirectionGroupValue(group);
    button.classList.toggle("active", normalizeDirectionValue(group, button.dataset.value) === selected);
  });
}

function setDirectionGroupValue(group, valueRaw) {
  const normalized = normalizeDirectionValue(group, valueRaw);
  if (group === "ns" && nsDirInput) {
    nsDirInput.value = normalized;
    return;
  }
  if (group === "ew" && ewDirInput) {
    ewDirInput.value = normalized;
    return;
  }
  if (group === "line1Ns" && line1NsDirInput) {
    line1NsDirInput.value = normalized;
    return;
  }
  if (group === "line1Ew" && line1EwDirInput) {
    line1EwDirInput.value = normalized;
    return;
  }
  if (group === "line2Ns" && line2NsDirInput) {
    line2NsDirInput.value = normalized;
    return;
  }
  if (group === "line2Ew" && line2EwDirInput) {
    line2EwDirInput.value = normalized;
    return;
  }
  if (group === "importantFlag" && importantFlagInput) {
    importantFlagInput.value = normalized;
    return;
  }
  if (group === "simpleRecordFlag" && simpleRecordFlagInput) {
    simpleRecordFlagInput.value = normalized;
    return;
  }
  if (group === "occurrenceSection" && occurrenceSectionInput) {
    occurrenceSectionInput.value = normalized;
    return;
  }
  if (group === "occurrenceSketch" && occurrenceSketchInput) {
    occurrenceSketchInput.value = normalized;
    return;
  }
  if (group === "layerRelative" && layerRelativeInput) {
    layerRelativeInput.value = normalized;
    return;
  }
  if (group === "planSizeMode" && planSizeModeInput) {
    planSizeModeInput.value = normalized;
    return;
  }
  if (group === "largeShapeType" && largeShapeTypeInput) {
    largeShapeTypeInput.value = normalized;
  }
}

function getDirectionGroupValue(group) {
  if (group === "ns") {
    return normalizeDirectionValue(group, nsDirInput?.value);
  }
  if (group === "ew") {
    return normalizeDirectionValue(group, ewDirInput?.value);
  }
  if (group === "line1Ns") {
    return normalizeDirectionValue(group, line1NsDirInput?.value);
  }
  if (group === "line1Ew") {
    return normalizeDirectionValue(group, line1EwDirInput?.value);
  }
  if (group === "line2Ns") {
    return normalizeDirectionValue(group, line2NsDirInput?.value);
  }
  if (group === "line2Ew") {
    return normalizeDirectionValue(group, line2EwDirInput?.value);
  }
  if (group === "importantFlag") {
    return normalizeDirectionValue(group, importantFlagInput?.value);
  }
  if (group === "simpleRecordFlag") {
    return normalizeDirectionValue(group, simpleRecordFlagInput?.value);
  }
  if (group === "occurrenceSection") {
    return normalizeDirectionValue(group, occurrenceSectionInput?.value);
  }
  if (group === "occurrenceSketch") {
    return normalizeDirectionValue(group, occurrenceSketchInput?.value);
  }
  if (group === "layerRelative") {
    return normalizeDirectionValue(group, layerRelativeInput?.value);
  }
  if (group === "planSizeMode") {
    return normalizeDirectionValue(group, planSizeModeInput?.value);
  }
  if (group === "largeShapeType") {
    return normalizeDirectionValue(group, largeShapeTypeInput?.value);
  }
  return "";
}

function deriveShapeLabelFromFileName(fileNameRaw) {
  const fileName = value(fileNameRaw);
  if (!fileName) {
    return "";
  }
  const baseName = fileName.split("/").pop() || fileName;
  const normalized = typeof baseName.normalize === "function" ? baseName.normalize("NFC") : baseName;
  const stem = normalized.replace(/\.[^.]+$/, "");
  const mapped = LARGE_SHAPE_FILE_LABEL_MAP[stem] || LARGE_SHAPE_FILE_LABEL_MAP[stem.toLowerCase()] || stem;
  return normalizeLargeShapeLabel(mapped);
}

function normalizeLargeShapeLabel(labelRaw) {
  const raw = value(labelRaw);
  const normalized = typeof raw.normalize === "function" ? raw.normalize("NFC") : raw;
  const withoutExt = normalized.replace(/\.[^.]+$/, "");
  const compact = withoutExt.replace(/\s+/g, "");
  const compactLower = compact.toLowerCase();
  if (LARGE_SHAPE_FILE_LABEL_MAP[compact]) {
    return LARGE_SHAPE_FILE_LABEL_MAP[compact];
  }
  if (LARGE_SHAPE_FILE_LABEL_MAP[compactLower]) {
    return LARGE_SHAPE_FILE_LABEL_MAP[compactLower];
  }
  if (compact === "くびれた骨") {
    return "くびれた形";
  }
  if (
    compact === "肋骨" ||
    compact === "肋骨（湾曲型）" ||
    compact === "肋骨（湾曲形）" ||
    compact === "肋骨(湾曲型)" ||
    compact === "肋骨(湾曲形)"
  ) {
    return "肋骨（湾曲形）";
  }
  if (compact === "C型") {
    return "C形";
  }
  if (compact === "菱形") {
    return "ひし形";
  }
  return withoutExt;
}

function toSafeAssetUrl(pathRaw) {
  const path = pathRaw == null ? "" : String(pathRaw).trim();
  if (!path) {
    return "";
  }
  if (path.startsWith("data:")) {
    return path;
  }
  try {
    return encodeURI(path);
  } catch (_error) {
    return path;
  }
}

function normalizeAssetPathKey(pathRaw) {
  const raw = pathRaw == null ? "" : String(pathRaw).trim();
  if (!raw) {
    return "";
  }
  if (raw.startsWith("data:")) {
    return raw;
  }
  let normalized = raw.replace(/[#?].*$/, "");
  try {
    normalized = decodeURI(normalized);
  } catch (_error) {
    // ignore decode failures and keep the original string
  }
  normalized = normalized.replace(/\\/g, "/").replace(/^(\.\/)+/, "").replace(/^\/+/, "");
  const repoMarker = "kaseki_mobile_app/";
  const markerIndex = normalized.lastIndexOf(repoMarker);
  if (markerIndex >= 0) {
    normalized = normalized.slice(markerIndex + repoMarker.length);
  }
  return normalized;
}

function getInlineLargeShapeDataUrl(pathRaw) {
  const key = normalizeAssetPathKey(pathRaw);
  if (!key || key.startsWith("data:")) {
    return "";
  }
  return value(INLINE_LARGE_SHAPE_DATA_MAP[key]);
}

function renderLargeShapeImageButtons() {
  if (!largeShapeImageButtons) {
    return;
  }
  const labels = Array.from(largeShapeImagePathMap.keys()).filter((label) => value(label));
  largeShapeImageButtons.innerHTML = labels
    .map(
      (label) =>
        `<button class="dir-tab" data-group="largeShapeType" data-value="${escapeHtml(label)}" type="button">${escapeHtml(label)}</button>`
    )
    .join("");
}

async function loadLargeShapeImageManifest() {
  try {
    const response = await fetch(`${LARGE_SHAPE_MANIFEST_PATH}?v=${Date.now()}`, { cache: "no-store" });
    if (!response.ok) {
      return;
    }
    const manifest = await response.json();
    const images = Array.isArray(manifest?.images) ? manifest.images : [];
    if (!images.length) {
      return;
    }
    const nextMap = new Map(Object.entries(DEFAULT_LARGE_SHAPE_IMAGE_PATHS));
    images.forEach((item) => {
      if (typeof item === "string") {
        const fileName = value(item);
        if (!fileName) {
          return;
        }
        const label = deriveShapeLabelFromFileName(fileName);
        if (!label || EXCLUDED_LARGE_SHAPE_LABELS.has(label)) {
          return;
        }
        if (!nextMap.has(label)) {
          nextMap.set(label, toSafeAssetUrl(`${LARGE_SHAPE_DIR_PATH}/${fileName}`));
        }
        return;
      }
      if (item && typeof item === "object") {
        const fileName = value(item.file || item.path || item.src);
        const explicitLabel = value(item.label);
        const label = explicitLabel ? normalizeLargeShapeLabel(explicitLabel) : deriveShapeLabelFromFileName(fileName);
        if (!fileName || !label || EXCLUDED_LARGE_SHAPE_LABELS.has(label)) {
          return;
        }
        const path = /^(\/|\.\/|https?:)/.test(fileName) ? fileName : `${LARGE_SHAPE_DIR_PATH}/${fileName}`;
        nextMap.set(label, toSafeAssetUrl(path));
      }
    });
    if (!nextMap.size) {
      return;
    }
    largeShapeImagePathMap = nextMap;
    planLargeShapeImageCache.clear();
    planLargeShapeTintedCanvasCache.clear();
    planLargeShapeTintedDataUrlCache.clear();
    renderLargeShapeImageButtons();
    syncDirectionTabsFromForm();
    renderOutputs();
  } catch (_error) {
    // ローカル file:// 実行時など manifest を読めない場合は既定画像のみで継続。
  }
}

function syncLargeShapeImagePreview(shapeTypeRaw) {
  if (!largeShapeImagePreview || !largeShapeImagePreviewTitle || !largeShapeImagePreviewImg) {
    return;
  }
  const shapeType = normalizeLargeShapeType(shapeTypeRaw);
  const candidates = getLargeShapeImagePathCandidates(shapeType);
  if (!candidates.length) {
    largeShapeImagePreview.classList.add("hidden");
    largeShapeImagePreviewTitle.textContent = "";
    largeShapeImagePreviewImg.removeAttribute("src");
    largeShapeImagePreviewImg.alt = "";
    return;
  }
  largeShapeImagePreview.classList.remove("hidden");
  largeShapeImagePreviewTitle.textContent = `${shapeType}`;
  largeShapeImagePreviewImg.alt = `${shapeType} 画像`;
  let candidateIndex = 0;
  largeShapeImagePreviewImg.onerror = () => {
    candidateIndex += 1;
    if (candidateIndex >= candidates.length) {
      largeShapeImagePreview.classList.add("hidden");
      return;
    }
    largeShapeImagePreviewImg.src = candidates[candidateIndex];
  };
  largeShapeImagePreviewImg.src = candidates[candidateIndex];
}

function syncLargeShapeSectionFromForm() {
  if (!largeShapeSection) {
    return;
  }
  const mode = normalizePlanSizeMode(planSizeModeInput?.value);
  if (planSizeModeInput) {
    planSizeModeInput.value = mode;
  }
  const isLarge = mode === "大きなもの";
  largeShapeSection.classList.toggle("hidden", !isLarge);
  if (!isLarge) {
    if (largeShapeTypeInput) {
      largeShapeTypeInput.value = "";
    }
    if (largeAxisDirectionInput) {
      largeAxisDirectionInput.value = "";
    }
    largeShapePanels.forEach((panel) => {
      panel.classList.add("hidden");
    });
    if (largeAxisDirectionRow) {
      largeAxisDirectionRow.classList.remove("hidden");
    }
    syncLargeShapeImagePreview("");
    return;
  }

  const shapeTypeRaw = normalizeLargeShapeType(largeShapeTypeInput?.value);
  const shapeType = shapeTypeRaw || "直線状";
  const isImageShape = isLargeShapeImageType(shapeType);
  if (largeShapeTypeInput) {
    largeShapeTypeInput.value = shapeType;
  }
  if (largeAxisDirectionInput) {
    largeAxisDirectionInput.value = isImageShape ? "" : normalizeLargeAxisDirection(largeAxisDirectionInput.value);
  }
  if (largeAxisDirectionRow) {
    largeAxisDirectionRow.classList.toggle("hidden", isImageShape);
  }
  largeShapePanels.forEach((panel) => {
    const panelType = value(panel.dataset.largeShapePanel);
    const shouldShow = panelType === shapeType || (panelType === "__IMAGE__" && isImageShape);
    panel.classList.toggle("hidden", !shouldShow);
  });
  syncLargeShapeImagePreview(isImageShape ? shapeType : "");
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
  syncDirectionTabsFromForm();
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
    document.querySelectorAll(".dir-tab").forEach((button) => {
      if (button.dataset.group === "ns") {
        button.classList.add("overwrite-updated");
      }
    });
  }
  if (normalizeEwDir(previousRecord.ewDir) !== normalizeEwDir(nextRecord.ewDir)) {
    document.querySelectorAll(".dir-tab").forEach((button) => {
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
    updateEditMissingRequiredHighlights();
    return;
  }

  if (target === layerOtherInput) {
    layerOtherInput.classList.remove("saved-carry-value");
    clearLayerSavedTabState();
  }
  updateEditMissingRequiredHighlights();
}

function setDefaultImageCornerDirections() {
  const defaults = {
    imgP1NsDir: "北から",
    imgP1EwDir: "東から",
    imgP2NsDir: "北から",
    imgP2EwDir: "東から",
    imgP3NsDir: "北から",
    imgP3EwDir: "東から",
    imgP4NsDir: "北から",
    imgP4EwDir: "東から",
  };
  Object.entries(defaults).forEach(([name, defaultValue]) => {
    const field = recordForm?.elements?.namedItem(name);
    if (field instanceof HTMLSelectElement || field instanceof HTMLInputElement) {
      field.value = defaultValue;
    }
  });
}

function clearImageCornerCmFields() {
  const names = ["imgP1NsCm", "imgP1EwCm", "imgP2NsCm", "imgP2EwCm", "imgP3NsCm", "imgP3EwCm", "imgP4NsCm", "imgP4EwCm"];
  names.forEach((name) => {
    const field = recordForm?.elements?.namedItem(name);
    if (field instanceof HTMLInputElement) {
      field.value = "";
    }
  });
}

function extractImageCornerFieldsFromFormData(formData) {
  const getNsDir = (name) => normalizeNsDir(value(formData.get(name)));
  const getEwDir = (name) => normalizeEwDir(value(formData.get(name)));
  const getCm = (name) => value(formData.get(name));

  return {
    imgP1NsDir: getNsDir("imgP1NsDir"),
    imgP1NsCm: getCm("imgP1NsCm"),
    imgP1EwDir: getEwDir("imgP1EwDir"),
    imgP1EwCm: getCm("imgP1EwCm"),
    imgP2NsDir: getNsDir("imgP2NsDir"),
    imgP2NsCm: getCm("imgP2NsCm"),
    imgP2EwDir: getEwDir("imgP2EwDir"),
    imgP2EwCm: getCm("imgP2EwCm"),
    imgP3NsDir: getNsDir("imgP3NsDir"),
    imgP3NsCm: getCm("imgP3NsCm"),
    imgP3EwDir: getEwDir("imgP3EwDir"),
    imgP3EwCm: getCm("imgP3EwCm"),
    imgP4NsDir: getNsDir("imgP4NsDir"),
    imgP4NsCm: getCm("imgP4NsCm"),
    imgP4EwDir: getEwDir("imgP4EwDir"),
    imgP4EwCm: getCm("imgP4EwCm"),
  };
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
  recordForm.elements.importantFlag.value = "無";
  recordForm.elements.simpleRecordFlag.value = "-";
  recordForm.elements.layerRelative.value = "";
  if (recordForm.elements.planSizeMode) {
    recordForm.elements.planSizeMode.value = "通常";
  }
  if (recordForm.elements.largeShapeType) {
    recordForm.elements.largeShapeType.value = "";
  }
  if (recordForm.elements.largeAxisDirection) {
    recordForm.elements.largeAxisDirection.value = "";
  }
  if (recordForm.elements.lineLengthCm) {
    recordForm.elements.lineLengthCm.value = "";
  }
  clearImageCornerCmFields();
  setDefaultImageCornerDirections();
  setLayerFromValue(PRESET_LAYER_NAMES[0]);

  nsDirInput.value = "北から";
  ewDirInput.value = "東から";
  syncDirectionTabsFromForm();

  currentPhotos = [];
  currentSectionDiagrams = [];
  renderSectionDiagramList();
  renderPhotoList();
  clearEditHistory();
  clearEditMissingRequiredHighlights();

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
  if (recordForm.elements.sectionDiagramDistanceChecked) {
    recordForm.elements.sectionDiagramDistanceChecked.checked = normalizeChecklistChecked(record.sectionDiagramDistanceChecked) === "1";
  }
  if (recordForm.elements.sectionDiagramHorizonChecked) {
    recordForm.elements.sectionDiagramHorizonChecked.checked = normalizeChecklistChecked(record.sectionDiagramHorizonChecked) === "1";
  }
  if (recordForm.elements.sectionDiagramLayerFaciesChecked) {
    recordForm.elements.sectionDiagramLayerFaciesChecked.checked =
      normalizeChecklistChecked(record.sectionDiagramLayerFaciesChecked) === "1";
  }
  if (recordForm.elements.photoClinometerChecked) {
    recordForm.elements.photoClinometerChecked.checked = normalizeChecklistChecked(record.photoClinometerChecked) === "1";
  }
  if (recordForm.elements.photoRulerChecked) {
    recordForm.elements.photoRulerChecked.checked = normalizeChecklistChecked(record.photoRulerChecked) === "1";
  }
  recordForm.elements.importantFlag.value = normalizeHasFlag(record.importantFlag);
  recordForm.elements.simpleRecordFlag.value = normalizeCircleDashFlag(record.simpleRecordFlag);
  recordForm.elements.layerRelative.value = normalizeLayerRelative(record.layerRelative);
  if (recordForm.elements.planSizeMode) {
    recordForm.elements.planSizeMode.value = normalizePlanSizeMode(record.planSizeMode);
  }
  if (recordForm.elements.largeShapeType) {
    recordForm.elements.largeShapeType.value = normalizeLargeShapeType(record.largeShapeType);
  }
  if (recordForm.elements.largeAxisDirection) {
    recordForm.elements.largeAxisDirection.value = normalizeLargeAxisDirection(record.largeAxisDirection);
  }
  if (recordForm.elements.lineLengthCm) {
    recordForm.elements.lineLengthCm.value = value(record.lineLengthCm);
  }

  nsDirInput.value = normalizeNsDir(record.nsDir);
  recordForm.elements.nsCm.value = record.nsCm || "";
  ewDirInput.value = normalizeEwDir(record.ewDir);
  recordForm.elements.ewCm.value = record.ewCm || "";
  syncDirectionTabsFromForm();

  setLayerFromValue(record.layerName);
  recordForm.elements.detail.value = record.detail || "";
  recordForm.elements.detailSub.value = record.detailSub || "";
  if (recordForm.elements.rectSide1Cm) {
    recordForm.elements.rectSide1Cm.value = record.rectSide1Cm || "";
  }
  if (recordForm.elements.rectSide2Cm) {
    recordForm.elements.rectSide2Cm.value = record.rectSide2Cm || "";
  }
  if (recordForm.elements.ellipseLongRadiusCm) {
    recordForm.elements.ellipseLongRadiusCm.value = record.ellipseLongRadiusCm || "";
  }
  if (recordForm.elements.ellipseShortRadiusCm) {
    recordForm.elements.ellipseShortRadiusCm.value = record.ellipseShortRadiusCm || "";
  }
  [
    "imgP1NsDir",
    "imgP1NsCm",
    "imgP1EwDir",
    "imgP1EwCm",
    "imgP2NsDir",
    "imgP2NsCm",
    "imgP2EwDir",
    "imgP2EwCm",
    "imgP3NsDir",
    "imgP3NsCm",
    "imgP3EwDir",
    "imgP3EwCm",
    "imgP4NsDir",
    "imgP4NsCm",
    "imgP4EwDir",
    "imgP4EwCm",
  ].forEach((name) => {
    const field = recordForm.elements.namedItem(name);
    if (field instanceof HTMLInputElement || field instanceof HTMLSelectElement) {
      field.value = value(record?.[name]);
    }
  });
  setDefaultImageCornerDirections();
  ["imgP1NsDir", "imgP1EwDir", "imgP2NsDir", "imgP2EwDir", "imgP3NsDir", "imgP3EwDir", "imgP4NsDir", "imgP4EwDir"].forEach((name) => {
    const field = recordForm.elements.namedItem(name);
    if (field instanceof HTMLSelectElement) {
      const savedValue = value(record?.[name]);
      if (savedValue) {
        field.value = savedValue;
      }
    }
  });
  recordForm.elements.layerRef.value = record.layerRef || "";
  recordForm.elements.layerFromCm.value = record.layerFromCm || "";
  recordForm.elements.notes.value = record.notes || "";
  clearCarryForwardSavedFields();
  clearOverwriteUpdatedState();
  currentSectionDiagrams = clonePhotos(record.sectionDiagrams || []);
  renderSectionDiagramList();
  currentPhotos = clonePhotos(record.photos || []);
  renderPhotoList();
  updateEditMissingRequiredHighlights();
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
  updateEditMissingRequiredHighlights();
}

function buildCurrentEditDraftRecord() {
  if (!recordForm) {
    return null;
  }
  const formData = new FormData(recordForm);
  const teamState = normalizeTeamState(value(editTeamInput?.value), value(editTeamOtherInput?.value));
  const specimenPrefix = normalizeSpecimenPrefix(value(formData.get("specimenPrefix")));
  const specimenSerial = compactNoSpaceValue(formData.get("specimenSerial"));
  const draftRawShapeType = value(formData.get("largeShapeType"));
  const draftShapeType = normalizeLargeShapeType(draftRawShapeType) || normalizeLargeShapeLabel(draftRawShapeType);
  const draftIsImageShape = isLargeShapeImageType(draftShapeType);
  const imageCornerFields = extractImageCornerFieldsFromFormData(formData);
  return {
    kuwaku: buildKuwaku(
      normalizeKuwakuHeadA(editKuwakuHeadAInput?.value),
      normalizeKuwakuHeadB(editKuwakuHeadBInput?.value),
      normalizeKuwakuBlock(editKuwakuBlockInput?.value),
      normalizeKuwakuNo(editKuwakuNoInput?.value)
    ),
    levelHeight: value(editLevelHeightInput?.value),
    date: value(editDateInput?.value),
    team: teamState.team,
    teamOther: teamState.teamOther,
    teamLead: value(editTeamLeadInput?.value),
    recorder: value(editRecorderInput?.value),
    specimenPrefix,
    specimenSerial,
    specimenNo: buildSpecimenNo(specimenPrefix, specimenSerial),
    analysisType: specimenPrefix === "a" ? normalizeAnalysisType(value(formData.get("analysisType"))) : "",
    nameMemo: value(formData.get("nameMemo")),
    importantFlag: value(formData.get("importantFlag")),
    simpleRecordFlag: value(formData.get("simpleRecordFlag")),
    discoverer: value(formData.get("discoverer")),
    identifier: value(formData.get("identifier")),
    levelUpperCm: value(formData.get("levelUpperCm")),
    levelLowerCm: value(formData.get("levelLowerCm")),
    occurrenceSection: value(formData.get("occurrenceSection")),
    occurrenceSketch: value(formData.get("occurrenceSketch")),
    sectionDiagramDistanceChecked: normalizeChecklistChecked(formData.get("sectionDiagramDistanceChecked")),
    sectionDiagramHorizonChecked: normalizeChecklistChecked(formData.get("sectionDiagramHorizonChecked")),
    sectionDiagramLayerFaciesChecked: normalizeChecklistChecked(formData.get("sectionDiagramLayerFaciesChecked")),
    photoClinometerChecked: normalizeChecklistChecked(formData.get("photoClinometerChecked")),
    photoRulerChecked: normalizeChecklistChecked(formData.get("photoRulerChecked")),
    nsDir: value(formData.get("nsDir")),
    nsCm: value(formData.get("nsCm")),
    ewDir: value(formData.get("ewDir")),
    ewCm: value(formData.get("ewCm")),
    planSizeMode: normalizePlanSizeMode(value(formData.get("planSizeMode"))),
    largeShapeType: draftShapeType,
    largeAxisDirection: draftIsImageShape ? "" : normalizeLargeAxisDirection(value(formData.get("largeAxisDirection"))),
    lineLengthCm: value(formData.get("lineLengthCm")),
    line1NsDir: "",
    line1NsCm: "",
    line1EwDir: "",
    line1EwCm: "",
    line2NsDir: "",
    line2NsCm: "",
    line2EwDir: "",
    line2EwCm: "",
    rectSide1Cm: value(formData.get("rectSide1Cm")),
    rectSide2Cm: value(formData.get("rectSide2Cm")),
    ellipseLongRadiusCm: value(formData.get("ellipseLongRadiusCm")),
    ellipseShortRadiusCm: value(formData.get("ellipseShortRadiusCm")),
    imgP1NsDir: imageCornerFields.imgP1NsDir,
    imgP1NsCm: imageCornerFields.imgP1NsCm,
    imgP1EwDir: imageCornerFields.imgP1EwDir,
    imgP1EwCm: imageCornerFields.imgP1EwCm,
    imgP2NsDir: imageCornerFields.imgP2NsDir,
    imgP2NsCm: imageCornerFields.imgP2NsCm,
    imgP2EwDir: imageCornerFields.imgP2EwDir,
    imgP2EwCm: imageCornerFields.imgP2EwCm,
    imgP3NsDir: imageCornerFields.imgP3NsDir,
    imgP3NsCm: imageCornerFields.imgP3NsCm,
    imgP3EwDir: imageCornerFields.imgP3EwDir,
    imgP3EwCm: imageCornerFields.imgP3EwCm,
    imgP4NsDir: imageCornerFields.imgP4NsDir,
    imgP4NsCm: imageCornerFields.imgP4NsCm,
    imgP4EwDir: imageCornerFields.imgP4EwDir,
    imgP4EwCm: imageCornerFields.imgP4EwCm,
    layerName: getSelectedLayerName(),
    unit: compactNoSpaceValue(formData.get("unit")),
    detail: compactNoSpaceValue(formData.get("detail")),
    detailSub: value(formData.get("detailSub")),
    layerRef: value(formData.get("layerRef")),
    layerRelative: value(formData.get("layerRelative")),
    layerFromCm: value(formData.get("layerFromCm")),
    notes: value(formData.get("notes")),
    sectionDiagrams: clonePhotos(currentSectionDiagrams),
    photos: clonePhotos(currentPhotos),
  };
}

function copyCurrentEditToInput() {
  if (getActiveTabId() !== "edit-tab") {
    return;
  }
  const draft = buildCurrentEditDraftRecord();
  if (!draft) {
    showToast("編集内容のコピーに失敗しました");
    return;
  }

  const kuwakuParts = parseKuwaku(draft.kuwaku);
  const teamState = normalizeTeamState(draft.team, draft.teamOther);
  if (siteForm?.elements) {
    siteForm.elements.kuwakuHeadA.value = kuwakuParts.headA || DEFAULT_KUWAKU_HEAD_A;
    siteForm.elements.kuwakuHeadB.value = kuwakuParts.headB || DEFAULT_KUWAKU_HEAD_B;
    siteForm.elements.kuwakuBlock.value = kuwakuParts.block || "";
    siteForm.elements.kuwakuNo.value = kuwakuParts.no || "";
    siteForm.elements.levelHeight.value = value(draft.levelHeight);
    siteForm.elements.date.value = value(draft.date);
    siteForm.elements.team.value = teamState.team;
    siteForm.elements.teamOther.value = teamState.teamOther;
    siteForm.elements.teamLead.value = value(draft.teamLead);
    siteForm.elements.recorder.value = value(draft.recorder);
    syncTeamOtherInput(siteForm.elements.team.value);
  }

  isOverwriteMode = false;
  overwriteOriginalRecord = null;
  clearOverwriteUpdatedState();
  clearEditHistory();
  editingRecordId = null;
  if (recordIdInput) {
    recordIdInput.value = "";
  }

  setActiveTab("input-tab");
  populateRecordForm({
    ...draft,
    id: "",
    sectionDiagrams: [],
    photos: [],
    sectionDiagramDistanceChecked: "",
    sectionDiagramHorizonChecked: "",
    sectionDiagramLayerFaciesChecked: "",
    photoClinometerChecked: "",
    photoRulerChecked: "",
  });
  editingRecordId = null;
  if (recordIdInput) {
    recordIdInput.value = "";
  }
  updateDuplicateSpecimenWarning();
  showToast("コピーして新規入力を作成しました");
}

function copySavedRecordToInput(recordId, preferredKuwaku = "") {
  const record = findRecord(recordId);
  if (!record) {
    showToast("対象データが見つかりません");
    return;
  }

  const kuwakuSource = value(preferredKuwaku) || value(record.kuwaku) || getRecordKuwaku(record);
  const kuwakuParts = parseKuwaku(kuwakuSource);
  const teamState = normalizeTeamState(value(record.team), value(record.teamOther));
  if (siteForm?.elements) {
    siteForm.elements.kuwakuHeadA.value = kuwakuParts.headA || DEFAULT_KUWAKU_HEAD_A;
    siteForm.elements.kuwakuHeadB.value = kuwakuParts.headB || DEFAULT_KUWAKU_HEAD_B;
    siteForm.elements.kuwakuBlock.value = kuwakuParts.block || "";
    siteForm.elements.kuwakuNo.value = kuwakuParts.no || "";
    siteForm.elements.levelHeight.value = value(record.levelHeight);
    siteForm.elements.date.value = value(record.date);
    siteForm.elements.team.value = teamState.team;
    siteForm.elements.teamOther.value = teamState.teamOther;
    siteForm.elements.teamLead.value = value(record.teamLead);
    siteForm.elements.recorder.value = value(record.recorder);
    syncTeamOtherInput(siteForm.elements.team.value);
  }

  isOverwriteMode = false;
  overwriteOriginalRecord = null;
  clearOverwriteUpdatedState();
  clearEditHistory();
  editingRecordId = null;
  if (recordIdInput) {
    recordIdInput.value = "";
  }

  setActiveTab("input-tab");
  populateRecordForm({
    ...record,
    id: "",
    sectionDiagrams: [],
    photos: [],
    sectionDiagramDistanceChecked: "",
    sectionDiagramHorizonChecked: "",
    sectionDiagramLayerFaciesChecked: "",
    photoClinometerChecked: "",
    photoRulerChecked: "",
  });
  editingRecordId = null;
  if (recordIdInput) {
    recordIdInput.value = "";
  }
  updateDuplicateSpecimenWarning();
  showToast("コピーして新規入力を作成しました");
}

function getRecordFormFieldByName(name) {
  if (!recordForm?.elements) {
    return null;
  }
  const field = recordForm.elements.namedItem(name);
  if (field instanceof Element) {
    return field;
  }
  if (field instanceof RadioNodeList && field.length > 0 && field[0] instanceof Element) {
    return field[0];
  }
  return null;
}

function markEditMissingFieldByName(name) {
  const field = getRecordFormFieldByName(name);
  if (field) {
    field.classList.add("edit-missing-field");
  }
}

function markEditMissingGroupByName(name) {
  const field = getRecordFormFieldByName(name);
  const group = field?.closest(".inline-fieldset");
  if (group) {
    group.classList.add("edit-missing-group");
  }
}

function clearEditMissingRequiredHighlights() {
  document.querySelectorAll(".edit-missing-field").forEach((element) => {
    element.classList.remove("edit-missing-field");
  });
  document.querySelectorAll(".edit-missing-group").forEach((element) => {
    element.classList.remove("edit-missing-group");
  });
}

function updateEditMissingRequiredHighlights() {
  clearEditMissingRequiredHighlights();
  if (getActiveTabId() !== "edit-tab") {
    return;
  }
  const draftRecord = buildCurrentEditDraftRecord();
  const missingKeys = getMissingRequiredKeys(draftRecord);
  if (!missingKeys.size) {
    return;
  }

  if (hasAnyMissingRequiredKey(missingKeys, ["kuwakuHeadA", "kuwakuHeadB", "kuwakuBlock", "kuwakuNo"])) {
    [editKuwakuHeadAInput, editKuwakuHeadBInput, editKuwakuBlockInput, editKuwakuNoInput].forEach((input) => {
      if (input) {
        input.classList.add("edit-missing-field");
      }
    });
  }
  if (missingKeys.has("levelHeight") && editLevelHeightInput) {
    editLevelHeightInput.classList.add("edit-missing-field");
  }
  if (missingKeys.has("date") && editDateInput) {
    editDateInput.classList.add("edit-missing-field");
  }
  if (missingKeys.has("team") && editTeamInput) {
    editTeamInput.classList.add("edit-missing-field");
  }
  if (missingKeys.has("teamOther") && editTeamOtherInput) {
    editTeamOtherInput.classList.add("edit-missing-field");
  }
  if (missingKeys.has("teamLead") && editTeamLeadInput) {
    editTeamLeadInput.classList.add("edit-missing-field");
  }
  if (missingKeys.has("recorder") && editRecorderInput) {
    editRecorderInput.classList.add("edit-missing-field");
  }

  if (missingKeys.has("specimenSerial")) {
    specimenSerialInput.classList.add("edit-missing-field");
  }
  if (missingKeys.has("analysisType") && analysisTypeSelect) {
    analysisTypeSelect.classList.add("edit-missing-field");
  }
  if (missingKeys.has("nameMemo")) {
    markEditMissingFieldByName("nameMemo");
  }
  if (missingKeys.has("importantFlag")) {
    markEditMissingGroupByName("importantFlag");
  }
  if (missingKeys.has("simpleRecordFlag")) {
    markEditMissingGroupByName("simpleRecordFlag");
  }
  if (missingKeys.has("discoverer")) {
    markEditMissingFieldByName("discoverer");
  }
  if (missingKeys.has("identifier")) {
    markEditMissingFieldByName("identifier");
  }
  if (missingKeys.has("levelUpperCm")) {
    markEditMissingFieldByName("levelUpperCm");
  }
  if (missingKeys.has("levelLowerCm")) {
    markEditMissingFieldByName("levelLowerCm");
  }
  if (missingKeys.has("occurrenceSection")) {
    markEditMissingGroupByName("occurrenceSection");
  }
  if (missingKeys.has("occurrenceSketch")) {
    markEditMissingGroupByName("occurrenceSketch");
  }
  if (missingKeys.has("nsDir")) {
    markEditMissingGroupByName("nsDir");
  }
  if (missingKeys.has("nsCm")) {
    markEditMissingFieldByName("nsCm");
  }
  if (missingKeys.has("ewDir")) {
    markEditMissingGroupByName("ewDir");
  }
  if (missingKeys.has("ewCm")) {
    markEditMissingFieldByName("ewCm");
  }
  if (missingKeys.has("largeShapeType")) {
    markEditMissingGroupByName("planSizeMode");
  }
  if (missingKeys.has("largeAxisDirection")) {
    markEditMissingFieldByName("largeAxisDirection");
  }
  if (missingKeys.has("lineLengthCm")) {
    markEditMissingFieldByName("lineLengthCm");
  }
  if (missingKeys.has("rectSide1Cm")) {
    markEditMissingFieldByName("rectSide1Cm");
  }
  if (missingKeys.has("rectSide2Cm")) {
    markEditMissingFieldByName("rectSide2Cm");
  }
  if (missingKeys.has("ellipseLongRadiusCm")) {
    markEditMissingFieldByName("ellipseLongRadiusCm");
  }
  if (missingKeys.has("ellipseShortRadiusCm")) {
    markEditMissingFieldByName("ellipseShortRadiusCm");
  }
  [
    "imgP1NsCm",
    "imgP1EwCm",
    "imgP2NsCm",
    "imgP2EwCm",
    "imgP3NsCm",
    "imgP3EwCm",
    "imgP4NsCm",
    "imgP4EwCm",
  ].forEach((name) => {
    if (missingKeys.has(name)) {
      markEditMissingFieldByName(name);
    }
  });
  if (missingKeys.has("layerName")) {
    markEditMissingGroupByName("layerName");
  }
  if (missingKeys.has("layerOther")) {
    layerOtherInput.classList.add("edit-missing-field");
  }
  if (missingKeys.has("unit")) {
    markEditMissingFieldByName("unit");
  }
  if (missingKeys.has("layerRef")) {
    markEditMissingFieldByName("layerRef");
  }
  if (missingKeys.has("layerRelative")) {
    markEditMissingGroupByName("layerRelative");
  }
  if (missingKeys.has("layerFromCm")) {
    markEditMissingFieldByName("layerFromCm");
  }
  if (
    hasAnyMissingRequiredKey(missingKeys, [
      "sectionDiagrams",
      "sectionDiagramDistanceChecked",
      "sectionDiagramHorizonChecked",
      "sectionDiagramLayerFaciesChecked",
    ])
  ) {
    const sectionWrap = recordForm?.querySelector(".diagram-upload-wrap");
    if (sectionWrap) {
      sectionWrap.classList.add("edit-missing-group");
    }
  }
  if (hasAnyMissingRequiredKey(missingKeys, ["photoClinometerChecked", "photoRulerChecked"])) {
    const photoWrap = recordForm?.querySelector(".photo-upload-wrap");
    if (photoWrap) {
      photoWrap.classList.add("edit-missing-group");
    }
  }
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
            <button type="button" data-action="copy-to-input" data-id="${record.id}" data-kuwaku="${escapeHtml(
              getRecordKuwaku(record)
            )}">コピーして新規入力</button>
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
  renderViewerOutput();
  renderExportOutput();
}

function renderExportOutput() {
  renderExportListRangeControls();
  renderExportCardRangeControls();
  renderExportPlanControls();
  updateExportButtonAvailability();
}

function renderExportListRangeControls() {
  const kuwakuScopedSource = getRecordsByExportRangeFilters({
    kuwakuValue: ALL_GRIDS_VALUE,
    categoryValue: exportListRangeCategory,
    statusValue: exportListRangeStatus,
    specimenFromRaw: exportListRangeSpecimenFrom,
    specimenToRaw: exportListRangeSpecimenTo,
    dateFromRaw: exportListRangeDateFrom,
    dateToRaw: exportListRangeDateTo,
  });
  const kuwakuOptions = collectExportKuwakuOptionsWithCounts(kuwakuScopedSource);
  if (!kuwakuOptions.some((item) => item.value === exportListRangeKuwaku)) {
    exportListRangeKuwaku = ALL_GRIDS_VALUE;
  }
  if (exportListRangeKuwakuSelect) {
    exportListRangeKuwakuSelect.innerHTML = kuwakuOptions
      .map(
        (item) =>
          `<option value="${escapeHtml(item.value)}" ${item.value === exportListRangeKuwaku ? "selected" : ""}>${escapeHtml(
            item.label
          )}</option>`
      )
      .join("");
  }

  const categoryScopedSource = getRecordsByExportRangeFilters({
    kuwakuValue: exportListRangeKuwaku,
    categoryValue: EXPORT_CATEGORY_ALL_VALUE,
    statusValue: exportListRangeStatus,
    specimenFromRaw: exportListRangeSpecimenFrom,
    specimenToRaw: exportListRangeSpecimenTo,
    dateFromRaw: exportListRangeDateFrom,
    dateToRaw: exportListRangeDateTo,
  });
  const categoryOptions = collectExportCategoryOptions(categoryScopedSource);
  if (!categoryOptions.some((item) => item.value === exportListRangeCategory)) {
    exportListRangeCategory = EXPORT_CATEGORY_ALL_VALUE;
  }
  if (exportListRangeCategorySelect) {
    exportListRangeCategorySelect.innerHTML = categoryOptions
      .map(
        (item) =>
          `<option value="${escapeHtml(item.value)}" ${item.value === exportListRangeCategory ? "selected" : ""}>${escapeHtml(
            item.label
          )}</option>`
      )
      .join("");
  }
  if (exportListRangeStatusSelect) {
    if (!["all", "complete", "incomplete"].includes(exportListRangeStatus)) {
      exportListRangeStatus = "all";
    }
    exportListRangeStatusSelect.value = exportListRangeStatus;
  }
  if (exportListRangeSpecimenFromInput) {
    exportListRangeSpecimenFromInput.value = exportListRangeSpecimenFrom;
  }
  if (exportListRangeSpecimenToInput) {
    exportListRangeSpecimenToInput.value = exportListRangeSpecimenTo;
  }
  if (exportListRangeDateFromInput) {
    exportListRangeDateFromInput.value = exportListRangeDateFrom;
  }
  if (exportListRangeDateToInput) {
    exportListRangeDateToInput.value = exportListRangeDateTo;
  }

  const filteredRecords = getListExportRecords();
  if (exportListRangeSummaryEl) {
    const hasData = filteredRecords.length > 0;
    exportListRangeSummaryEl.textContent = `対象件数: ${filteredRecords.length}件（${hasData ? "対象あり" : "対象なし"}）`;
    setAvailabilityClass(exportListRangeSummaryEl, hasData);
  }
}

function renderExportCardRangeControls() {
  const kuwakuScopedSource = getRecordsByExportRangeFilters({
    kuwakuValue: ALL_GRIDS_VALUE,
    categoryValue: exportCardRangeCategory,
    statusValue: exportCardRangeStatus,
    dateFromRaw: exportCardRangeDateFrom,
    dateToRaw: exportCardRangeDateTo,
  });
  const kuwakuOptions = collectExportKuwakuOptionsWithCounts(kuwakuScopedSource);
  if (!kuwakuOptions.some((item) => item.value === exportCardRangeKuwaku)) {
    exportCardRangeKuwaku = ALL_GRIDS_VALUE;
  }
  if (exportCardRangeKuwakuSelect) {
    exportCardRangeKuwakuSelect.innerHTML = kuwakuOptions
      .map(
        (item) =>
          `<option value="${escapeHtml(item.value)}" ${item.value === exportCardRangeKuwaku ? "selected" : ""}>${escapeHtml(
            item.label
          )}</option>`
      )
      .join("");
  }

  const categoryScopedSource = getRecordsByExportRangeFilters({
    kuwakuValue: exportCardRangeKuwaku,
    categoryValue: EXPORT_CATEGORY_ALL_VALUE,
    statusValue: exportCardRangeStatus,
    dateFromRaw: exportCardRangeDateFrom,
    dateToRaw: exportCardRangeDateTo,
  });
  const categoryOptions = collectExportCategoryOptions(categoryScopedSource);
  if (!categoryOptions.some((item) => item.value === exportCardRangeCategory)) {
    exportCardRangeCategory = EXPORT_CATEGORY_ALL_VALUE;
  }
  if (exportCardRangeCategorySelect) {
    exportCardRangeCategorySelect.innerHTML = categoryOptions
      .map(
        (item) =>
          `<option value="${escapeHtml(item.value)}" ${item.value === exportCardRangeCategory ? "selected" : ""}>${escapeHtml(
            item.label
          )}</option>`
      )
      .join("");
  }
  if (exportCardRangeStatusSelect) {
    if (!["all", "complete", "incomplete"].includes(exportCardRangeStatus)) {
      exportCardRangeStatus = "all";
    }
    exportCardRangeStatusSelect.value = exportCardRangeStatus;
  }
  if (exportCardRangeDateFromInput) {
    exportCardRangeDateFromInput.value = exportCardRangeDateFrom;
  }
  if (exportCardRangeDateToInput) {
    exportCardRangeDateToInput.value = exportCardRangeDateTo;
  }

  const filteredRecords = getCardExportRecords();
  if (exportCardRangeSummaryEl) {
    const hasData = filteredRecords.length > 0;
    exportCardRangeSummaryEl.textContent = `対象件数: ${filteredRecords.length}件（${hasData ? "対象あり" : "対象なし"}）`;
    setAvailabilityClass(exportCardRangeSummaryEl, hasData);
  }
}

function renderExportPlanControls() {
  const sortedRecords = getRecordsByExportRangeFilters({
    kuwakuValue: ALL_GRIDS_VALUE,
    categoryValue: EXPORT_CATEGORY_ALL_VALUE,
    statusValue: "all",
    dateFromRaw: exportPlanDateFrom,
    dateToRaw: exportPlanDateTo,
  });
  const kuwakuOptions = collectExportKuwakuOptionsWithCounts(sortedRecords).filter((item) => item.value !== ALL_GRIDS_VALUE);
  if (!kuwakuOptions.length) {
    exportPlanKuwaku = "";
  } else if (!kuwakuOptions.some((item) => item.value === exportPlanKuwaku)) {
    exportPlanKuwaku = kuwakuOptions[0].value;
  }
  if (exportPlanKuwakuSelect) {
    exportPlanKuwakuSelect.innerHTML = kuwakuOptions
      .map(
        (item) =>
          `<option value="${escapeHtml(item.value)}" ${item.value === exportPlanKuwaku ? "selected" : ""}>${escapeHtml(
            item.label
          )}</option>`
      )
      .join("");
  }

  const kuwakuScopedRecords =
    !exportPlanKuwaku
      ? []
      : sortedRecords.filter((record) => kuwakuValueForSelect(getRecordKuwaku(record)) === exportPlanKuwaku);
  const categoryScopedRecords =
    exportPlanCategory === EXPORT_CATEGORY_ALL_VALUE
      ? kuwakuScopedRecords
      : kuwakuScopedRecords.filter((record) => {
          const specimen = parseSpecimenNo(record.specimenNo, record.specimenPrefix, record.specimenSerial);
          return normalizeSpecimenPrefix(specimen.prefix) === exportPlanCategory;
        });

  const categoryOptions = collectExportCategoryOptions(kuwakuScopedRecords);
  if (!categoryOptions.some((item) => item.value === exportPlanCategory)) {
    exportPlanCategory = EXPORT_CATEGORY_ALL_VALUE;
  }
  if (exportPlanCategorySelect) {
    exportPlanCategorySelect.innerHTML = categoryOptions
      .map(
        (item) =>
          `<option value="${escapeHtml(item.value)}" ${item.value === exportPlanCategory ? "selected" : ""}>${escapeHtml(
            item.label
          )}</option>`
      )
      .join("");
  }
  if (exportPlanDateFromInput) {
    exportPlanDateFromInput.value = exportPlanDateFrom;
  }
  if (exportPlanDateToInput) {
    exportPlanDateToInput.value = exportPlanDateTo;
  }
  syncExportPlanModeControls(categoryScopedRecords);

  const groups = buildPlanPdfGroupsForExport({
    kuwakuValue: exportPlanKuwaku,
    categoryValue: exportPlanCategory,
    dateFromRaw: exportPlanDateFrom,
    dateToRaw: exportPlanDateTo,
    modeSelections: getExportPlanModeSelections(),
  });
  const recordCount = groups.reduce((sum, group) => sum + (Number.isFinite(group.count) ? group.count : 0), 0);
  if (exportPlanSummaryEl) {
    const hasData = recordCount > 0;
    exportPlanSummaryEl.textContent = `PDFページ対象: ${groups.length}ページ / 記録 ${recordCount}件（${hasData ? "対象あり" : "対象なし"}）`;
    setAvailabilityClass(exportPlanSummaryEl, hasData);
  }
}

function getExportPlanScopedRecords() {
  const sortedRecords = getRecordsByExportRangeFilters({
    kuwakuValue: ALL_GRIDS_VALUE,
    categoryValue: EXPORT_CATEGORY_ALL_VALUE,
    statusValue: "all",
    dateFromRaw: exportPlanDateFrom,
    dateToRaw: exportPlanDateTo,
  });
  if (!exportPlanKuwaku) {
    return [];
  }
  const kuwakuScopedRecords = sortedRecords.filter(
    (record) => kuwakuValueForSelect(getRecordKuwaku(record)) === exportPlanKuwaku
  );
  if (exportPlanCategory === EXPORT_CATEGORY_ALL_VALUE) {
    return kuwakuScopedRecords;
  }
  return kuwakuScopedRecords.filter((record) => {
    const specimen = parseSpecimenNo(record.specimenNo, record.specimenPrefix, record.specimenSerial);
    return normalizeSpecimenPrefix(specimen.prefix) === exportPlanCategory;
  });
}

function updateExportButtonAvailability() {
  const listCount = getListExportRecords().length;
  const cardCount = getCardExportRecords().length;
  const planGroups = buildPlanPdfGroupsForExport({
    kuwakuValue: exportPlanKuwaku,
    categoryValue: exportPlanCategory,
    dateFromRaw: exportPlanDateFrom,
    dateToRaw: exportPlanDateTo,
    modeSelections: getExportPlanModeSelections(),
  });
  const planRecordCount = planGroups.reduce((sum, group) => sum + (Number.isFinite(group.count) ? group.count : 0), 0);

  if (exportListCsvBtn) {
    exportListCsvBtn.disabled = listCount === 0;
  }
  if (exportListPdfBtn) {
    exportListPdfBtn.disabled = listCount === 0;
  }
  if (exportCardCsvBtn) {
    exportCardCsvBtn.disabled = cardCount === 0;
  }
  if (exportCardPdfBtn) {
    exportCardPdfBtn.disabled = cardCount === 0;
  }
  if (exportPlanPdfBtn) {
    exportPlanPdfBtn.disabled = planRecordCount === 0;
  }
}

function setAvailabilityClass(element, hasData) {
  if (!element) {
    return;
  }
  element.classList.toggle("has-data", hasData);
  element.classList.toggle("no-data", !hasData);
}

function syncExportPlanModeControls(recordsRaw) {
  const records = Array.isArray(recordsRaw) ? recordsRaw : [];
  const unitOptions = collectExportPlanValueOptions(records, (record) => unitValueForSelect(record.unit), unitLabelForSelect);
  exportPlanModeUnitValues = syncExportPlanModeValues(exportPlanModeUnitValues, unitOptions, {
    autoSelectAllWhenEmpty: !exportPlanModeUnitTouched,
  });
  renderExportPlanUnitButtons(exportPlanModeUnitButtons, unitOptions, exportPlanModeUnitValues);

  exportPlanModeDetailUnitValue = syncExportPlanSingleValue(exportPlanModeDetailUnitValue, unitOptions);
  renderExportPlanModeSelect(exportPlanModeDetailUnitSelect, unitOptions, exportPlanModeDetailUnitValue);

  const detailModeRecords = exportPlanModeDetailUnitValue
    ? filterPlanRecordsForMode(records, { unitValues: [exportPlanModeDetailUnitValue] })
    : [];
  const detailOptions = collectExportPlanValueOptions(
    detailModeRecords,
    (record) => detailValueForSelect(record.detail),
    detailLabelForSelect
  );
  exportPlanModeDetailValues = syncExportPlanModeValues(exportPlanModeDetailValues, detailOptions, {
    autoSelectAllWhenEmpty: !exportPlanModeDetailTouched,
  });
  renderExportPlanModeButtons(
    exportPlanModeDetailButtons,
    detailOptions,
    exportPlanModeDetailValues,
    "出力するサブユニットを選んでください"
  );

  exportPlanModeDetailSubUnitValue = syncExportPlanSingleValue(exportPlanModeDetailSubUnitValue, unitOptions);
  renderExportPlanModeSelect(exportPlanModeDetailSubUnitSelect, unitOptions, exportPlanModeDetailSubUnitValue);

  const detailSubBaseRecords = exportPlanModeDetailSubUnitValue
    ? filterPlanRecordsForMode(records, { unitValues: [exportPlanModeDetailSubUnitValue] })
    : [];
  const detailSubDetailOptions = collectExportPlanValueOptions(
    detailSubBaseRecords,
    (record) => detailValueForSelect(record.detail),
    detailLabelForSelect
  );
  exportPlanModeDetailSubDetailValue = syncExportPlanSingleValue(exportPlanModeDetailSubDetailValue, detailSubDetailOptions);
  renderExportPlanModeSelect(exportPlanModeDetailSubDetailSelect, detailSubDetailOptions, exportPlanModeDetailSubDetailValue);

  const detailSubRecords =
    exportPlanModeDetailSubUnitValue && exportPlanModeDetailSubDetailValue
      ? filterPlanRecordsForMode(records, {
          unitValues: [exportPlanModeDetailSubUnitValue],
          detailValues: [exportPlanModeDetailSubDetailValue],
        })
      : [];
  const detailSubOptions = collectExportPlanValueOptions(
    detailSubRecords,
    (record) => detailSubValueForSelect(record.detailSub),
    detailSubLabelForSelect
  );
  exportPlanModeDetailSubValues = syncExportPlanModeValues(exportPlanModeDetailSubValues, detailSubOptions, {
    autoSelectAllWhenEmpty: !exportPlanModeDetailSubTouched,
  });
  renderExportPlanModeButtons(
    exportPlanModeDetailSubButtons,
    detailSubOptions,
    exportPlanModeDetailSubValues,
    "出力するサブユニット細分を選んでください"
  );

  syncExportPlanModeCheckbox(exportPlanModeUnitCheck, unitOptions.length > 0, "unit");
  syncExportPlanModeCheckbox(exportPlanModeDetailCheck, !!exportPlanModeDetailUnitValue && detailOptions.length > 0, "detail");
  syncExportPlanModeCheckbox(
    exportPlanModeDetailSubCheck,
    !!exportPlanModeDetailSubUnitValue && !!exportPlanModeDetailSubDetailValue && detailSubOptions.length > 0,
    "detailSub"
  );

  const unitScopedRecords = filterPlanRecordsForMode(records, { unitValues: exportPlanModeUnitValues });
  const detailScopedRecords =
    exportPlanModeDetailUnitValue && exportPlanModeDetailValues.size
      ? filterPlanRecordsForMode(records, {
          unitValues: [exportPlanModeDetailUnitValue],
          detailValues: exportPlanModeDetailValues,
        })
      : [];
  const detailSubScopedRecords =
    exportPlanModeDetailSubUnitValue && exportPlanModeDetailSubDetailValue && exportPlanModeDetailSubValues.size
      ? filterPlanRecordsForMode(records, {
          unitValues: [exportPlanModeDetailSubUnitValue],
          detailValues: [exportPlanModeDetailSubDetailValue],
          detailSubValues: exportPlanModeDetailSubValues,
        })
      : [];

  renderExportPlanModeStats(exportPlanModeUnitStats, unitScopedRecords);
  renderExportPlanModeStats(exportPlanModeDetailStats, detailScopedRecords);
  renderExportPlanModeStats(exportPlanModeDetailSubStats, detailSubScopedRecords);
}

function syncExportPlanModeCheckbox(checkbox, hasOptions, modeKey) {
  if (!checkbox) {
    return;
  }
  if (!hasOptions) {
    checkbox.checked = false;
    checkbox.disabled = true;
    return;
  }
  checkbox.disabled = false;
  if (modeKey === "unit") {
    checkbox.checked = exportPlanModeUnitEnabled;
  } else if (modeKey === "detail") {
    checkbox.checked = exportPlanModeDetailEnabled;
  } else if (modeKey === "detailSub") {
    checkbox.checked = exportPlanModeDetailSubEnabled;
  }
}

function collectExportPlanValueOptions(recordsRaw, valueGetter, labelGetter) {
  const records = Array.isArray(recordsRaw) ? recordsRaw : [];
  const countMap = new Map();
  records.forEach((record) => {
    const optionValue = value(valueGetter(record));
    if (!optionValue) {
      return;
    }
    countMap.set(optionValue, (countMap.get(optionValue) || 0) + 1);
  });
  return Array.from(countMap.entries())
    .map(([optionValue, count]) => ({
      value: optionValue,
      label: `${labelGetter(optionValue)}（${count}件）`,
    }))
    .sort((a, b) => a.label.localeCompare(b.label, "ja", { numeric: true, sensitivity: "base" }));
}

function syncExportPlanModeValues(currentValuesRaw, options, config = {}) {
  const next = new Set();
  const currentValues = currentValuesRaw instanceof Set ? currentValuesRaw : new Set();
  const optionValues = (Array.isArray(options) ? options : []).map((option) => value(option.value)).filter(Boolean);
  const valid = new Set(optionValues);
  currentValues.forEach((selectedValue) => {
    const normalized = value(selectedValue);
    if (normalized && valid.has(normalized)) {
      next.add(normalized);
    }
  });
  if (!next.size && optionValues.length && config.autoSelectAllWhenEmpty) {
    optionValues.forEach((optionValue) => next.add(optionValue));
  }
  return next;
}

function syncExportPlanSingleValue(currentValueRaw, options) {
  const currentValue = value(currentValueRaw);
  const optionValues = (Array.isArray(options) ? options : []).map((option) => value(option.value)).filter(Boolean);
  if (!optionValues.length) {
    return "";
  }
  if (currentValue && optionValues.includes(currentValue)) {
    return currentValue;
  }
  return optionValues[0];
}

function renderExportPlanUnitButtons(container, options, selectedValuesRaw) {
  if (!container) {
    return;
  }
  const optionList = Array.isArray(options) ? options : [];
  if (!optionList.length) {
    container.innerHTML = '<span class="muted">候補なし</span>';
    return;
  }
  const selectedValues = normalizeSelectionSet(selectedValuesRaw);
  const validValues = new Set(optionList.map((option) => value(option.value)).filter(Boolean));
  let selectedCount = 0;
  validValues.forEach((optionValue) => {
    if (selectedValues.has(optionValue)) {
      selectedCount += 1;
    }
  });
  const allSelected = validValues.size > 0 && selectedCount === validValues.size;
  const showHint = selectedCount === 0;

  const allButtonHtml = `
    <div class="export-plan-unit-button-row all-row">
      <button type="button" class="export-plan-option-button export-plan-option-button-all ${
        allSelected ? "active" : ""
      }" data-value="${EXPORT_PLAN_ALL_UNITS_BUTTON_VALUE}">全ユニット</button>
      ${showHint ? '<span class="export-plan-select-hint">出力するユニットを選んでください</span>' : ""}
    </div>
  `;
  const unitButtonHtml = optionList
    .map(
      (option) =>
        `<button type="button" class="export-plan-option-button ${
          selectedValues.has(option.value) ? "active" : ""
        }" data-value="${escapeHtml(option.value)}">${escapeHtml(option.label)}</button>`
    )
    .join("");
  container.innerHTML = `${allButtonHtml}<div class="export-plan-unit-button-row unit-row">${unitButtonHtml}</div>`;
}

function renderExportPlanModeButtons(container, options, selectedValues, emptyHintText = "") {
  if (!container) {
    return;
  }
  const optionList = Array.isArray(options) ? options : [];
  if (!optionList.length) {
    container.innerHTML = '<span class="muted">候補なし</span>';
    return;
  }
  const selectedSet = normalizeSelectionSet(selectedValues);
  const buttonHtml = optionList
    .map(
      (option) =>
        `<button type="button" class="export-plan-option-button ${
          selectedSet.has(option.value) ? "active" : ""
        }" data-value="${escapeHtml(option.value)}">${escapeHtml(
          option.label
        )}</button>`
    )
    .join("");
  const hintHtml = !selectedSet.size && value(emptyHintText) ? `<span class="export-plan-select-hint">${escapeHtml(emptyHintText)}</span>` : "";
  container.innerHTML = `${buttonHtml}${hintHtml}`;
}

function renderExportPlanModeSelect(selectEl, options, selectedValue) {
  if (!selectEl) {
    return;
  }
  const optionList = Array.isArray(options) ? options : [];
  if (!optionList.length) {
    selectEl.innerHTML = "";
    selectEl.disabled = true;
    return;
  }
  selectEl.disabled = false;
  selectEl.innerHTML = optionList
    .map(
      (option) =>
        `<option value="${escapeHtml(option.value)}" ${option.value === selectedValue ? "selected" : ""}>${escapeHtml(
          option.label
        )}</option>`
    )
    .join("");
}

function toggleSelectionInSet(targetSet, optionValueRaw, checkedForced = null) {
  const optionValue = value(optionValueRaw);
  if (!optionValue || !(targetSet instanceof Set)) {
    return;
  }
  if (checkedForced === true) {
    targetSet.add(optionValue);
    return;
  }
  if (checkedForced === false) {
    targetSet.delete(optionValue);
    return;
  }
  if (targetSet.has(optionValue)) {
    targetSet.delete(optionValue);
  } else {
    targetSet.add(optionValue);
  }
}

function normalizeSelectionSet(valuesRaw) {
  const next = new Set();
  if (valuesRaw instanceof Set) {
    valuesRaw.forEach((item) => {
      const normalized = value(item);
      if (normalized) {
        next.add(normalized);
      }
    });
    return next;
  }
  if (Array.isArray(valuesRaw)) {
    valuesRaw.forEach((item) => {
      const normalized = value(item);
      if (normalized) {
        next.add(normalized);
      }
    });
  }
  return next;
}

function filterPlanRecordsForMode(recordsRaw, selections = {}) {
  let records = Array.isArray(recordsRaw) ? recordsRaw : [];
  const unitValues = normalizeSelectionSet(selections.unitValues);
  const detailValues = normalizeSelectionSet(selections.detailValues);
  const detailSubValues = normalizeSelectionSet(selections.detailSubValues);
  if (unitValues.size) {
    records = records.filter((record) => unitValues.has(unitValueForSelect(record.unit)));
  }
  if (detailValues.size) {
    records = records.filter((record) => detailValues.has(detailValueForSelect(record.detail)));
  }
  if (detailSubValues.size) {
    records = records.filter((record) => detailSubValues.has(detailSubValueForSelect(record.detailSub)));
  }
  return records;
}

function buildExportPlanModeStats(recordsRaw) {
  const records = Array.isArray(recordsRaw) ? recordsRaw : [];
  let plottedCount = 0;
  records.forEach((record) => {
    if (buildPlanDrawable(record)) {
      plottedCount += 1;
    }
  });
  return {
    total: records.length,
    missing: Math.max(0, records.length - plottedCount),
  };
}

function renderExportPlanModeStats(targetEl, records) {
  if (!targetEl) {
    return;
  }
  const stats = buildExportPlanModeStats(records);
  const hasData = stats.total > 0;
  targetEl.textContent = `対象件数: ${stats.total}件 / 平面位置未記入: ${stats.missing}件`;
  setAvailabilityClass(targetEl, hasData);
}

function getExportPlanModeSelections() {
  return {
    unit: {
      enabled: exportPlanModeUnitEnabled,
      unitValues: Array.from(exportPlanModeUnitValues),
    },
    detail: {
      enabled: exportPlanModeDetailEnabled,
      unitValue: exportPlanModeDetailUnitValue,
      detailValues: Array.from(exportPlanModeDetailValues),
    },
    detailSub: {
      enabled: exportPlanModeDetailSubEnabled,
      unitValue: exportPlanModeDetailSubUnitValue,
      detailValue: exportPlanModeDetailSubDetailValue,
      detailSubValues: Array.from(exportPlanModeDetailSubValues),
    },
  };
}

function collectExportKuwakuOptionsWithCounts(records) {
  const list = Array.isArray(records) ? records : [];
  if (!list.length) {
    return [{ value: ALL_GRIDS_VALUE, label: "全グリッド（0件）" }];
  }
  const countMap = new Map();
  list.forEach((record) => {
    const key = kuwakuValueForSelect(getRecordKuwaku(record));
    countMap.set(key, (countMap.get(key) || 0) + 1);
  });
  const options = Array.from(countMap.entries())
    .sort((a, b) => kuwakuLabelForSelect(a[0]).localeCompare(kuwakuLabelForSelect(b[0]), "ja", { numeric: true, sensitivity: "base" }))
    .map(([kuwakuValue, count]) => ({
      value: kuwakuValue,
      label: `${kuwakuLabelForSelect(kuwakuValue)}（${count}件）`,
    }));
  return [{ value: ALL_GRIDS_VALUE, label: `全グリッド（${list.length}件）` }, ...options];
}

function collectExportCategoryOptions(records) {
  const list = Array.isArray(records) ? records : [];
  const countMap = new Map();
  list.forEach((record) => {
    const specimen = parseSpecimenNo(record.specimenNo, record.specimenPrefix, record.specimenSerial);
    const prefix = normalizeSpecimenPrefix(specimen.prefix);
    if (!prefix) {
      return;
    }
    countMap.set(prefix, (countMap.get(prefix) || 0) + 1);
  });
  const totalCount = list.length;
  const options = [{ value: EXPORT_CATEGORY_ALL_VALUE, label: `全分類（${totalCount}件）` }];
  Object.keys(SPECIMEN_CATEGORY_MAP).forEach((prefix) => {
    const count = countMap.get(prefix) || 0;
    options.push({
      value: prefix,
      label: `${prefix}: ${SPECIMEN_CATEGORY_MAP[prefix] || ""}（${count}件）`,
    });
  });
  return options;
}

function getListExportRecords() {
  return getRecordsByExportRangeFilters({
    kuwakuValue: exportListRangeKuwaku,
    categoryValue: exportListRangeCategory,
    statusValue: exportListRangeStatus,
    specimenFromRaw: exportListRangeSpecimenFrom,
    specimenToRaw: exportListRangeSpecimenTo,
    dateFromRaw: exportListRangeDateFrom,
    dateToRaw: exportListRangeDateTo,
  });
}

function getCardExportRecords() {
  return getRecordsByExportRangeFilters({
    kuwakuValue: exportCardRangeKuwaku,
    categoryValue: exportCardRangeCategory,
    statusValue: exportCardRangeStatus,
    dateFromRaw: exportCardRangeDateFrom,
    dateToRaw: exportCardRangeDateTo,
  });
}

function getRecordsByExportRangeFilters(filters = {}) {
  const records = [...state.records].sort(compareRecordsByKuwakuThenSpecimen);
  const kuwakuValue = value(filters.kuwakuValue) || ALL_GRIDS_VALUE;
  const categoryValue = value(filters.categoryValue) || EXPORT_CATEGORY_ALL_VALUE;
  const statusValue = value(filters.statusValue) || "all";
  const dateFrom = normalizeDateForExportRange(filters.dateFromRaw);
  const dateTo = normalizeDateForExportRange(filters.dateToRaw);
  let minDate = dateFrom;
  let maxDate = dateTo;
  if (minDate && maxDate && minDate > maxDate) {
    minDate = dateTo;
    maxDate = dateFrom;
  }
  const fromSpecimen = parseSpecimenForExportRange(filters.specimenFromRaw);
  const toSpecimen = parseSpecimenForExportRange(filters.specimenToRaw);
  let minSpecimen = fromSpecimen;
  let maxSpecimen = toSpecimen;
  if (minSpecimen && maxSpecimen && compareRecordsBySpecimenNo(minSpecimen, maxSpecimen) > 0) {
    minSpecimen = toSpecimen;
    maxSpecimen = fromSpecimen;
  }
  return records.filter((record) => {
    if (kuwakuValue !== ALL_GRIDS_VALUE && kuwakuValueForSelect(getRecordKuwaku(record)) !== kuwakuValue) {
      return false;
    }
    if (categoryValue !== EXPORT_CATEGORY_ALL_VALUE) {
      const specimen = parseSpecimenNo(record.specimenNo, record.specimenPrefix, record.specimenSerial);
      if (normalizeSpecimenPrefix(specimen.prefix) !== categoryValue) {
        return false;
      }
    }
    if (statusValue === "complete" && !isRecordDataComplete(record)) {
      return false;
    }
    if (statusValue === "incomplete" && isRecordDataComplete(record)) {
      return false;
    }
    if (minDate || maxDate) {
      const recordDate = normalizeDateForExportRange(getRecordDate(record));
      if (!recordDate) {
        return false;
      }
      if (minDate && recordDate < minDate) {
        return false;
      }
      if (maxDate && recordDate > maxDate) {
        return false;
      }
    }
    if (minSpecimen && compareRecordsBySpecimenNo(record, minSpecimen) < 0) {
      return false;
    }
    if (maxSpecimen && compareRecordsBySpecimenNo(record, maxSpecimen) > 0) {
      return false;
    }
    return true;
  });
}

function parseSpecimenForExportRange(specimenRaw) {
  const parsed = parseSpecimenNo(specimenRaw);
  if (!value(parsed.serial)) {
    return null;
  }
  return {
    specimenNo: parsed.specimenNo,
    specimenPrefix: parsed.prefix,
    specimenSerial: parsed.serial,
    id: "__export-range__",
  };
}

function normalizeDateForExportRange(dateRaw) {
  const text = value(dateRaw);
  if (!text) {
    return "";
  }
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) {
    return text;
  }
  const ms = Date.parse(text);
  if (!Number.isFinite(ms)) {
    return "";
  }
  return new Date(ms).toISOString().slice(0, 10);
}

function renderListOutput() {
  updateOutputListSortHeader();
  if (!state.records.length) {
    syncOutputKuwakuSelect([]);
    outputListBody.innerHTML = "<tr><td colspan=\"18\">出力対象データがありません。</td></tr>";
    return;
  }

  const filteredRecords = getFilteredOutputRecords();
  if (!filteredRecords.length) {
    outputListBody.innerHTML = "<tr><td colspan=\"18\">選択した区画のデータがありません。</td></tr>";
    return;
  }

  outputListBody.innerHTML = sortOutputRecordsForList(filteredRecords)
    .map((record) => {
      const selectedClass = record.id === selectedCardRecordId ? "selected-card-row" : "";
      const cardButtonLabel = record.id === selectedCardRecordId ? "プレビュー表示中" : "カード";
      const kuwakuText = getRecordKuwaku(record);
      const kuwakuStyle = getKuwakuCellStyle(kuwakuText);
      const categoryColor = getRecordSpecimenColor(record);
      const categoryBackground = toRgbaColor(categoryColor, 0.2);
      const categoryBorderColor = toRgbaColor(categoryColor, 0.45);
      const unitStyle = getUnitCellStyle(record.unit);
      const missingRequiredKeys = getMissingRequiredKeys(record);
      const dataComplete = missingRequiredKeys.size === 0;
      const missingTitle = formatMissingRequiredTooltip(missingRequiredKeys);
      const kuwakuMissing = hasAnyMissingRequiredKey(missingRequiredKeys, ["kuwakuHeadA", "kuwakuHeadB", "kuwakuBlock", "kuwakuNo"]);
      const teamMissing = hasAnyMissingRequiredKey(missingRequiredKeys, ["team", "teamOther"]);
      const specimenMissing = hasAnyMissingRequiredKey(missingRequiredKeys, ["specimenSerial"]);
      const categoryMissing = hasAnyMissingRequiredKey(missingRequiredKeys, ["analysisType"]);
      const nameMissing = hasAnyMissingRequiredKey(missingRequiredKeys, ["nameMemo"]);
      const importantMissing = hasAnyMissingRequiredKey(missingRequiredKeys, ["importantFlag"]);
      const unitMissing = hasAnyMissingRequiredKey(missingRequiredKeys, ["unit"]);
      const discovererMissing = hasAnyMissingRequiredKey(missingRequiredKeys, ["discoverer"]);
      const identifierMissing = hasAnyMissingRequiredKey(missingRequiredKeys, ["identifier"]);
      const levelHeightMissing = hasAnyMissingRequiredKey(missingRequiredKeys, ["levelHeight"]);
      const levelReadMissing = hasAnyMissingRequiredKey(missingRequiredKeys, ["levelUpperCm", "levelLowerCm"]);
      const sectionMissing = hasAnyMissingRequiredKey(missingRequiredKeys, ["occurrenceSection", "sectionDiagrams"]);
      const sectionChecklistMissing = hasAnyMissingRequiredKey(missingRequiredKeys, [
        "sectionDiagramDistanceChecked",
        "sectionDiagramHorizonChecked",
        "sectionDiagramLayerFaciesChecked",
      ]);
      const sketchMissing = hasAnyMissingRequiredKey(missingRequiredKeys, ["occurrenceSketch"]);
      const positionMissing = hasAnyMissingRequiredKey(missingRequiredKeys, [
        "nsDir",
        "nsCm",
        "ewDir",
        "ewCm",
        "largeShapeType",
        "largeAxisDirection",
        "lineLengthCm",
        "rectSide1Cm",
        "rectSide2Cm",
        "ellipseLongRadiusCm",
        "ellipseShortRadiusCm",
      ]);
      return `
      <tr class="${selectedClass}">
        <td class="${listCellClass("kuwaku-color-cell", kuwakuMissing)}" style="background:${kuwakuStyle.background};color:${kuwakuStyle.color};border-color:${kuwakuStyle.border};" ${missingTitle}>${escapeHtml(
          kuwakuText
        )}</td>
        <td class="${listCellClass("", teamMissing)}" ${missingTitle}>${escapeHtml(getRecordTeamValue(record))}</td>
        <td class="${listCellClass(dataComplete ? "data-status-complete" : "data-status-incomplete", !dataComplete)}" ${missingTitle}>${
          dataComplete ? "○" : "未記入"
        }</td>
        <td class="${listCellClass("", specimenMissing)}" ${missingTitle}>${escapeHtml(record.specimenNo)}</td>
        <td class="${listCellClass("category-color-cell", categoryMissing)}" style="background:${categoryBackground};color:#111827;border-color:${categoryBorderColor};" ${missingTitle}>${escapeHtml(
          formatCategoryForRecord(record)
        )}</td>
        <td class="${listCellClass("", nameMissing)}" ${missingTitle}>${escapeHtml(record.nameMemo || "")}</td>
        <td class="${listCellClass(record.importantFlag === "有" ? "important-cell-important" : "", importantMissing)}" ${missingTitle}>${escapeHtml(
          record.importantFlag || ""
        )}</td>
        <td class="${listCellClass("unit-color-cell", unitMissing)}" style="background:${unitStyle.background};color:${unitStyle.color};border-color:${unitStyle.border};" ${missingTitle}>${escapeHtml(
          record.unit || ""
        )}</td>
        <td>${escapeHtml(formatDetailForRecord(record))}</td>
        <td class="${listCellClass("", discovererMissing)}" ${missingTitle}>${escapeHtml(record.discoverer || "")}</td>
        <td class="${listCellClass("", identifierMissing)}" ${missingTitle}>${escapeHtml(record.identifier || "")}</td>
        <td class="${listCellClass("", levelReadMissing)}" ${missingTitle}>${escapeHtml(formatLevelRead(record))}</td>
        <td class="${listCellClass("", levelHeightMissing || levelReadMissing)}" ${missingTitle}>${escapeHtml(formatRecordAltitudeM(record))}</td>
        <td class="${listCellClass("", sectionMissing || sectionChecklistMissing)}" ${missingTitle}>${escapeHtml(
          record.occurrenceSection || ""
        )}</td>
        <td class="${listCellClass("", sketchMissing)}" ${missingTitle}>${escapeHtml(record.occurrenceSketch || "")}</td>
        <td class="${listCellClass("", positionMissing)}" ${missingTitle}>${escapeHtml(formatPlanPosition(record))}</td>
        <td>${escapeHtml(record.notes || "")}</td>
        <td>
          <div class="row-actions">
            <button type="button" data-action="card" data-id="${record.id}">${cardButtonLabel}</button>
            <button type="button" data-action="copy-to-input" data-id="${record.id}" data-kuwaku="${escapeHtml(
              getRecordKuwaku(record)
            )}">コピーして新規入力</button>
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

function listCellClass(baseClass, isMissing) {
  const classes = [];
  if (value(baseClass)) {
    classes.push(value(baseClass));
  }
  if (isMissing) {
    classes.push("missing-required-cell");
  }
  return classes.join(" ");
}

function hasAnyMissingRequiredKey(missingKeys, keys) {
  if (!(missingKeys instanceof Set)) {
    return false;
  }
  return keys.some((key) => missingKeys.has(key));
}

function formatMissingRequiredTooltip(missingKeys) {
  if (!(missingKeys instanceof Set) || missingKeys.size === 0) {
    return "";
  }
  const labels = Array.from(missingKeys).map((key) => REQUIRED_FIELD_LABELS[key] || key);
  return `title="${escapeHtml(`未記入: ${labels.join(" / ")}`)}"`;
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
    case "dataStatus":
      compared = compareSortText(dataStatusSortText(a), dataStatusSortText(b));
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
    case "altitudeM":
      compared = compareNullableNumber(getRecordAltitudeMValue(a), getRecordAltitudeMValue(b));
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

function compareNullableNumber(a, b) {
  const aValid = Number.isFinite(a);
  const bValid = Number.isFinite(b);
  if (aValid && bValid) {
    return a - b;
  }
  if (aValid) {
    return -1;
  }
  if (bValid) {
    return 1;
  }
  return 0;
}

function dataStatusSortText(record) {
  return isRecordDataComplete(record) ? "0-○" : "1-未記入";
}

function formatPlanPosition(record) {
  const nsDir = value(record?.nsDir);
  const nsCm = formatCmValue(record?.nsCm);
  const ewDir = value(record?.ewDir);
  const ewCm = formatCmValue(record?.ewCm);
  const nsPart = `${nsDir}${nsCm}`;
  const ewPart = `${ewDir}${ewCm}`;
  let base = "";
  if (nsPart && ewPart) {
    base = `${nsPart} / ${ewPart}`;
  } else {
    base = nsPart || ewPart;
  }
  if (normalizePlanSizeMode(record?.planSizeMode) !== "大きなもの") {
    return base;
  }
  const axisDirection = normalizeLargeAxisDirection(record?.largeAxisDirection);
  const shapeType = normalizeLargeShapeType(record?.largeShapeType);
  const lineLength = shapeType === "直線状" ? formatCmValue(record?.lineLengthCm) : "";
  if (!base) {
    if (shapeType === "直線状") {
      if (axisDirection && lineLength) {
        return `方位:${axisDirection} / 長さ:${lineLength}`;
      }
      return axisDirection ? `方位:${axisDirection}` : lineLength ? `長さ:${lineLength}` : "";
    }
    return axisDirection ? `方位:${axisDirection}` : "";
  }
  if (shapeType === "直線状") {
    if (axisDirection && lineLength) {
      return `${base} / 方位:${axisDirection} / 長さ:${lineLength}`;
    }
    if (axisDirection) {
      return `${base} / 方位:${axisDirection}`;
    }
    if (lineLength) {
      return `${base} / 長さ:${lineLength}`;
    }
  }
  return axisDirection ? `${base} / 方位:${axisDirection}` : base;
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
        <div><span>サブユニット</span><strong>${escapeHtml(formatDetailForRecord(selectedRecord))}</strong></div>
        <div><span>地層中の位置</span><strong>${escapeHtml(formatLayerPosition(selectedRecord))}</strong></div>
        <div><span>発見者</span><strong>${escapeHtml(selectedRecord.discoverer || "")}</strong></div>
        <div><span>判定者</span><strong>${escapeHtml(selectedRecord.identifier || "")}</strong></div>
        <div><span>レベル読値(上面/下底)</span><strong>${escapeHtml(formatLevelRead(selectedRecord))}</strong></div>
        <div><span>産出状況断面</span><strong>${escapeHtml(selectedRecord.occurrenceSection || "")}</strong></div>
        <div><span>産状スケッチ</span><strong>${escapeHtml(selectedRecord.occurrenceSketch || "")}</strong></div>
        <div><span>平面位置</span><strong>${escapeHtml(formatPlanPosition(selectedRecord))}</strong></div>
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
  if (!planMapWrap || !planMapLegend || !planUnitSelect || !planDetailSelect || !planDetailSubSelect) {
    return;
  }
  planMapLegend.innerHTML = buildPlanLegendHtml();

  if (!state.records.length) {
    selectedPlanKuwaku = "";
    selectedPlanUnit = "";
    selectedPlanDetail = ALL_DETAILS_VALUE;
    selectedPlanDetailSub = ALL_DETAIL_SUBS_VALUE;
    syncPlanKuwakuSelect([]);
    planUnitSelect.innerHTML = "";
    planDetailSelect.innerHTML = "";
    planDetailSubSelect.innerHTML = "";
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
    selectedPlanDetailSub = ALL_DETAIL_SUBS_VALUE;
    planUnitSelect.innerHTML = "";
    planDetailSelect.innerHTML = "";
    planDetailSubSelect.innerHTML = "";
    const kuwakuLabelForMeta = selectedPlanKuwaku ? kuwakuLabelForSelect(selectedPlanKuwaku) : "-";
    planMapWrap.innerHTML = `
      <div class="plan-map-meta">
        <span>区画（グリッド）: ${escapeHtml(kuwakuLabelForMeta)}</span>
        <span>出力階層: -</span>
        <span>表示件数: 0件</span>
      </div>
      <p class="muted">この区画（グリッド）には表示対象データがありません。</p>
    `;
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
      : unitRecords.filter((record) => detailValueForSelect(record.detail) === selectedPlanDetail);

  const detailSubs = collectPlanDetailSubs(detailRecords);
  if (!detailSubs.some((detailSub) => detailSub.value === selectedPlanDetailSub)) {
    selectedPlanDetailSub = detailSubs[0].value;
  }
  planDetailSubSelect.innerHTML = detailSubs
    .map(
      (detailSub) =>
        `<option value="${escapeHtml(detailSub.value)}" ${detailSub.value === selectedPlanDetailSub ? "selected" : ""}>${escapeHtml(
          detailSub.label
        )}</option>`
    )
    .join("");

  const detailSubRecords =
    selectedPlanDetailSub === ALL_DETAIL_SUBS_VALUE
      ? detailRecords
      : detailRecords.filter((record) => detailSubValueForSelect(record.detailSub) === selectedPlanDetailSub);
  const drawables = detailSubRecords.map((record) => buildPlanDrawable(record)).filter(Boolean);
  const kuwakuLabelForMeta = selectedPlanKuwaku ? kuwakuLabelForSelect(selectedPlanKuwaku) : "-";
  const unitLabelForMeta = selectedPlanUnit === ALL_UNITS_VALUE ? "全ユニット" : unitLabelForSelect(selectedPlanUnit);
  const detailLabelForMeta =
    selectedPlanDetail === ALL_DETAILS_VALUE ? "全サブユニット" : detailLabelForSelect(selectedPlanDetail);
  const detailSubLabelForMeta =
    selectedPlanDetailSub === ALL_DETAIL_SUBS_VALUE ? "全細分" : detailSubLabelForSelect(selectedPlanDetailSub);
  const hierarchyLabelForMeta = `${unitLabelForMeta} > ${detailLabelForMeta} > ${detailSubLabelForMeta}`;
  const mapMetaHtml = `
    <div class="plan-map-meta">
      <span>区画（グリッド）: ${escapeHtml(kuwakuLabelForMeta)}</span>
      <span>出力階層: ${escapeHtml(hierarchyLabelForMeta)}</span>
      <span>表示件数: ${detailSubRecords.length}件</span>
    </div>
  `;

  if (!drawables.length) {
    planMapWrap.innerHTML = `
      ${mapMetaHtml}
      <p class="muted">このユニット/サブユニット/細分は、平面位置の入力が不足しているため表示できません。</p>
    `;
    return;
  }

  const verticalGrid = [100, 200, 300]
    .map((x) => `<line class="plan-grid-line" x1="${x}" y1="0" x2="${x}" y2="${PLAN_SIZE_CM}" />`)
    .join("");
  const horizontalGrid = [100, 200, 300]
    .map((y) => `<line class="plan-grid-line" x1="0" y1="${y}" x2="${PLAN_SIZE_CM}" y2="${y}" />`)
    .join("");
  const pointsSvg = drawables.map((drawable, index) => renderPlanDrawableSvg(drawable, index)).join("");
  const cornerLabels = buildPlanCornerLabels(selectedPlanKuwaku);

  planMapWrap.innerHTML = `
    ${mapMetaHtml}
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
        <rect class="plan-frame" x="0" y="0" width="${PLAN_SIZE_CM}" height="${PLAN_SIZE_CM}" />
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
  const detailSet = new Set(records.map((record) => detailValueForSelect(record.detail)));
  const detailOptions = Array.from(detailSet)
    .sort((a, b) => detailLabelForSelect(a).localeCompare(detailLabelForSelect(b), "ja", { numeric: true, sensitivity: "base" }))
    .map((detailValue) => ({
      value: detailValue,
      label: detailLabelForSelect(detailValue),
    }));
  return [{ value: ALL_DETAILS_VALUE, label: "全サブユニット" }, ...detailOptions];
}

function collectPlanDetailSubs(records) {
  const detailSubSet = new Set(records.map((record) => detailSubValueForSelect(record.detailSub)));
  const detailSubOptions = Array.from(detailSubSet)
    .sort((a, b) =>
      detailSubLabelForSelect(a).localeCompare(detailSubLabelForSelect(b), "ja", { numeric: true, sensitivity: "base" })
    )
    .map((detailSubValue) => ({
      value: detailSubValue,
      label: detailSubLabelForSelect(detailSubValue),
    }));
  return [{ value: ALL_DETAIL_SUBS_VALUE, label: "全細分" }, ...detailSubOptions];
}

function normalizeViewerVerticalScale(scaleRaw) {
  const num = Number(value(scaleRaw));
  if (!Number.isFinite(num)) {
    return 1;
  }
  return clamp(num, 1, 5);
}

function syncViewerVerticalScaleUi() {
  viewerVerticalScale = normalizeViewerVerticalScale(viewerVerticalScale);
  if (viewerZScaleInput) {
    viewerZScaleInput.value = String(viewerVerticalScale);
  }
  if (viewerZScaleValue) {
    viewerZScaleValue.textContent = `${viewerVerticalScale.toFixed(1)}x`;
  }
}

function applyViewerVerticalScale(zRaw, baseZRaw) {
  const z = Number(zRaw);
  const baseZ = Number(baseZRaw);
  if (!Number.isFinite(z) || !Number.isFinite(baseZ)) {
    return z;
  }
  return baseZ + (z - baseZ) * viewerVerticalScale;
}

function renderViewerOutput() {
  if (
    !viewerKuwakuSelect ||
    !viewerUnitSelect ||
    !viewerDetailSelect ||
    !viewerDetailSubSelect ||
    !viewerMapLegend ||
    !viewerStatus
  ) {
    return;
  }
  viewerMapLegend.innerHTML = buildPlanLegendHtml();
  syncViewerVerticalScaleUi();
  syncViewerViewButtons();

  const sortedRecords = [...state.records].sort(compareRecordsByKuwakuThenSpecimen);
  syncViewerKuwakuSelect(sortedRecords);

  if (!sortedRecords.length) {
    selectedViewerUnit = ALL_UNITS_VALUE;
    selectedViewerDetail = ALL_DETAILS_VALUE;
    selectedViewerDetailSub = ALL_DETAIL_SUBS_VALUE;
    viewerUnitSelect.innerHTML = "";
    viewerDetailSelect.innerHTML = "";
    viewerDetailSubSelect.innerHTML = "";
    if (viewerKuwakuInfo) {
      viewerKuwakuInfo.textContent = "区画: -";
    }
    viewerStatus.textContent = "表示対象データがありません。";
    clearViewerScene();
    return;
  }

  const kuwakuRecords =
    selectedViewerKuwaku === ALL_GRIDS_VALUE
      ? sortedRecords
      : sortedRecords.filter((record) => kuwakuValueForSelect(getRecordKuwaku(record)) === selectedViewerKuwaku);

  const kuwakuLabel = selectedViewerKuwaku === ALL_GRIDS_VALUE ? "全グリッド" : kuwakuLabelForSelect(selectedViewerKuwaku);
  if (viewerKuwakuInfo) {
    viewerKuwakuInfo.textContent = `区画: ${kuwakuLabel}`;
  }

  if (!kuwakuRecords.length) {
    selectedViewerUnit = ALL_UNITS_VALUE;
    selectedViewerDetail = ALL_DETAILS_VALUE;
    selectedViewerDetailSub = ALL_DETAIL_SUBS_VALUE;
    viewerUnitSelect.innerHTML = "";
    viewerDetailSelect.innerHTML = "";
    viewerDetailSubSelect.innerHTML = "";
    viewerStatus.textContent = "この条件では表示対象データがありません。";
    clearViewerScene();
    return;
  }

  const units = collectPlanUnits(kuwakuRecords);
  if (!units.some((unit) => unit.value === selectedViewerUnit)) {
    selectedViewerUnit = units[0]?.value || ALL_UNITS_VALUE;
  }
  viewerUnitSelect.innerHTML = units
    .map(
      (unit) =>
        `<option value="${escapeHtml(unit.value)}" ${unit.value === selectedViewerUnit ? "selected" : ""}>${escapeHtml(
          unit.label
        )}</option>`
    )
    .join("");
  const unitRecords =
    selectedViewerUnit === ALL_UNITS_VALUE
      ? kuwakuRecords
      : kuwakuRecords.filter((record) => unitValueForSelect(record.unit) === selectedViewerUnit);

  const details = collectPlanDetails(unitRecords);
  if (!details.some((detail) => detail.value === selectedViewerDetail)) {
    selectedViewerDetail = details[0]?.value || ALL_DETAILS_VALUE;
  }
  viewerDetailSelect.innerHTML = details
    .map(
      (detail) =>
        `<option value="${escapeHtml(detail.value)}" ${detail.value === selectedViewerDetail ? "selected" : ""}>${escapeHtml(
          detail.label
        )}</option>`
    )
    .join("");
  const detailRecords =
    selectedViewerDetail === ALL_DETAILS_VALUE
      ? unitRecords
      : unitRecords.filter((record) => detailValueForSelect(record.detail) === selectedViewerDetail);

  const detailSubs = collectPlanDetailSubs(detailRecords);
  if (!detailSubs.some((detailSub) => detailSub.value === selectedViewerDetailSub)) {
    selectedViewerDetailSub = detailSubs[0]?.value || ALL_DETAIL_SUBS_VALUE;
  }
  viewerDetailSubSelect.innerHTML = detailSubs
    .map(
      (detailSub) =>
        `<option value="${escapeHtml(detailSub.value)}" ${detailSub.value === selectedViewerDetailSub ? "selected" : ""}>${escapeHtml(
          detailSub.label
        )}</option>`
    )
    .join("");
  const detailSubRecords =
    selectedViewerDetailSub === ALL_DETAIL_SUBS_VALUE
      ? detailRecords
      : detailRecords.filter((record) => detailSubValueForSelect(record.detailSub) === selectedViewerDetailSub);

  const viewerCandidates = [];
  let missingPositionCount = 0;
  let missingAltitudeCount = 0;
  for (const record of detailSubRecords) {
    const drawable = buildPlanDrawable(record);
    if (!drawable) {
      missingPositionCount += 1;
      continue;
    }
    let altitudeM = getRecordAltitudeMValue(record);
    let altitudeEstimated = false;
    if (altitudeM == null) {
      missingAltitudeCount += 1;
      altitudeM = VIEWER_ALTITUDE_BASE_M;
      altitudeEstimated = true;
    }
    const kuwaku = parseKuwaku(getRecordKuwaku(record));
    const xIndex = kuwakuToViewerXIndex(kuwaku);
    const noIndex = parseGridNoToIndex(kuwaku.no);
    viewerCandidates.push({
      record,
      drawable,
      altitudeM,
      altitudeEstimated,
      grid: {
        kuwaku: buildKuwaku(kuwaku.headA, kuwaku.headB, kuwaku.block, kuwaku.no),
        headB: kuwaku.headB,
        block: kuwaku.block,
        no: kuwaku.no,
        xIndex,
        noIndex,
      },
    });
  }

  viewerStatus.textContent = `対象 ${detailSubRecords.length}件 / 3D表示 ${viewerCandidates.length}件 / 平面位置未記入 ${missingPositionCount}件 / 標高未記入 ${missingAltitudeCount}件（655mで仮表示） / 縦スケール ${viewerVerticalScale.toFixed(1)}x`;

  if (!viewerCandidates.length) {
    clearViewerScene();
    return;
  }

  if (!isViewerTabActive() && !viewer3d.initialized) {
    return;
  }
  if (!ensureViewerInitialized()) {
    return;
  }

  const metrics = buildViewerGridMetrics(viewerCandidates);
  const shapes = viewerCandidates.map((candidate) => buildViewerShapeFromCandidate(candidate, metrics)).filter(Boolean);
  renderViewerScene(shapes, metrics);
}

function isViewerTabActive() {
  return getActiveTabId() === "viewer-tab";
}

function syncViewerKuwakuSelect(records) {
  if (!viewerKuwakuSelect) {
    return;
  }
  const options = collectOutputKuwakuOptions(records);
  if (!options.some((item) => item.value === selectedViewerKuwaku)) {
    selectedViewerKuwaku = ALL_GRIDS_VALUE;
  }
  viewerKuwakuSelect.innerHTML = options
    .map(
      (item) =>
        `<option value="${escapeHtml(item.value)}" ${item.value === selectedViewerKuwaku ? "selected" : ""}>${escapeHtml(
          item.label
        )}</option>`
    )
    .join("");
}

function parseGridNoToIndex(noRaw) {
  const no = value(noRaw);
  if (/^-?\d+$/.test(no)) {
    return Number(no);
  }
  if (!no) {
    return 0;
  }
  return hashText(no) % 100;
}

function buildViewerGridMetrics(candidates) {
  const xIndexes = candidates.map((item) => item.grid.xIndex).filter((num) => Number.isFinite(num));
  const noIndexes = candidates.map((item) => item.grid.noIndex).filter((num) => Number.isFinite(num));
  const altitudes = candidates.map((item) => item.altitudeM).filter((num) => Number.isFinite(num));
  const minXIndex = xIndexes.length ? Math.min(...xIndexes) : 0;
  let maxXIndex = xIndexes.length ? Math.max(...xIndexes) : minXIndex;
  const presentHeads = new Set(candidates.map((item) => normalizeViewerHead(item?.grid?.headB)));
  presentHeads.forEach((head) => {
    const headIndex = VIEWER_HEAD_INDEX_MAP.get(head);
    if (!Number.isFinite(headIndex)) {
      return;
    }
    const fIndex = headIndex * 26 + 5;
    if (fIndex > maxXIndex) {
      maxXIndex = fIndex;
    }
  });
  const minNo = noIndexes.length ? Math.min(...noIndexes) : 0;
  const maxNo = noIndexes.length ? Math.max(...noIndexes) : minNo;
  const minZ = VIEWER_ALTITUDE_BASE_M;
  const observedMaxZ = altitudes.length ? Math.max(...altitudes) : minZ + 1;
  const maxZ = Math.max(minZ + 1, Math.ceil(observedMaxZ));
  return {
    minXIndex,
    maxXIndex,
    minNo,
    maxNo,
    minZ,
    maxZ,
    gridWidthM: Math.max(4, (maxXIndex - minXIndex + 1) * 4),
    gridHeightM: Math.max(4, (maxNo - minNo + 1) * 4),
  };
}

function buildViewerShapeFromCandidate(candidate, metrics) {
  const drawable = candidate.drawable;
  const altitudeM = candidate.altitudeM;
  const altitudeZ = applyViewerVerticalScale(altitudeM, metrics.minZ);
  const worldCenter = convertViewerPointCmToWorld(drawable.x, drawable.y, candidate.grid, metrics);
  const meta = {
    id: value(candidate.record.id),
    label: value(candidate.record.specimenNo),
    nameMemo: value(candidate.record.nameMemo),
    unit: value(candidate.record.unit),
    detail: buildDetailText(candidate.record.detail, candidate.record.detailSub),
    kuwaku: value(candidate.grid.kuwaku),
    altitudeM,
    altitudeEstimated: Boolean(candidate.altitudeEstimated),
    color: drawable.color,
  };
  if (drawable.type === "point") {
    return {
      type: "point",
      x: worldCenter.x,
      y: worldCenter.y,
      z: altitudeZ,
      ...meta,
    };
  }

  if (drawable.type === "line") {
    const p1 = convertViewerPointCmToWorld(drawable.x1, drawable.y1, candidate.grid, metrics);
    const p2 = convertViewerPointCmToWorld(drawable.x2, drawable.y2, candidate.grid, metrics);
    return {
      type: "line",
      points: [
        { x: p1.x, y: p1.y, z: altitudeZ },
        { x: p2.x, y: p2.y, z: altitudeZ },
      ],
      x: worldCenter.x,
      y: worldCenter.y,
      z: altitudeZ,
      ...meta,
    };
  }

  if (drawable.type === "rect") {
    const halfW = drawable.width / 2;
    const halfH = drawable.height / 2;
    const localCorners = [
      { x: drawable.x - halfW, y: drawable.y - halfH },
      { x: drawable.x + halfW, y: drawable.y - halfH },
      { x: drawable.x + halfW, y: drawable.y + halfH },
      { x: drawable.x - halfW, y: drawable.y + halfH },
    ].map((point) => rotatePlanPoint(point, { x: drawable.x, y: drawable.y }, drawable.rotationDeg));
    const points = localCorners.map((point) => {
      const world = convertViewerPointCmToWorld(point.x, point.y, candidate.grid, metrics);
      return { x: world.x, y: world.y, z: altitudeZ };
    });
    points.push(points[0]);
    return {
      type: "polyline",
      points,
      x: worldCenter.x,
      y: worldCenter.y,
      z: altitudeZ,
      ...meta,
    };
  }

  if (drawable.type === "ellipse") {
    const points = [];
    const segmentCount = 48;
    for (let i = 0; i <= segmentCount; i += 1) {
      const theta = (i / segmentCount) * Math.PI * 2;
      const local = {
        x: drawable.x + Math.cos(theta) * drawable.rx,
        y: drawable.y + Math.sin(theta) * drawable.ry,
      };
      const rotated = rotatePlanPoint(local, { x: drawable.x, y: drawable.y }, drawable.rotationDeg);
      const world = convertViewerPointCmToWorld(rotated.x, rotated.y, candidate.grid, metrics);
      points.push({ x: world.x, y: world.y, z: altitudeZ });
    }
    return {
      type: "polyline",
      points,
      x: worldCenter.x,
      y: worldCenter.y,
      z: altitudeZ,
      ...meta,
    };
  }

  if (drawable.type === "imageQuad") {
    const points = (drawable.points || []).map((point) => {
      const world = convertViewerPointCmToWorld(point.x, point.y, candidate.grid, metrics);
      return { x: world.x, y: world.y, z: altitudeZ };
    });
    return {
      type: "imageQuad",
      points,
      imageType: value(drawable.imageType),
      imagePath: value(drawable.imagePath),
      x: worldCenter.x,
      y: worldCenter.y,
      z: altitudeZ,
      ...meta,
    };
  }
  return null;
}

function rotatePlanPoint(point, center, rotationDegRaw) {
  const rotationDeg = Number(rotationDegRaw);
  if (!Number.isFinite(rotationDeg) || Math.abs(rotationDeg) < 1e-6) {
    return { x: point.x, y: point.y };
  }
  const rad = (rotationDeg * Math.PI) / 180;
  const cos = Math.cos(rad);
  const sin = Math.sin(rad);
  const dx = point.x - center.x;
  const dy = point.y - center.y;
  return {
    x: center.x + dx * cos - dy * sin,
    y: center.y + dx * sin + dy * cos,
  };
}

function convertViewerPointCmToWorld(xCmRaw, yCmRaw, grid, metrics) {
  const xCm = Number(xCmRaw);
  const yCm = Number(yCmRaw);
  const xIndex = Number(grid?.xIndex);
  const noIndex = Number(grid?.noIndex);
  const baseEast = (xIndex - metrics.minXIndex) * 4;
  const baseNorth = (metrics.maxNo - noIndex) * 4;
  return {
    x: baseEast + xCm / 100,
    y: baseNorth + (PLAN_SIZE_CM - yCm) / 100,
  };
}

function syncViewerViewButtons() {
  if (viewerViewTopBtn) {
    viewerViewTopBtn.classList.toggle("active", selectedViewerPerspective === "top");
  }
  if (viewerViewIsoBtn) {
    viewerViewIsoBtn.classList.toggle("active", selectedViewerPerspective === "iso");
  }
}

function ensureViewerInitialized() {
  if (!viewerCanvasWrap || viewer3d.initialized) {
    return viewer3d.initialized;
  }
  if (!window.THREE) {
    if (viewerStatus) {
      viewerStatus.textContent = "3Dライブラリの読み込みに失敗しました。";
    }
    return false;
  }
  try {
    viewer3d.scene = new THREE.Scene();
    viewer3d.scene.background = new THREE.Color(0xf8fafc);
    viewer3d.camera = new THREE.PerspectiveCamera(50, 1, 0.1, 10000);
    viewer3d.renderer = new THREE.WebGLRenderer({ antialias: true, alpha: true });
    viewer3d.renderer.setPixelRatio(Math.min(window.devicePixelRatio || 1, 2));
    viewer3d.renderer.domElement.setAttribute("aria-label", "3Dビュアー");
    viewerCanvasWrap.prepend(viewer3d.renderer.domElement);

    viewer3d.controls =
      typeof THREE.OrbitControls === "function"
        ? new THREE.OrbitControls(viewer3d.camera, viewer3d.renderer.domElement)
        : null;
    if (viewer3d.controls) {
      viewer3d.controls.enableDamping = true;
      viewer3d.controls.enablePan = true;
      viewer3d.controls.zoomSpeed = 0.45;
      viewer3d.controls.target.set(0, 0, 0);
    }

    const ambient = new THREE.AmbientLight(0xffffff, 0.88);
    const directional = new THREE.DirectionalLight(0xffffff, 0.55);
    directional.position.set(8, -6, 14);
    viewer3d.scene.add(ambient);
    viewer3d.scene.add(directional);

    viewer3d.gridGroup = new THREE.Group();
    viewer3d.dataGroup = new THREE.Group();
    viewer3d.labelGroup = new THREE.Group();
    viewer3d.scene.add(viewer3d.gridGroup);
    viewer3d.scene.add(viewer3d.dataGroup);
    viewer3d.scene.add(viewer3d.labelGroup);
    viewer3d.raycaster = new THREE.Raycaster();
    viewer3d.pointer = new THREE.Vector2();
    viewer3d.available = true;
    viewer3d.initialized = true;

    viewerCanvasWrap.addEventListener("pointermove", handleViewerPointerMove);
    viewerCanvasWrap.addEventListener("pointerleave", hideViewerTooltip);
    if (typeof ResizeObserver === "function") {
      viewer3d.resizeObserver = new ResizeObserver(() => {
        ensureViewerCanvasSize();
      });
      viewer3d.resizeObserver.observe(viewerCanvasWrap);
    } else {
      window.addEventListener("resize", ensureViewerCanvasSize);
    }
    ensureViewerCanvasSize();
    animateViewerScene();
    return true;
  } catch (_error) {
    viewer3d.available = false;
    viewer3d.initialized = false;
    if (viewerStatus) {
      viewerStatus.textContent = "3D表示の初期化に失敗しました。";
    }
    return false;
  }
}

function ensureViewerCanvasSize() {
  if (!viewer3d.initialized || !viewerCanvasWrap || !viewer3d.renderer || !viewer3d.camera) {
    return;
  }
  const rect = viewerCanvasWrap.getBoundingClientRect();
  const width = Math.max(16, Math.floor(rect.width));
  const height = Math.max(16, Math.floor(rect.height));
  viewer3d.renderer.setSize(width, height, false);
  viewer3d.camera.aspect = width / height;
  viewer3d.camera.updateProjectionMatrix();
}

function animateViewerScene() {
  if (!viewer3d.initialized || !viewer3d.renderer || !viewer3d.scene || !viewer3d.camera) {
    return;
  }
  viewer3d.frameHandle = window.requestAnimationFrame(animateViewerScene);
  if (viewer3d.controls) {
    viewer3d.controls.update();
  }
  viewer3d.renderer.render(viewer3d.scene, viewer3d.camera);
}

function clearViewerScene() {
  if (!viewer3d.initialized || !viewer3d.dataGroup || !viewer3d.labelGroup || !viewer3d.gridGroup) {
    return;
  }
  viewer3d.dataGroup.children.forEach((child) => disposeViewerObject3D(child));
  viewer3d.labelGroup.children.forEach((child) => disposeViewerObject3D(child));
  viewer3d.gridGroup.children.forEach((child) => disposeViewerObject3D(child));
  viewer3d.dataGroup.clear();
  viewer3d.labelGroup.clear();
  viewer3d.gridGroup.clear();
  viewer3d.pickMeshes = [];
  viewer3d.meshesByRecordId = new Map();
  hideViewerTooltip();
}

function disposeViewerObject3D(object) {
  if (!object) {
    return;
  }
  object.traverse?.((child) => {
    if (child.geometry?.dispose) {
      child.geometry.dispose();
    }
    if (child.material) {
      const materials = Array.isArray(child.material) ? child.material : [child.material];
      materials.forEach((material) => {
        if (material.map?.dispose) {
          material.map.dispose();
        }
        if (material.dispose) {
          material.dispose();
        }
      });
    }
  });
}

function renderViewerScene(shapes, metrics) {
  if (!viewer3d.initialized || !viewer3d.dataGroup || !viewer3d.labelGroup || !viewer3d.gridGroup) {
    return;
  }
  clearViewerScene();
  renderViewerGrid(metrics);
  viewer3d.renderNonce += 1;
  const renderNonce = viewer3d.renderNonce;

  shapes.forEach((shape) => {
    const color = new THREE.Color(shape.color || "#6b7280");
    if (shape.type === "point") {
      const pointMesh = new THREE.Mesh(
        new THREE.SphereGeometry(0.11, 14, 14),
        new THREE.MeshBasicMaterial({ color })
      );
      pointMesh.position.set(shape.x, shape.y, shape.z);
      viewer3d.dataGroup.add(pointMesh);
    } else if (shape.type === "line") {
      renderViewerSegment(shape.points[0], shape.points[1], color, 0.05);
    } else if (shape.type === "polyline") {
      for (let i = 0; i < shape.points.length - 1; i += 1) {
        renderViewerSegment(shape.points[i], shape.points[i + 1], color, 0.04);
      }
    } else if (shape.type === "imageQuad") {
      renderViewerImageQuad(shape, renderNonce);
    }

    const label = createViewerTextSprite(shape.label || "-", shape.color);
    label.position.set(shape.x, shape.y, shape.z + 0.16);
    viewer3d.labelGroup.add(label);

    const pickMesh = new THREE.Mesh(
      new THREE.SphereGeometry(0.2, 10, 10),
      new THREE.MeshBasicMaterial({ transparent: true, opacity: 0.001, depthWrite: false })
    );
    pickMesh.position.set(shape.x, shape.y, shape.z);
    pickMesh.userData = {
      label: shape.label,
      nameMemo: shape.nameMemo,
      unit: shape.unit,
      detail: shape.detail,
      kuwaku: shape.kuwaku,
      altitudeM: shape.altitudeM,
      altitudeEstimated: Boolean(shape.altitudeEstimated),
    };
    viewer3d.pickMeshes.push(pickMesh);
    viewer3d.dataGroup.add(pickMesh);
    viewer3d.meshesByRecordId.set(shape.id, shape);
  });

  viewer3d.bounds = computeViewerBounds(shapes, metrics);
  applyViewerPerspective();
}

function renderViewerGrid(metrics) {
  if (!viewer3d.gridGroup) {
    return;
  }
  const axisMinAltitude = Math.floor(metrics.minZ);
  const axisMaxAltitude = Math.max(axisMinAltitude + 1, Math.ceil(metrics.maxZ));
  const zBase = applyViewerVerticalScale(axisMinAltitude, metrics.minZ) - 0.05;
  const width = metrics.gridWidthM;
  const height = metrics.gridHeightM;
  const gridMat = new THREE.LineBasicMaterial({ color: 0xcbd5e1 });

  for (let x = 0; x <= width + 0.001; x += 4) {
    const geometry = new THREE.BufferGeometry().setFromPoints([
      new THREE.Vector3(x, 0, zBase),
      new THREE.Vector3(x, height, zBase),
    ]);
    viewer3d.gridGroup.add(new THREE.Line(geometry, gridMat));
  }
  for (let y = 0; y <= height + 0.001; y += 4) {
    const geometry = new THREE.BufferGeometry().setFromPoints([
      new THREE.Vector3(0, y, zBase),
      new THREE.Vector3(width, y, zBase),
    ]);
    viewer3d.gridGroup.add(new THREE.Line(geometry, gridMat));
  }

  for (let xIndex = metrics.minXIndex; xIndex <= metrics.maxXIndex; xIndex += 1) {
    const headBlock = viewerXIndexToHeadBlock(xIndex);
    for (let no = metrics.minNo; no <= metrics.maxNo; no += 1) {
      const label = `${headBlock.head}-${headBlock.block}-${no}`;
      const x = (xIndex - metrics.minXIndex) * 4 + 0.45;
      const y = (metrics.maxNo - no) * 4 + 3.55;
      const sprite = createViewerTextSprite(label, "#334155", {
        fontPx: 124,
        scaleX: 2.24,
        scaleY: 0.56,
      });
      sprite.position.set(x, y, zBase + 0.01);
      viewer3d.gridGroup.add(sprite);
    }
  }

  const zStart = applyViewerVerticalScale(axisMinAltitude, metrics.minZ);
  const zEnd = applyViewerVerticalScale(axisMaxAltitude, metrics.minZ);
  const zAxisMat = new THREE.LineBasicMaterial({ color: 0x334155 });
  const cornerAxes = [
    { x: 0, y: 0, dirX: -1, dirY: -1 },
    { x: width, y: 0, dirX: 1, dirY: -1 },
    { x: 0, y: height, dirX: -1, dirY: 1 },
    { x: width, y: height, dirX: 1, dirY: 1 },
  ];
  cornerAxes.forEach((corner) => {
    const zAxisGeom = new THREE.BufferGeometry().setFromPoints([
      new THREE.Vector3(corner.x, corner.y, zStart),
      new THREE.Vector3(corner.x, corner.y, zEnd),
    ]);
    viewer3d.gridGroup.add(new THREE.Line(zAxisGeom, zAxisMat));

    for (let altitude = axisMinAltitude; altitude <= axisMaxAltitude; altitude += 1) {
      const z = applyViewerVerticalScale(altitude, metrics.minZ);
      const tickGeom = new THREE.BufferGeometry().setFromPoints([
        new THREE.Vector3(corner.x, corner.y, z),
        new THREE.Vector3(corner.x + corner.dirX * 0.14, corner.y + corner.dirY * 0.14, z),
      ]);
      viewer3d.gridGroup.add(new THREE.Line(tickGeom, zAxisMat));
      const label = createViewerTextSprite(`${altitude}m`, "#334155", {
        fontPx: 87,
        scaleX: 1.57,
        scaleY: 0.39,
      });
      label.position.set(corner.x + corner.dirX * 0.62, corner.y + corner.dirY * 0.62, z);
      viewer3d.gridGroup.add(label);
    }
    const axisLabel = createViewerTextSprite("標高(m)", "#1e293b", {
      fontPx: 87,
      scaleX: 1.57,
      scaleY: 0.39,
    });
    axisLabel.position.set(corner.x + corner.dirX * 0.68, corner.y + corner.dirY * 0.68, zEnd + 0.08);
    viewer3d.gridGroup.add(axisLabel);
  });
}

function renderViewerSegment(start, end, color, radius = 0.04) {
  if (!viewer3d.dataGroup) {
    return;
  }
  const startVec = new THREE.Vector3(start.x, start.y, start.z);
  const endVec = new THREE.Vector3(end.x, end.y, end.z);
  const diff = new THREE.Vector3().subVectors(endVec, startVec);
  const length = diff.length();
  if (!Number.isFinite(length) || length <= 0.0001) {
    return;
  }
  const geometry = new THREE.CylinderGeometry(radius, radius, length, 8);
  const material = new THREE.MeshBasicMaterial({ color });
  const mesh = new THREE.Mesh(geometry, material);
  const mid = new THREE.Vector3().addVectors(startVec, endVec).multiplyScalar(0.5);
  mesh.position.copy(mid);
  const axis = new THREE.Vector3(0, 1, 0);
  mesh.quaternion.setFromUnitVectors(axis, diff.clone().normalize());
  viewer3d.dataGroup.add(mesh);
}

function renderViewerImageQuad(shape, renderNonce) {
  if (!viewer3d.dataGroup) {
    return;
  }
  const targetGroup = viewer3d.dataGroup;
  const points = Array.isArray(shape?.points) ? shape.points : [];
  const imagePath = value(shape?.imagePath) || getLargeShapeImagePath(shape?.imageType);
  if (points.length !== 4) {
    return;
  }
  // 画像が読めない場合でも位置が分かるように、外形線は先に描画する。
  for (let i = 0; i < points.length; i += 1) {
    const start = points[i];
    const end = points[(i + 1) % points.length];
    renderViewerSegment(start, end, new THREE.Color(shape?.color || "#6b7280"), 0.012);
  }
  if (!imagePath) {
    return;
  }
  const buildGeometry = () => {
    const segments = 28;
    const geometry = new THREE.BufferGeometry();
    const positions = [];
    const uvs = [];
    const indices = [];
    for (let y = 0; y <= segments; y += 1) {
      const v = 1 - y / segments;
      for (let x = 0; x <= segments; x += 1) {
        const u = x / segments;
        const p = interpolateViewerQuadPoint(points, u, v);
        positions.push(p.x, p.y, p.z);
        uvs.push(u, v);
      }
    }
    for (let y = 0; y < segments; y += 1) {
      for (let x = 0; x < segments; x += 1) {
        const row = y * (segments + 1);
        const nextRow = (y + 1) * (segments + 1);
        const a = row + x;
        const b = row + x + 1;
        const c = nextRow + x + 1;
        const d = nextRow + x;
        indices.push(a, b, d, b, c, d);
      }
    }
    const vertices = new Float32Array(positions);
    const uvArray = new Float32Array(uvs);
    geometry.setAttribute("position", new THREE.BufferAttribute(vertices, 3));
    geometry.setAttribute("uv", new THREE.BufferAttribute(uvArray, 2));
    geometry.setIndex(indices);
    geometry.computeVertexNormals();
    return geometry;
  };
  const addTexturedMesh = (texture, tintHex = "#ffffff") => {
    if (!viewer3d.initialized || !targetGroup || !targetGroup.parent) {
      return;
    }
    texture.needsUpdate = true;
    texture.minFilter = THREE.NearestFilter;
    texture.magFilter = THREE.NearestFilter;
    texture.generateMipmaps = false;
    // 上面視点（平面図と同じ向き）での貼り付けに合わせる。
    texture.flipY = false;
    if ("colorSpace" in texture && window.THREE?.SRGBColorSpace) {
      texture.colorSpace = window.THREE.SRGBColorSpace;
    }
    const geometry = buildGeometry();
    const material = new THREE.MeshBasicMaterial({
      map: texture,
      color: new THREE.Color(tintHex),
      transparent: true,
      side: THREE.DoubleSide,
      depthWrite: false,
      depthTest: false,
      alphaTest: 0,
      opacity: 1,
      polygonOffset: true,
      polygonOffsetFactor: -1,
      polygonOffsetUnits: -1,
    });
    const mesh = new THREE.Mesh(geometry, material);
    mesh.renderOrder = 6;
    targetGroup.add(mesh);
  };
  const loadFromImagePath = () =>
    getOrLoadPlanLargeShapeTintedCanvas(imagePath, shape.color, shape?.imageType)
      .then((canvas) => {
        const texture = new THREE.CanvasTexture(canvas);
        addTexturedMesh(texture, "#ffffff");
        return true;
      })
      .catch(() =>
        getOrLoadPlanLargeShapeImage(imagePath, shape?.imageType)
          .then((image) => {
            const texture = new THREE.Texture(image);
            const tint = parseHexColor(shape?.color).hex;
            addTexturedMesh(texture, tint);
            return true;
          })
          .catch(() => false)
      );

  const tintedDataUrl = getPlanLargeShapeTintedDataUrl(imagePath, shape?.color);
  if (tintedDataUrl) {
    getOrLoadPlanLargeShapeImage(tintedDataUrl, shape?.imageType)
      .then((image) => {
        const texture = new THREE.Texture(image);
        addTexturedMesh(texture, "#ffffff");
      })
      .catch(() => {
        void loadFromImagePath();
      });
    return;
  }
  void loadFromImagePath();
}

function addViewerImageStrokeOverlay(points, imageSource, colorRaw, targetGroup, renderNonce) {
  if (!Array.isArray(points) || points.length !== 4 || !targetGroup || !targetGroup.parent || !viewer3d.initialized) {
    return;
  }
  const canvas = ensureCanvasFromImageSource(imageSource);
  if (!canvas) {
    return;
  }
  const width = Number(canvas.width);
  const height = Number(canvas.height);
  if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 1 || height <= 1) {
    return;
  }
  const ctx = canvas.getContext("2d", { willReadFrequently: true });
  if (!ctx) {
    return;
  }
  let imageData;
  try {
    imageData = ctx.getImageData(0, 0, width, height);
  } catch (_error) {
    return;
  }
  const data = imageData.data;
  const minDim = Math.max(1, Math.min(width, height));
  const stride = Math.max(1, Math.floor(minDim / 180));
  const maxPoints = 18000;
  const positions = [];
  const zOffset = 0.004;
  let useLineMask = false;
  for (let y = 0; y < height && !useLineMask; y += Math.max(1, stride * 2)) {
    for (let x = 0; x < width; x += Math.max(1, stride * 2)) {
      const idx = (y * width + x) * 4;
      const alpha = data[idx + 3];
      if (alpha < 1) {
        continue;
      }
      const r = data[idx];
      const g = data[idx + 1];
      const b = data[idx + 2];
      const isNearWhite = r >= 244 && g >= 244 && b >= 244;
      if (!isNearWhite) {
        useLineMask = true;
        break;
      }
    }
  }
  for (let y = 0; y < height; y += stride) {
    for (let x = 0; x < width; x += stride) {
      const index = (y * width + x) * 4;
      const alpha = data[index + 3];
      if (alpha < 1) {
        continue;
      }
      if (useLineMask) {
        const r = data[index];
        const g = data[index + 1];
        const b = data[index + 2];
        const isNearWhite = r >= 244 && g >= 244 && b >= 244;
        if (isNearWhite) {
          continue;
        }
      }
      const u = width <= 1 ? 0 : x / (width - 1);
      const v = height <= 1 ? 0 : y / (height - 1);
      const world = interpolateViewerQuadPoint(points, u, v);
      positions.push(world.x, world.y, world.z + zOffset);
      if (positions.length / 3 >= maxPoints) {
        break;
      }
    }
    if (positions.length / 3 >= maxPoints) {
      break;
    }
  }
  if (positions.length < 9) {
    return;
  }
  const geometry = new THREE.BufferGeometry();
  geometry.setAttribute("position", new THREE.Float32BufferAttribute(positions, 3));
  const material = new THREE.PointsMaterial({
    color: new THREE.Color(parseHexColor(colorRaw).hex),
    size: 3.2,
    sizeAttenuation: false,
    transparent: true,
    opacity: 0.98,
    depthTest: false,
    depthWrite: false,
  });
  const cloud = new THREE.Points(geometry, material);
  cloud.renderOrder = 7;
  targetGroup.add(cloud);
}

function ensureCanvasFromImageSource(imageSource) {
  if (!imageSource) {
    return null;
  }
  const isCanvas =
    typeof HTMLCanvasElement !== "undefined" && imageSource instanceof HTMLCanvasElement && Number(imageSource.width) > 0;
  if (isCanvas) {
    return imageSource;
  }
  const width = Math.max(1, Number(imageSource.naturalWidth || imageSource.width) || 0);
  const height = Math.max(1, Number(imageSource.naturalHeight || imageSource.height) || 0);
  if (!width || !height) {
    return null;
  }
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const ctx = canvas.getContext("2d", { willReadFrequently: true });
  if (!ctx) {
    return null;
  }
  ctx.drawImage(imageSource, 0, 0, width, height);
  return canvas;
}

function interpolateViewerQuadPoint(points, uRaw, vRaw) {
  const u = clampNumber(Number(uRaw), 0, 1);
  const v = clampNumber(Number(vRaw), 0, 1);
  const p0 = points[0];
  const p1 = points[1];
  const p2 = points[2];
  const p3 = points[3];
  const w0 = (1 - u) * (1 - v);
  const w1 = u * (1 - v);
  const w2 = u * v;
  const w3 = (1 - u) * v;
  return {
    x: p0.x * w0 + p1.x * w1 + p2.x * w2 + p3.x * w3,
    y: p0.y * w0 + p1.y * w1 + p2.y * w2 + p3.y * w3,
    z: p0.z * w0 + p1.z * w1 + p2.z * w2 + p3.z * w3,
  };
}

function clampNumber(valueRaw, min, max) {
  const valueNum = Number(valueRaw);
  if (!Number.isFinite(valueNum)) {
    return min;
  }
  if (valueNum < min) {
    return min;
  }
  if (valueNum > max) {
    return max;
  }
  return valueNum;
}

function getOrLoadPlanLargeShapeImage(imagePathRaw, shapeTypeRaw = "") {
  const candidates = getLargeShapeImagePathCandidates(shapeTypeRaw, imagePathRaw);
  if (!candidates.length) {
    return Promise.reject(new Error("imagePath is empty"));
  }
  const cacheKey = candidates.join("|");
  const cached = planLargeShapeImageCache.get(cacheKey);
  if (cached) {
    return cached;
  }
  const promise = new Promise((resolve, reject) => {
    const tryLoad = (index) => {
      if (index >= candidates.length) {
        reject(new Error("image load failed"));
        return;
      }
      const image = new Image();
      image.onload = () => resolve(image);
      image.onerror = () => tryLoad(index + 1);
      image.src = candidates[index];
    };
    tryLoad(0);
  });
  promise.catch(() => {
    if (planLargeShapeImageCache.get(cacheKey) === promise) {
      planLargeShapeImageCache.delete(cacheKey);
    }
  });
  planLargeShapeImageCache.set(cacheKey, promise);
  return promise;
}

function dilateAlphaMask(alphaRaw, width, height, iterations = 1) {
  const w = Math.max(1, Math.floor(width));
  const h = Math.max(1, Math.floor(height));
  let src = alphaRaw instanceof Uint8ClampedArray ? alphaRaw : new Uint8ClampedArray(w * h);
  const iterCount = Math.max(0, Math.floor(iterations));
  for (let iter = 0; iter < iterCount; iter += 1) {
    const dst = new Uint8ClampedArray(src.length);
    for (let y = 0; y < h; y += 1) {
      for (let x = 0; x < w; x += 1) {
        let maxAlpha = 0;
        for (let dy = -1; dy <= 1; dy += 1) {
          const ny = y + dy;
          if (ny < 0 || ny >= h) {
            continue;
          }
          for (let dx = -1; dx <= 1; dx += 1) {
            const nx = x + dx;
            if (nx < 0 || nx >= w) {
              continue;
            }
            const alpha = src[ny * w + nx];
            if (alpha > maxAlpha) {
              maxAlpha = alpha;
              if (maxAlpha >= 255) {
                break;
              }
            }
          }
          if (maxAlpha >= 255) {
            break;
          }
        }
        dst[y * w + x] = maxAlpha;
      }
    }
    src = dst;
  }
  return src;
}

function getOrLoadPlanLargeShapeTintedCanvas(imagePathRaw, colorRaw, shapeTypeRaw = "") {
  const candidates = getLargeShapeImagePathCandidates(shapeTypeRaw, imagePathRaw);
  if (!candidates.length) {
    return Promise.reject(new Error("imagePath is empty"));
  }
  const tint = parseHexColor(colorRaw);
  const key = `${candidates.join("|")}::${tint.hex}`;
  const cached = planLargeShapeTintedCanvasCache.get(key);
  if (cached) {
    return cached;
  }
  const promise = getOrLoadPlanLargeShapeImage(candidates[0], shapeTypeRaw).then((image) => {
    const width = Math.max(1, Number(image.naturalWidth || image.width) || 1);
    const height = Math.max(1, Number(image.naturalHeight || image.height) || 1);
    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const ctx = canvas.getContext("2d", { willReadFrequently: true });
    if (!ctx) {
      throw new Error("2d context unavailable");
    }
    ctx.clearRect(0, 0, width, height);
    ctx.drawImage(image, 0, 0, width, height);
    let imageData;
    try {
      imageData = ctx.getImageData(0, 0, width, height);
    } catch (_error) {
      // ローカル環境等で getImageData が失敗する場合は、原画像のまま表示する。
      return canvas;
    }
    const data = imageData.data;
    const alphaMask = new Uint8ClampedArray(width * height);
    const lineMask = new Uint8ClampedArray(width * height);
    for (let i = 0, p = 0; i < data.length; i += 4, p += 1) {
      const alpha = data[i + 3];
      alphaMask[p] = alpha;
      if (alpha === 0) {
        continue;
      }
      const r = data[i];
      const g = data[i + 1];
      const b = data[i + 2];
      const isNearWhite = r >= 244 && g >= 244 && b >= 244;
      if (!isNearWhite) {
        lineMask[p] = alpha;
      }
    }
    let useLineMask = false;
    for (let i = 0; i < lineMask.length; i += 1) {
      if (lineMask[i] > 0) {
        useLineMask = true;
        break;
      }
    }
    const sourceMask = useLineMask ? lineMask : alphaMask;
    const expandedAlpha = dilateAlphaMask(sourceMask, width, height, IMAGE_SHAPE_CANVAS_DILATE_ITERATIONS);
    for (let i = 0; i < data.length; i += 4) {
      const alpha = expandedAlpha[i / 4];
      if (alpha === 0) {
        data[i + 3] = 0;
        continue;
      }
      data[i] = tint.r;
      data[i + 1] = tint.g;
      data[i + 2] = tint.b;
      data[i + 3] = alpha;
    }
    ctx.putImageData(imageData, 0, 0);
    return canvas;
  });
  promise.catch(() => {
    if (planLargeShapeTintedCanvasCache.get(key) === promise) {
      planLargeShapeTintedCanvasCache.delete(key);
    }
  });
  planLargeShapeTintedCanvasCache.set(key, promise);
  return promise;
}

function getPlanLargeShapeTintedDataUrl(imagePathRaw, colorRaw) {
  const candidates = getLargeShapeImagePathCandidates("", imagePathRaw);
  if (!candidates.length) {
    return "";
  }
  const tint = parseHexColor(colorRaw);
  const key = `${candidates.join("|")}::${tint.hex}`;
  const cached = planLargeShapeTintedDataUrlCache.get(key);
  if (cached === "loading") {
    return "";
  }
  if (typeof cached === "string") {
    return cached;
  }
  planLargeShapeTintedDataUrlCache.set(key, "loading");
  getOrLoadPlanLargeShapeTintedCanvas(candidates[0], tint.hex)
    .then((canvas) => {
      const dataUrl = canvas.toDataURL("image/png");
      planLargeShapeTintedDataUrlCache.set(key, dataUrl);
      renderOutputs();
    })
    .catch(() => {
      planLargeShapeTintedDataUrlCache.delete(key);
    });
  return "";
}

function createViewerTextSprite(textRaw, colorRaw, options = {}) {
  const fontPx = Math.max(12, Number(options?.fontPx) || 44);
  const scaleX = Math.max(0.05, Number(options?.scaleX) || 0.95);
  const scaleY = Math.max(0.05, Number(options?.scaleY) || 0.24);
  const canvas = document.createElement("canvas");
  canvas.width = 512;
  canvas.height = 128;
  const ctx = canvas.getContext("2d");
  ctx.clearRect(0, 0, canvas.width, canvas.height);
  ctx.font = `700 ${fontPx}px 'Yu Gothic', sans-serif`;
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.fillStyle = value(colorRaw) || "#111827";
  ctx.strokeStyle = "rgba(255,255,255,0.96)";
  ctx.lineWidth = 8;
  ctx.strokeText(value(textRaw) || "-", canvas.width / 2, canvas.height / 2);
  ctx.fillText(value(textRaw) || "-", canvas.width / 2, canvas.height / 2);
  const texture = new THREE.CanvasTexture(canvas);
  texture.needsUpdate = true;
  const material = new THREE.SpriteMaterial({ map: texture, transparent: true, depthTest: false });
  const sprite = new THREE.Sprite(material);
  sprite.scale.set(scaleX, scaleY, 1);
  return sprite;
}

function computeViewerBounds(shapes, metrics) {
  const xs = [];
  const ys = [];
  const zs = [];
  const axisMinZ = applyViewerVerticalScale(metrics.minZ, metrics.minZ);
  const axisMaxZ = applyViewerVerticalScale(metrics.maxZ, metrics.minZ);
  zs.push(axisMinZ, axisMaxZ);
  shapes.forEach((shape) => {
    if (shape.type === "point") {
      xs.push(shape.x);
      ys.push(shape.y);
      zs.push(shape.z);
      return;
    }
    (shape.points || []).forEach((point) => {
      xs.push(point.x);
      ys.push(point.y);
      zs.push(point.z);
    });
  });
  if (!xs.length || !ys.length || !zs.length) {
    return {
      minX: 0,
      maxX: metrics.gridWidthM,
      minY: 0,
      maxY: metrics.gridHeightM,
      minZ: axisMinZ,
      maxZ: axisMaxZ,
    };
  }
  return {
    minX: Math.min(...xs),
    maxX: Math.max(...xs),
    minY: Math.min(...ys),
    maxY: Math.max(...ys),
    minZ: Math.min(...zs),
    maxZ: Math.max(...zs),
  };
}

function applyViewerPerspective() {
  if (!viewer3d.initialized || !viewer3d.camera) {
    return;
  }
  const bounds = viewer3d.bounds;
  if (!bounds) {
    return;
  }
  const center = new THREE.Vector3(
    (bounds.minX + bounds.maxX) / 2,
    (bounds.minY + bounds.maxY) / 2,
    (bounds.minZ + bounds.maxZ) / 2
  );
  const spanX = Math.max(1, bounds.maxX - bounds.minX);
  const spanY = Math.max(1, bounds.maxY - bounds.minY);
  const spanZ = Math.max(1, bounds.maxZ - bounds.minZ);
  const scale = normalizeViewerVerticalScale(viewerVerticalScale);
  const unscaledSpanZ = Math.max(1, spanZ / scale);
  const baseDist = Math.max(spanX, spanY) * 1.1 + Math.max(8, unscaledSpanZ * 5);
  const zoomInFactor = Math.pow(scale, -0.7);
  const dist = baseDist * zoomInFactor;
  if (selectedViewerPerspective === "top") {
    viewer3d.camera.up.set(0, 1, 0);
    viewer3d.camera.position.set(center.x, center.y, center.z + dist);
  } else {
    viewer3d.camera.up.set(0, 0, 1);
    viewer3d.camera.position.set(center.x + dist * 0.75, center.y - dist * 0.75, center.z + dist * 0.6);
  }
  if (viewer3d.controls) {
    viewer3d.controls.target.copy(center);
    viewer3d.controls.update();
  } else {
    viewer3d.camera.lookAt(center);
  }
}

function handleViewerPointerMove(event) {
  if (!viewer3d.initialized || !viewer3d.raycaster || !viewer3d.camera || !viewer3d.pickMeshes.length || !viewerCanvasWrap) {
    hideViewerTooltip();
    return;
  }
  const rect = viewerCanvasWrap.getBoundingClientRect();
  if (rect.width <= 0 || rect.height <= 0) {
    hideViewerTooltip();
    return;
  }
  const x = ((event.clientX - rect.left) / rect.width) * 2 - 1;
  const y = -((event.clientY - rect.top) / rect.height) * 2 + 1;
  viewer3d.pointer.set(x, y);
  viewer3d.raycaster.setFromCamera(viewer3d.pointer, viewer3d.camera);
  const intersects = viewer3d.raycaster.intersectObjects(viewer3d.pickMeshes, false);
  if (!intersects.length) {
    hideViewerTooltip();
    return;
  }
  const picked = intersects[0].object?.userData || {};
  showViewerTooltip(event, picked);
}

function showViewerTooltip(event, data) {
  if (!viewerTooltip || !viewerCanvasWrap) {
    return;
  }
  const hasAltitude = Number.isFinite(Number(data.altitudeM));
  const altitudeTextBase = hasAltitude ? Number(data.altitudeM).toFixed(3).replace(/\.?0+$/, "") : "-";
  const altitudeText = data.altitudeEstimated ? `${altitudeTextBase}（仮）` : altitudeTextBase;
  viewerTooltip.innerHTML = `
    <div><strong>標本番号:</strong> ${escapeHtml(value(data.label) || "-")}</div>
    <div><strong>名称:</strong> ${escapeHtml(value(data.nameMemo) || "-")}</div>
    <div><strong>ユニット:</strong> ${escapeHtml(value(data.unit) || "-")}</div>
    <div><strong>サブユニット:</strong> ${escapeHtml(value(data.detail) || "-")}</div>
    <div><strong>区画:</strong> ${escapeHtml(value(data.kuwaku) || "-")}</div>
    <div><strong>標高(m):</strong> ${escapeHtml(altitudeText)}</div>
  `;
  viewerTooltip.hidden = false;
  const rect = viewerCanvasWrap.getBoundingClientRect();
  const maxX = Math.max(8, rect.width - 240);
  const maxY = Math.max(8, rect.height - 132);
  const x = clamp(event.clientX - rect.left + 14, 8, maxX);
  const y = clamp(event.clientY - rect.top + 12, 8, maxY);
  viewerTooltip.style.left = `${x}px`;
  viewerTooltip.style.top = `${y}px`;
}

function hideViewerTooltip() {
  if (!viewerTooltip) {
    return;
  }
  viewerTooltip.hidden = true;
}

function blockIndexToLabel(indexRaw) {
  let index = Number(indexRaw);
  if (!Number.isFinite(index) || index < 1) {
    return "A";
  }
  let label = "";
  while (index > 0) {
    const remainder = (index - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    index = Math.floor((index - 1) / 26);
  }
  return label || "A";
}

function normalizeViewerHead(headRaw) {
  const head = value(headRaw).toUpperCase();
  if (head === "Ⅲ" || head === "3" || head === "III") {
    return "Ⅲ";
  }
  if (head === "Ⅱ" || head === "2" || head === "II") {
    return "Ⅱ";
  }
  return "Ⅰ";
}

function kuwakuToViewerXIndex(kuwakuParts) {
  const parts = kuwakuParts || {};
  const head = normalizeViewerHead(parts.headB);
  const headIndex = VIEWER_HEAD_INDEX_MAP.get(head);
  const block = normalizeKuwakuBlock(parts.block);
  const baseLetterIndex = block ? block.charCodeAt(0) - 65 : 0;
  const letterIndex = Number.isFinite(baseLetterIndex) && baseLetterIndex >= 0 && baseLetterIndex < 26 ? baseLetterIndex : 0;
  if (!Number.isFinite(headIndex)) {
    return blockLabelToIndex(block) - 1;
  }
  return headIndex * 26 + letterIndex;
}

function viewerXIndexToHeadBlock(indexRaw) {
  const index = Number(indexRaw);
  if (!Number.isFinite(index)) {
    return { head: "Ⅰ", block: "A" };
  }
  const seqLength = VIEWER_HEAD_SEQUENCE.length;
  const normalized = ((Math.floor(index) % (26 * seqLength)) + 26 * seqLength) % (26 * seqLength);
  const headIndex = Math.floor(normalized / 26) % seqLength;
  const letterIndex = normalized % 26;
  return {
    head: VIEWER_HEAD_SEQUENCE[headIndex] || "Ⅰ",
    block: String.fromCharCode(65 + letterIndex),
  };
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

function detailValueForSelect(detailRaw) {
  const detail = value(detailRaw);
  return detail || EMPTY_DETAIL_VALUE;
}

function detailLabelForSelect(detailValue) {
  return detailValue === EMPTY_DETAIL_VALUE ? "（未設定）" : detailValue;
}

function detailSubValueForSelect(detailSubRaw) {
  const detailSub = value(detailSubRaw);
  return detailSub || EMPTY_DETAIL_SUB_VALUE;
}

function detailSubLabelForSelect(detailSubValue) {
  return detailSubValue === EMPTY_DETAIL_SUB_VALUE ? "（未設定）" : detailSubValue;
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

function buildPlanDrawableMeta(record) {
  const specimen = parseSpecimenNo(record.specimenNo, record.specimenPrefix, record.specimenSerial);
  const prefix = normalizeSpecimenPrefix(specimen.prefix);
  return {
    color: getSpecimenPrefixColor(prefix),
    label: record.specimenNo || "",
    nameMemo: value(record.nameMemo),
    unit: value(record.unit),
    detail: buildDetailText(record.detail, record.detailSub),
  };
}

function convertPositionToPlanCoords(nsDirRaw, nsCmRaw, ewDirRaw, ewCmRaw) {
  const nsCm = parseDistanceToCm(nsCmRaw);
  const ewCm = parseDistanceToCm(ewCmRaw);
  if (nsCm == null || ewCm == null) {
    return null;
  }
  const nsDir = normalizeNsDir(nsDirRaw);
  const ewDir = normalizeEwDir(ewDirRaw);
  const yRaw = nsDir === "北から" ? nsCm : PLAN_SIZE_CM - nsCm;
  const xRaw = ewDir === "西から" ? ewCm : PLAN_SIZE_CM - ewCm;
  return {
    x: xRaw,
    y: yRaw,
  };
}

function parseLargeAxisAzimuth(valueRaw) {
  const text = normalizeLargeAxisDirection(valueRaw);
  if (text === "NS") {
    return 0;
  }
  if (text === "EW") {
    return 90;
  }
  const matched = text.match(/^([NS])(\d+(?:\.\d+)?)([EW])$/);
  if (!matched) {
    return null;
  }
  const [, ns, degreeRaw, ew] = matched;
  const degree = Number(degreeRaw);
  if (!Number.isFinite(degree) || degree < 0 || degree > 90) {
    return null;
  }
  if (ns === "N" && ew === "E") {
    return degree;
  }
  if (ns === "N" && ew === "W") {
    return (360 - degree) % 360;
  }
  if (ns === "S" && ew === "E") {
    return 180 - degree;
  }
  if (ns === "S" && ew === "W") {
    return 180 + degree;
  }
  return null;
}

function azimuthToPlanUnitVector(azimuthDegRaw) {
  const azimuthDeg = Number(azimuthDegRaw);
  if (!Number.isFinite(azimuthDeg)) {
    return { dx: 0, dy: -1 };
  }
  const rad = (azimuthDeg * Math.PI) / 180;
  return {
    dx: Math.sin(rad),
    dy: -Math.cos(rad),
  };
}

function azimuthToSvgRotationDeg(azimuthDegRaw) {
  const azimuthDeg = Number(azimuthDegRaw);
  if (!Number.isFinite(azimuthDeg)) {
    return 0;
  }
  return azimuthDeg - 90;
}

function pointsToAzimuthDeg(pointA, pointB) {
  if (!pointA || !pointB) {
    return null;
  }
  const dx = pointB.x - pointA.x;
  const dy = pointB.y - pointA.y;
  const distance = Math.hypot(dx, dy);
  if (!Number.isFinite(distance) || distance <= 0) {
    return null;
  }
  const rad = Math.atan2(dx, -dy);
  const deg = (rad * 180) / Math.PI;
  return (deg + 360) % 360;
}

function parseImageQuadPlanPoints(record, center = null) {
  const corner1 = convertPositionToPlanCoords(record?.imgP1NsDir, record?.imgP1NsCm, record?.imgP1EwDir, record?.imgP1EwCm);
  const corner2 = convertPositionToPlanCoords(record?.imgP2NsDir, record?.imgP2NsCm, record?.imgP2EwDir, record?.imgP2EwCm);
  const corner3 = convertPositionToPlanCoords(record?.imgP3NsDir, record?.imgP3NsCm, record?.imgP3EwDir, record?.imgP3EwCm);
  const corner4 = convertPositionToPlanCoords(record?.imgP4NsDir, record?.imgP4NsCm, record?.imgP4EwDir, record?.imgP4EwCm);
  if (!corner1 || !corner2 || !corner3 || !corner4) {
    if (!center) {
      return null;
    }
    const candidates = [corner1, corner2, corner3, corner4, center].filter(Boolean);
    if (!candidates.length) {
      return null;
    }
    let minX = Math.min(...candidates.map((point) => point.x));
    let maxX = Math.max(...candidates.map((point) => point.x));
    let minY = Math.min(...candidates.map((point) => point.y));
    let maxY = Math.max(...candidates.map((point) => point.y));
    const fallbackHalfSize = 20;
    if (Math.abs(maxX - minX) < 1) {
      minX = center.x - fallbackHalfSize;
      maxX = center.x + fallbackHalfSize;
    }
    if (Math.abs(maxY - minY) < 1) {
      minY = center.y - fallbackHalfSize;
      maxY = center.y + fallbackHalfSize;
    }
    return [
      { x: minX, y: minY },
      { x: maxX, y: minY },
      { x: maxX, y: maxY },
      { x: minX, y: maxY },
    ];
  }
  return [corner1, corner2, corner3, corner4];
}

function collectImageCornerPoints(record) {
  const corners = [
    convertPositionToPlanCoords(record?.imgP1NsDir, record?.imgP1NsCm, record?.imgP1EwDir, record?.imgP1EwCm),
    convertPositionToPlanCoords(record?.imgP2NsDir, record?.imgP2NsCm, record?.imgP2EwDir, record?.imgP2EwCm),
    convertPositionToPlanCoords(record?.imgP3NsDir, record?.imgP3NsCm, record?.imgP3EwDir, record?.imgP3EwCm),
    convertPositionToPlanCoords(record?.imgP4NsDir, record?.imgP4NsCm, record?.imgP4EwDir, record?.imgP4EwCm),
  ].filter(Boolean);
  return corners;
}

function buildPlanDrawable(record) {
  const meta = buildPlanDrawableMeta(record);

  const planSizeMode = normalizePlanSizeMode(record.planSizeMode);
  const shapeType = normalizeLargeShapeType(record.largeShapeType);
  const normalizedShapeLabel = normalizeLargeShapeLabel(record.largeShapeType);
  const isImageShape = isLargeShapeImageType(shapeType);
  const hasMappedImageType = largeShapeImagePathMap.has(normalizedShapeLabel);
  const rawImageCorners = collectImageCornerPoints(record);
  let center = convertPositionToPlanCoords(record.nsDir, record.nsCm, record.ewDir, record.ewCm);
  if (!center && (isImageShape || rawImageCorners.length)) {
    if (rawImageCorners.length) {
      center = rawImageCorners.reduce(
        (acc, point) => ({ x: acc.x + point.x / rawImageCorners.length, y: acc.y + point.y / rawImageCorners.length }),
        { x: 0, y: 0 }
      );
    }
  }
  if (!center) {
    return null;
  }
  const resolvedImageType = isImageShape ? shapeType : hasMappedImageType ? normalizedShapeLabel : "";
  const shouldUseImageQuad =
    planSizeMode === "大きなもの" && (isImageShape || (resolvedImageType && rawImageCorners.length > 0));
  const axisAzimuth = shouldUseImageQuad ? null : parseLargeAxisAzimuth(record.largeAxisDirection);

  if (planSizeMode !== "大きなもの" || !shapeType) {
    return {
      type: "point",
      x: center.x,
      y: center.y,
      ...meta,
    };
  }

  if (shapeType === "直線状") {
    const lineLength = parseDistanceToCm(record.lineLengthCm);
    if (lineLength == null || lineLength <= 0) {
      return null;
    }
    if (axisAzimuth == null) {
      return null;
    }
    const unit = azimuthToPlanUnitVector(axisAzimuth);
    const halfLength = lineLength / 2;
    const x1 = center.x - unit.dx * halfLength;
    const y1 = center.y - unit.dy * halfLength;
    const x2 = center.x + unit.dx * halfLength;
    const y2 = center.y + unit.dy * halfLength;
    return {
      type: "line",
      x1,
      y1,
      x2,
      y2,
      x: center.x,
      y: center.y,
      ...meta,
    };
  }

  if (shapeType === "長方形") {
    const side1 = parseDistanceToCm(record.rectSide1Cm);
    const side2 = parseDistanceToCm(record.rectSide2Cm);
    if (side1 == null || side2 == null) {
      return null;
    }
    const longSide = Math.max(side1, side2);
    const shortSide = Math.min(side1, side2);
    const width = Math.max(1, longSide);
    const height = Math.max(1, shortSide);
    return {
      type: "rect",
      x: center.x,
      y: center.y,
      left: center.x - width / 2,
      top: center.y - height / 2,
      width,
      height,
      rotationDeg: azimuthToSvgRotationDeg(axisAzimuth ?? 90),
      ...meta,
    };
  }

  if (shapeType === "楕円") {
    const rx = parseDistanceToCm(record.ellipseLongRadiusCm);
    const ry = parseDistanceToCm(record.ellipseShortRadiusCm);
    if (rx == null || ry == null) {
      return null;
    }
    const longRadius = Math.max(rx, ry);
    const shortRadius = Math.min(rx, ry);
    return {
      type: "ellipse",
      x: center.x,
      y: center.y,
      rx: Math.min(Math.max(1, longRadius), PLAN_SIZE_CM),
      ry: Math.min(Math.max(1, shortRadius), PLAN_SIZE_CM),
      rotationDeg: azimuthToSvgRotationDeg(axisAzimuth ?? 90),
      ...meta,
    };
  }

  if (shouldUseImageQuad) {
    const points = parseImageQuadPlanPoints(record, center);
    if (!points) {
      return null;
    }
    const centroid = points.reduce(
      (acc, point) => ({ x: acc.x + point.x / points.length, y: acc.y + point.y / points.length }),
      { x: 0, y: 0 }
    );
    return {
      type: "imageQuad",
      points,
      imageType: resolvedImageType,
      imagePath: getLargeShapeImagePath(resolvedImageType),
      x: centroid.x,
      y: centroid.y,
      ...meta,
    };
  }

  return null;
}

function renderPlanDrawableSvg(drawable, index = 0) {
  const labelX = Math.min(PLAN_SIZE_CM - 2, drawable.x + 6);
  const labelY = Math.max(8, drawable.y - 6);
  const ariaLabel = [
    `標本番号 ${drawable.label || "未設定"}`,
    `化石・遺物名称 ${drawable.nameMemo || "未設定"}`,
    `ユニット ${drawable.unit || "未設定"}`,
    `サブユニット ${drawable.detail || "未設定"}`,
  ].join(" / ");

  let shapeSvg = "";
  if (drawable.type === "line") {
    shapeSvg = `<line class="plan-shape-line" x1="${drawable.x1}" y1="${drawable.y1}" x2="${drawable.x2}" y2="${drawable.y2}" stroke="${drawable.color}" />`;
  } else if (drawable.type === "rect") {
    const transform = Number.isFinite(drawable.rotationDeg)
      ? ` transform="rotate(${drawable.rotationDeg} ${drawable.x} ${drawable.y})"`
      : "";
    shapeSvg = `<rect class="plan-shape-rect" x="${drawable.left}" y="${drawable.top}" width="${drawable.width}" height="${drawable.height}" stroke="${drawable.color}"${transform} />`;
  } else if (drawable.type === "ellipse") {
    const transform = Number.isFinite(drawable.rotationDeg)
      ? ` transform="rotate(${drawable.rotationDeg} ${drawable.x} ${drawable.y})"`
      : "";
    shapeSvg = `<ellipse class="plan-shape-ellipse" cx="${drawable.x}" cy="${drawable.y}" rx="${drawable.rx}" ry="${drawable.ry}" stroke="${drawable.color}"${transform} />`;
  } else if (drawable.type === "imageQuad") {
    const imageSvg = buildPlanImageWarpSvg(drawable, index);
    const polygonPoints = (drawable.points || []).map((point) => `${point.x},${point.y}`).join(" ");
    const outlineSvg = `<polygon class="plan-shape-image-outline" points="${polygonPoints}" fill="none" stroke="${drawable.color}" />`;
    shapeSvg = imageSvg ? `${imageSvg}${outlineSvg}` : outlineSvg;
  } else {
    shapeSvg = `<circle class="plan-point-hit" cx="${drawable.x}" cy="${drawable.y}" r="5" fill="${drawable.color}" />`;
  }

  const hotspotSvg =
    drawable.type === "imageQuad"
      ? `<polygon class="plan-point-hotspot plan-image-hotspot" points="${(drawable.points || [])
          .map((point) => `${point.x},${point.y}`)
          .join(" ")}" fill="transparent" />`
      : `<circle class="plan-point-hotspot" cx="${drawable.x}" cy="${drawable.y}" r="12" fill="transparent" />`;

  return `
      <g
        class="plan-point-group"
        data-label="${escapeHtml(drawable.label || "")}"
        data-name-memo="${escapeHtml(drawable.nameMemo || "")}"
        data-unit="${escapeHtml(drawable.unit || "")}"
        data-detail="${escapeHtml(drawable.detail || "")}"
        data-x="${drawable.x}"
        data-y="${drawable.y}"
        tabindex="0"
        role="button"
        aria-label="${escapeHtml(ariaLabel)}"
      >
        ${shapeSvg}
        ${hotspotSvg}
        <text x="${labelX}" y="${labelY}">${escapeHtml(drawable.label || "")}</text>
      </g>
    `;
}

function parseHexColor(colorRaw, fallback = "#6b7280") {
  const text = value(colorRaw);
  const fallbackText = value(fallback) || "#6b7280";
  const normalized = /^#[0-9a-f]{3}$/i.test(text)
    ? `#${text[1]}${text[1]}${text[2]}${text[2]}${text[3]}${text[3]}`
    : /^#[0-9a-f]{6}$/i.test(text)
      ? text
      : fallbackText;
  const matched = normalized.match(/^#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i);
  if (!matched) {
    return { hex: "#6b7280", r: 107, g: 114, b: 128 };
  }
  return {
    hex: `#${matched[1]}${matched[2]}${matched[3]}`.toLowerCase(),
    r: Number.parseInt(matched[1], 16),
    g: Number.parseInt(matched[2], 16),
    b: Number.parseInt(matched[3], 16),
  };
}

function buildPlanImageWarpSvg(drawable, index = 0) {
  const points = Array.isArray(drawable?.points) ? drawable.points : [];
  const imagePath = value(drawable?.imagePath);
  const imageCandidates = getLargeShapeImagePathCandidates(drawable?.imageType, imagePath);
  const imageFallback = imageCandidates[0] || imagePath;
  const tintedDataUrl = getPlanLargeShapeTintedDataUrl(imageFallback, drawable?.color);
  const imageRef = tintedDataUrl || imageFallback;
  if (points.length !== 4 || !imageRef) {
    return "";
  }
  const [p1, p2, p3, p4] = points;
  const labelKey = value(drawable.label || "x").replace(/[^a-zA-Z0-9_-]/g, "");
  const clipIdA = `plan-img-clip-a-${index}-${labelKey || "x"}`;
  const clipIdB = `plan-img-clip-b-${index}-${labelKey || "x"}`;
  const clipIdRect = `plan-img-clip-rect-${index}-${labelKey || "x"}`;
  const matrixA = [p2.x - p1.x, p2.y - p1.y, p3.x - p2.x, p3.y - p2.y, p1.x, p1.y];
  const matrixB = [p3.x - p4.x, p3.y - p4.y, p4.x - p1.x, p4.y - p1.y, p1.x, p1.y];
  const triA = `${p1.x},${p1.y} ${p2.x},${p2.y} ${p3.x},${p3.y}`;
  const triB = `${p1.x},${p1.y} ${p3.x},${p3.y} ${p4.x},${p4.y}`;
  const quadPoints = `${p1.x},${p1.y} ${p2.x},${p2.y} ${p3.x},${p3.y} ${p4.x},${p4.y}`;
  const minX = Math.min(p1.x, p2.x, p3.x, p4.x);
  const maxX = Math.max(p1.x, p2.x, p3.x, p4.x);
  const minY = Math.min(p1.y, p2.y, p3.y, p4.y);
  const maxY = Math.max(p1.y, p2.y, p3.y, p4.y);
  const boxWidth = Math.max(1, maxX - minX);
  const boxHeight = Math.max(1, maxY - minY);
  const matrixAText = matrixA.map((num) => (Number.isFinite(num) ? Number(num).toFixed(4) : "0")).join(" ");
  const matrixBText = matrixB.map((num) => (Number.isFinite(num) ? Number(num).toFixed(4) : "0")).join(" ");
  return `
    <defs>
      <clipPath id="${clipIdA}">
        <polygon points="${triA}" />
      </clipPath>
      <clipPath id="${clipIdB}">
        <polygon points="${triB}" />
      </clipPath>
      <clipPath id="${clipIdRect}">
        <polygon points="${quadPoints}" />
      </clipPath>
    </defs>
    <image href="${escapeHtml(imageRef)}" x="${minX}" y="${minY}" width="${boxWidth}" height="${boxHeight}" preserveAspectRatio="none" clip-path="url(#${clipIdRect})" />
    <image href="${escapeHtml(imageRef)}" x="0" y="0" width="1" height="1" preserveAspectRatio="none" transform="matrix(${matrixAText})" clip-path="url(#${clipIdA})" />
    <image href="${escapeHtml(imageRef)}" x="0" y="0" width="1" height="1" preserveAspectRatio="none" transform="matrix(${matrixBText})" clip-path="url(#${clipIdB})" />
  `;
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
  const order = ["m", "b", "l", "s", "i", "g", "h", "a"];
  return order
    .map((prefix) => {
      const color = getSpecimenPrefixColor(prefix);
      const label = SPECIMEN_CATEGORY_MAP[prefix] || "";
      return `<span class="plan-legend-item"><span class="plan-legend-dot" style="background:${color}"></span>${prefix}: ${label}</span>`;
    })
    .join("");
}

function getSpecimenPrefixColor(prefixRaw) {
  const prefix = normalizeSpecimenPrefix(prefixRaw);
  return SPECIMEN_POINT_COLORS[prefix] || SPECIMEN_POINT_COLORS.h;
}

function getRecordSpecimenColor(record) {
  const specimen = parseSpecimenNo(record?.specimenNo, record?.specimenPrefix, record?.specimenSerial);
  return getSpecimenPrefixColor(specimen.prefix);
}

function getKuwakuCellStyle(kuwakuRaw) {
  const kuwaku = normalizeKuwakuText(kuwakuRaw);
  if (!kuwaku) {
    return {
      background: "#f3f4f6",
      border: "#d1d5db",
      color: "#111827",
    };
  }
  const parts = parseKuwaku(kuwaku);
  const blockIndex = blockLabelToIndex(parts.block);
  const noSeed = /^-?\d+$/.test(parts.no) ? Number(parts.no) : (hashText(parts.no || kuwaku) % 97) + 1;
  const headSeed = hashText(`${parts.headA}-${parts.headB}`) % 360;
  const hue = (((blockIndex * 41 + noSeed * 17 + headSeed) % 360) + 360) % 360;
  const sat = 66;
  const bgLightness = 93;
  const borderLightness = 82;
  return {
    background: `hsl(${hue}, ${sat}%, ${bgLightness}%)`,
    border: `hsl(${hue}, 48%, ${borderLightness}%)`,
    color: "#111827",
  };
}

function getUnitCellStyle(unitRaw) {
  const normalized = compactNoSpaceValue(unitRaw).toUpperCase();
  if (UNIT_CELL_COLOR_MAP[normalized]) {
    return UNIT_CELL_COLOR_MAP[normalized];
  }
  if (!normalized) {
    return { background: "#f3f4f6", border: "#d1d5db", color: "#111827" };
  }
  const hue = hashText(normalized) % 360;
  return {
    background: `hsl(${hue}, 58%, 93%)`,
    border: `hsl(${hue}, 38%, 80%)`,
    color: "#111827",
  };
}

function blockLabelToIndex(blockRaw) {
  const block = normalizeKuwakuBlock(blockRaw);
  if (!block || !/^[A-Z]+$/.test(block)) {
    return (hashText(block) % 26) + 1;
  }
  let index = 0;
  for (const char of block) {
    index = index * 26 + (char.charCodeAt(0) - 64);
  }
  return index;
}

function hashText(textRaw) {
  const text = value(textRaw);
  let hash = 0;
  for (let i = 0; i < text.length; i += 1) {
    hash = (hash * 31 + text.charCodeAt(i)) >>> 0;
  }
  return hash;
}

function toRgbaColor(hexColorRaw, alphaRaw) {
  const hexColor = value(hexColorRaw).replace("#", "");
  const alpha = clamp(Number(alphaRaw), 0, 1);
  if (!/^[0-9a-fA-F]{6}$/.test(hexColor)) {
    return `rgba(107, 114, 128, ${alpha})`;
  }
  const r = Number.parseInt(hexColor.slice(0, 2), 16);
  const g = Number.parseInt(hexColor.slice(2, 4), 16);
  const b = Number.parseInt(hexColor.slice(4, 6), 16);
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
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
      <div><strong>サブユニット:</strong> ${escapeHtml(detail)}</div>
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
  const rawLargeShapeType = value(item.largeShapeType);
  const normalizedLargeShapeType = normalizeLargeShapeType(rawLargeShapeType) || normalizeLargeShapeLabel(rawLargeShapeType);

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
    sectionDiagramDistanceChecked: normalizeChecklistChecked(value(item.sectionDiagramDistanceChecked)),
    sectionDiagramHorizonChecked: normalizeChecklistChecked(value(item.sectionDiagramHorizonChecked)),
    sectionDiagramLayerFaciesChecked: normalizeChecklistChecked(value(item.sectionDiagramLayerFaciesChecked)),
    photoClinometerChecked: normalizeChecklistChecked(value(item.photoClinometerChecked)),
    photoRulerChecked: normalizeChecklistChecked(value(item.photoRulerChecked)),
    nsDir: normalizeNsDir(value(item.nsDir)),
    nsCm: value(item.nsCm),
    ewDir: normalizeEwDir(value(item.ewDir)),
    ewCm: value(item.ewCm),
    planSizeMode: normalizePlanSizeMode(value(item.planSizeMode)),
    largeShapeType: normalizedLargeShapeType,
    largeAxisDirection: normalizeLargeAxisDirection(value(item.largeAxisDirection)),
    lineLengthCm: value(item.lineLengthCm),
    line1NsDir: normalizeNsDir(value(item.line1NsDir)),
    line1NsCm: value(item.line1NsCm),
    line1EwDir: normalizeEwDir(value(item.line1EwDir)),
    line1EwCm: value(item.line1EwCm),
    line2NsDir: normalizeNsDir(value(item.line2NsDir)),
    line2NsCm: value(item.line2NsCm),
    line2EwDir: normalizeEwDir(value(item.line2EwDir)),
    line2EwCm: value(item.line2EwCm),
    rectSide1Cm: value(item.rectSide1Cm),
    rectSide2Cm: value(item.rectSide2Cm),
    ellipseLongRadiusCm: value(item.ellipseLongRadiusCm),
    ellipseShortRadiusCm: value(item.ellipseShortRadiusCm),
    imgP1NsDir: normalizeNsDir(value(item.imgP1NsDir)),
    imgP1NsCm: value(item.imgP1NsCm),
    imgP1EwDir: normalizeEwDir(value(item.imgP1EwDir)),
    imgP1EwCm: value(item.imgP1EwCm),
    imgP2NsDir: normalizeNsDir(value(item.imgP2NsDir)),
    imgP2NsCm: value(item.imgP2NsCm),
    imgP2EwDir: normalizeEwDir(value(item.imgP2EwDir)),
    imgP2EwCm: value(item.imgP2EwCm),
    imgP3NsDir: normalizeNsDir(value(item.imgP3NsDir)),
    imgP3NsCm: value(item.imgP3NsCm),
    imgP3EwDir: normalizeEwDir(value(item.imgP3EwDir)),
    imgP3EwCm: value(item.imgP3EwCm),
    imgP4NsDir: normalizeNsDir(value(item.imgP4NsDir)),
    imgP4NsCm: value(item.imgP4NsCm),
    imgP4EwDir: normalizeEwDir(value(item.imgP4EwDir)),
    imgP4EwCm: value(item.imgP4EwCm),
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
  if (!value(normalized.lineLengthCm) && normalized.largeShapeType === "直線状") {
    const p1 = convertPositionToPlanCoords(normalized.line1NsDir, normalized.line1NsCm, normalized.line1EwDir, normalized.line1EwCm);
    const p2 = convertPositionToPlanCoords(normalized.line2NsDir, normalized.line2NsCm, normalized.line2EwDir, normalized.line2EwCm);
    if (p1 && p2) {
      const distance = Math.hypot(p2.x - p1.x, p2.y - p1.y);
      if (Number.isFinite(distance) && distance > 0) {
        normalized.lineLengthCm = trimNumericText(distance.toFixed(1));
      }
    }
  }
  return normalized;
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
    const canApplyRemote =
      activeTabId === "output-tab" || activeTabId === "plan-tab" || activeTabId === "viewer-tab" || activeTabId === "export-tab";
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
    if (activeTabId === "output-tab" || activeTabId === "plan-tab" || activeTabId === "viewer-tab" || activeTabId === "export-tab") {
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
  if (!force && tabAtRequest !== "output-tab" && tabAtRequest !== "plan-tab" && tabAtRequest !== "viewer-tab" && tabAtRequest !== "export-tab") {
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
    let mergedStateForSave = normalizeState(state);
    try {
      const remoteResponse = await requestCloud("load");
      const remoteState = normalizeState(remoteResponse?.state);
      mergedStateForSave = mergeStatesForCloud(remoteState, mergedStateForSave);
      const remoteUpdatedAt = value(remoteResponse?.updatedAt) || getStateUpdatedAt(remoteState);
      if (remoteUpdatedAt) {
        cloudLastPulledAt = remoteUpdatedAt;
      }
    } catch (_error) {
      // 読込に失敗しても、従来どおり端末側データで保存処理は継続する。
      mergedStateForSave = normalizeState(state);
    }
    const payload = {
      clientId: cloudClientId,
      updatedAt: getStateUpdatedAt(mergedStateForSave) || nowIso(),
      state: mergedStateForSave,
    };
    const response = await requestCloud("save", payload);
    state = normalizeState(mergedStateForSave);
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    } catch (_error) {
      // 端末容量不足時でもクラウド保存結果は維持する。
    }
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
    const saved = normalizeCloudEndpoint(value(localStorage.getItem(CLOUD_ENDPOINT_KEY)));
    if (saved) {
      return saved;
    }
    return DEFAULT_CLOUD_ENDPOINT;
  } catch (_error) {
    return DEFAULT_CLOUD_ENDPOINT;
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

function mergeStatesForCloud(remoteStateRaw, localStateRaw) {
  const remoteState = normalizeState(remoteStateRaw);
  const localState = normalizeState(localStateRaw);
  const mergedSite = mergeSiteForCloud(remoteState.site, localState.site);
  const recordById = new Map();

  const upsertRecord = (recordRaw, preferIncomingOnTie) => {
    const record = normalizeRecord(recordRaw, mergedSite);
    const recordId = value(record?.id);
    if (!recordId) {
      return;
    }
    const existing = recordById.get(recordId);
    if (!existing) {
      recordById.set(recordId, record);
      return;
    }
    recordById.set(recordId, chooseNewerRecordForCloud(existing, record, preferIncomingOnTie));
  };

  remoteState.records.forEach((record) => upsertRecord(record, false));
  localState.records.forEach((record) => upsertRecord(record, true));

  return normalizeState({
    site: mergedSite,
    records: Array.from(recordById.values()),
  });
}

function mergeSiteForCloud(remoteSiteRaw, localSiteRaw) {
  const remoteSite = remoteSiteRaw && typeof remoteSiteRaw === "object" ? remoteSiteRaw : {};
  const localSite = localSiteRaw && typeof localSiteRaw === "object" ? localSiteRaw : {};
  const preferLocal = isIsoTimestampGreaterOrEqual(localSite.updatedAt, remoteSite.updatedAt);
  const primary = preferLocal ? localSite : remoteSite;
  const secondary = preferLocal ? remoteSite : localSite;
  const mergedUpdatedAt = pickLatestIsoTimestamp(localSite.updatedAt, remoteSite.updatedAt);

  return {
    kuwaku: value(primary.kuwaku) || value(secondary.kuwaku) || DEFAULT_KUWAKU,
    kuwakuHeadA: value(primary.kuwakuHeadA) || value(secondary.kuwakuHeadA) || DEFAULT_KUWAKU_HEAD_A,
    kuwakuHeadB: value(primary.kuwakuHeadB) || value(secondary.kuwakuHeadB) || DEFAULT_KUWAKU_HEAD_B,
    kuwakuBlock: value(primary.kuwakuBlock) || value(secondary.kuwakuBlock),
    kuwakuNo: value(primary.kuwakuNo) || value(secondary.kuwakuNo),
    levelHeight: value(primary.levelHeight) || value(secondary.levelHeight),
    date: value(primary.date) || value(secondary.date),
    team: value(primary.team) || value(secondary.team),
    teamOther: value(primary.teamOther) || value(secondary.teamOther),
    teamLead: value(primary.teamLead) || value(secondary.teamLead),
    recorder: value(primary.recorder) || value(secondary.recorder),
    updatedAt: mergedUpdatedAt || value(primary.updatedAt) || value(secondary.updatedAt),
  };
}

function chooseNewerRecordForCloud(existingRecordRaw, incomingRecordRaw, preferIncomingOnTie = true) {
  const existingRecord = normalizeRecord(existingRecordRaw);
  const incomingRecord = normalizeRecord(incomingRecordRaw);
  const existingMs = getRecordUpdatedAtMs(existingRecord);
  const incomingMs = getRecordUpdatedAtMs(incomingRecord);
  let winner = existingRecord;
  let loser = incomingRecord;

  if (Number.isFinite(incomingMs) && Number.isFinite(existingMs)) {
    if (incomingMs > existingMs || (incomingMs === existingMs && preferIncomingOnTie)) {
      winner = incomingRecord;
      loser = existingRecord;
    }
  } else if (Number.isFinite(incomingMs) && !Number.isFinite(existingMs)) {
    winner = incomingRecord;
    loser = existingRecord;
  } else if (!Number.isFinite(incomingMs) && !Number.isFinite(existingMs) && preferIncomingOnTie) {
    winner = incomingRecord;
    loser = existingRecord;
  }

  const mergedHistory = mergeRecordHistoryEntries(existingRecord?.history, incomingRecord?.history);
  return {
    ...loser,
    ...winner,
    history: mergedHistory.length ? mergedHistory : normalizeRecordHistory(winner?.history),
  };
}

function mergeRecordHistoryEntries(historyA, historyB) {
  const mergedMap = new Map();
  const upsert = (entryRaw) => {
    const entry = normalizeRecordHistory([entryRaw])[0];
    if (!entry) {
      return;
    }
    const key = value(entry.id) || `${value(entry.at)}::${value(entry.action)}::${value(entry.content)}`;
    if (!key) {
      return;
    }
    const existing = mergedMap.get(key);
    if (!existing) {
      mergedMap.set(key, entry);
      return;
    }
    const existingMs = Date.parse(value(existing.at));
    const incomingMs = Date.parse(value(entry.at));
    if (Number.isFinite(incomingMs) && (!Number.isFinite(existingMs) || incomingMs >= existingMs)) {
      mergedMap.set(key, entry);
    }
  };

  normalizeRecordHistory(historyA).forEach(upsert);
  normalizeRecordHistory(historyB).forEach(upsert);

  return Array.from(mergedMap.values())
    .sort((a, b) => {
      const aMs = Date.parse(value(a.at));
      const bMs = Date.parse(value(b.at));
      if (Number.isFinite(aMs) && Number.isFinite(bMs) && aMs !== bMs) {
        return aMs - bMs;
      }
      return value(a.id).localeCompare(value(b.id));
    })
    .slice(-50);
}

function getRecordUpdatedAtMs(record) {
  const updatedMs = Date.parse(value(record?.updatedAt));
  if (Number.isFinite(updatedMs)) {
    return updatedMs;
  }
  const createdMs = Date.parse(value(record?.createdAt));
  return Number.isFinite(createdMs) ? createdMs : Number.NaN;
}

function pickLatestIsoTimestamp(tsA, tsB) {
  const aMs = Date.parse(value(tsA));
  const bMs = Date.parse(value(tsB));
  if (Number.isFinite(aMs) && Number.isFinite(bMs)) {
    return new Date(Math.max(aMs, bMs)).toISOString();
  }
  if (Number.isFinite(aMs)) {
    return new Date(aMs).toISOString();
  }
  if (Number.isFinite(bMs)) {
    return new Date(bMs).toISOString();
  }
  return "";
}

function isIsoTimestampGreaterOrEqual(tsA, tsB) {
  const aMs = Date.parse(value(tsA));
  const bMs = Date.parse(value(tsB));
  if (Number.isFinite(aMs) && Number.isFinite(bMs)) {
    return aMs >= bMs;
  }
  if (Number.isFinite(aMs)) {
    return true;
  }
  if (Number.isFinite(bMs)) {
    return false;
  }
  return true;
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
    "標高(m)",
    "産出状況断面",
    "産状スケッチ",
    "平面位置_NS",
    "平面位置_NS_cm",
    "平面位置_EW",
    "平面位置_EW_cm",
    "備考（観察事項など）",
  ];

  const rows = getListExportRecords().map((record) => [
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
    formatCmValue(record.levelUpperCm),
    formatCmValue(record.levelLowerCm),
    formatRecordAltitudeM(record),
    record.occurrenceSection,
    record.occurrenceSketch,
    record.nsDir,
    formatCmValue(record.nsCm),
    record.ewDir,
    formatCmValue(record.ewCm),
    record.notes,
  ]);

  return [header, ...rows].map((row) => row.map(csvCell).join(",")).join("\n");
}

function buildCardCsv() {
  const records = getCardExportRecords();
  const header = [
    "区画",
    "標本番号",
    "分類",
    "化石・遺物名称",
    "重要品指定",
    "簡易記載",
    "地層名",
    "ユニット",
    "サブユニット",
    "細分",
    "地層中の位置",
    "レベル読値",
    "平面位置",
    "発見者",
    "判定者",
    "産出状況断面",
    "産状スケッチ",
    "備考（観察事項など）",
    "産出状況断面図枚数",
    "写真枚数",
  ];
  const rows = records.map((record) => [
    getRecordKuwaku(record),
    record.specimenNo,
    formatCategoryForRecord(record),
    record.nameMemo,
    record.importantFlag,
    record.simpleRecordFlag,
    record.layerName,
    record.unit,
    record.detail,
    record.detailSub,
    formatLayerPosition(record),
    formatLevelRead(record),
    formatPlanPosition(record),
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

function exportListPdf() {
  const records = getListExportRecords();
  if (!records.length) {
    showToast("PDF出力対象データがありません");
    return;
  }
  const selectedGridLabel =
    exportListRangeKuwaku === ALL_GRIDS_VALUE ? "全グリッド" : kuwakuLabelForSelect(exportListRangeKuwaku);
  const categoryLabel =
    exportListRangeCategory === EXPORT_CATEGORY_ALL_VALUE
      ? "全分類"
      : `${exportListRangeCategory}: ${SPECIMEN_CATEGORY_MAP[exportListRangeCategory] || ""}`;
  const statusLabel =
    exportListRangeStatus === "complete" ? "必須完了のみ" : exportListRangeStatus === "incomplete" ? "未記入ありのみ" : "すべて";
  const specimenRangeLabel =
    exportListRangeSpecimenFrom || exportListRangeSpecimenTo
      ? `${value(exportListRangeSpecimenFrom) || "-"} 〜 ${value(exportListRangeSpecimenTo) || "-"}`
      : "指定なし";
  const html = buildListPdfHtml(records, {
    selectedGridLabel,
    categoryLabel,
    statusLabel,
    specimenRangeLabel,
  });
  const opened = openPdfPrintWindow({
    title: "化石遺物リストout＿出力.pdf",
    pageSize: "A3 landscape",
    bodyHtml: html,
  });
  if (opened) {
    showToast("リストPDFの印刷画面を開きました（保存先でPDFを選択）");
  }
}

function exportCardPdf() {
  const records = getCardExportRecords();
  if (!records.length) {
    showToast("カードPDFの出力対象データがありません");
    return;
  }
  const html = buildCardPdfHtml(records);
  const opened = openPdfPrintWindow({
    title: "化石遺物カードout＿出力.pdf",
    pageSize: "A4 portrait",
    bodyHtml: html,
  });
  if (opened) {
    showToast("カードPDFの印刷画面を開きました（保存先でPDFを選択）");
  }
}

function exportPlanPdf() {
  const groups = buildPlanPdfGroupsForExport({
    kuwakuValue: exportPlanKuwaku,
    categoryValue: exportPlanCategory,
    dateFromRaw: exportPlanDateFrom,
    dateToRaw: exportPlanDateTo,
    modeSelections: getExportPlanModeSelections(),
  });
  if (!groups.length) {
    showToast("平面図PDFの出力対象がありません（平面位置を確認してください）");
    return;
  }
  const html = buildPlanPdfHtml(groups);
  const opened = openPdfPrintWindow({
    title: "層準別平面図out＿出力.pdf",
    pageSize: "A4 portrait",
    bodyHtml: html,
  });
  if (opened) {
    showToast("平面図PDFの印刷画面を開きました（保存先でPDFを選択）");
  }
}

function buildListPdfHtml(records, options = {}) {
  const generatedAt = formatPdfGeneratedAt(new Date());
  const selectedGrid = value(options.selectedGridLabel) || "全グリッド";
  const categoryLabel = value(options.categoryLabel);
  const statusLabel = value(options.statusLabel);
  const specimenRangeLabel = value(options.specimenRangeLabel);
  const rows = records
    .map((record) => {
      const complete = isRecordDataComplete(record);
      return `
        <tr>
          <td>${escapeHtml(getRecordKuwaku(record))}</td>
          <td>${escapeHtml(getRecordTeamValue(record))}</td>
          <td>${escapeHtml(record.specimenNo || "")}</td>
          <td>${escapeHtml(formatCategoryForRecord(record))}</td>
          <td>${escapeHtml(record.nameMemo || "")}</td>
          <td>${escapeHtml(record.importantFlag || "")}</td>
          <td>${escapeHtml(record.unit || "")}</td>
          <td>${escapeHtml(formatDetailForRecord(record))}</td>
          <td>${escapeHtml(record.discoverer || "")}</td>
          <td>${escapeHtml(record.identifier || "")}</td>
          <td>${escapeHtml(formatLevelRead(record))}</td>
          <td>${escapeHtml(formatRecordAltitudeM(record))}</td>
          <td>${escapeHtml(formatPlanPosition(record))}</td>
          <td class="${complete ? "pdf-status-complete" : "pdf-status-incomplete"}">${complete ? "○" : "未記入"}</td>
        </tr>
      `;
    })
    .join("");

  return `
    <section class="pdf-page">
      <header class="pdf-header">
        <h1>化石遺物リスト</h1>
        <div class="pdf-meta">
          <span>区画: ${escapeHtml(selectedGrid)}</span>
          <span>分類: ${escapeHtml(categoryLabel || "全分類")}</span>
          <span>データ状態: ${escapeHtml(statusLabel || "すべて")}</span>
          <span>標本番号: ${escapeHtml(specimenRangeLabel || "指定なし")}</span>
          <span>出力日時: ${escapeHtml(generatedAt)}</span>
          <span>件数: ${records.length}</span>
        </div>
      </header>
      <table class="pdf-table pdf-list-table">
        <thead>
          <tr>
            <th>区画</th>
            <th>発掘班</th>
            <th>標本番号</th>
            <th>分類</th>
            <th>名称</th>
            <th>重要品</th>
            <th>ユニット</th>
            <th>サブユニット</th>
            <th>発見者</th>
            <th>判定者</th>
            <th>レベル読値</th>
            <th>標高(m)</th>
            <th>平面位置</th>
            <th>データ</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    </section>
  `;
}

function buildCardPdfHtml(records) {
  const generatedAt = formatPdfGeneratedAt(new Date());
  return records
    .map((record, index) => {
      const sectionDiagramsHtml = buildPdfImageGrid(record.sectionDiagrams, "断面図なし", "pdf-card-image-grid");
      const photosHtml = buildPdfImageGrid(record.photos, "写真なし", "pdf-card-image-grid");
      return `
        <section class="pdf-page pdf-card-page ${index < records.length - 1 ? "pdf-page-break" : ""}">
          <header class="pdf-header">
            <h1>化石遺物カード</h1>
            <div class="pdf-meta">
              <span>区画: ${escapeHtml(getRecordKuwaku(record))}</span>
              <span>標本番号: ${escapeHtml(record.specimenNo || "")}</span>
              <span>出力日時: ${escapeHtml(generatedAt)}</span>
            </div>
          </header>
          <table class="pdf-table pdf-card-table">
            <tbody>
              <tr><th>分類</th><td>${escapeHtml(formatCategoryForRecord(record))}</td><th>重要品指定</th><td>${escapeHtml(
                record.importantFlag || ""
              )}</td></tr>
              <tr><th>化石・遺物名称</th><td>${escapeHtml(record.nameMemo || "")}</td><th>簡易記載</th><td>${escapeHtml(
                record.simpleRecordFlag || "-"
              )}</td></tr>
              <tr><th>地層名</th><td>${escapeHtml(record.layerName || "")}</td><th>ユニット</th><td>${escapeHtml(
                record.unit || ""
              )}</td></tr>
              <tr><th>サブユニット</th><td>${escapeHtml(record.detail || "")}</td><th>細分</th><td>${escapeHtml(
                record.detailSub || ""
              )}</td></tr>
              <tr><th>地層中の位置</th><td colspan="3">${escapeHtml(formatLayerPosition(record))}</td></tr>
              <tr><th>発見者氏名</th><td>${escapeHtml(record.discoverer || "")}</td><th>判定者氏名</th><td>${escapeHtml(
                record.identifier || ""
              )}</td></tr>
              <tr><th>レベル読値</th><td>${escapeHtml(formatLevelRead(record))}</td><th>平面位置</th><td>${escapeHtml(
                formatPlanPosition(record)
              )}</td></tr>
              <tr><th>産出状況断面</th><td>${escapeHtml(record.occurrenceSection || "")}</td><th>産状スケッチ</th><td>${escapeHtml(
                record.occurrenceSketch || ""
              )}</td></tr>
              <tr><th>備考（観察事項など）</th><td colspan="3">${escapeHtml(record.notes || "")}</td></tr>
            </tbody>
          </table>
          <section class="pdf-image-section">
            <h2>産出状況断面図</h2>
            ${sectionDiagramsHtml}
          </section>
          <section class="pdf-image-section">
            <h2>写真</h2>
            ${photosHtml}
          </section>
        </section>
      `;
    })
    .join("");
}

function buildPlanPdfGroups() {
  const kuwakuRecords = getFilteredPlanRecords();
  if (!kuwakuRecords.length) {
    return [];
  }

  const unitValues =
    selectedPlanUnit === ALL_UNITS_VALUE
      ? collectPlanUnits(kuwakuRecords)
          .map((item) => item.value)
          .filter((valueRaw) => valueRaw !== ALL_UNITS_VALUE)
      : [selectedPlanUnit];
  const groups = [];

  unitValues.forEach((unitValue) => {
    const unitRecords =
      unitValue === ALL_UNITS_VALUE
        ? kuwakuRecords
        : kuwakuRecords.filter((record) => unitValueForSelect(record.unit) === unitValue);
    if (!unitRecords.length) {
      return;
    }

    const detailValues =
      selectedPlanDetail === ALL_DETAILS_VALUE
        ? collectPlanDetails(unitRecords)
            .map((item) => item.value)
            .filter((valueRaw) => valueRaw !== ALL_DETAILS_VALUE)
        : [selectedPlanDetail];

    detailValues.forEach((detailValue) => {
      const detailRecords =
        detailValue === ALL_DETAILS_VALUE
          ? unitRecords
          : unitRecords.filter((record) => detailValueForSelect(record.detail) === detailValue);
      if (!detailRecords.length) {
        return;
      }

      const detailSubValues =
        selectedPlanDetailSub === ALL_DETAIL_SUBS_VALUE
          ? collectPlanDetailSubs(detailRecords)
              .map((item) => item.value)
              .filter((valueRaw) => valueRaw !== ALL_DETAIL_SUBS_VALUE)
          : [selectedPlanDetailSub];

      detailSubValues.forEach((detailSubValue) => {
        const scopedRecords =
          detailSubValue === ALL_DETAIL_SUBS_VALUE
            ? detailRecords
            : detailRecords.filter((record) => detailSubValueForSelect(record.detailSub) === detailSubValue);
        const drawables = scopedRecords.map((record) => buildPlanDrawable(record)).filter(Boolean);
        if (!drawables.length) {
          return;
        }
        groups.push({
          unitValue,
          detailValue,
          detailSubValue,
          drawables,
          count: scopedRecords.length,
        });
      });
    });
  });

  return groups;
}

function buildPlanPdfGroupsForExport(filters = {}) {
  const allRecords = [...state.records].sort(compareRecordsByKuwakuThenSpecimen);
  if (!allRecords.length) {
    return [];
  }

  const kuwakuValue = value(filters.kuwakuValue);
  if (!kuwakuValue || kuwakuValue === ALL_GRIDS_VALUE) {
    return [];
  }
  const categoryValue = value(filters.categoryValue) || EXPORT_CATEGORY_ALL_VALUE;
  const scopedRecords = getRecordsByExportRangeFilters({
    kuwakuValue,
    categoryValue,
    statusValue: "all",
    dateFromRaw: filters.dateFromRaw,
    dateToRaw: filters.dateToRaw,
  });
  if (!scopedRecords.length) {
    return [];
  }

  const modeSelections = filters.modeSelections || {};
  const groups = [];
  const unitSelection = modeSelections.unit || {};
  const detailSelection = modeSelections.detail || {};
  const detailSubSelection = modeSelections.detailSub || {};

  if (unitSelection.enabled) {
    const selectedUnitValues = Array.from(normalizeSelectionSet(unitSelection.unitValues)).sort((a, b) =>
      unitLabelForSelect(a).localeCompare(unitLabelForSelect(b), "ja", { numeric: true, sensitivity: "base" })
    );
    selectedUnitValues.forEach((unitValue) => {
      const records = filterPlanRecordsForMode(scopedRecords, { unitValues: [unitValue] });
      if (!records.length) {
        return;
      }
      const drawables = records.map((record) => buildPlanDrawable(record)).filter(Boolean);
      if (!drawables.length) {
        return;
      }
      groups.push({
        kuwakuValue,
        modeLabel: "ユニット別",
        unitValue,
        detailValue: "",
        detailSubValue: "",
        unitLabel: unitLabelForSelect(unitValue),
        detailLabel: "-",
        detailSubLabel: "-",
        drawables,
        count: records.length,
      });
    });
  }

  if (detailSelection.enabled) {
    const selectedUnitValue = value(detailSelection.unitValue);
    if (selectedUnitValue) {
      const selectedDetailValues = Array.from(normalizeSelectionSet(detailSelection.detailValues));
      const unitRecords = filterPlanRecordsForMode(scopedRecords, {
        unitValues: [selectedUnitValue],
      });
      selectedDetailValues
        .sort((a, b) =>
          detailLabelForSelect(a).localeCompare(detailLabelForSelect(b), "ja", { numeric: true, sensitivity: "base" })
        )
        .forEach((detailValue) => {
          const records = filterPlanRecordsForMode(unitRecords, { detailValues: [detailValue] });
          if (!records.length) {
            return;
          }
          const drawables = records.map((record) => buildPlanDrawable(record)).filter(Boolean);
          if (!drawables.length) {
            return;
          }
          groups.push({
            kuwakuValue,
            modeLabel: "サブユニット別",
            unitValue: selectedUnitValue,
            detailValue,
            detailSubValue: "",
            unitLabel: unitLabelForSelect(selectedUnitValue),
            detailLabel: detailLabelForSelect(detailValue),
            detailSubLabel: "-",
            drawables,
            count: records.length,
          });
        });
    }
  }

  if (detailSubSelection.enabled) {
    const selectedUnitValue = value(detailSubSelection.unitValue);
    const selectedDetailValue = value(detailSubSelection.detailValue);
    if (selectedUnitValue && selectedDetailValue) {
      const selectedDetailSubValues = Array.from(normalizeSelectionSet(detailSubSelection.detailSubValues));
      const unitRecords = filterPlanRecordsForMode(scopedRecords, {
        unitValues: [selectedUnitValue],
      });
      const detailRecords = filterPlanRecordsForMode(unitRecords, {
        detailValues: [selectedDetailValue],
      });
      selectedDetailSubValues
        .sort((a, b) =>
          detailSubLabelForSelect(a).localeCompare(detailSubLabelForSelect(b), "ja", {
            numeric: true,
            sensitivity: "base",
          })
        )
        .forEach((detailSubValue) => {
          const records = filterPlanRecordsForMode(detailRecords, { detailSubValues: [detailSubValue] });
          if (!records.length) {
            return;
          }
          const drawables = records.map((record) => buildPlanDrawable(record)).filter(Boolean);
          if (!drawables.length) {
            return;
          }
          groups.push({
            kuwakuValue,
            modeLabel: "サブユニット細分別",
            unitValue: selectedUnitValue,
            detailValue: selectedDetailValue,
            detailSubValue,
            unitLabel: unitLabelForSelect(selectedUnitValue),
            detailLabel: detailLabelForSelect(selectedDetailValue),
            detailSubLabel: detailSubLabelForSelect(detailSubValue),
            drawables,
            count: records.length,
          });
        });
    }
  }

  return groups;
}

function buildPlanPdfHtml(groups) {
  const generatedAt = formatPdfGeneratedAt(new Date());

  return groups
    .map((group, index) => {
      const kuwakuValue = value(group.kuwakuValue) || selectedPlanKuwaku;
      const kuwakuLabel = kuwakuValue ? kuwakuLabelForSelect(kuwakuValue) : "-";
      const unitLabel =
        value(group.unitLabel) || (group.unitValue === ALL_UNITS_VALUE ? "全ユニット" : unitLabelForSelect(group.unitValue));
      const detailLabel =
        value(group.detailLabel) ||
        (group.detailValue === ALL_DETAILS_VALUE ? "全サブユニット" : detailLabelForSelect(group.detailValue));
      const detailSubLabel =
        value(group.detailSubLabel) ||
        (group.detailSubValue === ALL_DETAIL_SUBS_VALUE ? "全細分" : detailSubLabelForSelect(group.detailSubValue));
      const mapSvg = buildPlanPdfMapSvg(group.drawables, kuwakuValue);
      return `
        <section class="pdf-page ${index < groups.length - 1 ? "pdf-page-break" : ""}">
          <header class="pdf-header">
            <h1>層準別平面図</h1>
            <div class="pdf-meta">
              <span>区画: ${escapeHtml(kuwakuLabel)}</span>
              <span>出力モード: ${escapeHtml(value(group.modeLabel) || "-")}</span>
              <span>ユニット: ${escapeHtml(unitLabel)}</span>
              <span>サブユニット: ${escapeHtml(detailLabel)}</span>
              <span>細分: ${escapeHtml(detailSubLabel)}</span>
              <span>件数: ${group.count}</span>
              <span>出力日時: ${escapeHtml(generatedAt)}</span>
            </div>
          </header>
          <div class="pdf-plan-wrap">
            ${mapSvg}
          </div>
          <div class="pdf-plan-legend">${buildPlanLegendHtml()}</div>
        </section>
      `;
    })
    .join("");
}

function buildPlanPdfMapSvg(drawables, kuwakuRaw) {
  const verticalGrid = [100, 200, 300]
    .map((x) => `<line class="pdf-plan-grid-line" x1="${x}" y1="0" x2="${x}" y2="${PLAN_SIZE_CM}" />`)
    .join("");
  const horizontalGrid = [100, 200, 300]
    .map((y) => `<line class="pdf-plan-grid-line" x1="0" y1="${y}" x2="${PLAN_SIZE_CM}" y2="${y}" />`)
    .join("");
  const pointsSvg = drawables.map((drawable, index) => renderPlanPdfDrawableSvg(drawable, index)).join("");
  const cornerLabels = buildPlanCornerLabels(kuwakuRaw);

  return `
    <div class="pdf-plan-shell">
      <div class="pdf-plan-axis north">北</div>
      <div class="pdf-plan-axis east">東</div>
      <div class="pdf-plan-axis south">南</div>
      <div class="pdf-plan-axis west">西</div>
      <div class="pdf-plan-corner top-left">${escapeHtml(cornerLabels.topLeft)}</div>
      <div class="pdf-plan-corner top-right">${escapeHtml(cornerLabels.topRight)}</div>
      <div class="pdf-plan-corner bottom-left">${escapeHtml(cornerLabels.bottomLeft)}</div>
      <div class="pdf-plan-corner bottom-right">${escapeHtml(cornerLabels.bottomRight)}</div>
      <svg class="pdf-plan-svg" viewBox="0 0 ${PLAN_SIZE_CM} ${PLAN_SIZE_CM}" aria-label="層準別平面図">
        <rect class="pdf-plan-frame" x="0" y="0" width="${PLAN_SIZE_CM}" height="${PLAN_SIZE_CM}" />
        ${verticalGrid}
        ${horizontalGrid}
        ${pointsSvg}
      </svg>
    </div>
  `;
}

function renderPlanPdfDrawableSvg(drawable, index = 0) {
  const labelX = Math.min(PLAN_SIZE_CM - 2, drawable.x + 6);
  const labelY = Math.max(8, drawable.y - 6);
  let shapeSvg = "";
  if (drawable.type === "line") {
    shapeSvg = `<line class="pdf-plan-shape-line" x1="${drawable.x1}" y1="${drawable.y1}" x2="${drawable.x2}" y2="${drawable.y2}" stroke="${drawable.color}" />`;
  } else if (drawable.type === "rect") {
    const transform = Number.isFinite(drawable.rotationDeg)
      ? ` transform="rotate(${drawable.rotationDeg} ${drawable.x} ${drawable.y})"`
      : "";
    shapeSvg = `<rect class="pdf-plan-shape-rect" x="${drawable.left}" y="${drawable.top}" width="${drawable.width}" height="${drawable.height}" stroke="${drawable.color}"${transform} />`;
  } else if (drawable.type === "ellipse") {
    const transform = Number.isFinite(drawable.rotationDeg)
      ? ` transform="rotate(${drawable.rotationDeg} ${drawable.x} ${drawable.y})"`
      : "";
    shapeSvg = `<ellipse class="pdf-plan-shape-ellipse" cx="${drawable.x}" cy="${drawable.y}" rx="${drawable.rx}" ry="${drawable.ry}" stroke="${drawable.color}"${transform} />`;
  } else if (drawable.type === "imageQuad") {
    shapeSvg = buildPlanImageWarpSvg(drawable, `pdf-${index}`);
  } else {
    shapeSvg = `<circle cx="${drawable.x}" cy="${drawable.y}" r="5" fill="${drawable.color}" />`;
  }
  return `
    <g>
      ${shapeSvg}
      <text x="${labelX}" y="${labelY}">${escapeHtml(drawable.label || "")}</text>
    </g>
  `;
}

function buildPdfImageGrid(itemsRaw, emptyText, gridClassRaw = "") {
  const items = Array.isArray(itemsRaw) ? itemsRaw : [];
  if (!items.length) {
    return `<p class="pdf-muted">${escapeHtml(emptyText)}</p>`;
  }
  const gridClass = value(gridClassRaw).replace(/[^a-zA-Z0-9_-]/g, "");
  const classNames = gridClass ? `pdf-image-grid ${gridClass}` : "pdf-image-grid";
  return `
    <div class="${classNames}">
      ${items
        .map(
          (item) => `
            <figure>
              <img src="${item.dataUrl}" alt="${escapeHtml(item.name || "image")}" />
              <figcaption>${escapeHtml(item.caption || "")}</figcaption>
            </figure>
          `
        )
        .join("")}
    </div>
  `;
}

function openPdfPrintWindow({ title, pageSize, bodyHtml }) {
  const safeTitle = escapeHtml(title || "出力");
  const safeBody = bodyHtml || "<p>出力データがありません。</p>";
  const htmlText = `
    <!doctype html>
    <html lang="ja">
      <head>
        <meta charset="utf-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <title>${safeTitle}</title>
        <style>${buildPdfPrintStyles(pageSize)}</style>
      </head>
      <body>${safeBody}</body>
    </html>
  `;

  const printWindow = window.open("", "_blank");
  if (printWindow) {
    printWindow.document.open();
    printWindow.document.write(htmlText);
    printWindow.document.close();

    let hasPrinted = false;
    const triggerPrint = () => {
      if (hasPrinted) {
        return;
      }
      hasPrinted = true;
      try {
        printWindow.focus();
        printWindow.print();
      } catch (_error) {
        // ignore
      }
    };
    printWindow.onload = () => {
      window.setTimeout(triggerPrint, 220);
    };
    window.setTimeout(triggerPrint, 900);
    return true;
  }

  try {
    const frame = document.createElement("iframe");
    frame.style.position = "fixed";
    frame.style.right = "0";
    frame.style.bottom = "0";
    frame.style.width = "0";
    frame.style.height = "0";
    frame.style.border = "0";
    frame.setAttribute("aria-hidden", "true");
    frame.srcdoc = htmlText;
    document.body.appendChild(frame);
    frame.onload = () => {
      const frameWindow = frame.contentWindow;
      if (frameWindow) {
        try {
          frameWindow.focus();
          frameWindow.print();
        } catch (_error) {
          // ignore
        }
      }
      window.setTimeout(() => {
        frame.remove();
      }, 1500);
    };
    return true;
  } catch (_error) {
    showToast("PDF出力の画面を開けませんでした（ポップアップ設定を確認）");
    return false;
  }
}

function buildPdfPrintStyles(pageSizeRaw) {
  const pageSize = value(pageSizeRaw) || "A4 portrait";
  return `
    @page {
      size: ${pageSize};
      margin: 10mm;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: "Hiragino Kaku Gothic ProN", "Yu Gothic", sans-serif;
      color: #111827;
      font-size: 11px;
      line-height: 1.45;
    }
    .pdf-page {
      width: 100%;
    }
    .pdf-page-break {
      page-break-after: always;
    }
    .pdf-header h1 {
      margin: 0;
      font-size: 18px;
    }
    .pdf-meta {
      margin-top: 6px;
      display: flex;
      flex-wrap: wrap;
      gap: 8px 12px;
      color: #334155;
      font-size: 12px;
    }
    .pdf-table {
      width: 100%;
      border-collapse: collapse;
      margin-top: 8px;
    }
    .pdf-table th, .pdf-table td {
      border: 1px solid #9ca3af;
      padding: 4px 6px;
      vertical-align: top;
    }
    .pdf-table th {
      background: #f3f4f6;
      white-space: nowrap;
    }
    .pdf-list-table td:nth-child(14) {
      text-align: center;
      font-weight: 700;
    }
    .pdf-status-complete {
      color: #111827;
    }
    .pdf-status-incomplete {
      color: #b42318;
    }
    .pdf-image-section h2 {
      margin: 10px 0 5px;
      font-size: 12px;
    }
    .pdf-image-grid {
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(42mm, 1fr));
      gap: 6px;
    }
    .pdf-image-grid figure {
      margin: 0;
      border: 1px solid #d1d5db;
      border-radius: 4px;
      padding: 4px;
    }
    .pdf-image-grid img {
      width: 100%;
      height: auto;
      max-height: 38mm;
      object-fit: cover;
      display: block;
    }
    .pdf-image-grid figcaption {
      margin-top: 3px;
      color: #475569;
      font-size: 10px;
    }
    .pdf-card-image-grid {
      grid-template-columns: 1fr;
      gap: 8px;
    }
    .pdf-card-image-grid img {
      max-height: 72mm;
      object-fit: contain;
      background: #f8fafc;
    }
    .pdf-muted {
      color: #6b7280;
      margin: 4px 0;
    }
    .pdf-plan-wrap {
      margin-top: 8px;
      display: flex;
      justify-content: center;
    }
    .pdf-plan-shell {
      position: relative;
      width: 162mm;
      height: 162mm;
      border: 1px solid #9ca3af;
      background: #fff;
      padding: 12mm;
    }
    .pdf-plan-svg {
      width: 100%;
      height: 100%;
      display: block;
      overflow: visible;
    }
    .pdf-plan-svg .pdf-plan-frame {
      fill: #fff;
      stroke: #334155;
      stroke-width: 2;
    }
    .pdf-plan-svg .pdf-plan-grid-line {
      stroke: #94a3b8;
      stroke-width: 0.9;
      stroke-dasharray: 4 3;
    }
    .pdf-plan-svg text {
      font-size: 12px;
      font-weight: 700;
      fill: #0f172a;
    }
    .pdf-plan-svg .pdf-plan-shape-line {
      stroke-width: 4;
      fill: none;
      stroke-linecap: round;
      stroke-dasharray: none;
    }
    .pdf-plan-svg .pdf-plan-shape-rect,
    .pdf-plan-svg .pdf-plan-shape-ellipse {
      stroke-width: 3.2;
      fill: none;
    }
    .pdf-plan-axis {
      position: absolute;
      font-size: 13px;
      font-weight: 700;
      color: #1f2937;
    }
    .pdf-plan-axis.north { top: 5px; left: 50%; transform: translateX(-50%); }
    .pdf-plan-axis.south { bottom: 5px; left: 50%; transform: translateX(-50%); }
    .pdf-plan-axis.east { right: 5px; top: 50%; transform: translateY(-50%); }
    .pdf-plan-axis.west { left: 5px; top: 50%; transform: translateY(-50%); }
    .pdf-plan-corner {
      position: absolute;
      font-size: 12px;
      font-weight: 700;
      background: #eff6ff;
      border: 1px solid #93c5fd;
      border-radius: 4px;
      padding: 1px 5px;
    }
    .pdf-plan-corner.top-left { top: 10mm; left: 10mm; }
    .pdf-plan-corner.top-right { top: 10mm; right: 10mm; }
    .pdf-plan-corner.bottom-left { bottom: 10mm; left: 10mm; }
    .pdf-plan-corner.bottom-right { bottom: 10mm; right: 10mm; }
    .pdf-plan-legend {
      margin-top: 6px;
      display: flex;
      flex-wrap: wrap;
      gap: 6px;
      justify-content: center;
    }
    .plan-legend-item {
      display: inline-flex;
      align-items: center;
      gap: 4px;
      border: 1px solid #cbd5e1;
      border-radius: 999px;
      padding: 2px 8px;
      font-size: 11px;
      color: #334155;
      background: #f8fafc;
    }
    .plan-legend-dot {
      width: 10px;
      height: 10px;
      border-radius: 50%;
      border: 1px solid rgba(0, 0, 0, 0.25);
      display: inline-block;
    }
  `;
}

function formatPdfGeneratedAt(dateRaw) {
  const date = dateRaw instanceof Date ? dateRaw : new Date();
  if (!Number.isFinite(date.getTime())) {
    return "-";
  }
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const dd = String(date.getDate()).padStart(2, "0");
  const hh = String(date.getHours()).padStart(2, "0");
  const min = String(date.getMinutes()).padStart(2, "0");
  return `${yyyy}/${mm}/${dd} ${hh}:${min}`;
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
    ["層理面や鍵層名", recordFormData.get("layerRef")],
    ["地層中の位置（上/下）", recordFormData.get("layerRelative")],
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
    return "a: 分析用試料を選んだ場合は、区分を選択してください";
  }
  if (normalizeNeedFlag(value(recordFormData.get("occurrenceSection"))) === "要" && currentSectionDiagrams.length === 0) {
    return "産出状況断面が「要」の場合は、産出状況断面図添付を追加してください";
  }

  return "";
}

function isRecordDataComplete(record) {
  return getMissingRequiredKeys(record).size === 0;
}

function getMissingRequiredKeys(record) {
  const missing = new Set();
  if (!record) {
    return missing;
  }

  const kuwaku = parseKuwaku(value(record.kuwaku));
  if (!value(kuwaku.headA)) {
    missing.add("kuwakuHeadA");
  }
  if (!value(kuwaku.headB)) {
    missing.add("kuwakuHeadB");
  }
  if (!value(kuwaku.block)) {
    missing.add("kuwakuBlock");
  }
  if (!value(kuwaku.no)) {
    missing.add("kuwakuNo");
  }

  if (!value(record.levelHeight)) {
    missing.add("levelHeight");
  }
  if (!value(record.date)) {
    missing.add("date");
  }
  if (!value(record.team)) {
    missing.add("team");
  }
  if (!value(record.teamLead)) {
    missing.add("teamLead");
  }
  if (!value(record.recorder)) {
    missing.add("recorder");
  }

  const specimen = parseSpecimenNo(record.specimenNo, record.specimenPrefix, record.specimenSerial);
  if (!value(specimen.serial)) {
    missing.add("specimenSerial");
  }
  if (!value(record.nameMemo)) {
    missing.add("nameMemo");
  }
  if (!value(record.importantFlag)) {
    missing.add("importantFlag");
  }
  if (!value(record.simpleRecordFlag)) {
    missing.add("simpleRecordFlag");
  }
  if (!value(record.discoverer)) {
    missing.add("discoverer");
  }
  if (!value(record.identifier)) {
    missing.add("identifier");
  }
  if (!value(record.levelUpperCm)) {
    missing.add("levelUpperCm");
  }
  if (!value(record.levelLowerCm)) {
    missing.add("levelLowerCm");
  }
  if (!value(record.occurrenceSection)) {
    missing.add("occurrenceSection");
  }
  if (!value(record.occurrenceSketch)) {
    missing.add("occurrenceSketch");
  }
  if (!value(record.nsDir)) {
    missing.add("nsDir");
  }
  if (!value(record.nsCm)) {
    missing.add("nsCm");
  }
  if (!value(record.ewDir)) {
    missing.add("ewDir");
  }
  if (!value(record.ewCm)) {
    missing.add("ewCm");
  }

  const planSizeMode = normalizePlanSizeMode(value(record.planSizeMode));
  if (planSizeMode === "大きなもの") {
    const largeShapeType = normalizeLargeShapeType(value(record.largeShapeType));
    const isImageShape = isLargeShapeImageType(largeShapeType);
    if (!largeShapeType) {
      missing.add("largeShapeType");
    }
    if (!isImageShape && parseLargeAxisAzimuth(value(record.largeAxisDirection)) == null) {
      missing.add("largeAxisDirection");
    }
    if (largeShapeType === "直線状") {
      if (!value(record.lineLengthCm)) {
        missing.add("lineLengthCm");
      }
    } else if (largeShapeType === "長方形") {
      if (!value(record.rectSide1Cm)) {
        missing.add("rectSide1Cm");
      }
      if (!value(record.rectSide2Cm)) {
        missing.add("rectSide2Cm");
      }
    } else if (largeShapeType === "楕円") {
      if (!value(record.ellipseLongRadiusCm)) {
        missing.add("ellipseLongRadiusCm");
      }
      if (!value(record.ellipseShortRadiusCm)) {
        missing.add("ellipseShortRadiusCm");
      }
    } else if (isImageShape) {
      [
        "imgP1NsCm",
        "imgP1EwCm",
        "imgP2NsCm",
        "imgP2EwCm",
        "imgP3NsCm",
        "imgP3EwCm",
        "imgP4NsCm",
        "imgP4EwCm",
      ].forEach((key) => {
        if (!value(record[key])) {
          missing.add(key);
        }
      });
    }
  }

  const layerName = normalizeLayerName(value(record.layerName));
  if (!value(layerName)) {
    missing.add("layerName");
  }
  if (!value(record.unit)) {
    missing.add("unit");
  }
  if (!value(record.layerRef)) {
    missing.add("layerRef");
  }
  if (!value(record.layerRelative)) {
    missing.add("layerRelative");
  }
  if (!value(record.layerFromCm)) {
    missing.add("layerFromCm");
  }

  if (value(record.team) === OTHER_TEAM_NAME && !value(record.teamOther)) {
    missing.add("teamOther");
  }
  if (layerName === OTHER_LAYER_NAME && !value(extractOtherLayerText(layerName))) {
    missing.add("layerOther");
  }

  const specimenPrefix = normalizeSpecimenPrefix(value(record.specimenPrefix));
  if (specimenPrefix === "a" && !normalizeAnalysisType(value(record.analysisType))) {
    missing.add("analysisType");
  }

  const sectionDiagrams = Array.isArray(record.sectionDiagrams) ? record.sectionDiagrams : [];
  if (normalizeNeedFlag(value(record.occurrenceSection)) === "要" && !sectionDiagrams.length) {
    missing.add("sectionDiagrams");
  }
  if (sectionDiagrams.length) {
    if (normalizeChecklistChecked(record.sectionDiagramDistanceChecked) !== "1") {
      missing.add("sectionDiagramDistanceChecked");
    }
    if (normalizeChecklistChecked(record.sectionDiagramHorizonChecked) !== "1") {
      missing.add("sectionDiagramHorizonChecked");
    }
    if (normalizeChecklistChecked(record.sectionDiagramLayerFaciesChecked) !== "1") {
      missing.add("sectionDiagramLayerFaciesChecked");
    }
  }
  const photos = Array.isArray(record.photos) ? record.photos : [];
  if (photos.length) {
    if (normalizeChecklistChecked(record.photoClinometerChecked) !== "1") {
      missing.add("photoClinometerChecked");
    }
    if (normalizeChecklistChecked(record.photoRulerChecked) !== "1") {
      missing.add("photoRulerChecked");
    }
  }

  return missing;
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

function normalizeChecklistChecked(valueRaw) {
  const text = value(valueRaw).toLowerCase();
  if (text === "1" || text === "true" || text === "on" || text === "yes" || text === "checked" || text === "○" || text === "◯") {
    return "1";
  }
  return "";
}

function validateAttachmentChecklistForSave(formData) {
  if (!(formData instanceof FormData)) {
    return "";
  }
  if (currentSectionDiagrams.length > 0 && !areSectionDiagramChecklistComplete(formData)) {
    return "産出状況断面図のチェックをすべて入れてください";
  }
  if (currentPhotos.length > 0 && !arePhotoChecklistComplete(formData)) {
    return "写真添付のチェックをすべて入れてください";
  }
  return "";
}

function areSectionDiagramChecklistComplete(formData) {
  return (
    normalizeChecklistChecked(formData.get("sectionDiagramDistanceChecked")) === "1" &&
    normalizeChecklistChecked(formData.get("sectionDiagramHorizonChecked")) === "1" &&
    normalizeChecklistChecked(formData.get("sectionDiagramLayerFaciesChecked")) === "1"
  );
}

function arePhotoChecklistComplete(formData) {
  return (
    normalizeChecklistChecked(formData.get("photoClinometerChecked")) === "1" &&
    normalizeChecklistChecked(formData.get("photoRulerChecked")) === "1"
  );
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

function normalizePlanSizeMode(valueRaw) {
  const mode = value(valueRaw);
  if (mode === "大きなもの" || mode === "大きいもの") {
    return "大きなもの";
  }
  return "通常";
}

function normalizeLargeShapeType(valueRaw) {
  const raw = value(valueRaw);
  const normalizedRaw = typeof raw.normalize === "function" ? raw.normalize("NFC") : raw;
  let text = normalizeLargeShapeLabel(normalizedRaw);
  if (text === "ゾウ下顎臼歯" || text === "ゾウ上顎臼歯") {
    return "";
  }
  if (text === "直線状" || text === "長方形" || text === "楕円" || largeShapeImagePathMap.has(text)) {
    return text;
  }
  return "";
}

function isLargeShapeImageType(shapeTypeRaw) {
  return largeShapeImagePathMap.has(normalizeLargeShapeType(shapeTypeRaw));
}

function getLargeShapeImagePath(shapeTypeRaw) {
  const shapeType = normalizeLargeShapeType(shapeTypeRaw);
  if (!shapeType) {
    return "";
  }
  const fallbackList = Array.isArray(LARGE_SHAPE_IMAGE_FALLBACK_PATHS[shapeType])
    ? LARGE_SHAPE_IMAGE_FALLBACK_PATHS[shapeType]
    : [];
  for (const fallback of fallbackList) {
    const safe = toSafeAssetUrl(fallback);
    if (safe) {
      return safe;
    }
  }
  return toSafeAssetUrl(largeShapeImagePathMap.get(shapeType) || "");
}

function getLargeShapeImagePathCandidates(shapeTypeRaw, imagePathRaw = "") {
  const candidates = [];
  const pushCandidate = (pathRaw) => {
    const safe = toSafeAssetUrl(pathRaw);
    if (!safe || candidates.includes(safe)) {
      return;
    }
    candidates.push(safe);
  };
  const pushInlineCandidate = (pathRaw) => {
    const inline = getInlineLargeShapeDataUrl(pathRaw);
    if (!inline) {
      return;
    }
    pushCandidate(inline);
  };
  const shapeType = normalizeLargeShapeType(shapeTypeRaw);
  const explicitPath = toSafeAssetUrl(imagePathRaw);
  const hasExplicitPath = Boolean(explicitPath);
  if (hasExplicitPath) {
    pushInlineCandidate(explicitPath);
    pushCandidate(explicitPath);
  }
  if (shapeType) {
    const fallbackList = LARGE_SHAPE_IMAGE_FALLBACK_PATHS[shapeType] || [];
    const mappedPath = largeShapeImagePathMap.get(shapeType) || "";
    if (!hasExplicitPath) {
      fallbackList.forEach((pathRaw) => {
        pushInlineCandidate(pathRaw);
        pushCandidate(pathRaw);
      });
      pushInlineCandidate(mappedPath);
      pushCandidate(mappedPath);
    } else {
      pushInlineCandidate(mappedPath);
      pushCandidate(mappedPath);
      fallbackList.forEach((pathRaw) => {
        pushInlineCandidate(pathRaw);
        pushCandidate(pathRaw);
      });
    }
  }
  return candidates;
}

function normalizeLargeAxisDirection(valueRaw) {
  const text = value(valueRaw)
    .toUpperCase()
    .replace(/\s+/g, "")
    .replace(/[°度]/g, "")
    .replace(/[-_]/g, "");
  if (!text) {
    return "";
  }
  if (text === "NS" || text === "SN") {
    return "NS";
  }
  if (text === "EW" || text === "WE") {
    return "EW";
  }
  const matched = text.match(/^([NS])(\d+(?:\.\d+)?)([EW])$/);
  if (!matched) {
    return text;
  }
  const [, ns, degreeRaw, ew] = matched;
  const degree = Number(degreeRaw);
  if (!Number.isFinite(degree)) {
    return text;
  }
  const degreeText = Number.isInteger(degree) ? String(degree) : String(degree).replace(/\.?0+$/, "");
  return `${ns}${degreeText}${ew}`;
}

function normalizeDirectionValue(group, valueRaw) {
  if (group === "ns") {
    return normalizeNsDir(valueRaw);
  }
  if (group === "ew") {
    return normalizeEwDir(valueRaw);
  }
  if (group === "line1Ns" || group === "line2Ns") {
    return normalizeNsDir(valueRaw);
  }
  if (group === "line1Ew" || group === "line2Ew") {
    return normalizeEwDir(valueRaw);
  }
  if (group === "importantFlag") {
    return normalizeHasFlag(valueRaw) || "無";
  }
  if (group === "simpleRecordFlag") {
    return normalizeCircleDashFlag(valueRaw);
  }
  if (group === "occurrenceSection" || group === "occurrenceSketch") {
    return normalizeNeedFlag(valueRaw);
  }
  if (group === "layerRelative") {
    return normalizeLayerRelative(valueRaw);
  }
  if (group === "planSizeMode") {
    return normalizePlanSizeMode(valueRaw);
  }
  if (group === "largeShapeType") {
    return normalizeLargeShapeType(valueRaw);
  }
  return value(valueRaw);
}

function normalizeLayerRelative(valueRaw) {
  const text = value(valueRaw);
  if (text === "上") {
    return "上";
  }
  if (text === "下") {
    return "下";
  }
  const hasUpper = text.includes("上");
  const hasLower = text.includes("下");
  if (hasUpper && !hasLower) {
    return "上";
  }
  if (hasLower && !hasUpper) {
    return "下";
  }
  return "";
}

function formatLevelRead(record) {
  const upper = value(record?.levelUpperCm);
  const lower = value(record?.levelLowerCm);
  if (!upper && !lower) {
    return "";
  }
  return `${formatCmValue(upper, "-")} / ${formatCmValue(lower, "-")}`;
}

function getRecordAltitudeMValue(record) {
  const levelHeightM = parseDistanceToCm(getRecordLevelHeight(record));
  const lowerCm = parseDistanceToCm(record?.levelLowerCm);
  if (levelHeightM == null || lowerCm == null) {
    return null;
  }
  const altitudeM = levelHeightM - lowerCm / 100;
  return Number.isFinite(altitudeM) ? altitudeM : null;
}

function formatRecordAltitudeM(record) {
  const altitudeM = getRecordAltitudeMValue(record);
  if (altitudeM == null) {
    return "";
  }
  return altitudeM.toFixed(3).replace(/\.?0+$/, "");
}

function formatLayerPosition(record) {
  const ref = value(record?.layerRef);
  const fromCm = formatCmValue(record?.layerFromCm);
  const relative = value(record?.layerRelative);
  const line1 = ref;
  let line2 = "";
  if (relative && fromCm) {
    line2 = `${relative} に ${fromCm}`;
  } else if (relative) {
    line2 = relative;
  } else if (fromCm) {
    line2 = fromCm;
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

function formatCmValue(cmRaw, fallback = "") {
  const text = value(cmRaw);
  if (!text) {
    return fallback;
  }
  if (/^[-ー－]+$/.test(text)) {
    return text;
  }
  const normalized = text.replace(/\s*(cm|㎝)$/i, "");
  if (!normalized) {
    return fallback;
  }
  return `${normalized}cm`;
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
