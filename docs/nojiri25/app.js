const STORAGE_KEY = "nojiri-kaseki-mobile-localonly-v1";
const CLOUD_ENDPOINT_KEY = "nojiri-kaseki-cloud-endpoint-localonly-v1";
const CLOUD_CLIENT_ID_KEY = "nojiri-kaseki-cloud-client-id-localonly-v1";
const TOTAL_STATION_SETUP_KEY = "nojiri-total-station-setup-v1";
const DEFAULT_CLOUD_ENDPOINT = "";
const CLOUD_PULL_INTERVAL_MS = 20000;
const CLOUD_SAVE_DEBOUNCE_MS = 900;
const CLOUD_AUTO_PULL_ENABLED = false;
const TEAM_ROSTER_FILE_NAME = "第25次班長記載名簿.xlsx";
const TEAM_ROSTER_DAY_COL_START = 10;
const TEAM_ROSTER_DAY_COL_END = 41;
const TEAM_ROSTER_TEAM_COL = 45;
const TEAM_ROSTER_ROLE_COL = 49;
const TEAM_ROSTER_NAME_COL = 3;
const DEFAULT_SPECIMEN_PREFIX = "m";

function normalizeAlphanumericWidth(inputText) {
  return String(inputText)
    .replace(/[０-９Ａ-Ｚａ-ｚ]/g, (char) => String.fromCharCode(char.charCodeAt(0) - 0xfee0))
    .replace(/－/g, "-");
}

function normalizeTypedAlphanumericWidth(event) {
  if (event.isComposing) {
    return;
  }
  const target = event.target;
  if (!(target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement)) {
    return;
  }
  if (target instanceof HTMLInputElement && ["file", "checkbox", "radio", "date", "time", "color", "range", "button", "submit", "reset", "hidden"].includes(target.type)) {
    return;
  }
  const normalized = normalizeAlphanumericWidth(target.value);
  if (normalized === target.value) {
    return;
  }
  const selectionStart = target.selectionStart;
  const selectionEnd = target.selectionEnd;
  target.value = normalized;
  if (selectionStart != null && selectionEnd != null) {
    try {
      target.setSelectionRange(selectionStart, selectionEnd);
    } catch (_error) {
      // number inputなど、選択範囲を設定できない入力欄では値の変換だけを行う。
    }
  }
}

document.addEventListener("input", normalizeTypedAlphanumericWidth, true);
document.addEventListener("change", normalizeTypedAlphanumericWidth, true);
document.addEventListener("compositionend", normalizeTypedAlphanumericWidth, true);

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
const OUTPUT_CELL_EDIT_LABELS = {
  kuwaku: "区画",
  team: "発掘班",
  date: "日付",
  specimenNo: "標本番号",
  category: "分類",
  nameMemo: "名称",
  importantFlag: "重要品指定",
  unit: "ユニット",
  detail: "サブユニット",
  discoverer: "発見者",
  identifier: "判定者",
  levelRead: "レベル読値",
  altitudeM: "標高（m）",
  occurrenceSection: "産出状況断面",
  occurrenceSketch: "産状スケッチ",
  position: "平面位置",
  notes: "備考",
};
const OUTPUT_LIST_COLUMN_DEFS = [
  { key: "kuwaku", label: "区画" },
  { key: "team", label: "発掘班" },
  { key: "date", label: "日付" },
  { key: "dataStatus", label: "データ" },
  { key: "specimenNo", label: "標本番号" },
  { key: "category", label: "分類" },
  { key: "nameMemo", label: "名称" },
  { key: "importantFlag", label: "重要品指定" },
  { key: "unit", label: "ユニット" },
  { key: "detail", label: "サブユニット" },
  { key: "discoverer", label: "発見者" },
  { key: "identifier", label: "判定者" },
  { key: "levelRead", label: "レベル読値" },
  { key: "altitudeM", label: "標高(m)" },
  { key: "occurrenceSection", label: "産出状況断面" },
  { key: "occurrenceSketch", label: "産状スケッチ" },
  { key: "position", label: "平面位置" },
  { key: "notes", label: "備考" },
  { key: "actions", label: "操作" },
];
const CUSTOM_LARGE_SHAPE_TYPE = "カスタム画像";
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
  scribe: "記載者",
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
  multiPoints: "平面位置（複数点）",
  positionMethod: "平面位置の測定方法",
  tsStationPeg: "TS設置点杭名称",
  tsStationXNorthM: "TS設置点x（南正）",
  tsStationYEastM: "TS設置点y（西正）",
  tsStationAltitudeM: "TS設置点標高",
  tsBacksightPeg: "TS後視点杭名称",
  tsBacksightXNorthM: "TS後視点x（南正）",
  tsBacksightYEastM: "TS後視点y（西正）",
  tsBacksightAltitudeM: "TS後視点標高",
  tsInstrumentHeightM: "TS機械高",
  tsTargetHeightM: "TS目標高",
  tsObservationMode: "TS標本位置入力方法",
  tsPointXNorthM: "TS設置点から南への距離（南正）",
  tsPointYEastM: "TS設置点から西への距離",
  tsPointAltitudeM: "TS標本z標高",
  tsSlopeDistanceM: "TS斜距離",
  tsInclinationDeg: "TS傾斜度",
  tsInclinationMin: "TS傾斜分",
  tsInclinationSec: "TS傾斜秒",
  tsDirectionDeg: "TS方向角度",
  tsDirectionMin: "TS方向角分",
  tsDirectionSec: "TS方向角秒",
  largeShapeType: "大きなもの形状",
  largeAxisDirection: "長軸・長辺・長半径方向（例:N30W）",
  largeAxisPlungeDeg: "プランジ角（度）",
  largeAxisPlungeDir8: "プランジ方向（8方位）",
  planeStrikeDirection: "面の走向（例:N30E）",
  planeDipDeg: "面の傾斜（度）",
  planeDipDir8: "面の傾斜方向（8方位）",
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
  imgRotateDeg: "画像回転角（北から）",
  imgFrameWidthCm: "画像外枠 辺1",
  imgFrameHeightCm: "画像外枠 辺2",
  imgSkewXDeg: "画像 横変形角",
  imgSkewYDeg: "画像 縦変形角",
  imgLockAspectRatio: "画像 縦横比固定",
  customLargeImageName: "画像名",
  customLargeImageDataUrl: "画像ファイル",
  customLargeImageAspect: "画像 縦横比",
  layerName: "地層名",
  layerOther: "地層名（その他）",
  unit: "ユニット",
  layerColor: "色",
  layerLithology: "岩相",
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
  { key: "layerColor", label: "色" },
  { key: "layerLithology", label: "岩相" },
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
const DEFAULT_LAYER_NAME = "2.立が鼻砂部層";
const LAYER_COLOR_OPTIONS = [
  "黒色", "暗灰色", "灰色", "暗灰褐色", "灰褐色", "褐色",
  "明褐色", "茶褐色", "黄褐色", "緑灰色", "青灰色", "紫灰色",
];

function getLayerColor(record) {
  const explicit = value(record?.layerColor);
  if (explicit) return explicit;
  const legacy = value(record?.layerFacies);
  return [...LAYER_COLOR_OPTIONS].sort((a, b) => b.length - a.length).find((color) => legacy.startsWith(color)) || "";
}

function getLayerLithology(record) {
  const explicit = value(record?.layerLithology);
  if (explicit) return explicit;
  const legacy = value(record?.layerFacies);
  const color = getLayerColor(record);
  return color && legacy.startsWith(color) ? legacy.slice(color.length) : legacy;
}

function composeLayerFacies(colorRaw, lithologyRaw) {
  return `${value(colorRaw)}${value(lithologyRaw)}`;
}
const LEGACY_LAYER_NAME_ALIASES = {
  "2.立が花砂部層": "2.立が鼻砂部層",
};
const PRESET_TEAMS = ["1", "2", "3", "4", "その他"];
const OTHER_TEAM_NAME = "その他";
const DEFAULT_KUWAKU_HEAD_A = "25";
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
const CUSTOM_IMAGE_SHAPE_CANVAS_DILATE_ITERATIONS = 1;
const IMAGE_QUAD_TILT_Z_SCALE = 0.38;
const IMAGE_QUAD_TILT_Z_LIMIT_M = 1.2;
const PLAN_IMAGE_TINTED_DATA_URL_MAX_LENGTH = 180000;
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
    scribe: "",
    updatedAt: "",
  },
  records: [],
});

function createInitialOutputColumnVisibilityMap() {
  const visibilityMap = {};
  OUTPUT_LIST_COLUMN_DEFS.forEach((column) => {
    visibilityMap[column.key] = true;
  });
  return visibilityMap;
}

let stateNeedsRewriteAfterLoad = false;
let state = loadState();
let teamRosterAssignmentMap = new Map();
let teamRosterLoaded = false;
let editingRecordId = null;
let activeEditRecordContext = null;
let currentSectionDiagrams = [];
let currentPhotos = [];
let selectedCardRecordId = "";
let selectedOutputKuwaku = ALL_GRIDS_VALUE;
let selectedOutputCategory = EXPORT_CATEGORY_ALL_VALUE;
let selectedOutputStatus = "all";
let selectedOutputDate = "";
let outputSearchText = "";
let outputFilterMemory = {
  kuwaku: ALL_GRIDS_VALUE,
  category: EXPORT_CATEGORY_ALL_VALUE,
  status: "all",
  date: "",
  searchText: "",
};
let outputListColumnVisibility = createInitialOutputColumnVisibilityMap();
let activeOutputCellEdit = null;
let selectedPlanKuwaku = "";
let selectedPlanCategory = EXPORT_CATEGORY_ALL_VALUE;
let selectedPlanUnit = "";
let selectedPlanDetail = ALL_DETAILS_VALUE;
let selectedPlanDetailSub = ALL_DETAIL_SUBS_VALUE;
let selectedViewerKuwaku = ALL_GRIDS_VALUE;
let selectedViewerCategory = EXPORT_CATEGORY_ALL_VALUE;
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
const rosterStatusEl = document.getElementById("roster-status");
const recordIdInput = document.getElementById("record-id-input");
const recordSubmitBtn = document.getElementById("record-submit-btn");
const recordCopyToInputBtn = document.getElementById("record-copy-to-input-btn");
const recordCopyToInputTopBtn = document.getElementById("record-copy-to-input-top-btn");
const recordPrevBtn = document.getElementById("record-prev-btn");
const recordNextBtn = document.getElementById("record-next-btn");
const recordResetBtn = document.getElementById("record-reset-btn");
const recordNewBtn = document.getElementById("record-new-btn");
const recordTableBody = document.getElementById("record-table-body");
const editRecordTableBody = document.getElementById("edit-record-table-body");
const outputListBody = document.getElementById("output-list-body");
const outputListTable = document.getElementById("output-list-table");
const cardOutputList = document.getElementById("card-output-list");
const outputKuwakuSelect = document.getElementById("output-kuwaku-select");
const outputCategorySelect = document.getElementById("output-category-select");
const outputStatusSelect = document.getElementById("output-status-select");
const outputDateSelect = document.getElementById("output-date-select");
const outputSearchInput = document.getElementById("output-search-input");
const outputFilterSummary = document.getElementById("output-filter-summary");
const outputColumnToggleRow = document.getElementById("output-column-toggle-row");
const planKuwakuSelect = document.getElementById("plan-kuwaku-select");
const planCategorySelect = document.getElementById("plan-category-select");
const planUnitSelect = document.getElementById("plan-unit-select");
const planDetailSelect = document.getElementById("plan-detail-select");
const planDetailSubSelect = document.getElementById("plan-detail-sub-select");
const planMapLegend = document.getElementById("plan-map-legend");
const planMapWrap = document.getElementById("plan-map-wrap");
const planKuwakuInfo = document.getElementById("plan-kuwaku-info");
const viewerKuwakuSelect = document.getElementById("viewer-kuwaku-select");
const viewerCategorySelect = document.getElementById("viewer-category-select");
const viewerUnitSelect = document.getElementById("viewer-unit-select");
const viewerDetailSelect = document.getElementById("viewer-detail-select");
const viewerDetailSubSelect = document.getElementById("viewer-detail-sub-select");
const viewerKuwakuInfo = document.getElementById("viewer-kuwaku-info");
const viewerMapLegend = document.getElementById("viewer-map-legend");
const viewerCanvasWrap = document.getElementById("viewer-canvas-wrap");
const viewerTooltip = document.getElementById("viewer-tooltip");
const viewerStatus = document.getElementById("viewer-status");
const viewerViewTopBtn = document.getElementById("viewer-view-top-btn");
const viewerViewSeBtn = document.getElementById("viewer-view-se-btn");
const viewerViewEastBtn = document.getElementById("viewer-view-east-btn");
const viewerViewWestBtn = document.getElementById("viewer-view-west-btn");
const viewerViewSouthBtn = document.getElementById("viewer-view-south-btn");
const viewerViewNorthBtn = document.getElementById("viewer-view-north-btn");
const viewerZScaleInput = document.getElementById("viewer-z-scale-input");
const viewerZScaleValue = document.getElementById("viewer-z-scale-value");
const largeAxisDirectionRow = document.getElementById("large-axis-direction-row");
const largeAxisPlungeRow = document.getElementById("large-axis-plunge-row");
const largeAxisPlungeDirRow = document.getElementById("large-axis-plunge-dir-row");
const planeAttitudeRow = document.getElementById("plane-attitude-row");
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
const largeAxisPlungeInput = document.getElementById("large-axis-plunge-input");
const largeAxisPlungeDirInput = document.getElementById("large-axis-plunge-dir-input");
const planeStrikeInput = document.getElementById("plane-strike-input");
const planeDipInput = document.getElementById("plane-dip-input");
const planeDipDirInput = document.getElementById("plane-dip-dir-input");
const largeShapeImageButtons = document.getElementById("large-shape-image-buttons");
const largeShapeImagePreview = document.getElementById("large-shape-image-preview");
const largeShapeImagePreviewTitle = document.getElementById("large-shape-image-preview-title");
const largeShapeImagePreviewImg = document.getElementById("large-shape-image-preview-img");
const customLargeImageControls = document.getElementById("custom-large-image-controls");
const customLargeImageNameInput = document.getElementById("custom-large-image-name-input");
const customLargeImageFileInput = document.getElementById("custom-large-image-file-input");
const customLargeImageDataUrlInput = document.getElementById("custom-large-image-data-url-input");
const customLargeImageAspectInput = document.getElementById("custom-large-image-aspect-input");
const customLargeImageClearBtn = document.getElementById("custom-large-image-clear-btn");
const customLargeImageStatus = document.getElementById("custom-large-image-status");
const line1NsDirInput = document.getElementById("line1-ns-dir-input");
const line1EwDirInput = document.getElementById("line1-ew-dir-input");
const line2NsDirInput = document.getElementById("line2-ns-dir-input");
const line2EwDirInput = document.getElementById("line2-ew-dir-input");
const multiPointSection = document.getElementById("multi-point-section");
const multiPointRows = document.getElementById("multi-point-rows");
const multiPointAddBtn = document.getElementById("multi-point-add-btn");
const largeShapeSection = document.getElementById("large-shape-section");
const largeShapePanels = document.querySelectorAll(".large-shape-panel[data-large-shape-panel]");
const layerTabButtons = document.querySelectorAll(".layer-tab");
const layerNameInput = document.getElementById("layer-name-input");
const layerOtherInput = document.getElementById("layer-other-input");
const unitInput = document.getElementById("unit-input");
const unitTabs = document.getElementById("unit-tabs");
const teamOtherInput = document.getElementById("team-other-input");

const sectionDiagramCameraInput = document.getElementById("section-diagram-camera-input");
const sectionDiagramInput = document.getElementById("section-diagram-input");
const sectionDiagramList = document.getElementById("section-diagram-list");
const photoCameraInput = document.getElementById("photo-camera-input");
const photoInput = document.getElementById("photo-input");
const photoList = document.getElementById("photo-list");
const sectionDiagramCameraBtn = document.getElementById("section-diagram-camera-btn");
const photoCameraBtn = document.getElementById("photo-camera-btn");
const cameraCaptureModal = document.getElementById("camera-capture-modal");
const cameraCaptureVideo = document.getElementById("camera-capture-video");
const cameraCaptureStatus = document.getElementById("camera-capture-status");
const cameraCaptureShutterBtn = document.getElementById("camera-capture-shutter-btn");
const cameraCaptureCloseBtn = document.getElementById("camera-capture-close-btn");
const cameraCaptureCancelBtn = document.getElementById("camera-capture-cancel-btn");

let activeCameraStream = null;
let activeCameraDestination = "photo";

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
const positionPreviewBtn = document.getElementById("position-preview-btn");
const positionPreviewModal = document.getElementById("position-preview-modal");
const positionPreviewCloseBtn = document.getElementById("position-preview-close-btn");
const positionPreviewMeta = document.getElementById("position-preview-meta");
const positionPreviewMap = document.getElementById("position-preview-map");
const tsStationPointSelect = document.getElementById("ts-station-point-select");
const tsBacksightPointSelect = document.getElementById("ts-backsight-point-select");
const cellEditModal = document.getElementById("cell-edit-modal");
const cellEditForm = document.getElementById("cell-edit-form");
const cellEditTitle = document.getElementById("cell-edit-title");
const cellEditMeta = document.getElementById("cell-edit-meta");
const cellEditFields = document.getElementById("cell-edit-fields");
const cellEditCloseBtn = document.getElementById("cell-edit-close-btn");
const cellEditCancelBtn = document.getElementById("cell-edit-cancel-btn");
const cellEditSaveBtn = document.getElementById("cell-edit-save-btn");
let hoverEditMenuEl = null;
let hoverEditMenuRecordId = "";
let hoverEditMenuKuwaku = "";
let positionPreviewRecordOverride = null;
const TOUCH_LONG_PRESS_MS = 520;
const TOUCH_LONG_PRESS_MOVE_THRESHOLD_PX = 16;
const viewerTouchLongPressState = {
  pointerId: null,
  startX: 0,
  startY: 0,
  timer: 0,
  triggered: false,
};

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
  shiftPanActive: false,
  defaultLeftMouseAction: null,
};
const planLargeShapeImageCache = new Map();
const planLargeShapeTintedCanvasCache = new Map();
const planLargeShapeTintedDataUrlCache = new Map();

initialize();

function initialize() {
  bindEvents();
  initializeTotalStationGridPointSelectors();
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
  void loadTeamRosterFromDefaultFile();
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
  if (positionPreviewBtn) {
    positionPreviewBtn.addEventListener("click", () => {
      openPositionPreviewModal();
    });
  }
  if (positionPreviewCloseBtn) {
    positionPreviewCloseBtn.addEventListener("click", () => {
      closePositionPreviewModal();
    });
  }
  if (positionPreviewModal) {
    positionPreviewModal.addEventListener("click", (event) => {
      if (event.target === positionPreviewModal) {
        closePositionPreviewModal();
      }
    });
  }
  if (cellEditCloseBtn) {
    cellEditCloseBtn.addEventListener("click", () => {
      closeOutputCellEditModal();
    });
  }
  if (cellEditCancelBtn) {
    cellEditCancelBtn.addEventListener("click", () => {
      closeOutputCellEditModal();
    });
  }
  if (cellEditModal) {
    cellEditModal.addEventListener("click", (event) => {
      if (event.target === cellEditModal) {
        closeOutputCellEditModal();
      }
    });
  }
  if (cellEditForm) {
    const handleCellEditSave = (event) => {
      event.preventDefault();
      saveOutputCellEditFromModal();
    };
    cellEditForm.addEventListener("submit", handleCellEditSave);
    cellEditSaveBtn?.addEventListener("click", handleCellEditSave);
  }
  document.addEventListener("keydown", (event) => {
    if (event.key !== "Escape") {
      return;
    }
    hideHoverEditMenu();
    closeOutputCellEditModal();
    if (positionPreviewModal && !positionPreviewModal.classList.contains("hidden")) {
      closePositionPreviewModal();
    }
  });
  document.addEventListener("pointerdown", (event) => {
    if (!hoverEditMenuEl || hoverEditMenuEl.hidden) {
      return;
    }
    const target = event.target;
    if (target instanceof Node && hoverEditMenuEl.contains(target)) {
      return;
    }
    hideHoverEditMenu();
  });
  document.addEventListener(
    "scroll",
    () => {
      hideHoverEditMenu();
    },
    true
  );
  if (viewerCanvasWrap) {
    viewerCanvasWrap.addEventListener("click", (event) => {
      const target = event.target;
      if (!(target instanceof Element)) {
        return;
      }
      const fallbackButton = target.closest("button[data-action='viewer-open-plan']");
      if (!fallbackButton || !viewerCanvasWrap.contains(fallbackButton)) {
        return;
      }
      event.preventDefault();
      selectedPlanKuwaku = selectedViewerKuwaku;
      selectedPlanCategory = selectedViewerCategory;
      selectedPlanUnit = selectedViewerUnit;
      selectedPlanDetail = selectedViewerDetail;
      selectedPlanDetailSub = selectedViewerDetailSub;
      setActiveTab("plan-tab");
    });
  }

  recordForm.addEventListener("input", handleRecordFormFieldEdit);
  recordForm.addEventListener("change", handleRecordFormFieldEdit);
  if (customLargeImageFileInput) {
    customLargeImageFileInput.addEventListener("change", async (event) => {
      const input = event.target;
      if (!(input instanceof HTMLInputElement)) {
        return;
      }
      const file = Array.from(input.files || [])[0];
      if (!file) {
        return;
      }
      await setCustomLargeImageFromFile(file);
      input.value = "";
    });
  }
  if (customLargeImageClearBtn) {
    customLargeImageClearBtn.addEventListener("click", () => {
      clearCustomLargeImageFields();
      syncLargeShapeImagePreviewForCurrentForm();
      updateEditMissingRequiredHighlights();
    });
  }
  if (editTabPanel) {
    editTabPanel.addEventListener("input", () => {
      updateEditMissingRequiredHighlights();
      renderRecordTable();
    });
    editTabPanel.addEventListener("change", () => {
      updateEditMissingRequiredHighlights();
      renderRecordTable();
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
      renderRecordTable();
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
      scribe: value(formData.get("scribe")),
      updatedAt: nowIso(),
    };
    selectedPlanKuwaku = kuwakuValueForSelect(nextSiteKuwaku);
    persist("区画（グリッド）情報を保存しました");
    renderRecordTable();
    renderOutputs();
  });

  siteForm.elements.team.addEventListener("change", () => {
    syncTeamOtherInput(siteForm.elements.team.value);
    applyTeamRosterAutofill();
  });
  siteForm.elements.date.addEventListener("change", () => {
    applyTeamRosterAutofill();
  });
  ["kuwakuHeadA", "kuwakuHeadB", "kuwakuBlock", "kuwakuNo"].forEach((name) => {
    const input = siteForm.elements[name];
    if (!(input instanceof Element)) {
      return;
    }
    input.addEventListener("input", updateDuplicateSpecimenWarning);
    input.addEventListener("change", updateDuplicateSpecimenWarning);
    input.addEventListener("input", () => {
      if (isTotalStationMeasurementSelected()) applyTotalStationPosition();
    });
    input.addEventListener("change", () => {
      if (isTotalStationMeasurementSelected()) applyTotalStationPosition();
    });
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
  if (multiPointAddBtn && multiPointRows) {
    multiPointAddBtn.addEventListener("click", () => {
      multiPointRows.append(createMultiPointRowElement(createDefaultPlanMultiPoint()));
      syncMultiPointRemoveButtonState();
      updateMultiPointCalculationResults();
    });
  }
  if (multiPointRows) {
    multiPointRows.addEventListener("input", updateMultiPointCalculationResults);
    multiPointRows.addEventListener("change", updateMultiPointCalculationResults);
    multiPointRows.addEventListener("click", (event) => {
      const target = event.target;
      if (!(target instanceof Element)) {
        return;
      }
      const removeButton = target.closest("[data-multi-point-remove]");
      if (!(removeButton instanceof HTMLButtonElement)) {
        return;
      }
      const row = removeButton.closest("[data-multi-point-row]");
      if (!(row instanceof HTMLElement)) {
        return;
      }
      const rows = multiPointRows.querySelectorAll("[data-multi-point-row]");
      if (rows.length <= 1) {
        return;
      }
      row.remove();
      syncMultiPointRemoveButtonState();
      updateMultiPointCalculationResults();
    });
  }

  layerTabButtons.forEach((button) => {
    button.addEventListener("click", () => {
      clearLayerSavedTabState();
      layerOtherInput.classList.remove("saved-carry-value");
      activateLayerTab(button.dataset.layer);
    });
  });

  if (unitInput) {
    unitInput.addEventListener("input", syncUnitTabSelection);
  }
  if (unitTabs) {
    unitTabs.addEventListener("click", (event) => {
      const button = event.target.closest(".unit-tab");
      if (!button || !unitInput) return;
      unitInput.value = value(button.dataset.unit);
      unitInput.classList.remove("saved-carry-value");
      syncUnitTabSelection();
      updateEditMissingRequiredHighlights();
      renderRecordTable();
    });
  }

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
    const found = isEditTab
      ? findRecordByEditContext(editingId, activeEditRecordContext?.recordIndex, activeEditRecordContext?.recordSnapshot)
      : findRecord(editingId);
    if (isEditTab && !found) {
      showToast("編集対象が見つかりません。リストから編集を選び直してください");
      return;
    }
    const specimenNo = buildSpecimenNo(specimenPrefix, specimenSerial);
    const parsedKuwakuForDuplicateCheck = parseKuwaku(recordKuwaku);
    const hasKuwakuForDuplicateCheck = value(parsedKuwakuForDuplicateCheck.block) && value(parsedKuwakuForDuplicateCheck.no);
    const duplicateRecord = hasKuwakuForDuplicateCheck
      ? findDuplicateRecordByKuwakuAndSpecimen(recordKuwaku, specimenNo, value(found?.id))
      : null;
    if (duplicateRecord) {
      showToast(`この区画には ${specimenNo} がすでにあります`);
      return;
    }
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
    const positionMethod = normalizePositionMethod(formData.get("positionMethod"));
    const planSizeMode = normalizePlanSizeMode(value(formData.get("planSizeMode")));
    const planMultiPoints = planSizeMode === "複数点" ? readMultiPointRowsFromForm() : [];
    const rawLargeShapeType = value(formData.get("largeShapeType"));
    const largeShapeType =
      planSizeMode === "大きなもの"
        ? normalizeLargeShapeType(rawLargeShapeType) || normalizeLargeShapeLabel(rawLargeShapeType)
        : "";
    const isLineShape = largeShapeType === "直線状";
    const usesAxisDirection = isLineShape || largeShapeType === "長方形" || largeShapeType === "楕円";
    const largeAxisDirection = planSizeMode === "大きなもの" ? normalizeLargeAxisDirection(value(formData.get("largeAxisDirection"))) : "";
    const largeAxisPlungeDeg =
      planSizeMode === "大きなもの" ? normalizeLargeAxisPlungeDeg(value(formData.get("largeAxisPlungeDeg"))) : "";
    const largeAxisPlungeDir8 =
      planSizeMode === "大きなもの" ? normalizeCompass8Direction(value(formData.get("largeAxisPlungeDir8"))) : "";
    const planeStrikeDirection =
      planSizeMode === "大きなもの"
        ? normalizePlaneStrikeDirection(value(formData.get("planeStrikeDirection")) || (usesAxisDirection ? largeAxisDirection : ""))
        : "";
    const planeDipDeg = planSizeMode === "大きなもの" ? normalizePlaneDipDeg(value(formData.get("planeDipDeg"))) : "";
    const planeDipDir8 = planSizeMode === "大きなもの" ? normalizeCompass8Direction(value(formData.get("planeDipDir8"))) : "";
    const lineLengthCm = value(formData.get("lineLengthCm"));
    const rectSide1Cm = value(formData.get("rectSide1Cm"));
    const rectSide2Cm = value(formData.get("rectSide2Cm"));
    const ellipseLongRadiusCm = value(formData.get("ellipseLongRadiusCm"));
    const ellipseShortRadiusCm = value(formData.get("ellipseShortRadiusCm"));
    const altitudeInputEnabled = normalizeToggleFlag(formData.get("altitudeInputEnabled"));
    const altitudeDirectM = altitudeInputEnabled === "1" ? value(formData.get("altitudeDirectM")) : "";
    const imageCornerFields = extractImageCornerFieldsFromFormData(formData);
    const isLargeImageShape = planSizeMode === "大きなもの" && isLargeShapeImageType(largeShapeType);
    const isCustomLargeImageShape = isLargeImageShape && isCustomLargeShapeType(largeShapeType);
    const keepImageCornerFields = planSizeMode === "大きなもの";
    const imageTransformFields = extractImageTransformFieldsFromFormData(formData);
    const customLargeImageName = isCustomLargeImageShape
      ? normalizeCustomLargeImageName(value(formData.get("customLargeImageName")))
      : "";
    const customLargeImageDataUrl = isCustomLargeImageShape ? normalizeCustomLargeImageDataUrl(value(formData.get("customLargeImageDataUrl"))) : "";

    const recordBase = {
      id: recordId,
      kuwaku: recordKuwaku,
      specimenPrefix,
      specimenSerial,
      specimenNo,
      category: categoryFromPrefix(specimenPrefix),
      analysisType,
      levelHeight: positionMethod === "totalStation" ? "" : recordSiteSnapshot.levelHeight,
      date: recordSiteSnapshot.date,
      team: recordTeamState.team,
      teamOther: recordTeamState.teamOther,
      teamLead: recordSiteSnapshot.teamLead,
      recorder: recordSiteSnapshot.recorder,
      nameMemo: value(formData.get("nameMemo")),
      unit: compactNoSpaceValue(formData.get("unit")),
      discoverer: value(formData.get("discoverer")),
      identifier: value(formData.get("identifier")),
      levelUpperCm: positionMethod === "totalStation" ? "" : value(formData.get("levelUpperCm")),
      levelLowerCm: positionMethod === "totalStation" ? "" : value(formData.get("levelLowerCm")),
      altitudeInputEnabled,
      altitudeDirectM,
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
      positionMethod,
      tsCoordinateConvention: positionMethod === "totalStation" ? "southWestPositive" : "",
      tsStationPeg: normalizeTotalStationPointName(formData.get("tsStationPeg")),
      tsStationXNorthM: value(formData.get("tsStationXNorthM")),
      tsStationYEastM: value(formData.get("tsStationYEastM")),
      tsStationAltitudeM: value(formData.get("tsStationAltitudeM")),
      tsBacksightPeg: normalizeTotalStationPointName(formData.get("tsBacksightPeg")),
      tsBacksightXNorthM: value(formData.get("tsBacksightXNorthM")),
      tsBacksightYEastM: value(formData.get("tsBacksightYEastM")),
      tsBacksightAltitudeM: value(formData.get("tsBacksightAltitudeM")),
      tsInstrumentHeightM: value(formData.get("tsInstrumentHeightM")),
      tsTargetHeightM: value(formData.get("tsTargetHeightM")),
      tsObservationMode: value(formData.get("tsObservationMode")) === "polar" ? "polar" : "coordinate",
      tsPointCoordinateMode: "stationOffsetSouthWest",
      tsPointXNorthM: value(formData.get("tsPointXNorthM")),
      tsPointYEastM: value(formData.get("tsPointYEastM")),
      tsPointAltitudeM: value(formData.get("tsPointAltitudeM")),
      tsSlopeDistanceM: value(formData.get("tsSlopeDistanceM")),
      tsInclinationDeg: value(formData.get("tsInclinationDeg")),
      tsInclinationMin: value(formData.get("tsInclinationMin")),
      tsInclinationSec: value(formData.get("tsInclinationSec")),
      tsDirectionDeg: value(formData.get("tsDirectionDeg")),
      tsDirectionMin: value(formData.get("tsDirectionMin")),
      tsDirectionSec: value(formData.get("tsDirectionSec")),
      multiPoints: planMultiPoints,
      planSizeMode,
      largeShapeType,
      largeAxisDirection: isLargeImageShape || !usesAxisDirection ? "" : largeAxisDirection,
      largeAxisPlungeDeg: isLargeImageShape || !isLineShape ? "" : largeAxisPlungeDeg,
      largeAxisPlungeDir8: isLargeImageShape || !isLineShape ? "" : largeAxisPlungeDir8,
      planeStrikeDirection: isLineShape ? "" : planeStrikeDirection,
      planeDipDeg: isLineShape ? "" : planeDipDeg,
      planeDipDir8: isLineShape ? "" : planeDipDir8,
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
      imgRotateDeg: isLargeImageShape ? imageTransformFields.imgRotateDeg : "",
      imgFrameWidthCm: isLargeImageShape ? imageTransformFields.imgFrameWidthCm : "",
      imgFrameHeightCm: isLargeImageShape ? imageTransformFields.imgFrameHeightCm : "",
      imgSkewXDeg: isLargeImageShape ? imageTransformFields.imgSkewXDeg : "",
      imgSkewYDeg: isLargeImageShape ? imageTransformFields.imgSkewYDeg : "",
      imgFlipH: isLargeImageShape ? imageTransformFields.imgFlipH : "0",
      imgFlipV: isLargeImageShape ? imageTransformFields.imgFlipV : "0",
      imgLockAspectRatio: isLargeImageShape ? imageTransformFields.imgLockAspectRatio : "0",
      imgUseOriginalColor: isLargeImageShape ? imageTransformFields.imgUseOriginalColor : "0",
      customLargeImageName,
      customLargeImageDataUrl,
      customLargeImageAspect: isCustomLargeImageShape ? imageTransformFields.customLargeImageAspect : "",
      importantFlag: normalizeHasFlag(value(formData.get("importantFlag"))),
      simpleRecordFlag: normalizeCircleDashFlag(value(formData.get("simpleRecordFlag"))),
      layerName: getSelectedLayerName(),
      detail: compactNoSpaceValue(formData.get("detail")),
      detailSub: value(formData.get("detailSub")),
      layerColor: value(formData.get("layerColor")),
      layerLithology: value(formData.get("layerLithology")),
      layerFacies: composeLayerFacies(formData.get("layerColor"), formData.get("layerLithology")),
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

    const targetIndex = isEditTab
      ? state.records.findIndex((item) => item === found)
      : state.records.findIndex((item) => item.id === recordBase.id);
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
      const savedRecordIndex = state.records.findIndex((item) => item === record);
      activeEditRecordContext = {
        recordId: value(record.id),
        recordIndex: String(savedRecordIndex >= 0 ? savedRecordIndex : ""),
        recordSnapshot: buildCellEditRecordSnapshot(record),
      };
      renderEditHistory(record);
      updateEditMissingRequiredHighlights();
      return;
    }

    const carryForward = {
      layerName: record.layerName,
      unit: record.unit,
      detail: record.detail,
      detailSub: record.detailSub,
      layerColor: getLayerColor(record),
      layerLithology: getLayerLithology(record),
      layerFacies: record.layerFacies,
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
  if (recordNewBtn) {
    recordNewBtn.addEventListener("click", () => {
      const shouldClear = window.confirm(
        "現在入力している詳細情報は消去されます。\n新規入力を始める場合は「OK」を押してください。"
      );
      if (!shouldClear) {
        return;
      }
      resetRecordForm({ showMessage: true });
    });
  }
  if (recordPrevBtn) {
    recordPrevBtn.addEventListener("click", () => {
      moveToPreviousSpecimenWithoutSave();
    });
  }
  if (recordNextBtn) {
    recordNextBtn.addEventListener("click", () => {
      moveToNextSpecimenWithoutSave();
    });
  }
  if (recordCopyToInputBtn) {
    recordCopyToInputBtn.addEventListener("click", () => {
      copyCurrentEditToInput();
    });
  }
  if (recordCopyToInputTopBtn) {
    recordCopyToInputTopBtn.addEventListener("click", () => {
      copyCurrentEditToInput();
    });
  }

  if (recordTableBody) {
    recordTableBody.addEventListener("click", handleRecordTableActionClick);
  }
  if (editRecordTableBody) {
    editRecordTableBody.addEventListener("click", handleRecordTableActionClick);
  }

  outputListBody.addEventListener("click", (event) => {
    const button = event.target.closest("button[data-action]");
    if (!button) {
      return;
    }
    const action = button.dataset.action;
    const recordId = button.dataset.id;
    const row = button.closest("tr[data-record-index]");
    const recordIndex = value(button.dataset.recordIndex) || value(row?.dataset?.recordIndex);
    const record = findRecordByEditContext(recordId, recordIndex, null);

    if (action === "edit") {
      const rowKuwaku = value(button.dataset.kuwaku);
      openRecordForEdit(recordId, rowKuwaku, recordIndex);
      return;
    }
    if (action === "copy-to-input") {
      if (!record) {
        showToast("対象データが見つかりません");
        return;
      }
      const rowKuwaku = value(button.dataset.kuwaku);
      copySavedRecordToInput(value(record.id) || recordId, rowKuwaku, record);
      return;
    }
    if (action === "insert-row") {
      if (!record) {
        showToast("対象データが見つかりません");
        return;
      }
      const rowKuwaku = value(button.dataset.kuwaku);
      insertRowFromList(value(record.id) || recordId, rowKuwaku, record);
      return;
    }
    if (action === "position-preview") {
      if (!record) {
        showToast("対象データが見つかりません");
        return;
      }
      openPositionPreviewModal(record);
      return;
    }
    if (action === "delete") {
      if (!record) {
        showToast("対象データが見つかりません");
        return;
      }
      const answer = window.confirm(`標本番号 ${record.specimenNo} を削除しますか？`);
      if (!answer) {
        return;
      }
      const deletingId = value(record.id) || recordId;
      state.records = state.records.filter((item) => item !== record);
      if (editingRecordId === deletingId) {
        resetRecordForm({ showMessage: false });
      }
      if (selectedCardRecordId === deletingId) {
        selectedCardRecordId = "";
      }
      persist("記録を削除しました");
      renderRecordTable();
      renderOutputs();
      return;
    }
  });
  outputListBody.addEventListener("dblclick", (event) => {
    const target = event.target;
    if (!(target instanceof Element)) {
      return;
    }
    if (target.closest("button")) {
      return;
    }
    const cell = target.closest("td[data-cell-edit-key]");
    const row = target.closest("tr[data-record-id]");
    const editKey = value(cell?.dataset?.cellEditKey);
    const recordId = value(row?.dataset?.recordId);
    const recordIndex = value(row?.dataset?.recordIndex);
    if (!cell || !row || !editKey || !recordId || !outputListBody.contains(row)) {
      return;
    }
    event.preventDefault();
    openOutputCellEditModal(recordId, editKey, recordIndex);
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
      rememberOutputFilters();
      renderOutputs();
    });
  }
  if (outputCategorySelect) {
    outputCategorySelect.addEventListener("change", () => {
      selectedOutputCategory = value(outputCategorySelect.value) || EXPORT_CATEGORY_ALL_VALUE;
      selectedCardRecordId = "";
      rememberOutputFilters();
      renderOutputs();
    });
  }
  if (outputStatusSelect) {
    outputStatusSelect.addEventListener("change", () => {
      selectedOutputStatus = value(outputStatusSelect.value) || "all";
      selectedCardRecordId = "";
      rememberOutputFilters();
      renderOutputs();
    });
  }
  if (outputDateSelect) {
    outputDateSelect.addEventListener("change", () => {
      selectedOutputDate = value(outputDateSelect.value);
      selectedCardRecordId = "";
      rememberOutputFilters();
      renderOutputs();
    });
  }
  if (outputSearchInput) {
    outputSearchInput.addEventListener("input", () => {
      outputSearchText = value(outputSearchInput.value);
      selectedCardRecordId = "";
      rememberOutputFilters();
      renderOutputs();
    });
  }
  if (outputColumnToggleRow) {
    outputColumnToggleRow.addEventListener("click", (event) => {
      const target = event.target;
      if (!(target instanceof Element)) {
        return;
      }
      const button = target.closest("button[data-col-key]");
      if (!button || !outputColumnToggleRow.contains(button)) {
        return;
      }
      toggleOutputListColumnVisibility(value(button.dataset.colKey));
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
  if (planCategorySelect) {
    planCategorySelect.addEventListener("change", () => {
      selectedPlanCategory = value(planCategorySelect.value) || EXPORT_CATEGORY_ALL_VALUE;
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
  if (viewerCategorySelect) {
    viewerCategorySelect.addEventListener("change", () => {
      selectedViewerCategory = value(viewerCategorySelect.value) || EXPORT_CATEGORY_ALL_VALUE;
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
  if (viewerViewSeBtn) {
    viewerViewSeBtn.addEventListener("click", () => {
      selectedViewerPerspective = "se";
      applyViewerPerspective();
      syncViewerViewButtons();
    });
  }
  if (viewerViewEastBtn) {
    viewerViewEastBtn.addEventListener("click", () => {
      selectedViewerPerspective = "east";
      applyViewerPerspective();
      syncViewerViewButtons();
    });
  }
  if (viewerViewWestBtn) {
    viewerViewWestBtn.addEventListener("click", () => {
      selectedViewerPerspective = "west";
      applyViewerPerspective();
      syncViewerViewButtons();
    });
  }
  if (viewerViewSouthBtn) {
    viewerViewSouthBtn.addEventListener("click", () => {
      selectedViewerPerspective = "south";
      applyViewerPerspective();
      syncViewerViewButtons();
    });
  }
  if (viewerViewNorthBtn) {
    viewerViewNorthBtn.addEventListener("click", () => {
      selectedViewerPerspective = "north";
      applyViewerPerspective();
      syncViewerViewButtons();
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
      const dataUrl = await loadImageFileDataUrlWithFallback(file, {
        maxLength: 1280,
        quality: 0.72,
        mimeType: "image/jpeg",
      });
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
      const dataUrl = await loadImageFileDataUrlWithFallback(file, {
        maxLength: 1280,
        quality: 0.72,
        mimeType: "image/jpeg",
      });
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

function stopBrowserCamera() {
  if (activeCameraStream) {
    activeCameraStream.getTracks().forEach((track) => track.stop());
    activeCameraStream = null;
  }
  if (cameraCaptureVideo) {
    cameraCaptureVideo.srcObject = null;
  }
  if (cameraCaptureShutterBtn) {
    cameraCaptureShutterBtn.disabled = true;
  }
}

function closeBrowserCamera() {
  stopBrowserCamera();
  cameraCaptureModal?.classList.add("hidden");
}

async function openBrowserCamera(destination) {
  if (!cameraCaptureModal || !cameraCaptureVideo || !cameraCaptureStatus || !cameraCaptureShutterBtn) {
    return;
  }
  activeCameraDestination = destination === "sectionDiagram" ? "sectionDiagram" : "photo";
  cameraCaptureModal.classList.remove("hidden");
  cameraCaptureStatus.textContent = "カメラを準備しています。許可の確認が表示されたら「許可」を選んでください。";
  cameraCaptureShutterBtn.disabled = true;
  stopBrowserCamera();

  if (!navigator.mediaDevices?.getUserMedia) {
    cameraCaptureStatus.textContent = "このブラウザではカメラを直接使用できません。「保存フォルダから選択」を使用してください。";
    return;
  }

  try {
    activeCameraStream = await navigator.mediaDevices.getUserMedia({
      video: { facingMode: { ideal: "environment" } },
      audio: false,
    });
    cameraCaptureVideo.srcObject = activeCameraStream;
    await cameraCaptureVideo.play();
    cameraCaptureShutterBtn.disabled = false;
    cameraCaptureStatus.textContent = "映像を確認して「撮影して追加」を押してください。続けて複数枚撮影できます。";
  } catch (_error) {
    stopBrowserCamera();
    cameraCaptureStatus.textContent = "カメラを使用できません。ブラウザのサイト設定で、このページのカメラを「許可」にしてください。";
  }
}

function cameraFrameToFile() {
  return new Promise((resolve, reject) => {
    if (!cameraCaptureVideo?.videoWidth || !cameraCaptureVideo?.videoHeight) {
      reject(new Error("Camera image unavailable"));
      return;
    }
    const canvas = document.createElement("canvas");
    canvas.width = cameraCaptureVideo.videoWidth;
    canvas.height = cameraCaptureVideo.videoHeight;
    const context = canvas.getContext("2d");
    if (!context) {
      reject(new Error("Canvas context unavailable"));
      return;
    }
    context.drawImage(cameraCaptureVideo, 0, 0, canvas.width, canvas.height);
    canvas.toBlob((blob) => {
      if (!blob) {
        reject(new Error("Camera capture failed"));
        return;
      }
      const name = `camera-${new Date().toISOString().replace(/[:.]/g, "-")}.jpg`;
      resolve(new File([blob], name, { type: "image/jpeg", lastModified: Date.now() }));
    }, "image/jpeg", 0.9);
  });
}

sectionDiagramCameraBtn?.addEventListener("click", () => openBrowserCamera("sectionDiagram"));
photoCameraBtn?.addEventListener("click", () => openBrowserCamera("photo"));
cameraCaptureCloseBtn?.addEventListener("click", closeBrowserCamera);
cameraCaptureCancelBtn?.addEventListener("click", closeBrowserCamera);
cameraCaptureModal?.addEventListener("click", (event) => {
  if (event.target === cameraCaptureModal) {
    closeBrowserCamera();
  }
});
cameraCaptureShutterBtn?.addEventListener("click", async () => {
  cameraCaptureShutterBtn.disabled = true;
  cameraCaptureStatus.textContent = "撮影した画像を追加しています。";
  try {
    const file = await cameraFrameToFile();
    if (activeCameraDestination === "sectionDiagram") {
      await addSectionDiagramsFromFiles([file]);
    } else {
      await addPhotosFromFiles([file]);
    }
    cameraCaptureStatus.textContent = "撮影画像を追加しました。続けて撮影できます。";
  } catch (_error) {
    cameraCaptureStatus.textContent = "撮影に失敗しました。カメラ映像が表示されていることを確認してください。";
  } finally {
    cameraCaptureShutterBtn.disabled = !activeCameraStream;
  }
});
document.addEventListener("keydown", (event) => {
  if (event.key === "Escape" && !cameraCaptureModal?.classList.contains("hidden")) {
    closeBrowserCamera();
  }
});
window.addEventListener("pagehide", stopBrowserCamera);

function setActiveTab(tabId) {
  const previousTabId = getActiveTabId();
  hideHoverEditMenu();
  closeOutputCellEditModal();
  if (previousTabId === "output-tab") {
    rememberOutputFilters();
  }
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
    renderViewerOutput({ preserveCamera: true });
    ensureViewerCanvasSize();
  }
  if (tabId === "plan-tab") {
    renderPlanOutput();
  }
  if (tabId === "output-tab") {
    restoreOutputFilters();
    renderListOutput();
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
    if (recordCopyToInputTopBtn) {
      recordCopyToInputTopBtn.classList.remove("hidden");
    }
    return;
  }

  if (recordForm.parentElement !== recordFormHost) {
    recordFormHost.appendChild(recordForm);
  }
  if (recordCopyToInputBtn) {
    recordCopyToInputBtn.classList.remove("hidden");
  }
  if (recordCopyToInputTopBtn) {
    recordCopyToInputTopBtn.classList.add("hidden");
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

function rememberOutputFilters() {
  outputFilterMemory = {
    kuwaku: value(selectedOutputKuwaku) || ALL_GRIDS_VALUE,
    category: value(selectedOutputCategory) || EXPORT_CATEGORY_ALL_VALUE,
    status: ["all", "complete", "incomplete"].includes(value(selectedOutputStatus)) ? value(selectedOutputStatus) : "all",
    date: value(selectedOutputDate),
    searchText: value(outputSearchText),
  };
}

function restoreOutputFilters() {
  selectedOutputKuwaku = value(outputFilterMemory.kuwaku) || ALL_GRIDS_VALUE;
  selectedOutputCategory = value(outputFilterMemory.category) || EXPORT_CATEGORY_ALL_VALUE;
  selectedOutputStatus = ["all", "complete", "incomplete"].includes(value(outputFilterMemory.status))
    ? value(outputFilterMemory.status)
    : "all";
  selectedOutputDate = value(outputFilterMemory.date);
  outputSearchText = value(outputFilterMemory.searchText);
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
  if (siteForm.elements.scribe) {
    siteForm.elements.scribe.value = state.site.scribe || "";
  }
}

async function loadTeamRosterFromDefaultFile() {
  setTeamRosterStatus("名簿を読込中…");
  try {
    const rows = Array.isArray(window.__TEAM_ROSTER_ROWS__) ? window.__TEAM_ROSTER_ROWS__ : [];
    teamRosterAssignmentMap = buildTeamRosterAssignmentMap(rows);
    if (!teamRosterAssignmentMap.size) throw new Error("roster empty");
    teamRosterLoaded = true;
    setTeamRosterStatus("名簿を自動読込しました");
    applyTeamRosterAutofill();
  } catch (_error) {
    setTeamRosterStatus(`名簿の自動読込に失敗: ${TEAM_ROSTER_FILE_NAME}`);
  }
}

function buildTeamRosterAssignmentMap(rows) {
  const assignmentMap = new Map();
  if (!Array.isArray(rows) || rows.length === 0) return assignmentMap;
  const headerRow = Array.isArray(rows[0]) ? rows[0] : [];
  const dayBlocks = buildRosterDayBlocks(headerRow);
  for (let rowIndex = 1; rowIndex < rows.length; rowIndex += 1) {
    const row = Array.isArray(rows[rowIndex]) ? rows[rowIndex] : [];
    const name = value(row[TEAM_ROSTER_NAME_COL - 1]);
    const team = normalizeRosterTeam(row[TEAM_ROSTER_TEAM_COL - 1]);
    const roleText = value(row[TEAM_ROSTER_ROLE_COL - 1]);
    if (!name || !team || !roleText) continue;
    for (const block of dayBlocks) {
      if (!rowHasAttendanceMark(row, block.startCol, block.endCol)) continue;
      const key = `${block.day}-${team}`;
      if (!assignmentMap.has(key)) assignmentMap.set(key, { teamLeads: new Set(), recorders: new Set() });
      const target = assignmentMap.get(key);
      if (/班長/.test(roleText)) target.teamLeads.add(name);
      if (/記載/.test(roleText)) target.recorders.add(name);
    }
  }
  return assignmentMap;
}

function buildRosterDayBlocks(headerRow) {
  const blocks = [];
  let current = null;
  for (let col = TEAM_ROSTER_DAY_COL_START; col <= TEAM_ROSTER_DAY_COL_END; col += 1) {
    const match = value(headerRow[col - 1]).match(/(\d{1,2})日/);
    if (match) {
      if (current) {
        current.endCol = col - 1;
        blocks.push(current);
      }
      current = { day: Number(match[1]), startCol: col, endCol: col };
    }
  }
  if (current) {
    current.endCol = TEAM_ROSTER_DAY_COL_END;
    blocks.push(current);
  }
  return blocks;
}

function rowHasAttendanceMark(row, startCol, endCol) {
  for (let col = startCol; col <= endCol; col += 1) {
    const mark = value(row[col - 1]).replace(/\s+/g, "");
    if (mark === "○" || mark === "◯") return true;
  }
  return false;
}

function normalizeRosterTeam(teamRaw) {
  const text = value(teamRaw);
  if (!text) return "";
  return text.replace(/[^\d]/g, "") || text;
}

function setTeamRosterStatus(text) {
  if (rosterStatusEl) rosterStatusEl.textContent = text;
}

function applyTeamRosterAutofill() {
  if (!siteForm?.elements) return;
  const team = value(siteForm.elements.team?.value);
  const date = value(siteForm.elements.date?.value);
  if (!teamRosterLoaded || !team || !date || team === OTHER_TEAM_NAME) return;
  const day = extractDayFromIsoDate(date);
  if (!day) return;
  const assignment = teamRosterAssignmentMap.get(`${day}-${normalizeRosterTeam(team)}`);
  if (!assignment) return;
  if (siteForm.elements.teamLead) siteForm.elements.teamLead.value = Array.from(assignment.teamLeads).join(", ");
  if (siteForm.elements.recorder) siteForm.elements.recorder.value = Array.from(assignment.recorders).join(", ");
}

function extractDayFromIsoDate(dateRaw) {
  const match = value(dateRaw).match(/^\d{4}-\d{2}-(\d{2})$/);
  if (!match) return null;
  const day = Number(match[1]);
  return Number.isFinite(day) ? day : null;
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
}

function moveToPreviousSpecimenWithoutSave() {
  const activeTabId = getActiveTabId();
  if (activeTabId !== "input-tab" && activeTabId !== "edit-tab") {
    return;
  }
  const currentKuwaku = currentKuwakuForDuplicateWarning(activeTabId);
  if (!currentKuwaku) {
    showToast("区画（グリッド）を入力してください");
    return;
  }
  const currentPrefix = normalizeSpecimenPrefix(specimenPrefixInput?.value);
  const currentSerial = compactNoSpaceValue(specimenSerialInput?.value);
  if (!state.records.length) {
    let prevSerial = "1";
    if (currentSerial && /^\d+$/.test(currentSerial)) {
      prevSerial = String(Math.max(1, Number(currentSerial) - 1));
    } else if (currentSerial) {
      prevSerial = currentSerial;
    }
    applyNextNavigationTarget({
      kuwaku: currentKuwaku,
      prefix: currentPrefix,
      serial: prevSerial,
    });
    showToast(`前へ: ${buildSpecimenNo(currentPrefix, prevSerial)}`);
    return;
  }

  const prevRecord = findPreviousRecordByGridPrefixThenSerial(currentKuwaku, currentPrefix, currentSerial);
  if (!prevRecord) {
    showToast("前のデータが見つかりません");
    return;
  }
  const prevKuwaku = getRecordKuwaku(prevRecord);
  const prevSpecimen = parseSpecimenNo(prevRecord.specimenNo, prevRecord.specimenPrefix, prevRecord.specimenSerial);
  if (activeTabId === "edit-tab") {
    openRecordForEdit(prevRecord.id, prevKuwaku);
    showToast(`前へ: ${prevKuwaku} / ${prevSpecimen.specimenNo}`);
    return;
  }
  loadRecordIntoInputForNavigation(prevRecord, prevKuwaku);
  showToast(`前へ: ${prevKuwaku} / ${prevSpecimen.specimenNo}`);
}

function moveToNextSpecimenWithoutSave() {
  const activeTabId = getActiveTabId();
  if (activeTabId !== "input-tab" && activeTabId !== "edit-tab") {
    return;
  }
  const currentKuwaku = currentKuwakuForDuplicateWarning(activeTabId);
  if (!currentKuwaku) {
    showToast("区画（グリッド）を入力してください");
    return;
  }
  const currentPrefix = normalizeSpecimenPrefix(specimenPrefixInput?.value);
  const currentSerial = compactNoSpaceValue(specimenSerialInput?.value);
  if (!state.records.length) {
    const nextSerial = currentSerial ? (/^\d+$/.test(currentSerial) ? String(Number(currentSerial) + 1) : `${currentSerial}1`) : "1";
    applyNextNavigationTarget({
      kuwaku: currentKuwaku,
      prefix: currentPrefix,
      serial: nextSerial,
    });
    showToast(`次へ: ${buildSpecimenNo(currentPrefix, nextSerial)}`);
    return;
  }

  const nextRecord = findNextRecordByGridPrefixThenSerial(currentKuwaku, currentPrefix, currentSerial);
  if (!nextRecord) {
    showToast("次のデータが見つかりません");
    return;
  }

  const nextKuwaku = getRecordKuwaku(nextRecord);
  const nextSpecimen = parseSpecimenNo(nextRecord.specimenNo, nextRecord.specimenPrefix, nextRecord.specimenSerial);
  if (activeTabId === "edit-tab") {
    openRecordForEdit(nextRecord.id, nextKuwaku);
    showToast(`次へ: ${nextKuwaku} / ${nextSpecimen.specimenNo}`);
    return;
  }
  loadRecordIntoInputForNavigation(nextRecord, nextKuwaku);
  showToast(`次へ: ${nextKuwaku} / ${nextSpecimen.specimenNo}`);
}

function findPreviousRecordByGridPrefixThenSerial(currentKuwakuRaw, currentPrefixRaw, currentSerialRaw) {
  const currentKuwakuValue = kuwakuValueForSelect(currentKuwakuRaw);
  const currentPrefix = normalizeSpecimenPrefix(currentPrefixRaw);
  const currentSerial = compactNoSpaceValue(currentSerialRaw);
  const sorted = [...state.records].sort(compareRecordsByKuwakuThenSpecimen);
  if (!sorted.length) {
    return null;
  }

  const groupedByGrid = new Map();
  const gridOrder = [];
  sorted.forEach((record) => {
    const gridValue = kuwakuValueForSelect(getRecordKuwaku(record));
    if (!groupedByGrid.has(gridValue)) {
      groupedByGrid.set(gridValue, []);
      gridOrder.push(gridValue);
    }
    groupedByGrid.get(gridValue).push(record);
  });

  const currentGridRecords = groupedByGrid.get(currentKuwakuValue) || [];
  const gridStartIndex = resolvePreviousGridStartIndex(gridOrder, currentKuwakuValue);
  if (!currentGridRecords.length) {
    for (let step = 0; step < gridOrder.length; step += 1) {
      const gridValue = gridOrder[(gridStartIndex - step + gridOrder.length) % gridOrder.length];
      const records = groupedByGrid.get(gridValue) || [];
      if (records.length) {
        return records[records.length - 1];
      }
    }
    return sorted[sorted.length - 1];
  }

  const samePrefixRecords = currentGridRecords
    .filter((record) => normalizeSpecimenPrefix(parseSpecimenNo(record.specimenNo, record.specimenPrefix, record.specimenSerial).prefix) === currentPrefix)
    .sort(compareRecordsBySpecimenNo);
  if (samePrefixRecords.length) {
    if (currentSerial) {
      const exactIndex = samePrefixRecords.findIndex((record) => {
        const specimen = parseSpecimenNo(record.specimenNo, record.specimenPrefix, record.specimenSerial);
        return compactNoSpaceValue(specimen.serial) === currentSerial;
      });
      if (exactIndex > 0) {
        return samePrefixRecords[exactIndex - 1];
      }
      if (exactIndex < 0) {
        for (let i = samePrefixRecords.length - 1; i >= 0; i -= 1) {
          const specimen = parseSpecimenNo(
            samePrefixRecords[i].specimenNo,
            samePrefixRecords[i].specimenPrefix,
            samePrefixRecords[i].specimenSerial
          );
          if (compareSpecimenSerialOnly(compactNoSpaceValue(specimen.serial), currentSerial) < 0) {
            return samePrefixRecords[i];
          }
        }
      }
    } else {
      return samePrefixRecords[samePrefixRecords.length - 1];
    }
  }

  const sortedPrefixes = Array.from(
    new Set(
      currentGridRecords.map((record) =>
        normalizeSpecimenPrefix(parseSpecimenNo(record.specimenNo, record.specimenPrefix, record.specimenSerial).prefix)
      )
    )
  ).sort((a, b) => a.localeCompare(b, "ja", { sensitivity: "base" }));
  const prevPrefixes = sortedPrefixes
    .filter((prefix) => prefix.localeCompare(currentPrefix, "ja", { sensitivity: "base" }) < 0)
    .reverse();
  for (const prefix of prevPrefixes) {
    const prefixRecords = currentGridRecords
      .filter((record) => {
        const specimen = parseSpecimenNo(record.specimenNo, record.specimenPrefix, record.specimenSerial);
        return normalizeSpecimenPrefix(specimen.prefix) === prefix;
      })
      .sort(compareRecordsBySpecimenNo);
    if (prefixRecords.length) {
      return prefixRecords[prefixRecords.length - 1];
    }
  }

  for (let step = 1; step <= gridOrder.length; step += 1) {
    const gridValue = gridOrder[(gridStartIndex - step + gridOrder.length) % gridOrder.length];
    const records = groupedByGrid.get(gridValue) || [];
    if (records.length) {
      return records[records.length - 1];
    }
  }
  return sorted[sorted.length - 1];
}

function findNextRecordByGridPrefixThenSerial(currentKuwakuRaw, currentPrefixRaw, currentSerialRaw) {
  const currentKuwakuValue = kuwakuValueForSelect(currentKuwakuRaw);
  const currentPrefix = normalizeSpecimenPrefix(currentPrefixRaw);
  const currentSerial = compactNoSpaceValue(currentSerialRaw);
  const sorted = [...state.records].sort(compareRecordsByKuwakuThenSpecimen);
  if (!sorted.length) {
    return null;
  }

  const groupedByGrid = new Map();
  const gridOrder = [];
  sorted.forEach((record) => {
    const gridValue = kuwakuValueForSelect(getRecordKuwaku(record));
    if (!groupedByGrid.has(gridValue)) {
      groupedByGrid.set(gridValue, []);
      gridOrder.push(gridValue);
    }
    groupedByGrid.get(gridValue).push(record);
  });

  const currentGridRecords = groupedByGrid.get(currentKuwakuValue) || [];
  const gridStartIndex = resolveNextGridStartIndex(gridOrder, currentKuwakuValue);
  if (!currentGridRecords.length) {
    for (let step = 0; step < gridOrder.length; step += 1) {
      const gridValue = gridOrder[(gridStartIndex + step) % gridOrder.length];
      const records = groupedByGrid.get(gridValue) || [];
      if (records.length) {
        return records[0];
      }
    }
    return sorted[0];
  }

  const samePrefixRecords = currentGridRecords
    .filter((record) => normalizeSpecimenPrefix(parseSpecimenNo(record.specimenNo, record.specimenPrefix, record.specimenSerial).prefix) === currentPrefix)
    .sort(compareRecordsBySpecimenNo);

  if (samePrefixRecords.length) {
    if (currentSerial) {
      const exactIndex = samePrefixRecords.findIndex((record) => {
        const specimen = parseSpecimenNo(record.specimenNo, record.specimenPrefix, record.specimenSerial);
        return compactNoSpaceValue(specimen.serial) === currentSerial;
      });
      if (exactIndex >= 0 && exactIndex + 1 < samePrefixRecords.length) {
        return samePrefixRecords[exactIndex + 1];
      }
      if (exactIndex < 0) {
        const nextBySerial = samePrefixRecords.find((record) => {
          const specimen = parseSpecimenNo(record.specimenNo, record.specimenPrefix, record.specimenSerial);
          return compareSpecimenSerialOnly(compactNoSpaceValue(specimen.serial), currentSerial) > 0;
        });
        if (nextBySerial) {
          return nextBySerial;
        }
      }
    } else {
      return samePrefixRecords[0];
    }
  }

  const sortedPrefixes = Array.from(
    new Set(
      currentGridRecords.map((record) =>
        normalizeSpecimenPrefix(parseSpecimenNo(record.specimenNo, record.specimenPrefix, record.specimenSerial).prefix)
      )
    )
  ).sort((a, b) => a.localeCompare(b, "ja", { sensitivity: "base" }));
  const nextPrefixes = sortedPrefixes.filter(
    (prefix) => prefix.localeCompare(currentPrefix, "ja", { sensitivity: "base" }) > 0
  );
  for (const prefix of nextPrefixes) {
    const prefixRecords = currentGridRecords
      .filter((record) => {
        const specimen = parseSpecimenNo(record.specimenNo, record.specimenPrefix, record.specimenSerial);
        return normalizeSpecimenPrefix(specimen.prefix) === prefix;
      })
      .sort(compareRecordsBySpecimenNo);
    if (prefixRecords.length) {
      return prefixRecords[0];
    }
  }

  for (let step = 1; step <= gridOrder.length; step += 1) {
    const gridValue = gridOrder[(gridStartIndex + step) % gridOrder.length];
    const records = groupedByGrid.get(gridValue) || [];
    if (records.length) {
      return records[0];
    }
  }
  return sorted[0];
}

function resolveNextGridStartIndex(gridOrder, currentGridValue) {
  const exactIndex = gridOrder.indexOf(currentGridValue);
  if (exactIndex >= 0) {
    return exactIndex;
  }
  const currentLabel = kuwakuLabelForSelect(currentGridValue);
  const insertedIndex = gridOrder.findIndex(
    (value) =>
      kuwakuLabelForSelect(value).localeCompare(currentLabel, "ja", {
        numeric: true,
        sensitivity: "base",
      }) > 0
  );
  return insertedIndex >= 0 ? insertedIndex : 0;
}

function resolvePreviousGridStartIndex(gridOrder, currentGridValue) {
  const exactIndex = gridOrder.indexOf(currentGridValue);
  if (exactIndex >= 0) {
    return exactIndex;
  }
  const currentLabel = kuwakuLabelForSelect(currentGridValue);
  const insertedIndex = gridOrder.findIndex(
    (value) =>
      kuwakuLabelForSelect(value).localeCompare(currentLabel, "ja", {
        numeric: true,
        sensitivity: "base",
      }) > 0
  );
  if (insertedIndex >= 0) {
    return (insertedIndex - 1 + gridOrder.length) % gridOrder.length;
  }
  return Math.max(0, gridOrder.length - 1);
}

function compareSpecimenSerialOnly(aSerialRaw, bSerialRaw) {
  const aSerial = compactNoSpaceValue(aSerialRaw);
  const bSerial = compactNoSpaceValue(bSerialRaw);
  const aIsNumber = /^\d+$/.test(aSerial);
  const bIsNumber = /^\d+$/.test(bSerial);
  if (aIsNumber && bIsNumber) {
    return Number(aSerial) - Number(bSerial);
  }
  return aSerial.localeCompare(bSerial, "ja", { numeric: true, sensitivity: "base" });
}

function applyNextNavigationTarget({ kuwaku = "", prefix = DEFAULT_SPECIMEN_PREFIX, serial = "" } = {}) {
  const activeTabId = getActiveTabId();
  const kuwakuParts = parseKuwaku(kuwaku);
  if (activeTabId === "edit-tab") {
    if (editKuwakuHeadAInput) {
      editKuwakuHeadAInput.value = kuwakuParts.headA || DEFAULT_KUWAKU_HEAD_A;
    }
    if (editKuwakuHeadBInput) {
      editKuwakuHeadBInput.value = kuwakuParts.headB || DEFAULT_KUWAKU_HEAD_B;
    }
    if (editKuwakuBlockInput) {
      editKuwakuBlockInput.value = kuwakuParts.block || "";
    }
    if (editKuwakuNoInput) {
      editKuwakuNoInput.value = kuwakuParts.no || "";
    }
  } else {
    if (siteForm?.elements?.kuwakuHeadA) {
      siteForm.elements.kuwakuHeadA.value = kuwakuParts.headA || DEFAULT_KUWAKU_HEAD_A;
    }
    if (siteForm?.elements?.kuwakuHeadB) {
      siteForm.elements.kuwakuHeadB.value = kuwakuParts.headB || DEFAULT_KUWAKU_HEAD_B;
    }
    if (siteForm?.elements?.kuwakuBlock) {
      siteForm.elements.kuwakuBlock.value = kuwakuParts.block || "";
    }
    if (siteForm?.elements?.kuwakuNo) {
      siteForm.elements.kuwakuNo.value = kuwakuParts.no || "";
    }
    syncInputSiteStateFromForm();
  }
  activateSpecimenPrefix(prefix);
  specimenSerialInput.value = compactNoSpaceValue(serial);
  updateSpecimenNoFromParts();
  if (getActiveTabId() === "edit-tab") {
    renderRecordTable();
    updateEditMissingRequiredHighlights();
  } else if (getActiveTabId() === "input-tab") {
    renderRecordTable();
  }
}

function loadRecordIntoInputForNavigation(record, preferredKuwaku = "") {
  if (!record) {
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
    syncInputSiteStateFromForm();
  }
  if (getActiveTabId() !== "input-tab") {
    setActiveTab("input-tab");
  }
  populateRecordForm({
    ...record,
    id: "",
  });
  isOverwriteMode = false;
  overwriteOriginalRecord = null;
  clearOverwriteUpdatedState();
  clearEditHistory();
  editingRecordId = null;
  activeEditRecordContext = null;
  if (recordIdInput) {
    recordIdInput.value = "";
  }
  renderRecordTable();
  updateDuplicateSpecimenWarning();
}

function syncInputSiteStateFromForm() {
  if (!siteForm?.elements) {
    return;
  }
  const kuwakuHeadA = normalizeKuwakuHeadA(siteForm.elements.kuwakuHeadA?.value);
  const kuwakuHeadB = normalizeKuwakuHeadB(siteForm.elements.kuwakuHeadB?.value);
  const kuwakuBlock = normalizeKuwakuBlock(siteForm.elements.kuwakuBlock?.value);
  const kuwakuNo = normalizeKuwakuNo(siteForm.elements.kuwakuNo?.value);
  const teamState = normalizeTeamState(value(siteForm.elements.team?.value), value(siteForm.elements.teamOther?.value));
  state.site = {
    ...state.site,
    kuwaku: buildKuwaku(kuwakuHeadA, kuwakuHeadB, kuwakuBlock, kuwakuNo),
    kuwakuHeadA,
    kuwakuHeadB,
    kuwakuBlock,
    kuwakuNo,
    levelHeight: value(siteForm.elements.levelHeight?.value),
    date: value(siteForm.elements.date?.value),
    team: teamState.team,
    teamOther: teamState.teamOther,
    teamLead: value(siteForm.elements.teamLead?.value),
    recorder: value(siteForm.elements.recorder?.value),
  };
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
  if (group === "tsSetupNsDir" || group === "tsSetupEwDir") {
    applyTotalStationPosition();
  }
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
  setDirectionGroupValue("tsSetupNsDir", recordForm?.elements?.tsSetupNsDir?.value);
  setDirectionGroupValue("tsSetupEwDir", recordForm?.elements?.tsSetupEwDir?.value);
  setDirectionGroupValue("layerRelative", layerRelativeInput?.value);
  setDirectionGroupValue("planSizeMode", planSizeModeInput?.value);
  setDirectionGroupValue("largeShapeType", largeShapeTypeInput?.value);
  setDirectionGroupValue("plungeDir8", largeAxisPlungeDirInput?.value);
  setDirectionGroupValue("planeDipDir8", planeDipDirInput?.value);
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
  if (group === "tsSetupNsDir" && recordForm?.elements?.tsSetupNsDir) {
    recordForm.elements.tsSetupNsDir.value = normalized;
    return;
  }
  if (group === "tsSetupEwDir" && recordForm?.elements?.tsSetupEwDir) {
    recordForm.elements.tsSetupEwDir.value = normalized;
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
    return;
  }
  if (group === "plungeDir8" && largeAxisPlungeDirInput) {
    largeAxisPlungeDirInput.value = normalized;
    return;
  }
  if (group === "planeDipDir8" && planeDipDirInput) {
    planeDipDirInput.value = normalized;
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
  if (group === "tsSetupNsDir") {
    return normalizeDirectionValue(group, recordForm?.elements?.tsSetupNsDir?.value);
  }
  if (group === "tsSetupEwDir") {
    return normalizeDirectionValue(group, recordForm?.elements?.tsSetupEwDir?.value);
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
  if (group === "plungeDir8") {
    return normalizeDirectionValue(group, largeAxisPlungeDirInput?.value);
  }
  if (group === "planeDipDir8") {
    return normalizeDirectionValue(group, planeDipDirInput?.value);
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

function isCustomLargeShapeType(shapeTypeRaw) {
  return normalizeLargeShapeType(shapeTypeRaw) === CUSTOM_LARGE_SHAPE_TYPE;
}

function getImageShapeDilateIterations(shapeTypeRaw) {
  return isCustomLargeShapeType(shapeTypeRaw)
    ? CUSTOM_IMAGE_SHAPE_CANVAS_DILATE_ITERATIONS
    : IMAGE_SHAPE_CANVAS_DILATE_ITERATIONS;
}

function normalizeCustomLargeImageName(nameRaw) {
  return value(nameRaw);
}

function normalizeCustomLargeImageDataUrl(dataUrlRaw) {
  const dataUrl = value(dataUrlRaw);
  return dataUrl.startsWith("data:image/") ? dataUrl : "";
}

function normalizeCustomLargeImageAspect(aspectRaw) {
  const text = value(aspectRaw).replace(",", ".");
  if (!text) {
    return "";
  }
  const matched = text.match(/\d+(?:\.\d+)?/);
  if (!matched) {
    return "";
  }
  const ratio = Number(matched[0]);
  if (!Number.isFinite(ratio) || ratio <= 0) {
    return "";
  }
  return String(Number(ratio.toFixed(8)));
}

function formatLengthInputValue(lengthRaw) {
  const length = Number(lengthRaw);
  if (!Number.isFinite(length)) {
    return "";
  }
  if (Math.abs(length) < 0.0005) {
    return "0";
  }
  return String(Number(length.toFixed(3)));
}

function normalizePlanMultiPointDir(valueRaw, axis = "ns") {
  return axis === "ew" ? normalizeEwDir(valueRaw) : normalizeNsDir(valueRaw);
}

function normalizePlanMultiPointDistance(valueRaw) {
  const distance = parseDistanceToCm(valueRaw);
  if (distance == null || distance < 0) {
    return "";
  }
  return formatLengthInputValue(distance);
}

function isTotalStationMeasurementSelected() {
  return Boolean(recordForm) && normalizePositionMethod(new FormData(recordForm).get("positionMethod")) === "totalStation";
}

function isTotalStationPolarMeasurementSelected() {
  if (!isTotalStationMeasurementSelected()) return false;
  return value(new FormData(recordForm).get("tsObservationMode")) === "polar";
}

function gridReferenceNameForKuwaku(kuwakuRaw) {
  const parts = parseKuwaku(kuwakuLabelForSelect(kuwakuRaw));
  return normalizeTotalStationPointName(`${value(parts.headB)}-${value(parts.block)}-${value(parts.no)}`);
}

function getCurrentMultiPointGridReference() {
  return findTotalStationGridReferencePoint(currentInputGridReferenceName());
}

function getTotalStationMultiPointOffsetsForForm(point) {
  if (point?.coordinateMode === "stationOffsetSouthWest") {
    return point;
  }
  const gridReference = getCurrentMultiPointGridReference();
  const stationX = parseTotalStationNumber(recordForm?.elements?.tsStationXNorthM?.value);
  const stationY = parseTotalStationNumber(recordForm?.elements?.tsStationYEastM?.value);
  const xEastM = Number(point?.xEastM);
  const yNorthM = Number(point?.yNorthM);
  if (gridReference && [stationX, stationY, xEastM, yNorthM].every((number) => Number.isFinite(number))) {
    const pointX = gridReference.x + PLAN_SIZE_CM / 100 - yNorthM;
    const pointY = gridReference.y - xEastM;
    return {
      southFromStationM: String(pointX - stationX),
      westFromStationM: String(pointY - stationY),
      zAltitudeM: value(point?.zAltitudeM),
      coordinateMode: "stationOffsetSouthWest",
    };
  }
  return {
    southFromStationM: value(point?.southFromStationM),
    westFromStationM: value(point?.westFromStationM),
    zAltitudeM: value(point?.zAltitudeM),
    coordinateMode: "stationOffsetSouthWest",
  };
}

function getGridMultiPointForForm(point) {
  if (point?.coordinateMode !== "stationOffsetSouthWest") return point;
  const gridReference = getCurrentMultiPointGridReference();
  const stationX = parseTotalStationNumber(recordForm?.elements?.tsStationXNorthM?.value);
  const stationY = parseTotalStationNumber(recordForm?.elements?.tsStationYEastM?.value);
  const south = Number(point?.southFromStationM);
  const west = Number(point?.westFromStationM);
  if (gridReference && [stationX, stationY, south, west].every((number) => Number.isFinite(number))) {
    const pointX = stationX + south;
    const pointY = stationY + west;
    return {
      xEastM: String(-(pointY - gridReference.y)),
      yNorthM: String(PLAN_SIZE_CM / 100 - (pointX - gridReference.x)),
      zAltitudeM: value(point?.zAltitudeM),
    };
  }
  return { xEastM: "", yNorthM: "", zAltitudeM: value(point?.zAltitudeM) };
}

function convertTotalStationMultiPointToPlanCoords(record, point) {
  const stationX = parseTotalStationNumber(record?.tsStationXNorthM);
  const stationY = parseTotalStationNumber(record?.tsStationYEastM);
  const south = parseTotalStationNumber(point?.southFromStationM);
  const west = parseTotalStationNumber(point?.westFromStationM);
  if ([stationX, stationY, south, west].some((number) => number == null)) return null;
  const pointX = stationX + south;
  const pointY = stationY + west;
  const gridReference = findTotalStationGridReferencePoint(gridReferenceNameForKuwaku(getRecordKuwaku(record)));
  const zRaw = value(point?.zAltitudeM);
  const z = zRaw && Number.isFinite(Number(zRaw)) ? Number(zRaw) : null;
  if (gridReference) {
    return {
      x: -(pointY - gridReference.y) * 100,
      y: (pointX - gridReference.x) * 100,
      z,
    };
  }
  const peg = parseTotalStationPeg(record?.tsStationPeg);
  const grid = parseKuwaku(getRecordKuwaku(record));
  if (!peg || !grid.block || !grid.no) return null;
  const pegX = (blockLabelToIndex(peg.block) - blockLabelToIndex(grid.block)) * PLAN_SIZE_CM;
  const pegY = (Number(peg.no) - Number(grid.no)) * PLAN_SIZE_CM;
  return { x: pegX - west * 100, y: pegY + south * 100, z };
}

function convertTotalStationPolarMultiPointToPlanCoords(record, point) {
  const stationX = parseTotalStationNumber(record?.tsStationXNorthM);
  const stationY = parseTotalStationNumber(record?.tsStationYEastM);
  const stationZ = parseTotalStationNumber(record?.tsStationAltitudeM);
  const backX = parseTotalStationNumber(record?.tsBacksightXNorthM);
  const backY = parseTotalStationNumber(record?.tsBacksightYEastM);
  const instrumentHeight = parseTotalStationNumber(record?.tsInstrumentHeightM);
  const targetHeight = parseTotalStationNumber(record?.tsTargetHeightM);
  const distance = parseTotalStationNumber(point?.slopeDistanceM);
  const inclination = dmsToDegrees(point?.inclinationDeg, point?.inclinationMin, point?.inclinationSec);
  const direction = dmsToDegrees(point?.directionDeg, point?.directionMin, point?.directionSec);
  if ([stationX, stationY, stationZ, backX, backY, instrumentHeight, targetHeight, distance, inclination, direction].some((number) => number == null)) return null;
  if (distance < 0 || (stationX === backX && stationY === backY)) return null;
  const inclinationRad = inclination * Math.PI / 180;
  const baseAzimuth = Math.atan2(backY - stationY, backX - stationX);
  const azimuth = baseAzimuth + direction * Math.PI / 180;
  const horizontal = distance * Math.cos(inclinationRad);
  const pointX = stationX + horizontal * Math.cos(azimuth);
  const pointY = stationY + horizontal * Math.sin(azimuth);
  const z = stationZ + instrumentHeight + distance * Math.sin(inclinationRad) - targetHeight;
  const gridReference = findTotalStationGridReferencePoint(gridReferenceNameForKuwaku(getRecordKuwaku(record)));
  if (gridReference) {
    return { x: -(pointY - gridReference.y) * 100, y: (pointX - gridReference.x) * 100, z };
  }
  const peg = parseTotalStationPeg(record?.tsStationPeg);
  const grid = parseKuwaku(getRecordKuwaku(record));
  if (!peg || !grid.block || !grid.no) return null;
  const pegX = (blockLabelToIndex(peg.block) - blockLabelToIndex(grid.block)) * PLAN_SIZE_CM;
  const pegY = (Number(peg.no) - Number(grid.no)) * PLAN_SIZE_CM;
  return {
    x: pegX - (pointY - stationY) * 100,
    y: pegY + (pointX - stationX) * 100,
    z,
  };
}

function createDefaultPlanMultiPoint() {
  if (isTotalStationPolarMeasurementSelected()) {
    return {
      slopeDistanceM: "", inclinationDeg: "", inclinationMin: "0", inclinationSec: "0",
      directionDeg: "", directionMin: "0", directionSec: "0", coordinateMode: "stationPolar",
    };
  }
  return isTotalStationMeasurementSelected()
    ? { southFromStationM: "", westFromStationM: "", zAltitudeM: "", coordinateMode: "stationOffsetSouthWest" }
    : { xEastM: "", yNorthM: "", zAltitudeM: "" };
}

function normalizePlanMultiPointEntry(entryRaw) {
  const entry = entryRaw && typeof entryRaw === "object" ? entryRaw : {};
  const coordinateMode = value(entry.coordinateMode);
  if (coordinateMode === "stationPolar" || value(entry.slopeDistanceM)) {
    const normalized = {
      slopeDistanceM: value(entry.slopeDistanceM),
      inclinationDeg: value(entry.inclinationDeg),
      inclinationMin: value(entry.inclinationMin),
      inclinationSec: value(entry.inclinationSec),
      directionDeg: value(entry.directionDeg),
      directionMin: value(entry.directionMin),
      directionSec: value(entry.directionSec),
      coordinateMode: "stationPolar",
    };
    if ([normalized.slopeDistanceM, normalized.inclinationDeg, normalized.inclinationMin, normalized.inclinationSec,
      normalized.directionDeg, normalized.directionMin, normalized.directionSec].some((raw) => parseTotalStationNumber(raw) == null)) {
      return null;
    }
    return normalized;
  }
  if (coordinateMode === "stationOffsetSouthWest" || value(entry.southFromStationM) || value(entry.westFromStationM)) {
    const southFromStationM = value(entry.southFromStationM);
    const westFromStationM = value(entry.westFromStationM);
    const zAltitudeM = value(entry.zAltitudeM);
    if (!Number.isFinite(Number(southFromStationM)) || !Number.isFinite(Number(westFromStationM))) {
      return null;
    }
    return { southFromStationM, westFromStationM, zAltitudeM, coordinateMode: "stationOffsetSouthWest" };
  }
  let xEastM = value(entry.xEastM);
  let yNorthM = value(entry.yNorthM);
  const zAltitudeM = value(entry.zAltitudeM);
  if (!xEastM && value(entry.ewCm)) {
    const ewCm = Number(entry.ewCm);
    xEastM = String((normalizeEwDir(entry.ewDir) === "西から" ? ewCm : PLAN_SIZE_CM - ewCm) / 100);
  }
  if (!yNorthM && value(entry.nsCm)) {
    const nsCm = Number(entry.nsCm);
    yNorthM = String((normalizeNsDir(entry.nsDir) === "南から" ? nsCm : PLAN_SIZE_CM - nsCm) / 100);
  }
  if (!Number.isFinite(Number(xEastM)) || !Number.isFinite(Number(yNorthM))) {
    return null;
  }
  return { xEastM, yNorthM, zAltitudeM };
}

function normalizePlanMultiPoints(pointsRaw) {
  const points = Array.isArray(pointsRaw) ? pointsRaw : [];
  const normalized = [];
  const seen = new Set();
  points.forEach((point) => {
    const entry = normalizePlanMultiPointEntry(point);
    if (!entry) {
      return;
    }
    const key = entry.coordinateMode === "stationPolar"
      ? `polar|${entry.slopeDistanceM}|${entry.inclinationDeg}|${entry.inclinationMin}|${entry.inclinationSec}|${entry.directionDeg}|${entry.directionMin}|${entry.directionSec}`
      : entry.coordinateMode === "stationOffsetSouthWest"
        ? `ts|${entry.southFromStationM}|${entry.westFromStationM}|${entry.zAltitudeM}`
        : `grid|${entry.xEastM}|${entry.yNorthM}|${entry.zAltitudeM}`;
    if (seen.has(key)) {
      return;
    }
    seen.add(key);
    normalized.push(entry);
  });
  return normalized;
}

function createMultiPointRowElement(pointRaw = {}) {
  const normalized = normalizePlanMultiPointEntry(pointRaw) || createDefaultPlanMultiPoint();
  const isTotalStation = isTotalStationMeasurementSelected();
  const isPolar = isTotalStationPolarMeasurementSelected();
  const row = document.createElement("div");
  row.className = "multi-point-row";
  row.dataset.multiPointRow = "1";

  const makeField = (labelText, dataKey, fieldValue) => {
    const label = document.createElement("label");
    label.textContent = labelText;
    const wrap = document.createElement("div");
    wrap.className = "multi-point-input";
    const input = document.createElement("input");
    input.type = "text";
    input.placeholder = "m";
    input.dataset[dataKey] = "1";
    input.value = value(fieldValue);
    const unit = document.createElement("span");
    unit.textContent = "m";
    wrap.append(input, unit);
    label.append(wrap);
    return label;
  };
  const makeDmsField = (labelText, prefix, values) => {
    const label = document.createElement("label");
    label.textContent = labelText;
    const wrap = document.createElement("div");
    wrap.className = "multi-point-dms-input";
    [["Deg", "°"], ["Min", "′"], ["Sec", "″"]].forEach(([suffix, unit], index) => {
      const input = document.createElement("input");
      input.type = "text";
      input.dataset[`multiPoint${prefix}${suffix}`] = "1";
      input.value = value(values[index]);
      const unitEl = document.createElement("span");
      unitEl.textContent = unit;
      wrap.append(input, unitEl);
    });
    label.append(wrap);
    return label;
  };
  let firstLabel;
  let secondLabel;
  let thirdLabel;
  if (isPolar) {
    const polarPoint = normalized.coordinateMode === "stationPolar" ? normalized : createDefaultPlanMultiPoint();
    firstLabel = makeField("斜距離", "multiPointSlopeDistanceM", polarPoint.slopeDistanceM);
    secondLabel = makeDmsField("傾斜", "Inclination", [polarPoint.inclinationDeg, polarPoint.inclinationMin, polarPoint.inclinationSec]);
    thirdLabel = makeDmsField("方向角", "Direction", [polarPoint.directionDeg, polarPoint.directionMin, polarPoint.directionSec]);
  } else if (isTotalStation) {
    const offsets = getTotalStationMultiPointOffsetsForForm(normalized);
    firstLabel = makeField("X　設置点から南へ（南が正）", "multiPointSouthFromStationM", offsets.southFromStationM);
    secondLabel = makeField("Y　設置点から西へ（西が正）", "multiPointWestFromStationM", offsets.westFromStationM);
    thirdLabel = makeField("Z　標高", "multiPointZAltitudeM", normalized.zAltitudeM);
  } else {
    const gridPoint = getGridMultiPointForForm(normalized);
    firstLabel = makeField("x 東西（東が正）", "multiPointXEastM", gridPoint.xEastM);
    secondLabel = makeField("y 南北（北が正）", "multiPointYNorthM", gridPoint.yNorthM);
    thirdLabel = makeField("z 高度", "multiPointZAltitudeM", normalized.zAltitudeM);
  }

  const removeButton = document.createElement("button");
  removeButton.type = "button";
  removeButton.className = "multi-point-remove";
  removeButton.dataset.multiPointRemove = "1";
  removeButton.textContent = "削除";

  row.append(firstLabel, secondLabel, thirdLabel, removeButton);
  if (isTotalStation) {
    const result = document.createElement("p");
    result.className = "multi-point-calculation-result hint-text";
    result.dataset.multiPointCalculationResult = "1";
    result.textContent = "入力すると、北から・西からの距離を計算します。";
    row.append(result);
  }
  return row;
}

function renderMultiPointRows(pointsRaw = []) {
  if (!multiPointRows) {
    return;
  }
  const points = normalizePlanMultiPoints(pointsRaw);
  const source = points.length ? points : [createDefaultPlanMultiPoint()];
  multiPointRows.innerHTML = "";
  source.forEach((point) => {
    multiPointRows.append(createMultiPointRowElement(point));
  });
  syncMultiPointRemoveButtonState();
  updateMultiPointCalculationResults();
}

function syncMultiPointCoordinateModeUi() {
  const isTotalStation = isTotalStationMeasurementSelected();
  const isPolar = isTotalStationPolarMeasurementSelected();
  const hint = document.getElementById("multi-point-hint");
  if (hint) {
    hint.textContent = isPolar
      ? "各点の斜距離、傾斜（度・分・秒）、方向角（度・分・秒）を入力します。"
      : isTotalStation
        ? "各点のX（設置点から南・南が正）、Y（設置点から西・西が正）、Z（標高）をm単位で入力します。"
        : "各点の x（東西・東が正）、y（南北・北が正）、z（高度）をm単位で入力できます。";
  }
  const firstRow = multiPointRows?.querySelector("[data-multi-point-row]");
  if (!firstRow) return;
  const rowMode = firstRow.querySelector("[data-multi-point-slope-distance-m]")
    ? "polar"
    : firstRow.querySelector("[data-multi-point-south-from-station-m]")
      ? "offset"
      : "grid";
  const targetMode = isPolar ? "polar" : isTotalStation ? "offset" : "grid";
  if (rowMode !== targetMode) {
    const points = readMultiPointRowsFromForm();
    renderMultiPointRows(points);
  }
}

function syncMultiPointRemoveButtonState() {
  if (!multiPointRows) {
    return;
  }
  const rows = [...multiPointRows.querySelectorAll("[data-multi-point-row]")];
  const canRemove = rows.length > 1;
  rows.forEach((row) => {
    const removeButton = row.querySelector("[data-multi-point-remove]");
    if (removeButton instanceof HTMLButtonElement) {
      removeButton.disabled = !canRemove;
    }
  });
}

function readMultiPointRowsFromForm() {
  if (!multiPointRows) {
    return [];
  }
  const rows = [...multiPointRows.querySelectorAll("[data-multi-point-row]")];
  const points = rows.map((row) => {
    const slopeField = row.querySelector("[data-multi-point-slope-distance-m]");
    if (slopeField) {
      return {
        slopeDistanceM: value(slopeField.value),
        inclinationDeg: value(row.querySelector("[data-multi-point-inclination-deg]")?.value),
        inclinationMin: value(row.querySelector("[data-multi-point-inclination-min]")?.value),
        inclinationSec: value(row.querySelector("[data-multi-point-inclination-sec]")?.value),
        directionDeg: value(row.querySelector("[data-multi-point-direction-deg]")?.value),
        directionMin: value(row.querySelector("[data-multi-point-direction-min]")?.value),
        directionSec: value(row.querySelector("[data-multi-point-direction-sec]")?.value),
        coordinateMode: "stationPolar",
      };
    }
    const southField = row.querySelector("[data-multi-point-south-from-station-m]");
    const westField = row.querySelector("[data-multi-point-west-from-station-m]");
    if (southField || westField) {
      return {
        southFromStationM: value(southField?.value),
        westFromStationM: value(westField?.value),
        zAltitudeM: value(row.querySelector("[data-multi-point-z-altitude-m]")?.value),
        coordinateMode: "stationOffsetSouthWest",
      };
    }
    return {
      xEastM: value(row.querySelector("[data-multi-point-x-east-m]")?.value),
      yNorthM: value(row.querySelector("[data-multi-point-y-north-m]")?.value),
      zAltitudeM: value(row.querySelector("[data-multi-point-z-altitude-m]")?.value),
    };
  });
  return normalizePlanMultiPoints(points);
}

function readMultiPointRowFromElement(row) {
  if (!(row instanceof Element)) return null;
  const slopeField = row.querySelector("[data-multi-point-slope-distance-m]");
  if (slopeField) {
    return normalizePlanMultiPointEntry({
      slopeDistanceM: value(slopeField.value),
      inclinationDeg: value(row.querySelector("[data-multi-point-inclination-deg]")?.value),
      inclinationMin: value(row.querySelector("[data-multi-point-inclination-min]")?.value),
      inclinationSec: value(row.querySelector("[data-multi-point-inclination-sec]")?.value),
      directionDeg: value(row.querySelector("[data-multi-point-direction-deg]")?.value),
      directionMin: value(row.querySelector("[data-multi-point-direction-min]")?.value),
      directionSec: value(row.querySelector("[data-multi-point-direction-sec]")?.value),
      coordinateMode: "stationPolar",
    });
  }
  const southField = row.querySelector("[data-multi-point-south-from-station-m]");
  const westField = row.querySelector("[data-multi-point-west-from-station-m]");
  if (southField || westField) {
    return normalizePlanMultiPointEntry({
      southFromStationM: value(southField?.value),
      westFromStationM: value(westField?.value),
      zAltitudeM: value(row.querySelector("[data-multi-point-z-altitude-m]")?.value),
      coordinateMode: "stationOffsetSouthWest",
    });
  }
  return null;
}

function formatGridEdgeCalculationResult(coord, { prefix = "計算結果：" } = {}) {
  if (!coord || !Number.isFinite(Number(coord.x)) || !Number.isFinite(Number(coord.y))) return "";
  const hasAltitude = coord.z !== null && coord.z !== undefined && coord.z !== "" && Number.isFinite(Number(coord.z));
  const altitude = hasAltitude ? `、標高 ${Number(Number(coord.z).toFixed(4))} m` : "";
  return `${prefix}北から ${formatLengthInputValue(coord.y)} cm、西から ${formatLengthInputValue(coord.x)} cm${altitude}`;
}

function updateMultiPointCalculationResults() {
  if (!multiPointRows || !isTotalStationMeasurementSelected()) return;
  const draftRecord = buildCurrentRecordDraftForPositionPreview();
  if (!draftRecord) return;
  const rows = [...multiPointRows.querySelectorAll("[data-multi-point-row]")];
  rows.forEach((row, index) => {
    const resultEl = row.querySelector("[data-multi-point-calculation-result]");
    if (!resultEl) return;
    const point = readMultiPointRowFromElement(row);
    const coord = point?.coordinateMode === "stationPolar"
      ? convertTotalStationPolarMultiPointToPlanCoords(draftRecord, point)
      : point?.coordinateMode === "stationOffsetSouthWest"
        ? convertTotalStationMultiPointToPlanCoords(draftRecord, point)
        : null;
    const resultText = formatGridEdgeCalculationResult(coord, { prefix: `計算結果 ${index + 1}：` });
    resultEl.textContent = resultText || `計算結果 ${index + 1}：入力値を確認してください。`;
    resultEl.classList.toggle("total-station-result-ok", Boolean(resultText));
  });
}

function collectPlanMultiPointCoords(record) {
  const points = normalizePlanMultiPoints(record?.multiPoints);
  if (!points.length) {
    return [];
  }
  const coords = [];
  const seen = new Set();
  points.forEach((point) => {
    let coord;
    if (point.coordinateMode === "stationPolar") {
      coord = convertTotalStationPolarMultiPointToPlanCoords(record, point);
    } else if (point.coordinateMode === "stationOffsetSouthWest") {
      coord = convertTotalStationMultiPointToPlanCoords(record, point);
    } else {
      coord = {
        x: Number(point.xEastM) * 100,
        y: PLAN_SIZE_CM - Number(point.yNorthM) * 100,
        z: Number.isFinite(Number(point.zAltitudeM)) ? Number(point.zAltitudeM) : null,
      };
    }
    if (!coord) {
      return;
    }
    const key = `${coord.x.toFixed(4)}|${coord.y.toFixed(4)}`;
    if (seen.has(key)) {
      return;
    }
    seen.add(key);
    coords.push(coord);
  });
  return coords;
}

function readImageAspectRatio(dataUrlRaw) {
  const dataUrl = normalizeCustomLargeImageDataUrl(dataUrlRaw);
  if (!dataUrl) {
    return Promise.resolve(null);
  }
  return new Promise((resolve) => {
    const image = new Image();
    image.onload = () => {
      const width = Math.max(1, Number(image.naturalWidth || image.width) || 0);
      const height = Math.max(1, Number(image.naturalHeight || image.height) || 0);
      if (!width || !height) {
        resolve(null);
        return;
      }
      resolve(width / height);
    };
    image.onerror = () => resolve(null);
    image.src = dataUrl;
  });
}

async function updateCustomLargeImageAspectFromDataUrl(dataUrlRaw = customLargeImageDataUrlInput?.value) {
  const dataUrl = normalizeCustomLargeImageDataUrl(dataUrlRaw);
  if (!customLargeImageAspectInput) {
    return null;
  }
  if (!dataUrl) {
    customLargeImageAspectInput.value = "";
    return null;
  }
  const ratio = await readImageAspectRatio(dataUrl);
  customLargeImageAspectInput.value = normalizeCustomLargeImageAspect(ratio);
  return ratio && ratio > 0 ? ratio : null;
}

function deriveCustomLargeImageNameFromFileName(fileNameRaw) {
  const fileName = value(fileNameRaw);
  if (!fileName) {
    return "";
  }
  const base = fileName.split("/").pop() || fileName;
  const normalized = typeof base.normalize === "function" ? base.normalize("NFC") : base;
  return value(normalized.replace(/\.[^.]+$/, ""));
}

function syncCustomLargeImageStatus() {
  if (!customLargeImageStatus) {
    return;
  }
  const imageName = normalizeCustomLargeImageName(customLargeImageNameInput?.value);
  const hasImage = Boolean(normalizeCustomLargeImageDataUrl(customLargeImageDataUrlInput?.value));
  if (hasImage) {
    customLargeImageStatus.textContent = imageName ? `画像設定済み: ${imageName}` : "画像設定済み";
    return;
  }
  customLargeImageStatus.textContent = "画像未選択";
}

function clearCustomLargeImageFields() {
  if (customLargeImageNameInput) {
    customLargeImageNameInput.value = "";
  }
  if (customLargeImageDataUrlInput) {
    customLargeImageDataUrlInput.value = "";
  }
  if (customLargeImageAspectInput) {
    customLargeImageAspectInput.value = "";
  }
  if (customLargeImageFileInput) {
    customLargeImageFileInput.value = "";
  }
  syncCustomLargeImageStatus();
}

async function setCustomLargeImageFromFile(file) {
  const sourceFile = file instanceof File ? file : null;
  if (!sourceFile) {
    return;
  }
  try {
    const dataUrl = await loadImageFileDataUrlWithFallback(sourceFile, {
      maxLength: 1200,
      quality: 0.85,
      mimeType: "",
    });
    if (customLargeImageDataUrlInput) {
      customLargeImageDataUrlInput.value = normalizeCustomLargeImageDataUrl(dataUrl);
    }
    await updateCustomLargeImageAspectFromDataUrl(dataUrl);
    if (customLargeImageNameInput && !value(customLargeImageNameInput.value)) {
      customLargeImageNameInput.value = deriveCustomLargeImageNameFromFileName(sourceFile.name);
    }
    void syncImageFrameSizeFromAspectLock("customLargeImageDataUrl").then(() => {
      syncLargeShapeImagePreviewTransform();
    });
    syncCustomLargeImageStatus();
    syncLargeShapeImagePreviewForCurrentForm();
  } catch (_error) {
    showToast("画像アップロードに失敗しました");
  }
}

function syncLargeShapeImagePreviewForCurrentForm() {
  const shapeType = normalizeLargeShapeType(largeShapeTypeInput?.value);
  if (!isLargeShapeImageType(shapeType)) {
    syncLargeShapeImagePreview("");
    return;
  }
  const customPath = isCustomLargeShapeType(shapeType) ? normalizeCustomLargeImageDataUrl(customLargeImageDataUrlInput?.value) : "";
  const customName = isCustomLargeShapeType(shapeType) ? normalizeCustomLargeImageName(customLargeImageNameInput?.value) : "";
  syncLargeShapeImagePreview(shapeType, customPath, customName);
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
  const labels = Array.from(new Set([...Array.from(largeShapeImagePathMap.keys()).filter((label) => value(label)), CUSTOM_LARGE_SHAPE_TYPE]));
  largeShapeImageButtons.innerHTML = labels
    .map(
      (label) => {
        const buttonLabel = label === CUSTOM_LARGE_SHAPE_TYPE ? "画像アップロード" : label;
        return `<button class="dir-tab" data-group="largeShapeType" data-value="${escapeHtml(label)}" type="button">${escapeHtml(
          buttonLabel
        )}</button>`;
      }
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

function syncLargeShapeImagePreview(shapeTypeRaw, explicitImagePathRaw = "", explicitTitleRaw = "") {
  if (!largeShapeImagePreview || !largeShapeImagePreviewTitle || !largeShapeImagePreviewImg) {
    return;
  }
  const shapeType = normalizeLargeShapeType(shapeTypeRaw);
  const explicitImagePath = normalizeCustomLargeImageDataUrl(explicitImagePathRaw);
  const candidates = getLargeShapeImagePathCandidates(shapeType, explicitImagePath);
  if (!candidates.length) {
    largeShapeImagePreview.classList.add("hidden");
    largeShapeImagePreviewTitle.textContent = "";
    largeShapeImagePreviewImg.removeAttribute("src");
    largeShapeImagePreviewImg.alt = "";
    largeShapeImagePreviewImg.style.transform = "";
    largeShapeImagePreviewImg.style.width = "";
    largeShapeImagePreviewImg.style.height = "";
    largeShapeImagePreviewImg.style.maxWidth = "";
    largeShapeImagePreviewImg.style.maxHeight = "";
    largeShapeImagePreviewImg.style.objectFit = "";
    return;
  }
  largeShapeImagePreview.classList.remove("hidden");
  const previewTitle = value(explicitTitleRaw) || shapeType;
  largeShapeImagePreviewTitle.textContent = previewTitle;
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
  syncLargeShapeImagePreviewTransform();
}

function syncLargeShapeSectionFromForm() {
  if (!largeShapeSection && !multiPointSection) {
    return;
  }
  const mode = normalizePlanSizeMode(planSizeModeInput?.value);
  if (planSizeModeInput) {
    planSizeModeInput.value = mode;
  }
  const isLarge = mode === "大きなもの";
  const isMultiPoint = mode === "複数点";
  if (largeShapeSection) {
    largeShapeSection.classList.toggle("hidden", !isLarge);
  }
  if (multiPointSection) {
    multiPointSection.classList.toggle("hidden", !isMultiPoint);
  }
  if (isMultiPoint && multiPointRows && !multiPointRows.querySelector("[data-multi-point-row]")) {
    renderMultiPointRows([]);
  }
  if (!isLarge) {
    if (largeShapeTypeInput) {
      largeShapeTypeInput.value = "";
    }
    if (largeAxisDirectionInput) {
      largeAxisDirectionInput.value = "";
    }
    if (largeAxisPlungeInput) {
      largeAxisPlungeInput.value = "";
    }
    if (largeAxisPlungeDirInput) {
      largeAxisPlungeDirInput.value = "";
    }
    if (planeStrikeInput) {
      planeStrikeInput.value = "";
    }
    if (planeDipInput) {
      planeDipInput.value = "";
    }
    if (planeDipDirInput) {
      planeDipDirInput.value = "";
    }
    largeShapePanels.forEach((panel) => {
      panel.classList.add("hidden");
    });
    if (largeAxisDirectionRow) {
      largeAxisDirectionRow.classList.remove("hidden");
    }
    if (largeAxisPlungeRow) {
      largeAxisPlungeRow.classList.remove("hidden");
    }
    if (largeAxisPlungeDirRow) {
      largeAxisPlungeDirRow.classList.remove("hidden");
    }
    if (planeAttitudeRow) {
      planeAttitudeRow.classList.add("hidden");
    }
    if (customLargeImageControls) {
      customLargeImageControls.classList.add("hidden");
    }
    syncLargeShapeImagePreview("");
    return;
  }

  const shapeTypeRaw = normalizeLargeShapeType(largeShapeTypeInput?.value);
  const shapeType = shapeTypeRaw || "直線状";
  const isImageShape = isLargeShapeImageType(shapeType);
  const isCustomImageShape = isCustomLargeShapeType(shapeType);
  const isLineShape = shapeType === "直線状";
  const usesAxisDirection = isLineShape || shapeType === "長方形" || shapeType === "楕円";
  if (largeShapeTypeInput) {
    largeShapeTypeInput.value = shapeType;
  }
  if (largeAxisDirectionInput) {
    const fallbackAxis = !isLineShape && planeStrikeInput ? planeStrikeInput.value : "";
    largeAxisDirectionInput.value = usesAxisDirection ? normalizeLargeAxisDirection(largeAxisDirectionInput.value || fallbackAxis) : "";
  }
  if (largeAxisPlungeInput) {
    largeAxisPlungeInput.value = isLineShape ? normalizeLargeAxisPlungeDeg(largeAxisPlungeInput.value) : "";
  }
  if (largeAxisPlungeDirInput) {
    largeAxisPlungeDirInput.value = isLineShape ? normalizeCompass8Direction(largeAxisPlungeDirInput.value) : "";
  }
  if (planeStrikeInput) {
    const fallbackStrike = usesAxisDirection && largeAxisDirectionInput ? largeAxisDirectionInput.value : "";
    planeStrikeInput.value = isLineShape ? "" : normalizePlaneStrikeDirection(planeStrikeInput.value || fallbackStrike);
  }
  if (planeDipInput) {
    planeDipInput.value = isLineShape ? "" : normalizePlaneDipDeg(planeDipInput.value);
  }
  if (planeDipDirInput) {
    planeDipDirInput.value = isLineShape ? "" : normalizeCompass8Direction(planeDipDirInput.value);
  }
  if (largeAxisDirectionRow) {
    largeAxisDirectionRow.classList.toggle("hidden", !usesAxisDirection);
  }
  if (largeAxisPlungeRow) {
    largeAxisPlungeRow.classList.toggle("hidden", !isLineShape);
  }
  if (largeAxisPlungeDirRow) {
    largeAxisPlungeDirRow.classList.toggle("hidden", !isLineShape);
  }
  if (planeAttitudeRow) {
    planeAttitudeRow.classList.toggle("hidden", isLineShape);
  }
  if (customLargeImageControls) {
    customLargeImageControls.classList.toggle("hidden", !isCustomImageShape);
  }
  largeShapePanels.forEach((panel) => {
    const panelType = value(panel.dataset.largeShapePanel);
    const shouldShow = panelType === shapeType || (panelType === "__IMAGE__" && isImageShape);
    panel.classList.toggle("hidden", !shouldShow);
  });
  syncCustomLargeImageStatus();
  if (isImageShape) {
    syncLargeShapeImagePreviewForCurrentForm();
    void syncImageFrameSizeFromAspectLock("largeShapeType").then(() => {
      syncLargeShapeImagePreviewTransform();
    });
  } else {
    syncLargeShapeImagePreview("");
  }
}

function activateLayerTab(layerRaw) {
  const layer = PRESET_LAYER_NAMES.includes(value(layerRaw)) ? value(layerRaw) : DEFAULT_LAYER_NAME;
  layerNameInput.value = layer;
  layerTabButtons.forEach((button) => {
    button.classList.toggle("active", button.dataset.layer === layer);
  });

  const isOther = layer === OTHER_LAYER_NAME;
  layerOtherInput.classList.toggle("hidden", !isOther);
  if (!isOther) {
    layerOtherInput.value = "";
  }
  renderUnitTabsForLayer(layer);
}

function getUnitOptionsForLayer(layerRaw) {
  const layer = value(layerRaw);
  if (layer === "1.芙蓉湖砂シルト部層") return ["F1", "F2", "F3", "F4"];
  if (layer === "2.立が鼻砂部層") return ["T1", "T2", "T3", "T4", "T5", "T6", "T7"];
  if (layer === "3.海端砂シルト部層") return ["U1", "U2", "U3"];
  return [];
}

function renderUnitTabsForLayer(layerRaw) {
  if (!unitTabs) return;
  const options = getUnitOptionsForLayer(layerRaw);
  unitTabs.innerHTML = options
    .map((unit) => `<button class="unit-tab" data-unit="${unit}" type="button">${unit}</button>`)
    .join("");
  unitTabs.classList.toggle("hidden", !options.length);
  syncUnitTabSelection();
}

function syncUnitTabSelection() {
  if (!unitTabs || !unitInput) return;
  const selectedUnit = value(unitInput.value).toUpperCase();
  unitTabs.querySelectorAll(".unit-tab").forEach((button) => {
    button.classList.toggle("active", value(button.dataset.unit).toUpperCase() === selectedUnit);
  });
}

function setLayerFromValue(layerRaw) {
  const layerValue = normalizeLayerName(value(layerRaw));
  if (!layerValue) {
    activateLayerTab(DEFAULT_LAYER_NAME);
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
  const selected = value(layerNameInput.value) || DEFAULT_LAYER_NAME;
  if (selected !== OTHER_LAYER_NAME) {
    return selected;
  }

  const otherText = value(layerOtherInput.value);
  return otherText ? `${OTHER_LAYER_NAME}:${otherText}` : OTHER_LAYER_NAME;
}

function applyCarryForwardFields(saved) {
  setLayerFromValue(value(saved?.layerName));
  recordForm.elements.unit.value = value(saved?.unit);
  syncUnitTabSelection();
  recordForm.elements.detail.value = value(saved?.detail);
  recordForm.elements.detailSub.value = value(saved?.detailSub);
  recordForm.elements.layerColor.value = getLayerColor(saved);
  recordForm.elements.layerLithology.value = getLayerLithology(saved);
  recordForm.elements.layerFacies.value = composeLayerFacies(getLayerColor(saved), getLayerLithology(saved));
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
  if (getLayerColor(saved)) recordForm.elements.layerColor.classList.add("saved-carry-value");
  if (getLayerLithology(saved)) recordForm.elements.layerLithology.classList.add("saved-carry-value");
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
  recordForm.elements.layerColor.classList.remove("saved-carry-value");
  recordForm.elements.layerLithology.classList.remove("saved-carry-value");
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
    "largeAxisPlungeDeg",
    "largeAxisPlungeDir8",
    "planeStrikeDirection",
    "planeDipDeg",
    "planeDipDir8",
    "imgRotateDeg",
    "imgFrameWidthCm",
    "imgFrameHeightCm",
    "imgSkewXDeg",
    "imgSkewYDeg",
    "imgFlipH",
    "imgFlipV",
    "imgLockAspectRatio",
    "customLargeImageAspect",
    "detail",
    "detailSub",
    "layerColor",
    "layerLithology",
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

function syncAltitudeDirectInputUi({ clearWhenDisabled = false } = {}) {
  const altitudeToggleField = recordForm?.elements?.altitudeInputEnabled;
  const altitudeInputField = recordForm?.elements?.altitudeDirectM;
  if (!(altitudeInputField instanceof HTMLInputElement)) {
    return;
  }
  const isEnabled = altitudeToggleField instanceof HTMLInputElement ? altitudeToggleField.checked : false;
  altitudeInputField.disabled = !isEnabled;
  if (!isEnabled && clearWhenDisabled) {
    altitudeInputField.value = "";
  }
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
      target.name === "layerColor" ||
      target.name === "layerLithology" ||
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
  if (target instanceof HTMLInputElement && target.name === "altitudeInputEnabled") {
    syncAltitudeDirectInputUi({ clearWhenDisabled: true });
  }
  if (target instanceof HTMLElement && target.closest("#large-shape-section")) {
    const targetName =
      target instanceof HTMLInputElement || target instanceof HTMLSelectElement || target instanceof HTMLTextAreaElement
        ? target.name
        : "";
    syncCustomLargeImageStatus();
    if (target === customLargeImageNameInput || target === customLargeImageDataUrlInput || target === customLargeImageFileInput) {
      if (target === customLargeImageDataUrlInput) {
        void updateCustomLargeImageAspectFromDataUrl(customLargeImageDataUrlInput?.value);
      }
      syncLargeShapeImagePreviewForCurrentForm();
    }
    void syncImageFrameSizeFromAspectLock(targetName).then(() => {
      syncLargeShapeImagePreviewTransform();
      updateEditMissingRequiredHighlights();
    });
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

  const normalized = {
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
  return normalized;
}

function normalizeToggleFlag(valueRaw) {
  const text = value(valueRaw).toLowerCase();
  return text === "1" || text === "true" || text === "on" || text === "yes" ? "1" : "0";
}

function normalizeImageRotationDeg(valueRaw) {
  const text = value(valueRaw).replace(/[°度]/g, "");
  if (!text) {
    return "";
  }
  const matched = text.match(/-?\d+(?:\.\d+)?/);
  if (!matched) {
    return "";
  }
  const num = Number(matched[0]);
  if (!Number.isFinite(num)) {
    return "";
  }
  const normalized = ((num % 360) + 360) % 360;
  return Number.isInteger(normalized) ? String(normalized) : String(normalized).replace(/\.?0+$/, "");
}

function normalizeImageSkewDeg(valueRaw) {
  const text = value(valueRaw).replace(/[°度]/g, "");
  if (!text) {
    return "";
  }
  const matched = text.match(/-?\d+(?:\.\d+)?/);
  if (!matched) {
    return "";
  }
  const num = Number(matched[0]);
  if (!Number.isFinite(num)) {
    return "";
  }
  const limited = clamp(num, -80, 80);
  return Number.isInteger(limited) ? String(limited) : String(limited).replace(/\.?0+$/, "");
}

function extractImageTransformFieldsFromFormData(formData) {
  return {
    imgRotateDeg: normalizeImageRotationDeg(formData.get("imgRotateDeg")),
    imgFrameWidthCm: value(formData.get("imgFrameWidthCm")),
    imgFrameHeightCm: value(formData.get("imgFrameHeightCm")),
    imgSkewXDeg: normalizeImageSkewDeg(formData.get("imgSkewXDeg")),
    imgSkewYDeg: normalizeImageSkewDeg(formData.get("imgSkewYDeg")),
    imgFlipH: normalizeToggleFlag(formData.get("imgFlipH")),
    imgFlipV: normalizeToggleFlag(formData.get("imgFlipV")),
    imgLockAspectRatio: normalizeToggleFlag(formData.get("imgLockAspectRatio")),
    imgUseOriginalColor: normalizeToggleFlag(formData.get("imgUseOriginalColor")),
    customLargeImageAspect: normalizeCustomLargeImageAspect(formData.get("customLargeImageAspect")),
  };
}

function syncLargeShapeImagePreviewTransform() {
  if (!largeShapeImagePreviewImg || !recordForm) {
    return;
  }
  const shapeType = normalizeLargeShapeType(largeShapeTypeInput?.value);
  const isImageShape = isLargeShapeImageType(shapeType);
  if (!isImageShape) {
    largeShapeImagePreviewImg.style.transform = "";
    largeShapeImagePreviewImg.style.width = "";
    largeShapeImagePreviewImg.style.height = "";
    largeShapeImagePreviewImg.style.maxWidth = "";
    largeShapeImagePreviewImg.style.maxHeight = "";
    largeShapeImagePreviewImg.style.objectFit = "";
    return;
  }
  const formData = new FormData(recordForm);
  const frameWidthCm = parseDistanceToCm(formData.get("imgFrameWidthCm"));
  const frameHeightCm = parseDistanceToCm(formData.get("imgFrameHeightCm"));
  const rotate = Number(normalizeImageRotationDeg(formData.get("imgRotateDeg")) || "0");
  const skewX = Number(normalizeImageSkewDeg(formData.get("imgSkewXDeg")) || "0");
  const skewY = Number(normalizeImageSkewDeg(formData.get("imgSkewYDeg")) || "0");
  const flipH = normalizeToggleFlag(formData.get("imgFlipH")) === "1";
  const flipV = normalizeToggleFlag(formData.get("imgFlipV")) === "1";
  const scaleX = flipH ? -1 : 1;
  const scaleY = flipV ? -1 : 1;

  if (frameWidthCm != null && frameWidthCm > 0 && frameHeightCm != null && frameHeightCm > 0) {
    const maxSidePx = 280;
    const minSidePx = 52;
    const dominant = Math.max(frameWidthCm, frameHeightCm);
    let widthPx = (frameWidthCm / dominant) * maxSidePx;
    let heightPx = (frameHeightCm / dominant) * maxSidePx;
    const shortSide = Math.min(widthPx, heightPx);
    if (shortSide < minSidePx) {
      const scaleUp = minSidePx / shortSide;
      widthPx *= scaleUp;
      heightPx *= scaleUp;
    }
    largeShapeImagePreviewImg.style.width = `${Math.round(widthPx)}px`;
    largeShapeImagePreviewImg.style.height = `${Math.round(heightPx)}px`;
    largeShapeImagePreviewImg.style.maxWidth = "none";
    largeShapeImagePreviewImg.style.maxHeight = "none";
    // 平面図/3Dと同じく、PNG全体を外枠にマッピングする。
    largeShapeImagePreviewImg.style.objectFit = "fill";
  } else {
    largeShapeImagePreviewImg.style.width = "";
    largeShapeImagePreviewImg.style.height = "";
    largeShapeImagePreviewImg.style.maxWidth = "";
    largeShapeImagePreviewImg.style.maxHeight = "";
    largeShapeImagePreviewImg.style.objectFit = "contain";
  }

  largeShapeImagePreviewImg.style.transformOrigin = "center center";
  largeShapeImagePreviewImg.style.transform = `rotate(${rotate}deg) scale(${scaleX}, ${scaleY}) skew(${skewX}deg, ${skewY}deg)`;
}

async function syncImageFrameSizeFromAspectLock(triggerFieldName = "") {
  if (!recordForm?.elements) {
    return;
  }
  const planSizeMode = normalizePlanSizeMode(recordForm.elements.planSizeMode?.value);
  const shapeType = normalizeLargeShapeType(recordForm.elements.largeShapeType?.value);
  if (planSizeMode !== "大きなもの" || !isCustomLargeShapeType(shapeType)) {
    return;
  }
  const lockInput = recordForm.elements.imgLockAspectRatio;
  if (!(lockInput instanceof HTMLInputElement) || !lockInput.checked) {
    return;
  }
  const widthInput = recordForm.elements.imgFrameWidthCm;
  const heightInput = recordForm.elements.imgFrameHeightCm;
  if (!(widthInput instanceof HTMLInputElement) || !(heightInput instanceof HTMLInputElement)) {
    return;
  }

  let aspectRatio = Number(normalizeCustomLargeImageAspect(customLargeImageAspectInput?.value));
  if (!(aspectRatio > 0)) {
    const loadedAspect = await updateCustomLargeImageAspectFromDataUrl(customLargeImageDataUrlInput?.value);
    aspectRatio = Number(loadedAspect);
  }
  if (!(aspectRatio > 0)) {
    return;
  }

  const widthCm = parseDistanceToCm(widthInput.value);
  const heightCm = parseDistanceToCm(heightInput.value);
  const hasWidth = Number.isFinite(widthCm) && widthCm > 0;
  const hasHeight = Number.isFinite(heightCm) && heightCm > 0;
  if (triggerFieldName === "imgFrameWidthCm" && hasWidth) {
    heightInput.value = formatLengthInputValue(widthCm / aspectRatio);
    return;
  }
  if (triggerFieldName === "imgFrameHeightCm" && hasHeight) {
    widthInput.value = formatLengthInputValue(heightCm * aspectRatio);
    return;
  }
  if (hasWidth && !hasHeight) {
    heightInput.value = formatLengthInputValue(widthCm / aspectRatio);
    return;
  }
  if (!hasWidth && hasHeight) {
    widthInput.value = formatLengthInputValue(heightCm * aspectRatio);
  }
}

function resetRecordForm({ showMessage }) {
  recordForm.reset();
  isOverwriteMode = false;
  overwriteOriginalRecord = null;
  editingRecordId = null;
  activeEditRecordContext = null;
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
  if (recordForm.elements.altitudeInputEnabled instanceof HTMLInputElement) {
    recordForm.elements.altitudeInputEnabled.checked = false;
  }
  if (recordForm.elements.altitudeDirectM instanceof HTMLInputElement) {
    recordForm.elements.altitudeDirectM.value = "";
  }
  if (recordForm.elements.planSizeMode) {
    recordForm.elements.planSizeMode.value = "通常";
  }
  renderMultiPointRows([]);
  if (recordForm.elements.largeShapeType) {
    recordForm.elements.largeShapeType.value = "";
  }
  if (recordForm.elements.customLargeImageName) {
    recordForm.elements.customLargeImageName.value = "";
  }
  if (recordForm.elements.customLargeImageDataUrl) {
    recordForm.elements.customLargeImageDataUrl.value = "";
  }
  if (recordForm.elements.customLargeImageAspect) {
    recordForm.elements.customLargeImageAspect.value = "";
  }
  if (customLargeImageFileInput) {
    customLargeImageFileInput.value = "";
  }
  if (recordForm.elements.largeAxisDirection) {
    recordForm.elements.largeAxisDirection.value = "";
  }
  if (recordForm.elements.largeAxisPlungeDeg) {
    recordForm.elements.largeAxisPlungeDeg.value = "";
  }
  if (recordForm.elements.largeAxisPlungeDir8) {
    recordForm.elements.largeAxisPlungeDir8.value = "";
  }
  if (recordForm.elements.planeStrikeDirection) {
    recordForm.elements.planeStrikeDirection.value = "";
  }
  if (recordForm.elements.planeDipDeg) {
    recordForm.elements.planeDipDeg.value = "";
  }
  if (recordForm.elements.planeDipDir8) {
    recordForm.elements.planeDipDir8.value = "";
  }
  if (recordForm.elements.lineLengthCm) {
    recordForm.elements.lineLengthCm.value = "";
  }
  ["imgRotateDeg", "imgFrameWidthCm", "imgFrameHeightCm", "imgSkewXDeg", "imgSkewYDeg"].forEach((name) => {
    const field = recordForm?.elements?.namedItem(name);
    if (field instanceof HTMLInputElement) {
      field.value = "";
    }
  });
  if (recordForm.elements.imgFlipH instanceof HTMLInputElement) {
    recordForm.elements.imgFlipH.checked = false;
  }
  if (recordForm.elements.imgFlipV instanceof HTMLInputElement) {
    recordForm.elements.imgFlipV.checked = false;
  }
  if (recordForm.elements.imgLockAspectRatio instanceof HTMLInputElement) {
    recordForm.elements.imgLockAspectRatio.checked = false;
  }
  if (recordForm.elements.imgUseOriginalColor instanceof HTMLInputElement) {
    recordForm.elements.imgUseOriginalColor.checked = false;
  }
  clearImageCornerCmFields();
  setDefaultImageCornerDirections();
  setLayerFromValue(DEFAULT_LAYER_NAME);
  syncCustomLargeImageStatus();

  nsDirInput.value = "北から";
  ewDirInput.value = "東から";
  restoreSavedTotalStationSetup();
  syncDirectionTabsFromForm();
  syncAltitudeDirectInputUi({ clearWhenDisabled: true });

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
  if (recordForm.elements.altitudeInputEnabled instanceof HTMLInputElement) {
    recordForm.elements.altitudeInputEnabled.checked = normalizeToggleFlag(record.altitudeInputEnabled) === "1";
  }
  if (recordForm.elements.altitudeDirectM instanceof HTMLInputElement) {
    recordForm.elements.altitudeDirectM.value = value(record.altitudeDirectM);
  }
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
  if (recordForm.elements.customLargeImageName) {
    recordForm.elements.customLargeImageName.value = normalizeCustomLargeImageName(record.customLargeImageName);
  }
  if (recordForm.elements.customLargeImageDataUrl) {
    recordForm.elements.customLargeImageDataUrl.value = normalizeCustomLargeImageDataUrl(record.customLargeImageDataUrl);
  }
  if (recordForm.elements.customLargeImageAspect) {
    recordForm.elements.customLargeImageAspect.value = normalizeCustomLargeImageAspect(record.customLargeImageAspect);
  }
  if (customLargeImageFileInput) {
    customLargeImageFileInput.value = "";
  }
  if (recordForm.elements.largeAxisDirection) {
    recordForm.elements.largeAxisDirection.value = normalizeLargeAxisDirection(record.largeAxisDirection);
  }
  if (recordForm.elements.largeAxisPlungeDeg) {
    recordForm.elements.largeAxisPlungeDeg.value = normalizeLargeAxisPlungeDeg(record.largeAxisPlungeDeg);
  }
  if (recordForm.elements.largeAxisPlungeDir8) {
    recordForm.elements.largeAxisPlungeDir8.value = normalizeCompass8Direction(record.largeAxisPlungeDir8);
  }
  if (recordForm.elements.planeStrikeDirection) {
    recordForm.elements.planeStrikeDirection.value = normalizePlaneStrikeDirection(record.planeStrikeDirection);
  }
  if (recordForm.elements.planeDipDeg) {
    recordForm.elements.planeDipDeg.value = normalizePlaneDipDeg(record.planeDipDeg);
  }
  if (recordForm.elements.planeDipDir8) {
    recordForm.elements.planeDipDir8.value = normalizeCompass8Direction(record.planeDipDir8);
  }
  if (recordForm.elements.lineLengthCm) {
    recordForm.elements.lineLengthCm.value = value(record.lineLengthCm);
  }
  if (recordForm.elements.imgRotateDeg) {
    recordForm.elements.imgRotateDeg.value = normalizeImageRotationDeg(record.imgRotateDeg);
  }
  if (recordForm.elements.imgFrameWidthCm) {
    recordForm.elements.imgFrameWidthCm.value = value(record.imgFrameWidthCm);
  }
  if (recordForm.elements.imgFrameHeightCm) {
    recordForm.elements.imgFrameHeightCm.value = value(record.imgFrameHeightCm);
  }
  if (recordForm.elements.imgSkewXDeg) {
    recordForm.elements.imgSkewXDeg.value = normalizeImageSkewDeg(record.imgSkewXDeg);
  }
  if (recordForm.elements.imgSkewYDeg) {
    recordForm.elements.imgSkewYDeg.value = normalizeImageSkewDeg(record.imgSkewYDeg);
  }
  if (recordForm.elements.imgFlipH instanceof HTMLInputElement) {
    recordForm.elements.imgFlipH.checked = normalizeToggleFlag(record.imgFlipH) === "1";
  }
  if (recordForm.elements.imgFlipV instanceof HTMLInputElement) {
    recordForm.elements.imgFlipV.checked = normalizeToggleFlag(record.imgFlipV) === "1";
  }
  if (recordForm.elements.imgLockAspectRatio instanceof HTMLInputElement) {
    recordForm.elements.imgLockAspectRatio.checked = normalizeToggleFlag(record.imgLockAspectRatio) === "1";
  }
  if (recordForm.elements.imgUseOriginalColor instanceof HTMLInputElement) {
    recordForm.elements.imgUseOriginalColor.checked = normalizeToggleFlag(record.imgUseOriginalColor) === "1";
  }

  nsDirInput.value = normalizeNsDir(record.nsDir);
  recordForm.elements.nsCm.value = record.nsCm || "";
  ewDirInput.value = normalizeEwDir(record.ewDir);
  recordForm.elements.ewCm.value = record.ewCm || "";
  setPositionMeasurementFields(record);
  renderMultiPointRows(record.multiPoints);
  syncDirectionTabsFromForm();
  syncAltitudeDirectInputUi();

  setLayerFromValue(record.layerName);
  recordForm.elements.detail.value = record.detail || "";
  recordForm.elements.detailSub.value = record.detailSub || "";
  recordForm.elements.layerColor.value = getLayerColor(record);
  recordForm.elements.layerLithology.value = getLayerLithology(record);
  recordForm.elements.layerFacies.value = composeLayerFacies(getLayerColor(record), getLayerLithology(record));
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
  syncCustomLargeImageStatus();
  void syncImageFrameSizeFromAspectLock("populate").then(() => {
    syncLargeShapeImagePreviewTransform();
  });
  updateEditMissingRequiredHighlights();
}

function openRecordForEdit(recordId, preferredKuwaku = "", recordIndexRaw = "") {
  const record = findRecordByEditContext(recordId, recordIndexRaw, null);
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
  const recordIndex = state.records.findIndex((item) => item === record);
  activeEditRecordContext = {
    recordId: value(record.id),
    recordIndex: String(recordIndex >= 0 ? recordIndex : ""),
    recordSnapshot: buildCellEditRecordSnapshot(record),
  };
  clearOverwriteUpdatedState();
  populateRecordForm(record);
  renderEditHistory(record);
  setActiveTab("edit-tab");
  renderRecordTable();
  updateEditMissingRequiredHighlights();
}

function scrollToDetailInputTop() {
  if (!recordForm) {
    return;
  }
  const anchor = recordForm.querySelector(".detail-input-title-row") || recordForm;
  window.requestAnimationFrame(() => {
    if (anchor instanceof HTMLElement) {
      anchor.scrollIntoView({ behavior: "smooth", block: "start" });
    }
  });
}

function buildCurrentEditDraftRecord() {
  if (!recordForm) {
    return null;
  }
  const formData = new FormData(recordForm);
  const teamState = normalizeTeamState(value(editTeamInput?.value), value(editTeamOtherInput?.value));
  const draftPlanSizeMode = normalizePlanSizeMode(value(formData.get("planSizeMode")));
  const draftMultiPoints = draftPlanSizeMode === "複数点" ? readMultiPointRowsFromForm() : [];
  const specimenPrefix = normalizeSpecimenPrefix(value(formData.get("specimenPrefix")));
  const specimenSerial = compactNoSpaceValue(formData.get("specimenSerial"));
  const draftRawShapeType = value(formData.get("largeShapeType"));
  const draftShapeType = normalizeLargeShapeType(draftRawShapeType) || normalizeLargeShapeLabel(draftRawShapeType);
  const draftIsImageShape = isLargeShapeImageType(draftShapeType);
  const draftIsCustomImageShape = draftIsImageShape && isCustomLargeShapeType(draftShapeType);
  const draftIsLineShape = draftShapeType === "直線状";
  const draftUsesAxisDirection = draftIsLineShape || draftShapeType === "長方形" || draftShapeType === "楕円";
  const imageCornerFields = extractImageCornerFieldsFromFormData(formData);
  const imageTransformFields = extractImageTransformFieldsFromFormData(formData);
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
    altitudeInputEnabled: normalizeToggleFlag(formData.get("altitudeInputEnabled")),
    altitudeDirectM:
      normalizeToggleFlag(formData.get("altitudeInputEnabled")) === "1" ? value(formData.get("altitudeDirectM")) : "",
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
    positionMethod: normalizePositionMethod(formData.get("positionMethod")),
    tsCoordinateConvention: "southWestPositive",
    tsStationPeg: normalizeTotalStationPointName(formData.get("tsStationPeg")),
    tsStationXNorthM: value(formData.get("tsStationXNorthM")),
    tsStationYEastM: value(formData.get("tsStationYEastM")),
    tsStationAltitudeM: value(formData.get("tsStationAltitudeM")),
    tsBacksightPeg: normalizeTotalStationPointName(formData.get("tsBacksightPeg")),
    tsBacksightXNorthM: value(formData.get("tsBacksightXNorthM")),
    tsBacksightYEastM: value(formData.get("tsBacksightYEastM")),
    tsBacksightAltitudeM: value(formData.get("tsBacksightAltitudeM")),
    tsInstrumentHeightM: value(formData.get("tsInstrumentHeightM")),
    tsTargetHeightM: value(formData.get("tsTargetHeightM")),
    tsObservationMode: value(formData.get("tsObservationMode")) === "polar" ? "polar" : "coordinate",
    tsPointCoordinateMode: "stationOffsetSouthWest",
    multiPoints: draftMultiPoints,
    planSizeMode: draftPlanSizeMode,
    largeShapeType: draftShapeType,
    largeAxisDirection:
      draftIsImageShape || !draftUsesAxisDirection ? "" : normalizeLargeAxisDirection(value(formData.get("largeAxisDirection"))),
    largeAxisPlungeDeg: draftIsImageShape || !draftIsLineShape ? "" : normalizeLargeAxisPlungeDeg(value(formData.get("largeAxisPlungeDeg"))),
    largeAxisPlungeDir8: draftIsImageShape || !draftIsLineShape ? "" : normalizeCompass8Direction(value(formData.get("largeAxisPlungeDir8"))),
    planeStrikeDirection: draftIsLineShape
      ? ""
      : normalizePlaneStrikeDirection(value(formData.get("planeStrikeDirection")) || (draftUsesAxisDirection ? value(formData.get("largeAxisDirection")) : "")),
    planeDipDeg: draftIsLineShape ? "" : normalizePlaneDipDeg(value(formData.get("planeDipDeg"))),
    planeDipDir8: draftIsLineShape ? "" : normalizeCompass8Direction(value(formData.get("planeDipDir8"))),
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
    imgRotateDeg: draftIsImageShape ? imageTransformFields.imgRotateDeg : "",
    imgFrameWidthCm: draftIsImageShape ? imageTransformFields.imgFrameWidthCm : "",
    imgFrameHeightCm: draftIsImageShape ? imageTransformFields.imgFrameHeightCm : "",
    imgSkewXDeg: draftIsImageShape ? imageTransformFields.imgSkewXDeg : "",
    imgSkewYDeg: draftIsImageShape ? imageTransformFields.imgSkewYDeg : "",
    imgFlipH: draftIsImageShape ? imageTransformFields.imgFlipH : "0",
    imgFlipV: draftIsImageShape ? imageTransformFields.imgFlipV : "0",
    imgLockAspectRatio: draftIsImageShape ? imageTransformFields.imgLockAspectRatio : "0",
    imgUseOriginalColor: draftIsImageShape ? imageTransformFields.imgUseOriginalColor : "0",
    customLargeImageName: draftIsCustomImageShape ? normalizeCustomLargeImageName(value(formData.get("customLargeImageName"))) : "",
    customLargeImageDataUrl: draftIsCustomImageShape ? normalizeCustomLargeImageDataUrl(value(formData.get("customLargeImageDataUrl"))) : "",
    customLargeImageAspect: draftIsCustomImageShape ? imageTransformFields.customLargeImageAspect : "",
    layerName: getSelectedLayerName(),
    unit: compactNoSpaceValue(formData.get("unit")),
    detail: compactNoSpaceValue(formData.get("detail")),
    detailSub: value(formData.get("detailSub")),
    layerColor: value(formData.get("layerColor")),
    layerLithology: value(formData.get("layerLithology")),
    layerFacies: composeLayerFacies(formData.get("layerColor"), formData.get("layerLithology")),
    layerRef: value(formData.get("layerRef")),
    layerRelative: value(formData.get("layerRelative")),
    layerFromCm: value(formData.get("layerFromCm")),
    notes: value(formData.get("notes")),
    sectionDiagrams: clonePhotos(currentSectionDiagrams),
    photos: clonePhotos(currentPhotos),
  };
}

function copyCurrentEditToInput() {
  const activeTabId = getActiveTabId();
  if (activeTabId !== "edit-tab" && activeTabId !== "input-tab") {
    return;
  }
  const draft = buildCurrentEditDraftRecord();
  if (!draft) {
    showToast("入力内容のコピーに失敗しました");
    return;
  }
  if (activeTabId === "input-tab" && siteForm) {
    const siteFormData = new FormData(siteForm);
    const siteKuwakuHeadA = normalizeKuwakuHeadA(siteFormData.get("kuwakuHeadA"));
    const siteKuwakuHeadB = normalizeKuwakuHeadB(siteFormData.get("kuwakuHeadB"));
    const siteKuwakuBlock = normalizeKuwakuBlock(siteFormData.get("kuwakuBlock"));
    const siteKuwakuNo = normalizeKuwakuNo(siteFormData.get("kuwakuNo"));
    const siteTeamState = normalizeTeamState(value(siteFormData.get("team")), value(siteFormData.get("teamOther")));
    draft.kuwaku = buildKuwaku(siteKuwakuHeadA, siteKuwakuHeadB, siteKuwakuBlock, siteKuwakuNo);
    draft.levelHeight = value(siteFormData.get("levelHeight"));
    draft.date = value(siteFormData.get("date"));
    draft.team = siteTeamState.team;
    draft.teamOther = siteTeamState.teamOther;
    draft.teamLead = value(siteFormData.get("teamLead"));
    draft.recorder = value(siteFormData.get("recorder"));
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
  activeEditRecordContext = null;
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
  activeEditRecordContext = null;
  if (recordIdInput) {
    recordIdInput.value = "";
  }
  updateDuplicateSpecimenWarning();
  showToast("コピーして新規入力を作成しました");
}

function copySavedRecordToInput(recordId, preferredKuwaku = "", recordRaw = null, options = {}) {
  const record = recordRaw && typeof recordRaw === "object" ? recordRaw : findRecord(recordId);
  if (!record) {
    showToast("対象データが見つかりません");
    return;
  }
  const shouldShowToast = options && typeof options === "object" ? options.showToast !== false : true;

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
  activeEditRecordContext = null;
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
  activeEditRecordContext = null;
  if (recordIdInput) {
    recordIdInput.value = "";
  }
  updateDuplicateSpecimenWarning();
  if (shouldShowToast) {
    showToast("コピーして新規入力を作成しました");
  }
}

function insertRowFromList(recordId, preferredKuwaku = "", recordRaw = null) {
  const record = recordRaw && typeof recordRaw === "object" ? recordRaw : findRecord(recordId);
  if (!record) {
    showToast("対象データが見つかりません");
    return;
  }
  const kuwaku = normalizeKuwakuText(value(preferredKuwaku) || value(record.kuwaku) || getRecordKuwaku(record));
  const teamState = normalizeTeamState(value(record.team), value(record.teamOther));
  const nowIsoValue = nowIso();
  const initialPrefix = DEFAULT_SPECIMEN_PREFIX;
  const initialSerial = findSmallestUnusedSpecimenSerial(kuwaku, initialPrefix, "");
  const insertedBase = {
    id: newId("record"),
    kuwaku,
    specimenPrefix: initialPrefix,
    specimenSerial: initialSerial,
    specimenNo: buildSpecimenNo(initialPrefix, initialSerial),
    category: categoryFromPrefix(initialPrefix),
    analysisType: "",
    levelHeight: "",
    date: value(record.date),
    team: teamState.team,
    teamOther: teamState.teamOther,
    teamLead: value(record.teamLead),
    recorder: value(record.recorder),
    nameMemo: "",
    unit: "",
    discoverer: "",
    identifier: "",
    levelUpperCm: "",
    levelLowerCm: "",
    altitudeInputEnabled: "",
    altitudeDirectM: "",
    occurrenceSection: "要",
    occurrenceSketch: "要",
    sectionDiagrams: [],
    photos: [],
    sectionDiagramDistanceChecked: "",
    sectionDiagramHorizonChecked: "",
    sectionDiagramLayerFaciesChecked: "",
    photoClinometerChecked: "",
    photoRulerChecked: "",
    nsDir: "北から",
    nsCm: "",
    ewDir: "東から",
    ewCm: "",
    multiPoints: [],
    planSizeMode: "通常",
    largeShapeType: "",
    largeAxisDirection: "",
    largeAxisPlungeDeg: "",
    largeAxisPlungeDir8: "",
    planeStrikeDirection: "",
    planeDipDeg: "",
    planeDipDir8: "",
    lineLengthCm: "",
    line1NsDir: "",
    line1NsCm: "",
    line1EwDir: "",
    line1EwCm: "",
    line2NsDir: "",
    line2NsCm: "",
    line2EwDir: "",
    line2EwCm: "",
    rectSide1Cm: "",
    rectSide2Cm: "",
    ellipseLongRadiusCm: "",
    ellipseShortRadiusCm: "",
    imgP1NsDir: "",
    imgP1NsCm: "",
    imgP1EwDir: "",
    imgP1EwCm: "",
    imgP2NsDir: "",
    imgP2NsCm: "",
    imgP2EwDir: "",
    imgP2EwCm: "",
    imgP3NsDir: "",
    imgP3NsCm: "",
    imgP3EwDir: "",
    imgP3EwCm: "",
    imgP4NsDir: "",
    imgP4NsCm: "",
    imgP4EwDir: "",
    imgP4EwCm: "",
    imgRotateDeg: "",
    imgFrameWidthCm: "",
    imgFrameHeightCm: "",
    imgSkewXDeg: "",
    imgSkewYDeg: "",
    imgFlipH: "0",
    imgFlipV: "0",
    imgLockAspectRatio: "0",
    imgUseOriginalColor: "0",
    customLargeImageName: "",
    customLargeImageDataUrl: "",
    customLargeImageAspect: "",
    importantFlag: "無",
    simpleRecordFlag: "-",
    layerName: DEFAULT_LAYER_NAME,
    detail: "",
    detailSub: "",
    layerColor: "",
    layerLithology: "",
    layerFacies: "",
    layerRef: "",
    layerFromCm: "",
    layerRelative: "",
    notes: "",
    createdAt: nowIsoValue,
    updatedAt: nowIsoValue,
    deletedAt: "",
  };
  const insertedRecord = {
    ...insertedBase,
    history: buildNextRecordHistory(null, insertedBase, "行挿入"),
  };
  state.records.unshift(insertedRecord);
  persist("行挿入しました");
  if (getActiveTabId() !== "output-tab") {
    setActiveTab("output-tab");
  }
  selectedCardRecordId = "";
  renderRecordTable();
  renderOutputs();
  const insertedIndex = state.records.findIndex((item) => item === insertedRecord);
  window.requestAnimationFrame(() => {
    openOutputCellEditModal(insertedRecord.id, "category", String(insertedIndex >= 0 ? insertedIndex : ""));
  });
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
  if (name === "customLargeImageDataUrl") {
    if (customLargeImageFileInput) {
      customLargeImageFileInput.classList.add("edit-missing-field");
    }
    if (customLargeImageControls) {
      customLargeImageControls.classList.add("edit-missing-group");
    }
    return;
  }
  if (name === "customLargeImageName") {
    if (customLargeImageNameInput) {
      customLargeImageNameInput.classList.add("edit-missing-field");
    }
    if (customLargeImageControls) {
      customLargeImageControls.classList.add("edit-missing-group");
    }
    return;
  }
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
  if (missingKeys.has("multiPoints")) {
    markEditMissingGroupByName("planSizeMode");
  }
  if (missingKeys.has("largeShapeType")) {
    markEditMissingGroupByName("planSizeMode");
  }
  if (missingKeys.has("largeAxisDirection")) {
    markEditMissingFieldByName("largeAxisDirection");
  }
  if (missingKeys.has("largeAxisPlungeDir8")) {
    markEditMissingGroupByName("largeAxisPlungeDir8");
  }
  if (missingKeys.has("planeDipDir8")) {
    markEditMissingGroupByName("planeDipDir8");
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
    "imgRotateDeg",
    "imgFrameWidthCm",
    "imgFrameHeightCm",
    "customLargeImageName",
    "customLargeImageDataUrl",
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
  if (!recordTableBody && !editRecordTableBody) {
    return;
  }
  if (!state.records.length) {
    renderRecordTableBodyRows(recordTableBody, [], "まだ入力データがありません。");
    renderRecordTableBodyRows(editRecordTableBody, [], "まだ入力データがありません。");
    return;
  }
  const inputRecords = getInputRecordsForCurrentKuwaku();
  const editRecords = getEditRecordsForCurrentKuwaku();
  renderRecordTableBodyRows(
    recordTableBody,
    inputRecords,
    inputRecords.length ? "" : "現在の区画（グリッド）の入力データがありません。"
  );
  renderRecordTableBodyRows(
    editRecordTableBody,
    editRecords,
    editRecords.length ? "" : "編集中の区画（グリッド）の入力データがありません。"
  );
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

function getEditRecordsForCurrentKuwaku() {
  const editKuwaku = currentKuwakuForDuplicateWarning("edit-tab");
  const sortedRecords = [...state.records].sort(compareRecordsByKuwakuThenSpecimen);
  if (!editKuwaku) {
    return [];
  }
  const currentValue = kuwakuValueForSelect(editKuwaku);
  return sortedRecords.filter((record) => kuwakuValueForSelect(getRecordKuwaku(record)) === currentValue);
}

function renderRecordTableBodyRows(targetBody, records, emptyMessage) {
  if (!targetBody) {
    return;
  }
  if (!records.length) {
    targetBody.innerHTML = `<tr><td colspan="8">${escapeHtml(emptyMessage || "表示対象データがありません。")}</td></tr>`;
    return;
  }
  const recordIndexMap = new Map();
  state.records.forEach((record, index) => {
    if (record && typeof record === "object") {
      recordIndexMap.set(record, index);
    }
  });
  targetBody.innerHTML = records
    .map((record) => {
      const recordIndex = Number(recordIndexMap.get(record));
      return buildRecordTableRowHtml(record, Number.isInteger(recordIndex) && recordIndex >= 0 ? recordIndex : "");
    })
    .join("");
}

function buildRecordTableRowHtml(record, recordIndexRaw = "") {
  const recordIndex = value(recordIndexRaw);
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
            <button type="button" data-action="insert-row" data-id="${record.id}" data-kuwaku="${escapeHtml(
              getRecordKuwaku(record)
            )}" data-record-index="${escapeHtml(recordIndex)}">行挿入</button>
            <button type="button" data-action="edit" data-id="${record.id}" data-kuwaku="${escapeHtml(
              getRecordKuwaku(record)
            )}" data-record-index="${escapeHtml(recordIndex)}">編集</button>
            <button type="button" data-action="copy-to-input" data-id="${record.id}" data-kuwaku="${escapeHtml(
              getRecordKuwaku(record)
            )}" data-record-index="${escapeHtml(recordIndex)}">コピーして新規入力</button>
            <button class="danger" type="button" data-action="delete" data-id="${record.id}" data-record-index="${escapeHtml(recordIndex)}">削除</button>
          </div>
        </td>
      </tr>
      `;
}

function handleRecordTableActionClick(event) {
  const button = event.target.closest("button[data-action]");
  if (!button) {
    return;
  }
  const sourceTableBody = event.currentTarget;
  const shouldScrollToDetailTop = sourceTableBody === recordTableBody || sourceTableBody === editRecordTableBody;
  const recordId = button.dataset.id;
  const row = button.closest("tr[data-record-index]");
  const recordIndex = value(button.dataset.recordIndex) || value(row?.dataset?.recordIndex);
  const action = button.dataset.action;
  const record = findRecordByEditContext(recordId, recordIndex, null);
  if (!record) {
    showToast("対象データが見つかりません");
    return;
  }

  if (action === "edit") {
    const rowKuwaku = value(button.dataset.kuwaku);
    openRecordForEdit(record.id, rowKuwaku, recordIndex);
    if (shouldScrollToDetailTop) {
      scrollToDetailInputTop();
    }
    return;
  }
  if (action === "copy-to-input") {
    const rowKuwaku = value(button.dataset.kuwaku);
    copySavedRecordToInput(recordId, rowKuwaku, record);
    return;
  }
  if (action === "insert-row") {
    const rowKuwaku = value(button.dataset.kuwaku);
    insertRowFromList(recordId, rowKuwaku, record);
    return;
  }
  if (action === "delete") {
    const answer = window.confirm(`標本番号 ${record.specimenNo} を削除しますか？`);
    if (!answer) {
      return;
    }
    const deletingId = value(record.id) || recordId;
    state.records = state.records.filter((item) => item !== record);
    if (editingRecordId === deletingId) {
      resetRecordForm({ showMessage: false });
    }
    persist("記録を削除しました");
    renderRecordTable();
    renderOutputs();
  }
}

function renderOutputs() {
  renderListOutput();
  if (getActiveTabId() === "plan-tab") {
    renderPlanOutput();
  }
  if (getActiveTabId() === "viewer-tab") {
    renderViewerOutput({ preserveCamera: true });
  }
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

function filterRecordsByCategory(recordsRaw, categoryValueRaw) {
  const records = Array.isArray(recordsRaw) ? recordsRaw : [];
  const categoryValue = value(categoryValueRaw) || EXPORT_CATEGORY_ALL_VALUE;
  if (!categoryValue || categoryValue === EXPORT_CATEGORY_ALL_VALUE) {
    return records;
  }
  return records.filter((record) => {
    const specimen = parseSpecimenNo(record?.specimenNo, record?.specimenPrefix, record?.specimenSerial);
    const prefix = normalizeSpecimenPrefix(specimen.prefix);
    return prefix === categoryValue;
  });
}

function syncPlanCategorySelect(recordsRaw) {
  if (!planCategorySelect) {
    return;
  }
  const records = Array.isArray(recordsRaw) ? recordsRaw : [];
  const options = collectExportCategoryOptions(records);
  if (!options.some((item) => item.value === selectedPlanCategory)) {
    selectedPlanCategory = EXPORT_CATEGORY_ALL_VALUE;
  }
  planCategorySelect.innerHTML = options
    .map(
      (item) =>
        `<option value="${escapeHtml(item.value)}" ${item.value === selectedPlanCategory ? "selected" : ""}>${escapeHtml(
          item.label
        )}</option>`
    )
    .join("");
}

function syncViewerCategorySelect(recordsRaw) {
  if (!viewerCategorySelect) {
    return;
  }
  const records = Array.isArray(recordsRaw) ? recordsRaw : [];
  const options = collectExportCategoryOptions(records);
  if (!options.some((item) => item.value === selectedViewerCategory)) {
    selectedViewerCategory = EXPORT_CATEGORY_ALL_VALUE;
  }
  viewerCategorySelect.innerHTML = options
    .map(
      (item) =>
        `<option value="${escapeHtml(item.value)}" ${item.value === selectedViewerCategory ? "selected" : ""}>${escapeHtml(
          item.label
        )}</option>`
    )
    .join("");
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

function parseRecordIndex(recordIndexRaw) {
  const index = Number(recordIndexRaw);
  if (!Number.isInteger(index) || index < 0) {
    return -1;
  }
  return index;
}

function buildCellEditRecordSnapshot(record) {
  return {
    id: value(record?.id),
    specimenNo: parseSpecimenNo(record?.specimenNo, record?.specimenPrefix, record?.specimenSerial).specimenNo,
    kuwaku: normalizeKuwakuText(getRecordKuwaku(record)),
    createdAt: value(record?.createdAt),
  };
}

function matchesCellEditRecordSnapshot(record, snapshotRaw) {
  const snapshot = snapshotRaw && typeof snapshotRaw === "object" ? snapshotRaw : {};
  if (!record) {
    return false;
  }
  if (value(snapshot.id) && value(record.id) !== value(snapshot.id)) {
    return false;
  }
  if (value(snapshot.specimenNo)) {
    const specimenNo = parseSpecimenNo(record.specimenNo, record.specimenPrefix, record.specimenSerial).specimenNo;
    if (specimenNo !== value(snapshot.specimenNo)) {
      return false;
    }
  }
  if (value(snapshot.kuwaku) && normalizeKuwakuText(getRecordKuwaku(record)) !== value(snapshot.kuwaku)) {
    return false;
  }
  if (value(snapshot.createdAt) && value(record.createdAt) !== value(snapshot.createdAt)) {
    return false;
  }
  return true;
}

function findRecordByEditContext(recordIdRaw, recordIndexRaw = "", snapshotRaw = null) {
  const recordId = value(recordIdRaw);
  const recordIndex = parseRecordIndex(recordIndexRaw);
  const snapshot = snapshotRaw && typeof snapshotRaw === "object" ? snapshotRaw : null;

  if (recordIndex >= 0 && recordIndex < state.records.length) {
    const byIndex = state.records[recordIndex];
    if (byIndex && (!recordId || value(byIndex.id) === recordId) && matchesCellEditRecordSnapshot(byIndex, snapshot)) {
      return byIndex;
    }
  }

  if (snapshot) {
    const exactMatches = state.records.filter((item) => matchesCellEditRecordSnapshot(item, snapshot));
    if (exactMatches.length === 1) {
      return exactMatches[0];
    }
  }

  if (recordId) {
    const idMatches = state.records.filter((item) => value(item?.id) === recordId);
    if (idMatches.length === 1) {
      return idMatches[0];
    }
    if (snapshot) {
      const idAndSnapshotMatches = idMatches.filter((item) => matchesCellEditRecordSnapshot(item, snapshot));
      if (idAndSnapshotMatches.length === 1) {
        return idAndSnapshotMatches[0];
      }
    }
  }

  return null;
}

function openOutputCellEditModal(recordId, editKey, recordIndexRaw = "") {
  if (!cellEditModal || !cellEditTitle || !cellEditMeta || !cellEditFields) {
    return;
  }
  const record = findRecordByEditContext(recordId, recordIndexRaw, null);
  if (!record) {
    showToast("対象データが見つかりません");
    return;
  }
  const label = OUTPUT_CELL_EDIT_LABELS[editKey];
  if (!label) {
    return;
  }
  const fieldsHtml = buildOutputCellEditFieldsHtml(record, editKey);
  if (!fieldsHtml) {
    showToast("このセルは直接編集に対応していません");
    return;
  }

  activeOutputCellEdit = {
    recordId: value(record.id),
    editKey: value(editKey),
    recordIndex: String(parseRecordIndex(recordIndexRaw)),
    recordSnapshot: buildCellEditRecordSnapshot(record),
  };
  cellEditTitle.textContent = `${label}を編集`;
  cellEditMeta.textContent = `標本番号 ${record.specimenNo || "-"} / 区画 ${getRecordKuwaku(record) || "-"}`;
  cellEditFields.innerHTML = fieldsHtml;
  cellEditModal.classList.remove("hidden");
  bindOutputCellEditDynamicFields(editKey);
  const focusTarget = cellEditFields.querySelector("input, select, textarea");
  if (focusTarget instanceof HTMLElement) {
    focusTarget.focus();
    if (focusTarget instanceof HTMLInputElement || focusTarget instanceof HTMLTextAreaElement) {
      focusTarget.select?.();
    }
  }
}

function closeOutputCellEditModal() {
  if (!cellEditModal || !cellEditFields) {
    return;
  }
  cellEditModal.classList.add("hidden");
  cellEditFields.innerHTML = "";
  activeOutputCellEdit = null;
}

function buildOutputCellEditFieldsHtml(record, editKey) {
  const teamState = normalizeTeamState(record?.team, record?.teamOther);
  const teamOtherHiddenClass = teamState.team === OTHER_TEAM_NAME ? "" : " hidden";
  const parsedSpecimen = parseSpecimenNo(record?.specimenNo, record?.specimenPrefix, record?.specimenSerial);
  const specimenPrefix = normalizeSpecimenPrefix(parsedSpecimen.prefix);
  const analysisHiddenClass = specimenPrefix === "a" ? "" : " hidden";
  const analysisType = normalizeAnalysisType(record?.analysisType);
  const altitudeDirectEnabled = normalizeToggleFlag(record?.altitudeInputEnabled) === "1";
  const altitudeDirectHiddenClass = altitudeDirectEnabled ? "" : " hidden";

  switch (editKey) {
    case "kuwaku":
      return `
        <label>
          <span class="label-title">区画（グリッド）</span>
          <input name="kuwaku" type="text" value="${escapeHtml(getRecordKuwaku(record))}" placeholder="例: 25-Ⅰ-C-4" />
        </label>
      `;
    case "team":
      return `
        <label>
          <span class="label-title">発掘班</span>
          <select name="team" data-cell-edit-team-select>
            <option value="" ${teamState.team === "" ? "selected" : ""}>未設定</option>
            <option value="1" ${teamState.team === "1" ? "selected" : ""}>1</option>
            <option value="2" ${teamState.team === "2" ? "selected" : ""}>2</option>
            <option value="3" ${teamState.team === "3" ? "selected" : ""}>3</option>
            <option value="4" ${teamState.team === "4" ? "selected" : ""}>4</option>
            <option value="その他" ${teamState.team === "その他" ? "selected" : ""}>その他</option>
          </select>
        </label>
        <label data-cell-edit-team-other class="${teamOtherHiddenClass}">
          <span class="label-title">発掘班（その他）</span>
          <input name="teamOther" type="text" value="${escapeHtml(teamState.teamOther)}" />
        </label>
      `;
    case "date":
      return `
        <label>
          <span class="label-title">日付</span>
          <input name="date" type="date" value="${escapeHtml(normalizeDateForExportRange(getRecordDate(record)))}" />
        </label>
      `;
    case "specimenNo":
      return `
        <label>
          <span class="label-title">標本番号</span>
          <input name="specimenNo" type="text" value="${escapeHtml(parsedSpecimen.specimenNo)}" placeholder="例: m-1" />
        </label>
      `;
    case "category":
      return `
        <label>
          <span class="label-title">分類</span>
          <select name="specimenPrefix" data-cell-edit-prefix-select>
            ${buildOutputCellEditPrefixOptionsHtml(specimenPrefix)}
          </select>
        </label>
        <label>
          <span class="label-title">番号</span>
          <div class="specimen-number-row">
            <span data-cell-edit-prefix-label class="specimen-prefix-chip">${escapeHtml(specimenPrefix)}</span>
            <span>-</span>
            <input name="specimenSerial" type="text" inputmode="numeric" value="${escapeHtml(
              parsedSpecimen.serial
            )}" placeholder="例: 1" data-cell-edit-specimen-serial />
          </div>
          <span class="hint-text">この区画・分類で未使用の最小番号を自動入力しています。変更もできます。</span>
        </label>
        <label data-cell-edit-analysis-row class="${analysisHiddenClass}">
          <span class="label-title">分析用試料の区分</span>
          <select name="analysisType">
            ${buildOutputCellEditAnalysisOptionsHtml(analysisType)}
          </select>
        </label>
      `;
    case "nameMemo":
      return `
        <label>
          <span class="label-title">化石・遺物名称</span>
          <input name="nameMemo" type="text" value="${escapeHtml(value(record?.nameMemo))}" />
        </label>
      `;
    case "importantFlag":
      return `
        <label>
          <span class="label-title">重要品指定</span>
          <select name="importantFlag">
            <option value="無" ${value(record?.importantFlag) === "無" ? "selected" : ""}>無</option>
            <option value="有" ${value(record?.importantFlag) === "有" ? "selected" : ""}>有</option>
          </select>
        </label>
      `;
    case "unit":
      return `
        <label>
          <span class="label-title">ユニット</span>
          <input name="unit" type="text" value="${escapeHtml(value(record?.unit))}" />
        </label>
      `;
    case "detail":
      return `
        <label>
          <span class="label-title">サブユニット</span>
          <input name="detail" type="text" value="${escapeHtml(value(record?.detail))}" />
        </label>
        <label>
          <span class="label-title">細分</span>
          <input name="detailSub" type="text" value="${escapeHtml(value(record?.detailSub))}" />
        </label>
      `;
    case "discoverer":
      return `
        <label>
          <span class="label-title">発見者氏名</span>
          <input name="discoverer" type="text" value="${escapeHtml(value(record?.discoverer))}" />
        </label>
      `;
    case "identifier":
      return `
        <label>
          <span class="label-title">判定者氏名</span>
          <input name="identifier" type="text" value="${escapeHtml(value(record?.identifier))}" />
        </label>
      `;
    case "levelRead":
      return `
        <label>
          <span class="label-title">レベル読値（上面）</span>
          <div class="inline-cm-row compact-cm-row">
            <input name="levelUpperCm" type="text" value="${escapeHtml(value(record?.levelUpperCm))}" />
            <span>cm</span>
          </div>
        </label>
        <label>
          <span class="label-title">レベル読値（下底）</span>
          <div class="inline-cm-row compact-cm-row">
            <input name="levelLowerCm" type="text" value="${escapeHtml(value(record?.levelLowerCm))}" />
            <span>cm</span>
          </div>
        </label>
      `;
    case "altitudeM":
      return `
        <label>
          <span class="label-title">レベル高</span>
          <div class="inline-cm-row compact-cm-row">
            <input name="levelHeight" type="text" value="${escapeHtml(value(record?.levelHeight))}" />
            <span>m</span>
          </div>
        </label>
        <label class="level-altitude-toggle">
          <input name="altitudeInputEnabled" type="checkbox" value="1" ${altitudeDirectEnabled ? "checked" : ""} data-cell-edit-altitude-check />
          標高（下底）を直接入力
        </label>
        <label data-cell-edit-altitude-direct class="${altitudeDirectHiddenClass}">
          <span class="label-title">標高（下底）</span>
          <div class="inline-cm-row compact-cm-row">
            <input name="altitudeDirectM" type="text" value="${escapeHtml(value(record?.altitudeDirectM))}" />
            <span>m</span>
          </div>
        </label>
      `;
    case "occurrenceSection":
      return `
        <label>
          <span class="label-title">産出状況断面</span>
          <select name="occurrenceSection">
            <option value="要" ${value(record?.occurrenceSection) === "否" ? "" : "selected"}>要</option>
            <option value="否" ${value(record?.occurrenceSection) === "否" ? "selected" : ""}>否</option>
          </select>
        </label>
      `;
    case "occurrenceSketch":
      return `
        <label>
          <span class="label-title">産状スケッチ</span>
          <select name="occurrenceSketch">
            <option value="要" ${value(record?.occurrenceSketch) === "否" ? "" : "selected"}>要</option>
            <option value="否" ${value(record?.occurrenceSketch) === "否" ? "selected" : ""}>否</option>
          </select>
        </label>
      `;
    case "position":
      return `
        <label>
          <span class="label-title">北から / 南から</span>
          <div class="inline-cm-row compact-cm-row">
            <select name="nsDir">
              <option value="北から" ${normalizeNsDir(record?.nsDir) === "北から" ? "selected" : ""}>北から</option>
              <option value="南から" ${normalizeNsDir(record?.nsDir) === "南から" ? "selected" : ""}>南から</option>
            </select>
            <input name="nsCm" type="text" value="${escapeHtml(value(record?.nsCm))}" />
            <span>cm</span>
          </div>
        </label>
        <label>
          <span class="label-title">東から / 西から</span>
          <div class="inline-cm-row compact-cm-row">
            <select name="ewDir">
              <option value="東から" ${normalizeEwDir(record?.ewDir) === "東から" ? "selected" : ""}>東から</option>
              <option value="西から" ${normalizeEwDir(record?.ewDir) === "西から" ? "selected" : ""}>西から</option>
            </select>
            <input name="ewCm" type="text" value="${escapeHtml(value(record?.ewCm))}" />
            <span>cm</span>
          </div>
        </label>
      `;
    case "notes":
      return `
        <label>
          <span class="label-title">備考</span>
          <textarea name="notes" rows="4">${escapeHtml(value(record?.notes))}</textarea>
        </label>
      `;
    default:
      return "";
  }
}

function bindOutputCellEditDynamicFields(editKey) {
  if (!cellEditFields) {
    return;
  }
  if (editKey === "team") {
    const teamSelect = cellEditFields.querySelector("[data-cell-edit-team-select]");
    const teamOtherLabel = cellEditFields.querySelector("[data-cell-edit-team-other]");
    if (teamSelect instanceof HTMLSelectElement && teamOtherLabel instanceof HTMLElement) {
      const toggle = () => {
        const isOther = value(teamSelect.value) === OTHER_TEAM_NAME;
        teamOtherLabel.classList.toggle("hidden", !isOther);
      };
      teamSelect.addEventListener("change", toggle);
      toggle();
    }
  }
  if (editKey === "category") {
    const prefixSelect = cellEditFields.querySelector("[data-cell-edit-prefix-select]");
    const analysisRow = cellEditFields.querySelector("[data-cell-edit-analysis-row]");
    const specimenSerialInput = cellEditFields.querySelector("[data-cell-edit-specimen-serial]");
    const prefixLabel = cellEditFields.querySelector("[data-cell-edit-prefix-label]");
    if (prefixSelect instanceof HTMLSelectElement && analysisRow instanceof HTMLElement) {
      const toggle = (useAutoSerial = false) => {
        const prefix = normalizeSpecimenPrefix(prefixSelect.value);
        analysisRow.classList.toggle("hidden", prefix !== "a");
        if (prefixLabel instanceof HTMLElement) {
          prefixLabel.textContent = prefix;
        }
        if (specimenSerialInput instanceof HTMLInputElement) {
          if (useAutoSerial || !compactNoSpaceValue(specimenSerialInput.value)) {
            specimenSerialInput.value = findSmallestUnusedSpecimenSerialForActiveEdit(prefix);
          }
        }
      };
      prefixSelect.addEventListener("change", () => toggle(true));
      toggle(false);
    }
  }
  if (editKey === "altitudeM") {
    const check = cellEditFields.querySelector("[data-cell-edit-altitude-check]");
    const directRow = cellEditFields.querySelector("[data-cell-edit-altitude-direct]");
    const directInput = directRow?.querySelector("input[name='altitudeDirectM']");
    if (check instanceof HTMLInputElement && directRow instanceof HTMLElement) {
      const toggle = () => {
        const enabled = check.checked;
        directRow.classList.toggle("hidden", !enabled);
        if (directInput instanceof HTMLInputElement) {
          directInput.disabled = !enabled;
          if (!enabled) {
            directInput.value = "";
          }
        }
      };
      check.addEventListener("change", toggle);
      toggle();
    }
  }
}

function saveOutputCellEditFromModal() {
  if (!activeOutputCellEdit || !cellEditForm) {
    return;
  }
  const record = findRecordByEditContext(
    activeOutputCellEdit.recordId,
    activeOutputCellEdit.recordIndex,
    activeOutputCellEdit.recordSnapshot
  );
  if (!record) {
    closeOutputCellEditModal();
    showToast("対象データが見つかりません（一覧を再表示してから再度編集してください）");
    return;
  }
  const formData = createOutputCellEditFormData();
  const result = applyOutputCellEditToRecord(record, activeOutputCellEdit.editKey, formData);
  if (!result.ok) {
    if (result.message) {
      showToast(result.message);
    }
    return;
  }
  record.updatedAt = nowIso();
  const specimenNoForMessage = parseSpecimenNo(record.specimenNo, record.specimenPrefix, record.specimenSerial).specimenNo || record.specimenNo;
  persist(`標本番号 ${specimenNoForMessage || "-"} を上書き保存しました`);
  renderRecordTable();
  renderOutputs();
  closeOutputCellEditModal();
}

function createOutputCellEditFormData() {
  const values = {};
  cellEditFields?.querySelectorAll("input[name], select[name], textarea[name]").forEach((field) => {
    if (!(field instanceof HTMLInputElement || field instanceof HTMLSelectElement || field instanceof HTMLTextAreaElement) || field.disabled) return;
    const name = value(field.name);
    if (!name) return;
    if (field instanceof HTMLInputElement && (field.type === "checkbox" || field.type === "radio")) {
      if (field.checked) values[name] = field.value;
      return;
    }
    values[name] = field.value;
  });
  return { get: (name) => (Object.prototype.hasOwnProperty.call(values, name) ? values[name] : null) };
}

function applyOutputCellEditToRecord(record, editKey, formData) {
  if (!(record && formData && typeof formData.get === "function")) {
    return { ok: false, message: "編集データを取得できませんでした" };
  }
  if (editKey === "kuwaku") {
    const parsed = parseKuwaku(formData.get("kuwaku"));
    const nextKuwaku = buildKuwaku(parsed.headA, parsed.headB, parsed.block, parsed.no);
    if (!parsed.block || !parsed.no || !nextKuwaku) {
      return { ok: false, message: "区画（例: 25-Ⅰ-C-4）を入力してください" };
    }
    const duplicate = findDuplicateRecordByKuwakuAndSpecimen(nextKuwaku, record.specimenNo, record.id);
    if (duplicate) {
      return { ok: false, message: `この区画には ${record.specimenNo} がすでにあります` };
    }
    record.kuwaku = nextKuwaku;
    return { ok: true };
  }
  if (editKey === "team") {
    const teamState = normalizeTeamState(formData.get("team"), formData.get("teamOther"));
    record.team = teamState.team;
    record.teamOther = teamState.teamOther;
    return { ok: true };
  }
  if (editKey === "date") {
    record.date = value(formData.get("date"));
    return { ok: true };
  }
  if (editKey === "specimenNo") {
    const parsed = parseSpecimenNo(formData.get("specimenNo"), record.specimenPrefix, record.specimenSerial);
    if (!parsed.serial) {
      return { ok: false, message: "標本番号を入力してください（例: m-1）" };
    }
    const nextSpecimenNo = buildSpecimenNo(parsed.prefix, parsed.serial);
    const duplicate = findDuplicateRecordByKuwakuAndSpecimen(getRecordKuwaku(record), nextSpecimenNo, record.id);
    if (duplicate) {
      return { ok: false, message: `この区画には ${nextSpecimenNo} がすでにあります` };
    }
    record.specimenPrefix = normalizeSpecimenPrefix(parsed.prefix);
    record.specimenSerial = compactNoSpaceValue(parsed.serial);
    record.specimenNo = nextSpecimenNo;
    record.category = categoryFromPrefix(record.specimenPrefix);
    if (record.specimenPrefix !== "a") {
      record.analysisType = "";
    }
    return { ok: true };
  }
  if (editKey === "category") {
    const nextPrefix = normalizeSpecimenPrefix(formData.get("specimenPrefix"));
    const nextSerial = compactNoSpaceValue(formData.get("specimenSerial"));
    if (!nextSerial) {
      return { ok: false, message: "番号を入力してください" };
    }
    if (!/^\d+$/.test(nextSerial)) {
      return { ok: false, message: "番号は半角数字で入力してください" };
    }
    const nextSpecimenNo = buildSpecimenNo(nextPrefix, nextSerial);
    const duplicate = findDuplicateRecordByKuwakuAndSpecimen(getRecordKuwaku(record), nextSpecimenNo, record.id);
    if (duplicate) {
      return { ok: false, message: `この区画には ${nextSpecimenNo} がすでにあります` };
    }
    record.specimenPrefix = nextPrefix;
    record.specimenSerial = nextSerial;
    record.specimenNo = nextSpecimenNo;
    record.category = categoryFromPrefix(nextPrefix);
    record.analysisType = nextPrefix === "a" ? normalizeAnalysisType(formData.get("analysisType")) : "";
    return { ok: true };
  }
  if (editKey === "nameMemo") {
    record.nameMemo = value(formData.get("nameMemo"));
    return { ok: true };
  }
  if (editKey === "importantFlag") {
    record.importantFlag = normalizeHasFlag(formData.get("importantFlag")) || "無";
    return { ok: true };
  }
  if (editKey === "unit") {
    record.unit = compactNoSpaceValue(formData.get("unit"));
    return { ok: true };
  }
  if (editKey === "detail") {
    record.detail = compactNoSpaceValue(formData.get("detail"));
    record.detailSub = value(formData.get("detailSub"));
    return { ok: true };
  }
  if (editKey === "discoverer") {
    record.discoverer = value(formData.get("discoverer"));
    return { ok: true };
  }
  if (editKey === "identifier") {
    record.identifier = value(formData.get("identifier"));
    return { ok: true };
  }
  if (editKey === "levelRead") {
    record.levelUpperCm = value(formData.get("levelUpperCm"));
    record.levelLowerCm = value(formData.get("levelLowerCm"));
    return { ok: true };
  }
  if (editKey === "altitudeM") {
    record.levelHeight = value(formData.get("levelHeight"));
    const useDirectAltitude = normalizeToggleFlag(formData.get("altitudeInputEnabled")) === "1";
    record.altitudeInputEnabled = useDirectAltitude ? "1" : "";
    record.altitudeDirectM = useDirectAltitude ? value(formData.get("altitudeDirectM")) : "";
    return { ok: true };
  }
  if (editKey === "occurrenceSection") {
    record.occurrenceSection = normalizeNeedFlag(formData.get("occurrenceSection"));
    return { ok: true };
  }
  if (editKey === "occurrenceSketch") {
    record.occurrenceSketch = normalizeNeedFlag(formData.get("occurrenceSketch"));
    return { ok: true };
  }
  if (editKey === "position") {
    record.nsDir = normalizeNsDir(formData.get("nsDir"));
    record.nsCm = value(formData.get("nsCm"));
    record.ewDir = normalizeEwDir(formData.get("ewDir"));
    record.ewCm = value(formData.get("ewCm"));
    return { ok: true };
  }
  if (editKey === "notes") {
    record.notes = value(formData.get("notes"));
    return { ok: true };
  }
  return { ok: false, message: "このセルは直接編集に対応していません" };
}

function findSmallestUnusedSpecimenSerialForActiveEdit(prefixRaw) {
  if (!activeOutputCellEdit) return "1";
  const record = findRecordByEditContext(
    activeOutputCellEdit.recordId,
    activeOutputCellEdit.recordIndex,
    activeOutputCellEdit.recordSnapshot
  );
  return record ? findSmallestUnusedSpecimenSerial(getRecordKuwaku(record), prefixRaw, record.id) : "1";
}

function findSmallestUnusedSpecimenSerial(kuwakuRaw, prefixRaw, excludeRecordIdRaw) {
  const kuwaku = normalizeKuwakuText(kuwakuRaw);
  const prefix = normalizeSpecimenPrefix(prefixRaw);
  const excludeRecordId = value(excludeRecordIdRaw);
  const used = new Set();
  state.records.forEach((item) => {
    if (!item || value(item.id) === excludeRecordId || normalizeKuwakuText(getRecordKuwaku(item)) !== kuwaku) return;
    const specimen = parseSpecimenNo(item.specimenNo, item.specimenPrefix, item.specimenSerial);
    if (normalizeSpecimenPrefix(specimen.prefix) !== prefix) return;
    const serial = compactNoSpaceValue(specimen.serial);
    if (/^\d+$/.test(serial)) used.add(Number(serial));
  });
  let next = 1;
  while (used.has(next)) next += 1;
  return String(next);
}

function buildOutputCellEditPrefixOptionsHtml(selectedPrefixRaw) {
  const selectedPrefix = normalizeSpecimenPrefix(selectedPrefixRaw);
  const order = ["m", "b", "l", "s", "i", "g", "h", "a"];
  return order
    .map((prefix) => {
      const label = `${prefix}: ${SPECIMEN_CATEGORY_MAP[prefix] || ""}`;
      return `<option value="${prefix}" ${prefix === selectedPrefix ? "selected" : ""}>${escapeHtml(label)}</option>`;
    })
    .join("");
}

function buildOutputCellEditAnalysisOptionsHtml(selectedTypeRaw) {
  const selectedType = normalizeAnalysisType(selectedTypeRaw);
  const order = ["A", "C", "M", "F", "P", "B", "I", "D", "R", "S", "H", "MG"];
  const options = [`<option value="" ${selectedType ? "" : "selected"}>未設定</option>`];
  order.forEach((code) => {
    const displayCode = code === "MG" ? "Mg" : code;
    const optionValue = `${displayCode}: ${ANALYSIS_TYPE_MAP[code]}`;
    options.push(`<option value="${escapeHtml(optionValue)}" ${optionValue === selectedType ? "selected" : ""}>${escapeHtml(optionValue)}</option>`);
  });
  return options.join("");
}

function isOutputListColumnVisible(columnKeyRaw) {
  const columnKey = value(columnKeyRaw);
  if (!columnKey) {
    return true;
  }
  return outputListColumnVisibility[columnKey] !== false;
}

function getOutputListVisibleColumnCount() {
  const visibleCount = OUTPUT_LIST_COLUMN_DEFS.reduce((count, column) => {
    return count + (isOutputListColumnVisible(column.key) ? 1 : 0);
  }, 0);
  return visibleCount > 0 ? visibleCount : 1;
}

function renderOutputColumnToggleRow() {
  if (!outputColumnToggleRow) {
    return;
  }
  outputColumnToggleRow.innerHTML = OUTPUT_LIST_COLUMN_DEFS.map((column) => {
    const visible = isOutputListColumnVisible(column.key);
    return `<button type="button" class="output-column-toggle-btn ${visible ? "active" : ""}" data-col-key="${escapeHtml(
      column.key
    )}" aria-pressed="${visible ? "true" : "false"}">${escapeHtml(column.label)}</button>`;
  }).join("");
}

function applyOutputListColumnVisibility() {
  if (!outputListTable) {
    return;
  }
  const allCells = outputListTable.querySelectorAll("[data-col-key]");
  allCells.forEach((cell) => {
    const columnKey = value(cell.dataset.colKey);
    if (!columnKey) {
      return;
    }
    cell.hidden = !isOutputListColumnVisible(columnKey);
  });
}

function toggleOutputListColumnVisibility(columnKeyRaw) {
  const columnKey = value(columnKeyRaw);
  if (!columnKey || !Object.prototype.hasOwnProperty.call(outputListColumnVisibility, columnKey)) {
    return;
  }
  if (isOutputListColumnVisible(columnKey) && getOutputListVisibleColumnCount() <= 1) {
    showToast("少なくとも1列は表示してください");
    return;
  }
  outputListColumnVisibility[columnKey] = !isOutputListColumnVisible(columnKey);
  renderListOutput();
}

function renderListOutput() {
  renderOutputColumnToggleRow();
  updateOutputListSortHeader();
  if (!state.records.length) {
    syncOutputKuwakuSelect([]);
    syncOutputCategorySelect([]);
    syncOutputStatusSelect();
    syncOutputDateSelect([]);
    syncOutputSearchInput();
    updateOutputFilterSummary(0, 0);
    outputListBody.innerHTML = `<tr><td colspan="${getOutputListVisibleColumnCount()}">出力対象データがありません。</td></tr>`;
    applyOutputListColumnVisibility();
    return;
  }

  const filteredRecords = getFilteredOutputRecords();
  if (!filteredRecords.length) {
    outputListBody.innerHTML = `<tr><td colspan="${getOutputListVisibleColumnCount()}">条件に一致するデータがありません。</td></tr>`;
    applyOutputListColumnVisibility();
    return;
  }

  const recordIndexMap = new Map();
  state.records.forEach((record, index) => {
    if (record && typeof record === "object") {
      recordIndexMap.set(record, index);
    }
  });

  outputListBody.innerHTML = sortOutputRecordsForList(filteredRecords)
    .map((record) => {
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
        "imgRotateDeg",
        "imgFrameWidthCm",
        "imgFrameHeightCm",
        "customLargeImageName",
        "customLargeImageDataUrl",
      ]);
      const recordIndex = Number(recordIndexMap.get(record));
      return `
      <tr data-record-id="${escapeHtml(value(record.id))}" data-record-index="${Number.isInteger(recordIndex) && recordIndex >= 0 ? recordIndex : ""}">
        <td data-col-key="kuwaku" data-cell-edit-key="kuwaku" class="${listCellClass("kuwaku-color-cell", kuwakuMissing)}" style="background:${kuwakuStyle.background};color:${kuwakuStyle.color};border-color:${kuwakuStyle.border};" ${missingTitle}>${escapeHtml(
          kuwakuText
        )}</td>
        <td data-col-key="team" data-cell-edit-key="team" class="${listCellClass("", teamMissing)}" ${missingTitle}>${escapeHtml(getRecordTeamValue(record))}</td>
        <td data-col-key="date" data-cell-edit-key="date">${escapeHtml(getRecordDate(record))}</td>
        <td data-col-key="dataStatus" class="${listCellClass(dataComplete ? "data-status-complete" : "data-status-incomplete", !dataComplete)}" ${missingTitle}>${
          dataComplete ? "○" : "未記入"
        }</td>
        <td data-col-key="specimenNo" data-cell-edit-key="specimenNo" class="${listCellClass("", specimenMissing)}" ${missingTitle}>${escapeHtml(record.specimenNo)}</td>
        <td data-col-key="category" data-cell-edit-key="category" class="${listCellClass("category-color-cell", categoryMissing)}" style="background:${categoryBackground};color:#111827;border-color:${categoryBorderColor};" ${missingTitle}>${escapeHtml(
          formatCategoryForRecord(record)
        )}</td>
        <td data-col-key="nameMemo" data-cell-edit-key="nameMemo" class="${listCellClass("", nameMissing)}" ${missingTitle}>${escapeHtml(record.nameMemo || "")}</td>
        <td data-col-key="importantFlag" data-cell-edit-key="importantFlag" class="${listCellClass(record.importantFlag === "有" ? "important-cell-important" : "", importantMissing)}" ${missingTitle}>${escapeHtml(
          record.importantFlag || ""
        )}</td>
        <td data-col-key="unit" data-cell-edit-key="unit" class="${listCellClass("unit-color-cell", unitMissing)}" style="background:${unitStyle.background};color:${unitStyle.color};border-color:${unitStyle.border};" ${missingTitle}>${escapeHtml(
          record.unit || ""
        )}</td>
        <td data-col-key="detail" data-cell-edit-key="detail">${escapeHtml(formatDetailForRecord(record))}</td>
        <td data-col-key="discoverer" data-cell-edit-key="discoverer" class="${listCellClass("", discovererMissing)}" ${missingTitle}>${escapeHtml(record.discoverer || "")}</td>
        <td data-col-key="identifier" data-cell-edit-key="identifier" class="${listCellClass("", identifierMissing)}" ${missingTitle}>${escapeHtml(record.identifier || "")}</td>
        <td data-col-key="levelRead" data-cell-edit-key="levelRead" class="${listCellClass("", levelReadMissing)}" ${missingTitle}>${escapeHtml(formatLevelRead(record))}</td>
        <td data-col-key="altitudeM" data-cell-edit-key="altitudeM" class="${listCellClass("", levelHeightMissing || levelReadMissing)}" ${missingTitle}>${escapeHtml(formatRecordAltitudeM(record))}</td>
        <td data-col-key="occurrenceSection" data-cell-edit-key="occurrenceSection" class="${listCellClass("", sectionMissing || sectionChecklistMissing)}" ${missingTitle}>${escapeHtml(
          record.occurrenceSection || ""
        )}</td>
        <td data-col-key="occurrenceSketch" data-cell-edit-key="occurrenceSketch" class="${listCellClass("", sketchMissing)}" ${missingTitle}>${escapeHtml(record.occurrenceSketch || "")}</td>
        <td data-col-key="position" data-cell-edit-key="position" class="${listCellClass("", positionMissing)}" ${missingTitle}>${escapeHtml(formatPlanPosition(record))}</td>
        <td data-col-key="notes" data-cell-edit-key="notes">${escapeHtml(record.notes || "")}</td>
        <td data-col-key="actions">
          <div class="row-actions">
            <button type="button" data-action="insert-row" data-id="${record.id}" data-kuwaku="${escapeHtml(
              getRecordKuwaku(record)
            )}" data-record-index="${Number.isInteger(recordIndex) && recordIndex >= 0 ? recordIndex : ""}">行挿入</button>
            <button type="button" data-action="copy-to-input" data-id="${record.id}" data-kuwaku="${escapeHtml(
              getRecordKuwaku(record)
            )}" data-record-index="${Number.isInteger(recordIndex) && recordIndex >= 0 ? recordIndex : ""}">コピーして新規入力</button>
            <button type="button" data-action="edit" data-id="${record.id}" data-kuwaku="${escapeHtml(
              getRecordKuwaku(record)
            )}" data-record-index="${Number.isInteger(recordIndex) && recordIndex >= 0 ? recordIndex : ""}">編集</button>
            <button type="button" data-action="position-preview" data-id="${record.id}" data-record-index="${Number.isInteger(recordIndex) && recordIndex >= 0 ? recordIndex : ""}">平面位置確認</button>
            <button class="danger" type="button" data-action="delete" data-id="${record.id}" data-record-index="${Number.isInteger(recordIndex) && recordIndex >= 0 ? recordIndex : ""}">削除</button>
          </div>
        </td>
      </tr>
      `;
    })
    .join("");
  applyOutputListColumnVisibility();
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
    case "date":
      compared = compareSortDate(getRecordDate(a), getRecordDate(b));
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

function compareSortDate(aRaw, bRaw) {
  const a = normalizeDateForExportRange(aRaw);
  const b = normalizeDateForExportRange(bRaw);
  if (a && b) {
    return a.localeCompare(b, "ja", { numeric: true, sensitivity: "base" });
  }
  if (a) {
    return -1;
  }
  if (b) {
    return 1;
  }
  return compareSortText(aRaw, bRaw);
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
  const totalStationCoords = normalizePositionMethod(record?.positionMethod) === "totalStation"
    ? convertPositionToPlanCoords(record?.nsDir, record?.nsCm, record?.ewDir, record?.ewCm)
    : null;
  if (totalStationCoords) {
    base = `北から${formatCmValue(totalStationCoords.y)}、西から${formatCmValue(totalStationCoords.x)}`;
  } else if (nsPart && ewPart) {
    base = `${nsPart} / ${ewPart}`;
  } else {
    base = nsPart || ewPart;
  }
  const planSizeMode = normalizePlanSizeMode(record?.planSizeMode);
  if (planSizeMode === "複数点") {
    const multiPointCoords = collectPlanMultiPointCoords(record);
    const multiPointCount = multiPointCoords.length;
    if (multiPointCount > 0) {
      if (normalizePositionMethod(record?.positionMethod) === "totalStation") {
        const details = multiPointCoords.map((coord, index) =>
          formatGridEdgeCalculationResult(coord, { prefix: `${index + 1}: ` })
        );
        return `複数点(${multiPointCount}点) / ${details.join(" / ")}`;
      }
      return base ? `複数点(${multiPointCount}点) / ${base}` : `複数点(${multiPointCount}点)`;
    }
  }
  if (planSizeMode !== "大きなもの") {
    return base;
  }
  const axisDirection = normalizeLargeAxisDirection(record?.largeAxisDirection);
  const plungeDeg = normalizeLargeAxisPlungeDeg(record?.largeAxisPlungeDeg);
  const plungeDir8 = normalizeCompass8Direction(record?.largeAxisPlungeDir8);
  const plungeText = plungeDeg ? `プランジ:${plungeDeg}${plungeDir8 ? `(${plungeDir8})` : ""}` : "";
  const planeStrike = normalizePlaneStrikeDirection(record?.planeStrikeDirection);
  const planeDip = normalizePlaneDipDeg(record?.planeDipDeg);
  const planeDipDir8 = normalizeCompass8Direction(record?.planeDipDir8);
  const planeAttitudeText =
    planeStrike && planeDip ? `走向傾斜:${planeStrike}/${planeDip}${planeDipDir8 ? `(${planeDipDir8})` : ""}` : "";
  const shapeType = normalizeLargeShapeType(record?.largeShapeType);
  const lineLength = shapeType === "直線状" ? formatCmValue(record?.lineLengthCm) : "";
  if (!base) {
    if (shapeType === "直線状") {
      if (axisDirection && lineLength) {
        return [`方位:${axisDirection}`, `長さ:${lineLength}`, plungeText].filter(Boolean).join(" / ");
      }
      return [axisDirection ? `方位:${axisDirection}` : "", lineLength ? `長さ:${lineLength}` : "", plungeText]
        .filter(Boolean)
        .join(" / ");
    }
    return [planeAttitudeText].filter(Boolean).join(" / ");
  }
  if (shapeType === "直線状") {
    if (axisDirection && lineLength) {
      return [base, `方位:${axisDirection}`, `長さ:${lineLength}`, plungeText].filter(Boolean).join(" / ");
    }
    if (axisDirection) {
      return [base, `方位:${axisDirection}`, plungeText].filter(Boolean).join(" / ");
    }
    if (lineLength) {
      return [base, `長さ:${lineLength}`, plungeText].filter(Boolean).join(" / ");
    }
  }
  if (shapeType !== "直線状" && planeAttitudeText) {
    return [base, planeAttitudeText].filter(Boolean).join(" / ");
  }
  if (axisDirection || plungeText) {
    return [base, axisDirection ? `方位:${axisDirection}` : "", plungeText].filter(Boolean).join(" / ");
  }
  return base;
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
        <div><span>色</span><strong>${escapeHtml(getLayerColor(selectedRecord))}</strong></div>
        <div><span>岩相</span><strong>${escapeHtml(getLayerLithology(selectedRecord))}</strong></div>
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
  const kuwakuScopedRecords =
    selectedOutputKuwaku === ALL_GRIDS_VALUE
      ? sortedRecords
      : sortedRecords.filter((record) => kuwakuValueForSelect(getRecordKuwaku(record)) === selectedOutputKuwaku);

  syncOutputCategorySelect(kuwakuScopedRecords);
  syncOutputStatusSelect();

  let filteredRecords = kuwakuScopedRecords;
  if (selectedOutputCategory && selectedOutputCategory !== EXPORT_CATEGORY_ALL_VALUE) {
    filteredRecords = filteredRecords.filter((record) => {
      const specimen = parseSpecimenNo(record.specimenNo, record.specimenPrefix, record.specimenSerial);
      return normalizeSpecimenPrefix(specimen.prefix) === selectedOutputCategory;
    });
  }
  if (selectedOutputStatus === "complete") {
    filteredRecords = filteredRecords.filter((record) => isRecordDataComplete(record));
  } else if (selectedOutputStatus === "incomplete") {
    filteredRecords = filteredRecords.filter((record) => !isRecordDataComplete(record));
  }
  syncOutputDateSelect(filteredRecords);
  if (selectedOutputDate) {
    filteredRecords = filteredRecords.filter(
      (record) => normalizeDateForExportRange(getRecordDate(record)) === selectedOutputDate
    );
  }

  syncOutputSearchInput();
  const searchText = value(outputSearchText).toLowerCase();
  if (searchText) {
    filteredRecords = filteredRecords.filter((record) => buildOutputFilterSearchText(record).includes(searchText));
  }

  updateOutputFilterSummary(filteredRecords.length, sortedRecords.length);
  return filteredRecords;
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

function syncOutputCategorySelect(records) {
  if (!outputCategorySelect) {
    return;
  }
  const options = collectExportCategoryOptions(records);
  if (!options.some((item) => item.value === selectedOutputCategory)) {
    selectedOutputCategory = EXPORT_CATEGORY_ALL_VALUE;
  }
  outputCategorySelect.innerHTML = options
    .map(
      (item) =>
        `<option value="${escapeHtml(item.value)}" ${item.value === selectedOutputCategory ? "selected" : ""}>${escapeHtml(
          item.label
        )}</option>`
    )
    .join("");
}

function syncOutputStatusSelect() {
  if (!outputStatusSelect) {
    return;
  }
  if (!["all", "complete", "incomplete"].includes(selectedOutputStatus)) {
    selectedOutputStatus = "all";
  }
  outputStatusSelect.value = selectedOutputStatus;
}

function syncOutputDateSelect(records) {
  if (!outputDateSelect) {
    return;
  }
  const options = collectOutputDateOptions(records);
  if (!options.some((item) => item.value === selectedOutputDate)) {
    selectedOutputDate = "";
  }
  outputDateSelect.innerHTML = options
    .map(
      (item) =>
        `<option value="${escapeHtml(item.value)}" ${item.value === selectedOutputDate ? "selected" : ""}>${escapeHtml(
          item.label
        )}</option>`
    )
    .join("");
}

function syncOutputSearchInput() {
  if (!outputSearchInput) {
    return;
  }
  if (outputSearchInput.value !== outputSearchText) {
    outputSearchInput.value = outputSearchText;
  }
}

function buildOutputFilterSearchText(record) {
  return [
    getRecordKuwaku(record),
    record.specimenNo,
    formatCategoryForRecord(record),
    record.nameMemo,
    record.unit,
    formatDetailForRecord(record),
    getRecordTeamValue(record),
    record.discoverer,
    record.identifier,
    formatPlanPosition(record),
    record.notes,
  ]
    .map((item) => value(item).toLowerCase())
    .join(" ");
}

function updateOutputFilterSummary(filteredCount, totalCount) {
  if (!outputFilterSummary) {
    return;
  }
  const filtered = Number.isFinite(filteredCount) ? filteredCount : 0;
  const total = Number.isFinite(totalCount) ? totalCount : 0;
  outputFilterSummary.textContent = `表示: ${filtered}件 / 全体: ${total}件`;
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

function collectOutputDateOptions(records) {
  if (!records.length) {
    return [{ value: "", label: "すべて" }];
  }
  const dateSet = new Set();
  records.forEach((record) => {
    const normalizedDate = normalizeDateForExportRange(getRecordDate(record));
    if (normalizedDate) {
      dateSet.add(normalizedDate);
    }
  });
  const options = Array.from(dateSet)
    .sort((a, b) => b.localeCompare(a, "ja", { numeric: true, sensitivity: "base" }))
    .map((dateValue) => ({
      value: dateValue,
      label: dateValue,
    }));
  return [{ value: "", label: "すべて" }, ...options];
}

function renderPlanOutput() {
  if (!planMapWrap || !planMapLegend || !planUnitSelect || !planDetailSelect || !planDetailSubSelect) {
    return;
  }
  planMapLegend.innerHTML = buildPlanLegendHtml();

  if (!state.records.length) {
    selectedPlanKuwaku = "";
    selectedPlanCategory = EXPORT_CATEGORY_ALL_VALUE;
    selectedPlanUnit = "";
    selectedPlanDetail = ALL_DETAILS_VALUE;
    selectedPlanDetailSub = ALL_DETAIL_SUBS_VALUE;
    syncPlanKuwakuSelect([]);
    if (planCategorySelect) {
      planCategorySelect.innerHTML = "";
    }
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
  syncPlanCategorySelect(kuwakuFilteredRecords);
  const categoryRecords = filterRecordsByCategory(kuwakuFilteredRecords, selectedPlanCategory);
  if (planKuwakuInfo) {
    const kuwakuLabel = selectedPlanKuwaku ? kuwakuLabelForSelect(selectedPlanKuwaku) : "-";
    planKuwakuInfo.textContent = `区画: ${kuwakuLabel}`;
  }
  if (!categoryRecords.length) {
    selectedPlanUnit = "";
    selectedPlanDetail = ALL_DETAILS_VALUE;
    selectedPlanDetailSub = ALL_DETAIL_SUBS_VALUE;
    planUnitSelect.innerHTML = "";
    planDetailSelect.innerHTML = "";
    planDetailSubSelect.innerHTML = "";
    const kuwakuLabelForMeta = selectedPlanKuwaku ? kuwakuLabelForSelect(selectedPlanKuwaku) : "-";
    const categoryLabelForMeta =
      selectedPlanCategory === EXPORT_CATEGORY_ALL_VALUE
        ? "全分類"
        : `${selectedPlanCategory}: ${SPECIMEN_CATEGORY_MAP[selectedPlanCategory] || ""}`;
    planMapWrap.innerHTML = `
      <div class="plan-map-meta">
        <span>区画（グリッド）: ${escapeHtml(kuwakuLabelForMeta)}</span>
        <span>分類: ${escapeHtml(categoryLabelForMeta)}</span>
        <span>出力階層: -</span>
        <span>表示件数: 0件</span>
      </div>
      <p class="muted">この条件には表示対象データがありません。</p>
    `;
    return;
  }

  const units = collectPlanUnits(categoryRecords);
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
      ? categoryRecords
      : categoryRecords.filter((record) => unitValueForSelect(record.unit) === selectedPlanUnit);

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
  const categoryLabelForMeta =
    selectedPlanCategory === EXPORT_CATEGORY_ALL_VALUE
      ? "全分類"
      : `${selectedPlanCategory}: ${SPECIMEN_CATEGORY_MAP[selectedPlanCategory] || ""}`;
  const unitLabelForMeta = selectedPlanUnit === ALL_UNITS_VALUE ? "全ユニット" : unitLabelForSelect(selectedPlanUnit);
  const detailLabelForMeta =
    selectedPlanDetail === ALL_DETAILS_VALUE ? "全サブユニット" : detailLabelForSelect(selectedPlanDetail);
  const detailSubLabelForMeta =
    selectedPlanDetailSub === ALL_DETAIL_SUBS_VALUE ? "全細分" : detailSubLabelForSelect(selectedPlanDetailSub);
  const hierarchyLabelForMeta = `${unitLabelForMeta} > ${detailLabelForMeta} > ${detailSubLabelForMeta}`;
  const mapMetaHtml = `
    <div class="plan-map-meta">
      <span>区画（グリッド）: ${escapeHtml(kuwakuLabelForMeta)}</span>
      <span>分類: ${escapeHtml(categoryLabelForMeta)}</span>
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
  const cornerLabelsSvg = buildPlanCornerLabelsSvg(cornerLabels);
  const viewBox = computePlanSvgViewBox(drawables);

  planMapWrap.innerHTML = `
    ${mapMetaHtml}
    <div class="plan-map-shell">
      <div class="plan-axis north">北</div>
      <div class="plan-axis east">東</div>
      <div class="plan-axis south">南</div>
      <div class="plan-axis west">西</div>
      <svg class="plan-map-svg" viewBox="${viewBox.minX} ${viewBox.minY} ${viewBox.width} ${viewBox.height}" aria-label="ユニット別平面図" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink">
        <rect class="plan-frame" x="0" y="0" width="${PLAN_SIZE_CM}" height="${PLAN_SIZE_CM}" />
        ${verticalGrid}
        ${horizontalGrid}
        ${cornerLabelsSvg}
        ${pointsSvg}
      </svg>
      <div class="plan-map-tooltip" hidden></div>
    </div>
  `;
  attachPlanMapTooltips();
}

function buildCurrentRecordDraftForPositionPreview() {
  if (!recordForm) {
    return null;
  }
  const formData = new FormData(recordForm);
  const isEditTab = getActiveTabId() === "edit-tab";
  const kuwaku = isEditTab
    ? buildKuwaku(
        normalizeKuwakuHeadA(editKuwakuHeadAInput?.value),
        normalizeKuwakuHeadB(editKuwakuHeadBInput?.value),
        normalizeKuwakuBlock(editKuwakuBlockInput?.value),
        normalizeKuwakuNo(editKuwakuNoInput?.value)
      )
    : buildKuwaku(
        normalizeKuwakuHeadA(siteForm?.elements?.kuwakuHeadA?.value),
        normalizeKuwakuHeadB(siteForm?.elements?.kuwakuHeadB?.value),
        normalizeKuwakuBlock(siteForm?.elements?.kuwakuBlock?.value),
        normalizeKuwakuNo(siteForm?.elements?.kuwakuNo?.value)
      );
  const specimenPrefix = normalizeSpecimenPrefix(value(formData.get("specimenPrefix")));
  const specimenSerial = compactNoSpaceValue(formData.get("specimenSerial"));
  const planSizeMode = normalizePlanSizeMode(value(formData.get("planSizeMode")));
  const draftMultiPoints = planSizeMode === "複数点" ? readMultiPointRowsFromForm() : [];
  const rawLargeShapeType = value(formData.get("largeShapeType"));
  const largeShapeType =
    planSizeMode === "大きなもの" ? normalizeLargeShapeType(rawLargeShapeType) || normalizeLargeShapeLabel(rawLargeShapeType) : "";
  const imageCornerFields = extractImageCornerFieldsFromFormData(formData);
  const imageTransformFields = extractImageTransformFieldsFromFormData(formData);
  const isImageShape = isLargeShapeImageType(largeShapeType);
  const isCustomImageShape = isImageShape && isCustomLargeShapeType(largeShapeType);
  return {
    id: value(editingRecordId || recordIdInput?.value),
    kuwaku,
    specimenPrefix,
    specimenSerial,
    specimenNo: buildSpecimenNo(specimenPrefix, specimenSerial),
    nameMemo: value(formData.get("nameMemo")),
    unit: compactNoSpaceValue(formData.get("unit")),
    detail: compactNoSpaceValue(formData.get("detail")),
    detailSub: value(formData.get("detailSub")),
    nsDir: normalizeNsDir(value(formData.get("nsDir"))),
    nsCm: value(formData.get("nsCm")),
    ewDir: normalizeEwDir(value(formData.get("ewDir"))),
    ewCm: value(formData.get("ewCm")),
    positionMethod: normalizePositionMethod(formData.get("positionMethod")),
    tsCoordinateConvention: "southWestPositive",
    tsStationPeg: normalizeTotalStationPointName(formData.get("tsStationPeg")),
    tsStationXNorthM: value(formData.get("tsStationXNorthM")),
    tsStationYEastM: value(formData.get("tsStationYEastM")),
    tsStationAltitudeM: value(formData.get("tsStationAltitudeM")),
    tsBacksightPeg: normalizeTotalStationPointName(formData.get("tsBacksightPeg")),
    tsBacksightXNorthM: value(formData.get("tsBacksightXNorthM")),
    tsBacksightYEastM: value(formData.get("tsBacksightYEastM")),
    tsBacksightAltitudeM: value(formData.get("tsBacksightAltitudeM")),
    tsInstrumentHeightM: value(formData.get("tsInstrumentHeightM")),
    tsTargetHeightM: value(formData.get("tsTargetHeightM")),
    tsObservationMode: value(formData.get("tsObservationMode")) === "polar" ? "polar" : "coordinate",
    tsPointCoordinateMode: "stationOffsetSouthWest",
    multiPoints: draftMultiPoints,
    planSizeMode,
    largeShapeType,
    largeAxisDirection: normalizeLargeAxisDirection(value(formData.get("largeAxisDirection"))),
    planeStrikeDirection: normalizePlaneStrikeDirection(value(formData.get("planeStrikeDirection"))),
    planeDipDeg: normalizePlaneDipDeg(value(formData.get("planeDipDeg"))),
    planeDipDir8: normalizeCompass8Direction(value(formData.get("planeDipDir8"))),
    lineLengthCm: value(formData.get("lineLengthCm")),
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
    imgRotateDeg: isImageShape ? imageTransformFields.imgRotateDeg : "",
    imgFrameWidthCm: isImageShape ? imageTransformFields.imgFrameWidthCm : "",
    imgFrameHeightCm: isImageShape ? imageTransformFields.imgFrameHeightCm : "",
    imgSkewXDeg: isImageShape ? imageTransformFields.imgSkewXDeg : "",
    imgSkewYDeg: isImageShape ? imageTransformFields.imgSkewYDeg : "",
    imgFlipH: isImageShape ? imageTransformFields.imgFlipH : "0",
    imgFlipV: isImageShape ? imageTransformFields.imgFlipV : "0",
    imgUseOriginalColor: isImageShape ? imageTransformFields.imgUseOriginalColor : "0",
    customLargeImageDataUrl: isCustomImageShape ? normalizeCustomLargeImageDataUrl(value(formData.get("customLargeImageDataUrl"))) : "",
  };
}

function renderPositionPreviewModalContent() {
  if (!positionPreviewMeta || !positionPreviewMap) {
    return false;
  }
  const draftRecord = positionPreviewRecordOverride || buildCurrentRecordDraftForPositionPreview();
  if (!draftRecord) {
    return false;
  }
  const currentDrawableRaw = buildPlanDrawable(draftRecord);
  if (!currentDrawableRaw) {
    const positionMethod = positionPreviewRecordOverride
      ? normalizePositionMethod(positionPreviewRecordOverride.positionMethod)
      : normalizePositionMethod(new FormData(recordForm).get("positionMethod"));
    showToast(
      positionMethod === "totalStation"
        ? getTotalStationInputError() || "トータルステーションの入力値を確認してください"
        : "平面位置の入力値を確認してください"
    );
    return false;
  }
  const kuwakuValue = kuwakuValueForSelect(getRecordKuwaku(draftRecord));
  const savedRecords = state.records.filter(
    (record) => kuwakuValueForSelect(getRecordKuwaku(record)) === kuwakuValue && value(record.id) !== value(draftRecord.id)
  );
  const savedDrawables = savedRecords
    .map((record) => {
      const drawable = buildPlanDrawable(record);
      return drawable ? { ...drawable, color: "#9ca3af", labelColor: "#6b7280" } : null;
    })
    .filter(Boolean);
  const currentDrawable = {
    ...currentDrawableRaw,
    color: "#dc2626",
    labelColor: "#dc2626",
    label: value(draftRecord.specimenNo) || "入力中",
  };
  const drawables = [...savedDrawables, currentDrawable];
  const verticalGrid = [100, 200, 300]
    .map((x) => `<line class="plan-grid-line" x1="${x}" y1="0" x2="${x}" y2="${PLAN_SIZE_CM}" />`)
    .join("");
  const horizontalGrid = [100, 200, 300]
    .map((y) => `<line class="plan-grid-line" x1="0" y1="${y}" x2="${PLAN_SIZE_CM}" y2="${y}" />`)
    .join("");
  const pointsSvg = drawables.map((drawable, index) => renderPlanDrawableSvg(drawable, index)).join("");
  const cornerLabels = buildPlanCornerLabels(kuwakuValue);
  const cornerLabelsSvg = buildPlanCornerLabelsSvg(cornerLabels);
  const viewBox = computePlanSvgViewBox(drawables);
  const kuwakuLabel = kuwakuValue === EMPTY_KUWAKU_VALUE ? "（未設定）" : kuwakuLabelForSelect(kuwakuValue);
  positionPreviewMeta.innerHTML = `
    <span>区画（グリッド）: ${escapeHtml(kuwakuLabel)}</span>
    <span>表示件数: ${drawables.length}件</span>
  `;
  positionPreviewMap.innerHTML = `
    <div class="plan-map-shell position-preview-shell">
      <div class="plan-axis north">北</div>
      <div class="plan-axis east">東</div>
      <div class="plan-axis south">南</div>
      <div class="plan-axis west">西</div>
      <svg class="plan-map-svg" viewBox="${viewBox.minX} ${viewBox.minY} ${viewBox.width} ${viewBox.height}" aria-label="平面位置確認" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink">
        <rect class="plan-frame" x="0" y="0" width="${PLAN_SIZE_CM}" height="${PLAN_SIZE_CM}" />
        ${verticalGrid}
        ${horizontalGrid}
        ${cornerLabelsSvg}
        ${pointsSvg}
      </svg>
    </div>
  `;
  return true;
}

function openPositionPreviewModal(recordOverride = null) {
  if (!positionPreviewModal) {
    return;
  }
  positionPreviewRecordOverride = recordOverride && typeof recordOverride === "object" ? recordOverride : null;
  const rendered = renderPositionPreviewModalContent();
  if (!rendered) {
    positionPreviewRecordOverride = null;
    return;
  }
  positionPreviewModal.classList.remove("hidden");
}

function closePositionPreviewModal() {
  if (!positionPreviewModal) {
    return;
  }
  positionPreviewModal.classList.add("hidden");
  positionPreviewRecordOverride = null;
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

function hideViewerFallbackPanel() {
  if (!viewerCanvasWrap) {
    return;
  }
  const panel = viewerCanvasWrap.querySelector(".viewer-fallback-panel");
  if (panel) {
    panel.remove();
  }
}

function showViewerFallbackPanel(reasonText, context = {}) {
  if (!viewerCanvasWrap) {
    return;
  }
  const reason = value(reasonText) || "この端末では3D表示が利用できません。";
  const kuwakuLabel = value(context.kuwakuLabel) || "-";
  const totalCount = Number.isFinite(Number(context.totalCount)) ? Number(context.totalCount) : 0;
  const drawableCount = Number.isFinite(Number(context.drawableCount)) ? Number(context.drawableCount) : 0;
  let panel = viewerCanvasWrap.querySelector(".viewer-fallback-panel");
  if (!panel) {
    panel = document.createElement("div");
    panel.className = "viewer-fallback-panel";
    viewerCanvasWrap.appendChild(panel);
  }
  panel.innerHTML = `
    <div class="viewer-fallback-card">
      <p class="viewer-fallback-title">この端末では3D表示が使えません</p>
      <p class="viewer-fallback-reason">${escapeHtml(reason)}</p>
      <p class="viewer-fallback-meta">区画: ${escapeHtml(kuwakuLabel)} / 対象 ${totalCount}件 / 描画可能 ${drawableCount}件</p>
      <button type="button" class="viewer-fallback-open-plan-btn" data-action="viewer-open-plan">平面図タブで確認</button>
    </div>
  `;
}

function renderViewerOutput(options = {}) {
  const preserveCamera = Boolean(options?.preserveCamera);
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
  hideViewerFallbackPanel();
  viewerMapLegend.innerHTML = buildPlanLegendHtml();
  syncViewerVerticalScaleUi();
  syncViewerViewButtons();

  const sortedRecords = [...state.records].sort(compareRecordsByKuwakuThenSpecimen);
  syncViewerKuwakuSelect(sortedRecords);

  if (!sortedRecords.length) {
    selectedViewerCategory = EXPORT_CATEGORY_ALL_VALUE;
    selectedViewerUnit = ALL_UNITS_VALUE;
    selectedViewerDetail = ALL_DETAILS_VALUE;
    selectedViewerDetailSub = ALL_DETAIL_SUBS_VALUE;
    if (viewerCategorySelect) {
      viewerCategorySelect.innerHTML = "";
    }
    viewerUnitSelect.innerHTML = "";
    viewerDetailSelect.innerHTML = "";
    viewerDetailSubSelect.innerHTML = "";
    if (viewerKuwakuInfo) {
      viewerKuwakuInfo.textContent = "区画: -";
    }
    viewerStatus.textContent = "表示対象データがありません。";
    clearViewerScene();
    hideViewerFallbackPanel();
    return;
  }

  const kuwakuRecords =
    selectedViewerKuwaku === ALL_GRIDS_VALUE
      ? sortedRecords
      : sortedRecords.filter((record) => kuwakuValueForSelect(getRecordKuwaku(record)) === selectedViewerKuwaku);
  syncViewerCategorySelect(kuwakuRecords);
  const categoryRecords = filterRecordsByCategory(kuwakuRecords, selectedViewerCategory);

  const kuwakuLabel = selectedViewerKuwaku === ALL_GRIDS_VALUE ? "全グリッド" : kuwakuLabelForSelect(selectedViewerKuwaku);
  if (viewerKuwakuInfo) {
    viewerKuwakuInfo.textContent = `区画: ${kuwakuLabel}`;
  }

  if (!categoryRecords.length) {
    selectedViewerUnit = ALL_UNITS_VALUE;
    selectedViewerDetail = ALL_DETAILS_VALUE;
    selectedViewerDetailSub = ALL_DETAIL_SUBS_VALUE;
    viewerUnitSelect.innerHTML = "";
    viewerDetailSelect.innerHTML = "";
    viewerDetailSubSelect.innerHTML = "";
    viewerStatus.textContent = "この条件では表示対象データがありません。";
    clearViewerScene();
    hideViewerFallbackPanel();
    return;
  }

  const units = collectPlanUnits(categoryRecords);
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
      ? categoryRecords
      : categoryRecords.filter((record) => unitValueForSelect(record.unit) === selectedViewerUnit);

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
    hideViewerFallbackPanel();
    return;
  }

  if (!isViewerTabActive() && !viewer3d.initialized) {
    return;
  }
  if (!ensureViewerInitialized()) {
    showViewerFallbackPanel(viewerStatus?.textContent, {
      kuwakuLabel,
      totalCount: detailSubRecords.length,
      drawableCount: viewerCandidates.length,
    });
    return;
  }
  hideViewerFallbackPanel();

  const metrics = buildViewerGridMetrics(viewerCandidates);
  const shapes = viewerCandidates.map((candidate) => buildViewerShapeFromCandidate(candidate, metrics)).filter(Boolean);
  renderViewerScene(shapes, metrics, { preserveCamera });
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

function normalizeAzimuth360(valueRaw) {
  const valueNum = Number(valueRaw);
  if (!Number.isFinite(valueNum)) {
    return null;
  }
  return ((valueNum % 360) + 360) % 360;
}

function angularDistanceDeg(aRaw, bRaw) {
  const a = normalizeAzimuth360(aRaw);
  const b = normalizeAzimuth360(bRaw);
  if (!Number.isFinite(a) || !Number.isFinite(b)) {
    return Infinity;
  }
  const diff = Math.abs(a - b) % 360;
  return diff > 180 ? 360 - diff : diff;
}

function azimuthToViewerHorizontalUnit(azimuthRaw) {
  const azimuth = Number(azimuthRaw);
  if (!Number.isFinite(azimuth)) {
    return { x: 0, y: 1 };
  }
  const rad = (azimuth * Math.PI) / 180;
  return {
    x: Math.sin(rad),
    y: Math.cos(rad),
  };
}

function resolvePlaneTiltParameters(record) {
  const strikeAzimuth = parseLargeAxisAzimuth(record?.planeStrikeDirection);
  const dipDeg = parseLargeAxisPlungeDeg(record?.planeDipDeg);
  const dipDirRaw = parseCompass8Azimuth(record?.planeDipDir8);
  if (!Number.isFinite(dipDeg) || dipDeg <= 0) {
    return {
      strikeAzimuth: Number.isFinite(strikeAzimuth) ? strikeAzimuth : null,
      dipAzimuth: null,
      dipDeg: null,
    };
  }
  if (Number.isFinite(strikeAzimuth)) {
    const rightDip = (strikeAzimuth + 90) % 360;
    const leftDip = (strikeAzimuth + 270) % 360;
    if (!Number.isFinite(dipDirRaw)) {
      return { strikeAzimuth, dipAzimuth: rightDip, dipDeg };
    }
    const rightDist = angularDistanceDeg(dipDirRaw, rightDip);
    const leftDist = angularDistanceDeg(dipDirRaw, leftDip);
    return {
      strikeAzimuth,
      dipAzimuth: rightDist <= leftDist ? rightDip : leftDip,
      dipDeg,
    };
  }
  return {
    strikeAzimuth: null,
    dipAzimuth: Number.isFinite(dipDirRaw) ? dipDirRaw : null,
    dipDeg,
  };
}

function buildViewerShapeFromCandidate(candidate, metrics) {
  const drawable = candidate.drawable;
  const altitudeM = candidate.altitudeM;
  const centerPlanPoint = { x: drawable.x, y: drawable.y };
  let tiltAzimuth = null;
  let tiltDeg = null;
  let lineDirectionAzimuth = null;
  let linePlungeDeg = null;
  let lineDownwardAzimuth = null;
  let planeStrikeAzimuth = null;
  let planeDipAzimuth = null;
  let planeDipDeg = null;
  if (drawable.type === "line") {
    lineDirectionAzimuth = parseLargeAxisAzimuth(candidate?.record?.largeAxisDirection);
    linePlungeDeg = parseLargeAxisPlungeDeg(candidate?.record?.largeAxisPlungeDeg);
    lineDownwardAzimuth = parseCompass8Azimuth(candidate?.record?.largeAxisPlungeDir8);
    tiltAzimuth = lineDownwardAzimuth ?? lineDirectionAzimuth;
    tiltDeg = linePlungeDeg;
  } else if (drawable.type === "rect" || drawable.type === "ellipse" || drawable.type === "imageQuad") {
    const planeTilt = resolvePlaneTiltParameters(candidate?.record);
    planeStrikeAzimuth = planeTilt.strikeAzimuth;
    planeDipAzimuth = planeTilt.dipAzimuth;
    planeDipDeg = planeTilt.dipDeg;
    tiltAzimuth = planeDipAzimuth;
    tiltDeg = planeDipDeg;
  }
  const getViewerZForPlanPoint = (planPointRaw) => {
    const planPoint = planPointRaw || centerPlanPoint;
    const deltaM = computeViewerPlungeDeltaM(planPoint, centerPlanPoint, tiltAzimuth, tiltDeg);
    return applyViewerVerticalScale(altitudeM + deltaM, metrics.minZ);
  };
  const getViewerZForImagePlanPoint = (planPointRaw) => {
    const planPoint = planPointRaw || centerPlanPoint;
    const rawDeltaM = computeViewerPlungeDeltaM(planPoint, centerPlanPoint, tiltAzimuth, tiltDeg);
    // 画像形状のフォールバック計算。極端値は抑制して表示破綻を防ぐ。
    const deltaM = Number.isFinite(rawDeltaM)
      ? clamp(rawDeltaM * IMAGE_QUAD_TILT_Z_SCALE, -IMAGE_QUAD_TILT_Z_LIMIT_M, IMAGE_QUAD_TILT_Z_LIMIT_M)
      : 0;
    const baseZ = applyViewerVerticalScale(altitudeM, metrics.minZ);
    return baseZ + deltaM;
  };
  const altitudeZ = getViewerZForPlanPoint(centerPlanPoint);
  const directBottomAltitudeEnabled = normalizeToggleFlag(candidate?.record?.altitudeInputEnabled) === "1";
  const bottomTargetZ = applyViewerVerticalScale(altitudeM, metrics.minZ);
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

  if (drawable.type === "multipoint") {
    let points = (drawable.points || [])
      .map((point) => {
        const world = convertViewerPointCmToWorld(point.x, point.y, candidate.grid, metrics);
        return { x: world.x, y: world.y, z: getViewerZForPlanPoint(point) };
      })
      .filter((point) => Number.isFinite(point.x) && Number.isFinite(point.y) && Number.isFinite(point.z));
    if (!points.length) {
      return null;
    }
    if (directBottomAltitudeEnabled) {
      const anchored = anchorViewerPointsToBottomAltitude(points, bottomTargetZ);
      points = anchored.points;
    }
    const hull = buildHullPointsFromSource(points, points[0]?.z);
    const centerZ = points.reduce((sum, point) => sum + point.z, 0) / points.length;
    return {
      type: "multipoint",
      points,
      hull,
      x: worldCenter.x,
      y: worldCenter.y,
      z: centerZ,
      ...meta,
    };
  }

  if (drawable.type === "line") {
    const viewerScale = normalizeViewerVerticalScale(viewerVerticalScale);
    const lineLengthCm =
      parseDistanceToCm(candidate?.record?.lineLengthCm) ?? Math.hypot(drawable.x2 - drawable.x1, drawable.y2 - drawable.y1);
    const halfLengthM = Math.max(0, lineLengthCm / 200);
    let axisAzimuth = normalizeAzimuth360(lineDirectionAzimuth ?? tiltAzimuth);
    if (!Number.isFinite(axisAzimuth)) {
      axisAzimuth = 0;
    }
    const downwardAzimuth = normalizeAzimuth360(lineDownwardAzimuth);
    if (Number.isFinite(downwardAzimuth)) {
      const forwardDist = angularDistanceDeg(downwardAzimuth, axisAzimuth);
      if (forwardDist > 90) {
        axisAzimuth = (axisAzimuth + 180) % 360;
      }
    }
    const plunge = Number.isFinite(linePlungeDeg) ? clamp(linePlungeDeg, 0, 90) : 0;
    const plungeRad = (plunge * Math.PI) / 180;
    const horizFactor = Math.cos(plungeRad);
    const verticalFactor = Math.sin(plungeRad);
    const axisUnit = azimuthToViewerHorizontalUnit(axisAzimuth);
    const vx = axisUnit.x * horizFactor;
    const vy = axisUnit.y * horizFactor;
    const vz = -verticalFactor * viewerScale;
    let linePoints = [
      {
        x: worldCenter.x - vx * halfLengthM,
        y: worldCenter.y - vy * halfLengthM,
        z: altitudeZ - vz * halfLengthM,
      },
      {
        x: worldCenter.x + vx * halfLengthM,
        y: worldCenter.y + vy * halfLengthM,
        z: altitudeZ + vz * halfLengthM,
      },
    ];
    let lineCenterZ = altitudeZ;
    if (directBottomAltitudeEnabled) {
      const anchored = anchorViewerPointsToBottomAltitude(linePoints, bottomTargetZ);
      linePoints = anchored.points;
      if (anchored.minZ != null && anchored.maxZ != null) {
        lineCenterZ = (anchored.minZ + anchored.maxZ) / 2;
      }
    }
    return {
      type: "line",
      points: linePoints,
      x: worldCenter.x,
      y: worldCenter.y,
      z: lineCenterZ,
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
    let points = localCorners.map((point) => {
      const world = convertViewerPointCmToWorld(point.x, point.y, candidate.grid, metrics);
      return { x: world.x, y: world.y, z: getViewerZForPlanPoint(point) };
    });
    if (directBottomAltitudeEnabled) {
      const anchored = anchorViewerPointsToBottomAltitude(points, bottomTargetZ);
      points = anchored.points;
    }
    points.push(points[0]);
    return {
      type: "polyline",
      points,
      x: worldCenter.x,
      y: worldCenter.y,
      z: directBottomAltitudeEnabled && points.length ? points.reduce((sum, p) => sum + p.z, 0) / points.length : altitudeZ,
      ...meta,
    };
  }

  if (drawable.type === "ellipse") {
    let points = [];
    const segmentCount = 48;
    for (let i = 0; i <= segmentCount; i += 1) {
      const theta = (i / segmentCount) * Math.PI * 2;
      const local = {
        x: drawable.x + Math.cos(theta) * drawable.rx,
        y: drawable.y + Math.sin(theta) * drawable.ry,
      };
      const rotated = rotatePlanPoint(local, { x: drawable.x, y: drawable.y }, drawable.rotationDeg);
      const world = convertViewerPointCmToWorld(rotated.x, rotated.y, candidate.grid, metrics);
      points.push({ x: world.x, y: world.y, z: getViewerZForPlanPoint(rotated) });
    }
    if (directBottomAltitudeEnabled) {
      const anchored = anchorViewerPointsToBottomAltitude(points, bottomTargetZ);
      points = anchored.points;
    }
    return {
      type: "polyline",
      points,
      x: worldCenter.x,
      y: worldCenter.y,
      z: directBottomAltitudeEnabled && points.length ? points.reduce((sum, p) => sum + p.z, 0) / points.length : altitudeZ,
      ...meta,
    };
  }

  if (drawable.type === "imageQuad") {
    const viewerScale = normalizeViewerVerticalScale(viewerVerticalScale);
    const dipDeg = Number.isFinite(planeDipDeg) ? clamp(planeDipDeg, 0, 90) : 0;
    const dipAzimuth = normalizeAzimuth360(
      planeDipAzimuth ??
        (Number.isFinite(planeStrikeAzimuth) ? (planeStrikeAzimuth + 90) % 360 : null)
    );
    const strikeAzimuth = normalizeAzimuth360(
      planeStrikeAzimuth ??
        (Number.isFinite(dipAzimuth) ? (dipAzimuth + 270) % 360 : null)
    );
    const canUsePlaneProjection = Number.isFinite(dipAzimuth) && Number.isFinite(strikeAzimuth) && dipDeg > 0;
    const strikeUnit = azimuthToViewerHorizontalUnit(strikeAzimuth);
    const dipUnit = azimuthToViewerHorizontalUnit(dipAzimuth);
    const dipRad = (dipDeg * Math.PI) / 180;
    const dipHorizFactor = Math.cos(dipRad);
    const dipVerticalFactor = Math.sin(dipRad);
    let points = (drawable.points || []).map((point) => {
      const world = convertViewerPointCmToWorld(point.x, point.y, candidate.grid, metrics);
      if (!canUsePlaneProjection) {
        return { x: world.x, y: world.y, z: getViewerZForImagePlanPoint(point) };
      }
      const relX = world.x - worldCenter.x;
      const relY = world.y - worldCenter.y;
      const strikeComp = relX * strikeUnit.x + relY * strikeUnit.y;
      const dipComp = relX * dipUnit.x + relY * dipUnit.y;
      return {
        x: worldCenter.x + strikeUnit.x * strikeComp + dipUnit.x * dipComp * dipHorizFactor,
        y: worldCenter.y + strikeUnit.y * strikeComp + dipUnit.y * dipComp * dipHorizFactor,
        z: altitudeZ - dipComp * dipVerticalFactor * viewerScale,
      };
    });
    if (directBottomAltitudeEnabled) {
      const anchored = anchorViewerPointsToBottomAltitude(points, bottomTargetZ);
      points = anchored.points;
    }
    return {
      type: "imageQuad",
      points,
      imageType: value(drawable.imageType),
      imagePath: value(drawable.imagePath),
      useOriginalImageColor: Boolean(drawable.useOriginalImageColor),
      x: worldCenter.x,
      y: worldCenter.y,
      z: directBottomAltitudeEnabled && points.length ? points.reduce((sum, p) => sum + p.z, 0) / points.length : altitudeZ,
      ...meta,
    };
  }
  return null;
}

function anchorViewerPointsToBottomAltitude(pointsRaw, targetBottomZRaw) {
  const points = Array.isArray(pointsRaw) ? pointsRaw : [];
  const targetBottomZ = Number(targetBottomZRaw);
  if (!points.length || !Number.isFinite(targetBottomZ)) {
    return { points, minZ: null, maxZ: null };
  }
  let minZ = Infinity;
  let maxZ = -Infinity;
  points.forEach((point) => {
    const z = Number(point?.z);
    if (!Number.isFinite(z)) {
      return;
    }
    if (z < minZ) {
      minZ = z;
    }
    if (z > maxZ) {
      maxZ = z;
    }
  });
  if (!Number.isFinite(minZ) || !Number.isFinite(maxZ)) {
    return { points, minZ: null, maxZ: null };
  }
  const deltaZ = targetBottomZ - minZ;
  if (Math.abs(deltaZ) < 1e-9) {
    return { points, minZ, maxZ };
  }
  const shifted = points.map((point) => ({
    x: point.x,
    y: point.y,
    z: Number.isFinite(Number(point?.z)) ? Number(point.z) + deltaZ : point.z,
  }));
  return {
    points: shifted,
    minZ: minZ + deltaZ,
    maxZ: maxZ + deltaZ,
  };
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

function computeViewerPlungeDeltaM(point, center, axisAzimuthRaw, plungeDegRaw) {
  if (!point || !center) {
    return 0;
  }
  const axisAzimuth = Number(axisAzimuthRaw);
  const plungeDeg = Number(plungeDegRaw);
  if (!Number.isFinite(axisAzimuth) || !Number.isFinite(plungeDeg) || plungeDeg <= 0) {
    return 0;
  }
  const unit = azimuthToPlanUnitVector(axisAzimuth);
  const alongAxisCm = (point.x - center.x) * unit.dx + (point.y - center.y) * unit.dy;
  const alongAxisM = alongAxisCm / 100;
  const tilt = Math.tan((plungeDeg * Math.PI) / 180);
  if (!Number.isFinite(tilt) || tilt === 0) {
    return 0;
  }
  // 方位方向側へ進むほど深くなる（標高は低くなる）向きで適用する。
  return -alongAxisM * tilt;
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
  if (viewerViewSeBtn) {
    viewerViewSeBtn.classList.toggle("active", selectedViewerPerspective === "se" || selectedViewerPerspective === "iso");
  }
  if (viewerViewEastBtn) {
    viewerViewEastBtn.classList.toggle("active", selectedViewerPerspective === "east");
  }
  if (viewerViewWestBtn) {
    viewerViewWestBtn.classList.toggle("active", selectedViewerPerspective === "west");
  }
  if (viewerViewSouthBtn) {
    viewerViewSouthBtn.classList.toggle("active", selectedViewerPerspective === "south");
  }
  if (viewerViewNorthBtn) {
    viewerViewNorthBtn.classList.toggle("active", selectedViewerPerspective === "north");
  }
}

function createViewerRendererWithFallback() {
  if (!window.THREE) {
    return { renderer: null, error: new Error("THREE未読込") };
  }
  const createRendererWithContext = (contextType, contextAttrs) => {
    const canvas = document.createElement("canvas");
    const gl =
      canvas.getContext(contextType, contextAttrs) ||
      (contextType === "webgl" ? canvas.getContext("experimental-webgl", contextAttrs) : null);
    if (!gl) {
      throw new Error(`${contextType} context unavailable`);
    }
    return new THREE.WebGLRenderer({
      canvas,
      context: gl,
      antialias: false,
      alpha: Boolean(contextAttrs?.alpha),
      powerPreference: "low-power",
      precision: "mediump",
    });
  };
  const attempts = [
    () =>
      new THREE.WebGLRenderer({
        antialias: false,
        alpha: false,
        stencil: false,
        powerPreference: "low-power",
        precision: "mediump",
      }),
    () => createRendererWithContext("webgl2", { antialias: false, alpha: false, depth: true, stencil: false }),
    () => createRendererWithContext("webgl", { antialias: false, alpha: false, depth: true, stencil: false }),
    () => createRendererWithContext("webgl", { antialias: false, alpha: true, depth: true, stencil: false }),
  ];
  if (typeof THREE.WebGL1Renderer === "function") {
    attempts.push(() => new THREE.WebGL1Renderer({ antialias: false, alpha: false, stencil: false }));
  }
  let lastError = null;
  for (const makeRenderer of attempts) {
    try {
      const renderer = makeRenderer();
      if (renderer) {
        return { renderer, error: null };
      }
    } catch (error) {
      lastError = error;
    }
  }
  const probeCanvas = document.createElement("canvas");
  const hasWebGl =
    Boolean(probeCanvas.getContext("webgl")) ||
    Boolean(probeCanvas.getContext("experimental-webgl")) ||
    Boolean(probeCanvas.getContext("webgl2"));
  if (!hasWebGl) {
    lastError = new Error("この端末・ブラウザでWebGLが利用できません");
  }
  return { renderer: null, error: lastError || new Error("WebGLRenderer生成失敗") };
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
    viewer3d.camera = new THREE.PerspectiveCamera(34, 1, 0.1, 10000);
    const rendererResult = createViewerRendererWithFallback();
    if (!rendererResult.renderer) {
      throw rendererResult.error || new Error("WebGLRenderer生成失敗");
    }
    viewer3d.renderer = rendererResult.renderer;
    viewer3d.renderer.setPixelRatio(Math.min(window.devicePixelRatio || 1, 2));
    viewer3d.renderer.domElement.setAttribute("aria-label", "3Dビュアー");
    if (typeof viewerCanvasWrap.prepend === "function") {
      viewerCanvasWrap.prepend(viewer3d.renderer.domElement);
    } else {
      viewerCanvasWrap.insertBefore(viewer3d.renderer.domElement, viewerCanvasWrap.firstChild);
    }

    viewer3d.controls =
      typeof THREE.OrbitControls === "function"
        ? new THREE.OrbitControls(viewer3d.camera, viewer3d.renderer.domElement)
        : null;
    if (viewer3d.controls) {
      viewer3d.controls.enableDamping = false;
      viewer3d.controls.enablePan = true;
      viewer3d.controls.rotateSpeed = 1.25;
      viewer3d.controls.zoomSpeed = 0.62;
      viewer3d.controls.panSpeed = 1.1;
      viewer3d.controls.screenSpacePanning = true;
      if (THREE.MOUSE) {
        viewer3d.controls.mouseButtons = {
          LEFT: THREE.MOUSE.ROTATE,
          MIDDLE: THREE.MOUSE.PAN,
          RIGHT: null,
        };
        viewer3d.defaultLeftMouseAction = THREE.MOUSE.ROTATE;
      }
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
    viewerCanvasWrap.addEventListener("pointerdown", handleViewerPointerDown);
    viewerCanvasWrap.addEventListener("pointerup", handleViewerPointerEnd);
    viewerCanvasWrap.addEventListener("pointercancel", handleViewerPointerEnd);
    viewerCanvasWrap.addEventListener("pointerleave", handleViewerPointerLeave);
    viewerCanvasWrap.addEventListener("contextmenu", handleViewerContextMenu);
    if (viewer3d.renderer?.domElement) {
      viewer3d.renderer.domElement.addEventListener("pointerdown", handleViewerControlPointerDown);
      viewer3d.renderer.domElement.addEventListener("pointerup", handleViewerControlPointerUp);
      viewer3d.renderer.domElement.addEventListener("pointercancel", handleViewerControlPointerUp);
    }
    window.addEventListener("pointerup", handleViewerControlPointerUp);
    window.addEventListener("blur", handleViewerControlPointerUp);
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
  } catch (error) {
    viewer3d.available = false;
    viewer3d.initialized = false;
    if (viewer3d.renderer) {
      try {
        viewer3d.renderer.dispose();
      } catch (_disposeError) {
        // noop
      }
      const canvas = viewer3d.renderer.domElement;
      if (canvas?.parentNode) {
        canvas.parentNode.removeChild(canvas);
      }
      viewer3d.renderer = null;
    }
    if (viewerStatus) {
      const reason = value(error?.message);
      viewerStatus.textContent = reason
        ? `3D表示の初期化に失敗しました（${reason}）。`
        : "3D表示の初期化に失敗しました。";
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

function captureViewerCameraState() {
  if (!viewer3d.initialized || !viewer3d.camera || !viewer3d.bounds) {
    return null;
  }
  const camera = viewer3d.camera;
  const position = {
    x: Number(camera.position?.x),
    y: Number(camera.position?.y),
    z: Number(camera.position?.z),
  };
  if (!Number.isFinite(position.x) || !Number.isFinite(position.y) || !Number.isFinite(position.z)) {
    return null;
  }
  const up = {
    x: Number(camera.up?.x),
    y: Number(camera.up?.y),
    z: Number(camera.up?.z),
  };
  const targetVec = viewer3d.controls?.target;
  const target = {
    x: Number(targetVec?.x),
    y: Number(targetVec?.y),
    z: Number(targetVec?.z),
  };
  if (!Number.isFinite(target.x) || !Number.isFinite(target.y) || !Number.isFinite(target.z)) {
    return null;
  }
  const dx = position.x - target.x;
  const dy = position.y - target.y;
  const dz = position.z - target.z;
  const distance = Math.hypot(dx, dy, dz);
  if (!Number.isFinite(distance) || distance < 0.5) {
    return null;
  }
  return {
    position,
    up:
      Number.isFinite(up.x) && Number.isFinite(up.y) && Number.isFinite(up.z)
        ? up
        : { x: 0, y: 0, z: 1 },
    target,
  };
}

function restoreViewerCameraState(viewState) {
  if (!viewState || !viewer3d.initialized || !viewer3d.camera) {
    return false;
  }
  const position = viewState.position || {};
  const up = viewState.up || {};
  const target = viewState.target || {};
  if (
    !Number.isFinite(position.x) ||
    !Number.isFinite(position.y) ||
    !Number.isFinite(position.z) ||
    !Number.isFinite(target.x) ||
    !Number.isFinite(target.y) ||
    !Number.isFinite(target.z)
  ) {
    return false;
  }
  viewer3d.camera.position.set(position.x, position.y, position.z);
  if (Number.isFinite(up.x) && Number.isFinite(up.y) && Number.isFinite(up.z)) {
    viewer3d.camera.up.set(up.x, up.y, up.z);
  }
  if (viewer3d.controls) {
    viewer3d.controls.target.set(target.x, target.y, target.z);
    viewer3d.controls.update();
  } else {
    viewer3d.camera.lookAt(new THREE.Vector3(target.x, target.y, target.z));
  }
  return true;
}

function isViewerCameraStateCompatible(viewState, bounds) {
  if (!viewState || !bounds) {
    return false;
  }
  const target = viewState.target || {};
  const position = viewState.position || {};
  if (
    !Number.isFinite(target.x) ||
    !Number.isFinite(target.y) ||
    !Number.isFinite(target.z) ||
    !Number.isFinite(position.x) ||
    !Number.isFinite(position.y) ||
    !Number.isFinite(position.z)
  ) {
    return false;
  }
  const spanX = Math.max(1, Number(bounds.maxX) - Number(bounds.minX));
  const spanY = Math.max(1, Number(bounds.maxY) - Number(bounds.minY));
  const spanZ = Math.max(1, Number(bounds.maxZ) - Number(bounds.minZ));
  const marginFactor = 3.5;
  const within = (valueNum, minNum, maxNum, span) =>
    valueNum >= minNum - span * marginFactor && valueNum <= maxNum + span * marginFactor;
  if (
    !within(target.x, bounds.minX, bounds.maxX, spanX) ||
    !within(target.y, bounds.minY, bounds.maxY, spanY) ||
    !within(target.z, bounds.minZ, bounds.maxZ, spanZ)
  ) {
    return false;
  }
  if (
    !within(position.x, bounds.minX, bounds.maxX, spanX) ||
    !within(position.y, bounds.minY, bounds.maxY, spanY) ||
    !within(position.z, bounds.minZ, bounds.maxZ, spanZ)
  ) {
    return false;
  }
  return true;
}

function renderViewerScene(shapes, metrics, options = {}) {
  if (!viewer3d.initialized || !viewer3d.dataGroup || !viewer3d.labelGroup || !viewer3d.gridGroup) {
    return;
  }
  const preserveCamera = Boolean(options?.preserveCamera);
  const previousViewState = preserveCamera ? captureViewerCameraState() : null;
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
    } else if (shape.type === "multipoint") {
      (shape.points || []).forEach((point) => {
        const pointMesh = new THREE.Mesh(
          new THREE.SphereGeometry(0.085, 12, 12),
          new THREE.MeshBasicMaterial({ color })
        );
        pointMesh.position.set(point.x, point.y, point.z);
        viewer3d.dataGroup.add(pointMesh);
      });
      const hull = Array.isArray(shape.hull) ? shape.hull : [];
      if (hull.length === 2) {
        renderViewerSegment(hull[0], hull[1], color, 0.025);
      } else if (hull.length >= 3) {
        for (let i = 0; i < hull.length; i += 1) {
          const start = hull[i];
          const end = hull[(i + 1) % hull.length];
          renderViewerSegment(start, end, color, 0.025);
        }
      }
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

    const pickTargets =
      shape.type === "multipoint" && Array.isArray(shape.points) && shape.points.length
        ? shape.points
        : [{ x: shape.x, y: shape.y, z: shape.z }];
    pickTargets.forEach((targetPoint) => {
      const pickMesh = new THREE.Mesh(
        new THREE.SphereGeometry(0.2, 10, 10),
        new THREE.MeshBasicMaterial({ transparent: true, opacity: 0.001, depthWrite: false })
      );
      pickMesh.position.set(targetPoint.x, targetPoint.y, targetPoint.z);
      pickMesh.userData = {
        id: shape.id,
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
    });
    viewer3d.meshesByRecordId.set(shape.id, shape);
  });

  viewer3d.bounds = computeViewerBounds(shapes, metrics);
  const shouldRestoreCamera =
    preserveCamera && isViewerCameraStateCompatible(previousViewState, viewer3d.bounds) && restoreViewerCameraState(previousViewState);
  if (!shouldRestoreCamera) {
    applyViewerPerspective();
  }
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

  const cardinalOffset = 1.18;
  const cardinalZ = zBase + 0.06;
  const cardinalLabels = [
    { text: "北", x: width / 2, y: height + cardinalOffset },
    { text: "南", x: width / 2, y: -cardinalOffset },
    { text: "東", x: width + cardinalOffset, y: height / 2 },
    { text: "西", x: -cardinalOffset, y: height / 2 },
  ];
  cardinalLabels.forEach((item) => {
    const sprite = createViewerTextSprite(item.text, "#0f172a", {
      fontPx: 120,
      scaleX: 1.8,
      scaleY: 0.45,
    });
    sprite.position.set(item.x, item.y, cardinalZ);
    viewer3d.gridGroup.add(sprite);
  });

  const colNorthY = height + 0.62;
  const colSouthY = -0.62;
  const rowWestX = -0.62;
  const rowEastX = width + 0.62;
  const colCount = Math.max(1, metrics.maxXIndex - metrics.minXIndex + 1);
  const rowCount = Math.max(1, metrics.maxNo - metrics.minNo + 1);
  const denseFactor = Math.max(colCount, rowCount);
  const frameLabelWidth = clamp(2.002 - denseFactor * 0.0208, 1.17, 1.664);
  const frameLabelHeight = clamp(0.598 - denseFactor * 0.00624, 0.312, 0.468);
  const stakeLabelWidth = clamp(1.3824 - denseFactor * 0.01296, 0.72, 1.0944);
  const stakeLabelHeight = clamp(0.432 - denseFactor * 0.004032, 0.2016, 0.3168);
  const labelPlaneZ = zBase + 0.004;
  for (let xIndex = metrics.minXIndex; xIndex <= metrics.maxXIndex; xIndex += 1) {
    const block = viewerXIndexToHeadBlock(xIndex).block;
    const centerX = (xIndex - metrics.minXIndex) * 4 + 2;
    const northLabel = createViewerTextPlane(block, "#1f3b35", {
      fontPx: 127,
      width: frameLabelWidth,
      height: frameLabelHeight,
    });
    northLabel.position.set(centerX, colNorthY, labelPlaneZ);
    viewer3d.gridGroup.add(northLabel);
    const southLabel = createViewerTextPlane(block, "#1f3b35", {
      fontPx: 127,
      width: frameLabelWidth,
      height: frameLabelHeight,
    });
    southLabel.position.set(centerX, colSouthY, labelPlaneZ);
    viewer3d.gridGroup.add(southLabel);
  }
  for (let no = metrics.minNo; no <= metrics.maxNo; no += 1) {
    const rowText = String(no);
    const centerY = (metrics.maxNo - no) * 4 + 2;
    const westLabel = createViewerTextPlane(rowText, "#1f3b35", {
      fontPx: 127,
      width: frameLabelWidth,
      height: frameLabelHeight,
    });
    westLabel.position.set(rowWestX, centerY, labelPlaneZ);
    viewer3d.gridGroup.add(westLabel);
    const eastLabel = createViewerTextPlane(rowText, "#1f3b35", {
      fontPx: 127,
      width: frameLabelWidth,
      height: frameLabelHeight,
    });
    eastLabel.position.set(rowEastX, centerY, labelPlaneZ);
    viewer3d.gridGroup.add(eastLabel);
  }
  const allCornerStakes = buildViewerAllGridCornerStakeLabels(metrics);
  allCornerStakes.forEach((corner) => {
    const label = createViewerTextPlane(corner.label, "#97a7bc", {
      fontPx: 98,
      width: stakeLabelWidth,
      height: stakeLabelHeight,
    });
    label.position.set(corner.x, corner.y, labelPlaneZ);
    viewer3d.gridGroup.add(label);
  });

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
      label.position.set(corner.x + corner.dirX * 0.95, corner.y + corner.dirY * 0.95, z + 0.12);
      viewer3d.gridGroup.add(label);
    }
    const axisLabel = createViewerTextSprite("標高(m)", "#1e293b", {
      fontPx: 87,
      scaleX: 1.57,
      scaleY: 0.39,
    });
    axisLabel.position.set(corner.x + corner.dirX * 1.5, corner.y + corner.dirY * 1.5, zEnd + 0.85);
    viewer3d.gridGroup.add(axisLabel);
  });
}

function buildViewerAllGridCornerStakeLabels(metrics) {
  const cols = Math.max(1, metrics.maxXIndex - metrics.minXIndex + 1);
  const rows = Math.max(1, metrics.maxNo - metrics.minNo + 1);
  const labels = [];
  const cellSize = 4;
  for (let xLine = 0; xLine <= cols; xLine += 1) {
    const block = incrementGridBlock(viewerXIndexToHeadBlock(metrics.minXIndex).block, xLine);
    const x = xLine * cellSize;
    for (let yLine = 0; yLine <= rows; yLine += 1) {
      const no = String(metrics.minNo + yLine);
      const y = metrics.gridHeightM - yLine * cellSize;
      labels.push({
        label: `${block}-${no}`,
        x,
        y,
      });
    }
  }
  return labels;
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
  const useOriginalImageColor = Boolean(shape?.useOriginalImageColor);
  if (points.length !== 4) {
    return;
  }
  if (!useOriginalImageColor) {
    // 画像が読めない場合でも位置が分かるように、外形線は先に描画する。
    for (let i = 0; i < points.length; i += 1) {
      const start = points[i];
      const end = points[(i + 1) % points.length];
      renderViewerSegment(start, end, new THREE.Color(shape?.color || "#6b7280"), 0.012);
    }
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
    texture.minFilter = THREE.LinearFilter;
    texture.magFilter = THREE.LinearFilter;
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
  const dilateIterations = getImageShapeDilateIterations(shape?.imageType);
  const loadFromImagePath = () => {
    if (useOriginalImageColor) {
      return getOrLoadPlanLargeShapeImage(imagePath, shape?.imageType)
        .then((image) => {
          const texture = new THREE.Texture(image);
          addTexturedMesh(texture, "#ffffff");
          return true;
        })
        .catch(() => false);
    }
    return getOrLoadPlanLargeShapeTintedCanvas(imagePath, shape?.color, shape?.imageType, { dilateIterations })
      .then((tintedCanvas) => {
        const texture = new THREE.CanvasTexture(tintedCanvas);
        addTexturedMesh(texture, "#ffffff");
        return true;
      })
      .catch(() =>
        getOrLoadPlanLargeShapeImage(imagePath, shape?.imageType).then((image) => {
          const texture = new THREE.Texture(image);
          addTexturedMesh(texture, "#ffffff");
          return true;
        })
      )
      .catch(() => false);
  };
  // 3Dは常に着色済みCanvas経路を優先し、黒線化しやすいdataURL経路は使わない。
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

function buildTintedCanvasFromImageSource(imageSource, colorRaw) {
  const sourceCanvas = ensureCanvasFromImageSource(imageSource);
  if (!sourceCanvas) {
    return null;
  }
  const width = Math.max(1, Number(sourceCanvas.width) || 0);
  const height = Math.max(1, Number(sourceCanvas.height) || 0);
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
  ctx.clearRect(0, 0, width, height);
  ctx.drawImage(sourceCanvas, 0, 0, width, height);
  ctx.globalCompositeOperation = "source-in";
  ctx.fillStyle = parseHexColor(colorRaw).hex;
  ctx.fillRect(0, 0, width, height);
  ctx.globalCompositeOperation = "source-over";
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

function getOrLoadPlanLargeShapeTintedCanvas(imagePathRaw, colorRaw, shapeTypeRaw = "", options = {}) {
  const candidates = getLargeShapeImagePathCandidates(shapeTypeRaw, imagePathRaw);
  if (!candidates.length) {
    return Promise.reject(new Error("imagePath is empty"));
  }
  const tint = parseHexColor(colorRaw);
  const requestedIterations = Number(options?.dilateIterations);
  const dilateIterations = Number.isFinite(requestedIterations)
    ? clamp(Math.round(requestedIterations), 0, 8)
    : getImageShapeDilateIterations(shapeTypeRaw);
  const key = `${candidates.join("|")}::${tint.hex}::d${dilateIterations}`;
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
      // getImageData が失敗する環境でも、透過マスクに分類色を掛けて表示する。
      ctx.globalCompositeOperation = "source-in";
      ctx.fillStyle = tint.hex;
      ctx.fillRect(0, 0, width, height);
      ctx.globalCompositeOperation = "source-over";
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
    const expandedAlpha = dilateAlphaMask(sourceMask, width, height, dilateIterations);
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

function getPlanLargeShapeTintedDataUrl(imagePathRaw, colorRaw, shapeTypeRaw = "", options = {}) {
  const candidates = getLargeShapeImagePathCandidates(shapeTypeRaw, imagePathRaw);
  if (!candidates.length) {
    return "";
  }
  const tint = parseHexColor(colorRaw);
  const requestedIterations = Number(options?.dilateIterations);
  const dilateIterations = Number.isFinite(requestedIterations)
    ? clamp(Math.round(requestedIterations), 0, 8)
    : getImageShapeDilateIterations(shapeTypeRaw);
  const key = `${candidates.join("|")}::${tint.hex}::d${dilateIterations}`;
  const cached = planLargeShapeTintedDataUrlCache.get(key);
  if (cached === "loading") {
    return "";
  }
  if (typeof cached === "string") {
    return cached;
  }
  const preferredSource = candidates.find((candidate) => !String(candidate).startsWith("data:")) || candidates[0];
  planLargeShapeTintedDataUrlCache.set(key, "loading");
  getOrLoadPlanLargeShapeTintedCanvas(preferredSource, tint.hex, shapeTypeRaw, { dilateIterations })
    .then((canvas) => {
      const dataUrl = canvas.toDataURL("image/png");
      if (dataUrl.length > PLAN_IMAGE_TINTED_DATA_URL_MAX_LENGTH) {
        planLargeShapeTintedDataUrlCache.set(key, "");
        renderOutputs();
        return;
      }
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
  const resolutionScale = clamp(Math.round(Number(options?.resolutionScale) || 2), 1, 4);
  const sizeAttenuation = options?.sizeAttenuation !== false;
  const canvas = document.createElement("canvas");
  canvas.width = 512 * resolutionScale;
  canvas.height = 128 * resolutionScale;
  const ctx = canvas.getContext("2d");
  ctx.clearRect(0, 0, canvas.width, canvas.height);
  ctx.font = `700 ${fontPx * resolutionScale}px 'Yu Gothic', sans-serif`;
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.fillStyle = value(colorRaw) || "#111827";
  ctx.strokeStyle = "rgba(255,255,255,0.96)";
  ctx.lineWidth = 8 * resolutionScale;
  ctx.strokeText(value(textRaw) || "-", canvas.width / 2, canvas.height / 2);
  ctx.fillText(value(textRaw) || "-", canvas.width / 2, canvas.height / 2);
  const texture = new THREE.CanvasTexture(canvas);
  texture.generateMipmaps = false;
  texture.minFilter = THREE.LinearFilter;
  texture.magFilter = THREE.LinearFilter;
  texture.needsUpdate = true;
  const material = new THREE.SpriteMaterial({ map: texture, transparent: true, depthTest: false, sizeAttenuation });
  const sprite = new THREE.Sprite(material);
  sprite.scale.set(scaleX, scaleY, 1);
  return sprite;
}

function createViewerTextPlane(textRaw, colorRaw, options = {}) {
  const fontPx = Math.max(12, Number(options?.fontPx) || 44);
  const width = Math.max(0.05, Number(options?.width) || 0.8);
  const height = Math.max(0.05, Number(options?.height) || 0.22);
  const resolutionScale = clamp(Math.round(Number(options?.resolutionScale) || 2), 1, 4);
  const canvas = document.createElement("canvas");
  canvas.width = 512 * resolutionScale;
  canvas.height = 128 * resolutionScale;
  const ctx = canvas.getContext("2d");
  ctx.clearRect(0, 0, canvas.width, canvas.height);
  ctx.font = `700 ${fontPx * resolutionScale}px 'Yu Gothic', sans-serif`;
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.fillStyle = value(colorRaw) || "#111827";
  ctx.strokeStyle = "rgba(255,255,255,0.96)";
  ctx.lineWidth = 8 * resolutionScale;
  ctx.strokeText(value(textRaw) || "-", canvas.width / 2, canvas.height / 2);
  ctx.fillText(value(textRaw) || "-", canvas.width / 2, canvas.height / 2);
  const texture = new THREE.CanvasTexture(canvas);
  texture.generateMipmaps = false;
  texture.minFilter = THREE.LinearFilter;
  texture.magFilter = THREE.LinearFilter;
  texture.needsUpdate = true;
  const material = new THREE.MeshBasicMaterial({
    map: texture,
    transparent: true,
    depthTest: false,
    side: THREE.DoubleSide,
  });
  const geometry = new THREE.PlaneGeometry(width, height);
  const mesh = new THREE.Mesh(geometry, material);
  mesh.renderOrder = 5;
  return mesh;
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
  const focus = new THREE.Vector3(center.x, center.y, bounds.minZ + spanZ * 0.16);
  const baseDist = Math.max(spanX, spanY) * 0.68 + Math.max(3.2, unscaledSpanZ * 2.2);
  const zoomInFactor = Math.pow(scale, -0.7);
  const dist = baseDist * zoomInFactor * 0.82;
  const perspective = selectedViewerPerspective === "iso" ? "se" : selectedViewerPerspective;
  if (perspective === "top") {
    viewer3d.camera.up.set(0, 1, 0);
    viewer3d.camera.position.set(focus.x, focus.y, focus.z + dist);
  } else {
    viewer3d.camera.up.set(0, 0, 1);
    const sideElev = dist * 0.14;
    if (perspective === "east") {
      viewer3d.camera.position.set(focus.x + dist, focus.y, focus.z + sideElev);
    } else if (perspective === "west") {
      viewer3d.camera.position.set(focus.x - dist, focus.y, focus.z + sideElev);
    } else if (perspective === "south") {
      viewer3d.camera.position.set(focus.x, focus.y - dist, focus.z + sideElev);
    } else if (perspective === "north") {
      viewer3d.camera.position.set(focus.x, focus.y + dist, focus.z + sideElev);
    } else {
      viewer3d.camera.position.set(focus.x + dist * 0.75, focus.y - dist * 0.75, focus.z + dist * 0.46);
    }
  }
  if (viewer3d.controls) {
    viewer3d.controls.target.copy(focus);
    viewer3d.controls.update();
  } else {
    viewer3d.camera.lookAt(focus);
  }
}

function handleViewerControlPointerDown(event) {
  if (!viewer3d.controls || !window.THREE?.MOUSE) {
    return;
  }
  if (isTouchLikePointerEvent(event)) {
    return;
  }
  if (Number(event.button) === 0 && event.shiftKey) {
    viewer3d.controls.mouseButtons.LEFT = THREE.MOUSE.PAN;
    viewer3d.shiftPanActive = true;
    return;
  }
  viewer3d.controls.mouseButtons.LEFT =
    viewer3d.defaultLeftMouseAction != null ? viewer3d.defaultLeftMouseAction : THREE.MOUSE.ROTATE;
  viewer3d.shiftPanActive = false;
}

function handleViewerControlPointerUp() {
  if (!viewer3d.controls || !window.THREE?.MOUSE) {
    return;
  }
  if (!viewer3d.shiftPanActive) {
    return;
  }
  viewer3d.controls.mouseButtons.LEFT =
    viewer3d.defaultLeftMouseAction != null ? viewer3d.defaultLeftMouseAction : THREE.MOUSE.ROTATE;
  viewer3d.shiftPanActive = false;
}

function handleViewerPointerMove(event) {
  if (isTouchLikePointerEvent(event)) {
    updateViewerTouchLongPressByMove(event);
    hideViewerTooltip();
    return;
  }
  const picked = pickViewerDataAtEvent(event);
  if (!picked) {
    hideViewerTooltip();
    return;
  }
  showViewerTooltip(event, picked);
}

function pickViewerDataAtEvent(event) {
  return pickViewerDataAtClient(event?.clientX, event?.clientY);
}

function pickViewerDataAtClient(clientXRaw, clientYRaw) {
  if (!viewer3d.initialized || !viewer3d.raycaster || !viewer3d.camera || !viewer3d.pickMeshes.length || !viewerCanvasWrap) {
    return null;
  }
  const clientX = Number(clientXRaw);
  const clientY = Number(clientYRaw);
  if (!Number.isFinite(clientX) || !Number.isFinite(clientY)) {
    return null;
  }
  const rect = viewerCanvasWrap.getBoundingClientRect();
  if (rect.width <= 0 || rect.height <= 0) {
    return null;
  }
  const x = ((clientX - rect.left) / rect.width) * 2 - 1;
  const y = -((clientY - rect.top) / rect.height) * 2 + 1;
  viewer3d.pointer.set(x, y);
  viewer3d.raycaster.setFromCamera(viewer3d.pointer, viewer3d.camera);
  const intersects = viewer3d.raycaster.intersectObjects(viewer3d.pickMeshes, false);
  if (!intersects.length) {
    return null;
  }
  return intersects[0].object?.userData || null;
}

function handleViewerPointerDown(event) {
  if (!isTouchLikePointerEvent(event)) {
    return;
  }
  cancelViewerTouchLongPress();
  viewerTouchLongPressState.pointerId = Number.isFinite(Number(event.pointerId)) ? event.pointerId : null;
  viewerTouchLongPressState.startX = Number(event.clientX) || 0;
  viewerTouchLongPressState.startY = Number(event.clientY) || 0;
  viewerTouchLongPressState.triggered = false;
  viewerTouchLongPressState.timer = window.setTimeout(() => {
    if (viewerTouchLongPressState.pointerId == null) {
      return;
    }
    viewerTouchLongPressState.timer = 0;
    const picked = pickViewerDataAtClient(viewerTouchLongPressState.startX, viewerTouchLongPressState.startY);
    if (!picked || !value(picked.id)) {
      return;
    }
    hideViewerTooltip();
    showHoverEditMenu(
      viewerTouchLongPressState.startX,
      viewerTouchLongPressState.startY,
      picked.id,
      picked.kuwaku,
      picked.label
    );
    viewerTouchLongPressState.triggered = true;
  }, TOUCH_LONG_PRESS_MS);
}

function updateViewerTouchLongPressByMove(event) {
  if (!isTouchLikePointerEvent(event) || viewerTouchLongPressState.pointerId == null) {
    return;
  }
  if (Number(event.pointerId) !== Number(viewerTouchLongPressState.pointerId)) {
    return;
  }
  const moved = pointerMovedBeyondThreshold(
    event.clientX,
    event.clientY,
    viewerTouchLongPressState.startX,
    viewerTouchLongPressState.startY,
    TOUCH_LONG_PRESS_MOVE_THRESHOLD_PX
  );
  if (moved) {
    cancelViewerTouchLongPress();
  }
}

function handleViewerPointerEnd(event) {
  if (!isTouchLikePointerEvent(event)) {
    return;
  }
  if (viewerTouchLongPressState.pointerId == null) {
    return;
  }
  if (Number(event.pointerId) !== Number(viewerTouchLongPressState.pointerId)) {
    return;
  }
  cancelViewerTouchLongPress();
}

function handleViewerPointerLeave(event) {
  hideViewerTooltip();
  if (isTouchLikePointerEvent(event)) {
    cancelViewerTouchLongPress();
  }
}

function cancelViewerTouchLongPress() {
  if (viewerTouchLongPressState.timer) {
    window.clearTimeout(viewerTouchLongPressState.timer);
  }
  viewerTouchLongPressState.timer = 0;
  viewerTouchLongPressState.pointerId = null;
  viewerTouchLongPressState.triggered = false;
}

function handleViewerContextMenu(event) {
  event.preventDefault();
  event.stopPropagation();
  const picked = pickViewerDataAtEvent(event);
  if (!picked || !value(picked.id)) {
    hideHoverEditMenu();
    return;
  }
  showHoverEditMenu(event.clientX, event.clientY, picked.id, picked.kuwaku, picked.label);
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
    id: value(record.id),
    kuwaku: value(record.kuwaku) || getRecordKuwaku(record),
    color: getSpecimenPrefixColor(prefix),
    label: record.specimenNo || "",
    nameMemo: value(record.nameMemo),
    unit: value(record.unit),
    detail: buildDetailText(record.detail, record.detailSub),
  };
}

function buildTotalStationPlanMarker(record) {
  if (normalizePositionMethod(record?.positionMethod) !== "totalStation") return null;
  const peg = parseTotalStationPeg(record?.tsBacksightPeg);
  const grid = parseKuwaku(getRecordKuwaku(record));
  const stationX = parseTotalStationNumber(record?.tsStationXNorthM);
  const stationY = parseTotalStationNumber(record?.tsStationYEastM);
  const backX = parseTotalStationNumber(record?.tsBacksightXNorthM);
  const backY = parseTotalStationNumber(record?.tsBacksightYEastM);
  if (!peg || !grid.block || !grid.no || [stationX, stationY, backX, backY].some((number) => number == null)) return null;
  const isSouthWestPositive = value(record?.tsCoordinateConvention) === "southWestPositive";
  const eastOffsetM = (stationY - backY) * (isSouthWestPositive ? -1 : 1);
  const southOffsetM = (stationX - backX) * (isSouthWestPositive ? 1 : -1);
  return {
    x: (blockLabelToIndex(peg.block) - blockLabelToIndex(grid.block)) * PLAN_SIZE_CM + eastOffsetM * 100,
    y: (Number(peg.no) - Number(grid.no)) * PLAN_SIZE_CM + southOffsetM * 100,
  };
}

function collectTotalStationPlanMarkers(records) {
  const unique = new Map();
  (Array.isArray(records) ? records : []).forEach((record) => {
    const marker = buildTotalStationPlanMarker(record);
    if (!marker) return;
    const key = `${marker.x.toFixed(3)}:${marker.y.toFixed(3)}`;
    if (!unique.has(key)) unique.set(key, marker);
  });
  return [...unique.values()];
}

function renderTotalStationPlanMarkersSvg(markers) {
  return markers.map((marker) => `<g class="plan-total-station" aria-label="トータルステーション設置点"><circle cx="${marker.x}" cy="${marker.y}" r="8" fill="#dbeafe" stroke="#0067c5" stroke-width="3" /><line x1="${marker.x - 11}" y1="${marker.y}" x2="${marker.x + 11}" y2="${marker.y}" stroke="#0067c5" stroke-width="2" /><line x1="${marker.x}" y1="${marker.y - 11}" x2="${marker.x}" y2="${marker.y + 11}" stroke="#0067c5" stroke-width="2" /><text x="${marker.x + 12}" y="${marker.y - 10}" fill="#0067c5" font-size="13" font-weight="800">TS</text></g>`).join("");
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

function parseLargeAxisPlungeDeg(valueRaw) {
  const text = value(valueRaw).replace(/[°度]/g, "");
  if (!text) {
    return null;
  }
  const matched = text.match(/-?\d+(?:\.\d+)?/);
  if (!matched) {
    return null;
  }
  const num = Number(matched[0]);
  if (!Number.isFinite(num)) {
    return null;
  }
  return clamp(Math.abs(num), 0, 90);
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

function buildConvexHull2d(pointsRaw) {
  const points = Array.isArray(pointsRaw) ? pointsRaw : [];
  const uniqueMap = new Map();
  points.forEach((point) => {
    const x = Number(point?.x);
    const y = Number(point?.y);
    if (!Number.isFinite(x) || !Number.isFinite(y)) {
      return;
    }
    const key = `${x.toFixed(6)}|${y.toFixed(6)}`;
    if (!uniqueMap.has(key)) {
      uniqueMap.set(key, { x, y });
    }
  });
  const unique = Array.from(uniqueMap.values());
  if (unique.length <= 2) {
    return unique;
  }
  unique.sort((a, b) => (a.x === b.x ? a.y - b.y : a.x - b.x));
  const cross = (origin, a, b) => (a.x - origin.x) * (b.y - origin.y) - (a.y - origin.y) * (b.x - origin.x);

  const lower = [];
  unique.forEach((point) => {
    while (lower.length >= 2 && cross(lower[lower.length - 2], lower[lower.length - 1], point) <= 0) {
      lower.pop();
    }
    lower.push(point);
  });

  const upper = [];
  for (let i = unique.length - 1; i >= 0; i -= 1) {
    const point = unique[i];
    while (upper.length >= 2 && cross(upper[upper.length - 2], upper[upper.length - 1], point) <= 0) {
      upper.pop();
    }
    upper.push(point);
  }

  lower.pop();
  upper.pop();
  return [...lower, ...upper];
}

function buildHullPointsFromSource(pointsRaw, zFallback = null) {
  const points = Array.isArray(pointsRaw) ? pointsRaw : [];
  if (points.length <= 2) {
    return points.filter((point) => Number.isFinite(Number(point?.x)) && Number.isFinite(Number(point?.y)));
  }
  const hull2d = buildConvexHull2d(points);
  if (!hull2d.length) {
    return [];
  }
  const keyOf = (xRaw, yRaw) => `${Number(xRaw).toFixed(6)}|${Number(yRaw).toFixed(6)}`;
  const pointMap = new Map();
  points.forEach((point) => {
    const x = Number(point?.x);
    const y = Number(point?.y);
    if (!Number.isFinite(x) || !Number.isFinite(y)) {
      return;
    }
    pointMap.set(keyOf(x, y), point);
  });
  const fallbackZ = Number.isFinite(Number(zFallback)) ? Number(zFallback) : 0;
  return hull2d.map((point) => {
    const original = pointMap.get(keyOf(point.x, point.y));
    if (original) {
      return original;
    }
    return { x: point.x, y: point.y, z: fallbackZ };
  });
}

function parseImageQuadPlanPoints(record, center = null) {
  const centerPointRaw = center
    ? { x: Number(center.x), y: Number(center.y) }
    : convertPositionToPlanCoords(record?.nsDir, record?.nsCm, record?.ewDir, record?.ewCm);
  const centerPoint =
    centerPointRaw && Number.isFinite(centerPointRaw.x) && Number.isFinite(centerPointRaw.y) ? centerPointRaw : null;
  const widthCm = parseDistanceToCm(record?.imgFrameWidthCm);
  const heightCm = parseDistanceToCm(record?.imgFrameHeightCm);
  // 回転角が未入力でも、画像外枠サイズは反映する（未入力時は0°扱い）。
  const rotationText = normalizeImageRotationDeg(record?.imgRotateDeg);
  if (centerPoint && widthCm != null && widthCm > 0 && heightCm != null && heightCm > 0) {
    const rotationDeg = Number(rotationText === "" ? "0" : rotationText);
    const skewXDeg = Number(normalizeImageSkewDeg(record?.imgSkewXDeg) || "0");
    const skewYDeg = Number(normalizeImageSkewDeg(record?.imgSkewYDeg) || "0");
    const flipH = normalizeToggleFlag(record?.imgFlipH) === "1";
    const flipV = normalizeToggleFlag(record?.imgFlipV) === "1";

    const up = azimuthToPlanUnitVector(rotationDeg);
    const down = { dx: -up.dx, dy: -up.dy };
    const right = azimuthToPlanUnitVector((rotationDeg + 90) % 360);
    const tanSkewX = Math.tan((skewXDeg * Math.PI) / 180);
    const tanSkewY = Math.tan((skewYDeg * Math.PI) / 180);
    const baseLocalCorners = [
      { x: -widthCm / 2, y: -heightCm / 2 }, // tl
      { x: widthCm / 2, y: -heightCm / 2 }, // tr
      { x: widthCm / 2, y: heightCm / 2 }, // br
      { x: -widthCm / 2, y: heightCm / 2 }, // bl
    ];
    const shearedWorldCorners = baseLocalCorners.map((local) => {
      const xSheared = local.x + tanSkewX * local.y;
      const ySheared = local.y + tanSkewY * local.x;
      return {
        x: centerPoint.x + right.dx * xSheared + down.dx * ySheared,
        y: centerPoint.y + right.dy * xSheared + down.dy * ySheared,
      };
    });
    let order = [0, 1, 2, 3];
    if (flipH && flipV) {
      order = [2, 3, 0, 1];
    } else if (flipH) {
      order = [1, 0, 3, 2];
    } else if (flipV) {
      order = [3, 2, 1, 0];
    }
    return order.map((index) => shearedWorldCorners[index]);
  }

  const corner1 = convertPositionToPlanCoords(record?.imgP1NsDir, record?.imgP1NsCm, record?.imgP1EwDir, record?.imgP1EwCm);
  const corner2 = convertPositionToPlanCoords(record?.imgP2NsDir, record?.imgP2NsCm, record?.imgP2EwDir, record?.imgP2EwCm);
  const corner3 = convertPositionToPlanCoords(record?.imgP3NsDir, record?.imgP3NsCm, record?.imgP3EwDir, record?.imgP3EwCm);
  const corner4 = convertPositionToPlanCoords(record?.imgP4NsDir, record?.imgP4NsCm, record?.imgP4EwDir, record?.imgP4EwCm);
  if (!corner1 || !corner2 || !corner3 || !corner4) {
    if (!centerPoint) {
      return null;
    }
    const candidates = [corner1, corner2, corner3, corner4, centerPoint].filter(Boolean);
    if (!candidates.length) {
      return null;
    }
    let minX = Math.min(...candidates.map((point) => point.x));
    let maxX = Math.max(...candidates.map((point) => point.x));
    let minY = Math.min(...candidates.map((point) => point.y));
    let maxY = Math.max(...candidates.map((point) => point.y));
    const fallbackHalfSize = 20;
    if (Math.abs(maxX - minX) < 1) {
      minX = centerPoint.x - fallbackHalfSize;
      maxX = centerPoint.x + fallbackHalfSize;
    }
    if (Math.abs(maxY - minY) < 1) {
      minY = centerPoint.y - fallbackHalfSize;
      maxY = centerPoint.y + fallbackHalfSize;
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
  if (planSizeMode === "複数点") {
    const multiPoints = collectPlanMultiPointCoords(record);
    const fallbackCenter = convertPositionToPlanCoords(record?.nsDir, record?.nsCm, record?.ewDir, record?.ewCm);
    const points = [];
    const seenPointKeys = new Set();
    [fallbackCenter, ...multiPoints].filter(Boolean).forEach((point) => {
      const pointKey = `${point.x.toFixed(4)}|${point.y.toFixed(4)}`;
      if (seenPointKeys.has(pointKey)) return;
      seenPointKeys.add(pointKey);
      points.push(point);
    });
    if (!points.length) {
      return null;
    }
    const centroid = points.reduce(
      (acc, point) => ({ x: acc.x + point.x / points.length, y: acc.y + point.y / points.length }),
      { x: 0, y: 0 }
    );
    return {
      type: "multipoint",
      points,
      hull: buildHullPointsFromSource(points),
      x: centroid.x,
      y: centroid.y,
      ...meta,
    };
  }

  const shapeType = normalizeLargeShapeType(record.largeShapeType);
  const normalizedShapeLabel = normalizeLargeShapeLabel(record.largeShapeType);
  const customImagePath = isCustomLargeShapeType(shapeType) ? normalizeCustomLargeImageDataUrl(record.customLargeImageDataUrl) : "";
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
  const orientationAzimuth =
    shapeType === "直線状"
      ? parseLargeAxisAzimuth(record.largeAxisDirection)
      : parseLargeAxisAzimuth(record.planeStrikeDirection) ?? parseLargeAxisAzimuth(record.largeAxisDirection);
  const axisAzimuth = shouldUseImageQuad ? null : orientationAzimuth;

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
    const useOriginalImageColor = normalizeToggleFlag(record?.imgUseOriginalColor) === "1";
    const centroid = points.reduce(
      (acc, point) => ({ x: acc.x + point.x / points.length, y: acc.y + point.y / points.length }),
      { x: 0, y: 0 }
    );
    return {
      type: "imageQuad",
      points,
      imageType: resolvedImageType,
      imagePath: customImagePath || getLargeShapeImagePath(resolvedImageType),
      useOriginalImageColor,
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
  } else if (drawable.type === "multipoint") {
    const points = Array.isArray(drawable.points) ? drawable.points : [];
    const hull = Array.isArray(drawable.hull) ? drawable.hull : [];
    const hullPointsText = hull.map((point) => `${point.x},${point.y}`).join(" ");
    let hullSvg = "";
    if (hull.length >= 3) {
      hullSvg = `<polygon class="plan-shape-multipoint-hull" points="${hullPointsText}" stroke="${drawable.color}" />`;
    } else if (hull.length === 2) {
      hullSvg = `<line class="plan-shape-multipoint-hull" x1="${hull[0].x}" y1="${hull[0].y}" x2="${hull[1].x}" y2="${hull[1].y}" stroke="${drawable.color}" />`;
    }
    const pointsSvg = points
      .map((point) => `<circle class="plan-shape-multipoint-dot" cx="${point.x}" cy="${point.y}" r="4.5" fill="${drawable.color}" />`)
      .join("");
    shapeSvg = `${hullSvg}${pointsSvg}`;
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
    if (drawable.useOriginalImageColor) {
      shapeSvg = imageSvg;
    } else {
      const polygonPoints = (drawable.points || []).map((point) => `${point.x},${point.y}`).join(" ");
      const outlineSvg = `<polygon class="plan-shape-image-outline" points="${polygonPoints}" fill="none" stroke="${drawable.color}" />`;
      shapeSvg = imageSvg ? `${imageSvg}${outlineSvg}` : outlineSvg;
    }
  } else {
    shapeSvg = `<circle class="plan-point-hit" cx="${drawable.x}" cy="${drawable.y}" r="5" fill="${drawable.color}" />`;
  }

  let hotspotSvg = `<circle class="plan-point-hotspot" cx="${drawable.x}" cy="${drawable.y}" r="12" fill="transparent" />`;
  if (drawable.type === "imageQuad") {
    hotspotSvg = `<polygon class="plan-point-hotspot plan-image-hotspot" points="${(drawable.points || [])
      .map((point) => `${point.x},${point.y}`)
      .join(" ")}" fill="transparent" />`;
  } else if (drawable.type === "multipoint") {
    const points = Array.isArray(drawable.points) ? drawable.points : [];
    const hull = Array.isArray(drawable.hull) ? drawable.hull : [];
    const hotspotPoints = points
      .map((point) => `<circle class="plan-point-hotspot" cx="${point.x}" cy="${point.y}" r="10" fill="transparent" />`)
      .join("");
    if (hull.length >= 3) {
      hotspotSvg = `<polygon class="plan-point-hotspot" points="${hull
        .map((point) => `${point.x},${point.y}`)
        .join(" ")}" fill="transparent" />${hotspotPoints}`;
    } else {
      hotspotSvg = hotspotPoints || hotspotSvg;
    }
  }

  return `
      <g
        class="plan-point-group"
        data-id="${escapeHtml(value(drawable.id) || "")}"
        data-kuwaku="${escapeHtml(value(drawable.kuwaku) || "")}"
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
        <text x="${labelX}" y="${labelY}"${drawable.labelColor ? ` style="fill:${escapeHtml(drawable.labelColor)};"` : ""}>${escapeHtml(drawable.label || "")}</text>
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
  const useOriginalImageColor = Boolean(drawable?.useOriginalImageColor);
  const imageCandidates = getLargeShapeImagePathCandidates(drawable?.imageType, imagePath);
  const imageFallback =
    imageCandidates.find((candidate) => !String(candidate).startsWith("data:")) || imageCandidates[0] || imagePath;
  let imageRef = imageFallback;
  if (!useOriginalImageColor) {
    const dilateIterations = getImageShapeDilateIterations(drawable?.imageType);
    const tintedDataUrlRaw = getPlanLargeShapeTintedDataUrl(imageFallback, drawable?.color, drawable?.imageType, {
      dilateIterations,
    });
    const tintedDataUrl = value(tintedDataUrlRaw);
    const canUseTintedDataUrl =
      tintedDataUrl &&
      tintedDataUrl.startsWith("data:") &&
      tintedDataUrl.length <= PLAN_IMAGE_TINTED_DATA_URL_MAX_LENGTH;
    imageRef = canUseTintedDataUrl ? tintedDataUrl : imageFallback;
  }
  if (points.length !== 4 || !imageRef) {
    return "";
  }
  const [p1, p2, p3, p4] = points;
  const isParallelogram =
    Number.isFinite(p1.x) &&
    Number.isFinite(p1.y) &&
    Number.isFinite(p2.x) &&
    Number.isFinite(p2.y) &&
    Number.isFinite(p3.x) &&
    Number.isFinite(p3.y) &&
    Number.isFinite(p4.x) &&
    Number.isFinite(p4.y) &&
    Math.abs((p1.x + p3.x) - (p2.x + p4.x)) <= 0.8 &&
    Math.abs((p1.y + p3.y) - (p2.y + p4.y)) <= 0.8;
  if (isParallelogram) {
    const matrix = [p2.x - p1.x, p2.y - p1.y, p4.x - p1.x, p4.y - p1.y, p1.x, p1.y];
    const matrixText = matrix.map((num) => (Number.isFinite(num) ? Number(num).toFixed(4) : "0")).join(" ");
    return `<image href="${escapeHtml(imageRef)}" xlink:href="${escapeHtml(
      imageRef
    )}" x="0" y="0" width="1" height="1" preserveAspectRatio="none" transform="matrix(${matrixText})" />`;
  }
  const labelKey = value(drawable.label || "x").replace(/[^a-zA-Z0-9_-]/g, "");
  const clipIdA = `plan-img-clip-a-${index}-${labelKey || "x"}`;
  const clipIdB = `plan-img-clip-b-${index}-${labelKey || "x"}`;
  const matrixA = [p2.x - p1.x, p2.y - p1.y, p3.x - p2.x, p3.y - p2.y, p1.x, p1.y];
  const matrixB = [p3.x - p4.x, p3.y - p4.y, p4.x - p1.x, p4.y - p1.y, p1.x, p1.y];
  const triA = `${p1.x},${p1.y} ${p2.x},${p2.y} ${p3.x},${p3.y}`;
  const triB = `${p1.x},${p1.y} ${p3.x},${p3.y} ${p4.x},${p4.y}`;
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
    </defs>
    <image href="${escapeHtml(imageRef)}" xlink:href="${escapeHtml(
      imageRef
    )}" x="0" y="0" width="1" height="1" preserveAspectRatio="none" transform="matrix(${matrixAText})" clip-path="url(#${clipIdA})" />
    <image href="${escapeHtml(imageRef)}" xlink:href="${escapeHtml(
      imageRef
    )}" x="0" y="0" width="1" height="1" preserveAspectRatio="none" transform="matrix(${matrixBText})" clip-path="url(#${clipIdB})" />
  `;
}

function getPlanDrawableExtent(drawable) {
  if (!drawable || typeof drawable !== "object") {
    return null;
  }
  if (drawable.type === "line") {
    return {
      minX: Math.min(Number(drawable.x1), Number(drawable.x2)),
      maxX: Math.max(Number(drawable.x1), Number(drawable.x2)),
      minY: Math.min(Number(drawable.y1), Number(drawable.y2)),
      maxY: Math.max(Number(drawable.y1), Number(drawable.y2)),
    };
  }
  if (drawable.type === "rect") {
    const cx = Number(drawable.x);
    const cy = Number(drawable.y);
    const w = Number(drawable.width);
    const h = Number(drawable.height);
    const halfDiag = Math.hypot(w / 2, h / 2);
    if (!Number.isFinite(cx) || !Number.isFinite(cy) || !Number.isFinite(halfDiag)) {
      return null;
    }
    return {
      minX: cx - halfDiag,
      maxX: cx + halfDiag,
      minY: cy - halfDiag,
      maxY: cy + halfDiag,
    };
  }
  if (drawable.type === "ellipse") {
    const cx = Number(drawable.x);
    const cy = Number(drawable.y);
    const radius = Math.max(Number(drawable.rx), Number(drawable.ry));
    if (!Number.isFinite(cx) || !Number.isFinite(cy) || !Number.isFinite(radius)) {
      return null;
    }
    return {
      minX: cx - radius,
      maxX: cx + radius,
      minY: cy - radius,
      maxY: cy + radius,
    };
  }
  if (drawable.type === "imageQuad" && Array.isArray(drawable.points) && drawable.points.length) {
    const xs = drawable.points.map((point) => Number(point?.x)).filter((num) => Number.isFinite(num));
    const ys = drawable.points.map((point) => Number(point?.y)).filter((num) => Number.isFinite(num));
    if (!xs.length || !ys.length) {
      return null;
    }
    return {
      minX: Math.min(...xs),
      maxX: Math.max(...xs),
      minY: Math.min(...ys),
      maxY: Math.max(...ys),
    };
  }
  if (drawable.type === "multipoint") {
    const points = [
      ...(Array.isArray(drawable.points) ? drawable.points : []),
      ...(Array.isArray(drawable.hull) ? drawable.hull : []),
    ];
    const xs = points.map((point) => Number(point?.x)).filter((num) => Number.isFinite(num));
    const ys = points.map((point) => Number(point?.y)).filter((num) => Number.isFinite(num));
    if (!xs.length || !ys.length) {
      return null;
    }
    return {
      minX: Math.min(...xs),
      maxX: Math.max(...xs),
      minY: Math.min(...ys),
      maxY: Math.max(...ys),
    };
  }
  const cx = Number(drawable.x);
  const cy = Number(drawable.y);
  if (!Number.isFinite(cx) || !Number.isFinite(cy)) {
    return null;
  }
  return {
    minX: cx,
    maxX: cx,
    minY: cy,
    maxY: cy,
  };
}

function computePlanSvgViewBox() {
  const leftPad = 40;
  const rightPad = 24;
  const topPad = 24;
  const bottomPad = 24;
  return {
    minX: -leftPad,
    minY: -topPad,
    width: PLAN_SIZE_CM + leftPad + rightPad,
    height: PLAN_SIZE_CM + topPad + bottomPad,
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

function buildPlanCornerLabelsSvg(cornerLabels) {
  const labels = cornerLabels || {};
  const tl = escapeHtml(value(labels.topLeft) || "-");
  const tr = escapeHtml(value(labels.topRight) || "-");
  const bl = escapeHtml(value(labels.bottomLeft) || "-");
  const br = escapeHtml(value(labels.bottomRight) || "-");
  return `
    <g class="plan-grid-corner-svg">
      <text x="4" y="4" text-anchor="start" dominant-baseline="hanging">${tl}</text>
      <text x="${PLAN_SIZE_CM - 4}" y="4" text-anchor="end" dominant-baseline="hanging">${tr}</text>
      <text x="4" y="${PLAN_SIZE_CM - 4}" text-anchor="start" dominant-baseline="ideographic">${bl}</text>
      <text x="${PLAN_SIZE_CM - 4}" y="${PLAN_SIZE_CM - 4}" text-anchor="end" dominant-baseline="ideographic">${br}</text>
    </g>
  `;
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

function buildPlanLegendHtml(options = {}) {
  const includeTotalStation = options.includeTotalStation !== false;
  const order = ["m", "b", "l", "s", "i", "g", "h", "a"];
  const specimenLegend = order
    .map((prefix) => {
      const color = getSpecimenPrefixColor(prefix);
      const label = SPECIMEN_CATEGORY_MAP[prefix] || "";
      return `<span class="plan-legend-item"><span class="plan-legend-dot" style="background:${color}"></span>${prefix}: ${label}</span>`;
    })
    .join("");
  const totalStationLegend = includeTotalStation
    ? '<span class="plan-legend-item"><span class="plan-legend-dot" style="background:#0067c5"></span>TS: トータルステーション設置位置</span>'
    : "";
  return `${specimenLegend}${totalStationLegend}`;
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
    const touchLongPressState = {
      pointerId: null,
      startX: 0,
      startY: 0,
      timer: 0,
    };
    const clearTouchLongPress = () => {
      if (touchLongPressState.timer) {
        window.clearTimeout(touchLongPressState.timer);
      }
      touchLongPressState.timer = 0;
      touchLongPressState.pointerId = null;
    };

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
    pointEl.addEventListener("contextmenu", (event) => {
      event.preventDefault();
      event.stopPropagation();
      const recordId = value(pointEl.dataset.id);
      if (!recordId) {
        hideHoverEditMenu();
        return;
      }
      show(pointEl, event);
      showHoverEditMenu(event.clientX, event.clientY, recordId, value(pointEl.dataset.kuwaku), value(pointEl.dataset.label));
    });
    pointEl.addEventListener("pointerdown", (event) => {
      if (!isTouchLikePointerEvent(event)) {
        return;
      }
      const recordId = value(pointEl.dataset.id);
      if (!recordId) {
        return;
      }
      clearTouchLongPress();
      touchLongPressState.pointerId = Number.isFinite(Number(event.pointerId)) ? event.pointerId : null;
      touchLongPressState.startX = Number(event.clientX) || 0;
      touchLongPressState.startY = Number(event.clientY) || 0;
      touchLongPressState.timer = window.setTimeout(() => {
        if (touchLongPressState.pointerId == null) {
          return;
        }
        touchLongPressState.timer = 0;
        show(pointEl, { clientX: touchLongPressState.startX, clientY: touchLongPressState.startY });
        showHoverEditMenu(
          touchLongPressState.startX,
          touchLongPressState.startY,
          recordId,
          value(pointEl.dataset.kuwaku),
          value(pointEl.dataset.label)
        );
      }, TOUCH_LONG_PRESS_MS);
    });
    pointEl.addEventListener("pointermove", (event) => {
      if (!isTouchLikePointerEvent(event) || touchLongPressState.pointerId == null) {
        return;
      }
      if (Number(event.pointerId) !== Number(touchLongPressState.pointerId)) {
        return;
      }
      const moved = pointerMovedBeyondThreshold(
        event.clientX,
        event.clientY,
        touchLongPressState.startX,
        touchLongPressState.startY,
        TOUCH_LONG_PRESS_MOVE_THRESHOLD_PX
      );
      if (moved) {
        clearTouchLongPress();
      }
    });
    pointEl.addEventListener("pointerup", (event) => {
      if (!isTouchLikePointerEvent(event) || touchLongPressState.pointerId == null) {
        return;
      }
      if (Number(event.pointerId) !== Number(touchLongPressState.pointerId)) {
        return;
      }
      clearTouchLongPress();
    });
    pointEl.addEventListener("pointercancel", clearTouchLongPress);
    pointEl.addEventListener("pointerleave", clearTouchLongPress);
  });

  shell.addEventListener("click", (event) => {
    if (event.target.closest(".plan-point-group")) {
      return;
    }
    hide();
    hideHoverEditMenu();
  });
  shell.addEventListener("contextmenu", (event) => {
    if (event.target.closest(".plan-point-group")) {
      return;
    }
    hideHoverEditMenu();
  });
}

function ensureHoverEditMenu() {
  if (hoverEditMenuEl) {
    return hoverEditMenuEl;
  }
  const menu = document.createElement("div");
  menu.className = "hover-edit-menu";
  menu.hidden = true;
  menu.innerHTML = `
    <div class="hover-edit-menu-title">このデータを編集</div>
    <button type="button" class="hover-edit-menu-button">編集</button>
  `;
  const button = menu.querySelector(".hover-edit-menu-button");
  if (button) {
    button.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      const recordId = value(hoverEditMenuRecordId);
      if (!recordId) {
        hideHoverEditMenu();
        return;
      }
      const preferredKuwaku = value(hoverEditMenuKuwaku);
      hideHoverEditMenu();
      openRecordForEdit(recordId, preferredKuwaku);
    });
  }
  menu.addEventListener("click", (event) => {
    event.stopPropagation();
  });
  document.body.appendChild(menu);
  hoverEditMenuEl = menu;
  return hoverEditMenuEl;
}

function showHoverEditMenu(clientXRaw, clientYRaw, recordIdRaw, kuwakuRaw = "", labelRaw = "") {
  const recordId = value(recordIdRaw);
  if (!recordId) {
    hideHoverEditMenu();
    return;
  }
  const menu = ensureHoverEditMenu();
  hoverEditMenuRecordId = recordId;
  hoverEditMenuKuwaku = value(kuwakuRaw);
  const title = menu.querySelector(".hover-edit-menu-title");
  const labelText = value(labelRaw);
  if (title) {
    title.textContent = labelText ? `${labelText} を編集` : "このデータを編集";
  }
  menu.hidden = false;
  const clientX = Number(clientXRaw);
  const clientY = Number(clientYRaw);
  const fallbackX = Math.max(0, Math.floor(window.innerWidth / 2));
  const fallbackY = Math.max(0, Math.floor(window.innerHeight / 2));
  const anchorX = Number.isFinite(clientX) ? clientX : fallbackX;
  const anchorY = Number.isFinite(clientY) ? clientY : fallbackY;
  const margin = 8;
  const offset = 12;
  const rect = menu.getBoundingClientRect();
  const maxLeft = Math.max(margin, window.innerWidth - rect.width - margin);
  const maxTop = Math.max(margin, window.innerHeight - rect.height - margin);
  const left = clamp(anchorX + offset, margin, maxLeft);
  const top = clamp(anchorY + offset, margin, maxTop);
  menu.style.left = `${left}px`;
  menu.style.top = `${top}px`;
}

function hideHoverEditMenu() {
  if (!hoverEditMenuEl) {
    return;
  }
  hoverEditMenuEl.hidden = true;
  hoverEditMenuRecordId = "";
  hoverEditMenuKuwaku = "";
}

function isTouchLikePointerEvent(event) {
  const pointerType = value(event?.pointerType).toLowerCase();
  return pointerType === "touch" || pointerType === "pen";
}

function pointerMovedBeyondThreshold(clientXRaw, clientYRaw, startXRaw, startYRaw, thresholdRaw = 12) {
  const clientX = Number(clientXRaw);
  const clientY = Number(clientYRaw);
  const startX = Number(startXRaw);
  const startY = Number(startYRaw);
  if (!Number.isFinite(clientX) || !Number.isFinite(clientY) || !Number.isFinite(startX) || !Number.isFinite(startY)) {
    return false;
  }
  const threshold = Math.max(1, Number(thresholdRaw) || 1);
  return Math.hypot(clientX - startX, clientY - startY) > threshold;
}

function positionTooltip(pointEl, mouseEvent, shell, svg, tooltip) {
  const shellRect = shell.getBoundingClientRect();

  let xLocal = 0;
  let yLocal = 0;
  if (mouseEvent && typeof mouseEvent.clientX === "number" && typeof mouseEvent.clientY === "number") {
    xLocal = mouseEvent.clientX - shellRect.left;
    yLocal = mouseEvent.clientY - shellRect.top;
  } else {
    const pointRect = pointEl.getBoundingClientRect();
    xLocal = pointRect.left - shellRect.left + pointRect.width / 2;
    yLocal = pointRect.top - shellRect.top + pointRect.height / 2;
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

function ensureUniqueRecordIds(recordsRaw) {
  if (!Array.isArray(recordsRaw)) {
    return [];
  }
  const seenIds = new Set();
  return recordsRaw.map((recordRaw) => {
    if (!recordRaw || typeof recordRaw !== "object") {
      return recordRaw;
    }
    const record = { ...recordRaw };
    const currentId = value(record.id);
    if (!currentId || seenIds.has(currentId)) {
      record.id = newId("record");
      seenIds.add(record.id);
      return record;
    }
    seenIds.add(currentId);
    return record;
  });
}

function normalizeState(candidate) {
  const safe = createInitialState();
  if (!candidate || typeof candidate !== "object") {
    return safe;
  }

  const kuwakuParts = parseKuwaku(value(candidate.site?.kuwaku));
  let kuwakuHeadA = normalizeKuwakuHeadA(value(candidate.site?.kuwakuHeadA) || kuwakuParts.headA || DEFAULT_KUWAKU_HEAD_A);
  const kuwakuHeadB = normalizeKuwakuHeadB(value(candidate.site?.kuwakuHeadB) || kuwakuParts.headB || DEFAULT_KUWAKU_HEAD_B);
  const kuwakuBlock = normalizeKuwakuBlock(value(candidate.site?.kuwakuBlock) || kuwakuParts.block);
  const kuwakuNo = normalizeKuwakuNo(value(candidate.site?.kuwakuNo) || kuwakuParts.no);
  const teamState = normalizeTeamState(value(candidate.site?.team), value(candidate.site?.teamOther));
  // 第25次専用。旧版が端末内に保存した先頭値「24」も自動的に「25」へ移行する。
  if (kuwakuHeadA === "24") kuwakuHeadA = DEFAULT_KUWAKU_HEAD_A;

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
    scribe: value(candidate.site?.scribe),
  };

  if (Array.isArray(candidate.records)) {
    safe.records = ensureUniqueRecordIds(candidate.records.map((item) => normalizeRecord(item, safe.site)).filter(Boolean));
    return safe;
  }

  const artifacts = Array.isArray(candidate.artifacts) ? candidate.artifacts : [];
  const cards = candidate.cards && typeof candidate.cards === "object" ? candidate.cards : {};
  const photos = candidate.photos && typeof candidate.photos === "object" ? candidate.photos : {};

  safe.records = ensureUniqueRecordIds(
    artifacts
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
        layerColor: value(card.layerColor),
        layerLithology: value(card.layerLithology),
        layerFacies: value(card.layerFacies),
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
    .filter(Boolean)
  );

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
    altitudeInputEnabled: normalizeToggleFlag(value(item.altitudeInputEnabled) || (value(item.altitudeDirectM) ? "1" : "")),
    altitudeDirectM: value(item.altitudeDirectM) || value(item.altitudeInputM),
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
    positionMethod: normalizePositionMethod(item.positionMethod),
    tsStationPeg: value(item.tsStationPeg),
    tsCoordinateConvention: value(item.tsCoordinateConvention),
    tsStationXNorthM: value(item.tsStationXNorthM),
    tsStationYEastM: value(item.tsStationYEastM),
    tsStationAltitudeM: value(item.tsStationAltitudeM),
    tsBacksightPeg: value(item.tsBacksightPeg),
    tsBacksightXNorthM: value(item.tsBacksightXNorthM),
    tsBacksightYEastM: value(item.tsBacksightYEastM),
    tsBacksightAltitudeM: value(item.tsBacksightAltitudeM),
    tsInstrumentHeightM: value(item.tsInstrumentHeightM),
    tsTargetHeightM: value(item.tsTargetHeightM),
    tsObservationMode: value(item.tsObservationMode) === "polar" ? "polar" : "coordinate",
    tsPointCoordinateMode: value(item.tsPointCoordinateMode),
    tsPointXNorthM: value(item.tsPointXNorthM),
    tsPointYEastM: value(item.tsPointYEastM),
    tsPointAltitudeM: value(item.tsPointAltitudeM),
    tsSlopeDistanceM: value(item.tsSlopeDistanceM),
    tsInclinationDeg: value(item.tsInclinationDeg),
    tsInclinationMin: value(item.tsInclinationMin),
    tsInclinationSec: value(item.tsInclinationSec),
    tsDirectionDeg: value(item.tsDirectionDeg),
    tsDirectionMin: value(item.tsDirectionMin),
    tsDirectionSec: value(item.tsDirectionSec),
    multiPoints: normalizePlanMultiPoints(item.multiPoints),
    planSizeMode: normalizePlanSizeMode(value(item.planSizeMode)),
    largeShapeType: normalizedLargeShapeType,
    largeAxisDirection: normalizeLargeAxisDirection(value(item.largeAxisDirection)),
    largeAxisPlungeDeg: normalizeLargeAxisPlungeDeg(value(item.largeAxisPlungeDeg)),
    largeAxisPlungeDir8: normalizeCompass8Direction(
      value(item.largeAxisPlungeDir8) || value(item.largeAxisPlungeDirection) || value(item.plungeDir8)
    ),
    planeStrikeDirection: normalizePlaneStrikeDirection(value(item.planeStrikeDirection)),
    planeDipDeg: normalizePlaneDipDeg(value(item.planeDipDeg)),
    planeDipDir8: normalizeCompass8Direction(value(item.planeDipDir8) || value(item.planeDipDirection)),
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
    imgRotateDeg: normalizeImageRotationDeg(value(item.imgRotateDeg)),
    imgFrameWidthCm: value(item.imgFrameWidthCm),
    imgFrameHeightCm: value(item.imgFrameHeightCm),
    imgSkewXDeg: normalizeImageSkewDeg(value(item.imgSkewXDeg)),
    imgSkewYDeg: normalizeImageSkewDeg(value(item.imgSkewYDeg)),
    imgFlipH: normalizeToggleFlag(value(item.imgFlipH)),
    imgFlipV: normalizeToggleFlag(value(item.imgFlipV)),
    imgLockAspectRatio: normalizeToggleFlag(value(item.imgLockAspectRatio)),
    imgUseOriginalColor: normalizeToggleFlag(value(item.imgUseOriginalColor) || value(item.useOriginalImageColor)),
    customLargeImageName: normalizeCustomLargeImageName(value(item.customLargeImageName)),
    customLargeImageDataUrl: normalizeCustomLargeImageDataUrl(value(item.customLargeImageDataUrl)),
    customLargeImageAspect: normalizeCustomLargeImageAspect(value(item.customLargeImageAspect)),
    importantFlag: normalizeHasFlag(value(item.importantFlag) || value(item.isImportant)),
    simpleRecordFlag: normalizeCircleDashFlag(value(item.simpleRecordFlag)),
    layerName: normalizeLayerName(value(item.layerName)),
    detail: compactNoSpaceValue(item.detail),
    detailSub: value(item.detailSub),
    layerColor: value(item.layerColor) || getLayerColor(item),
    layerLithology: value(item.layerLithology) || getLayerLithology(item),
    layerFacies: value(item.layerFacies) || composeLayerFacies(item.layerColor, item.layerLithology),
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
  const isLineShape = normalized.largeShapeType === "直線状";
  if (isLineShape) {
    normalized.planeStrikeDirection = "";
    normalized.planeDipDeg = "";
    normalized.planeDipDir8 = "";
  } else {
    if (!value(normalized.planeStrikeDirection) && value(normalized.largeAxisDirection)) {
      normalized.planeStrikeDirection = normalized.largeAxisDirection;
    }
    if (!value(normalized.largeAxisDirection) && value(normalized.planeStrikeDirection)) {
      normalized.largeAxisDirection = normalized.planeStrikeDirection;
    }
    if (!value(normalized.planeDipDeg) && value(normalized.largeAxisPlungeDeg)) {
      normalized.planeDipDeg = normalized.largeAxisPlungeDeg;
    }
    if (!value(normalized.planeDipDir8) && value(normalized.largeAxisPlungeDir8)) {
      normalized.planeDipDir8 = normalized.largeAxisPlungeDir8;
    }
    normalized.largeAxisPlungeDeg = "";
    normalized.largeAxisPlungeDir8 = "";
  }
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
    layerColor: getLayerColor(record),
    layerLithology: getLayerLithology(record),
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
  } catch (error) {
    void recoverFromQuotaError(successMessage, error);
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
      value(site.recorder) ||
      value(site.scribe)
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
    scribe: value(primary.scribe) || value(secondary.scribe),
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
    "色",
    "岩相",
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
    getLayerColor(record),
    getLayerLithology(record),
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
              <tr><th>色</th><td>${escapeHtml(getLayerColor(record))}</td><th>岩相</th><td>${escapeHtml(
                getLayerLithology(record)
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
  const rangeRecords = getRecordsByExportRangeFilters({
    kuwakuValue: ALL_GRIDS_VALUE,
    categoryValue,
    statusValue: "all",
    dateFromRaw: filters.dateFromRaw,
    dateToRaw: filters.dateToRaw,
  });
  const scopedItems = buildPlanPdfItemsAroundGrid(rangeRecords, kuwakuValue);
  if (!scopedItems.length) {
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
      const items = scopedItems.filter((item) => unitValueForSelect(item.record.unit) === unitValue);
      if (!items.length) {
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
        drawables: items.map((item) => item.drawable),
        records: items.map((item) => item.record),
        count: items.length,
      });
    });
  }

  if (detailSelection.enabled) {
    const selectedUnitValue = value(detailSelection.unitValue);
    if (selectedUnitValue) {
      const selectedDetailValues = Array.from(normalizeSelectionSet(detailSelection.detailValues));
      const unitItems = scopedItems.filter((item) => unitValueForSelect(item.record.unit) === selectedUnitValue);
      selectedDetailValues
        .sort((a, b) =>
          detailLabelForSelect(a).localeCompare(detailLabelForSelect(b), "ja", { numeric: true, sensitivity: "base" })
        )
        .forEach((detailValue) => {
          const items = unitItems.filter((item) => detailValueForSelect(item.record.detail) === detailValue);
          if (!items.length) {
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
            drawables: items.map((item) => item.drawable),
            records: items.map((item) => item.record),
            count: items.length,
          });
        });
    }
  }

  if (detailSubSelection.enabled) {
    const selectedUnitValue = value(detailSubSelection.unitValue);
    const selectedDetailValue = value(detailSubSelection.detailValue);
    if (selectedUnitValue && selectedDetailValue) {
      const selectedDetailSubValues = Array.from(normalizeSelectionSet(detailSubSelection.detailSubValues));
      const detailItems = scopedItems.filter(
        (item) =>
          unitValueForSelect(item.record.unit) === selectedUnitValue &&
          detailValueForSelect(item.record.detail) === selectedDetailValue
      );
      selectedDetailSubValues
        .sort((a, b) =>
          detailSubLabelForSelect(a).localeCompare(detailSubLabelForSelect(b), "ja", {
            numeric: true,
            sensitivity: "base",
          })
        )
        .forEach((detailSubValue) => {
          const items = detailItems.filter((item) => detailSubValueForSelect(item.record.detailSub) === detailSubValue);
          if (!items.length) {
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
            drawables: items.map((item) => item.drawable),
            records: items.map((item) => item.record),
            count: items.length,
          });
        });
    }
  }

  return groups;
}

function buildPlanPdfItemsAroundGrid(recordsRaw, targetKuwakuRaw) {
  const target = parseKuwaku(targetKuwakuRaw);
  const targetBlockIndex = blockLabelToIndex(target.block);
  const targetNo = Number(target.no);
  if (!target.block || !Number.isFinite(targetNo)) {
    return [];
  }
  const min = -100;
  const max = PLAN_SIZE_CM + 100;
  return (Array.isArray(recordsRaw) ? recordsRaw : [])
    .map((record) => {
      const recordGrid = parseKuwaku(getRecordKuwaku(record));
      const recordNo = Number(recordGrid.no);
      if (
        recordGrid.headA !== target.headA ||
        recordGrid.headB !== target.headB ||
        !recordGrid.block ||
        !Number.isFinite(recordNo)
      ) {
        return null;
      }
      const drawable = buildPlanDrawable(record);
      if (!drawable || drawable.type === "totalStation") {
        return null;
      }
      const offsetX = (blockLabelToIndex(recordGrid.block) - targetBlockIndex) * PLAN_SIZE_CM;
      const offsetY = (recordNo - targetNo) * PLAN_SIZE_CM;
      const translated = translatePlanDrawable(drawable, offsetX, offsetY);
      const extent = getPlanDrawableExtent(translated);
      if (!extent || extent.maxX < min || extent.minX > max || extent.maxY < min || extent.minY > max) {
        return null;
      }
      return { record, drawable: translated };
    })
    .filter(Boolean);
}

function translatePlanDrawable(drawableRaw, offsetXRaw, offsetYRaw) {
  const drawable = { ...drawableRaw };
  const offsetX = Number(offsetXRaw) || 0;
  const offsetY = Number(offsetYRaw) || 0;
  ["x", "x1", "x2", "left"].forEach((key) => {
    if (Number.isFinite(Number(drawable[key]))) drawable[key] = Number(drawable[key]) + offsetX;
  });
  ["y", "y1", "y2", "top"].forEach((key) => {
    if (Number.isFinite(Number(drawable[key]))) drawable[key] = Number(drawable[key]) + offsetY;
  });
  ["points", "hull"].forEach((key) => {
    if (Array.isArray(drawable[key])) {
      drawable[key] = drawable[key].map((point) => ({
        ...point,
        x: Number(point.x) + offsetX,
        y: Number(point.y) + offsetY,
      }));
    }
  });
  return drawable;
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
      const groupRecords = Array.isArray(group.records) ? group.records : [];
      const teamLabel = collectPlanPdfRecordLabels(groupRecords, (record) => getRecordTeamValue(record));
      const teamLeadLabel = collectPlanPdfRecordLabels(groupRecords, (record) => getRecordTeamLead(record));
      const recorderLabel = collectPlanPdfRecordLabels(groupRecords, (record) => getRecordRecorder(record));
      const memberLabel = collectPlanPdfRecordLabels(groupRecords, (record) => value(record?.discoverer));
      const mapSvg = buildPlanPdfMapSvg(group.drawables, kuwakuValue);
      const recordTable = buildPlanPdfRecordTable(groupRecords);
      return `
        <section class="pdf-page ${index < groups.length - 1 ? "pdf-page-break" : ""}">
          <header class="pdf-header">
            <h1>層準別平面図</h1>
            <div class="pdf-meta">
              <span>区画: ${escapeHtml(kuwakuLabel)}　発掘班: ${escapeHtml(teamLabel || "-")}</span>
              <span>班長: ${escapeHtml(teamLeadLabel || "-")}</span>
              <span>記載係: ${escapeHtml(recorderLabel || "-")}</span>
              <span>班員: ${escapeHtml(memberLabel || "-")}</span>
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
          <div class="pdf-plan-legend">${buildPlanLegendHtml({ includeTotalStation: false })}</div>
          ${recordTable}
        </section>
      `;
    })
    .join("");
}

function collectPlanPdfRecordLabels(recordsRaw, getter) {
  const labels = new Set();
  (Array.isArray(recordsRaw) ? recordsRaw : []).forEach((record) => {
    const label = value(typeof getter === "function" ? getter(record) : "");
    if (label) labels.add(label);
  });
  return Array.from(labels).join("、");
}

function buildPlanPdfRecordTable(recordsRaw) {
  const records = Array.isArray(recordsRaw) ? recordsRaw : [];
  if (!records.length) return "";
  const rows = records
    .map((record, index) => {
      const importantClass = normalizeHasFlag(record?.importantFlag) === "有" ? " pdf-plan-record-important" : "";
      return `<tr class="${importantClass.trim()}">
        <td>${escapeHtml(record.specimenNo || `（${index + 1}）`)}</td>
        <td>${escapeHtml(record.nameMemo || "-")}</td>
      </tr>`;
    })
    .join("");
  return `<section class="pdf-plan-records">
    <h2>表示している化石・遺物</h2>
    <table class="pdf-table pdf-plan-record-table">
      <thead><tr>
        <th>標本番号</th><th>化石・遺物名</th>
      </tr></thead>
      <tbody>${rows}</tbody>
    </table>
  </section>`;
}

function buildPlanPdfMapSvg(drawables, kuwakuRaw) {
  const verticalGrid = [100, 200, 300]
    .map((x) => `<line class="pdf-plan-grid-line" x1="${x}" y1="0" x2="${x}" y2="${PLAN_SIZE_CM}" />`)
    .join("");
  const horizontalGrid = [100, 200, 300]
    .map((y) => `<line class="pdf-plan-grid-line" x1="0" y1="${y}" x2="${PLAN_SIZE_CM}" y2="${y}" />`)
    .join("");
  const pointsSvg = drawables
    .filter((drawable) => drawable?.type !== "totalStation")
    .map((drawable, index) => renderPlanPdfDrawableSvg(drawable, index))
    .join("");
  const viewBox = { minX: -100, minY: -100, width: PLAN_SIZE_CM + 200, height: PLAN_SIZE_CM + 200 };
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
      <svg class="pdf-plan-svg" viewBox="${viewBox.minX} ${viewBox.minY} ${viewBox.width} ${viewBox.height}" aria-label="層準別平面図" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink">
        <rect class="pdf-plan-frame" x="0" y="0" width="${PLAN_SIZE_CM}" height="${PLAN_SIZE_CM}" />
        ${verticalGrid}
        ${horizontalGrid}
        ${pointsSvg}
      </svg>
    </div>
  `;
}

function renderPlanPdfDrawableSvg(drawable, index = 0) {
  const labelX = clamp(drawable.x + 6, -98, PLAN_SIZE_CM + 98);
  const labelY = clamp(drawable.y - 6, -92, PLAN_SIZE_CM + 98);
  let shapeSvg = "";
  if (drawable.type === "line") {
    shapeSvg = `<line class="pdf-plan-shape-line" x1="${drawable.x1}" y1="${drawable.y1}" x2="${drawable.x2}" y2="${drawable.y2}" stroke="${drawable.color}" />`;
  } else if (drawable.type === "multipoint") {
    const points = Array.isArray(drawable.points) ? drawable.points : [];
    const hull = Array.isArray(drawable.hull) ? drawable.hull : [];
    let hullSvg = "";
    if (hull.length >= 3) {
      hullSvg = `<polygon points="${hull.map((point) => `${point.x},${point.y}`).join(" ")}" fill="none" stroke="${drawable.color}" stroke-width="1.8" />`;
    } else if (hull.length === 2) {
      hullSvg = `<line x1="${hull[0].x}" y1="${hull[0].y}" x2="${hull[1].x}" y2="${hull[1].y}" stroke="${drawable.color}" stroke-width="1.8" />`;
    }
    const pointsSvg = points
      .map((point) => `<circle cx="${point.x}" cy="${point.y}" r="3.8" fill="${drawable.color}" stroke="#ffffff" stroke-width="0.9" />`)
      .join("");
    shapeSvg = `${hullSvg}${pointsSvg}`;
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
      -webkit-print-color-adjust: exact;
      print-color-adjust: exact;
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
      overflow: hidden;
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
    .pdf-plan-corner.top-left { top: 34mm; left: 34mm; }
    .pdf-plan-corner.top-right { top: 34mm; right: 34mm; }
    .pdf-plan-corner.bottom-left { bottom: 34mm; left: 34mm; }
    .pdf-plan-corner.bottom-right { bottom: 34mm; right: 34mm; }
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
      -webkit-print-color-adjust: exact;
      print-color-adjust: exact;
    }
    .pdf-plan-records {
      margin-top: 8px;
      page-break-inside: auto;
    }
    .pdf-plan-records h2 {
      margin: 0 0 4px;
      font-size: 13px;
    }
    .pdf-plan-record-table {
      table-layout: fixed;
      margin-top: 0;
      font-size: 10px;
      line-height: 1.35;
      overflow-wrap: anywhere;
    }
    .pdf-plan-record-table thead {
      display: table-header-group;
    }
    .pdf-plan-record-table tr {
      page-break-inside: avoid;
    }
    .pdf-plan-record-table th, .pdf-plan-record-table td {
      padding: 4px 6px;
    }
    .pdf-plan-record-table th:nth-child(1), .pdf-plan-record-table td:nth-child(1) { width: 28%; }
    .pdf-plan-record-table th:nth-child(2), .pdf-plan-record-table td:nth-child(2) { width: 72%; }
    .pdf-plan-record-table tr.pdf-plan-record-important td {
      text-decoration-line: underline;
      text-decoration-color: #dc2626;
      text-decoration-thickness: 2px;
      text-underline-offset: 3px;
      -webkit-print-color-adjust: exact;
      print-color-adjust: exact;
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

  const positionMethod = normalizePositionMethod(recordFormData.get("positionMethod"));
  const siteRequiredFields = [
    ["区画（グリッド）名の1番目", siteSnapshot.kuwakuHeadA],
    ["区画（グリッド）名の2番目", siteSnapshot.kuwakuHeadB],
    ["区画（グリッド）の英字", siteSnapshot.kuwakuBlock],
    ["区画（グリッド）の番号", siteSnapshot.kuwakuNo],
    ["日付", siteSnapshot.date],
    ["発掘班", siteSnapshot.team],
    ["記載者", siteSnapshot.scribe],
  ];
  if (positionMethod !== "totalStation") {
    siteRequiredFields.splice(4, 0, ["レベル高", siteSnapshot.levelHeight]);
  }
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
  if (positionMethod !== "totalStation") {
    recordRequiredFields.splice(6, 0,
      ["レベル読値（上面）", recordFormData.get("levelUpperCm")],
      ["レベル読値（下底）", recordFormData.get("levelLowerCm")]
    );
  } else {
    const tsError = getTotalStationInputError(true);
    if (tsError) {
      return tsError;
    }
  }
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
  const positionMethod = normalizePositionMethod(record.positionMethod);
  if (positionMethod !== "totalStation") {
    if (!value(record.levelUpperCm)) {
      missing.add("levelUpperCm");
    }
    if (!value(record.levelLowerCm)) {
      missing.add("levelLowerCm");
    }
  } else if (!value(record.altitudeDirectM)) {
    missing.add("altitudeDirectM");
  }
  if (!value(record.occurrenceSection)) {
    missing.add("occurrenceSection");
  }
  if (!value(record.occurrenceSketch)) {
    missing.add("occurrenceSketch");
  }
  const planSizeMode = normalizePlanSizeMode(value(record.planSizeMode));
  if (planSizeMode === "複数点") {
    if (!collectPlanMultiPointCoords(record).length) {
      missing.add("multiPoints");
    }
  } else {
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
  }

  if (planSizeMode === "大きなもの") {
    const largeShapeType = normalizeLargeShapeType(value(record.largeShapeType));
    const isImageShape = isLargeShapeImageType(largeShapeType);
    if (!largeShapeType) {
      missing.add("largeShapeType");
    }
    if (largeShapeType === "直線状") {
      if (parseLargeAxisAzimuth(value(record.largeAxisDirection)) == null) {
        missing.add("largeAxisDirection");
      }
      if (!value(record.lineLengthCm)) {
        missing.add("lineLengthCm");
      }
      const plungeDeg = parseLargeAxisPlungeDeg(value(record.largeAxisPlungeDeg));
      if (Number.isFinite(plungeDeg) && plungeDeg > 0 && !normalizeCompass8Direction(value(record.largeAxisPlungeDir8))) {
        missing.add("largeAxisPlungeDir8");
      }
    } else if (largeShapeType === "長方形") {
      if (!value(record.rectSide1Cm)) {
        missing.add("rectSide1Cm");
      }
      if (!value(record.rectSide2Cm)) {
        missing.add("rectSide2Cm");
      }
      const dipDeg = parseLargeAxisPlungeDeg(value(record.planeDipDeg));
      if (Number.isFinite(dipDeg) && dipDeg > 0 && !normalizeCompass8Direction(value(record.planeDipDir8))) {
        missing.add("planeDipDir8");
      }
    } else if (largeShapeType === "楕円") {
      if (!value(record.ellipseLongRadiusCm)) {
        missing.add("ellipseLongRadiusCm");
      }
      if (!value(record.ellipseShortRadiusCm)) {
        missing.add("ellipseShortRadiusCm");
      }
      const dipDeg = parseLargeAxisPlungeDeg(value(record.planeDipDeg));
      if (Number.isFinite(dipDeg) && dipDeg > 0 && !normalizeCompass8Direction(value(record.planeDipDir8))) {
        missing.add("planeDipDir8");
      }
    } else if (isImageShape) {
      const hasFrameSpec =
        parseDistanceToCm(record.imgFrameWidthCm) != null &&
        parseDistanceToCm(record.imgFrameHeightCm) != null &&
        normalizeImageRotationDeg(record.imgRotateDeg) !== "";
      if (!hasFrameSpec) {
        ["imgRotateDeg", "imgFrameWidthCm", "imgFrameHeightCm"].forEach((key) => {
          if (!value(record[key])) {
            missing.add(key);
          }
        });
      }
      if (isCustomLargeShapeType(largeShapeType)) {
        if (!normalizeCustomLargeImageName(record.customLargeImageName)) {
          missing.add("customLargeImageName");
        }
        if (!normalizeCustomLargeImageDataUrl(record.customLargeImageDataUrl)) {
          missing.add("customLargeImageDataUrl");
        }
      }
      const dipDeg = parseLargeAxisPlungeDeg(value(record.planeDipDeg));
      if (Number.isFinite(dipDeg) && dipDeg > 0 && !normalizeCompass8Direction(value(record.planeDipDir8))) {
        missing.add("planeDipDir8");
      }
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

function normalizeCompass8Direction(valueRaw) {
  const raw = value(valueRaw);
  if (!raw) {
    return "";
  }
  const text = raw.toUpperCase().replace(/\s+/g, "");
  if (text === "北" || text === "N") {
    return "北";
  }
  if (text === "北東" || text === "NE" || text === "東北") {
    return "北東";
  }
  if (text === "東" || text === "E") {
    return "東";
  }
  if (text === "南東" || text === "SE" || text === "東南") {
    return "南東";
  }
  if (text === "南" || text === "S") {
    return "南";
  }
  if (text === "南西" || text === "SW" || text === "西南") {
    return "南西";
  }
  if (text === "西" || text === "W") {
    return "西";
  }
  if (text === "北西" || text === "NW" || text === "西北") {
    return "北西";
  }
  return "";
}

function parseCompass8Azimuth(valueRaw) {
  const direction = normalizeCompass8Direction(valueRaw);
  if (!direction) {
    return null;
  }
  if (direction === "北") {
    return 0;
  }
  if (direction === "北東") {
    return 45;
  }
  if (direction === "東") {
    return 90;
  }
  if (direction === "南東") {
    return 135;
  }
  if (direction === "南") {
    return 180;
  }
  if (direction === "南西") {
    return 225;
  }
  if (direction === "西") {
    return 270;
  }
  if (direction === "北西") {
    return 315;
  }
  return null;
}

function normalizePlanSizeMode(valueRaw) {
  const mode = value(valueRaw);
  if (mode === "複数点" || mode === "複数") {
    return "複数点";
  }
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
  if (text === "画像アップロード") {
    text = CUSTOM_LARGE_SHAPE_TYPE;
  }
  if (text === "直線状" || text === "長方形" || text === "楕円" || text === CUSTOM_LARGE_SHAPE_TYPE || largeShapeImagePathMap.has(text)) {
    return text;
  }
  return "";
}

function isLargeShapeImageType(shapeTypeRaw) {
  const shapeType = normalizeLargeShapeType(shapeTypeRaw);
  if (!shapeType) {
    return false;
  }
  return shapeType === CUSTOM_LARGE_SHAPE_TYPE || largeShapeImagePathMap.has(shapeType);
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
  if (shapeType) {
    const fallbackList = LARGE_SHAPE_IMAGE_FALLBACK_PATHS[shapeType] || [];
    const mappedPath = largeShapeImagePathMap.get(shapeType) || "";
    // 形状タイプに対応する正規画像を常に優先する（古い保存パス対策）。
    pushInlineCandidate(mappedPath);
    pushCandidate(mappedPath);
    fallbackList.forEach((pathRaw) => {
      pushInlineCandidate(pathRaw);
      pushCandidate(pathRaw);
    });
  }
  if (hasExplicitPath) {
    // 明示パスは後方互換用の候補として最後に評価する。
    pushInlineCandidate(explicitPath);
    pushCandidate(explicitPath);
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

function normalizeLargeAxisPlungeDeg(valueRaw) {
  const plunge = parseLargeAxisPlungeDeg(valueRaw);
  if (!Number.isFinite(plunge)) {
    return "";
  }
  return Number.isInteger(plunge) ? String(plunge) : String(plunge).replace(/\.?0+$/, "");
}

function normalizePlaneStrikeDirection(valueRaw) {
  return normalizeLargeAxisDirection(valueRaw);
}

function normalizePlaneDipDeg(valueRaw) {
  return normalizeLargeAxisPlungeDeg(valueRaw);
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
  if (group === "tsSetupNsDir") {
    return value(valueRaw) === "南" ? "南" : "北";
  }
  if (group === "tsSetupEwDir") {
    return value(valueRaw) === "西" ? "西" : "東";
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
  if (group === "plungeDir8" || group === "planeDipDir8") {
    return normalizeCompass8Direction(valueRaw);
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
  const useDirectAltitude = normalizeToggleFlag(record?.altitudeInputEnabled) === "1";
  if (useDirectAltitude) {
    const directAltitudeM = parseDistanceToCm(record?.altitudeDirectM);
    return directAltitudeM != null ? directAltitudeM : null;
  }
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

function isLikelyQuotaExceededError(error) {
  if (!error) {
    return false;
  }
  const name = value(error.name).toLowerCase();
  const message = value(error.message).toLowerCase();
  return (
    name.includes("quota") ||
    name.includes("ns_error_dom_quota_reached") ||
    message.includes("quota") ||
    message.includes("storage")
  );
}

async function recoverFromQuotaError(successMessage, causeError = null) {
  if (quotaRecoveryInProgress) {
    showToast("写真データを圧縮中です。数秒待ってから再試行してください");
    return;
  }
  const quotaExceeded = isLikelyQuotaExceededError(causeError);
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
        scheduleCloudSave();
        return;
      } catch (_error) {
        // 次の圧縮段階で再試行
      }
    }
    if (cloudEndpoint) {
      const pushed = await pushStateToCloud({ showToastOnSuccess: false, silentOnError: true });
      if (pushed) {
        if (successMessage) {
          showToast(`${successMessage}（端末容量超過のためクラウド保存）`);
        } else {
          showToast("端末容量超過のためクラウド保存しました");
        }
        return;
      }
    }
    if (quotaExceeded) {
      showToast("端末保存の容量を超えました。画像を減らすか圧縮して再試行してください");
    } else {
      showToast("保存に失敗しました。写真を一部削除して再試行してください");
    }
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
    if (isCustomLargeShapeType(record.largeShapeType)) {
      const imageResult = await recompressImageDataUrl(record.customLargeImageDataUrl, maxLength, quality);
      if (imageResult.changed) {
        record.customLargeImageDataUrl = imageResult.dataUrl;
        changed = true;
      }
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
  if (customLargeImageDataUrlInput) {
    const currentCustomImageResult = await recompressImageDataUrl(customLargeImageDataUrlInput.value, maxLength, quality);
    if (currentCustomImageResult.changed) {
      customLargeImageDataUrlInput.value = currentCustomImageResult.dataUrl;
      await updateCustomLargeImageAspectFromDataUrl(currentCustomImageResult.dataUrl);
      syncLargeShapeImagePreviewForCurrentForm();
      changed = true;
    }
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

async function recompressImageDataUrl(dataUrlRaw, maxLength, quality) {
  const dataUrl = normalizeCustomLargeImageDataUrl(dataUrlRaw);
  if (!dataUrl) {
    return { dataUrl, changed: false };
  }
  const matchedMimeType = dataUrl.match(/^data:([^;,]+)/i);
  const sourceMimeType = normalizeImageMimeType(matchedMimeType?.[1]);
  const outputMimeType = sourceMimeType === "image/png" ? "image/png" : "image/jpeg";
  try {
    const recompressed = await resizeDataUrlImage(dataUrl, maxLength, quality, outputMimeType);
    if (recompressed && recompressed.length < dataUrl.length) {
      return { dataUrl: recompressed, changed: true };
    }
  } catch (_error) {
    // 元画像を維持
  }
  return { dataUrl, changed: false };
}

const SUPPORTED_UPLOAD_IMAGE_MIME_TYPES = new Set(["image/jpeg", "image/png"]);

function normalizeImageMimeType(mimeTypeRaw) {
  const mimeType = value(mimeTypeRaw).toLowerCase();
  if (!mimeType) {
    return "";
  }
  if (mimeType === "image/jpg") {
    return "image/jpeg";
  }
  return mimeType;
}

function inferImageMimeTypeFromFileName(fileNameRaw) {
  const fileName = value(fileNameRaw).toLowerCase();
  if (fileName.endsWith(".jpg") || fileName.endsWith(".jpeg")) {
    return "image/jpeg";
  }
  if (fileName.endsWith(".png")) {
    return "image/png";
  }
  return "";
}

function resolveImageMimeType(file, preferredMimeTypeRaw = "") {
  const preferredMimeType = normalizeImageMimeType(preferredMimeTypeRaw);
  if (SUPPORTED_UPLOAD_IMAGE_MIME_TYPES.has(preferredMimeType)) {
    return preferredMimeType;
  }
  const fileMimeType = normalizeImageMimeType(file?.type);
  if (SUPPORTED_UPLOAD_IMAGE_MIME_TYPES.has(fileMimeType)) {
    return fileMimeType;
  }
  const fromName = inferImageMimeTypeFromFileName(file?.name);
  if (SUPPORTED_UPLOAD_IMAGE_MIME_TYPES.has(fromName)) {
    return fromName;
  }
  return "";
}

function normalizeImageDataUrlForUpload(dataUrlRaw, file, preferredMimeTypeRaw = "") {
  const dataUrl = value(dataUrlRaw);
  if (!dataUrl || !dataUrl.startsWith("data:")) {
    return "";
  }
  if (dataUrl.startsWith("data:image/")) {
    return dataUrl;
  }
  const resolvedMimeType = resolveImageMimeType(file, preferredMimeTypeRaw);
  if (!resolvedMimeType) {
    return "";
  }
  if (/^data:;base64,/i.test(dataUrl)) {
    return dataUrl.replace(/^data:;base64,/i, `data:${resolvedMimeType};base64,`);
  }
  return dataUrl.replace(/^data:[^;,]+/i, `data:${resolvedMimeType}`);
}

function resizeDataUrlImage(dataUrl, maxLength, quality = 0.72, mimeType = "image/jpeg") {
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
      resolve(canvas.toDataURL(mimeType, quality));
    };
    image.onerror = () => reject(new Error("Failed to load image"));
    image.src = dataUrl;
  });
}

function resizeImage(file, maxLength, quality = 0.72, mimeType = "image/jpeg") {
  return new Promise((resolve, reject) => {
    const outputMimeType = resolveImageMimeType(file, mimeType) || "image/jpeg";
    const reader = new FileReader();
    reader.onload = () => {
      const normalizedDataUrl = normalizeImageDataUrlForUpload(reader.result, file, outputMimeType);
      if (!normalizedDataUrl) {
        reject(new Error("Unsupported image type"));
        return;
      }
      resizeDataUrlImage(normalizedDataUrl, maxLength, quality, outputMimeType).then(resolve).catch(reject);
    };
    reader.onerror = () => reject(new Error("Failed to read file"));
    reader.readAsDataURL(file);
  });
}

function readFileAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const sourceFile = file instanceof File ? file : null;
    if (!sourceFile) {
      reject(new Error("Invalid file"));
      return;
    }
    const reader = new FileReader();
    reader.onload = () => {
      resolve(String(reader.result || ""));
    };
    reader.onerror = () => reject(new Error("Failed to read file"));
    reader.readAsDataURL(sourceFile);
  });
}

async function loadImageFileDataUrlWithFallback(
  file,
  { maxLength = 1280, quality = 0.72, mimeType = "image/jpeg", allowOriginalFallback = true } = {}
) {
  const sourceFile = file instanceof File ? file : null;
  if (!sourceFile) {
    throw new Error("Invalid file");
  }
  const outputMimeType = resolveImageMimeType(sourceFile, mimeType) || normalizeImageMimeType(mimeType) || "image/jpeg";
  try {
    return await resizeImage(sourceFile, maxLength, quality, outputMimeType);
  } catch (resizeError) {
    if (!allowOriginalFallback) {
      throw resizeError;
    }
    const originalDataUrl = normalizeImageDataUrlForUpload(await readFileAsDataUrl(sourceFile), sourceFile, outputMimeType);
    if (!originalDataUrl) {
      throw resizeError;
    }
    return originalDataUrl;
  }
}

function normalizePositionMethod(valueRaw) {
  return value(valueRaw) === "totalStation" ? "totalStation" : "grid";
}

function normalizeTotalStationPointName(valueRaw) {
  const normalized = value(valueRaw)
    .replace(/[！-～]/g, (char) => String.fromCharCode(char.charCodeAt(0) - 0xfee0))
    .replace(/[－—–ー―]/g, "-")
    .replace(/\s+/g, "")
    .toUpperCase();
  return normalized.replace(/^[1IⅠ](?=-)/, "Ⅰ");
}

function getTotalStationGridReferencePoints() {
  const source = Array.isArray(window.NOJIRI_GRID_REFERENCE_POINTS) ? window.NOJIRI_GRID_REFERENCE_POINTS : [];
  return source
    .map((point) => ({
      name: normalizeTotalStationPointName(point?.name),
      x: Number(point?.x),
      y: Number(point?.y),
    }))
    .filter((point) => point.name && Number.isFinite(point.x) && Number.isFinite(point.y));
}

function findTotalStationGridReferencePoint(nameRaw) {
  const name = normalizeTotalStationPointName(nameRaw);
  return getTotalStationGridReferencePoints().find((point) => point.name === name) || null;
}

function applyTotalStationGridReferencePoint(kind, nameRaw, { showMessage = false } = {}) {
  if (!recordForm) return false;
  const point = findTotalStationGridReferencePoint(nameRaw);
  if (!point) return false;
  const isStation = kind === "station";
  const nameField = recordForm.elements[isStation ? "tsStationPeg" : "tsBacksightPeg"];
  const xField = recordForm.elements[isStation ? "tsStationXNorthM" : "tsBacksightXNorthM"];
  const yField = recordForm.elements[isStation ? "tsStationYEastM" : "tsBacksightYEastM"];
  if (nameField instanceof HTMLInputElement) nameField.value = point.name;
  if (xField instanceof HTMLInputElement) xField.value = String(point.x);
  if (yField instanceof HTMLInputElement) yField.value = String(point.y);
  const select = isStation ? tsStationPointSelect : tsBacksightPointSelect;
  if (select instanceof HTMLSelectElement) select.value = point.name;
  applyTotalStationPosition();
  if (showMessage) showToast(`${isStation ? "設置点" : "後視点"} ${point.name} のX・Yを入力しました`);
  return true;
}

function syncTotalStationGridPointSelectors() {
  if (!recordForm) return;
  const stationName = normalizeTotalStationPointName(recordForm.elements.tsStationPeg?.value);
  const backsightName = normalizeTotalStationPointName(recordForm.elements.tsBacksightPeg?.value);
  if (tsStationPointSelect instanceof HTMLSelectElement) {
    tsStationPointSelect.value = findTotalStationGridReferencePoint(stationName) ? stationName : "";
  }
  if (tsBacksightPointSelect instanceof HTMLSelectElement) {
    tsBacksightPointSelect.value = findTotalStationGridReferencePoint(backsightName) ? backsightName : "";
  }
}

function initializeTotalStationGridPointSelectors() {
  const points = getTotalStationGridReferencePoints();
  const options = points
    .map((point) => `<option value="${escapeHtml(point.name)}">${escapeHtml(point.name)}（X ${point.x}、Y ${point.y}）</option>`)
    .join("");
  [tsStationPointSelect, tsBacksightPointSelect].forEach((select) => {
    if (select instanceof HTMLSelectElement) select.insertAdjacentHTML("beforeend", options);
  });
  tsStationPointSelect?.addEventListener("change", () => {
    if (value(tsStationPointSelect.value)) applyTotalStationGridReferencePoint("station", tsStationPointSelect.value, { showMessage: true });
  });
  tsBacksightPointSelect?.addEventListener("change", () => {
    if (value(tsBacksightPointSelect.value)) applyTotalStationGridReferencePoint("backsight", tsBacksightPointSelect.value, { showMessage: true });
  });
  const stationNameField = recordForm?.elements?.tsStationPeg;
  const backsightNameField = recordForm?.elements?.tsBacksightPeg;
  stationNameField?.addEventListener("input", () => {
    const normalized = normalizeTotalStationPointName(stationNameField.value);
    if (normalized !== stationNameField.value) stationNameField.value = normalized;
    if (findTotalStationGridReferencePoint(normalized)) applyTotalStationGridReferencePoint("station", normalized);
  });
  backsightNameField?.addEventListener("input", () => {
    const normalized = normalizeTotalStationPointName(backsightNameField.value);
    if (normalized !== backsightNameField.value) backsightNameField.value = normalized;
    if (findTotalStationGridReferencePoint(normalized)) applyTotalStationGridReferencePoint("backsight", normalized);
  });
  stationNameField?.addEventListener("change", () => {
    const normalized = normalizeTotalStationPointName(stationNameField.value);
    stationNameField.value = normalized;
    if (!applyTotalStationGridReferencePoint("station", normalized)) syncTotalStationGridPointSelectors();
  });
  backsightNameField?.addEventListener("change", () => {
    const normalized = normalizeTotalStationPointName(backsightNameField.value);
    backsightNameField.value = normalized;
    if (!applyTotalStationGridReferencePoint("backsight", normalized)) syncTotalStationGridPointSelectors();
  });
}

function setPositionMeasurementFields(record = {}) {
  if (!recordForm) {
    return;
  }
  const method = normalizePositionMethod(record.positionMethod);
  const radio = recordForm.querySelector(`input[name="positionMethod"][value="${method}"]`);
  if (radio instanceof HTMLInputElement) {
    radio.checked = true;
  }
  [
    "tsStationPeg", "tsStationXNorthM", "tsStationYEastM", "tsStationAltitudeM",
    "tsBacksightPeg", "tsBacksightXNorthM", "tsBacksightYEastM", "tsBacksightAltitudeM",
    "tsInstrumentHeightM", "tsTargetHeightM", "tsPointXNorthM", "tsPointYEastM", "tsPointAltitudeM",
    "tsSlopeDistanceM", "tsInclinationDeg", "tsInclinationMin", "tsInclinationSec",
    "tsDirectionDeg", "tsDirectionMin", "tsDirectionSec",
  ].forEach((name) => {
    const field = recordForm.elements[name];
    if (field instanceof HTMLInputElement || field instanceof HTMLSelectElement) {
      const isLegacyCoordinateField = !value(record.tsCoordinateConvention) && [
        "tsStationXNorthM", "tsStationYEastM", "tsBacksightXNorthM", "tsBacksightYEastM",
        "tsPointXNorthM", "tsPointYEastM",
      ].includes(name);
      const rawValue = value(record[name]);
      const numericValue = isLegacyCoordinateField ? parseTotalStationNumber(rawValue) : null;
      field.value = numericValue == null ? rawValue || field.value : String(-numericValue);
    }
  });
  const stationPegField = recordForm.elements.tsStationPeg;
  const backsightPegField = recordForm.elements.tsBacksightPeg;
  if (stationPegField instanceof HTMLInputElement && value(stationPegField.value)) {
    stationPegField.value = normalizeTotalStationPointName(stationPegField.value);
  }
  if (backsightPegField instanceof HTMLInputElement && value(backsightPegField.value)) {
    backsightPegField.value = normalizeTotalStationPointName(backsightPegField.value);
  }
  if (stationPegField instanceof HTMLInputElement && !value(stationPegField.value) && value(record.tsBacksightPeg)) {
    stationPegField.value = normalizeTotalStationPointName(record.tsBacksightPeg);
  }
  if (value(record.tsObservationMode) !== "polar" && value(record.tsPointCoordinateMode) === "stationOffsetNorthWest") {
    const formerNorthOffset = parseTotalStationNumber(record.tsPointXNorthM);
    const southOffsetField = recordForm.elements.tsPointXNorthM;
    if (formerNorthOffset != null && southOffsetField instanceof HTMLInputElement) {
      southOffsetField.value = String(-formerNorthOffset);
    }
  } else if (value(record.tsObservationMode) !== "polar" && value(record.tsPointCoordinateMode) !== "stationOffsetSouthWest") {
    const legacyConvention = !value(record.tsCoordinateConvention);
    const stationXRaw = parseTotalStationNumber(record.tsStationXNorthM);
    const stationYRaw = parseTotalStationNumber(record.tsStationYEastM);
    const pointXRaw = parseTotalStationNumber(record.tsPointXNorthM);
    const pointYRaw = parseTotalStationNumber(record.tsPointYEastM);
    if ([stationXRaw, stationYRaw, pointXRaw, pointYRaw].every((number) => number != null)) {
      const stationX = legacyConvention ? -stationXRaw : stationXRaw;
      const stationY = legacyConvention ? -stationYRaw : stationYRaw;
      const pointX = legacyConvention ? -pointXRaw : pointXRaw;
      const pointY = legacyConvention ? -pointYRaw : pointYRaw;
      const southOffsetField = recordForm.elements.tsPointXNorthM;
      const westOffsetField = recordForm.elements.tsPointYEastM;
      if (southOffsetField instanceof HTMLInputElement) southOffsetField.value = String(pointX - stationX);
      if (westOffsetField instanceof HTMLInputElement) westOffsetField.value = String(pointY - stationY);
    }
  }
  const observationMode = value(record.tsObservationMode) === "polar" ? "polar" : "coordinate";
  const observationRadio = recordForm.querySelector(`input[name="tsObservationMode"][value="${observationMode}"]`);
  if (observationRadio instanceof HTMLInputElement) observationRadio.checked = true;
  syncTotalStationGridPointSelectors();
  syncPositionMeasurementUi();
}

function currentInputKuwakuParts() {
  const isEdit = getActiveTabId() === "edit-tab";
  const block = isEdit ? value(editKuwakuBlockInput?.value) : value(siteForm?.elements?.kuwakuBlock?.value);
  const no = isEdit ? value(editKuwakuNoInput?.value) : value(siteForm?.elements?.kuwakuNo?.value);
  return { block: block.toUpperCase(), no };
}

function currentInputGridReferenceName() {
  const isEdit = getActiveTabId() === "edit-tab";
  const headB = isEdit ? value(editKuwakuHeadBInput?.value) : value(siteForm?.elements?.kuwakuHeadB?.value);
  const { block, no } = currentInputKuwakuParts();
  return normalizeTotalStationPointName(`${headB}-${block}-${no}`);
}

function parseTotalStationPeg(valueRaw) {
  const parts = normalizeTotalStationPointName(valueRaw).split("-").map((part) => part.trim()).filter(Boolean);
  if (parts.length < 2) {
    return null;
  }
  const block = parts[parts.length - 2];
  const no = parts[parts.length - 1];
  if (!/^[A-Z]+$/.test(block) || !/^-?\d+$/.test(no)) {
    return null;
  }
  return { block, no };
}

function parseTotalStationNumber(valueRaw) {
  const normalized = value(valueRaw)
    .replace(/[０-９]/g, (char) => String.fromCharCode(char.charCodeAt(0) - 0xfee0))
    .replace(/[，、]/g, ".")
    .replace(/[＋]/g, "+")
    .replace(/[−ー―]/g, "-");
  if (!normalized) {
    return null;
  }
  const number = Number(normalized);
  return Number.isFinite(number) ? number : null;
}

function readTotalStationSetupFromForm() {
  if (!recordForm) {
    return null;
  }
  const data = new FormData(recordForm);
  return {
    tsCoordinateConvention: "southWestPositive",
    tsStationPeg: normalizeTotalStationPointName(data.get("tsStationPeg")),
    tsStationXNorthM: value(data.get("tsStationXNorthM")),
    tsStationYEastM: value(data.get("tsStationYEastM")),
    tsStationAltitudeM: value(data.get("tsStationAltitudeM")),
    tsBacksightPeg: normalizeTotalStationPointName(data.get("tsBacksightPeg")),
    tsBacksightXNorthM: value(data.get("tsBacksightXNorthM")),
    tsBacksightYEastM: value(data.get("tsBacksightYEastM")),
    tsBacksightAltitudeM: value(data.get("tsBacksightAltitudeM")),
    tsInstrumentHeightM: value(data.get("tsInstrumentHeightM")),
    tsTargetHeightM: value(data.get("tsTargetHeightM")),
  };
}

function restoreSavedTotalStationSetup() {
  if (!recordForm) {
    return false;
  }
  let saved = null;
  try {
    saved = JSON.parse(localStorage.getItem(TOTAL_STATION_SETUP_KEY) || "null");
  } catch (_error) {
    return false;
  }
  if (!saved || typeof saved !== "object") {
    return false;
  }
  if (!value(saved.tsStationXNorthM) || !value(saved.tsBacksightPeg)) {
    return false;
  }
  ["tsStationPeg", "tsStationXNorthM", "tsStationYEastM", "tsStationAltitudeM", "tsBacksightPeg",
    "tsBacksightXNorthM", "tsBacksightYEastM", "tsBacksightAltitudeM", "tsInstrumentHeightM", "tsTargetHeightM"].forEach((name) => {
    const field = recordForm.elements[name];
    if (field instanceof HTMLInputElement) {
      const isLegacyCoordinateField = !value(saved.tsCoordinateConvention) && [
        "tsStationXNorthM", "tsStationYEastM", "tsBacksightXNorthM", "tsBacksightYEastM",
      ].includes(name);
      const rawValue = value(saved[name]);
      const numericValue = isLegacyCoordinateField ? parseTotalStationNumber(rawValue) : null;
      field.value = numericValue == null ? rawValue || field.value : String(-numericValue);
    }
  });
  const stationPegField = recordForm.elements.tsStationPeg;
  const backsightPegField = recordForm.elements.tsBacksightPeg;
  if (stationPegField instanceof HTMLInputElement && value(stationPegField.value)) {
    stationPegField.value = normalizeTotalStationPointName(stationPegField.value);
  }
  if (backsightPegField instanceof HTMLInputElement && value(backsightPegField.value)) {
    backsightPegField.value = normalizeTotalStationPointName(backsightPegField.value);
  }
  if (stationPegField instanceof HTMLInputElement && !value(stationPegField.value)) {
    stationPegField.value = normalizeTotalStationPointName(saved.tsBacksightPeg);
  }
  const status = document.getElementById("ts-setup-save-status");
  if (status) {
    status.textContent = `保存済み：設置点 ${value(stationPegField?.value).toUpperCase()}／後視点 ${value(saved.tsBacksightPeg).toUpperCase()}`;
  }
  syncTotalStationGridPointSelectors();
  return true;
}

function saveTotalStationSetup() {
  const setup = readTotalStationSetupFromForm();
  if (!setup || !value(setup.tsStationPeg)) {
    showToast("設置点の杭（点）名称を入力してください");
    return;
  }
  if (!value(setup.tsBacksightPeg)) {
    showToast("後視点の杭（点）名称を入力してください");
    return;
  }
  const requiredNumbers = [
    ["設置点x", setup.tsStationXNorthM], ["設置点y", setup.tsStationYEastM],
    ["設置点標高", setup.tsStationAltitudeM], ["後視点x", setup.tsBacksightXNorthM],
    ["後視点y", setup.tsBacksightYEastM], ["後視点標高", setup.tsBacksightAltitudeM],
    ["機械高", setup.tsInstrumentHeightM], ["目標高", setup.tsTargetHeightM],
  ];
  for (const [label, raw] of requiredNumbers) {
    if (parseTotalStationNumber(raw) == null) {
      showToast(`${label}を数値で入力してください`);
      return;
    }
  }
  try {
    localStorage.setItem(TOTAL_STATION_SETUP_KEY, JSON.stringify(setup));
  } catch (_error) {
    showToast("トータルステーション設置位置を保存できませんでした");
    return;
  }
  const status = document.getElementById("ts-setup-save-status");
  if (status) {
    status.textContent = `保存済み：設置点 ${setup.tsStationPeg}／後視点 ${setup.tsBacksightPeg}`;
  }
  showToast("トータルステーション設置位置を保存しました");
}

function getTotalStationInputError(requireAltitude = false) {
  if (!recordForm) {
    return "入力情報を取得できませんでした";
  }
  const data = new FormData(recordForm);
  const grid = currentInputKuwakuParts();
  if (!value(data.get("tsStationPeg"))) {
    return "設置点の杭（点）名称を入力してください";
  }
  if (!value(data.get("tsBacksightPeg"))) {
    return "後視点の杭（点）名称を入力してください";
  }
  if (!/^[A-Z]+$/.test(grid.block) || !/^-?\d+$/.test(grid.no)) {
    return "先に区画（グリッド）の英字と番号を入力してください";
  }
  const requiredNumbers = [
    ["設置点x", data.get("tsStationXNorthM")], ["設置点y", data.get("tsStationYEastM")],
    ["設置点標高", data.get("tsStationAltitudeM")], ["後視点x", data.get("tsBacksightXNorthM")],
    ["後視点y", data.get("tsBacksightYEastM")], ["後視点標高", data.get("tsBacksightAltitudeM")],
    ["機械高", data.get("tsInstrumentHeightM")], ["目標高", data.get("tsTargetHeightM")],
  ];
  for (const [label, raw] of requiredNumbers) {
    if (parseTotalStationNumber(raw) == null) {
      return `${label}を数値で入力してください`;
    }
  }
  const mode = value(data.get("tsObservationMode")) === "polar" ? "polar" : "coordinate";
  const observationFields = mode === "coordinate"
    ? [["設置点から南への距離", data.get("tsPointXNorthM")], ["設置点から西への距離", data.get("tsPointYEastM")], ["標高", data.get("tsPointAltitudeM")]]
    : [["斜距離", data.get("tsSlopeDistanceM")], ["傾斜度", data.get("tsInclinationDeg")], ["傾斜分", data.get("tsInclinationMin")], ["傾斜秒", data.get("tsInclinationSec")], ["方向角度", data.get("tsDirectionDeg")], ["方向角分", data.get("tsDirectionMin")], ["方向角秒", data.get("tsDirectionSec")]];
  for (const [label, raw] of observationFields) {
    if (parseTotalStationNumber(raw) == null) return `${label}を数値で入力してください`;
  }
  return "";
}

function dmsToDegrees(degRaw, minRaw, secRaw) {
  const deg = parseTotalStationNumber(degRaw);
  const min = parseTotalStationNumber(minRaw);
  const sec = parseTotalStationNumber(secRaw);
  if ([deg, min, sec].some((number) => number == null) || min < 0 || min >= 60 || sec < 0 || sec >= 60) return null;
  const sign = deg < 0 || /^\s*-/.test(value(degRaw)) ? -1 : 1;
  return sign * (Math.abs(deg) + min / 60 + sec / 3600);
}

function calculateTotalStationPosition() {
  if (!recordForm || normalizePositionMethod(new FormData(recordForm).get("positionMethod")) !== "totalStation") {
    return null;
  }
  const data = new FormData(recordForm);
  const stationPegRaw = normalizeTotalStationPointName(data.get("tsStationPeg") || data.get("tsBacksightPeg"));
  const peg = parseTotalStationPeg(stationPegRaw);
  const grid = currentInputKuwakuParts();
  const stationX = parseTotalStationNumber(data.get("tsStationXNorthM"));
  const stationY = parseTotalStationNumber(data.get("tsStationYEastM"));
  const stationZ = parseTotalStationNumber(data.get("tsStationAltitudeM"));
  const backX = parseTotalStationNumber(data.get("tsBacksightXNorthM"));
  const backY = parseTotalStationNumber(data.get("tsBacksightYEastM"));
  const instrumentHeight = parseTotalStationNumber(data.get("tsInstrumentHeightM"));
  const targetHeight = parseTotalStationNumber(data.get("tsTargetHeightM"));
  if (!stationPegRaw || !/^[A-Z]+$/.test(grid.block) || !/^-?\d+$/.test(grid.no) || [stationX, stationY, stationZ, backX, backY, instrumentHeight, targetHeight].some((number) => number == null)) return null;
  const mode = value(data.get("tsObservationMode")) === "polar" ? "polar" : "coordinate";
  let pointX = parseTotalStationNumber(data.get("tsPointXNorthM"));
  let pointY = parseTotalStationNumber(data.get("tsPointYEastM"));
  let pointZ = parseTotalStationNumber(data.get("tsPointAltitudeM"));
  if (mode === "coordinate" && pointX != null && pointY != null) {
    const southFromStationM = pointX;
    const westFromStationM = pointY;
    pointX = stationX + southFromStationM;
    pointY = stationY + westFromStationM;
  } else if (mode === "polar") {
    const distance = parseTotalStationNumber(data.get("tsSlopeDistanceM"));
    const inclination = dmsToDegrees(data.get("tsInclinationDeg"), data.get("tsInclinationMin"), data.get("tsInclinationSec"));
    const direction = dmsToDegrees(data.get("tsDirectionDeg"), data.get("tsDirectionMin"), data.get("tsDirectionSec"));
    if ([distance, inclination, direction].some((number) => number == null) || distance < 0 || (stationX === backX && stationY === backY)) return null;
    const inclinationRad = inclination * Math.PI / 180;
    const baseAzimuth = Math.atan2(backY - stationY, backX - stationX);
    const azimuth = baseAzimuth + direction * Math.PI / 180;
    const horizontal = distance * Math.cos(inclinationRad);
    pointX = stationX + horizontal * Math.cos(azimuth);
    pointY = stationY + horizontal * Math.sin(azimuth);
    pointZ = stationZ + instrumentHeight + distance * Math.sin(inclinationRad) - targetHeight;
  }
  if ([pointX, pointY, pointZ].some((number) => number == null)) return null;
  const specimenSouthFromStationM = pointX - stationX;
  const specimenWestFromStationM = pointY - stationY;
  const specimenEastFromPegCm = -specimenWestFromStationM * 100;
  const specimenNorthFromPegCm = -specimenSouthFromStationM * 100;
  const gridPegLabel = currentInputGridReferenceName();
  const gridReference = findTotalStationGridReferencePoint(gridPegLabel);
  let stationPlanX;
  let stationPlanY;
  let xPlanCm;
  let yPlanCm;
  if (gridReference) {
    stationPlanX = -(stationY - gridReference.y) * 100;
    stationPlanY = (stationX - gridReference.x) * 100;
    xPlanCm = -(pointY - gridReference.y) * 100;
    yPlanCm = (pointX - gridReference.x) * 100;
  } else if (peg) {
    const pegX = (blockLabelToIndex(peg.block) - blockLabelToIndex(grid.block)) * PLAN_SIZE_CM;
    const pegY = (Number(peg.no) - Number(grid.no)) * PLAN_SIZE_CM;
    stationPlanX = pegX;
    stationPlanY = pegY;
    xPlanCm = pegX + specimenEastFromPegCm;
    yPlanCm = pegY - specimenNorthFromPegCm;
  } else {
    return null;
  }
  return {
    pegLabel: stationPegRaw.toUpperCase(),
    specimenEastFromPegCm,
    specimenNorthFromPegCm,
    gridPegLabel,
    specimenEastFromGridPegCm: xPlanCm,
    specimenSouthFromGridPegCm: yPlanCm,
    xPlanCm,
    yPlanCm,
    stationPlanX, stationPlanY, pointX, pointY,
    specimenSouthFromStationM, specimenWestFromStationM,
    altitudeM: pointZ,
  };
}

function applyTotalStationPosition() {
  updateMultiPointCalculationResults();
  const result = calculateTotalStationPosition();
  const resultEl = document.getElementById("total-station-result");
  if (!result) {
    if (resultEl) {
      resultEl.textContent = getTotalStationInputError() || "入力値を確認してください。";
      resultEl.classList.remove("total-station-result-ok");
    }
    return;
  }
  nsDirInput.value = "北から";
  ewDirInput.value = "西から";
  recordForm.elements.nsCm.value = formatLengthInputValue(result.yPlanCm);
  recordForm.elements.ewCm.value = formatLengthInputValue(result.xPlanCm);
  const altitudeToggle = recordForm.elements.altitudeInputEnabled;
  const altitudeField = recordForm.elements.altitudeDirectM;
  if (result.altitudeM != null) {
    if (altitudeToggle instanceof HTMLInputElement) {
      altitudeToggle.checked = true;
    }
    if (altitudeField instanceof HTMLInputElement) {
      altitudeField.value = String(Number(result.altitudeM.toFixed(4)));
    }
  }
  syncDirectionTabsFromForm();
  syncAltitudeDirectInputUi();
  if (resultEl) {
    resultEl.textContent = formatGridEdgeCalculationResult({
      x: result.xPlanCm,
      y: result.yPlanCm,
      z: result.altitudeM,
    });
    resultEl.classList.add("total-station-result-ok");
  }
}

function syncPositionMeasurementUi() {
  if (!recordForm) {
    return;
  }
  const method = normalizePositionMethod(new FormData(recordForm).get("positionMethod"));
  document.getElementById("grid-position-fields")?.classList.toggle("hidden", method !== "grid");
  document.getElementById("total-station-section")?.classList.toggle("hidden", method !== "totalStation");
  document.getElementById("ts-specimen-position-section")?.classList.toggle("hidden", method !== "totalStation");
  const observationMode = value(new FormData(recordForm).get("tsObservationMode")) === "polar" ? "polar" : "coordinate";
  document.getElementById("ts-coordinate-fields")?.classList.toggle("hidden", observationMode !== "coordinate");
  document.getElementById("ts-polar-fields")?.classList.toggle("hidden", observationMode !== "polar");
  document.getElementById("level-reading-section")?.classList.toggle("hidden", method === "totalStation");
  siteForm?.querySelector(".site-level")?.classList.toggle("hidden", method === "totalStation");
  syncMultiPointCoordinateModeUi();
  if (method === "totalStation") {
    applyTotalStationPosition();
  }
}

if (recordForm) {
  document.getElementById("ts-setup-save-btn")?.addEventListener("click", saveTotalStationSetup);
  recordForm.addEventListener("change", (event) => {
    if (
      event.target instanceof Element &&
      (event.target.matches('[name="positionMethod"]') ||
        event.target.closest("#total-station-section, #ts-specimen-position-section"))
    ) {
      syncPositionMeasurementUi();
      applyTotalStationPosition();
    }
  });
  recordForm.addEventListener("input", (event) => {
    if (
      event.target instanceof Element &&
      event.target.closest("#total-station-section, #ts-specimen-position-section")
    ) {
      applyTotalStationPosition();
    }
  });
  syncPositionMeasurementUi();
}

if ("serviceWorker" in navigator && location.protocol !== "file:") {
  const offlineStatus = document.getElementById("offline-status");
  navigator.serviceWorker.register("./sw.js")
    .then(() => navigator.serviceWorker.ready)
    .then(() => {
      if (offlineStatus) {
        offlineStatus.textContent = "オフライン利用準備完了";
        offlineStatus.classList.add("ready");
      }
    })
    .catch(() => {
      if (offlineStatus) {
        offlineStatus.textContent = "オフライン利用の準備に失敗";
        offlineStatus.classList.add("error");
      }
    });
}
