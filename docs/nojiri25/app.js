if (!String.prototype.replaceAll) {
  String.prototype.replaceAll = function replaceAllPolyfill(searchValue, replaceValue) {
    var source = String(this);
    if (searchValue instanceof RegExp) {
      if (!searchValue.global) {
        throw new TypeError("String.prototype.replaceAll called with a non-global RegExp");
      }
      return source.replace(searchValue, replaceValue);
    }
    return source.split(String(searchValue)).join(String(replaceValue));
  };
}

function _toConsumableArray(r) { return _arrayWithoutHoles(r) || _iterableToArray(r) || _unsupportedIterableToArray(r) || _nonIterableSpread(); }
function _nonIterableSpread() { throw new TypeError("Invalid attempt to spread non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); }
function _iterableToArray(r) { if ("undefined" != typeof Symbol && null != r[Symbol.iterator] || null != r["@@iterator"]) return Array.from(r); }
function _arrayWithoutHoles(r) { if (Array.isArray(r)) return _arrayLikeToArray(r); }
function _createForOfIteratorHelper(r, e) { var t = "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"]; if (!t) { if (Array.isArray(r) || (t = _unsupportedIterableToArray(r)) || e && r && "number" == typeof r.length) { t && (r = t); var _n = 0, F = function F() {}; return { s: F, n: function n() { return _n >= r.length ? { done: !0 } : { done: !1, value: r[_n++] }; }, e: function e(r) { throw r; }, f: F }; } throw new TypeError("Invalid attempt to iterate non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); } var o, a = !0, u = !1; return { s: function s() { t = t.call(r); }, n: function n() { var r = t.next(); return a = r.done, r; }, e: function e(r) { u = !0, o = r; }, f: function f() { try { a || null == t["return"] || t["return"](); } finally { if (u) throw o; } } }; }
function _slicedToArray(r, e) { return _arrayWithHoles(r) || _iterableToArrayLimit(r, e) || _unsupportedIterableToArray(r, e) || _nonIterableRest(); }
function _nonIterableRest() { throw new TypeError("Invalid attempt to destructure non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); }
function _unsupportedIterableToArray(r, a) { if (r) { if ("string" == typeof r) return _arrayLikeToArray(r, a); var t = {}.toString.call(r).slice(8, -1); return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? _arrayLikeToArray(r, a) : void 0; } }
function _arrayLikeToArray(r, a) { (null == a || a > r.length) && (a = r.length); for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e]; return n; }
function _iterableToArrayLimit(r, l) { var t = null == r ? null : "undefined" != typeof Symbol && r[Symbol.iterator] || r["@@iterator"]; if (null != t) { var e, n, i, u, a = [], f = !0, o = !1; try { if (i = (t = t.call(r)).next, 0 === l) { if (Object(t) !== t) return; f = !1; } else for (; !(f = (e = i.call(t)).done) && (a.push(e.value), a.length !== l); f = !0); } catch (r) { o = !0, n = r; } finally { try { if (!f && null != t["return"] && (u = t["return"](), Object(u) !== u)) return; } finally { if (o) throw n; } } return a; } }
function _arrayWithHoles(r) { if (Array.isArray(r)) return r; }
function ownKeys(e, r) { var t = Object.keys(e); if (Object.getOwnPropertySymbols) { var o = Object.getOwnPropertySymbols(e); r && (o = o.filter(function (r) { return Object.getOwnPropertyDescriptor(e, r).enumerable; })), t.push.apply(t, o); } return t; }
function _objectSpread(e) { for (var r = 1; r < arguments.length; r++) { var t = null != arguments[r] ? arguments[r] : {}; r % 2 ? ownKeys(Object(t), !0).forEach(function (r) { _defineProperty(e, r, t[r]); }) : Object.getOwnPropertyDescriptors ? Object.defineProperties(e, Object.getOwnPropertyDescriptors(t)) : ownKeys(Object(t)).forEach(function (r) { Object.defineProperty(e, r, Object.getOwnPropertyDescriptor(t, r)); }); } return e; }
function _defineProperty(e, r, t) { return (r = _toPropertyKey(r)) in e ? Object.defineProperty(e, r, { value: t, enumerable: !0, configurable: !0, writable: !0 }) : e[r] = t, e; }
function _toPropertyKey(t) { var i = _toPrimitive(t, "string"); return "symbol" == _typeof(i) ? i : i + ""; }
function _toPrimitive(t, r) { if ("object" != _typeof(t) || !t) return t; var e = t[Symbol.toPrimitive]; if (void 0 !== e) { var i = e.call(t, r || "default"); if ("object" != _typeof(i)) return i; throw new TypeError("@@toPrimitive must return a primitive value."); } return ("string" === r ? String : Number)(t); }
function _regenerator() { var e, t, r = "function" == typeof Symbol ? Symbol : {}, n = r.iterator || "@@iterator", o = r.toStringTag || "@@toStringTag"; function i(r, n, o, i) { var c = n && n.prototype instanceof Generator ? n : Generator, u = Object.create(c.prototype); return _regeneratorDefine2(u, "_invoke", function (r, n, o) { var i, c, u, f = 0, p = o || [], y = !1, G = { p: 0, n: 0, v: e, a: d, f: d.bind(e, 4), d: function d(t, r) { return i = t, c = 0, u = e, G.n = r, a; } }; function d(r, n) { for (c = r, u = n, t = 0; !y && f && !o && t < p.length; t++) { var o, i = p[t], d = G.p, l = i[2]; r > 3 ? (o = l === n) && (u = i[(c = i[4]) ? 5 : (c = 3, 3)], i[4] = i[5] = e) : i[0] <= d && ((o = r < 2 && d < i[1]) ? (c = 0, G.v = n, G.n = i[1]) : d < l && (o = r < 3 || i[0] > n || n > l) && (i[4] = r, i[5] = n, G.n = l, c = 0)); } if (o || r > 1) return a; throw y = !0, n; } return function (o, p, l) { if (f > 1) throw TypeError("Generator is already running"); for (y && 1 === p && d(p, l), c = p, u = l; (t = c < 2 ? e : u) || !y;) { i || (c ? c < 3 ? (c > 1 && (G.n = -1), d(c, u)) : G.n = u : G.v = u); try { if (f = 2, i) { if (c || (o = "next"), t = i[o]) { if (!(t = t.call(i, u))) throw TypeError("iterator result is not an object"); if (!t.done) return t; u = t.value, c < 2 && (c = 0); } else 1 === c && (t = i["return"]) && t.call(i), c < 2 && (u = TypeError("The iterator does not provide a '" + o + "' method"), c = 1); i = e; } else if ((t = (y = G.n < 0) ? u : r.call(n, G)) !== a) break; } catch (t) { i = e, c = 1, u = t; } finally { f = 1; } } return { value: t, done: y }; }; }(r, o, i), !0), u; } var a = {}; function Generator() {} function GeneratorFunction() {} function GeneratorFunctionPrototype() {} t = Object.getPrototypeOf; var c = [][n] ? t(t([][n]())) : (_regeneratorDefine2(t = {}, n, function () { return this; }), t), u = GeneratorFunctionPrototype.prototype = Generator.prototype = Object.create(c); function f(e) { return Object.setPrototypeOf ? Object.setPrototypeOf(e, GeneratorFunctionPrototype) : (e.__proto__ = GeneratorFunctionPrototype, _regeneratorDefine2(e, o, "GeneratorFunction")), e.prototype = Object.create(u), e; } return GeneratorFunction.prototype = GeneratorFunctionPrototype, _regeneratorDefine2(u, "constructor", GeneratorFunctionPrototype), _regeneratorDefine2(GeneratorFunctionPrototype, "constructor", GeneratorFunction), GeneratorFunction.displayName = "GeneratorFunction", _regeneratorDefine2(GeneratorFunctionPrototype, o, "GeneratorFunction"), _regeneratorDefine2(u), _regeneratorDefine2(u, o, "Generator"), _regeneratorDefine2(u, n, function () { return this; }), _regeneratorDefine2(u, "toString", function () { return "[object Generator]"; }), (_regenerator = function _regenerator() { return { w: i, m: f }; })(); }
function _regeneratorDefine2(e, r, n, t) { var i = Object.defineProperty; try { i({}, "", {}); } catch (e) { i = 0; } _regeneratorDefine2 = function _regeneratorDefine(e, r, n, t) { function o(r, n) { _regeneratorDefine2(e, r, function (e) { return this._invoke(r, n, e); }); } r ? i ? i(e, r, { value: n, enumerable: !t, configurable: !t, writable: !t }) : e[r] = n : (o("next", 0), o("throw", 1), o("return", 2)); }, _regeneratorDefine2(e, r, n, t); }
function asyncGeneratorStep(n, t, e, r, o, a, c) { try { var i = n[a](c), u = i.value; } catch (n) { return void e(n); } i.done ? t(u) : Promise.resolve(u).then(r, o); }
function _asyncToGenerator(n) { return function () { var t = this, e = arguments; return new Promise(function (r, o) { var a = n.apply(t, e); function _next(n) { asyncGeneratorStep(a, r, o, _next, _throw, "next", n); } function _throw(n) { asyncGeneratorStep(a, r, o, _next, _throw, "throw", n); } _next(void 0); }); }; }
function _typeof(o) { "@babel/helpers - typeof"; return _typeof = "function" == typeof Symbol && "symbol" == typeof Symbol.iterator ? function (o) { return typeof o; } : function (o) { return o && "function" == typeof Symbol && o.constructor === Symbol && o !== Symbol.prototype ? "symbol" : typeof o; }, _typeof(o); }
var STORAGE_KEY = "nojiri-kaseki-mobile-localonly-v1";
var CLOUD_ENDPOINT_KEY = "nojiri-kaseki-cloud-endpoint-localonly-v1";
var CLOUD_CLIENT_ID_KEY = "nojiri-kaseki-cloud-client-id-localonly-v1";
var DEFAULT_CLOUD_ENDPOINT = "";
var CLOUD_PULL_INTERVAL_MS = 20000;
var CLOUD_SAVE_DEBOUNCE_MS = 900;
var CLOUD_AUTO_PULL_ENABLED = false;
var TEAM_ROSTER_FILE_NAME = "第25次班長記載名簿.xlsx";
var TEAM_ROSTER_DAY_COL_START = 10;
var TEAM_ROSTER_DAY_COL_END = 41;
var TEAM_ROSTER_TEAM_COL = 45;
var TEAM_ROSTER_ROLE_COL = 49;
var TEAM_ROSTER_NAME_COL = 3;
var DEFAULT_SPECIMEN_PREFIX = "m";
var SPECIMEN_CATEGORY_MAP = {
  m: "哺乳類",
  b: "植物",
  l: "生痕",
  s: "貝類",
  i: "昆虫",
  g: "人類考古",
  h: "その他",
  a: "分析用試料"
};
var VALID_SPECIMEN_PREFIXES = new Set(Object.keys(SPECIMEN_CATEGORY_MAP));
var ANALYSIS_TYPE_MAP = {
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
  MG: "はぎとり資料"
};
var OUTPUT_CELL_EDIT_LABELS = {
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
  notes: "備考"
};
var OUTPUT_LIST_COLUMN_DEFS = [{
  key: "kuwaku",
  label: "区画"
}, {
  key: "team",
  label: "発掘班"
}, {
  key: "date",
  label: "日付"
}, {
  key: "dataStatus",
  label: "データ"
}, {
  key: "specimenNo",
  label: "標本番号"
}, {
  key: "category",
  label: "分類"
}, {
  key: "nameMemo",
  label: "名称"
}, {
  key: "importantFlag",
  label: "重要品指定"
}, {
  key: "unit",
  label: "ユニット"
}, {
  key: "detail",
  label: "サブユニット"
}, {
  key: "discoverer",
  label: "発見者"
}, {
  key: "identifier",
  label: "判定者"
}, {
  key: "levelRead",
  label: "レベル読値"
}, {
  key: "altitudeM",
  label: "標高(m)"
}, {
  key: "occurrenceSection",
  label: "産出状況断面"
}, {
  key: "occurrenceSketch",
  label: "産状スケッチ"
}, {
  key: "position",
  label: "平面位置"
}, {
  key: "notes",
  label: "備考"
}, {
  key: "actions",
  label: "操作"
}];
var CUSTOM_LARGE_SHAPE_TYPE = "カスタム画像";
var REQUIRED_FIELD_LABELS = {
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
  multiPoints: "平面位置（複数点）",
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
  layerRef: "地層中の位置（層理面や鍵層名）",
  layerRelative: "地層中の位置（上/下）",
  layerFromCm: "地層中の位置（cm）"
};
var HISTORY_SNAPSHOT_FIELDS = [{
  key: "specimenNo",
  label: "標本番号"
}, {
  key: "nameMemo",
  label: "名称"
}, {
  key: "category",
  label: "分類"
}, {
  key: "layerName",
  label: "地層名"
}, {
  key: "unit",
  label: "ユニット"
}, {
  key: "detail",
  label: "サブユニット"
}, {
  key: "layerFacies",
  label: "層相"
}, {
  key: "layerPosition",
  label: "地層中の位置"
}];
var HISTORY_SNAPSHOT_FIELD_KEYS = new Set(HISTORY_SNAPSHOT_FIELDS.map(function (field) {
  return field.key;
}));
var PRESET_LAYER_NAMES = ["1.芙蓉湖砂シルト部層", "2.立が鼻砂部層", "3.海端砂シルト部層", "4.その他"];
var OTHER_LAYER_NAME = "4.その他";
var LEGACY_LAYER_NAME_ALIASES = {
  "2.立が花砂部層": "2.立が鼻砂部層"
};
var PRESET_TEAMS = ["1", "2", "3", "4", "その他"];
var OTHER_TEAM_NAME = "その他";
var DEFAULT_KUWAKU_HEAD_A = "25";
var DEFAULT_KUWAKU_HEAD_B = "Ⅰ";
var DEFAULT_KUWAKU = "".concat(DEFAULT_KUWAKU_HEAD_A, "-").concat(DEFAULT_KUWAKU_HEAD_B, "--");
var DEFAULT_LAYER_NAME = "2.立が鼻砂部層";
var ALL_GRIDS_VALUE = "__KUWAKU_ALL__";
var EMPTY_KUWAKU_VALUE = "__KUWAKU_EMPTY__";
var PLAN_SIZE_CM = 400;
var ALL_UNITS_VALUE = "__UNIT_ALL__";
var EMPTY_UNIT_VALUE = "__UNIT_EMPTY__";
var ALL_DETAILS_VALUE = "__DETAIL_ALL__";
var EMPTY_DETAIL_VALUE = "__DETAIL_EMPTY__";
var ALL_DETAIL_SUBS_VALUE = "__DETAIL_SUB_ALL__";
var EMPTY_DETAIL_SUB_VALUE = "__DETAIL_SUB_EMPTY__";
var EXPORT_PLAN_ALL_UNITS_BUTTON_VALUE = "__EXPORT_PLAN_ALL_UNITS__";
var EXPORT_CATEGORY_ALL_VALUE = "__EXPORT_CATEGORY_ALL__";
var SPECIMEN_POINT_COLORS = {
  m: "#d62828",
  a: "#5b21b6",
  b: "#2a9d8f",
  l: "#f4a261",
  s: "#457b9d",
  i: "#6d597a",
  g: "#8f5a2b",
  h: "#6b7280"
};
var IMAGE_SHAPE_CANVAS_DILATE_ITERATIONS = 5;
var CUSTOM_IMAGE_SHAPE_CANVAS_DILATE_ITERATIONS = 1;
var IMAGE_QUAD_TILT_Z_SCALE = 0.38;
var IMAGE_QUAD_TILT_Z_LIMIT_M = 1.2;
var PLAN_IMAGE_TINTED_DATA_URL_MAX_LENGTH = 180000;
var LARGE_SHAPE_DIR_PATH = "./shapes";
var LARGE_SHAPE_FILE_LABEL_MAP = {
  palmate_antler: "掌状角",
  incisor: "切歯",
  constricted_shape: "くびれた形",
  rib_curved: "肋骨（湾曲形）",
  triangle: "三角",
  c_shape: "C形",
  diamond_hira: "ひし形"
};
var DEFAULT_LARGE_SHAPE_IMAGE_PATHS = {
  掌状角: "".concat(LARGE_SHAPE_DIR_PATH, "/palmate_antler.png"),
  切歯: "".concat(LARGE_SHAPE_DIR_PATH, "/incisor.png"),
  くびれた形: "".concat(LARGE_SHAPE_DIR_PATH, "/constricted_shape.png"),
  "肋骨（湾曲形）": "".concat(LARGE_SHAPE_DIR_PATH, "/rib_curved.png"),
  三角: "".concat(LARGE_SHAPE_DIR_PATH, "/triangle.png"),
  C形: "".concat(LARGE_SHAPE_DIR_PATH, "/c_shape.png"),
  ひし形: "".concat(LARGE_SHAPE_DIR_PATH, "/diamond_hira.png")
};
var LARGE_SHAPE_IMAGE_FALLBACK_PATHS = {
  掌状角: ["./assets/large-shapes/palmate_antler.png"],
  切歯: ["./assets/large-shapes/incisor.png"],
  三角: ["./assets/large-shapes/triangle.png"],
  C形: ["./assets/large-shapes/c_shape.png"],
  くびれた形: ["./assets/large-shapes/constricted_shape.png"],
  ひし形: ["./assets/large-shapes/diamond_hira.png"],
  "肋骨（湾曲形）": ["./assets/large-shapes/rib_curved.png"]
};
var EXCLUDED_LARGE_SHAPE_LABELS = new Set(["菱形", "肋骨", "肋骨（湾曲型）"]);
var LARGE_SHAPE_MANIFEST_PATH = "".concat(LARGE_SHAPE_DIR_PATH, "/manifest.json");
var largeShapeImagePathMap = new Map(Object.entries(DEFAULT_LARGE_SHAPE_IMAGE_PATHS));
var INLINE_LARGE_SHAPE_DATA_MAP = typeof window !== "undefined" && window.__INLINE_LARGE_SHAPE_DATA_MAP__ && _typeof(window.__INLINE_LARGE_SHAPE_DATA_MAP__) === "object" ? window.__INLINE_LARGE_SHAPE_DATA_MAP__ : {};
var VIEWER_HEAD_SEQUENCE = ["Ⅲ", "Ⅰ", "Ⅱ"];
var VIEWER_HEAD_INDEX_MAP = new Map(VIEWER_HEAD_SEQUENCE.map(function (head, index) {
  return [head, index];
}));
var VIEWER_ALTITUDE_BASE_M = 655;
var UNIT_CELL_COLOR_MAP = {
  U1: {
    background: "hsl(272, 64%, 93%)",
    border: "hsl(272, 38%, 80%)",
    color: "#111827"
  },
  U2: {
    background: "hsl(286, 62%, 93%)",
    border: "hsl(286, 36%, 80%)",
    color: "#111827"
  },
  U3: {
    background: "hsl(258, 60%, 93%)",
    border: "hsl(258, 36%, 80%)",
    color: "#111827"
  },
  T1: {
    background: "hsl(28, 58%, 92%)",
    border: "hsl(28, 34%, 78%)",
    color: "#111827"
  },
  T2: {
    background: "hsl(34, 56%, 92%)",
    border: "hsl(34, 34%, 78%)",
    color: "#111827"
  },
  T3: {
    background: "hsl(20, 58%, 92%)",
    border: "hsl(20, 34%, 78%)",
    color: "#111827"
  },
  T4: {
    background: "hsl(196, 74%, 92%)",
    border: "hsl(196, 42%, 79%)",
    color: "#111827"
  },
  T5: {
    background: "hsl(52, 84%, 92%)",
    border: "hsl(52, 46%, 79%)",
    color: "#111827"
  },
  T6: {
    background: "hsl(0, 82%, 93%)",
    border: "hsl(0, 44%, 80%)",
    color: "#111827"
  },
  T7: {
    background: "hsl(88, 62%, 91%)",
    border: "hsl(88, 34%, 77%)",
    color: "#111827"
  }
};
var PHOTO_COMPRESSION_STEPS = [{
  maxLength: 1280,
  quality: 0.72
}, {
  maxLength: 960,
  quality: 0.62
}, {
  maxLength: 720,
  quality: 0.54
}, {
  maxLength: 560,
  quality: 0.46
}];
var createInitialState = function createInitialState() {
  return {
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
      updatedAt: ""
    },
    records: []
  };
};
function createInitialOutputColumnVisibilityMap() {
  var visibilityMap = {};
  OUTPUT_LIST_COLUMN_DEFS.forEach(function (column) {
    visibilityMap[column.key] = true;
  });
  return visibilityMap;
}
var stateNeedsRewriteAfterLoad = false;
var state = loadState();
var editingRecordId = null;
var activeEditRecordContext = null;
var currentSectionDiagrams = [];
var currentPhotos = [];
var selectedCardRecordId = "";
var selectedOutputKuwaku = ALL_GRIDS_VALUE;
var selectedOutputCategory = EXPORT_CATEGORY_ALL_VALUE;
var selectedOutputStatus = "all";
var selectedOutputDate = "";
var outputSearchText = "";
var outputFilterMemory = {
  kuwaku: ALL_GRIDS_VALUE,
  category: EXPORT_CATEGORY_ALL_VALUE,
  status: "all",
  date: "",
  searchText: ""
};
var outputListColumnVisibility = createInitialOutputColumnVisibilityMap();
var activeOutputCellEdit = null;
var selectedPlanKuwaku = "";
var selectedPlanCategory = EXPORT_CATEGORY_ALL_VALUE;
var selectedPlanUnit = "";
var selectedPlanDetail = ALL_DETAILS_VALUE;
var selectedPlanDetailSub = ALL_DETAIL_SUBS_VALUE;
var selectedViewerKuwaku = ALL_GRIDS_VALUE;
var selectedViewerCategory = EXPORT_CATEGORY_ALL_VALUE;
var selectedViewerUnit = ALL_UNITS_VALUE;
var selectedViewerDetail = ALL_DETAILS_VALUE;
var selectedViewerDetailSub = ALL_DETAIL_SUBS_VALUE;
var selectedViewerPerspective = "top";
var viewerVerticalScale = 1;
var exportListRangeKuwaku = ALL_GRIDS_VALUE;
var exportListRangeCategory = EXPORT_CATEGORY_ALL_VALUE;
var exportListRangeStatus = "all";
var exportListRangeSpecimenFrom = "";
var exportListRangeSpecimenTo = "";
var exportListRangeDateFrom = "";
var exportListRangeDateTo = "";
var exportCardRangeKuwaku = ALL_GRIDS_VALUE;
var exportCardRangeCategory = EXPORT_CATEGORY_ALL_VALUE;
var exportCardRangeStatus = "all";
var exportCardRangeDateFrom = "";
var exportCardRangeDateTo = "";
var exportPlanKuwaku = "";
var exportPlanCategory = EXPORT_CATEGORY_ALL_VALUE;
var exportPlanDateFrom = "";
var exportPlanDateTo = "";
var exportPlanModeUnitEnabled = true;
var exportPlanModeDetailEnabled = false;
var exportPlanModeDetailSubEnabled = false;
var exportPlanModeUnitValues = new Set();
var exportPlanModeDetailUnitValue = "";
var exportPlanModeDetailValues = new Set();
var exportPlanModeDetailSubUnitValue = "";
var exportPlanModeDetailSubDetailValue = "";
var exportPlanModeDetailSubValues = new Set();
var exportPlanModeUnitTouched = false;
var exportPlanModeDetailTouched = false;
var exportPlanModeDetailSubTouched = false;
var outputListSortKey = "kuwaku";
var outputListSortDirection = "asc";
var isOverwriteMode = false;
var overwriteOriginalRecord = null;
var toastTimer = null;
var quotaRecoveryInProgress = false;
var cloudEndpoint = loadCloudEndpoint();
var cloudClientId = loadOrCreateCloudClientId();
var cloudSaveTimer = null;
var cloudPullTimer = null;
var cloudPushInProgress = false;
var cloudPullInProgress = false;
var cloudLastSyncedAt = "";
var cloudLastPulledAt = "";
var cloudLastErrorAt = 0;
var teamRosterAssignmentMap = new Map();
var teamRosterLoaded = false;
var tabButtons = document.querySelectorAll(".tab-button");
var tabPanels = document.querySelectorAll(".tab-panel");
var siteForm = document.getElementById("site-form");
var recordForm = document.getElementById("record-form");
var recordFormHost = document.getElementById("record-form-host");
var editRecordFormHost = document.getElementById("edit-record-form-host");
var editTabPanel = document.getElementById("edit-tab");
var editHistoryPanel = document.getElementById("edit-history-panel");
var editHistoryList = document.getElementById("edit-history-list");
var editKuwakuHeadAInput = document.getElementById("edit-kuwaku-head-a");
var editKuwakuHeadBInput = document.getElementById("edit-kuwaku-head-b");
var editKuwakuBlockInput = document.getElementById("edit-kuwaku-block");
var editKuwakuNoInput = document.getElementById("edit-kuwaku-no");
var editLevelHeightInput = document.getElementById("edit-level-height");
var editDateInput = document.getElementById("edit-date");
var editTeamInput = document.getElementById("edit-team");
var editTeamOtherInput = document.getElementById("edit-team-other");
var editTeamLeadInput = document.getElementById("edit-team-lead");
var editRecorderInput = document.getElementById("edit-recorder");
var rosterStatusEl = document.getElementById("roster-status");
var recordIdInput = document.getElementById("record-id-input");
var recordSubmitBtn = document.getElementById("record-submit-btn");
var recordCopyToInputBtn = document.getElementById("record-copy-to-input-btn");
var recordCopyToInputTopBtn = document.getElementById("record-copy-to-input-top-btn");
var recordPrevBtn = document.getElementById("record-prev-btn");
var recordNextBtn = document.getElementById("record-next-btn");
var recordResetBtn = document.getElementById("record-reset-btn");
var recordNewBtn = document.getElementById("record-new-btn");
var recordTableBody = document.getElementById("record-table-body");
var editRecordTableBody = document.getElementById("edit-record-table-body");
var outputListBody = document.getElementById("output-list-body");
var outputListTable = document.getElementById("output-list-table");
var cardOutputList = document.getElementById("card-output-list");
var outputKuwakuSelect = document.getElementById("output-kuwaku-select");
var outputCategorySelect = document.getElementById("output-category-select");
var outputStatusSelect = document.getElementById("output-status-select");
var outputDateSelect = document.getElementById("output-date-select");
var outputSearchInput = document.getElementById("output-search-input");
var outputFilterSummary = document.getElementById("output-filter-summary");
var outputColumnToggleRow = document.getElementById("output-column-toggle-row");
var planKuwakuSelect = document.getElementById("plan-kuwaku-select");
var planCategorySelect = document.getElementById("plan-category-select");
var planUnitSelect = document.getElementById("plan-unit-select");
var planDetailSelect = document.getElementById("plan-detail-select");
var planDetailSubSelect = document.getElementById("plan-detail-sub-select");
var planMapLegend = document.getElementById("plan-map-legend");
var planMapWrap = document.getElementById("plan-map-wrap");
var planKuwakuInfo = document.getElementById("plan-kuwaku-info");
var viewerKuwakuSelect = document.getElementById("viewer-kuwaku-select");
var viewerCategorySelect = document.getElementById("viewer-category-select");
var viewerUnitSelect = document.getElementById("viewer-unit-select");
var viewerDetailSelect = document.getElementById("viewer-detail-select");
var viewerDetailSubSelect = document.getElementById("viewer-detail-sub-select");
var viewerKuwakuInfo = document.getElementById("viewer-kuwaku-info");
var viewerMapLegend = document.getElementById("viewer-map-legend");
var viewerCanvasWrap = document.getElementById("viewer-canvas-wrap");
var viewerTooltip = document.getElementById("viewer-tooltip");
var viewerStatus = document.getElementById("viewer-status");
var viewerViewTopBtn = document.getElementById("viewer-view-top-btn");
var viewerViewSeBtn = document.getElementById("viewer-view-se-btn");
var viewerViewEastBtn = document.getElementById("viewer-view-east-btn");
var viewerViewWestBtn = document.getElementById("viewer-view-west-btn");
var viewerViewSouthBtn = document.getElementById("viewer-view-south-btn");
var viewerViewNorthBtn = document.getElementById("viewer-view-north-btn");
var viewerZScaleInput = document.getElementById("viewer-z-scale-input");
var viewerZScaleValue = document.getElementById("viewer-z-scale-value");
var largeAxisDirectionRow = document.getElementById("large-axis-direction-row");
var largeAxisPlungeRow = document.getElementById("large-axis-plunge-row");
var largeAxisPlungeDirRow = document.getElementById("large-axis-plunge-dir-row");
var planeAttitudeRow = document.getElementById("plane-attitude-row");
var exportListRangeKuwakuSelect = document.getElementById("export-range-kuwaku-select");
var exportListRangeCategorySelect = document.getElementById("export-range-category-select");
var exportListRangeStatusSelect = document.getElementById("export-range-status-select");
var exportListRangeSpecimenFromInput = document.getElementById("export-range-specimen-from");
var exportListRangeSpecimenToInput = document.getElementById("export-range-specimen-to");
var exportListRangeDateFromInput = document.getElementById("export-range-date-from");
var exportListRangeDateToInput = document.getElementById("export-range-date-to");
var exportListRangeSummaryEl = document.getElementById("export-range-summary");
var exportCardRangeKuwakuSelect = document.getElementById("export-card-kuwaku-select");
var exportCardRangeCategorySelect = document.getElementById("export-card-category-select");
var exportCardRangeStatusSelect = document.getElementById("export-card-status-select");
var exportCardRangeDateFromInput = document.getElementById("export-card-date-from");
var exportCardRangeDateToInput = document.getElementById("export-card-date-to");
var exportCardRangeSummaryEl = document.getElementById("export-card-summary");
var exportPlanKuwakuSelect = document.getElementById("export-plan-kuwaku-select");
var exportPlanCategorySelect = document.getElementById("export-plan-category-select");
var exportPlanDateFromInput = document.getElementById("export-plan-date-from");
var exportPlanDateToInput = document.getElementById("export-plan-date-to");
var exportPlanModeUnitCheck = document.getElementById("export-plan-mode-unit-check");
var exportPlanModeUnitButtons = document.getElementById("export-plan-mode-unit-buttons");
var exportPlanModeUnitStats = document.getElementById("export-plan-mode-unit-stats");
var exportPlanModeDetailCheck = document.getElementById("export-plan-mode-detail-check");
var exportPlanModeDetailUnitSelect = document.getElementById("export-plan-mode-detail-unit-select");
var exportPlanModeDetailButtons = document.getElementById("export-plan-mode-detail-buttons");
var exportPlanModeDetailStats = document.getElementById("export-plan-mode-detail-stats");
var exportPlanModeDetailSubCheck = document.getElementById("export-plan-mode-detail-sub-check");
var exportPlanModeDetailSubUnitSelect = document.getElementById("export-plan-mode-detail-sub-unit-select");
var exportPlanModeDetailSubDetailSelect = document.getElementById("export-plan-mode-detail-sub-detail-select");
var exportPlanModeDetailSubButtons = document.getElementById("export-plan-mode-detail-sub-buttons");
var exportPlanModeDetailSubStats = document.getElementById("export-plan-mode-detail-sub-stats");
var exportPlanSummaryEl = document.getElementById("export-plan-summary");
var specimenTabButtons = document.querySelectorAll(".specimen-tab");
var specimenPrefixInput = document.getElementById("specimen-prefix-input");
var specimenSerialInput = document.getElementById("specimen-serial-input");
var specimenNoInput = document.getElementById("specimen-no-input");
var specimenPrefixLabel = document.getElementById("specimen-prefix-label");
var specimenDuplicateWarning = document.getElementById("specimen-duplicate-warning");
var analysisTypeRow = document.getElementById("analysis-type-row");
var analysisTypeSelect = document.getElementById("analysis-type-select");
var nsDirInput = document.getElementById("ns-dir-input");
var ewDirInput = document.getElementById("ew-dir-input");
var importantFlagInput = document.getElementById("important-flag-input");
var simpleRecordFlagInput = document.getElementById("simple-record-flag-input");
var occurrenceSectionInput = document.getElementById("occurrence-section-input");
var occurrenceSketchInput = document.getElementById("occurrence-sketch-input");
var layerRelativeInput = document.getElementById("layer-relative-input");
var planSizeModeInput = document.getElementById("plan-size-mode-input");
var largeShapeTypeInput = document.getElementById("large-shape-type-input");
var largeAxisDirectionInput = document.getElementById("large-axis-direction-input");
var largeAxisPlungeInput = document.getElementById("large-axis-plunge-input");
var largeAxisPlungeDirInput = document.getElementById("large-axis-plunge-dir-input");
var planeStrikeInput = document.getElementById("plane-strike-input");
var planeDipInput = document.getElementById("plane-dip-input");
var planeDipDirInput = document.getElementById("plane-dip-dir-input");
var largeShapeImageButtons = document.getElementById("large-shape-image-buttons");
var largeShapeImagePreview = document.getElementById("large-shape-image-preview");
var largeShapeImagePreviewTitle = document.getElementById("large-shape-image-preview-title");
var largeShapeImagePreviewImg = document.getElementById("large-shape-image-preview-img");
var customLargeImageControls = document.getElementById("custom-large-image-controls");
var customLargeImageNameInput = document.getElementById("custom-large-image-name-input");
var customLargeImageFileInput = document.getElementById("custom-large-image-file-input");
var customLargeImageDataUrlInput = document.getElementById("custom-large-image-data-url-input");
var customLargeImageAspectInput = document.getElementById("custom-large-image-aspect-input");
var customLargeImageClearBtn = document.getElementById("custom-large-image-clear-btn");
var customLargeImageStatus = document.getElementById("custom-large-image-status");
var line1NsDirInput = document.getElementById("line1-ns-dir-input");
var line1EwDirInput = document.getElementById("line1-ew-dir-input");
var line2NsDirInput = document.getElementById("line2-ns-dir-input");
var line2EwDirInput = document.getElementById("line2-ew-dir-input");
var multiPointSection = document.getElementById("multi-point-section");
var multiPointRows = document.getElementById("multi-point-rows");
var multiPointAddBtn = document.getElementById("multi-point-add-btn");
var largeShapeSection = document.getElementById("large-shape-section");
var largeShapePanels = document.querySelectorAll(".large-shape-panel[data-large-shape-panel]");
var layerTabButtons = document.querySelectorAll(".layer-tab");
var layerNameInput = document.getElementById("layer-name-input");
var layerOtherInput = document.getElementById("layer-other-input");
var teamOtherInput = document.getElementById("team-other-input");
var sectionDiagramCameraInput = document.getElementById("section-diagram-camera-input");
var sectionDiagramInput = document.getElementById("section-diagram-input");
var sectionDiagramList = document.getElementById("section-diagram-list");
var photoCameraInput = document.getElementById("photo-camera-input");
var photoInput = document.getElementById("photo-input");
var photoList = document.getElementById("photo-list");
var exportListCsvBtn = document.getElementById("export-list-csv-btn");
var exportCardCsvBtn = document.getElementById("export-card-csv-btn");
var exportListPdfBtn = document.getElementById("export-list-pdf-btn");
var exportCardPdfBtn = document.getElementById("export-card-pdf-btn");
var exportPlanPdfBtn = document.getElementById("export-plan-pdf-btn");
var exportJsonBtn = document.getElementById("export-json-btn");
var importJsonInput = document.getElementById("import-json-input");
var cloudEndpointInput = document.getElementById("cloud-endpoint-input");
var cloudConnectBtn = document.getElementById("cloud-connect-btn");
var cloudSyncBtn = document.getElementById("cloud-sync-btn");
var cloudDisableBtn = document.getElementById("cloud-disable-btn");
var cloudStatusEl = document.getElementById("cloud-status");
var toastEl = document.getElementById("toast");
var positionPreviewBtn = document.getElementById("position-preview-btn");
var positionPreviewModal = document.getElementById("position-preview-modal");
var positionPreviewCloseBtn = document.getElementById("position-preview-close-btn");
var positionPreviewMeta = document.getElementById("position-preview-meta");
var positionPreviewMap = document.getElementById("position-preview-map");
var cellEditModal = document.getElementById("cell-edit-modal");
var cellEditForm = document.getElementById("cell-edit-form");
var cellEditTitle = document.getElementById("cell-edit-title");
var cellEditMeta = document.getElementById("cell-edit-meta");
var cellEditFields = document.getElementById("cell-edit-fields");
var cellEditCloseBtn = document.getElementById("cell-edit-close-btn");
var cellEditSaveBtn = document.getElementById("cell-edit-save-btn");
var cellEditCancelBtn = document.getElementById("cell-edit-cancel-btn");
var hoverEditMenuEl = null;
var hoverEditMenuRecordId = "";
var hoverEditMenuKuwaku = "";
var TOUCH_LONG_PRESS_MS = 520;
var TOUCH_LONG_PRESS_MOVE_THRESHOLD_PX = 16;
var viewerTouchLongPressState = {
  pointerId: null,
  startX: 0,
  startY: 0,
  timer: 0,
  triggered: false
};
var viewer3d = {
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
  defaultLeftMouseAction: null
};
var planLargeShapeImageCache = new Map();
var planLargeShapeTintedCanvasCache = new Map();
var planLargeShapeTintedDataUrlCache = new Map();
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
  resetRecordForm({
    showMessage: false
  });
  renderRecordTable();
  renderOutputs();
  void loadLargeShapeImageManifest();
  void loadTeamRosterFromDefaultFile();
  void bootstrapCloudSync();
}
function bindEvents() {
  tabButtons.forEach(function (button) {
    button.addEventListener("click", function () {
      return setActiveTab(button.dataset.tab);
    });
  });
  if (cloudConnectBtn) {
    cloudConnectBtn.addEventListener("click", function () {
      void handleCloudConnect();
    });
  }
  if (cloudSyncBtn) {
    cloudSyncBtn.addEventListener("click", function () {
      void handleCloudManualReload();
    });
  }
  if (cloudDisableBtn) {
    cloudDisableBtn.addEventListener("click", function () {
      disableCloudSync({
        showToastMessage: true
      });
    });
  }
  if (cloudEndpointInput) {
    cloudEndpointInput.addEventListener("keydown", function (event) {
      if (event.key === "Enter") {
        event.preventDefault();
        void handleCloudConnect();
      }
    });
  }
  if (positionPreviewBtn) {
    positionPreviewBtn.addEventListener("click", function () {
      openPositionPreviewModal();
    });
  }
  if (positionPreviewCloseBtn) {
    positionPreviewCloseBtn.addEventListener("click", function () {
      closePositionPreviewModal();
    });
  }
  if (positionPreviewModal) {
    positionPreviewModal.addEventListener("click", function (event) {
      if (event.target === positionPreviewModal) {
        closePositionPreviewModal();
      }
    });
  }
  if (cellEditCloseBtn) {
    cellEditCloseBtn.addEventListener("click", function () {
      closeOutputCellEditModal();
    });
  }
  if (cellEditCancelBtn) {
    cellEditCancelBtn.addEventListener("click", function () {
      closeOutputCellEditModal();
    });
  }
  if (cellEditModal) {
    cellEditModal.addEventListener("click", function (event) {
      if (event.target === cellEditModal) {
        closeOutputCellEditModal();
      }
    });
  }
  if (cellEditForm) {
    var handleCellEditSave = function handleCellEditSave(event) {
      event.preventDefault();
      saveOutputCellEditFromModal();
    };
    cellEditForm.addEventListener("submit", handleCellEditSave);
    if (cellEditSaveBtn) {
      cellEditSaveBtn.addEventListener("click", handleCellEditSave);
    }
  }
  document.addEventListener("keydown", function (event) {
    if (event.key !== "Escape") {
      return;
    }
    hideHoverEditMenu();
    closeOutputCellEditModal();
    if (positionPreviewModal && !positionPreviewModal.classList.contains("hidden")) {
      closePositionPreviewModal();
    }
  });
  document.addEventListener("pointerdown", function (event) {
    if (!hoverEditMenuEl || hoverEditMenuEl.hidden) {
      return;
    }
    var target = event.target;
    if (target instanceof Node && hoverEditMenuEl.contains(target)) {
      return;
    }
    hideHoverEditMenu();
  });
  document.addEventListener("scroll", function () {
    hideHoverEditMenu();
  }, true);
  if (viewerCanvasWrap) {
    viewerCanvasWrap.addEventListener("click", function (event) {
      var target = event.target;
      if (!(target instanceof Element)) {
        return;
      }
      var fallbackButton = target.closest("button[data-action='viewer-open-plan']");
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
    customLargeImageFileInput.addEventListener("change", function () {
      var _ref = _asyncToGenerator(_regenerator().m(function _callee(event) {
        var input, file;
        return _regenerator().w(function (_context) {
          while (1) switch (_context.n) {
            case 0:
              input = event.target;
              if (input instanceof HTMLInputElement) {
                _context.n = 1;
                break;
              }
              return _context.a(2);
            case 1:
              file = Array.from(input.files || [])[0];
              if (file) {
                _context.n = 2;
                break;
              }
              return _context.a(2);
            case 2:
              _context.n = 3;
              return setCustomLargeImageFromFile(file);
            case 3:
              input.value = "";
            case 4:
              return _context.a(2);
          }
        }, _callee);
      }));
      return function (_x) {
        return _ref.apply(this, arguments);
      };
    }());
  }
  if (customLargeImageClearBtn) {
    customLargeImageClearBtn.addEventListener("click", function () {
      clearCustomLargeImageFields();
      syncLargeShapeImagePreviewForCurrentForm();
      updateEditMissingRequiredHighlights();
    });
  }
  if (editTabPanel) {
    editTabPanel.addEventListener("input", function () {
      updateEditMissingRequiredHighlights();
      renderRecordTable();
    });
    editTabPanel.addEventListener("change", function () {
      updateEditMissingRequiredHighlights();
      renderRecordTable();
    });
    editTabPanel.addEventListener("click", function (event) {
      var target = event.target;
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
  siteForm.addEventListener("submit", function (event) {
    event.preventDefault();
    var formData = new FormData(siteForm);
    var kuwakuHeadA = normalizeKuwakuHeadA(formData.get("kuwakuHeadA"));
    var kuwakuHeadB = normalizeKuwakuHeadB(formData.get("kuwakuHeadB"));
    var kuwakuBlock = normalizeKuwakuBlock(formData.get("kuwakuBlock"));
    var kuwakuNo = normalizeKuwakuNo(formData.get("kuwakuNo"));
    var teamState = normalizeTeamState(value(formData.get("team")), value(formData.get("teamOther")));
    var nextSiteKuwaku = buildKuwaku(kuwakuHeadA, kuwakuHeadB, kuwakuBlock, kuwakuNo);
    var normalizedKuwaku = parseKuwaku(nextSiteKuwaku);
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
      updatedAt: nowIso()
    };
    selectedPlanKuwaku = kuwakuValueForSelect(nextSiteKuwaku);
    persist("区画（グリッド）情報を保存しました");
    renderRecordTable();
    renderOutputs();
  });
  siteForm.elements.team.addEventListener("change", function () {
    syncTeamOtherInput(siteForm.elements.team.value);
    applyTeamRosterAutofill();
  });
  siteForm.elements.date.addEventListener("change", function () {
    applyTeamRosterAutofill();
  });
  ["kuwakuHeadA", "kuwakuHeadB", "kuwakuBlock", "kuwakuNo"].forEach(function (name) {
    var input = siteForm.elements[name];
    if (!(input instanceof Element)) {
      return;
    }
    input.addEventListener("input", updateDuplicateSpecimenWarning);
    input.addEventListener("change", updateDuplicateSpecimenWarning);
  });
  if (editTeamInput) {
    editTeamInput.addEventListener("change", function () {
      syncEditTeamOtherInput(editTeamInput.value);
      editTeamInput.classList.remove("overwrite-updated");
      if (editTeamOtherInput) {
        editTeamOtherInput.classList.remove("overwrite-updated");
      }
    });
  }
  specimenTabButtons.forEach(function (button) {
    button.addEventListener("click", function () {
      activateSpecimenPrefix(button.dataset.prefix);
      updateSpecimenNoFromParts();
    });
  });
  specimenSerialInput.addEventListener("input", function () {
    updateSpecimenNoFromParts();
  });
  if (analysisTypeSelect) {
    analysisTypeSelect.addEventListener("change", function () {
      analysisTypeSelect.value = normalizeAnalysisType(analysisTypeSelect.value);
      analysisTypeSelect.classList.remove("overwrite-updated");
    });
  }
  recordForm.addEventListener("click", function (event) {
    var target = event.target;
    if (!(target instanceof Element)) {
      return;
    }
    var button = target.closest(".dir-tab");
    if (!(button instanceof HTMLButtonElement)) {
      return;
    }
    activateDirectionTab(button.dataset.group, button.dataset.value);
  });
  if (multiPointAddBtn && multiPointRows) {
    multiPointAddBtn.addEventListener("click", function () {
      multiPointRows.append(createMultiPointRowElement(createDefaultPlanMultiPoint()));
      syncMultiPointRemoveButtonState();
    });
  }
  if (multiPointRows) {
    multiPointRows.addEventListener("click", function (event) {
      var target = event.target;
      if (!(target instanceof Element)) {
        return;
      }
      var removeButton = target.closest("[data-multi-point-remove]");
      if (!(removeButton instanceof HTMLButtonElement)) {
        return;
      }
      var row = removeButton.closest("[data-multi-point-row]");
      if (!(row instanceof HTMLElement)) {
        return;
      }
      var rows = multiPointRows.querySelectorAll("[data-multi-point-row]");
      if (rows.length <= 1) {
        return;
      }
      row.remove();
      syncMultiPointRemoveButtonState();
    });
  }
  layerTabButtons.forEach(function (button) {
    button.addEventListener("click", function () {
      clearLayerSavedTabState();
      layerOtherInput.classList.remove("saved-carry-value");
      activateLayerTab(button.dataset.layer);
    });
  });
  recordForm.addEventListener("submit", function (event) {
    var _activeEditRecordCont, _activeEditRecordCont2, _editSiteSnapshot, _editSiteSnapshot2, _editSiteSnapshot3, _editSiteSnapshot4, _editSiteSnapshot5, _editSiteSnapshot6, _siteSnapshot, _siteSnapshot2, _siteSnapshot3, _siteSnapshot4, _siteSnapshot5, _siteSnapshot6;
    event.preventDefault();
    var isEditTab = getActiveTabId() === "edit-tab";
    var recordKuwaku = "";
    var siteSnapshot = null;
    var editSiteSnapshot = null;
    if (isEditTab) {
      var headA = normalizeKuwakuHeadA(editKuwakuHeadAInput === null || editKuwakuHeadAInput === void 0 ? void 0 : editKuwakuHeadAInput.value);
      var headB = normalizeKuwakuHeadB(editKuwakuHeadBInput === null || editKuwakuHeadBInput === void 0 ? void 0 : editKuwakuHeadBInput.value);
      var block = normalizeKuwakuBlock(editKuwakuBlockInput === null || editKuwakuBlockInput === void 0 ? void 0 : editKuwakuBlockInput.value);
      var no = normalizeKuwakuNo(editKuwakuNoInput === null || editKuwakuNoInput === void 0 ? void 0 : editKuwakuNoInput.value);
      recordKuwaku = buildKuwaku(headA, headB, block, no);
      var editTeamState = normalizeTeamState(value(editTeamInput === null || editTeamInput === void 0 ? void 0 : editTeamInput.value), value(editTeamOtherInput === null || editTeamOtherInput === void 0 ? void 0 : editTeamOtherInput.value));
      editSiteSnapshot = {
        levelHeight: value(editLevelHeightInput === null || editLevelHeightInput === void 0 ? void 0 : editLevelHeightInput.value),
        date: value(editDateInput === null || editDateInput === void 0 ? void 0 : editDateInput.value),
        team: editTeamState.team,
        teamOther: editTeamState.teamOther,
        teamLead: value(editTeamLeadInput === null || editTeamLeadInput === void 0 ? void 0 : editTeamLeadInput.value),
        recorder: value(editRecorderInput === null || editRecorderInput === void 0 ? void 0 : editRecorderInput.value)
      };
    } else {
      var siteFormData = new FormData(siteForm);
      var siteKuwakuHeadA = normalizeKuwakuHeadA(siteFormData.get("kuwakuHeadA"));
      var siteKuwakuHeadB = normalizeKuwakuHeadB(siteFormData.get("kuwakuHeadB"));
      var siteKuwakuBlock = normalizeKuwakuBlock(siteFormData.get("kuwakuBlock"));
      var siteKuwakuNo = normalizeKuwakuNo(siteFormData.get("kuwakuNo"));
      var siteTeamState = normalizeTeamState(value(siteFormData.get("team")), value(siteFormData.get("teamOther")));
      var nextSiteKuwaku = buildKuwaku(siteKuwakuHeadA, siteKuwakuHeadB, siteKuwakuBlock, siteKuwakuNo);
      var normalizedKuwaku = parseKuwaku(nextSiteKuwaku);
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
        recorder: value(siteFormData.get("recorder"))
      };
      recordKuwaku = nextSiteKuwaku;
    }
    var formData = new FormData(recordForm);
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
        updatedAt: nowIso()
      };
      selectedPlanKuwaku = kuwakuValueForSelect(siteSnapshot.kuwaku);
    }
    var specimenPrefix = normalizeSpecimenPrefix(value(formData.get("specimenPrefix")));
    var specimenSerial = compactNoSpaceValue(formData.get("specimenSerial"));
    if (!specimenSerial) {
      showToast("標本番号は必須です");
      return;
    }
    if (!/^\d+$/.test(specimenSerial)) {
      showToast("標本番号の数字部分は半角数字で入力してください");
      return;
    }
    var analysisType = specimenPrefix === "a" ? normalizeAnalysisType(value(formData.get("analysisType"))) : "";
    if (specimenPrefix === "a" && !analysisType) {
      showToast("a: 分析用試料を選んだ場合は、区分を選択してください");
      return;
    }
    var attachmentChecklistError = validateAttachmentChecklistForSave(formData);
    if (attachmentChecklistError) {
      showToast(attachmentChecklistError);
      return;
    }
    var nowIsoValue = new Date().toISOString();
    var editingId = editingRecordId || recordIdInput.value;
    var found = isEditTab ? findRecordByEditContext(editingId, (_activeEditRecordCont = activeEditRecordContext) === null || _activeEditRecordCont === void 0 ? void 0 : _activeEditRecordCont.recordIndex, (_activeEditRecordCont2 = activeEditRecordContext) === null || _activeEditRecordCont2 === void 0 ? void 0 : _activeEditRecordCont2.recordSnapshot) : findRecord(editingId);
    if (isEditTab && !found) {
      showToast("編集対象が見つかりません。リストから編集を選び直してください");
      return;
    }
    var specimenNo = buildSpecimenNo(specimenPrefix, specimenSerial);
    var parsedKuwakuForDuplicateCheck = parseKuwaku(recordKuwaku);
    var hasKuwakuForDuplicateCheck = value(parsedKuwakuForDuplicateCheck.block) && value(parsedKuwakuForDuplicateCheck.no);
    var duplicateRecord = hasKuwakuForDuplicateCheck ? findDuplicateRecordByKuwakuAndSpecimen(recordKuwaku, specimenNo, value(found === null || found === void 0 ? void 0 : found.id)) : null;
    if (duplicateRecord) {
      showToast("\u3053\u306E\u533A\u753B\u306B\u306F ".concat(specimenNo, " \u304C\u3059\u3067\u306B\u3042\u308A\u307E\u3059"));
      return;
    }
    var saveAnswer = window.confirm(isEditTab ? "".concat(specimenNo, "\u306E\u60C5\u5831\u3092\u4E0A\u66F8\u304D\u4FDD\u5B58\u3057\u307E\u3059\u304B\uFF1F") : "".concat(specimenNo, "\u306E\u60C5\u5831\u3092\u4FDD\u5B58\u3057\u307E\u3059\u304B\uFF1F"));
    if (!saveAnswer) {
      return;
    }
    var hasSpecimenChanged = Boolean(found && found.specimenNo !== specimenNo);
    var recordId = isEditTab ? (found === null || found === void 0 ? void 0 : found.id) || editingId || newId("record") : hasSpecimenChanged ? newId("record") : (found === null || found === void 0 ? void 0 : found.id) || editingId || newId("record");
    var recordSiteSnapshot = isEditTab ? {
      levelHeight: value((_editSiteSnapshot = editSiteSnapshot) === null || _editSiteSnapshot === void 0 ? void 0 : _editSiteSnapshot.levelHeight),
      date: value((_editSiteSnapshot2 = editSiteSnapshot) === null || _editSiteSnapshot2 === void 0 ? void 0 : _editSiteSnapshot2.date),
      team: value((_editSiteSnapshot3 = editSiteSnapshot) === null || _editSiteSnapshot3 === void 0 ? void 0 : _editSiteSnapshot3.team),
      teamOther: value((_editSiteSnapshot4 = editSiteSnapshot) === null || _editSiteSnapshot4 === void 0 ? void 0 : _editSiteSnapshot4.teamOther),
      teamLead: value((_editSiteSnapshot5 = editSiteSnapshot) === null || _editSiteSnapshot5 === void 0 ? void 0 : _editSiteSnapshot5.teamLead),
      recorder: value((_editSiteSnapshot6 = editSiteSnapshot) === null || _editSiteSnapshot6 === void 0 ? void 0 : _editSiteSnapshot6.recorder)
    } : {
      levelHeight: value((_siteSnapshot = siteSnapshot) === null || _siteSnapshot === void 0 ? void 0 : _siteSnapshot.levelHeight),
      date: value((_siteSnapshot2 = siteSnapshot) === null || _siteSnapshot2 === void 0 ? void 0 : _siteSnapshot2.date),
      team: value((_siteSnapshot3 = siteSnapshot) === null || _siteSnapshot3 === void 0 ? void 0 : _siteSnapshot3.team),
      teamOther: value((_siteSnapshot4 = siteSnapshot) === null || _siteSnapshot4 === void 0 ? void 0 : _siteSnapshot4.teamOther),
      teamLead: value((_siteSnapshot5 = siteSnapshot) === null || _siteSnapshot5 === void 0 ? void 0 : _siteSnapshot5.teamLead),
      recorder: value((_siteSnapshot6 = siteSnapshot) === null || _siteSnapshot6 === void 0 ? void 0 : _siteSnapshot6.recorder)
    };
    var recordTeamState = normalizeTeamState(recordSiteSnapshot.team, recordSiteSnapshot.teamOther);
    var planSizeMode = normalizePlanSizeMode(value(formData.get("planSizeMode")));
    var planMultiPoints = planSizeMode === "複数点" ? readMultiPointRowsFromForm() : [];
    var rawLargeShapeType = value(formData.get("largeShapeType"));
    var largeShapeType = planSizeMode === "大きなもの" ? normalizeLargeShapeType(rawLargeShapeType) || normalizeLargeShapeLabel(rawLargeShapeType) : "";
    var isLineShape = largeShapeType === "直線状";
    var usesAxisDirection = isLineShape || largeShapeType === "長方形" || largeShapeType === "楕円";
    var largeAxisDirection = planSizeMode === "大きなもの" ? normalizeLargeAxisDirection(value(formData.get("largeAxisDirection"))) : "";
    var largeAxisPlungeDeg = planSizeMode === "大きなもの" ? normalizeLargeAxisPlungeDeg(value(formData.get("largeAxisPlungeDeg"))) : "";
    var largeAxisPlungeDir8 = planSizeMode === "大きなもの" ? normalizeCompass8Direction(value(formData.get("largeAxisPlungeDir8"))) : "";
    var planeStrikeDirection = planSizeMode === "大きなもの" ? normalizePlaneStrikeDirection(value(formData.get("planeStrikeDirection")) || (usesAxisDirection ? largeAxisDirection : "")) : "";
    var planeDipDeg = planSizeMode === "大きなもの" ? normalizePlaneDipDeg(value(formData.get("planeDipDeg"))) : "";
    var planeDipDir8 = planSizeMode === "大きなもの" ? normalizeCompass8Direction(value(formData.get("planeDipDir8"))) : "";
    var lineLengthCm = value(formData.get("lineLengthCm"));
    var rectSide1Cm = value(formData.get("rectSide1Cm"));
    var rectSide2Cm = value(formData.get("rectSide2Cm"));
    var ellipseLongRadiusCm = value(formData.get("ellipseLongRadiusCm"));
    var ellipseShortRadiusCm = value(formData.get("ellipseShortRadiusCm"));
    var altitudeInputEnabled = normalizeToggleFlag(formData.get("altitudeInputEnabled"));
    var altitudeDirectM = altitudeInputEnabled === "1" ? value(formData.get("altitudeDirectM")) : "";
    var imageCornerFields = extractImageCornerFieldsFromFormData(formData);
    var isLargeImageShape = planSizeMode === "大きなもの" && isLargeShapeImageType(largeShapeType);
    var isCustomLargeImageShape = isLargeImageShape && isCustomLargeShapeType(largeShapeType);
    var keepImageCornerFields = planSizeMode === "大きなもの";
    var imageTransformFields = extractImageTransformFieldsFromFormData(formData);
    var customLargeImageName = isCustomLargeImageShape ? normalizeCustomLargeImageName(value(formData.get("customLargeImageName"))) : "";
    var customLargeImageDataUrl = isCustomLargeImageShape ? normalizeCustomLargeImageDataUrl(value(formData.get("customLargeImageDataUrl"))) : "";
    var recordBase = {
      id: recordId,
      kuwaku: recordKuwaku,
      specimenPrefix: specimenPrefix,
      specimenSerial: specimenSerial,
      specimenNo: specimenNo,
      category: categoryFromPrefix(specimenPrefix),
      analysisType: analysisType,
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
      altitudeInputEnabled: altitudeInputEnabled,
      altitudeDirectM: altitudeDirectM,
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
      multiPoints: planMultiPoints,
      planSizeMode: planSizeMode,
      largeShapeType: largeShapeType,
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
      customLargeImageName: customLargeImageName,
      customLargeImageDataUrl: customLargeImageDataUrl,
      customLargeImageAspect: isCustomLargeImageShape ? imageTransformFields.customLargeImageAspect : "",
      importantFlag: normalizeHasFlag(value(formData.get("importantFlag"))),
      simpleRecordFlag: normalizeCircleDashFlag(value(formData.get("simpleRecordFlag"))),
      layerName: getSelectedLayerName(),
      detail: compactNoSpaceValue(formData.get("detail")),
      detailSub: value(formData.get("detailSub")),
      layerFacies: value(formData.get("layerFacies")),
      layerRef: value(formData.get("layerRef")),
      layerFromCm: value(formData.get("layerFromCm")),
      layerRelative: value(formData.get("layerRelative")),
      notes: value(formData.get("notes")),
      sectionDiagrams: clonePhotos(currentSectionDiagrams),
      photos: clonePhotos(currentPhotos),
      createdAt: (found === null || found === void 0 ? void 0 : found.createdAt) || nowIsoValue,
      updatedAt: nowIsoValue,
      deletedAt: ""
    };
    var targetIndex = isEditTab ? state.records.findIndex(function (item) {
      return item === found;
    }) : state.records.findIndex(function (item) {
      return item.id === recordBase.id;
    });
    var previousRecord = targetIndex >= 0 ? state.records[targetIndex] : null;
    var historyAction = isEditTab ? "上書き保存" : targetIndex >= 0 ? "更新保存" : "新規保存";
    var record = _objectSpread(_objectSpread({}, recordBase), {}, {
      history: buildNextRecordHistory(previousRecord, recordBase, historyAction)
    });
    var missingRequiredKeys = getMissingRequiredKeys(record);
    if (missingRequiredKeys.size > 0) {
      var keepSavingIncomplete = window.confirm("未記入がありますが保存しますか？");
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
      markOverwriteUpdatedState(found, record, value(found === null || found === void 0 ? void 0 : found.kuwaku), recordKuwaku);
      overwriteOriginalRecord = _objectSpread({}, record);
      var savedRecordIndex = state.records.findIndex(function (item) {
        return item === record;
      });
      activeEditRecordContext = {
        recordId: value(record.id),
        recordIndex: String(savedRecordIndex >= 0 ? savedRecordIndex : ""),
        recordSnapshot: buildCellEditRecordSnapshot(record)
      };
      renderEditHistory(record);
      updateEditMissingRequiredHighlights();
      return;
    }
    var carryForward = {
      layerName: record.layerName,
      unit: record.unit,
      detail: record.detail,
      detailSub: record.detailSub,
      layerFacies: record.layerFacies,
      layerRef: record.layerRef,
      layerFromCm: record.layerFromCm,
      layerRelative: record.layerRelative
    };
    persist("記録を保存しました");
    renderRecordTable();
    renderOutputs();
    resetRecordForm({
      showMessage: false
    });
    applyCarryForwardFields(carryForward);
    markCarryForwardSavedFields(carryForward);
  });
  recordResetBtn.addEventListener("click", function () {
    resetRecordForm({
      showMessage: true
    });
  });
  if (recordNewBtn) {
    recordNewBtn.addEventListener("click", function () {
      var shouldClear = window.confirm("詳細データは全て消去されます。よろしいですか？");
      if (!shouldClear) {
        return;
      }
      resetRecordForm({
        showMessage: true
      });
    });
  }
  if (recordPrevBtn) {
    recordPrevBtn.addEventListener("click", function () {
      moveToPreviousSpecimenWithoutSave();
    });
  }
  if (recordNextBtn) {
    recordNextBtn.addEventListener("click", function () {
      moveToNextSpecimenWithoutSave();
    });
  }
  if (recordCopyToInputBtn) {
    recordCopyToInputBtn.addEventListener("click", function () {
      copyCurrentEditToInput();
    });
  }
  if (recordCopyToInputTopBtn) {
    recordCopyToInputTopBtn.addEventListener("click", function () {
      copyCurrentEditToInput();
    });
  }
  if (recordTableBody) {
    recordTableBody.addEventListener("click", handleRecordTableActionClick);
  }
  if (editRecordTableBody) {
    editRecordTableBody.addEventListener("click", handleRecordTableActionClick);
  }
  outputListBody.addEventListener("click", function (event) {
    var _row$dataset;
    var button = event.target.closest("button[data-action]");
    if (!button) {
      return;
    }
    var action = button.dataset.action;
    var recordId = button.dataset.id;
    var row = button.closest("tr[data-record-index]");
    var recordIndex = value(button.dataset.recordIndex) || value(row === null || row === void 0 || (_row$dataset = row.dataset) === null || _row$dataset === void 0 ? void 0 : _row$dataset.recordIndex);
    var record = findRecordByEditContext(recordId, recordIndex, null);
    if (action === "edit") {
      var rowKuwaku = value(button.dataset.kuwaku);
      openRecordForEdit(recordId, rowKuwaku, recordIndex);
      return;
    }
    if (action === "copy-to-input") {
      if (!record) {
        showToast("対象データが見つかりません");
        return;
      }
      var _rowKuwaku = value(button.dataset.kuwaku);
      copySavedRecordToInput(value(record.id) || recordId, _rowKuwaku, record);
      return;
    }
    if (action === "insert-row") {
      if (!record) {
        showToast("対象データが見つかりません");
        return;
      }
      var _rowKuwaku2 = value(button.dataset.kuwaku);
      insertRowFromList(value(record.id) || recordId, _rowKuwaku2, record);
      return;
    }
    if (action === "delete") {
      if (!record) {
        showToast("対象データが見つかりません");
        return;
      }
      var answer = window.confirm("\u6A19\u672C\u756A\u53F7 ".concat(record.specimenNo, " \u3092\u524A\u9664\u3057\u307E\u3059\u304B\uFF1F"));
      if (!answer) {
        return;
      }
      var deletingId = value(record.id) || recordId;
      state.records = state.records.filter(function (item) {
        return item !== record;
      });
      if (editingRecordId === deletingId) {
        resetRecordForm({
          showMessage: false
        });
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
  outputListBody.addEventListener("dblclick", function (event) {
    var _cell$dataset, _row$dataset2, _row$dataset3;
    var target = event.target;
    if (!(target instanceof Element)) {
      return;
    }
    if (target.closest("button")) {
      return;
    }
    var cell = target.closest("td[data-cell-edit-key]");
    var row = target.closest("tr[data-record-id]");
    var editKey = value(cell === null || cell === void 0 || (_cell$dataset = cell.dataset) === null || _cell$dataset === void 0 ? void 0 : _cell$dataset.cellEditKey);
    var recordId = value(row === null || row === void 0 || (_row$dataset2 = row.dataset) === null || _row$dataset2 === void 0 ? void 0 : _row$dataset2.recordId);
    var recordIndex = value(row === null || row === void 0 || (_row$dataset3 = row.dataset) === null || _row$dataset3 === void 0 ? void 0 : _row$dataset3.recordIndex);
    if (!cell || !row || !editKey || !recordId || !outputListBody.contains(row)) {
      return;
    }
    event.preventDefault();
    openOutputCellEditModal(recordId, editKey, recordIndex);
  });
  if (outputListTable) {
    var handleSortHeader = function handleSortHeader(target) {
      if (!(target instanceof Element)) {
        return;
      }
      var header = target.closest("th[data-sort-key]");
      if (!header || !outputListTable.contains(header)) {
        return;
      }
      var sortKey = value(header.dataset.sortKey);
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
    outputListTable.addEventListener("click", function (event) {
      handleSortHeader(event.target);
    });
    outputListTable.addEventListener("keydown", function (event) {
      if (event.key !== "Enter" && event.key !== " ") {
        return;
      }
      event.preventDefault();
      handleSortHeader(event.target);
    });
  }
  if (outputKuwakuSelect) {
    outputKuwakuSelect.addEventListener("change", function () {
      selectedOutputKuwaku = value(outputKuwakuSelect.value) || ALL_GRIDS_VALUE;
      selectedCardRecordId = "";
      rememberOutputFilters();
      renderOutputs();
    });
  }
  if (outputCategorySelect) {
    outputCategorySelect.addEventListener("change", function () {
      selectedOutputCategory = value(outputCategorySelect.value) || EXPORT_CATEGORY_ALL_VALUE;
      selectedCardRecordId = "";
      rememberOutputFilters();
      renderOutputs();
    });
  }
  if (outputStatusSelect) {
    outputStatusSelect.addEventListener("change", function () {
      selectedOutputStatus = value(outputStatusSelect.value) || "all";
      selectedCardRecordId = "";
      rememberOutputFilters();
      renderOutputs();
    });
  }
  if (outputDateSelect) {
    outputDateSelect.addEventListener("change", function () {
      selectedOutputDate = value(outputDateSelect.value);
      selectedCardRecordId = "";
      rememberOutputFilters();
      renderOutputs();
    });
  }
  if (outputSearchInput) {
    outputSearchInput.addEventListener("input", function () {
      outputSearchText = value(outputSearchInput.value);
      selectedCardRecordId = "";
      rememberOutputFilters();
      renderOutputs();
    });
  }
  if (outputColumnToggleRow) {
    outputColumnToggleRow.addEventListener("click", function (event) {
      var target = event.target;
      if (!(target instanceof Element)) {
        return;
      }
      var button = target.closest("button[data-col-key]");
      if (!button || !outputColumnToggleRow.contains(button)) {
        return;
      }
      toggleOutputListColumnVisibility(value(button.dataset.colKey));
    });
  }
  if (exportListRangeKuwakuSelect) {
    exportListRangeKuwakuSelect.addEventListener("change", function () {
      exportListRangeKuwaku = value(exportListRangeKuwakuSelect.value) || ALL_GRIDS_VALUE;
      renderExportOutput();
    });
  }
  if (exportListRangeCategorySelect) {
    exportListRangeCategorySelect.addEventListener("change", function () {
      exportListRangeCategory = value(exportListRangeCategorySelect.value) || EXPORT_CATEGORY_ALL_VALUE;
      renderExportOutput();
    });
  }
  if (exportListRangeStatusSelect) {
    exportListRangeStatusSelect.addEventListener("change", function () {
      exportListRangeStatus = value(exportListRangeStatusSelect.value) || "all";
      renderExportOutput();
    });
  }
  if (exportListRangeSpecimenFromInput) {
    exportListRangeSpecimenFromInput.addEventListener("input", function () {
      exportListRangeSpecimenFrom = value(exportListRangeSpecimenFromInput.value);
      renderExportOutput();
    });
  }
  if (exportListRangeSpecimenToInput) {
    exportListRangeSpecimenToInput.addEventListener("input", function () {
      exportListRangeSpecimenTo = value(exportListRangeSpecimenToInput.value);
      renderExportOutput();
    });
  }
  if (exportListRangeDateFromInput) {
    exportListRangeDateFromInput.addEventListener("input", function () {
      exportListRangeDateFrom = value(exportListRangeDateFromInput.value);
      renderExportOutput();
    });
  }
  if (exportListRangeDateToInput) {
    exportListRangeDateToInput.addEventListener("input", function () {
      exportListRangeDateTo = value(exportListRangeDateToInput.value);
      renderExportOutput();
    });
  }
  if (exportCardRangeKuwakuSelect) {
    exportCardRangeKuwakuSelect.addEventListener("change", function () {
      exportCardRangeKuwaku = value(exportCardRangeKuwakuSelect.value) || ALL_GRIDS_VALUE;
      renderExportOutput();
    });
  }
  if (exportCardRangeCategorySelect) {
    exportCardRangeCategorySelect.addEventListener("change", function () {
      exportCardRangeCategory = value(exportCardRangeCategorySelect.value) || EXPORT_CATEGORY_ALL_VALUE;
      renderExportOutput();
    });
  }
  if (exportCardRangeStatusSelect) {
    exportCardRangeStatusSelect.addEventListener("change", function () {
      exportCardRangeStatus = value(exportCardRangeStatusSelect.value) || "all";
      renderExportOutput();
    });
  }
  if (exportCardRangeDateFromInput) {
    exportCardRangeDateFromInput.addEventListener("input", function () {
      exportCardRangeDateFrom = value(exportCardRangeDateFromInput.value);
      renderExportOutput();
    });
  }
  if (exportCardRangeDateToInput) {
    exportCardRangeDateToInput.addEventListener("input", function () {
      exportCardRangeDateTo = value(exportCardRangeDateToInput.value);
      renderExportOutput();
    });
  }
  if (planKuwakuSelect) {
    planKuwakuSelect.addEventListener("change", function () {
      selectedPlanKuwaku = value(planKuwakuSelect.value);
      renderPlanOutput();
    });
  }
  if (planCategorySelect) {
    planCategorySelect.addEventListener("change", function () {
      selectedPlanCategory = value(planCategorySelect.value) || EXPORT_CATEGORY_ALL_VALUE;
      renderPlanOutput();
    });
  }
  if (planUnitSelect) {
    planUnitSelect.addEventListener("change", function () {
      selectedPlanUnit = value(planUnitSelect.value);
      renderPlanOutput();
    });
  }
  if (planDetailSelect) {
    planDetailSelect.addEventListener("change", function () {
      selectedPlanDetail = value(planDetailSelect.value) || ALL_DETAILS_VALUE;
      renderPlanOutput();
    });
  }
  if (planDetailSubSelect) {
    planDetailSubSelect.addEventListener("change", function () {
      selectedPlanDetailSub = value(planDetailSubSelect.value) || ALL_DETAIL_SUBS_VALUE;
      renderPlanOutput();
    });
  }
  if (viewerKuwakuSelect) {
    viewerKuwakuSelect.addEventListener("change", function () {
      selectedViewerKuwaku = value(viewerKuwakuSelect.value) || ALL_GRIDS_VALUE;
      renderViewerOutput();
    });
  }
  if (viewerCategorySelect) {
    viewerCategorySelect.addEventListener("change", function () {
      selectedViewerCategory = value(viewerCategorySelect.value) || EXPORT_CATEGORY_ALL_VALUE;
      renderViewerOutput();
    });
  }
  if (viewerUnitSelect) {
    viewerUnitSelect.addEventListener("change", function () {
      selectedViewerUnit = value(viewerUnitSelect.value) || ALL_UNITS_VALUE;
      renderViewerOutput();
    });
  }
  if (viewerDetailSelect) {
    viewerDetailSelect.addEventListener("change", function () {
      selectedViewerDetail = value(viewerDetailSelect.value) || ALL_DETAILS_VALUE;
      renderViewerOutput();
    });
  }
  if (viewerDetailSubSelect) {
    viewerDetailSubSelect.addEventListener("change", function () {
      selectedViewerDetailSub = value(viewerDetailSubSelect.value) || ALL_DETAIL_SUBS_VALUE;
      renderViewerOutput();
    });
  }
  if (viewerViewTopBtn) {
    viewerViewTopBtn.addEventListener("click", function () {
      selectedViewerPerspective = "top";
      applyViewerPerspective();
      syncViewerViewButtons();
    });
  }
  if (viewerViewSeBtn) {
    viewerViewSeBtn.addEventListener("click", function () {
      selectedViewerPerspective = "se";
      applyViewerPerspective();
      syncViewerViewButtons();
    });
  }
  if (viewerViewEastBtn) {
    viewerViewEastBtn.addEventListener("click", function () {
      selectedViewerPerspective = "east";
      applyViewerPerspective();
      syncViewerViewButtons();
    });
  }
  if (viewerViewWestBtn) {
    viewerViewWestBtn.addEventListener("click", function () {
      selectedViewerPerspective = "west";
      applyViewerPerspective();
      syncViewerViewButtons();
    });
  }
  if (viewerViewSouthBtn) {
    viewerViewSouthBtn.addEventListener("click", function () {
      selectedViewerPerspective = "south";
      applyViewerPerspective();
      syncViewerViewButtons();
    });
  }
  if (viewerViewNorthBtn) {
    viewerViewNorthBtn.addEventListener("click", function () {
      selectedViewerPerspective = "north";
      applyViewerPerspective();
      syncViewerViewButtons();
    });
  }
  if (viewerZScaleInput) {
    viewerZScaleInput.addEventListener("input", function () {
      viewerVerticalScale = normalizeViewerVerticalScale(viewerZScaleInput.value);
      syncViewerVerticalScaleUi();
      if (viewer3d.initialized) {
        renderViewerOutput();
      }
    });
  }
  if (exportPlanKuwakuSelect) {
    exportPlanKuwakuSelect.addEventListener("change", function () {
      exportPlanKuwaku = value(exportPlanKuwakuSelect.value);
      renderExportOutput();
    });
  }
  if (exportPlanCategorySelect) {
    exportPlanCategorySelect.addEventListener("change", function () {
      exportPlanCategory = value(exportPlanCategorySelect.value) || EXPORT_CATEGORY_ALL_VALUE;
      renderExportOutput();
    });
  }
  if (exportPlanDateFromInput) {
    exportPlanDateFromInput.addEventListener("input", function () {
      exportPlanDateFrom = value(exportPlanDateFromInput.value);
      renderExportOutput();
    });
  }
  if (exportPlanDateToInput) {
    exportPlanDateToInput.addEventListener("input", function () {
      exportPlanDateTo = value(exportPlanDateToInput.value);
      renderExportOutput();
    });
  }
  if (exportPlanModeUnitCheck) {
    exportPlanModeUnitCheck.addEventListener("change", function () {
      exportPlanModeUnitEnabled = Boolean(exportPlanModeUnitCheck.checked);
      renderExportOutput();
    });
  }
  if (exportPlanModeUnitButtons) {
    exportPlanModeUnitButtons.addEventListener("click", function (event) {
      var button = event.target.closest("button[data-value]");
      if (!(button instanceof HTMLButtonElement)) {
        return;
      }
      var optionValue = value(button.dataset.value);
      exportPlanModeUnitTouched = true;
      if (optionValue === EXPORT_PLAN_ALL_UNITS_BUTTON_VALUE) {
        var unitValues = Array.from(collectExportPlanValueOptions(getExportPlanScopedRecords(), function (record) {
          return unitValueForSelect(record.unit);
        }, unitLabelForSelect)).map(function (item) {
          return value(item.value);
        }).filter(Boolean);
        var allSelected = unitValues.length > 0 && unitValues.every(function (unitValue) {
          return exportPlanModeUnitValues.has(unitValue);
        });
        exportPlanModeUnitValues = allSelected ? new Set() : new Set(unitValues);
      } else {
        toggleSelectionInSet(exportPlanModeUnitValues, optionValue);
      }
      renderExportOutput();
    });
  }
  if (exportPlanModeDetailCheck) {
    exportPlanModeDetailCheck.addEventListener("change", function () {
      exportPlanModeDetailEnabled = Boolean(exportPlanModeDetailCheck.checked);
      renderExportOutput();
    });
  }
  if (exportPlanModeDetailUnitSelect) {
    exportPlanModeDetailUnitSelect.addEventListener("change", function () {
      exportPlanModeDetailUnitValue = value(exportPlanModeDetailUnitSelect.value);
      exportPlanModeDetailTouched = false;
      renderExportOutput();
    });
  }
  if (exportPlanModeDetailButtons) {
    exportPlanModeDetailButtons.addEventListener("click", function (event) {
      var button = event.target.closest("button[data-value]");
      if (!(button instanceof HTMLButtonElement)) {
        return;
      }
      exportPlanModeDetailTouched = true;
      toggleSelectionInSet(exportPlanModeDetailValues, value(button.dataset.value));
      renderExportOutput();
    });
  }
  if (exportPlanModeDetailSubCheck) {
    exportPlanModeDetailSubCheck.addEventListener("change", function () {
      exportPlanModeDetailSubEnabled = Boolean(exportPlanModeDetailSubCheck.checked);
      renderExportOutput();
    });
  }
  if (exportPlanModeDetailSubUnitSelect) {
    exportPlanModeDetailSubUnitSelect.addEventListener("change", function () {
      exportPlanModeDetailSubUnitValue = value(exportPlanModeDetailSubUnitSelect.value);
      exportPlanModeDetailSubTouched = false;
      renderExportOutput();
    });
  }
  if (exportPlanModeDetailSubDetailSelect) {
    exportPlanModeDetailSubDetailSelect.addEventListener("change", function () {
      exportPlanModeDetailSubDetailValue = value(exportPlanModeDetailSubDetailSelect.value);
      exportPlanModeDetailSubTouched = false;
      renderExportOutput();
    });
  }
  if (exportPlanModeDetailSubButtons) {
    exportPlanModeDetailSubButtons.addEventListener("click", function (event) {
      var button = event.target.closest("button[data-value]");
      if (!(button instanceof HTMLButtonElement)) {
        return;
      }
      exportPlanModeDetailSubTouched = true;
      toggleSelectionInSet(exportPlanModeDetailSubValues, value(button.dataset.value));
      renderExportOutput();
    });
  }
  if (editKuwakuHeadAInput) {
    editKuwakuHeadAInput.addEventListener("input", function () {
      editKuwakuHeadAInput.classList.remove("overwrite-updated");
      updateDuplicateSpecimenWarning();
    });
  }
  if (editKuwakuHeadBInput) {
    editKuwakuHeadBInput.addEventListener("input", function () {
      editKuwakuHeadBInput.classList.remove("overwrite-updated");
      updateDuplicateSpecimenWarning();
    });
  }
  if (editKuwakuBlockInput) {
    editKuwakuBlockInput.addEventListener("input", function () {
      editKuwakuBlockInput.classList.remove("overwrite-updated");
      updateDuplicateSpecimenWarning();
    });
  }
  if (editKuwakuNoInput) {
    editKuwakuNoInput.addEventListener("input", function () {
      editKuwakuNoInput.classList.remove("overwrite-updated");
      updateDuplicateSpecimenWarning();
    });
  }
  if (editLevelHeightInput) {
    editLevelHeightInput.addEventListener("input", function () {
      editLevelHeightInput.classList.remove("overwrite-updated");
    });
  }
  if (editDateInput) {
    editDateInput.addEventListener("input", function () {
      editDateInput.classList.remove("overwrite-updated");
    });
  }
  if (editTeamOtherInput) {
    editTeamOtherInput.addEventListener("input", function () {
      editTeamOtherInput.classList.remove("overwrite-updated");
    });
  }
  if (editTeamLeadInput) {
    editTeamLeadInput.addEventListener("input", function () {
      editTeamLeadInput.classList.remove("overwrite-updated");
    });
  }
  if (editRecorderInput) {
    editRecorderInput.addEventListener("input", function () {
      editRecorderInput.classList.remove("overwrite-updated");
    });
  }
  var sectionDiagramInputs = [sectionDiagramCameraInput, sectionDiagramInput].filter(Boolean);
  sectionDiagramInputs.forEach(function (input) {
    input.addEventListener("change", function () {
      var _ref2 = _asyncToGenerator(_regenerator().m(function _callee2(event) {
        return _regenerator().w(function (_context2) {
          while (1) switch (_context2.n) {
            case 0:
              _context2.n = 1;
              return addSectionDiagramsFromFiles(event.target.files);
            case 1:
              event.target.value = "";
            case 2:
              return _context2.a(2);
          }
        }, _callee2);
      }));
      return function (_x2) {
        return _ref2.apply(this, arguments);
      };
    }());
  });
  sectionDiagramList.addEventListener("input", function (event) {
    var input = event.target.closest("input[data-diagram-id]");
    if (!input) {
      return;
    }
    var target = currentSectionDiagrams.find(function (item) {
      return item.id === input.dataset.diagramId;
    });
    if (!target) {
      return;
    }
    target.caption = input.value;
    persist();
  });
  sectionDiagramList.addEventListener("click", function (event) {
    var button = event.target.closest("button[data-remove-diagram-id]");
    if (!button) {
      return;
    }
    currentSectionDiagrams = currentSectionDiagrams.filter(function (item) {
      return item.id !== button.dataset.removeDiagramId;
    });
    renderSectionDiagramList();
    persist("断面図を削除しました");
  });
  var photoInputs = [photoCameraInput, photoInput].filter(Boolean);
  photoInputs.forEach(function (input) {
    input.addEventListener("change", function () {
      var _ref3 = _asyncToGenerator(_regenerator().m(function _callee3(event) {
        return _regenerator().w(function (_context3) {
          while (1) switch (_context3.n) {
            case 0:
              _context3.n = 1;
              return addPhotosFromFiles(event.target.files);
            case 1:
              event.target.value = "";
            case 2:
              return _context3.a(2);
          }
        }, _callee3);
      }));
      return function (_x3) {
        return _ref3.apply(this, arguments);
      };
    }());
  });
  photoList.addEventListener("input", function (event) {
    var input = event.target.closest("input[data-photo-id]");
    if (!input) {
      return;
    }
    var target = currentPhotos.find(function (photo) {
      return photo.id === input.dataset.photoId;
    });
    if (!target) {
      return;
    }
    target.caption = input.value;
    persist();
  });
  photoList.addEventListener("click", function (event) {
    var button = event.target.closest("button[data-remove-photo-id]");
    if (!button) {
      return;
    }
    currentPhotos = currentPhotos.filter(function (photo) {
      return photo.id !== button.dataset.removePhotoId;
    });
    renderPhotoList();
    persist("写真を削除しました");
  });
  exportListCsvBtn.addEventListener("click", function () {
    if (!getListExportRecords().length) {
      showToast("CSV出力対象データがありません");
      return;
    }
    var csv = buildListCsv();
    downloadFile("nojiri-kaseki-list-".concat(timestamp(), ".csv"), csv, "text/csv;charset=utf-8");
    showToast("リストCSVを書き出しました");
  });
  exportCardCsvBtn.addEventListener("click", function () {
    if (!getCardExportRecords().length) {
      showToast("カードCSVの出力対象データがありません");
      return;
    }
    var csv = buildCardCsv();
    downloadFile("nojiri-kaseki-card-".concat(timestamp(), ".csv"), csv, "text/csv;charset=utf-8");
    showToast("カードCSVを書き出しました");
  });
  if (exportListPdfBtn) {
    exportListPdfBtn.addEventListener("click", function () {
      exportListPdf();
    });
  }
  if (exportCardPdfBtn) {
    exportCardPdfBtn.addEventListener("click", function () {
      exportCardPdf();
    });
  }
  if (exportPlanPdfBtn) {
    exportPlanPdfBtn.addEventListener("click", function () {
      exportPlanPdf();
    });
  }
  exportJsonBtn.addEventListener("click", function () {
    var json = JSON.stringify(state, null, 2);
    downloadFile("nojiri-kaseki-".concat(timestamp(), ".json"), json, "application/json");
    showToast("JSONを書き出しました");
  });
  importJsonInput.addEventListener("change", function () {
    var _ref4 = _asyncToGenerator(_regenerator().m(function _callee4(event) {
      var _Array$from, _Array$from2, file, text, imported, _t;
      return _regenerator().w(function (_context4) {
        while (1) switch (_context4.p = _context4.n) {
          case 0:
            _Array$from = Array.from(event.target.files || []), _Array$from2 = _slicedToArray(_Array$from, 1), file = _Array$from2[0];
            if (file) {
              _context4.n = 1;
              break;
            }
            return _context4.a(2);
          case 1:
            _context4.p = 1;
            _context4.n = 2;
            return file.text();
          case 2:
            text = _context4.v;
            imported = JSON.parse(text);
            state = normalizeState(imported);
            hydrateSiteForm();
            resetRecordForm({
              showMessage: false
            });
            renderRecordTable();
            renderOutputs();
            persist("JSONを読み込みました");
            _context4.n = 4;
            break;
          case 3:
            _context4.p = 3;
            _t = _context4.v;
            showToast("JSON読み込みに失敗しました");
          case 4:
            _context4.p = 4;
            importJsonInput.value = "";
            return _context4.f(4);
          case 5:
            return _context4.a(2);
        }
      }, _callee4, null, [[1, 3, 4, 5]]);
    }));
    return function (_x4) {
      return _ref4.apply(this, arguments);
    };
  }());
}
function addSectionDiagramsFromFiles(_x5) {
  return _addSectionDiagramsFromFiles.apply(this, arguments);
}
function _addSectionDiagramsFromFiles() {
  _addSectionDiagramsFromFiles = _asyncToGenerator(_regenerator().m(function _callee5(fileList) {
    var files, added, _i8, _files, file, dataUrl, _t2;
    return _regenerator().w(function (_context5) {
      while (1) switch (_context5.p = _context5.n) {
        case 0:
          files = Array.from(fileList || []);
          if (files.length) {
            _context5.n = 1;
            break;
          }
          return _context5.a(2);
        case 1:
          added = false;
          _i8 = 0, _files = files;
        case 2:
          if (!(_i8 < _files.length)) {
            _context5.n = 7;
            break;
          }
          file = _files[_i8];
          _context5.p = 3;
          _context5.n = 4;
          return loadImageFileDataUrlWithFallback(file, {
            maxLength: 1280,
            quality: 0.72,
            mimeType: "image/jpeg"
          });
        case 4:
          dataUrl = _context5.v;
          currentSectionDiagrams.push({
            id: newId("diagram"),
            name: file.name || "diagram.jpg",
            dataUrl: dataUrl,
            caption: "",
            createdAt: new Date().toISOString()
          });
          added = true;
          _context5.n = 6;
          break;
        case 5:
          _context5.p = 5;
          _t2 = _context5.v;
          showToast("\u65AD\u9762\u56F3\u8FFD\u52A0\u306B\u5931\u6557: ".concat(file.name));
        case 6:
          _i8++;
          _context5.n = 2;
          break;
        case 7:
          if (added) {
            _context5.n = 8;
            break;
          }
          return _context5.a(2);
        case 8:
          renderSectionDiagramList();
          updateEditMissingRequiredHighlights();
          persist();
        case 9:
          return _context5.a(2);
      }
    }, _callee5, null, [[3, 5]]);
  }));
  return _addSectionDiagramsFromFiles.apply(this, arguments);
}
function addPhotosFromFiles(_x6) {
  return _addPhotosFromFiles.apply(this, arguments);
}
function _addPhotosFromFiles() {
  _addPhotosFromFiles = _asyncToGenerator(_regenerator().m(function _callee6(fileList) {
    var files, added, _i9, _files2, file, dataUrl, _t3;
    return _regenerator().w(function (_context6) {
      while (1) switch (_context6.p = _context6.n) {
        case 0:
          files = Array.from(fileList || []);
          if (files.length) {
            _context6.n = 1;
            break;
          }
          return _context6.a(2);
        case 1:
          added = false;
          _i9 = 0, _files2 = files;
        case 2:
          if (!(_i9 < _files2.length)) {
            _context6.n = 7;
            break;
          }
          file = _files2[_i9];
          _context6.p = 3;
          _context6.n = 4;
          return loadImageFileDataUrlWithFallback(file, {
            maxLength: 1280,
            quality: 0.72,
            mimeType: "image/jpeg"
          });
        case 4:
          dataUrl = _context6.v;
          currentPhotos.push({
            id: newId("photo"),
            name: file.name || "photo.jpg",
            dataUrl: dataUrl,
            caption: "",
            createdAt: new Date().toISOString()
          });
          added = true;
          _context6.n = 6;
          break;
        case 5:
          _context6.p = 5;
          _t3 = _context6.v;
          showToast("\u5199\u771F\u8FFD\u52A0\u306B\u5931\u6557: ".concat(file.name));
        case 6:
          _i9++;
          _context6.n = 2;
          break;
        case 7:
          if (added) {
            _context6.n = 8;
            break;
          }
          return _context6.a(2);
        case 8:
          renderPhotoList();
          persist();
        case 9:
          return _context6.a(2);
      }
    }, _callee6, null, [[3, 5]]);
  }));
  return _addPhotosFromFiles.apply(this, arguments);
}
function setActiveTab(tabId) {
  var previousTabId = getActiveTabId();
  hideHoverEditMenu();
  closeOutputCellEditModal();
  if (previousTabId === "output-tab") {
    rememberOutputFilters();
  }
  tabButtons.forEach(function (button) {
    button.classList.toggle("active", button.dataset.tab === tabId);
  });
  tabPanels.forEach(function (panel) {
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
    renderViewerOutput({
      preserveCamera: true
    });
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
    void pullStateFromCloud({
      force: false,
      showToastOnSuccess: false,
      silentOnError: true
    });
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
      resetRecordForm({
        showMessage: false
      });
    }
    recordSubmitBtn.textContent = "記録を保存";
  }
}
function getActiveTabId() {
  var activePanel = document.querySelector(".tab-panel.active");
  return (activePanel === null || activePanel === void 0 ? void 0 : activePanel.id) || "";
}
function rememberOutputFilters() {
  outputFilterMemory = {
    kuwaku: value(selectedOutputKuwaku) || ALL_GRIDS_VALUE,
    category: value(selectedOutputCategory) || EXPORT_CATEGORY_ALL_VALUE,
    status: ["all", "complete", "incomplete"].includes(value(selectedOutputStatus)) ? value(selectedOutputStatus) : "all",
    date: value(selectedOutputDate),
    searchText: value(outputSearchText)
  };
}
function restoreOutputFilters() {
  selectedOutputKuwaku = value(outputFilterMemory.kuwaku) || ALL_GRIDS_VALUE;
  selectedOutputCategory = value(outputFilterMemory.category) || EXPORT_CATEGORY_ALL_VALUE;
  selectedOutputStatus = ["all", "complete", "incomplete"].includes(value(outputFilterMemory.status)) ? value(outputFilterMemory.status) : "all";
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
  var teamState = normalizeTeamState(state.site.team, state.site.teamOther);
  siteForm.elements.team.value = teamState.team || "";
  siteForm.elements.teamOther.value = teamState.teamOther || "";
  syncTeamOtherInput(siteForm.elements.team.value);
  siteForm.elements.teamLead.value = state.site.teamLead || "";
  siteForm.elements.recorder.value = state.site.recorder || "";
  if (siteForm.elements.scribe) {
    siteForm.elements.scribe.value = state.site.scribe || "";
  }
}
function loadTeamRosterFromDefaultFile() {
  return _loadTeamRosterFromDefaultFile.apply(this, arguments);
}
function _loadTeamRosterFromDefaultFile() {
  _loadTeamRosterFromDefaultFile = _asyncToGenerator(_regenerator().m(function _callee7() {
    var rows, _t4;
    return _regenerator().w(function (_context7) {
      while (1) switch (_context7.p = _context7.n) {
        case 0:
          setTeamRosterStatus("名簿を読込中…");
          _context7.p = 1;
          rows = Array.isArray(window.__TEAM_ROSTER_ROWS__) ? window.__TEAM_ROSTER_ROWS__ : [];
          teamRosterAssignmentMap = buildTeamRosterAssignmentMap(rows);
          if (teamRosterAssignmentMap.size) {
            _context7.n = 2;
            break;
          }
          throw new Error("roster empty");
        case 2:
          teamRosterLoaded = true;
          setTeamRosterStatus("名簿を自動読込しました");
          applyTeamRosterAutofill();
          _context7.n = 4;
          break;
        case 3:
          _context7.p = 3;
          _t4 = _context7.v;
          setTeamRosterStatus("\u540D\u7C3F\u306E\u81EA\u52D5\u8AAD\u8FBC\u306B\u5931\u6557: ".concat(TEAM_ROSTER_FILE_NAME));
        case 4:
          return _context7.a(2);
      }
    }, _callee7, null, [[1, 3]]);
  }));
  return _loadTeamRosterFromDefaultFile.apply(this, arguments);
}
function buildTeamRosterAssignmentMap(rows) {
  var assignmentMap = new Map();
  if (!Array.isArray(rows) || rows.length === 0) {
    return assignmentMap;
  }
  var headerRow = Array.isArray(rows[0]) ? rows[0] : [];
  var dayBlocks = buildRosterDayBlocks(headerRow);
  for (var rowIndex = 1; rowIndex < rows.length; rowIndex += 1) {
    var row = Array.isArray(rows[rowIndex]) ? rows[rowIndex] : [];
    var name = value(row[TEAM_ROSTER_NAME_COL - 1]);
    var team = normalizeRosterTeam(row[TEAM_ROSTER_TEAM_COL - 1]);
    var roleText = value(row[TEAM_ROSTER_ROLE_COL - 1]);
    if (!name || !team || !roleText) {
      continue;
    }
    var _iterator = _createForOfIteratorHelper(dayBlocks),
      _step;
    try {
      for (_iterator.s(); !(_step = _iterator.n()).done;) {
        var block = _step.value;
        if (!rowHasAttendanceMark(row, block.startCol, block.endCol)) {
          continue;
        }
        var key = "".concat(block.day, "-").concat(team);
        if (!assignmentMap.has(key)) {
          assignmentMap.set(key, {
            teamLeads: new Set(),
            recorders: new Set()
          });
        }
        var target = assignmentMap.get(key);
        if (/班長/.test(roleText)) {
          target.teamLeads.add(name);
        }
        if (/記載/.test(roleText)) {
          target.recorders.add(name);
        }
      }
    } catch (err) {
      _iterator.e(err);
    } finally {
      _iterator.f();
    }
  }
  return assignmentMap;
}
function buildRosterDayBlocks(headerRow) {
  var blocks = [];
  var current = null;
  for (var col = TEAM_ROSTER_DAY_COL_START; col <= TEAM_ROSTER_DAY_COL_END; col += 1) {
    var text = value(headerRow[col - 1]);
    var match = text.match(/(\d{1,2})日/);
    if (match) {
      if (current) {
        current.endCol = col - 1;
        blocks.push(current);
      }
      current = {
        day: Number(match[1]),
        startCol: col,
        endCol: col
      };
    }
  }
  if (current) {
    current.endCol = TEAM_ROSTER_DAY_COL_END;
    blocks.push(current);
  }
  return blocks;
}
function rowHasAttendanceMark(row, startCol, endCol) {
  for (var col = startCol; col <= endCol; col += 1) {
    var mark = value(row[col - 1]).replace(/\s+/g, "");
    if (mark === "○" || mark === "◯") {
      return true;
    }
  }
  return false;
}
function normalizeRosterTeam(teamRaw) {
  var text = value(teamRaw);
  if (!text) {
    return "";
  }
  var normalized = text.replace(/[^\d]/g, "");
  return normalized || text;
}
function setTeamRosterStatus(text) {
  if (rosterStatusEl) {
    rosterStatusEl.textContent = text;
  }
}
function applyTeamRosterAutofill() {
  var _siteForm$elements$te, _siteForm$elements$da;
  if (!(siteForm !== null && siteForm !== void 0 && siteForm.elements)) {
    return;
  }
  var team = value((_siteForm$elements$te = siteForm.elements.team) === null || _siteForm$elements$te === void 0 ? void 0 : _siteForm$elements$te.value);
  var date = value((_siteForm$elements$da = siteForm.elements.date) === null || _siteForm$elements$da === void 0 ? void 0 : _siteForm$elements$da.value);
  if (!teamRosterLoaded || !team || !date || team === OTHER_TEAM_NAME) {
    return;
  }
  var day = extractDayFromIsoDate(date);
  if (!day) {
    return;
  }
  var key = "".concat(day, "-").concat(normalizeRosterTeam(team));
  var assignment = teamRosterAssignmentMap.get(key);
  if (!assignment) {
    return;
  }
  var leads = Array.from(assignment.teamLeads).join(", ");
  var recorders = Array.from(assignment.recorders).join(", ");
  if (siteForm.elements.teamLead) {
    siteForm.elements.teamLead.value = leads;
  }
  if (siteForm.elements.recorder) {
    siteForm.elements.recorder.value = recorders;
  }
}
function extractDayFromIsoDate(dateRaw) {
  var text = value(dateRaw);
  var match = text.match(/^\d{4}-\d{2}-(\d{2})$/);
  if (!match) {
    return null;
  }
  var day = Number(match[1]);
  return Number.isFinite(day) ? day : null;
}
function activateSpecimenPrefix(prefixRaw) {
  var prefix = normalizeSpecimenPrefix(prefixRaw);
  specimenPrefixInput.value = prefix;
  specimenPrefixLabel.textContent = prefix;
  specimenPrefixLabel.dataset.prefix = prefix;
  specimenTabButtons.forEach(function (button) {
    button.classList.toggle("active", normalizeSpecimenPrefix(button.dataset.prefix) === prefix);
  });
  syncAnalysisTypeInput(prefix);
}
function updateSpecimenNoFromParts() {
  var prefix = normalizeSpecimenPrefix(specimenPrefixInput.value);
  var serial = value(specimenSerialInput.value);
  specimenPrefixInput.value = prefix;
  specimenNoInput.value = buildSpecimenNo(prefix, serial);
  updateDuplicateSpecimenWarning();
}
function updateDuplicateSpecimenWarning() {
  if (!specimenDuplicateWarning) {
    return;
  }
  var activeTabId = getActiveTabId();
  if (activeTabId !== "input-tab" && activeTabId !== "edit-tab") {
    hideDuplicateSpecimenWarning();
    return;
  }
  var specimenSerial = compactNoSpaceValue(specimenSerialInput === null || specimenSerialInput === void 0 ? void 0 : specimenSerialInput.value);
  var specimenPrefix = normalizeSpecimenPrefix(specimenPrefixInput === null || specimenPrefixInput === void 0 ? void 0 : specimenPrefixInput.value);
  var specimenNo = buildSpecimenNo(specimenPrefix, specimenSerial);
  if (!specimenNo || !specimenSerial) {
    hideDuplicateSpecimenWarning();
    return;
  }
  var kuwaku = currentKuwakuForDuplicateWarning(activeTabId);
  if (!kuwaku) {
    hideDuplicateSpecimenWarning();
    return;
  }
  var excludeRecordId = activeTabId === "edit-tab" ? value(editingRecordId || (recordIdInput === null || recordIdInput === void 0 ? void 0 : recordIdInput.value)) : "";
  var duplicate = findDuplicateRecordByKuwakuAndSpecimen(kuwaku, specimenNo, excludeRecordId);
  if (!duplicate) {
    hideDuplicateSpecimenWarning();
    return;
  }
  specimenDuplicateWarning.textContent = "\u8B66\u544A: \u3053\u306E\u533A\u753B\u306B\u306F ".concat(specimenNo, " \u304C\u3059\u3067\u306B\u3042\u308A\u307E\u3059");
  specimenDuplicateWarning.classList.remove("hidden");
}
function currentKuwakuForDuplicateWarning() {
  var _siteForm$elements, _siteForm$elements2, _siteForm$elements3, _siteForm$elements4;
  var activeTabId = arguments.length > 0 && arguments[0] !== undefined ? arguments[0] : getActiveTabId();
  if (activeTabId === "edit-tab") {
    var _headA = normalizeKuwakuHeadA(editKuwakuHeadAInput === null || editKuwakuHeadAInput === void 0 ? void 0 : editKuwakuHeadAInput.value);
    var _headB = normalizeKuwakuHeadB(editKuwakuHeadBInput === null || editKuwakuHeadBInput === void 0 ? void 0 : editKuwakuHeadBInput.value);
    var _block = normalizeKuwakuBlock(editKuwakuBlockInput === null || editKuwakuBlockInput === void 0 ? void 0 : editKuwakuBlockInput.value);
    var _no = normalizeKuwakuNo(editKuwakuNoInput === null || editKuwakuNoInput === void 0 ? void 0 : editKuwakuNoInput.value);
    if (!_block || !_no) {
      return "";
    }
    return buildKuwaku(_headA, _headB, _block, _no);
  }
  var headA = normalizeKuwakuHeadA(siteForm === null || siteForm === void 0 || (_siteForm$elements = siteForm.elements) === null || _siteForm$elements === void 0 || (_siteForm$elements = _siteForm$elements.kuwakuHeadA) === null || _siteForm$elements === void 0 ? void 0 : _siteForm$elements.value);
  var headB = normalizeKuwakuHeadB(siteForm === null || siteForm === void 0 || (_siteForm$elements2 = siteForm.elements) === null || _siteForm$elements2 === void 0 || (_siteForm$elements2 = _siteForm$elements2.kuwakuHeadB) === null || _siteForm$elements2 === void 0 ? void 0 : _siteForm$elements2.value);
  var block = normalizeKuwakuBlock(siteForm === null || siteForm === void 0 || (_siteForm$elements3 = siteForm.elements) === null || _siteForm$elements3 === void 0 || (_siteForm$elements3 = _siteForm$elements3.kuwakuBlock) === null || _siteForm$elements3 === void 0 ? void 0 : _siteForm$elements3.value);
  var no = normalizeKuwakuNo(siteForm === null || siteForm === void 0 || (_siteForm$elements4 = siteForm.elements) === null || _siteForm$elements4 === void 0 || (_siteForm$elements4 = _siteForm$elements4.kuwakuNo) === null || _siteForm$elements4 === void 0 ? void 0 : _siteForm$elements4.value);
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
  var activeTabId = getActiveTabId();
  if (activeTabId !== "input-tab" && activeTabId !== "edit-tab") {
    return;
  }
  var currentKuwaku = currentKuwakuForDuplicateWarning(activeTabId);
  if (!currentKuwaku) {
    showToast("区画（グリッド）を入力してください");
    return;
  }
  var currentPrefix = normalizeSpecimenPrefix(specimenPrefixInput === null || specimenPrefixInput === void 0 ? void 0 : specimenPrefixInput.value);
  var currentSerial = compactNoSpaceValue(specimenSerialInput === null || specimenSerialInput === void 0 ? void 0 : specimenSerialInput.value);
  if (!state.records.length) {
    var prevSerial = "1";
    if (currentSerial && /^\d+$/.test(currentSerial)) {
      prevSerial = String(Math.max(1, Number(currentSerial) - 1));
    } else if (currentSerial) {
      prevSerial = currentSerial;
    }
    applyNextNavigationTarget({
      kuwaku: currentKuwaku,
      prefix: currentPrefix,
      serial: prevSerial
    });
    showToast("\u524D\u3078: ".concat(buildSpecimenNo(currentPrefix, prevSerial)));
    return;
  }
  var prevRecord = findPreviousRecordByGridPrefixThenSerial(currentKuwaku, currentPrefix, currentSerial);
  if (!prevRecord) {
    showToast("前のデータが見つかりません");
    return;
  }
  var prevKuwaku = getRecordKuwaku(prevRecord);
  var prevSpecimen = parseSpecimenNo(prevRecord.specimenNo, prevRecord.specimenPrefix, prevRecord.specimenSerial);
  if (activeTabId === "edit-tab") {
    openRecordForEdit(prevRecord.id, prevKuwaku);
    showToast("\u524D\u3078: ".concat(prevKuwaku, " / ").concat(prevSpecimen.specimenNo));
    return;
  }
  loadRecordIntoInputForNavigation(prevRecord, prevKuwaku);
  showToast("\u524D\u3078: ".concat(prevKuwaku, " / ").concat(prevSpecimen.specimenNo));
}
function moveToNextSpecimenWithoutSave() {
  var activeTabId = getActiveTabId();
  if (activeTabId !== "input-tab" && activeTabId !== "edit-tab") {
    return;
  }
  var currentKuwaku = currentKuwakuForDuplicateWarning(activeTabId);
  if (!currentKuwaku) {
    showToast("区画（グリッド）を入力してください");
    return;
  }
  var currentPrefix = normalizeSpecimenPrefix(specimenPrefixInput === null || specimenPrefixInput === void 0 ? void 0 : specimenPrefixInput.value);
  var currentSerial = compactNoSpaceValue(specimenSerialInput === null || specimenSerialInput === void 0 ? void 0 : specimenSerialInput.value);
  if (!state.records.length) {
    var nextSerial = currentSerial ? /^\d+$/.test(currentSerial) ? String(Number(currentSerial) + 1) : "".concat(currentSerial, "1") : "1";
    applyNextNavigationTarget({
      kuwaku: currentKuwaku,
      prefix: currentPrefix,
      serial: nextSerial
    });
    showToast("\u6B21\u3078: ".concat(buildSpecimenNo(currentPrefix, nextSerial)));
    return;
  }
  var nextRecord = findNextRecordByGridPrefixThenSerial(currentKuwaku, currentPrefix, currentSerial);
  if (!nextRecord) {
    showToast("次のデータが見つかりません");
    return;
  }
  var nextKuwaku = getRecordKuwaku(nextRecord);
  var nextSpecimen = parseSpecimenNo(nextRecord.specimenNo, nextRecord.specimenPrefix, nextRecord.specimenSerial);
  if (activeTabId === "edit-tab") {
    openRecordForEdit(nextRecord.id, nextKuwaku);
    showToast("\u6B21\u3078: ".concat(nextKuwaku, " / ").concat(nextSpecimen.specimenNo));
    return;
  }
  loadRecordIntoInputForNavigation(nextRecord, nextKuwaku);
  showToast("\u6B21\u3078: ".concat(nextKuwaku, " / ").concat(nextSpecimen.specimenNo));
}
function findPreviousRecordByGridPrefixThenSerial(currentKuwakuRaw, currentPrefixRaw, currentSerialRaw) {
  var currentKuwakuValue = kuwakuValueForSelect(currentKuwakuRaw);
  var currentPrefix = normalizeSpecimenPrefix(currentPrefixRaw);
  var currentSerial = compactNoSpaceValue(currentSerialRaw);
  var sorted = _toConsumableArray(state.records).sort(compareRecordsByKuwakuThenSpecimen);
  if (!sorted.length) {
    return null;
  }
  var groupedByGrid = new Map();
  var gridOrder = [];
  sorted.forEach(function (record) {
    var gridValue = kuwakuValueForSelect(getRecordKuwaku(record));
    if (!groupedByGrid.has(gridValue)) {
      groupedByGrid.set(gridValue, []);
      gridOrder.push(gridValue);
    }
    groupedByGrid.get(gridValue).push(record);
  });
  var currentGridRecords = groupedByGrid.get(currentKuwakuValue) || [];
  var gridStartIndex = resolvePreviousGridStartIndex(gridOrder, currentKuwakuValue);
  if (!currentGridRecords.length) {
    for (var step = 0; step < gridOrder.length; step += 1) {
      var gridValue = gridOrder[(gridStartIndex - step + gridOrder.length) % gridOrder.length];
      var records = groupedByGrid.get(gridValue) || [];
      if (records.length) {
        return records[records.length - 1];
      }
    }
    return sorted[sorted.length - 1];
  }
  var samePrefixRecords = currentGridRecords.filter(function (record) {
    return normalizeSpecimenPrefix(parseSpecimenNo(record.specimenNo, record.specimenPrefix, record.specimenSerial).prefix) === currentPrefix;
  }).sort(compareRecordsBySpecimenNo);
  if (samePrefixRecords.length) {
    if (currentSerial) {
      var exactIndex = samePrefixRecords.findIndex(function (record) {
        var specimen = parseSpecimenNo(record.specimenNo, record.specimenPrefix, record.specimenSerial);
        return compactNoSpaceValue(specimen.serial) === currentSerial;
      });
      if (exactIndex > 0) {
        return samePrefixRecords[exactIndex - 1];
      }
      if (exactIndex < 0) {
        for (var i = samePrefixRecords.length - 1; i >= 0; i -= 1) {
          var specimen = parseSpecimenNo(samePrefixRecords[i].specimenNo, samePrefixRecords[i].specimenPrefix, samePrefixRecords[i].specimenSerial);
          if (compareSpecimenSerialOnly(compactNoSpaceValue(specimen.serial), currentSerial) < 0) {
            return samePrefixRecords[i];
          }
        }
      }
    } else {
      return samePrefixRecords[samePrefixRecords.length - 1];
    }
  }
  var sortedPrefixes = Array.from(new Set(currentGridRecords.map(function (record) {
    return normalizeSpecimenPrefix(parseSpecimenNo(record.specimenNo, record.specimenPrefix, record.specimenSerial).prefix);
  }))).sort(function (a, b) {
    return a.localeCompare(b, "ja", {
      sensitivity: "base"
    });
  });
  var prevPrefixes = sortedPrefixes.filter(function (prefix) {
    return prefix.localeCompare(currentPrefix, "ja", {
      sensitivity: "base"
    }) < 0;
  }).reverse();
  var _iterator2 = _createForOfIteratorHelper(prevPrefixes),
    _step2;
  try {
    var _loop = function _loop() {
        var prefix = _step2.value;
        var prefixRecords = currentGridRecords.filter(function (record) {
          var specimen = parseSpecimenNo(record.specimenNo, record.specimenPrefix, record.specimenSerial);
          return normalizeSpecimenPrefix(specimen.prefix) === prefix;
        }).sort(compareRecordsBySpecimenNo);
        if (prefixRecords.length) {
          return {
            v: prefixRecords[prefixRecords.length - 1]
          };
        }
      },
      _ret;
    for (_iterator2.s(); !(_step2 = _iterator2.n()).done;) {
      _ret = _loop();
      if (_ret) return _ret.v;
    }
  } catch (err) {
    _iterator2.e(err);
  } finally {
    _iterator2.f();
  }
  for (var _step3 = 1; _step3 <= gridOrder.length; _step3 += 1) {
    var _gridValue = gridOrder[(gridStartIndex - _step3 + gridOrder.length) % gridOrder.length];
    var _records = groupedByGrid.get(_gridValue) || [];
    if (_records.length) {
      return _records[_records.length - 1];
    }
  }
  return sorted[sorted.length - 1];
}
function findNextRecordByGridPrefixThenSerial(currentKuwakuRaw, currentPrefixRaw, currentSerialRaw) {
  var currentKuwakuValue = kuwakuValueForSelect(currentKuwakuRaw);
  var currentPrefix = normalizeSpecimenPrefix(currentPrefixRaw);
  var currentSerial = compactNoSpaceValue(currentSerialRaw);
  var sorted = _toConsumableArray(state.records).sort(compareRecordsByKuwakuThenSpecimen);
  if (!sorted.length) {
    return null;
  }
  var groupedByGrid = new Map();
  var gridOrder = [];
  sorted.forEach(function (record) {
    var gridValue = kuwakuValueForSelect(getRecordKuwaku(record));
    if (!groupedByGrid.has(gridValue)) {
      groupedByGrid.set(gridValue, []);
      gridOrder.push(gridValue);
    }
    groupedByGrid.get(gridValue).push(record);
  });
  var currentGridRecords = groupedByGrid.get(currentKuwakuValue) || [];
  var gridStartIndex = resolveNextGridStartIndex(gridOrder, currentKuwakuValue);
  if (!currentGridRecords.length) {
    for (var step = 0; step < gridOrder.length; step += 1) {
      var gridValue = gridOrder[(gridStartIndex + step) % gridOrder.length];
      var records = groupedByGrid.get(gridValue) || [];
      if (records.length) {
        return records[0];
      }
    }
    return sorted[0];
  }
  var samePrefixRecords = currentGridRecords.filter(function (record) {
    return normalizeSpecimenPrefix(parseSpecimenNo(record.specimenNo, record.specimenPrefix, record.specimenSerial).prefix) === currentPrefix;
  }).sort(compareRecordsBySpecimenNo);
  if (samePrefixRecords.length) {
    if (currentSerial) {
      var exactIndex = samePrefixRecords.findIndex(function (record) {
        var specimen = parseSpecimenNo(record.specimenNo, record.specimenPrefix, record.specimenSerial);
        return compactNoSpaceValue(specimen.serial) === currentSerial;
      });
      if (exactIndex >= 0 && exactIndex + 1 < samePrefixRecords.length) {
        return samePrefixRecords[exactIndex + 1];
      }
      if (exactIndex < 0) {
        var nextBySerial = samePrefixRecords.find(function (record) {
          var specimen = parseSpecimenNo(record.specimenNo, record.specimenPrefix, record.specimenSerial);
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
  var sortedPrefixes = Array.from(new Set(currentGridRecords.map(function (record) {
    return normalizeSpecimenPrefix(parseSpecimenNo(record.specimenNo, record.specimenPrefix, record.specimenSerial).prefix);
  }))).sort(function (a, b) {
    return a.localeCompare(b, "ja", {
      sensitivity: "base"
    });
  });
  var nextPrefixes = sortedPrefixes.filter(function (prefix) {
    return prefix.localeCompare(currentPrefix, "ja", {
      sensitivity: "base"
    }) > 0;
  });
  var _iterator3 = _createForOfIteratorHelper(nextPrefixes),
    _step4;
  try {
    var _loop2 = function _loop2() {
        var prefix = _step4.value;
        var prefixRecords = currentGridRecords.filter(function (record) {
          var specimen = parseSpecimenNo(record.specimenNo, record.specimenPrefix, record.specimenSerial);
          return normalizeSpecimenPrefix(specimen.prefix) === prefix;
        }).sort(compareRecordsBySpecimenNo);
        if (prefixRecords.length) {
          return {
            v: prefixRecords[0]
          };
        }
      },
      _ret2;
    for (_iterator3.s(); !(_step4 = _iterator3.n()).done;) {
      _ret2 = _loop2();
      if (_ret2) return _ret2.v;
    }
  } catch (err) {
    _iterator3.e(err);
  } finally {
    _iterator3.f();
  }
  for (var _step5 = 1; _step5 <= gridOrder.length; _step5 += 1) {
    var _gridValue2 = gridOrder[(gridStartIndex + _step5) % gridOrder.length];
    var _records2 = groupedByGrid.get(_gridValue2) || [];
    if (_records2.length) {
      return _records2[0];
    }
  }
  return sorted[0];
}
function resolveNextGridStartIndex(gridOrder, currentGridValue) {
  var exactIndex = gridOrder.indexOf(currentGridValue);
  if (exactIndex >= 0) {
    return exactIndex;
  }
  var currentLabel = kuwakuLabelForSelect(currentGridValue);
  var insertedIndex = gridOrder.findIndex(function (value) {
    return kuwakuLabelForSelect(value).localeCompare(currentLabel, "ja", {
      numeric: true,
      sensitivity: "base"
    }) > 0;
  });
  return insertedIndex >= 0 ? insertedIndex : 0;
}
function resolvePreviousGridStartIndex(gridOrder, currentGridValue) {
  var exactIndex = gridOrder.indexOf(currentGridValue);
  if (exactIndex >= 0) {
    return exactIndex;
  }
  var currentLabel = kuwakuLabelForSelect(currentGridValue);
  var insertedIndex = gridOrder.findIndex(function (value) {
    return kuwakuLabelForSelect(value).localeCompare(currentLabel, "ja", {
      numeric: true,
      sensitivity: "base"
    }) > 0;
  });
  if (insertedIndex >= 0) {
    return (insertedIndex - 1 + gridOrder.length) % gridOrder.length;
  }
  return Math.max(0, gridOrder.length - 1);
}
function compareSpecimenSerialOnly(aSerialRaw, bSerialRaw) {
  var aSerial = compactNoSpaceValue(aSerialRaw);
  var bSerial = compactNoSpaceValue(bSerialRaw);
  var aIsNumber = /^\d+$/.test(aSerial);
  var bIsNumber = /^\d+$/.test(bSerial);
  if (aIsNumber && bIsNumber) {
    return Number(aSerial) - Number(bSerial);
  }
  return aSerial.localeCompare(bSerial, "ja", {
    numeric: true,
    sensitivity: "base"
  });
}
function applyNextNavigationTarget() {
  var _ref5 = arguments.length > 0 && arguments[0] !== undefined ? arguments[0] : {},
    _ref5$kuwaku = _ref5.kuwaku,
    kuwaku = _ref5$kuwaku === void 0 ? "" : _ref5$kuwaku,
    _ref5$prefix = _ref5.prefix,
    prefix = _ref5$prefix === void 0 ? DEFAULT_SPECIMEN_PREFIX : _ref5$prefix,
    _ref5$serial = _ref5.serial,
    serial = _ref5$serial === void 0 ? "" : _ref5$serial;
  var activeTabId = getActiveTabId();
  var kuwakuParts = parseKuwaku(kuwaku);
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
    var _siteForm$elements5, _siteForm$elements6, _siteForm$elements7, _siteForm$elements8;
    if (siteForm !== null && siteForm !== void 0 && (_siteForm$elements5 = siteForm.elements) !== null && _siteForm$elements5 !== void 0 && _siteForm$elements5.kuwakuHeadA) {
      siteForm.elements.kuwakuHeadA.value = kuwakuParts.headA || DEFAULT_KUWAKU_HEAD_A;
    }
    if (siteForm !== null && siteForm !== void 0 && (_siteForm$elements6 = siteForm.elements) !== null && _siteForm$elements6 !== void 0 && _siteForm$elements6.kuwakuHeadB) {
      siteForm.elements.kuwakuHeadB.value = kuwakuParts.headB || DEFAULT_KUWAKU_HEAD_B;
    }
    if (siteForm !== null && siteForm !== void 0 && (_siteForm$elements7 = siteForm.elements) !== null && _siteForm$elements7 !== void 0 && _siteForm$elements7.kuwakuBlock) {
      siteForm.elements.kuwakuBlock.value = kuwakuParts.block || "";
    }
    if (siteForm !== null && siteForm !== void 0 && (_siteForm$elements8 = siteForm.elements) !== null && _siteForm$elements8 !== void 0 && _siteForm$elements8.kuwakuNo) {
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
function loadRecordIntoInputForNavigation(record) {
  var preferredKuwaku = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : "";
  if (!record) {
    return;
  }
  var kuwakuSource = value(preferredKuwaku) || value(record.kuwaku) || getRecordKuwaku(record);
  var kuwakuParts = parseKuwaku(kuwakuSource);
  var teamState = normalizeTeamState(value(record.team), value(record.teamOther));
  if (siteForm !== null && siteForm !== void 0 && siteForm.elements) {
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
  populateRecordForm(_objectSpread(_objectSpread({}, record), {}, {
    id: ""
  }));
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
  var _siteForm$elements$ku, _siteForm$elements$ku2, _siteForm$elements$ku3, _siteForm$elements$ku4, _siteForm$elements$te2, _siteForm$elements$te3, _siteForm$elements$le, _siteForm$elements$da2, _siteForm$elements$te4, _siteForm$elements$re;
  if (!(siteForm !== null && siteForm !== void 0 && siteForm.elements)) {
    return;
  }
  var kuwakuHeadA = normalizeKuwakuHeadA((_siteForm$elements$ku = siteForm.elements.kuwakuHeadA) === null || _siteForm$elements$ku === void 0 ? void 0 : _siteForm$elements$ku.value);
  var kuwakuHeadB = normalizeKuwakuHeadB((_siteForm$elements$ku2 = siteForm.elements.kuwakuHeadB) === null || _siteForm$elements$ku2 === void 0 ? void 0 : _siteForm$elements$ku2.value);
  var kuwakuBlock = normalizeKuwakuBlock((_siteForm$elements$ku3 = siteForm.elements.kuwakuBlock) === null || _siteForm$elements$ku3 === void 0 ? void 0 : _siteForm$elements$ku3.value);
  var kuwakuNo = normalizeKuwakuNo((_siteForm$elements$ku4 = siteForm.elements.kuwakuNo) === null || _siteForm$elements$ku4 === void 0 ? void 0 : _siteForm$elements$ku4.value);
  var teamState = normalizeTeamState(value((_siteForm$elements$te2 = siteForm.elements.team) === null || _siteForm$elements$te2 === void 0 ? void 0 : _siteForm$elements$te2.value), value((_siteForm$elements$te3 = siteForm.elements.teamOther) === null || _siteForm$elements$te3 === void 0 ? void 0 : _siteForm$elements$te3.value));
  state.site = _objectSpread(_objectSpread({}, state.site), {}, {
    kuwaku: buildKuwaku(kuwakuHeadA, kuwakuHeadB, kuwakuBlock, kuwakuNo),
    kuwakuHeadA: kuwakuHeadA,
    kuwakuHeadB: kuwakuHeadB,
    kuwakuBlock: kuwakuBlock,
    kuwakuNo: kuwakuNo,
    levelHeight: value((_siteForm$elements$le = siteForm.elements.levelHeight) === null || _siteForm$elements$le === void 0 ? void 0 : _siteForm$elements$le.value),
    date: value((_siteForm$elements$da2 = siteForm.elements.date) === null || _siteForm$elements$da2 === void 0 ? void 0 : _siteForm$elements$da2.value),
    team: teamState.team,
    teamOther: teamState.teamOther,
    teamLead: value((_siteForm$elements$te4 = siteForm.elements.teamLead) === null || _siteForm$elements$te4 === void 0 ? void 0 : _siteForm$elements$te4.value),
    recorder: value((_siteForm$elements$re = siteForm.elements.recorder) === null || _siteForm$elements$re === void 0 ? void 0 : _siteForm$elements$re.value)
  });
}
function syncAnalysisTypeInput(prefixRaw) {
  if (!analysisTypeRow || !analysisTypeSelect) {
    return;
  }
  var prefix = normalizeSpecimenPrefix(prefixRaw);
  var isAnalysis = prefix === "a";
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
  setDirectionGroupValue("ns", nsDirInput === null || nsDirInput === void 0 ? void 0 : nsDirInput.value);
  setDirectionGroupValue("ew", ewDirInput === null || ewDirInput === void 0 ? void 0 : ewDirInput.value);
  setDirectionGroupValue("line1Ns", line1NsDirInput === null || line1NsDirInput === void 0 ? void 0 : line1NsDirInput.value);
  setDirectionGroupValue("line1Ew", line1EwDirInput === null || line1EwDirInput === void 0 ? void 0 : line1EwDirInput.value);
  setDirectionGroupValue("line2Ns", line2NsDirInput === null || line2NsDirInput === void 0 ? void 0 : line2NsDirInput.value);
  setDirectionGroupValue("line2Ew", line2EwDirInput === null || line2EwDirInput === void 0 ? void 0 : line2EwDirInput.value);
  setDirectionGroupValue("importantFlag", importantFlagInput === null || importantFlagInput === void 0 ? void 0 : importantFlagInput.value);
  setDirectionGroupValue("simpleRecordFlag", simpleRecordFlagInput === null || simpleRecordFlagInput === void 0 ? void 0 : simpleRecordFlagInput.value);
  setDirectionGroupValue("occurrenceSection", occurrenceSectionInput === null || occurrenceSectionInput === void 0 ? void 0 : occurrenceSectionInput.value);
  setDirectionGroupValue("occurrenceSketch", occurrenceSketchInput === null || occurrenceSketchInput === void 0 ? void 0 : occurrenceSketchInput.value);
  setDirectionGroupValue("layerRelative", layerRelativeInput === null || layerRelativeInput === void 0 ? void 0 : layerRelativeInput.value);
  setDirectionGroupValue("planSizeMode", planSizeModeInput === null || planSizeModeInput === void 0 ? void 0 : planSizeModeInput.value);
  setDirectionGroupValue("largeShapeType", largeShapeTypeInput === null || largeShapeTypeInput === void 0 ? void 0 : largeShapeTypeInput.value);
  setDirectionGroupValue("plungeDir8", largeAxisPlungeDirInput === null || largeAxisPlungeDirInput === void 0 ? void 0 : largeAxisPlungeDirInput.value);
  setDirectionGroupValue("planeDipDir8", planeDipDirInput === null || planeDipDirInput === void 0 ? void 0 : planeDipDirInput.value);
  syncLargeShapeSectionFromForm();
  document.querySelectorAll(".dir-tab").forEach(function (button) {
    var group = value(button.dataset.group);
    var selected = getDirectionGroupValue(group);
    button.classList.toggle("active", normalizeDirectionValue(group, button.dataset.value) === selected);
  });
}
function setDirectionGroupValue(group, valueRaw) {
  var normalized = normalizeDirectionValue(group, valueRaw);
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
    return normalizeDirectionValue(group, nsDirInput === null || nsDirInput === void 0 ? void 0 : nsDirInput.value);
  }
  if (group === "ew") {
    return normalizeDirectionValue(group, ewDirInput === null || ewDirInput === void 0 ? void 0 : ewDirInput.value);
  }
  if (group === "line1Ns") {
    return normalizeDirectionValue(group, line1NsDirInput === null || line1NsDirInput === void 0 ? void 0 : line1NsDirInput.value);
  }
  if (group === "line1Ew") {
    return normalizeDirectionValue(group, line1EwDirInput === null || line1EwDirInput === void 0 ? void 0 : line1EwDirInput.value);
  }
  if (group === "line2Ns") {
    return normalizeDirectionValue(group, line2NsDirInput === null || line2NsDirInput === void 0 ? void 0 : line2NsDirInput.value);
  }
  if (group === "line2Ew") {
    return normalizeDirectionValue(group, line2EwDirInput === null || line2EwDirInput === void 0 ? void 0 : line2EwDirInput.value);
  }
  if (group === "importantFlag") {
    return normalizeDirectionValue(group, importantFlagInput === null || importantFlagInput === void 0 ? void 0 : importantFlagInput.value);
  }
  if (group === "simpleRecordFlag") {
    return normalizeDirectionValue(group, simpleRecordFlagInput === null || simpleRecordFlagInput === void 0 ? void 0 : simpleRecordFlagInput.value);
  }
  if (group === "occurrenceSection") {
    return normalizeDirectionValue(group, occurrenceSectionInput === null || occurrenceSectionInput === void 0 ? void 0 : occurrenceSectionInput.value);
  }
  if (group === "occurrenceSketch") {
    return normalizeDirectionValue(group, occurrenceSketchInput === null || occurrenceSketchInput === void 0 ? void 0 : occurrenceSketchInput.value);
  }
  if (group === "layerRelative") {
    return normalizeDirectionValue(group, layerRelativeInput === null || layerRelativeInput === void 0 ? void 0 : layerRelativeInput.value);
  }
  if (group === "planSizeMode") {
    return normalizeDirectionValue(group, planSizeModeInput === null || planSizeModeInput === void 0 ? void 0 : planSizeModeInput.value);
  }
  if (group === "largeShapeType") {
    return normalizeDirectionValue(group, largeShapeTypeInput === null || largeShapeTypeInput === void 0 ? void 0 : largeShapeTypeInput.value);
  }
  if (group === "plungeDir8") {
    return normalizeDirectionValue(group, largeAxisPlungeDirInput === null || largeAxisPlungeDirInput === void 0 ? void 0 : largeAxisPlungeDirInput.value);
  }
  if (group === "planeDipDir8") {
    return normalizeDirectionValue(group, planeDipDirInput === null || planeDipDirInput === void 0 ? void 0 : planeDipDirInput.value);
  }
  return "";
}
function deriveShapeLabelFromFileName(fileNameRaw) {
  var fileName = value(fileNameRaw);
  if (!fileName) {
    return "";
  }
  var baseName = fileName.split("/").pop() || fileName;
  var normalized = typeof baseName.normalize === "function" ? baseName.normalize("NFC") : baseName;
  var stem = normalized.replace(/\.[^.]+$/, "");
  var mapped = LARGE_SHAPE_FILE_LABEL_MAP[stem] || LARGE_SHAPE_FILE_LABEL_MAP[stem.toLowerCase()] || stem;
  return normalizeLargeShapeLabel(mapped);
}
function normalizeLargeShapeLabel(labelRaw) {
  var raw = value(labelRaw);
  var normalized = typeof raw.normalize === "function" ? raw.normalize("NFC") : raw;
  var withoutExt = normalized.replace(/\.[^.]+$/, "");
  var compact = withoutExt.replace(/\s+/g, "");
  var compactLower = compact.toLowerCase();
  if (LARGE_SHAPE_FILE_LABEL_MAP[compact]) {
    return LARGE_SHAPE_FILE_LABEL_MAP[compact];
  }
  if (LARGE_SHAPE_FILE_LABEL_MAP[compactLower]) {
    return LARGE_SHAPE_FILE_LABEL_MAP[compactLower];
  }
  if (compact === "くびれた骨") {
    return "くびれた形";
  }
  if (compact === "肋骨" || compact === "肋骨（湾曲型）" || compact === "肋骨（湾曲形）" || compact === "肋骨(湾曲型)" || compact === "肋骨(湾曲形)") {
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
  return isCustomLargeShapeType(shapeTypeRaw) ? CUSTOM_IMAGE_SHAPE_CANVAS_DILATE_ITERATIONS : IMAGE_SHAPE_CANVAS_DILATE_ITERATIONS;
}
function normalizeCustomLargeImageName(nameRaw) {
  return value(nameRaw);
}
function normalizeCustomLargeImageDataUrl(dataUrlRaw) {
  var dataUrl = value(dataUrlRaw);
  return dataUrl.startsWith("data:image/") ? dataUrl : "";
}
function normalizeCustomLargeImageAspect(aspectRaw) {
  var text = value(aspectRaw).replace(",", ".");
  if (!text) {
    return "";
  }
  var matched = text.match(/\d+(?:\.\d+)?/);
  if (!matched) {
    return "";
  }
  var ratio = Number(matched[0]);
  if (!Number.isFinite(ratio) || ratio <= 0) {
    return "";
  }
  return String(Number(ratio.toFixed(8)));
}
function formatLengthInputValue(lengthRaw) {
  var length = Number(lengthRaw);
  if (!Number.isFinite(length)) {
    return "";
  }
  if (Math.abs(length) < 0.0005) {
    return "0";
  }
  return String(Number(length.toFixed(3)));
}
function normalizePlanMultiPointDir(valueRaw) {
  var axis = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : "ns";
  return axis === "ew" ? normalizeEwDir(valueRaw) : normalizeNsDir(valueRaw);
}
function normalizePlanMultiPointDistance(valueRaw) {
  var distance = parseDistanceToCm(valueRaw);
  if (distance == null || distance < 0) {
    return "";
  }
  return formatLengthInputValue(distance);
}
function createDefaultPlanMultiPoint() {
  return {
    nsDir: "北から",
    nsCm: "",
    ewDir: "東から",
    ewCm: ""
  };
}
function normalizePlanMultiPointEntry(entryRaw) {
  var entry = entryRaw && _typeof(entryRaw) === "object" ? entryRaw : {};
  var nsDir = normalizePlanMultiPointDir(value(entry.nsDir), "ns");
  var ewDir = normalizePlanMultiPointDir(value(entry.ewDir), "ew");
  var nsCm = normalizePlanMultiPointDistance(value(entry.nsCm));
  var ewCm = normalizePlanMultiPointDistance(value(entry.ewCm));
  if (!nsCm || !ewCm) {
    return null;
  }
  return {
    nsDir: nsDir,
    nsCm: nsCm,
    ewDir: ewDir,
    ewCm: ewCm
  };
}
function normalizePlanMultiPoints(pointsRaw) {
  var points = Array.isArray(pointsRaw) ? pointsRaw : [];
  var normalized = [];
  var seen = new Set();
  points.forEach(function (point) {
    var entry = normalizePlanMultiPointEntry(point);
    if (!entry) {
      return;
    }
    var key = "".concat(entry.nsDir, "|").concat(entry.nsCm, "|").concat(entry.ewDir, "|").concat(entry.ewCm);
    if (seen.has(key)) {
      return;
    }
    seen.add(key);
    normalized.push(entry);
  });
  return normalized;
}
function createMultiPointRowElement() {
  var pointRaw = arguments.length > 0 && arguments[0] !== undefined ? arguments[0] : {};
  var normalized = normalizePlanMultiPointEntry(pointRaw) || _objectSpread(_objectSpread({}, createDefaultPlanMultiPoint()), {}, {
    nsDir: normalizePlanMultiPointDir(value(pointRaw === null || pointRaw === void 0 ? void 0 : pointRaw.nsDir), "ns"),
    ewDir: normalizePlanMultiPointDir(value(pointRaw === null || pointRaw === void 0 ? void 0 : pointRaw.ewDir), "ew")
  });
  var row = document.createElement("div");
  row.className = "multi-point-row";
  row.dataset.multiPointRow = "1";
  var nsLabel = document.createElement("label");
  nsLabel.textContent = "平面位置[北から・南から]";
  var nsWrap = document.createElement("div");
  nsWrap.className = "multi-point-input";
  var nsDirSelect = document.createElement("select");
  nsDirSelect.dataset.multiPointNsDir = "1";
  ["北から", "南から"].forEach(function (dir) {
    var option = document.createElement("option");
    option.value = dir;
    option.textContent = dir;
    nsDirSelect.append(option);
  });
  nsDirSelect.value = normalized.nsDir;
  var nsCmInput = document.createElement("input");
  nsCmInput.type = "text";
  nsCmInput.placeholder = "cm";
  nsCmInput.inputMode = "decimal";
  nsCmInput.dataset.multiPointNsCm = "1";
  nsCmInput.value = value(normalized.nsCm);
  var nsUnit = document.createElement("span");
  nsUnit.textContent = "cm";
  nsWrap.append(nsDirSelect, nsCmInput, nsUnit);
  nsLabel.append(nsWrap);
  var ewLabel = document.createElement("label");
  ewLabel.textContent = "平面位置[東から・西から]";
  var ewWrap = document.createElement("div");
  ewWrap.className = "multi-point-input";
  var ewDirSelect = document.createElement("select");
  ewDirSelect.dataset.multiPointEwDir = "1";
  ["東から", "西から"].forEach(function (dir) {
    var option = document.createElement("option");
    option.value = dir;
    option.textContent = dir;
    ewDirSelect.append(option);
  });
  ewDirSelect.value = normalized.ewDir;
  var ewCmInput = document.createElement("input");
  ewCmInput.type = "text";
  ewCmInput.placeholder = "cm";
  ewCmInput.inputMode = "decimal";
  ewCmInput.dataset.multiPointEwCm = "1";
  ewCmInput.value = value(normalized.ewCm);
  var ewUnit = document.createElement("span");
  ewUnit.textContent = "cm";
  ewWrap.append(ewDirSelect, ewCmInput, ewUnit);
  ewLabel.append(ewWrap);
  var removeButton = document.createElement("button");
  removeButton.type = "button";
  removeButton.className = "multi-point-remove";
  removeButton.dataset.multiPointRemove = "1";
  removeButton.textContent = "削除";
  row.append(nsLabel, ewLabel, removeButton);
  return row;
}
function renderMultiPointRows() {
  var pointsRaw = arguments.length > 0 && arguments[0] !== undefined ? arguments[0] : [];
  if (!multiPointRows) {
    return;
  }
  var points = normalizePlanMultiPoints(pointsRaw);
  var source = points.length ? points : [createDefaultPlanMultiPoint()];
  multiPointRows.innerHTML = "";
  source.forEach(function (point) {
    multiPointRows.append(createMultiPointRowElement(point));
  });
  syncMultiPointRemoveButtonState();
}
function syncMultiPointRemoveButtonState() {
  if (!multiPointRows) {
    return;
  }
  var rows = _toConsumableArray(multiPointRows.querySelectorAll("[data-multi-point-row]"));
  var canRemove = rows.length > 1;
  rows.forEach(function (row) {
    var removeButton = row.querySelector("[data-multi-point-remove]");
    if (removeButton instanceof HTMLButtonElement) {
      removeButton.disabled = !canRemove;
    }
  });
}
function readMultiPointRowsFromForm() {
  if (!multiPointRows) {
    return [];
  }
  var rows = _toConsumableArray(multiPointRows.querySelectorAll("[data-multi-point-row]"));
  var points = rows.map(function (row) {
    var _row$querySelector, _row$querySelector2, _row$querySelector3, _row$querySelector4;
    return {
      nsDir: value((_row$querySelector = row.querySelector("[data-multi-point-ns-dir]")) === null || _row$querySelector === void 0 ? void 0 : _row$querySelector.value),
      nsCm: value((_row$querySelector2 = row.querySelector("[data-multi-point-ns-cm]")) === null || _row$querySelector2 === void 0 ? void 0 : _row$querySelector2.value),
      ewDir: value((_row$querySelector3 = row.querySelector("[data-multi-point-ew-dir]")) === null || _row$querySelector3 === void 0 ? void 0 : _row$querySelector3.value),
      ewCm: value((_row$querySelector4 = row.querySelector("[data-multi-point-ew-cm]")) === null || _row$querySelector4 === void 0 ? void 0 : _row$querySelector4.value)
    };
  });
  return normalizePlanMultiPoints(points);
}
function collectPlanMultiPointCoords(record) {
  var points = normalizePlanMultiPoints(record === null || record === void 0 ? void 0 : record.multiPoints);
  if (!points.length) {
    return [];
  }
  var coords = [];
  var seen = new Set();
  points.forEach(function (point) {
    var coord = convertPositionToPlanCoords(point.nsDir, point.nsCm, point.ewDir, point.ewCm);
    if (!coord) {
      return;
    }
    var key = "".concat(coord.x.toFixed(4), "|").concat(coord.y.toFixed(4));
    if (seen.has(key)) {
      return;
    }
    seen.add(key);
    coords.push(coord);
  });
  return coords;
}
function readImageAspectRatio(dataUrlRaw) {
  var dataUrl = normalizeCustomLargeImageDataUrl(dataUrlRaw);
  if (!dataUrl) {
    return Promise.resolve(null);
  }
  return new Promise(function (resolve) {
    var image = new Image();
    image.onload = function () {
      var width = Math.max(1, Number(image.naturalWidth || image.width) || 0);
      var height = Math.max(1, Number(image.naturalHeight || image.height) || 0);
      if (!width || !height) {
        resolve(null);
        return;
      }
      resolve(width / height);
    };
    image.onerror = function () {
      return resolve(null);
    };
    image.src = dataUrl;
  });
}
function updateCustomLargeImageAspectFromDataUrl() {
  return _updateCustomLargeImageAspectFromDataUrl.apply(this, arguments);
}
function _updateCustomLargeImageAspectFromDataUrl() {
  _updateCustomLargeImageAspectFromDataUrl = _asyncToGenerator(_regenerator().m(function _callee8() {
    var dataUrlRaw,
      dataUrl,
      ratio,
      _args8 = arguments;
    return _regenerator().w(function (_context8) {
      while (1) switch (_context8.n) {
        case 0:
          dataUrlRaw = _args8.length > 0 && _args8[0] !== undefined ? _args8[0] : customLargeImageDataUrlInput === null || customLargeImageDataUrlInput === void 0 ? void 0 : customLargeImageDataUrlInput.value;
          dataUrl = normalizeCustomLargeImageDataUrl(dataUrlRaw);
          if (customLargeImageAspectInput) {
            _context8.n = 1;
            break;
          }
          return _context8.a(2, null);
        case 1:
          if (dataUrl) {
            _context8.n = 2;
            break;
          }
          customLargeImageAspectInput.value = "";
          return _context8.a(2, null);
        case 2:
          _context8.n = 3;
          return readImageAspectRatio(dataUrl);
        case 3:
          ratio = _context8.v;
          customLargeImageAspectInput.value = normalizeCustomLargeImageAspect(ratio);
          return _context8.a(2, ratio && ratio > 0 ? ratio : null);
      }
    }, _callee8);
  }));
  return _updateCustomLargeImageAspectFromDataUrl.apply(this, arguments);
}
function deriveCustomLargeImageNameFromFileName(fileNameRaw) {
  var fileName = value(fileNameRaw);
  if (!fileName) {
    return "";
  }
  var base = fileName.split("/").pop() || fileName;
  var normalized = typeof base.normalize === "function" ? base.normalize("NFC") : base;
  return value(normalized.replace(/\.[^.]+$/, ""));
}
function syncCustomLargeImageStatus() {
  if (!customLargeImageStatus) {
    return;
  }
  var imageName = normalizeCustomLargeImageName(customLargeImageNameInput === null || customLargeImageNameInput === void 0 ? void 0 : customLargeImageNameInput.value);
  var hasImage = Boolean(normalizeCustomLargeImageDataUrl(customLargeImageDataUrlInput === null || customLargeImageDataUrlInput === void 0 ? void 0 : customLargeImageDataUrlInput.value));
  if (hasImage) {
    customLargeImageStatus.textContent = imageName ? "\u753B\u50CF\u8A2D\u5B9A\u6E08\u307F: ".concat(imageName) : "画像設定済み";
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
function setCustomLargeImageFromFile(_x7) {
  return _setCustomLargeImageFromFile.apply(this, arguments);
}
function _setCustomLargeImageFromFile() {
  _setCustomLargeImageFromFile = _asyncToGenerator(_regenerator().m(function _callee9(file) {
    var sourceFile, dataUrl, _t5;
    return _regenerator().w(function (_context9) {
      while (1) switch (_context9.p = _context9.n) {
        case 0:
          sourceFile = file instanceof File ? file : null;
          if (sourceFile) {
            _context9.n = 1;
            break;
          }
          return _context9.a(2);
        case 1:
          _context9.p = 1;
          _context9.n = 2;
          return loadImageFileDataUrlWithFallback(sourceFile, {
            maxLength: 1200,
            quality: 0.85,
            mimeType: ""
          });
        case 2:
          dataUrl = _context9.v;
          if (customLargeImageDataUrlInput) {
            customLargeImageDataUrlInput.value = normalizeCustomLargeImageDataUrl(dataUrl);
          }
          _context9.n = 3;
          return updateCustomLargeImageAspectFromDataUrl(dataUrl);
        case 3:
          if (customLargeImageNameInput && !value(customLargeImageNameInput.value)) {
            customLargeImageNameInput.value = deriveCustomLargeImageNameFromFileName(sourceFile.name);
          }
          void syncImageFrameSizeFromAspectLock("customLargeImageDataUrl").then(function () {
            syncLargeShapeImagePreviewTransform();
          });
          syncCustomLargeImageStatus();
          syncLargeShapeImagePreviewForCurrentForm();
          _context9.n = 5;
          break;
        case 4:
          _context9.p = 4;
          _t5 = _context9.v;
          showToast("画像アップロードに失敗しました");
        case 5:
          return _context9.a(2);
      }
    }, _callee9, null, [[1, 4]]);
  }));
  return _setCustomLargeImageFromFile.apply(this, arguments);
}
function syncLargeShapeImagePreviewForCurrentForm() {
  var shapeType = normalizeLargeShapeType(largeShapeTypeInput === null || largeShapeTypeInput === void 0 ? void 0 : largeShapeTypeInput.value);
  if (!isLargeShapeImageType(shapeType)) {
    syncLargeShapeImagePreview("");
    return;
  }
  var customPath = isCustomLargeShapeType(shapeType) ? normalizeCustomLargeImageDataUrl(customLargeImageDataUrlInput === null || customLargeImageDataUrlInput === void 0 ? void 0 : customLargeImageDataUrlInput.value) : "";
  var customName = isCustomLargeShapeType(shapeType) ? normalizeCustomLargeImageName(customLargeImageNameInput === null || customLargeImageNameInput === void 0 ? void 0 : customLargeImageNameInput.value) : "";
  syncLargeShapeImagePreview(shapeType, customPath, customName);
}
function toSafeAssetUrl(pathRaw) {
  var path = pathRaw == null ? "" : String(pathRaw).trim();
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
  var raw = pathRaw == null ? "" : String(pathRaw).trim();
  if (!raw) {
    return "";
  }
  if (raw.startsWith("data:")) {
    return raw;
  }
  var normalized = raw.replace(/[#?].*$/, "");
  try {
    normalized = decodeURI(normalized);
  } catch (_error) {}
  normalized = normalized.replace(/\\/g, "/").replace(/^(\.\/)+/, "").replace(/^\/+/, "");
  var repoMarker = "kaseki_mobile_app/";
  var markerIndex = normalized.lastIndexOf(repoMarker);
  if (markerIndex >= 0) {
    normalized = normalized.slice(markerIndex + repoMarker.length);
  }
  return normalized;
}
function getInlineLargeShapeDataUrl(pathRaw) {
  var key = normalizeAssetPathKey(pathRaw);
  if (!key || key.startsWith("data:")) {
    return "";
  }
  return value(INLINE_LARGE_SHAPE_DATA_MAP[key]);
}
function renderLargeShapeImageButtons() {
  if (!largeShapeImageButtons) {
    return;
  }
  var labels = Array.from(new Set([].concat(_toConsumableArray(Array.from(largeShapeImagePathMap.keys()).filter(function (label) {
    return value(label);
  })), [CUSTOM_LARGE_SHAPE_TYPE])));
  largeShapeImageButtons.innerHTML = labels.map(function (label) {
    var buttonLabel = label === CUSTOM_LARGE_SHAPE_TYPE ? "画像アップロード" : label;
    return "<button class=\"dir-tab\" data-group=\"largeShapeType\" data-value=\"".concat(escapeHtml(label), "\" type=\"button\">").concat(escapeHtml(buttonLabel), "</button>");
  }).join("");
}
function loadLargeShapeImageManifest() {
  return _loadLargeShapeImageManifest.apply(this, arguments);
}
function _loadLargeShapeImageManifest() {
  _loadLargeShapeImageManifest = _asyncToGenerator(_regenerator().m(function _callee0() {
    var response, manifest, images, nextMap, _t6;
    return _regenerator().w(function (_context0) {
      while (1) switch (_context0.p = _context0.n) {
        case 0:
          _context0.p = 0;
          _context0.n = 1;
          return fetch("".concat(LARGE_SHAPE_MANIFEST_PATH, "?v=").concat(Date.now()), {
            cache: "no-store"
          });
        case 1:
          response = _context0.v;
          if (response.ok) {
            _context0.n = 2;
            break;
          }
          return _context0.a(2);
        case 2:
          _context0.n = 3;
          return response.json();
        case 3:
          manifest = _context0.v;
          images = Array.isArray(manifest === null || manifest === void 0 ? void 0 : manifest.images) ? manifest.images : [];
          if (images.length) {
            _context0.n = 4;
            break;
          }
          return _context0.a(2);
        case 4:
          nextMap = new Map(Object.entries(DEFAULT_LARGE_SHAPE_IMAGE_PATHS));
          images.forEach(function (item) {
            if (typeof item === "string") {
              var fileName = value(item);
              if (!fileName) {
                return;
              }
              var label = deriveShapeLabelFromFileName(fileName);
              if (!label || EXCLUDED_LARGE_SHAPE_LABELS.has(label)) {
                return;
              }
              if (!nextMap.has(label)) {
                nextMap.set(label, toSafeAssetUrl("".concat(LARGE_SHAPE_DIR_PATH, "/").concat(fileName)));
              }
              return;
            }
            if (item && _typeof(item) === "object") {
              var _fileName = value(item.file || item.path || item.src);
              var explicitLabel = value(item.label);
              var _label2 = explicitLabel ? normalizeLargeShapeLabel(explicitLabel) : deriveShapeLabelFromFileName(_fileName);
              if (!_fileName || !_label2 || EXCLUDED_LARGE_SHAPE_LABELS.has(_label2)) {
                return;
              }
              var path = /^(\/|\.\/|https?:)/.test(_fileName) ? _fileName : "".concat(LARGE_SHAPE_DIR_PATH, "/").concat(_fileName);
              nextMap.set(_label2, toSafeAssetUrl(path));
            }
          });
          if (nextMap.size) {
            _context0.n = 5;
            break;
          }
          return _context0.a(2);
        case 5:
          largeShapeImagePathMap = nextMap;
          planLargeShapeImageCache.clear();
          planLargeShapeTintedCanvasCache.clear();
          planLargeShapeTintedDataUrlCache.clear();
          renderLargeShapeImageButtons();
          syncDirectionTabsFromForm();
          renderOutputs();
          _context0.n = 7;
          break;
        case 6:
          _context0.p = 6;
          _t6 = _context0.v;
        case 7:
          return _context0.a(2);
      }
    }, _callee0, null, [[0, 6]]);
  }));
  return _loadLargeShapeImageManifest.apply(this, arguments);
}
function syncLargeShapeImagePreview(shapeTypeRaw) {
  var explicitImagePathRaw = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : "";
  var explicitTitleRaw = arguments.length > 2 && arguments[2] !== undefined ? arguments[2] : "";
  if (!largeShapeImagePreview || !largeShapeImagePreviewTitle || !largeShapeImagePreviewImg) {
    return;
  }
  var shapeType = normalizeLargeShapeType(shapeTypeRaw);
  var explicitImagePath = normalizeCustomLargeImageDataUrl(explicitImagePathRaw);
  var candidates = getLargeShapeImagePathCandidates(shapeType, explicitImagePath);
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
  var previewTitle = value(explicitTitleRaw) || shapeType;
  largeShapeImagePreviewTitle.textContent = previewTitle;
  largeShapeImagePreviewImg.alt = "".concat(shapeType, " \u753B\u50CF");
  var candidateIndex = 0;
  largeShapeImagePreviewImg.onerror = function () {
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
  var mode = normalizePlanSizeMode(planSizeModeInput === null || planSizeModeInput === void 0 ? void 0 : planSizeModeInput.value);
  if (planSizeModeInput) {
    planSizeModeInput.value = mode;
  }
  var isLarge = mode === "大きなもの";
  var isMultiPoint = mode === "複数点";
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
    largeShapePanels.forEach(function (panel) {
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
  var shapeTypeRaw = normalizeLargeShapeType(largeShapeTypeInput === null || largeShapeTypeInput === void 0 ? void 0 : largeShapeTypeInput.value);
  var shapeType = shapeTypeRaw || "直線状";
  var isImageShape = isLargeShapeImageType(shapeType);
  var isCustomImageShape = isCustomLargeShapeType(shapeType);
  var isLineShape = shapeType === "直線状";
  var usesAxisDirection = isLineShape || shapeType === "長方形" || shapeType === "楕円";
  if (largeShapeTypeInput) {
    largeShapeTypeInput.value = shapeType;
  }
  if (largeAxisDirectionInput) {
    var fallbackAxis = !isLineShape && planeStrikeInput ? planeStrikeInput.value : "";
    largeAxisDirectionInput.value = usesAxisDirection ? normalizeLargeAxisDirection(largeAxisDirectionInput.value || fallbackAxis) : "";
  }
  if (largeAxisPlungeInput) {
    largeAxisPlungeInput.value = isLineShape ? normalizeLargeAxisPlungeDeg(largeAxisPlungeInput.value) : "";
  }
  if (largeAxisPlungeDirInput) {
    largeAxisPlungeDirInput.value = isLineShape ? normalizeCompass8Direction(largeAxisPlungeDirInput.value) : "";
  }
  if (planeStrikeInput) {
    var fallbackStrike = usesAxisDirection && largeAxisDirectionInput ? largeAxisDirectionInput.value : "";
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
  largeShapePanels.forEach(function (panel) {
    var panelType = value(panel.dataset.largeShapePanel);
    var shouldShow = panelType === shapeType || panelType === "__IMAGE__" && isImageShape;
    panel.classList.toggle("hidden", !shouldShow);
  });
  syncCustomLargeImageStatus();
  if (isImageShape) {
    syncLargeShapeImagePreviewForCurrentForm();
    void syncImageFrameSizeFromAspectLock("largeShapeType").then(function () {
      syncLargeShapeImagePreviewTransform();
    });
  } else {
    syncLargeShapeImagePreview("");
  }
}
function activateLayerTab(layerRaw) {
  var layer = PRESET_LAYER_NAMES.includes(value(layerRaw)) ? value(layerRaw) : DEFAULT_LAYER_NAME;
  layerNameInput.value = layer;
  layerTabButtons.forEach(function (button) {
    button.classList.toggle("active", button.dataset.layer === layer);
  });
  var isOther = layer === OTHER_LAYER_NAME;
  layerOtherInput.classList.toggle("hidden", !isOther);
  if (!isOther) {
    layerOtherInput.value = "";
  }
}
function setLayerFromValue(layerRaw) {
  var layerValue = normalizeLayerName(value(layerRaw));
  if (!layerValue) {
    activateLayerTab(DEFAULT_LAYER_NAME);
    return;
  }
  if (PRESET_LAYER_NAMES.includes(layerValue) && layerValue !== OTHER_LAYER_NAME) {
    activateLayerTab(layerValue);
    return;
  }
  activateLayerTab(OTHER_LAYER_NAME);
  var otherText = extractOtherLayerText(layerValue);
  layerOtherInput.value = otherText;
}
function getSelectedLayerName() {
  var selected = value(layerNameInput.value) || DEFAULT_LAYER_NAME;
  if (selected !== OTHER_LAYER_NAME) {
    return selected;
  }
  var otherText = value(layerOtherInput.value);
  return otherText ? "".concat(OTHER_LAYER_NAME, ":").concat(otherText) : OTHER_LAYER_NAME;
}
function applyCarryForwardFields(saved) {
  setLayerFromValue(value(saved === null || saved === void 0 ? void 0 : saved.layerName));
  recordForm.elements.unit.value = value(saved === null || saved === void 0 ? void 0 : saved.unit);
  recordForm.elements.detail.value = value(saved === null || saved === void 0 ? void 0 : saved.detail);
  recordForm.elements.detailSub.value = value(saved === null || saved === void 0 ? void 0 : saved.detailSub);
  recordForm.elements.layerFacies.value = value(saved === null || saved === void 0 ? void 0 : saved.layerFacies);
  recordForm.elements.layerRef.value = value(saved === null || saved === void 0 ? void 0 : saved.layerRef);
  recordForm.elements.layerFromCm.value = value(saved === null || saved === void 0 ? void 0 : saved.layerFromCm);
  recordForm.elements.layerRelative.value = value(saved === null || saved === void 0 ? void 0 : saved.layerRelative);
  syncDirectionTabsFromForm();
}
function markCarryForwardSavedFields(saved) {
  clearCarryForwardSavedFields();
  if (value(saved === null || saved === void 0 ? void 0 : saved.unit)) {
    recordForm.elements.unit.classList.add("saved-carry-value");
  }
  if (value(saved === null || saved === void 0 ? void 0 : saved.detail)) {
    recordForm.elements.detail.classList.add("saved-carry-value");
  }
  if (value(saved === null || saved === void 0 ? void 0 : saved.detailSub)) {
    recordForm.elements.detailSub.classList.add("saved-carry-value");
  }
  if (value(saved === null || saved === void 0 ? void 0 : saved.layerFacies)) {
    recordForm.elements.layerFacies.classList.add("saved-carry-value");
  }
  if (value(saved === null || saved === void 0 ? void 0 : saved.layerRef)) {
    recordForm.elements.layerRef.classList.add("saved-carry-value");
  }
  if (value(saved === null || saved === void 0 ? void 0 : saved.layerFromCm)) {
    recordForm.elements.layerFromCm.classList.add("saved-carry-value");
  }
  if (value(saved === null || saved === void 0 ? void 0 : saved.layerRelative)) {
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
  recordForm.elements.layerFacies.classList.remove("saved-carry-value");
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
  recordForm.querySelectorAll(".overwrite-updated").forEach(function (element) {
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
  var fields = ["specimenSerial", "nameMemo", "unit", "discoverer", "identifier", "levelUpperCm", "levelLowerCm", "occurrenceSection", "occurrenceSketch", "importantFlag", "simpleRecordFlag", "analysisType", "nsCm", "ewCm", "largeAxisPlungeDeg", "largeAxisPlungeDir8", "planeStrikeDirection", "planeDipDeg", "planeDipDir8", "imgRotateDeg", "imgFrameWidthCm", "imgFrameHeightCm", "imgSkewXDeg", "imgSkewYDeg", "imgFlipH", "imgFlipV", "imgLockAspectRatio", "customLargeImageAspect", "detail", "detailSub", "layerFacies", "layerRef", "layerFromCm", "layerRelative", "notes"];
  fields.forEach(function (name) {
    var element = recordForm.elements[name];
    if (!(element instanceof Element)) {
      return;
    }
    var prev = value(previousRecord === null || previousRecord === void 0 ? void 0 : previousRecord[name]);
    var next = value(nextRecord === null || nextRecord === void 0 ? void 0 : nextRecord[name]);
    if (prev !== next) {
      element.classList.add("overwrite-updated");
    }
  });
  var prevPrefix = normalizeSpecimenPrefix(previousRecord.specimenPrefix);
  var nextPrefix = normalizeSpecimenPrefix(nextRecord.specimenPrefix);
  if (prevPrefix !== nextPrefix) {
    specimenPrefixLabel.classList.add("overwrite-updated");
    specimenTabButtons.forEach(function (button) {
      if (normalizeSpecimenPrefix(button.dataset.prefix) === nextPrefix) {
        button.classList.add("overwrite-updated");
      }
    });
  }
  if (normalizeNsDir(previousRecord.nsDir) !== normalizeNsDir(nextRecord.nsDir)) {
    document.querySelectorAll(".dir-tab").forEach(function (button) {
      if (button.dataset.group === "ns") {
        button.classList.add("overwrite-updated");
      }
    });
  }
  if (normalizeEwDir(previousRecord.ewDir) !== normalizeEwDir(nextRecord.ewDir)) {
    document.querySelectorAll(".dir-tab").forEach(function (button) {
      if (button.dataset.group === "ew") {
        button.classList.add("overwrite-updated");
      }
    });
  }
  if (value(previousRecord.layerName) !== value(nextRecord.layerName)) {
    layerTabButtons.forEach(function (button) {
      if (button.classList.contains("active")) {
        button.classList.add("overwrite-updated");
      }
    });
    layerOtherInput.classList.add("overwrite-updated");
  }
  var previousParts = parseKuwaku(previousKuwakuRaw);
  var nextParts = parseKuwaku(nextKuwakuRaw);
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
  if (editTeamInput && (value(previousRecord.team) !== value(nextRecord.team) || value(previousRecord.teamOther) !== value(nextRecord.teamOther))) {
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
  var activeButton = Array.from(layerTabButtons).find(function (button) {
    return button.classList.contains("active");
  });
  if (activeButton) {
    activeButton.classList.add("saved-carry-value");
  }
}
function clearLayerSavedTabState() {
  layerTabButtons.forEach(function (button) {
    button.classList.remove("saved-carry-value");
  });
}
function syncAltitudeDirectInputUi() {
  var _recordForm$elements, _recordForm$elements2;
  var _ref6 = arguments.length > 0 && arguments[0] !== undefined ? arguments[0] : {},
    _ref6$clearWhenDisabl = _ref6.clearWhenDisabled,
    clearWhenDisabled = _ref6$clearWhenDisabl === void 0 ? false : _ref6$clearWhenDisabl;
  var altitudeToggleField = recordForm === null || recordForm === void 0 || (_recordForm$elements = recordForm.elements) === null || _recordForm$elements === void 0 ? void 0 : _recordForm$elements.altitudeInputEnabled;
  var altitudeInputField = recordForm === null || recordForm === void 0 || (_recordForm$elements2 = recordForm.elements) === null || _recordForm$elements2 === void 0 ? void 0 : _recordForm$elements2.altitudeDirectM;
  if (!(altitudeInputField instanceof HTMLInputElement)) {
    return;
  }
  var isEnabled = altitudeToggleField instanceof HTMLInputElement ? altitudeToggleField.checked : false;
  altitudeInputField.disabled = !isEnabled;
  if (!isEnabled && clearWhenDisabled) {
    altitudeInputField.value = "";
  }
}
function handleRecordFormFieldEdit(event) {
  var target = event.target;
  if (!(target instanceof Element)) {
    return;
  }
  target.classList.remove("overwrite-updated");
  var isCarryField = target instanceof HTMLInputElement && (target.name === "unit" || target.name === "detail" || target.name === "detailSub" || target.name === "layerFacies" || target.name === "layerRef" || target.name === "layerFromCm" || target.name === "layerRelative");
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
    syncAltitudeDirectInputUi({
      clearWhenDisabled: true
    });
  }
  if (target instanceof HTMLElement && target.closest("#large-shape-section")) {
    var targetName = target instanceof HTMLInputElement || target instanceof HTMLSelectElement || target instanceof HTMLTextAreaElement ? target.name : "";
    syncCustomLargeImageStatus();
    if (target === customLargeImageNameInput || target === customLargeImageDataUrlInput || target === customLargeImageFileInput) {
      if (target === customLargeImageDataUrlInput) {
        void updateCustomLargeImageAspectFromDataUrl(customLargeImageDataUrlInput === null || customLargeImageDataUrlInput === void 0 ? void 0 : customLargeImageDataUrlInput.value);
      }
      syncLargeShapeImagePreviewForCurrentForm();
    }
    void syncImageFrameSizeFromAspectLock(targetName).then(function () {
      syncLargeShapeImagePreviewTransform();
      updateEditMissingRequiredHighlights();
    });
  }
  updateEditMissingRequiredHighlights();
}
function setDefaultImageCornerDirections() {
  var defaults = {
    imgP1NsDir: "北から",
    imgP1EwDir: "東から",
    imgP2NsDir: "北から",
    imgP2EwDir: "東から",
    imgP3NsDir: "北から",
    imgP3EwDir: "東から",
    imgP4NsDir: "北から",
    imgP4EwDir: "東から"
  };
  Object.entries(defaults).forEach(function (_ref7) {
    var _recordForm$elements3;
    var _ref8 = _slicedToArray(_ref7, 2),
      name = _ref8[0],
      defaultValue = _ref8[1];
    var field = recordForm === null || recordForm === void 0 || (_recordForm$elements3 = recordForm.elements) === null || _recordForm$elements3 === void 0 ? void 0 : _recordForm$elements3.namedItem(name);
    if (field instanceof HTMLSelectElement || field instanceof HTMLInputElement) {
      field.value = defaultValue;
    }
  });
}
function clearImageCornerCmFields() {
  var names = ["imgP1NsCm", "imgP1EwCm", "imgP2NsCm", "imgP2EwCm", "imgP3NsCm", "imgP3EwCm", "imgP4NsCm", "imgP4EwCm"];
  names.forEach(function (name) {
    var _recordForm$elements4;
    var field = recordForm === null || recordForm === void 0 || (_recordForm$elements4 = recordForm.elements) === null || _recordForm$elements4 === void 0 ? void 0 : _recordForm$elements4.namedItem(name);
    if (field instanceof HTMLInputElement) {
      field.value = "";
    }
  });
}
function extractImageCornerFieldsFromFormData(formData) {
  var getNsDir = function getNsDir(name) {
    return normalizeNsDir(value(formData.get(name)));
  };
  var getEwDir = function getEwDir(name) {
    return normalizeEwDir(value(formData.get(name)));
  };
  var getCm = function getCm(name) {
    return value(formData.get(name));
  };
  var normalized = {
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
    imgP4EwCm: getCm("imgP4EwCm")
  };
  return normalized;
}
function normalizeToggleFlag(valueRaw) {
  var text = value(valueRaw).toLowerCase();
  return text === "1" || text === "true" || text === "on" || text === "yes" ? "1" : "0";
}
function normalizeImageRotationDeg(valueRaw) {
  var text = value(valueRaw).replace(/[°度]/g, "");
  if (!text) {
    return "";
  }
  var matched = text.match(/-?\d+(?:\.\d+)?/);
  if (!matched) {
    return "";
  }
  var num = Number(matched[0]);
  if (!Number.isFinite(num)) {
    return "";
  }
  var normalized = (num % 360 + 360) % 360;
  return Number.isInteger(normalized) ? String(normalized) : String(normalized).replace(/\.?0+$/, "");
}
function normalizeImageSkewDeg(valueRaw) {
  var text = value(valueRaw).replace(/[°度]/g, "");
  if (!text) {
    return "";
  }
  var matched = text.match(/-?\d+(?:\.\d+)?/);
  if (!matched) {
    return "";
  }
  var num = Number(matched[0]);
  if (!Number.isFinite(num)) {
    return "";
  }
  var limited = clamp(num, -80, 80);
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
    customLargeImageAspect: normalizeCustomLargeImageAspect(formData.get("customLargeImageAspect"))
  };
}
function syncLargeShapeImagePreviewTransform() {
  if (!largeShapeImagePreviewImg || !recordForm) {
    return;
  }
  var shapeType = normalizeLargeShapeType(largeShapeTypeInput === null || largeShapeTypeInput === void 0 ? void 0 : largeShapeTypeInput.value);
  var isImageShape = isLargeShapeImageType(shapeType);
  if (!isImageShape) {
    largeShapeImagePreviewImg.style.transform = "";
    largeShapeImagePreviewImg.style.width = "";
    largeShapeImagePreviewImg.style.height = "";
    largeShapeImagePreviewImg.style.maxWidth = "";
    largeShapeImagePreviewImg.style.maxHeight = "";
    largeShapeImagePreviewImg.style.objectFit = "";
    return;
  }
  var formData = new FormData(recordForm);
  var frameWidthCm = parseDistanceToCm(formData.get("imgFrameWidthCm"));
  var frameHeightCm = parseDistanceToCm(formData.get("imgFrameHeightCm"));
  var rotate = Number(normalizeImageRotationDeg(formData.get("imgRotateDeg")) || "0");
  var skewX = Number(normalizeImageSkewDeg(formData.get("imgSkewXDeg")) || "0");
  var skewY = Number(normalizeImageSkewDeg(formData.get("imgSkewYDeg")) || "0");
  var flipH = normalizeToggleFlag(formData.get("imgFlipH")) === "1";
  var flipV = normalizeToggleFlag(formData.get("imgFlipV")) === "1";
  var scaleX = flipH ? -1 : 1;
  var scaleY = flipV ? -1 : 1;
  if (frameWidthCm != null && frameWidthCm > 0 && frameHeightCm != null && frameHeightCm > 0) {
    var maxSidePx = 280;
    var minSidePx = 52;
    var dominant = Math.max(frameWidthCm, frameHeightCm);
    var widthPx = frameWidthCm / dominant * maxSidePx;
    var heightPx = frameHeightCm / dominant * maxSidePx;
    var shortSide = Math.min(widthPx, heightPx);
    if (shortSide < minSidePx) {
      var scaleUp = minSidePx / shortSide;
      widthPx *= scaleUp;
      heightPx *= scaleUp;
    }
    largeShapeImagePreviewImg.style.width = "".concat(Math.round(widthPx), "px");
    largeShapeImagePreviewImg.style.height = "".concat(Math.round(heightPx), "px");
    largeShapeImagePreviewImg.style.maxWidth = "none";
    largeShapeImagePreviewImg.style.maxHeight = "none";
    largeShapeImagePreviewImg.style.objectFit = "fill";
  } else {
    largeShapeImagePreviewImg.style.width = "";
    largeShapeImagePreviewImg.style.height = "";
    largeShapeImagePreviewImg.style.maxWidth = "";
    largeShapeImagePreviewImg.style.maxHeight = "";
    largeShapeImagePreviewImg.style.objectFit = "contain";
  }
  largeShapeImagePreviewImg.style.transformOrigin = "center center";
  largeShapeImagePreviewImg.style.transform = "rotate(".concat(rotate, "deg) scale(").concat(scaleX, ", ").concat(scaleY, ") skew(").concat(skewX, "deg, ").concat(skewY, "deg)");
}
function syncImageFrameSizeFromAspectLock() {
  return _syncImageFrameSizeFromAspectLock.apply(this, arguments);
}
function _syncImageFrameSizeFromAspectLock() {
  _syncImageFrameSizeFromAspectLock = _asyncToGenerator(_regenerator().m(function _callee1() {
    var _recordForm$elements$, _recordForm$elements$2;
    var triggerFieldName,
      planSizeMode,
      shapeType,
      lockInput,
      widthInput,
      heightInput,
      aspectRatio,
      loadedAspect,
      widthCm,
      heightCm,
      hasWidth,
      hasHeight,
      _args1 = arguments;
    return _regenerator().w(function (_context1) {
      while (1) switch (_context1.n) {
        case 0:
          triggerFieldName = _args1.length > 0 && _args1[0] !== undefined ? _args1[0] : "";
          if (recordForm !== null && recordForm !== void 0 && recordForm.elements) {
            _context1.n = 1;
            break;
          }
          return _context1.a(2);
        case 1:
          planSizeMode = normalizePlanSizeMode((_recordForm$elements$ = recordForm.elements.planSizeMode) === null || _recordForm$elements$ === void 0 ? void 0 : _recordForm$elements$.value);
          shapeType = normalizeLargeShapeType((_recordForm$elements$2 = recordForm.elements.largeShapeType) === null || _recordForm$elements$2 === void 0 ? void 0 : _recordForm$elements$2.value);
          if (!(planSizeMode !== "大きなもの" || !isCustomLargeShapeType(shapeType))) {
            _context1.n = 2;
            break;
          }
          return _context1.a(2);
        case 2:
          lockInput = recordForm.elements.imgLockAspectRatio;
          if (!(!(lockInput instanceof HTMLInputElement) || !lockInput.checked)) {
            _context1.n = 3;
            break;
          }
          return _context1.a(2);
        case 3:
          widthInput = recordForm.elements.imgFrameWidthCm;
          heightInput = recordForm.elements.imgFrameHeightCm;
          if (!(!(widthInput instanceof HTMLInputElement) || !(heightInput instanceof HTMLInputElement))) {
            _context1.n = 4;
            break;
          }
          return _context1.a(2);
        case 4:
          aspectRatio = Number(normalizeCustomLargeImageAspect(customLargeImageAspectInput === null || customLargeImageAspectInput === void 0 ? void 0 : customLargeImageAspectInput.value));
          if (aspectRatio > 0) {
            _context1.n = 6;
            break;
          }
          _context1.n = 5;
          return updateCustomLargeImageAspectFromDataUrl(customLargeImageDataUrlInput === null || customLargeImageDataUrlInput === void 0 ? void 0 : customLargeImageDataUrlInput.value);
        case 5:
          loadedAspect = _context1.v;
          aspectRatio = Number(loadedAspect);
        case 6:
          if (aspectRatio > 0) {
            _context1.n = 7;
            break;
          }
          return _context1.a(2);
        case 7:
          widthCm = parseDistanceToCm(widthInput.value);
          heightCm = parseDistanceToCm(heightInput.value);
          hasWidth = Number.isFinite(widthCm) && widthCm > 0;
          hasHeight = Number.isFinite(heightCm) && heightCm > 0;
          if (!(triggerFieldName === "imgFrameWidthCm" && hasWidth)) {
            _context1.n = 8;
            break;
          }
          heightInput.value = formatLengthInputValue(widthCm / aspectRatio);
          return _context1.a(2);
        case 8:
          if (!(triggerFieldName === "imgFrameHeightCm" && hasHeight)) {
            _context1.n = 9;
            break;
          }
          widthInput.value = formatLengthInputValue(heightCm * aspectRatio);
          return _context1.a(2);
        case 9:
          if (!(hasWidth && !hasHeight)) {
            _context1.n = 10;
            break;
          }
          heightInput.value = formatLengthInputValue(widthCm / aspectRatio);
          return _context1.a(2);
        case 10:
          if (!hasWidth && hasHeight) {
            widthInput.value = formatLengthInputValue(heightCm * aspectRatio);
          }
        case 11:
          return _context1.a(2);
      }
    }, _callee1);
  }));
  return _syncImageFrameSizeFromAspectLock.apply(this, arguments);
}
function resetRecordForm(_ref9) {
  var showMessage = _ref9.showMessage;
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
  ["imgRotateDeg", "imgFrameWidthCm", "imgFrameHeightCm", "imgSkewXDeg", "imgSkewYDeg"].forEach(function (name) {
    var _recordForm$elements5;
    var field = recordForm === null || recordForm === void 0 || (_recordForm$elements5 = recordForm.elements) === null || _recordForm$elements5 === void 0 ? void 0 : _recordForm$elements5.namedItem(name);
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
  syncDirectionTabsFromForm();
  syncAltitudeDirectInputUi({
    clearWhenDisabled: true
  });
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
  var parsedSpecimen = parseSpecimenNo(record.specimenNo, record.specimenPrefix, record.specimenSerial);
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
    recordForm.elements.sectionDiagramLayerFaciesChecked.checked = normalizeChecklistChecked(record.sectionDiagramLayerFaciesChecked) === "1";
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
  renderMultiPointRows(record.multiPoints);
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
  syncDirectionTabsFromForm();
  syncAltitudeDirectInputUi();
  setLayerFromValue(record.layerName);
  recordForm.elements.detail.value = record.detail || "";
  recordForm.elements.detailSub.value = record.detailSub || "";
  recordForm.elements.layerFacies.value = record.layerFacies || "";
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
  ["imgP1NsDir", "imgP1NsCm", "imgP1EwDir", "imgP1EwCm", "imgP2NsDir", "imgP2NsCm", "imgP2EwDir", "imgP2EwCm", "imgP3NsDir", "imgP3NsCm", "imgP3EwDir", "imgP3EwCm", "imgP4NsDir", "imgP4NsCm", "imgP4EwDir", "imgP4EwCm"].forEach(function (name) {
    var field = recordForm.elements.namedItem(name);
    if (field instanceof HTMLInputElement || field instanceof HTMLSelectElement) {
      field.value = value(record === null || record === void 0 ? void 0 : record[name]);
    }
  });
  setDefaultImageCornerDirections();
  ["imgP1NsDir", "imgP1EwDir", "imgP2NsDir", "imgP2EwDir", "imgP3NsDir", "imgP3EwDir", "imgP4NsDir", "imgP4EwDir"].forEach(function (name) {
    var field = recordForm.elements.namedItem(name);
    if (field instanceof HTMLSelectElement) {
      var savedValue = value(record === null || record === void 0 ? void 0 : record[name]);
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
  void syncImageFrameSizeFromAspectLock("populate").then(function () {
    syncLargeShapeImagePreviewTransform();
  });
  updateEditMissingRequiredHighlights();
}
function openRecordForEdit(recordId) {
  var _state$site3, _state$site4;
  var preferredKuwaku = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : "";
  var recordIndexRaw = arguments.length > 2 && arguments[2] !== undefined ? arguments[2] : "";
  var record = findRecordByEditContext(recordId, recordIndexRaw, null);
  if (!record) {
    showToast("対象データが見つかりません");
    return;
  }
  var kuwakuSource = value(preferredKuwaku) || value(record.kuwaku) || getRecordKuwaku(record);
  var kuwakuParts = parseKuwaku(kuwakuSource);
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
    var _state$site;
    editLevelHeightInput.value = value(record.levelHeight) || value((_state$site = state.site) === null || _state$site === void 0 ? void 0 : _state$site.levelHeight);
  }
  if (editDateInput) {
    var _state$site2;
    editDateInput.value = value(record.date) || value((_state$site2 = state.site) === null || _state$site2 === void 0 ? void 0 : _state$site2.date);
  }
  var editTeamState = normalizeTeamState(value(record.team) || value((_state$site3 = state.site) === null || _state$site3 === void 0 ? void 0 : _state$site3.team), value(record.teamOther) || value((_state$site4 = state.site) === null || _state$site4 === void 0 ? void 0 : _state$site4.teamOther));
  if (editTeamInput) {
    editTeamInput.value = editTeamState.team;
  }
  if (editTeamOtherInput) {
    editTeamOtherInput.value = editTeamState.teamOther;
  }
  syncEditTeamOtherInput(editTeamState.team);
  if (editTeamLeadInput) {
    var _state$site5;
    editTeamLeadInput.value = value(record.teamLead) || value((_state$site5 = state.site) === null || _state$site5 === void 0 ? void 0 : _state$site5.teamLead);
  }
  if (editRecorderInput) {
    var _state$site6;
    editRecorderInput.value = value(record.recorder) || value((_state$site6 = state.site) === null || _state$site6 === void 0 ? void 0 : _state$site6.recorder);
  }
  isOverwriteMode = true;
  overwriteOriginalRecord = _objectSpread({}, record);
  var recordIndex = state.records.findIndex(function (item) {
    return item === record;
  });
  activeEditRecordContext = {
    recordId: value(record.id),
    recordIndex: String(recordIndex >= 0 ? recordIndex : ""),
    recordSnapshot: buildCellEditRecordSnapshot(record)
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
  var anchor = recordForm.querySelector(".detail-input-title-row") || recordForm;
  window.requestAnimationFrame(function () {
    if (anchor instanceof HTMLElement) {
      anchor.scrollIntoView({
        behavior: "smooth",
        block: "start"
      });
    }
  });
}
function buildCurrentEditDraftRecord() {
  if (!recordForm) {
    return null;
  }
  var formData = new FormData(recordForm);
  var teamState = normalizeTeamState(value(editTeamInput === null || editTeamInput === void 0 ? void 0 : editTeamInput.value), value(editTeamOtherInput === null || editTeamOtherInput === void 0 ? void 0 : editTeamOtherInput.value));
  var draftPlanSizeMode = normalizePlanSizeMode(value(formData.get("planSizeMode")));
  var draftMultiPoints = draftPlanSizeMode === "複数点" ? readMultiPointRowsFromForm() : [];
  var specimenPrefix = normalizeSpecimenPrefix(value(formData.get("specimenPrefix")));
  var specimenSerial = compactNoSpaceValue(formData.get("specimenSerial"));
  var draftRawShapeType = value(formData.get("largeShapeType"));
  var draftShapeType = normalizeLargeShapeType(draftRawShapeType) || normalizeLargeShapeLabel(draftRawShapeType);
  var draftIsImageShape = isLargeShapeImageType(draftShapeType);
  var draftIsCustomImageShape = draftIsImageShape && isCustomLargeShapeType(draftShapeType);
  var draftIsLineShape = draftShapeType === "直線状";
  var draftUsesAxisDirection = draftIsLineShape || draftShapeType === "長方形" || draftShapeType === "楕円";
  var imageCornerFields = extractImageCornerFieldsFromFormData(formData);
  var imageTransformFields = extractImageTransformFieldsFromFormData(formData);
  return {
    kuwaku: buildKuwaku(normalizeKuwakuHeadA(editKuwakuHeadAInput === null || editKuwakuHeadAInput === void 0 ? void 0 : editKuwakuHeadAInput.value), normalizeKuwakuHeadB(editKuwakuHeadBInput === null || editKuwakuHeadBInput === void 0 ? void 0 : editKuwakuHeadBInput.value), normalizeKuwakuBlock(editKuwakuBlockInput === null || editKuwakuBlockInput === void 0 ? void 0 : editKuwakuBlockInput.value), normalizeKuwakuNo(editKuwakuNoInput === null || editKuwakuNoInput === void 0 ? void 0 : editKuwakuNoInput.value)),
    levelHeight: value(editLevelHeightInput === null || editLevelHeightInput === void 0 ? void 0 : editLevelHeightInput.value),
    date: value(editDateInput === null || editDateInput === void 0 ? void 0 : editDateInput.value),
    team: teamState.team,
    teamOther: teamState.teamOther,
    teamLead: value(editTeamLeadInput === null || editTeamLeadInput === void 0 ? void 0 : editTeamLeadInput.value),
    recorder: value(editRecorderInput === null || editRecorderInput === void 0 ? void 0 : editRecorderInput.value),
    specimenPrefix: specimenPrefix,
    specimenSerial: specimenSerial,
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
    altitudeDirectM: normalizeToggleFlag(formData.get("altitudeInputEnabled")) === "1" ? value(formData.get("altitudeDirectM")) : "",
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
    multiPoints: draftMultiPoints,
    planSizeMode: draftPlanSizeMode,
    largeShapeType: draftShapeType,
    largeAxisDirection: draftIsImageShape || !draftUsesAxisDirection ? "" : normalizeLargeAxisDirection(value(formData.get("largeAxisDirection"))),
    largeAxisPlungeDeg: draftIsImageShape || !draftIsLineShape ? "" : normalizeLargeAxisPlungeDeg(value(formData.get("largeAxisPlungeDeg"))),
    largeAxisPlungeDir8: draftIsImageShape || !draftIsLineShape ? "" : normalizeCompass8Direction(value(formData.get("largeAxisPlungeDir8"))),
    planeStrikeDirection: draftIsLineShape ? "" : normalizePlaneStrikeDirection(value(formData.get("planeStrikeDirection")) || (draftUsesAxisDirection ? value(formData.get("largeAxisDirection")) : "")),
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
    layerFacies: value(formData.get("layerFacies")),
    layerRef: value(formData.get("layerRef")),
    layerRelative: value(formData.get("layerRelative")),
    layerFromCm: value(formData.get("layerFromCm")),
    notes: value(formData.get("notes")),
    sectionDiagrams: clonePhotos(currentSectionDiagrams),
    photos: clonePhotos(currentPhotos)
  };
}
function copyCurrentEditToInput() {
  var activeTabId = getActiveTabId();
  if (activeTabId !== "edit-tab" && activeTabId !== "input-tab") {
    return;
  }
  var draft = buildCurrentEditDraftRecord();
  if (!draft) {
    showToast("入力内容のコピーに失敗しました");
    return;
  }
  if (activeTabId === "input-tab" && siteForm) {
    var siteFormData = new FormData(siteForm);
    var siteKuwakuHeadA = normalizeKuwakuHeadA(siteFormData.get("kuwakuHeadA"));
    var siteKuwakuHeadB = normalizeKuwakuHeadB(siteFormData.get("kuwakuHeadB"));
    var siteKuwakuBlock = normalizeKuwakuBlock(siteFormData.get("kuwakuBlock"));
    var siteKuwakuNo = normalizeKuwakuNo(siteFormData.get("kuwakuNo"));
    var siteTeamState = normalizeTeamState(value(siteFormData.get("team")), value(siteFormData.get("teamOther")));
    draft.kuwaku = buildKuwaku(siteKuwakuHeadA, siteKuwakuHeadB, siteKuwakuBlock, siteKuwakuNo);
    draft.levelHeight = value(siteFormData.get("levelHeight"));
    draft.date = value(siteFormData.get("date"));
    draft.team = siteTeamState.team;
    draft.teamOther = siteTeamState.teamOther;
    draft.teamLead = value(siteFormData.get("teamLead"));
    draft.recorder = value(siteFormData.get("recorder"));
  }
  var kuwakuParts = parseKuwaku(draft.kuwaku);
  var teamState = normalizeTeamState(draft.team, draft.teamOther);
  if (siteForm !== null && siteForm !== void 0 && siteForm.elements) {
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
  populateRecordForm(_objectSpread(_objectSpread({}, draft), {}, {
    id: "",
    sectionDiagrams: [],
    photos: [],
    sectionDiagramDistanceChecked: "",
    sectionDiagramHorizonChecked: "",
    sectionDiagramLayerFaciesChecked: "",
    photoClinometerChecked: "",
    photoRulerChecked: ""
  }));
  editingRecordId = null;
  activeEditRecordContext = null;
  if (recordIdInput) {
    recordIdInput.value = "";
  }
  updateDuplicateSpecimenWarning();
  showToast("コピーして新規入力を作成しました");
}
function copySavedRecordToInput(recordId) {
  var preferredKuwaku = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : "";
  var recordRaw = arguments.length > 2 && arguments[2] !== undefined ? arguments[2] : null;
  var options = arguments.length > 3 && arguments[3] !== undefined ? arguments[3] : {};
  var record = recordRaw && _typeof(recordRaw) === "object" ? recordRaw : findRecord(recordId);
  if (!record) {
    showToast("対象データが見つかりません");
    return;
  }
  var shouldShowToast = options && _typeof(options) === "object" ? options.showToast !== false : true;
  var kuwakuSource = value(preferredKuwaku) || value(record.kuwaku) || getRecordKuwaku(record);
  var kuwakuParts = parseKuwaku(kuwakuSource);
  var teamState = normalizeTeamState(value(record.team), value(record.teamOther));
  if (siteForm !== null && siteForm !== void 0 && siteForm.elements) {
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
  populateRecordForm(_objectSpread(_objectSpread({}, record), {}, {
    id: "",
    sectionDiagrams: [],
    photos: [],
    sectionDiagramDistanceChecked: "",
    sectionDiagramHorizonChecked: "",
    sectionDiagramLayerFaciesChecked: "",
    photoClinometerChecked: "",
    photoRulerChecked: ""
  }));
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
function insertRowFromList(recordId) {
  var preferredKuwaku = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : "";
  var recordRaw = arguments.length > 2 && arguments[2] !== undefined ? arguments[2] : null;
  var record = recordRaw && _typeof(recordRaw) === "object" ? recordRaw : findRecord(recordId);
  if (!record) {
    showToast("対象データが見つかりません");
    return;
  }
  var kuwaku = normalizeKuwakuText(value(preferredKuwaku) || value(record.kuwaku) || getRecordKuwaku(record));
  var teamState = normalizeTeamState(value(record.team), value(record.teamOther));
  var nowIsoValue = nowIso();
  var insertedBase = {
    id: newId("record"),
    kuwaku: kuwaku,
    specimenPrefix: DEFAULT_SPECIMEN_PREFIX,
    specimenSerial: "",
    specimenNo: "",
    category: categoryFromPrefix(DEFAULT_SPECIMEN_PREFIX),
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
    layerFacies: "",
    layerRef: "",
    layerFromCm: "",
    layerRelative: "",
    notes: "",
    createdAt: nowIsoValue,
    updatedAt: nowIsoValue,
    deletedAt: ""
  };
  var insertedRecord = _objectSpread(_objectSpread({}, insertedBase), {}, {
    history: buildNextRecordHistory(null, insertedBase, "行挿入")
  });
  state.records.unshift(insertedRecord);
  persist("行挿入しました");
  if (getActiveTabId() !== "output-tab") {
    setActiveTab("output-tab");
  }
  selectedCardRecordId = "";
  renderRecordTable();
  renderOutputs();
  var insertedIndex = state.records.findIndex(function (item) {
    return item === insertedRecord;
  });
  window.requestAnimationFrame(function () {
    openOutputCellEditModal(insertedRecord.id, "category", String(insertedIndex >= 0 ? insertedIndex : ""));
  });
}
function getRecordFormFieldByName(name) {
  if (!(recordForm !== null && recordForm !== void 0 && recordForm.elements)) {
    return null;
  }
  var field = recordForm.elements.namedItem(name);
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
  var field = getRecordFormFieldByName(name);
  if (field) {
    field.classList.add("edit-missing-field");
  }
}
function markEditMissingGroupByName(name) {
  var field = getRecordFormFieldByName(name);
  var group = field === null || field === void 0 ? void 0 : field.closest(".inline-fieldset");
  if (group) {
    group.classList.add("edit-missing-group");
  }
}
function clearEditMissingRequiredHighlights() {
  document.querySelectorAll(".edit-missing-field").forEach(function (element) {
    element.classList.remove("edit-missing-field");
  });
  document.querySelectorAll(".edit-missing-group").forEach(function (element) {
    element.classList.remove("edit-missing-group");
  });
}
function updateEditMissingRequiredHighlights() {
  clearEditMissingRequiredHighlights();
  if (getActiveTabId() !== "edit-tab") {
    return;
  }
  var draftRecord = buildCurrentEditDraftRecord();
  var missingKeys = getMissingRequiredKeys(draftRecord);
  if (!missingKeys.size) {
    return;
  }
  if (hasAnyMissingRequiredKey(missingKeys, ["kuwakuHeadA", "kuwakuHeadB", "kuwakuBlock", "kuwakuNo"])) {
    [editKuwakuHeadAInput, editKuwakuHeadBInput, editKuwakuBlockInput, editKuwakuNoInput].forEach(function (input) {
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
  ["imgRotateDeg", "imgFrameWidthCm", "imgFrameHeightCm", "customLargeImageName", "customLargeImageDataUrl", "imgP1NsCm", "imgP1EwCm", "imgP2NsCm", "imgP2EwCm", "imgP3NsCm", "imgP3EwCm", "imgP4NsCm", "imgP4EwCm"].forEach(function (name) {
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
  if (hasAnyMissingRequiredKey(missingKeys, ["sectionDiagrams", "sectionDiagramDistanceChecked", "sectionDiagramHorizonChecked", "sectionDiagramLayerFaciesChecked"])) {
    var sectionWrap = recordForm === null || recordForm === void 0 ? void 0 : recordForm.querySelector(".diagram-upload-wrap");
    if (sectionWrap) {
      sectionWrap.classList.add("edit-missing-group");
    }
  }
  if (hasAnyMissingRequiredKey(missingKeys, ["photoClinometerChecked", "photoRulerChecked"])) {
    var photoWrap = recordForm === null || recordForm === void 0 ? void 0 : recordForm.querySelector(".photo-upload-wrap");
    if (photoWrap) {
      photoWrap.classList.add("edit-missing-group");
    }
  }
}
function syncEditHistoryVisibility() {
  var activeTabId = arguments.length > 0 && arguments[0] !== undefined ? arguments[0] : getActiveTabId();
  if (!editHistoryPanel) {
    return;
  }
  var shouldShow = activeTabId === "edit-tab" && Boolean(editingRecordId);
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
  var history = normalizeRecordHistory(record === null || record === void 0 ? void 0 : record.history);
  if (!history.length) {
    editHistoryList.innerHTML = "<p class=\"muted\">履歴はまだありません。</p>";
  } else {
    var displayHistory = history.map(function (entry, index) {
      return {
        entry: entry,
        prevEntry: index > 0 ? history[index - 1] : null
      };
    }).reverse();
    editHistoryList.innerHTML = displayHistory.map(function (_ref0) {
      var entry = _ref0.entry,
        prevEntry = _ref0.prevEntry;
      var contentHtml = renderHistoryContentHtml(entry, prevEntry);
      return "\n          <article class=\"edit-history-item\">\n            <p><strong>\u5165\u529B\u5185\u5BB9:</strong> ".concat(contentHtml, "</p>\n            <p><strong>\u5E74\u30FB\u6708\u65E5\u30FB\u6642\u9593:</strong> ").concat(escapeHtml(formatHistoryDateTime(entry.at)), "</p>\n          </article>\n        ");
    }).join("");
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
  var inputRecords = getInputRecordsForCurrentKuwaku();
  var editRecords = getEditRecordsForCurrentKuwaku();
  renderRecordTableBodyRows(recordTableBody, inputRecords, inputRecords.length ? "" : "現在の区画（グリッド）の入力データがありません。");
  renderRecordTableBodyRows(editRecordTableBody, editRecords, editRecords.length ? "" : "編集中の区画（グリッド）の入力データがありません。");
}
function getInputRecordsForCurrentKuwaku() {
  var _state$site7;
  var currentKuwaku = value((_state$site7 = state.site) === null || _state$site7 === void 0 ? void 0 : _state$site7.kuwaku);
  var sortedRecords = _toConsumableArray(state.records).sort(compareRecordsByKuwakuThenSpecimen);
  if (!currentKuwaku || isDefaultKuwaku(currentKuwaku)) {
    return sortedRecords;
  }
  var currentValue = kuwakuValueForSelect(currentKuwaku);
  return sortedRecords.filter(function (record) {
    return kuwakuValueForSelect(getRecordKuwaku(record)) === currentValue;
  });
}
function getEditRecordsForCurrentKuwaku() {
  var editKuwaku = currentKuwakuForDuplicateWarning("edit-tab");
  var sortedRecords = _toConsumableArray(state.records).sort(compareRecordsByKuwakuThenSpecimen);
  if (!editKuwaku) {
    return [];
  }
  var currentValue = kuwakuValueForSelect(editKuwaku);
  return sortedRecords.filter(function (record) {
    return kuwakuValueForSelect(getRecordKuwaku(record)) === currentValue;
  });
}
function renderRecordTableBodyRows(targetBody, records, emptyMessage) {
  if (!targetBody) {
    return;
  }
  if (!records.length) {
    targetBody.innerHTML = "<tr><td colspan=\"8\">".concat(escapeHtml(emptyMessage || "表示対象データがありません。"), "</td></tr>");
    return;
  }
  var recordIndexMap = new Map();
  state.records.forEach(function (record, index) {
    if (record && _typeof(record) === "object") {
      recordIndexMap.set(record, index);
    }
  });
  targetBody.innerHTML = records.map(function (record) {
    var recordIndex = Number(recordIndexMap.get(record));
    return buildRecordTableRowHtml(record, Number.isInteger(recordIndex) && recordIndex >= 0 ? recordIndex : "");
  }).join("");
}
function buildRecordTableRowHtml(record) {
  var recordIndexRaw = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : "";
  var recordIndex = value(recordIndexRaw);
  return "\n      <tr>\n        <td>".concat(escapeHtml(getRecordKuwaku(record)), "</td>\n        <td>").concat(escapeHtml(getRecordTeamValue(record)), "</td>\n        <td>").concat(escapeHtml(record.specimenNo), "</td>\n        <td>").concat(escapeHtml(formatCategoryForRecord(record)), "</td>\n        <td>").concat(escapeHtml(record.nameMemo || ""), "</td>\n        <td>").concat(escapeHtml(record.discoverer || ""), "</td>\n        <td>").concat(escapeHtml(formatLevelRead(record)), "</td>\n        <td>\n          <div class=\"row-actions\">\n            <button type=\"button\" data-action=\"insert-row\" data-id=\"").concat(record.id, "\" data-kuwaku=\"").concat(escapeHtml(getRecordKuwaku(record)), "\" data-record-index=\"").concat(escapeHtml(recordIndex), "\">\u884C\u633F\u5165</button>\n            <button type=\"button\" data-action=\"edit\" data-id=\"").concat(record.id, "\" data-kuwaku=\"").concat(escapeHtml(getRecordKuwaku(record)), "\" data-record-index=\"").concat(escapeHtml(recordIndex), "\">\u7DE8\u96C6</button>\n            <button type=\"button\" data-action=\"copy-to-input\" data-id=\"").concat(record.id, "\" data-kuwaku=\"").concat(escapeHtml(getRecordKuwaku(record)), "\" data-record-index=\"").concat(escapeHtml(recordIndex), "\">\u30B3\u30D4\u30FC\u3057\u3066\u65B0\u898F\u5165\u529B</button>\n            <button class=\"danger\" type=\"button\" data-action=\"delete\" data-id=\"").concat(record.id, "\" data-record-index=\"").concat(escapeHtml(recordIndex), "\">\u524A\u9664</button>\n          </div>\n        </td>\n      </tr>\n      ");
}
function handleRecordTableActionClick(event) {
  var _row$dataset4;
  var button = event.target.closest("button[data-action]");
  if (!button) {
    return;
  }
  var sourceTableBody = event.currentTarget;
  var shouldScrollToDetailTop = sourceTableBody === recordTableBody || sourceTableBody === editRecordTableBody;
  var recordId = button.dataset.id;
  var row = button.closest("tr[data-record-index]");
  var recordIndex = value(button.dataset.recordIndex) || value(row === null || row === void 0 || (_row$dataset4 = row.dataset) === null || _row$dataset4 === void 0 ? void 0 : _row$dataset4.recordIndex);
  var action = button.dataset.action;
  var record = findRecordByEditContext(recordId, recordIndex, null);
  if (!record) {
    showToast("対象データが見つかりません");
    return;
  }
  if (action === "edit") {
    var rowKuwaku = value(button.dataset.kuwaku);
    openRecordForEdit(record.id, rowKuwaku, recordIndex);
    if (shouldScrollToDetailTop) {
      scrollToDetailInputTop();
    }
    return;
  }
  if (action === "copy-to-input") {
    var _rowKuwaku3 = value(button.dataset.kuwaku);
    copySavedRecordToInput(recordId, _rowKuwaku3, record);
    return;
  }
  if (action === "insert-row") {
    var _rowKuwaku4 = value(button.dataset.kuwaku);
    insertRowFromList(recordId, _rowKuwaku4, record);
    return;
  }
  if (action === "delete") {
    var answer = window.confirm("\u6A19\u672C\u756A\u53F7 ".concat(record.specimenNo, " \u3092\u524A\u9664\u3057\u307E\u3059\u304B\uFF1F"));
    if (!answer) {
      return;
    }
    var deletingId = value(record.id) || recordId;
    state.records = state.records.filter(function (item) {
      return item !== record;
    });
    if (editingRecordId === deletingId) {
      resetRecordForm({
        showMessage: false
      });
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
    renderViewerOutput({
      preserveCamera: true
    });
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
  var kuwakuScopedSource = getRecordsByExportRangeFilters({
    kuwakuValue: ALL_GRIDS_VALUE,
    categoryValue: exportListRangeCategory,
    statusValue: exportListRangeStatus,
    specimenFromRaw: exportListRangeSpecimenFrom,
    specimenToRaw: exportListRangeSpecimenTo,
    dateFromRaw: exportListRangeDateFrom,
    dateToRaw: exportListRangeDateTo
  });
  var kuwakuOptions = collectExportKuwakuOptionsWithCounts(kuwakuScopedSource);
  if (!kuwakuOptions.some(function (item) {
    return item.value === exportListRangeKuwaku;
  })) {
    exportListRangeKuwaku = ALL_GRIDS_VALUE;
  }
  if (exportListRangeKuwakuSelect) {
    exportListRangeKuwakuSelect.innerHTML = kuwakuOptions.map(function (item) {
      return "<option value=\"".concat(escapeHtml(item.value), "\" ").concat(item.value === exportListRangeKuwaku ? "selected" : "", ">").concat(escapeHtml(item.label), "</option>");
    }).join("");
  }
  var categoryScopedSource = getRecordsByExportRangeFilters({
    kuwakuValue: exportListRangeKuwaku,
    categoryValue: EXPORT_CATEGORY_ALL_VALUE,
    statusValue: exportListRangeStatus,
    specimenFromRaw: exportListRangeSpecimenFrom,
    specimenToRaw: exportListRangeSpecimenTo,
    dateFromRaw: exportListRangeDateFrom,
    dateToRaw: exportListRangeDateTo
  });
  var categoryOptions = collectExportCategoryOptions(categoryScopedSource);
  if (!categoryOptions.some(function (item) {
    return item.value === exportListRangeCategory;
  })) {
    exportListRangeCategory = EXPORT_CATEGORY_ALL_VALUE;
  }
  if (exportListRangeCategorySelect) {
    exportListRangeCategorySelect.innerHTML = categoryOptions.map(function (item) {
      return "<option value=\"".concat(escapeHtml(item.value), "\" ").concat(item.value === exportListRangeCategory ? "selected" : "", ">").concat(escapeHtml(item.label), "</option>");
    }).join("");
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
  var filteredRecords = getListExportRecords();
  if (exportListRangeSummaryEl) {
    var hasData = filteredRecords.length > 0;
    exportListRangeSummaryEl.textContent = "\u5BFE\u8C61\u4EF6\u6570: ".concat(filteredRecords.length, "\u4EF6\uFF08").concat(hasData ? "対象あり" : "対象なし", "\uFF09");
    setAvailabilityClass(exportListRangeSummaryEl, hasData);
  }
}
function renderExportCardRangeControls() {
  var kuwakuScopedSource = getRecordsByExportRangeFilters({
    kuwakuValue: ALL_GRIDS_VALUE,
    categoryValue: exportCardRangeCategory,
    statusValue: exportCardRangeStatus,
    dateFromRaw: exportCardRangeDateFrom,
    dateToRaw: exportCardRangeDateTo
  });
  var kuwakuOptions = collectExportKuwakuOptionsWithCounts(kuwakuScopedSource);
  if (!kuwakuOptions.some(function (item) {
    return item.value === exportCardRangeKuwaku;
  })) {
    exportCardRangeKuwaku = ALL_GRIDS_VALUE;
  }
  if (exportCardRangeKuwakuSelect) {
    exportCardRangeKuwakuSelect.innerHTML = kuwakuOptions.map(function (item) {
      return "<option value=\"".concat(escapeHtml(item.value), "\" ").concat(item.value === exportCardRangeKuwaku ? "selected" : "", ">").concat(escapeHtml(item.label), "</option>");
    }).join("");
  }
  var categoryScopedSource = getRecordsByExportRangeFilters({
    kuwakuValue: exportCardRangeKuwaku,
    categoryValue: EXPORT_CATEGORY_ALL_VALUE,
    statusValue: exportCardRangeStatus,
    dateFromRaw: exportCardRangeDateFrom,
    dateToRaw: exportCardRangeDateTo
  });
  var categoryOptions = collectExportCategoryOptions(categoryScopedSource);
  if (!categoryOptions.some(function (item) {
    return item.value === exportCardRangeCategory;
  })) {
    exportCardRangeCategory = EXPORT_CATEGORY_ALL_VALUE;
  }
  if (exportCardRangeCategorySelect) {
    exportCardRangeCategorySelect.innerHTML = categoryOptions.map(function (item) {
      return "<option value=\"".concat(escapeHtml(item.value), "\" ").concat(item.value === exportCardRangeCategory ? "selected" : "", ">").concat(escapeHtml(item.label), "</option>");
    }).join("");
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
  var filteredRecords = getCardExportRecords();
  if (exportCardRangeSummaryEl) {
    var hasData = filteredRecords.length > 0;
    exportCardRangeSummaryEl.textContent = "\u5BFE\u8C61\u4EF6\u6570: ".concat(filteredRecords.length, "\u4EF6\uFF08").concat(hasData ? "対象あり" : "対象なし", "\uFF09");
    setAvailabilityClass(exportCardRangeSummaryEl, hasData);
  }
}
function renderExportPlanControls() {
  var sortedRecords = getRecordsByExportRangeFilters({
    kuwakuValue: ALL_GRIDS_VALUE,
    categoryValue: EXPORT_CATEGORY_ALL_VALUE,
    statusValue: "all",
    dateFromRaw: exportPlanDateFrom,
    dateToRaw: exportPlanDateTo
  });
  var kuwakuOptions = collectExportKuwakuOptionsWithCounts(sortedRecords).filter(function (item) {
    return item.value !== ALL_GRIDS_VALUE;
  });
  if (!kuwakuOptions.length) {
    exportPlanKuwaku = "";
  } else if (!kuwakuOptions.some(function (item) {
    return item.value === exportPlanKuwaku;
  })) {
    exportPlanKuwaku = kuwakuOptions[0].value;
  }
  if (exportPlanKuwakuSelect) {
    exportPlanKuwakuSelect.innerHTML = kuwakuOptions.map(function (item) {
      return "<option value=\"".concat(escapeHtml(item.value), "\" ").concat(item.value === exportPlanKuwaku ? "selected" : "", ">").concat(escapeHtml(item.label), "</option>");
    }).join("");
  }
  var kuwakuScopedRecords = !exportPlanKuwaku ? [] : sortedRecords.filter(function (record) {
    return kuwakuValueForSelect(getRecordKuwaku(record)) === exportPlanKuwaku;
  });
  var categoryScopedRecords = exportPlanCategory === EXPORT_CATEGORY_ALL_VALUE ? kuwakuScopedRecords : kuwakuScopedRecords.filter(function (record) {
    var specimen = parseSpecimenNo(record.specimenNo, record.specimenPrefix, record.specimenSerial);
    return normalizeSpecimenPrefix(specimen.prefix) === exportPlanCategory;
  });
  var categoryOptions = collectExportCategoryOptions(kuwakuScopedRecords);
  if (!categoryOptions.some(function (item) {
    return item.value === exportPlanCategory;
  })) {
    exportPlanCategory = EXPORT_CATEGORY_ALL_VALUE;
  }
  if (exportPlanCategorySelect) {
    exportPlanCategorySelect.innerHTML = categoryOptions.map(function (item) {
      return "<option value=\"".concat(escapeHtml(item.value), "\" ").concat(item.value === exportPlanCategory ? "selected" : "", ">").concat(escapeHtml(item.label), "</option>");
    }).join("");
  }
  if (exportPlanDateFromInput) {
    exportPlanDateFromInput.value = exportPlanDateFrom;
  }
  if (exportPlanDateToInput) {
    exportPlanDateToInput.value = exportPlanDateTo;
  }
  syncExportPlanModeControls(categoryScopedRecords);
  var groups = buildPlanPdfGroupsForExport({
    kuwakuValue: exportPlanKuwaku,
    categoryValue: exportPlanCategory,
    dateFromRaw: exportPlanDateFrom,
    dateToRaw: exportPlanDateTo,
    modeSelections: getExportPlanModeSelections()
  });
  var recordCount = groups.reduce(function (sum, group) {
    return sum + (Number.isFinite(group.count) ? group.count : 0);
  }, 0);
  if (exportPlanSummaryEl) {
    var hasData = recordCount > 0;
    exportPlanSummaryEl.textContent = "PDF\u30DA\u30FC\u30B8\u5BFE\u8C61: ".concat(groups.length, "\u30DA\u30FC\u30B8 / \u8A18\u9332 ").concat(recordCount, "\u4EF6\uFF08").concat(hasData ? "対象あり" : "対象なし", "\uFF09");
    setAvailabilityClass(exportPlanSummaryEl, hasData);
  }
}
function getExportPlanScopedRecords() {
  var sortedRecords = getRecordsByExportRangeFilters({
    kuwakuValue: ALL_GRIDS_VALUE,
    categoryValue: EXPORT_CATEGORY_ALL_VALUE,
    statusValue: "all",
    dateFromRaw: exportPlanDateFrom,
    dateToRaw: exportPlanDateTo
  });
  if (!exportPlanKuwaku) {
    return [];
  }
  var kuwakuScopedRecords = sortedRecords.filter(function (record) {
    return kuwakuValueForSelect(getRecordKuwaku(record)) === exportPlanKuwaku;
  });
  if (exportPlanCategory === EXPORT_CATEGORY_ALL_VALUE) {
    return kuwakuScopedRecords;
  }
  return kuwakuScopedRecords.filter(function (record) {
    var specimen = parseSpecimenNo(record.specimenNo, record.specimenPrefix, record.specimenSerial);
    return normalizeSpecimenPrefix(specimen.prefix) === exportPlanCategory;
  });
}
function updateExportButtonAvailability() {
  var listCount = getListExportRecords().length;
  var cardCount = getCardExportRecords().length;
  var planGroups = buildPlanPdfGroupsForExport({
    kuwakuValue: exportPlanKuwaku,
    categoryValue: exportPlanCategory,
    dateFromRaw: exportPlanDateFrom,
    dateToRaw: exportPlanDateTo,
    modeSelections: getExportPlanModeSelections()
  });
  var planRecordCount = planGroups.reduce(function (sum, group) {
    return sum + (Number.isFinite(group.count) ? group.count : 0);
  }, 0);
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
  var records = Array.isArray(recordsRaw) ? recordsRaw : [];
  var unitOptions = collectExportPlanValueOptions(records, function (record) {
    return unitValueForSelect(record.unit);
  }, unitLabelForSelect);
  exportPlanModeUnitValues = syncExportPlanModeValues(exportPlanModeUnitValues, unitOptions, {
    autoSelectAllWhenEmpty: !exportPlanModeUnitTouched
  });
  renderExportPlanUnitButtons(exportPlanModeUnitButtons, unitOptions, exportPlanModeUnitValues);
  exportPlanModeDetailUnitValue = syncExportPlanSingleValue(exportPlanModeDetailUnitValue, unitOptions);
  renderExportPlanModeSelect(exportPlanModeDetailUnitSelect, unitOptions, exportPlanModeDetailUnitValue);
  var detailModeRecords = exportPlanModeDetailUnitValue ? filterPlanRecordsForMode(records, {
    unitValues: [exportPlanModeDetailUnitValue]
  }) : [];
  var detailOptions = collectExportPlanValueOptions(detailModeRecords, function (record) {
    return detailValueForSelect(record.detail);
  }, detailLabelForSelect);
  exportPlanModeDetailValues = syncExportPlanModeValues(exportPlanModeDetailValues, detailOptions, {
    autoSelectAllWhenEmpty: !exportPlanModeDetailTouched
  });
  renderExportPlanModeButtons(exportPlanModeDetailButtons, detailOptions, exportPlanModeDetailValues, "出力するサブユニットを選んでください");
  exportPlanModeDetailSubUnitValue = syncExportPlanSingleValue(exportPlanModeDetailSubUnitValue, unitOptions);
  renderExportPlanModeSelect(exportPlanModeDetailSubUnitSelect, unitOptions, exportPlanModeDetailSubUnitValue);
  var detailSubBaseRecords = exportPlanModeDetailSubUnitValue ? filterPlanRecordsForMode(records, {
    unitValues: [exportPlanModeDetailSubUnitValue]
  }) : [];
  var detailSubDetailOptions = collectExportPlanValueOptions(detailSubBaseRecords, function (record) {
    return detailValueForSelect(record.detail);
  }, detailLabelForSelect);
  exportPlanModeDetailSubDetailValue = syncExportPlanSingleValue(exportPlanModeDetailSubDetailValue, detailSubDetailOptions);
  renderExportPlanModeSelect(exportPlanModeDetailSubDetailSelect, detailSubDetailOptions, exportPlanModeDetailSubDetailValue);
  var detailSubRecords = exportPlanModeDetailSubUnitValue && exportPlanModeDetailSubDetailValue ? filterPlanRecordsForMode(records, {
    unitValues: [exportPlanModeDetailSubUnitValue],
    detailValues: [exportPlanModeDetailSubDetailValue]
  }) : [];
  var detailSubOptions = collectExportPlanValueOptions(detailSubRecords, function (record) {
    return detailSubValueForSelect(record.detailSub);
  }, detailSubLabelForSelect);
  exportPlanModeDetailSubValues = syncExportPlanModeValues(exportPlanModeDetailSubValues, detailSubOptions, {
    autoSelectAllWhenEmpty: !exportPlanModeDetailSubTouched
  });
  renderExportPlanModeButtons(exportPlanModeDetailSubButtons, detailSubOptions, exportPlanModeDetailSubValues, "出力するサブユニット細分を選んでください");
  syncExportPlanModeCheckbox(exportPlanModeUnitCheck, unitOptions.length > 0, "unit");
  syncExportPlanModeCheckbox(exportPlanModeDetailCheck, !!exportPlanModeDetailUnitValue && detailOptions.length > 0, "detail");
  syncExportPlanModeCheckbox(exportPlanModeDetailSubCheck, !!exportPlanModeDetailSubUnitValue && !!exportPlanModeDetailSubDetailValue && detailSubOptions.length > 0, "detailSub");
  var unitScopedRecords = filterPlanRecordsForMode(records, {
    unitValues: exportPlanModeUnitValues
  });
  var detailScopedRecords = exportPlanModeDetailUnitValue && exportPlanModeDetailValues.size ? filterPlanRecordsForMode(records, {
    unitValues: [exportPlanModeDetailUnitValue],
    detailValues: exportPlanModeDetailValues
  }) : [];
  var detailSubScopedRecords = exportPlanModeDetailSubUnitValue && exportPlanModeDetailSubDetailValue && exportPlanModeDetailSubValues.size ? filterPlanRecordsForMode(records, {
    unitValues: [exportPlanModeDetailSubUnitValue],
    detailValues: [exportPlanModeDetailSubDetailValue],
    detailSubValues: exportPlanModeDetailSubValues
  }) : [];
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
  var records = Array.isArray(recordsRaw) ? recordsRaw : [];
  var countMap = new Map();
  records.forEach(function (record) {
    var optionValue = value(valueGetter(record));
    if (!optionValue) {
      return;
    }
    countMap.set(optionValue, (countMap.get(optionValue) || 0) + 1);
  });
  return Array.from(countMap.entries()).map(function (_ref1) {
    var _ref10 = _slicedToArray(_ref1, 2),
      optionValue = _ref10[0],
      count = _ref10[1];
    return {
      value: optionValue,
      label: "".concat(labelGetter(optionValue), "\uFF08").concat(count, "\u4EF6\uFF09")
    };
  }).sort(function (a, b) {
    return a.label.localeCompare(b.label, "ja", {
      numeric: true,
      sensitivity: "base"
    });
  });
}
function syncExportPlanModeValues(currentValuesRaw, options) {
  var config = arguments.length > 2 && arguments[2] !== undefined ? arguments[2] : {};
  var next = new Set();
  var currentValues = currentValuesRaw instanceof Set ? currentValuesRaw : new Set();
  var optionValues = (Array.isArray(options) ? options : []).map(function (option) {
    return value(option.value);
  }).filter(Boolean);
  var valid = new Set(optionValues);
  currentValues.forEach(function (selectedValue) {
    var normalized = value(selectedValue);
    if (normalized && valid.has(normalized)) {
      next.add(normalized);
    }
  });
  if (!next.size && optionValues.length && config.autoSelectAllWhenEmpty) {
    optionValues.forEach(function (optionValue) {
      return next.add(optionValue);
    });
  }
  return next;
}
function syncExportPlanSingleValue(currentValueRaw, options) {
  var currentValue = value(currentValueRaw);
  var optionValues = (Array.isArray(options) ? options : []).map(function (option) {
    return value(option.value);
  }).filter(Boolean);
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
  var optionList = Array.isArray(options) ? options : [];
  if (!optionList.length) {
    container.innerHTML = '<span class="muted">候補なし</span>';
    return;
  }
  var selectedValues = normalizeSelectionSet(selectedValuesRaw);
  var validValues = new Set(optionList.map(function (option) {
    return value(option.value);
  }).filter(Boolean));
  var selectedCount = 0;
  validValues.forEach(function (optionValue) {
    if (selectedValues.has(optionValue)) {
      selectedCount += 1;
    }
  });
  var allSelected = validValues.size > 0 && selectedCount === validValues.size;
  var showHint = selectedCount === 0;
  var allButtonHtml = "\n    <div class=\"export-plan-unit-button-row all-row\">\n      <button type=\"button\" class=\"export-plan-option-button export-plan-option-button-all ".concat(allSelected ? "active" : "", "\" data-value=\"").concat(EXPORT_PLAN_ALL_UNITS_BUTTON_VALUE, "\">\u5168\u30E6\u30CB\u30C3\u30C8</button>\n      ").concat(showHint ? '<span class="export-plan-select-hint">出力するユニットを選んでください</span>' : "", "\n    </div>\n  ");
  var unitButtonHtml = optionList.map(function (option) {
    return "<button type=\"button\" class=\"export-plan-option-button ".concat(selectedValues.has(option.value) ? "active" : "", "\" data-value=\"").concat(escapeHtml(option.value), "\">").concat(escapeHtml(option.label), "</button>");
  }).join("");
  container.innerHTML = "".concat(allButtonHtml, "<div class=\"export-plan-unit-button-row unit-row\">").concat(unitButtonHtml, "</div>");
}
function renderExportPlanModeButtons(container, options, selectedValues) {
  var emptyHintText = arguments.length > 3 && arguments[3] !== undefined ? arguments[3] : "";
  if (!container) {
    return;
  }
  var optionList = Array.isArray(options) ? options : [];
  if (!optionList.length) {
    container.innerHTML = '<span class="muted">候補なし</span>';
    return;
  }
  var selectedSet = normalizeSelectionSet(selectedValues);
  var buttonHtml = optionList.map(function (option) {
    return "<button type=\"button\" class=\"export-plan-option-button ".concat(selectedSet.has(option.value) ? "active" : "", "\" data-value=\"").concat(escapeHtml(option.value), "\">").concat(escapeHtml(option.label), "</button>");
  }).join("");
  var hintHtml = !selectedSet.size && value(emptyHintText) ? "<span class=\"export-plan-select-hint\">".concat(escapeHtml(emptyHintText), "</span>") : "";
  container.innerHTML = "".concat(buttonHtml).concat(hintHtml);
}
function renderExportPlanModeSelect(selectEl, options, selectedValue) {
  if (!selectEl) {
    return;
  }
  var optionList = Array.isArray(options) ? options : [];
  if (!optionList.length) {
    selectEl.innerHTML = "";
    selectEl.disabled = true;
    return;
  }
  selectEl.disabled = false;
  selectEl.innerHTML = optionList.map(function (option) {
    return "<option value=\"".concat(escapeHtml(option.value), "\" ").concat(option.value === selectedValue ? "selected" : "", ">").concat(escapeHtml(option.label), "</option>");
  }).join("");
}
function toggleSelectionInSet(targetSet, optionValueRaw) {
  var checkedForced = arguments.length > 2 && arguments[2] !== undefined ? arguments[2] : null;
  var optionValue = value(optionValueRaw);
  if (!optionValue || !(targetSet instanceof Set)) {
    return;
  }
  if (checkedForced === true) {
    targetSet.add(optionValue);
    return;
  }
  if (checkedForced === false) {
    targetSet["delete"](optionValue);
    return;
  }
  if (targetSet.has(optionValue)) {
    targetSet["delete"](optionValue);
  } else {
    targetSet.add(optionValue);
  }
}
function normalizeSelectionSet(valuesRaw) {
  var next = new Set();
  if (valuesRaw instanceof Set) {
    valuesRaw.forEach(function (item) {
      var normalized = value(item);
      if (normalized) {
        next.add(normalized);
      }
    });
    return next;
  }
  if (Array.isArray(valuesRaw)) {
    valuesRaw.forEach(function (item) {
      var normalized = value(item);
      if (normalized) {
        next.add(normalized);
      }
    });
  }
  return next;
}
function filterPlanRecordsForMode(recordsRaw) {
  var selections = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : {};
  var records = Array.isArray(recordsRaw) ? recordsRaw : [];
  var unitValues = normalizeSelectionSet(selections.unitValues);
  var detailValues = normalizeSelectionSet(selections.detailValues);
  var detailSubValues = normalizeSelectionSet(selections.detailSubValues);
  if (unitValues.size) {
    records = records.filter(function (record) {
      return unitValues.has(unitValueForSelect(record.unit));
    });
  }
  if (detailValues.size) {
    records = records.filter(function (record) {
      return detailValues.has(detailValueForSelect(record.detail));
    });
  }
  if (detailSubValues.size) {
    records = records.filter(function (record) {
      return detailSubValues.has(detailSubValueForSelect(record.detailSub));
    });
  }
  return records;
}
function buildExportPlanModeStats(recordsRaw) {
  var records = Array.isArray(recordsRaw) ? recordsRaw : [];
  var plottedCount = 0;
  records.forEach(function (record) {
    if (buildPlanDrawable(record)) {
      plottedCount += 1;
    }
  });
  return {
    total: records.length,
    missing: Math.max(0, records.length - plottedCount)
  };
}
function renderExportPlanModeStats(targetEl, records) {
  if (!targetEl) {
    return;
  }
  var stats = buildExportPlanModeStats(records);
  var hasData = stats.total > 0;
  targetEl.textContent = "\u5BFE\u8C61\u4EF6\u6570: ".concat(stats.total, "\u4EF6 / \u5E73\u9762\u4F4D\u7F6E\u672A\u8A18\u5165: ").concat(stats.missing, "\u4EF6");
  setAvailabilityClass(targetEl, hasData);
}
function getExportPlanModeSelections() {
  return {
    unit: {
      enabled: exportPlanModeUnitEnabled,
      unitValues: Array.from(exportPlanModeUnitValues)
    },
    detail: {
      enabled: exportPlanModeDetailEnabled,
      unitValue: exportPlanModeDetailUnitValue,
      detailValues: Array.from(exportPlanModeDetailValues)
    },
    detailSub: {
      enabled: exportPlanModeDetailSubEnabled,
      unitValue: exportPlanModeDetailSubUnitValue,
      detailValue: exportPlanModeDetailSubDetailValue,
      detailSubValues: Array.from(exportPlanModeDetailSubValues)
    }
  };
}
function collectExportKuwakuOptionsWithCounts(records) {
  var list = Array.isArray(records) ? records : [];
  if (!list.length) {
    return [{
      value: ALL_GRIDS_VALUE,
      label: "全グリッド（0件）"
    }];
  }
  var countMap = new Map();
  list.forEach(function (record) {
    var key = kuwakuValueForSelect(getRecordKuwaku(record));
    countMap.set(key, (countMap.get(key) || 0) + 1);
  });
  var options = Array.from(countMap.entries()).sort(function (a, b) {
    return kuwakuLabelForSelect(a[0]).localeCompare(kuwakuLabelForSelect(b[0]), "ja", {
      numeric: true,
      sensitivity: "base"
    });
  }).map(function (_ref11) {
    var _ref12 = _slicedToArray(_ref11, 2),
      kuwakuValue = _ref12[0],
      count = _ref12[1];
    return {
      value: kuwakuValue,
      label: "".concat(kuwakuLabelForSelect(kuwakuValue), "\uFF08").concat(count, "\u4EF6\uFF09")
    };
  });
  return [{
    value: ALL_GRIDS_VALUE,
    label: "\u5168\u30B0\u30EA\u30C3\u30C9\uFF08".concat(list.length, "\u4EF6\uFF09")
  }].concat(_toConsumableArray(options));
}
function collectExportCategoryOptions(records) {
  var list = Array.isArray(records) ? records : [];
  var countMap = new Map();
  list.forEach(function (record) {
    var specimen = parseSpecimenNo(record.specimenNo, record.specimenPrefix, record.specimenSerial);
    var prefix = normalizeSpecimenPrefix(specimen.prefix);
    if (!prefix) {
      return;
    }
    countMap.set(prefix, (countMap.get(prefix) || 0) + 1);
  });
  var totalCount = list.length;
  var options = [{
    value: EXPORT_CATEGORY_ALL_VALUE,
    label: "\u5168\u5206\u985E\uFF08".concat(totalCount, "\u4EF6\uFF09")
  }];
  Object.keys(SPECIMEN_CATEGORY_MAP).forEach(function (prefix) {
    var count = countMap.get(prefix) || 0;
    options.push({
      value: prefix,
      label: "".concat(prefix, ": ").concat(SPECIMEN_CATEGORY_MAP[prefix] || "", "\uFF08").concat(count, "\u4EF6\uFF09")
    });
  });
  return options;
}
function filterRecordsByCategory(recordsRaw, categoryValueRaw) {
  var records = Array.isArray(recordsRaw) ? recordsRaw : [];
  var categoryValue = value(categoryValueRaw) || EXPORT_CATEGORY_ALL_VALUE;
  if (!categoryValue || categoryValue === EXPORT_CATEGORY_ALL_VALUE) {
    return records;
  }
  return records.filter(function (record) {
    var specimen = parseSpecimenNo(record === null || record === void 0 ? void 0 : record.specimenNo, record === null || record === void 0 ? void 0 : record.specimenPrefix, record === null || record === void 0 ? void 0 : record.specimenSerial);
    var prefix = normalizeSpecimenPrefix(specimen.prefix);
    return prefix === categoryValue;
  });
}
function syncPlanCategorySelect(recordsRaw) {
  if (!planCategorySelect) {
    return;
  }
  var records = Array.isArray(recordsRaw) ? recordsRaw : [];
  var options = collectExportCategoryOptions(records);
  if (!options.some(function (item) {
    return item.value === selectedPlanCategory;
  })) {
    selectedPlanCategory = EXPORT_CATEGORY_ALL_VALUE;
  }
  planCategorySelect.innerHTML = options.map(function (item) {
    return "<option value=\"".concat(escapeHtml(item.value), "\" ").concat(item.value === selectedPlanCategory ? "selected" : "", ">").concat(escapeHtml(item.label), "</option>");
  }).join("");
}
function syncViewerCategorySelect(recordsRaw) {
  if (!viewerCategorySelect) {
    return;
  }
  var records = Array.isArray(recordsRaw) ? recordsRaw : [];
  var options = collectExportCategoryOptions(records);
  if (!options.some(function (item) {
    return item.value === selectedViewerCategory;
  })) {
    selectedViewerCategory = EXPORT_CATEGORY_ALL_VALUE;
  }
  viewerCategorySelect.innerHTML = options.map(function (item) {
    return "<option value=\"".concat(escapeHtml(item.value), "\" ").concat(item.value === selectedViewerCategory ? "selected" : "", ">").concat(escapeHtml(item.label), "</option>");
  }).join("");
}
function getListExportRecords() {
  return getRecordsByExportRangeFilters({
    kuwakuValue: exportListRangeKuwaku,
    categoryValue: exportListRangeCategory,
    statusValue: exportListRangeStatus,
    specimenFromRaw: exportListRangeSpecimenFrom,
    specimenToRaw: exportListRangeSpecimenTo,
    dateFromRaw: exportListRangeDateFrom,
    dateToRaw: exportListRangeDateTo
  });
}
function getCardExportRecords() {
  return getRecordsByExportRangeFilters({
    kuwakuValue: exportCardRangeKuwaku,
    categoryValue: exportCardRangeCategory,
    statusValue: exportCardRangeStatus,
    dateFromRaw: exportCardRangeDateFrom,
    dateToRaw: exportCardRangeDateTo
  });
}
function getRecordsByExportRangeFilters() {
  var filters = arguments.length > 0 && arguments[0] !== undefined ? arguments[0] : {};
  var records = _toConsumableArray(state.records).sort(compareRecordsByKuwakuThenSpecimen);
  var kuwakuValue = value(filters.kuwakuValue) || ALL_GRIDS_VALUE;
  var categoryValue = value(filters.categoryValue) || EXPORT_CATEGORY_ALL_VALUE;
  var statusValue = value(filters.statusValue) || "all";
  var dateFrom = normalizeDateForExportRange(filters.dateFromRaw);
  var dateTo = normalizeDateForExportRange(filters.dateToRaw);
  var minDate = dateFrom;
  var maxDate = dateTo;
  if (minDate && maxDate && minDate > maxDate) {
    minDate = dateTo;
    maxDate = dateFrom;
  }
  var fromSpecimen = parseSpecimenForExportRange(filters.specimenFromRaw);
  var toSpecimen = parseSpecimenForExportRange(filters.specimenToRaw);
  var minSpecimen = fromSpecimen;
  var maxSpecimen = toSpecimen;
  if (minSpecimen && maxSpecimen && compareRecordsBySpecimenNo(minSpecimen, maxSpecimen) > 0) {
    minSpecimen = toSpecimen;
    maxSpecimen = fromSpecimen;
  }
  return records.filter(function (record) {
    if (kuwakuValue !== ALL_GRIDS_VALUE && kuwakuValueForSelect(getRecordKuwaku(record)) !== kuwakuValue) {
      return false;
    }
    if (categoryValue !== EXPORT_CATEGORY_ALL_VALUE) {
      var specimen = parseSpecimenNo(record.specimenNo, record.specimenPrefix, record.specimenSerial);
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
      var recordDate = normalizeDateForExportRange(getRecordDate(record));
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
  var parsed = parseSpecimenNo(specimenRaw);
  if (!value(parsed.serial)) {
    return null;
  }
  return {
    specimenNo: parsed.specimenNo,
    specimenPrefix: parsed.prefix,
    specimenSerial: parsed.serial,
    id: "__export-range__"
  };
}
function normalizeDateForExportRange(dateRaw) {
  var text = value(dateRaw);
  if (!text) {
    return "";
  }
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) {
    return text;
  }
  var ms = Date.parse(text);
  if (!Number.isFinite(ms)) {
    return "";
  }
  return new Date(ms).toISOString().slice(0, 10);
}
function parseRecordIndex(recordIndexRaw) {
  var index = Number(recordIndexRaw);
  if (!Number.isInteger(index) || index < 0) {
    return -1;
  }
  return index;
}
function buildCellEditRecordSnapshot(record) {
  return {
    id: value(record === null || record === void 0 ? void 0 : record.id),
    specimenNo: parseSpecimenNo(record === null || record === void 0 ? void 0 : record.specimenNo, record === null || record === void 0 ? void 0 : record.specimenPrefix, record === null || record === void 0 ? void 0 : record.specimenSerial).specimenNo,
    kuwaku: normalizeKuwakuText(getRecordKuwaku(record)),
    createdAt: value(record === null || record === void 0 ? void 0 : record.createdAt)
  };
}
function matchesCellEditRecordSnapshot(record, snapshotRaw) {
  var snapshot = snapshotRaw && _typeof(snapshotRaw) === "object" ? snapshotRaw : {};
  if (!record) {
    return false;
  }
  if (value(snapshot.id) && value(record.id) !== value(snapshot.id)) {
    return false;
  }
  if (value(snapshot.specimenNo)) {
    var specimenNo = parseSpecimenNo(record.specimenNo, record.specimenPrefix, record.specimenSerial).specimenNo;
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
function findRecordByEditContext(recordIdRaw) {
  var recordIndexRaw = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : "";
  var snapshotRaw = arguments.length > 2 && arguments[2] !== undefined ? arguments[2] : null;
  var recordId = value(recordIdRaw);
  var recordIndex = parseRecordIndex(recordIndexRaw);
  var snapshot = snapshotRaw && _typeof(snapshotRaw) === "object" ? snapshotRaw : null;
  if (recordIndex >= 0 && recordIndex < state.records.length) {
    var byIndex = state.records[recordIndex];
    if (byIndex && (!recordId || value(byIndex.id) === recordId) && matchesCellEditRecordSnapshot(byIndex, snapshot)) {
      return byIndex;
    }
  }
  if (snapshot) {
    var exactMatches = state.records.filter(function (item) {
      return matchesCellEditRecordSnapshot(item, snapshot);
    });
    if (exactMatches.length === 1) {
      return exactMatches[0];
    }
  }
  if (recordId) {
    var idMatches = state.records.filter(function (item) {
      return value(item === null || item === void 0 ? void 0 : item.id) === recordId;
    });
    if (idMatches.length === 1) {
      return idMatches[0];
    }
    if (snapshot) {
      var idAndSnapshotMatches = idMatches.filter(function (item) {
        return matchesCellEditRecordSnapshot(item, snapshot);
      });
      if (idAndSnapshotMatches.length === 1) {
        return idAndSnapshotMatches[0];
      }
    }
  }
  return null;
}
function openOutputCellEditModal(recordId, editKey) {
  var recordIndexRaw = arguments.length > 2 && arguments[2] !== undefined ? arguments[2] : "";
  if (!cellEditModal || !cellEditTitle || !cellEditMeta || !cellEditFields) {
    return;
  }
  var record = findRecordByEditContext(recordId, recordIndexRaw, null);
  if (!record) {
    showToast("対象データが見つかりません");
    return;
  }
  var label = OUTPUT_CELL_EDIT_LABELS[editKey];
  if (!label) {
    return;
  }
  var fieldsHtml = buildOutputCellEditFieldsHtml(record, editKey);
  if (!fieldsHtml) {
    showToast("このセルは直接編集に対応していません");
    return;
  }
  activeOutputCellEdit = {
    recordId: value(record.id),
    editKey: value(editKey),
    recordIndex: String(parseRecordIndex(recordIndexRaw)),
    recordSnapshot: buildCellEditRecordSnapshot(record)
  };
  cellEditTitle.textContent = "".concat(label, "\u3092\u7DE8\u96C6");
  cellEditMeta.textContent = "\u6A19\u672C\u756A\u53F7 ".concat(record.specimenNo || "-", " / \u533A\u753B ").concat(getRecordKuwaku(record) || "-");
  cellEditFields.innerHTML = fieldsHtml;
  cellEditModal.classList.remove("hidden");
  bindOutputCellEditDynamicFields(editKey);
  var focusTarget = cellEditFields.querySelector("input, select, textarea");
  if (focusTarget instanceof HTMLElement) {
    focusTarget.focus();
    if (focusTarget instanceof HTMLInputElement || focusTarget instanceof HTMLTextAreaElement) {
      var _focusTarget$select;
      (_focusTarget$select = focusTarget.select) === null || _focusTarget$select === void 0 || _focusTarget$select.call(focusTarget);
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
  var teamState = normalizeTeamState(record === null || record === void 0 ? void 0 : record.team, record === null || record === void 0 ? void 0 : record.teamOther);
  var teamOtherHiddenClass = teamState.team === OTHER_TEAM_NAME ? "" : " hidden";
  var parsedSpecimen = parseSpecimenNo(record === null || record === void 0 ? void 0 : record.specimenNo, record === null || record === void 0 ? void 0 : record.specimenPrefix, record === null || record === void 0 ? void 0 : record.specimenSerial);
  var specimenPrefix = normalizeSpecimenPrefix(parsedSpecimen.prefix);
  var analysisHiddenClass = specimenPrefix === "a" ? "" : " hidden";
  var analysisType = normalizeAnalysisType(record === null || record === void 0 ? void 0 : record.analysisType);
  var altitudeDirectEnabled = normalizeToggleFlag(record === null || record === void 0 ? void 0 : record.altitudeInputEnabled) === "1";
  var altitudeDirectHiddenClass = altitudeDirectEnabled ? "" : " hidden";
  switch (editKey) {
    case "kuwaku":
      return "\n        <label>\n          <span class=\"label-title\">\u533A\u753B\uFF08\u30B0\u30EA\u30C3\u30C9\uFF09</span>\n          <input name=\"kuwaku\" type=\"text\" value=\"".concat(escapeHtml(getRecordKuwaku(record)), "\" placeholder=\"\u4F8B: 24-\u2160-C-4\" />\n        </label>\n      ");
    case "team":
      return "\n        <label>\n          <span class=\"label-title\">\u767A\u6398\u73ED</span>\n          <select name=\"team\" data-cell-edit-team-select>\n            <option value=\"\" ".concat(teamState.team === "" ? "selected" : "", ">\u672A\u8A2D\u5B9A</option>\n            <option value=\"1\" ").concat(teamState.team === "1" ? "selected" : "", ">1</option>\n            <option value=\"2\" ").concat(teamState.team === "2" ? "selected" : "", ">2</option>\n            <option value=\"3\" ").concat(teamState.team === "3" ? "selected" : "", ">3</option>\n            <option value=\"4\" ").concat(teamState.team === "4" ? "selected" : "", ">4</option>\n            <option value=\"\u305D\u306E\u4ED6\" ").concat(teamState.team === "その他" ? "selected" : "", ">\u305D\u306E\u4ED6</option>\n          </select>\n        </label>\n        <label data-cell-edit-team-other class=\"").concat(teamOtherHiddenClass, "\">\n          <span class=\"label-title\">\u767A\u6398\u73ED\uFF08\u305D\u306E\u4ED6\uFF09</span>\n          <input name=\"teamOther\" type=\"text\" value=\"").concat(escapeHtml(teamState.teamOther), "\" />\n        </label>\n      ");
    case "date":
      return "\n        <label>\n          <span class=\"label-title\">\u65E5\u4ED8</span>\n          <input name=\"date\" type=\"date\" value=\"".concat(escapeHtml(normalizeDateForExportRange(getRecordDate(record))), "\" />\n        </label>\n      ");
    case "specimenNo":
      return "\n        <label>\n          <span class=\"label-title\">\u6A19\u672C\u756A\u53F7</span>\n          <input name=\"specimenNo\" type=\"text\" value=\"".concat(escapeHtml(parsedSpecimen.specimenNo), "\" placeholder=\"\u4F8B: m-1\" />\n        </label>\n      ");
    case "category":
      return "\n        <label>\n          <span class=\"label-title\">\u5206\u985E</span>\n          <select name=\"specimenPrefix\" data-cell-edit-prefix-select>\n            ".concat(buildOutputCellEditPrefixOptionsHtml(specimenPrefix), "\n          </select>\n        </label>\n        <label>\n          <span class=\"label-title\">\u6A19\u672C\u756A\u53F7</span>\n          <input name=\"specimenNo\" type=\"text\" value=\"").concat(escapeHtml(parsedSpecimen.specimenNo), "\" placeholder=\"\u4F8B: ").concat(specimenPrefix, "-1\" data-cell-edit-specimen-no />\n        </label>\n        <label data-cell-edit-analysis-row class=\"").concat(analysisHiddenClass, "\">\n          <span class=\"label-title\">\u5206\u6790\u7528\u8A66\u6599\u306E\u533A\u5206</span>\n          <select name=\"analysisType\">\n            ").concat(buildOutputCellEditAnalysisOptionsHtml(analysisType), "\n          </select>\n        </label>\n      ");
    case "nameMemo":
      return "\n        <label>\n          <span class=\"label-title\">\u5316\u77F3\u30FB\u907A\u7269\u540D\u79F0</span>\n          <input name=\"nameMemo\" type=\"text\" value=\"".concat(escapeHtml(value(record === null || record === void 0 ? void 0 : record.nameMemo)), "\" />\n        </label>\n      ");
    case "importantFlag":
      return "\n        <label>\n          <span class=\"label-title\">\u91CD\u8981\u54C1\u6307\u5B9A</span>\n          <select name=\"importantFlag\">\n            <option value=\"\u7121\" ".concat(value(record === null || record === void 0 ? void 0 : record.importantFlag) === "無" ? "selected" : "", ">\u7121</option>\n            <option value=\"\u6709\" ").concat(value(record === null || record === void 0 ? void 0 : record.importantFlag) === "有" ? "selected" : "", ">\u6709</option>\n          </select>\n        </label>\n      ");
    case "unit":
      return "\n        <label>\n          <span class=\"label-title\">\u30E6\u30CB\u30C3\u30C8</span>\n          <input name=\"unit\" type=\"text\" value=\"".concat(escapeHtml(value(record === null || record === void 0 ? void 0 : record.unit)), "\" />\n        </label>\n      ");
    case "detail":
      return "\n        <label>\n          <span class=\"label-title\">\u30B5\u30D6\u30E6\u30CB\u30C3\u30C8</span>\n          <input name=\"detail\" type=\"text\" value=\"".concat(escapeHtml(value(record === null || record === void 0 ? void 0 : record.detail)), "\" />\n        </label>\n        <label>\n          <span class=\"label-title\">\u7D30\u5206</span>\n          <input name=\"detailSub\" type=\"text\" value=\"").concat(escapeHtml(value(record === null || record === void 0 ? void 0 : record.detailSub)), "\" />\n        </label>\n      ");
    case "discoverer":
      return "\n        <label>\n          <span class=\"label-title\">\u767A\u898B\u8005\u6C0F\u540D</span>\n          <input name=\"discoverer\" type=\"text\" value=\"".concat(escapeHtml(value(record === null || record === void 0 ? void 0 : record.discoverer)), "\" />\n        </label>\n      ");
    case "identifier":
      return "\n        <label>\n          <span class=\"label-title\">\u5224\u5B9A\u8005\u6C0F\u540D</span>\n          <input name=\"identifier\" type=\"text\" value=\"".concat(escapeHtml(value(record === null || record === void 0 ? void 0 : record.identifier)), "\" />\n        </label>\n      ");
    case "levelRead":
      return "\n        <label>\n          <span class=\"label-title\">\u30EC\u30D9\u30EB\u8AAD\u5024\uFF08\u4E0A\u9762\uFF09</span>\n          <div class=\"inline-cm-row compact-cm-row\">\n            <input name=\"levelUpperCm\" type=\"text\" value=\"".concat(escapeHtml(value(record === null || record === void 0 ? void 0 : record.levelUpperCm)), "\" />\n            <span>cm</span>\n          </div>\n        </label>\n        <label>\n          <span class=\"label-title\">\u30EC\u30D9\u30EB\u8AAD\u5024\uFF08\u4E0B\u5E95\uFF09</span>\n          <div class=\"inline-cm-row compact-cm-row\">\n            <input name=\"levelLowerCm\" type=\"text\" value=\"").concat(escapeHtml(value(record === null || record === void 0 ? void 0 : record.levelLowerCm)), "\" />\n            <span>cm</span>\n          </div>\n        </label>\n      ");
    case "altitudeM":
      return "\n        <label>\n          <span class=\"label-title\">\u30EC\u30D9\u30EB\u9AD8</span>\n          <div class=\"inline-cm-row compact-cm-row\">\n            <input name=\"levelHeight\" type=\"text\" value=\"".concat(escapeHtml(value(record === null || record === void 0 ? void 0 : record.levelHeight)), "\" />\n            <span>m</span>\n          </div>\n        </label>\n        <label class=\"level-altitude-toggle\">\n          <input name=\"altitudeInputEnabled\" type=\"checkbox\" value=\"1\" ").concat(altitudeDirectEnabled ? "checked" : "", " data-cell-edit-altitude-check />\n          \u6A19\u9AD8\uFF08\u4E0B\u5E95\uFF09\u3092\u76F4\u63A5\u5165\u529B\n        </label>\n        <label data-cell-edit-altitude-direct class=\"").concat(altitudeDirectHiddenClass, "\">\n          <span class=\"label-title\">\u6A19\u9AD8\uFF08\u4E0B\u5E95\uFF09</span>\n          <div class=\"inline-cm-row compact-cm-row\">\n            <input name=\"altitudeDirectM\" type=\"text\" value=\"").concat(escapeHtml(value(record === null || record === void 0 ? void 0 : record.altitudeDirectM)), "\" />\n            <span>m</span>\n          </div>\n        </label>\n      ");
    case "occurrenceSection":
      return "\n        <label>\n          <span class=\"label-title\">\u7523\u51FA\u72B6\u6CC1\u65AD\u9762</span>\n          <select name=\"occurrenceSection\">\n            <option value=\"\u8981\" ".concat(value(record === null || record === void 0 ? void 0 : record.occurrenceSection) === "否" ? "" : "selected", ">\u8981</option>\n            <option value=\"\u5426\" ").concat(value(record === null || record === void 0 ? void 0 : record.occurrenceSection) === "否" ? "selected" : "", ">\u5426</option>\n          </select>\n        </label>\n      ");
    case "occurrenceSketch":
      return "\n        <label>\n          <span class=\"label-title\">\u7523\u72B6\u30B9\u30B1\u30C3\u30C1</span>\n          <select name=\"occurrenceSketch\">\n            <option value=\"\u8981\" ".concat(value(record === null || record === void 0 ? void 0 : record.occurrenceSketch) === "否" ? "" : "selected", ">\u8981</option>\n            <option value=\"\u5426\" ").concat(value(record === null || record === void 0 ? void 0 : record.occurrenceSketch) === "否" ? "selected" : "", ">\u5426</option>\n          </select>\n        </label>\n      ");
    case "position":
      return "\n        <label>\n          <span class=\"label-title\">\u5317\u304B\u3089 / \u5357\u304B\u3089</span>\n          <div class=\"inline-cm-row compact-cm-row\">\n            <select name=\"nsDir\">\n              <option value=\"\u5317\u304B\u3089\" ".concat(normalizeNsDir(record === null || record === void 0 ? void 0 : record.nsDir) === "北から" ? "selected" : "", ">\u5317\u304B\u3089</option>\n              <option value=\"\u5357\u304B\u3089\" ").concat(normalizeNsDir(record === null || record === void 0 ? void 0 : record.nsDir) === "南から" ? "selected" : "", ">\u5357\u304B\u3089</option>\n            </select>\n            <input name=\"nsCm\" type=\"text\" value=\"").concat(escapeHtml(value(record === null || record === void 0 ? void 0 : record.nsCm)), "\" />\n            <span>cm</span>\n          </div>\n        </label>\n        <label>\n          <span class=\"label-title\">\u6771\u304B\u3089 / \u897F\u304B\u3089</span>\n          <div class=\"inline-cm-row compact-cm-row\">\n            <select name=\"ewDir\">\n              <option value=\"\u6771\u304B\u3089\" ").concat(normalizeEwDir(record === null || record === void 0 ? void 0 : record.ewDir) === "東から" ? "selected" : "", ">\u6771\u304B\u3089</option>\n              <option value=\"\u897F\u304B\u3089\" ").concat(normalizeEwDir(record === null || record === void 0 ? void 0 : record.ewDir) === "西から" ? "selected" : "", ">\u897F\u304B\u3089</option>\n            </select>\n            <input name=\"ewCm\" type=\"text\" value=\"").concat(escapeHtml(value(record === null || record === void 0 ? void 0 : record.ewCm)), "\" />\n            <span>cm</span>\n          </div>\n        </label>\n      ");
    case "notes":
      return "\n        <label>\n          <span class=\"label-title\">\u5099\u8003</span>\n          <textarea name=\"notes\" rows=\"4\">".concat(escapeHtml(value(record === null || record === void 0 ? void 0 : record.notes)), "</textarea>\n        </label>\n      ");
    default:
      return "";
  }
}
function bindOutputCellEditDynamicFields(editKey) {
  if (!cellEditFields) {
    return;
  }
  if (editKey === "team") {
    var teamSelect = cellEditFields.querySelector("[data-cell-edit-team-select]");
    var teamOtherLabel = cellEditFields.querySelector("[data-cell-edit-team-other]");
    if (teamSelect instanceof HTMLSelectElement && teamOtherLabel instanceof HTMLElement) {
      var toggle = function toggle() {
        var isOther = value(teamSelect.value) === OTHER_TEAM_NAME;
        teamOtherLabel.classList.toggle("hidden", !isOther);
      };
      teamSelect.addEventListener("change", toggle);
      toggle();
    }
  }
  if (editKey === "category") {
    var prefixSelect = cellEditFields.querySelector("[data-cell-edit-prefix-select]");
    var analysisRow = cellEditFields.querySelector("[data-cell-edit-analysis-row]");
    var specimenNoInput = cellEditFields.querySelector("[data-cell-edit-specimen-no]");
    if (prefixSelect instanceof HTMLSelectElement && analysisRow instanceof HTMLElement) {
      var _toggle = function _toggle() {
        var prefix = normalizeSpecimenPrefix(prefixSelect.value);
        analysisRow.classList.toggle("hidden", prefix !== "a");
        if (specimenNoInput instanceof HTMLInputElement) {
          var parsed = parseSpecimenNo(specimenNoInput.value, prefix, "");
          var serial = compactNoSpaceValue(parsed.serial);
          if (!serial || normalizeSpecimenPrefix(parsed.prefix) !== prefix) {
            serial = findSmallestUnusedSpecimenSerialForActiveEdit(prefix);
          }
          specimenNoInput.value = buildSpecimenNo(prefix, serial);
          specimenNoInput.placeholder = "\u4F8B: ".concat(prefix, "-").concat(serial || "1");
        }
      };
      prefixSelect.addEventListener("change", _toggle);
      _toggle();
    }
  }
  if (editKey === "altitudeM") {
    var check = cellEditFields.querySelector("[data-cell-edit-altitude-check]");
    var directRow = cellEditFields.querySelector("[data-cell-edit-altitude-direct]");
    var directInput = directRow === null || directRow === void 0 ? void 0 : directRow.querySelector("input[name='altitudeDirectM']");
    if (check instanceof HTMLInputElement && directRow instanceof HTMLElement) {
      var _toggle2 = function _toggle2() {
        var enabled = check.checked;
        directRow.classList.toggle("hidden", !enabled);
        if (directInput instanceof HTMLInputElement) {
          directInput.disabled = !enabled;
          if (!enabled) {
            directInput.value = "";
          }
        }
      };
      check.addEventListener("change", _toggle2);
      _toggle2();
    }
  }
}
function saveOutputCellEditFromModal() {
  if (!activeOutputCellEdit || !cellEditForm) {
    return;
  }
  var record = findRecordByEditContext(activeOutputCellEdit.recordId, activeOutputCellEdit.recordIndex, activeOutputCellEdit.recordSnapshot);
  if (!record) {
    closeOutputCellEditModal();
    showToast("対象データが見つかりません（一覧を再表示してから再度編集してください）");
    return;
  }
  var formData = createOutputCellEditFormData();
  var result = applyOutputCellEditToRecord(record, activeOutputCellEdit.editKey, formData);
  if (!result.ok) {
    if (result.message) {
      showToast(result.message);
    }
    return;
  }
  record.updatedAt = nowIso();
  var specimenNoForMessage = parseSpecimenNo(record.specimenNo, record.specimenPrefix, record.specimenSerial).specimenNo || record.specimenNo;
  persist("\u6A19\u672C\u756A\u53F7 ".concat(specimenNoForMessage || "-", " \u3092\u4E0A\u66F8\u304D\u4FDD\u5B58\u3057\u307E\u3057\u305F"));
  renderRecordTable();
  renderOutputs();
  closeOutputCellEditModal();
}
function createOutputCellEditFormData() {
  var values = {};
  if (cellEditFields) {
    var fields = cellEditFields.querySelectorAll("input[name], select[name], textarea[name]");
    fields.forEach(function (field) {
      if (!(field instanceof HTMLInputElement || field instanceof HTMLSelectElement || field instanceof HTMLTextAreaElement) || field.disabled) {
        return;
      }
      var name = value(field.name);
      if (!name) {
        return;
      }
      if (field instanceof HTMLInputElement && (field.type === "checkbox" || field.type === "radio")) {
        if (field.checked) {
          values[name] = field.value;
        }
        return;
      }
      values[name] = field.value;
    });
  }
  return {
    get: function get(name) {
      return Object.prototype.hasOwnProperty.call(values, name) ? values[name] : null;
    }
  };
}
function applyOutputCellEditToRecord(record, editKey, formData) {
  if (!(record && formData && typeof formData.get === "function")) {
    return {
      ok: false,
      message: "編集データを取得できませんでした"
    };
  }
  if (editKey === "kuwaku") {
    var parsed = parseKuwaku(formData.get("kuwaku"));
    var nextKuwaku = buildKuwaku(parsed.headA, parsed.headB, parsed.block, parsed.no);
    if (!parsed.block || !parsed.no || !nextKuwaku) {
      return {
        ok: false,
        message: "区画（例: 24-Ⅰ-C-4）を入力してください"
      };
    }
    var duplicate = findDuplicateRecordByKuwakuAndSpecimen(nextKuwaku, record.specimenNo, record.id);
    if (duplicate) {
      return {
        ok: false,
        message: "\u3053\u306E\u533A\u753B\u306B\u306F ".concat(record.specimenNo, " \u304C\u3059\u3067\u306B\u3042\u308A\u307E\u3059")
      };
    }
    record.kuwaku = nextKuwaku;
    return {
      ok: true
    };
  }
  if (editKey === "team") {
    var teamState = normalizeTeamState(formData.get("team"), formData.get("teamOther"));
    record.team = teamState.team;
    record.teamOther = teamState.teamOther;
    return {
      ok: true
    };
  }
  if (editKey === "date") {
    record.date = value(formData.get("date"));
    return {
      ok: true
    };
  }
  if (editKey === "specimenNo") {
    var _parsed = parseSpecimenNo(formData.get("specimenNo"), record.specimenPrefix, record.specimenSerial);
    if (!_parsed.serial) {
      return {
        ok: false,
        message: "標本番号を入力してください（例: m-1）"
      };
    }
    var nextSpecimenNo = buildSpecimenNo(_parsed.prefix, _parsed.serial);
    var _duplicate = findDuplicateRecordByKuwakuAndSpecimen(getRecordKuwaku(record), nextSpecimenNo, record.id);
    if (_duplicate) {
      return {
        ok: false,
        message: "\u3053\u306E\u533A\u753B\u306B\u306F ".concat(nextSpecimenNo, " \u304C\u3059\u3067\u306B\u3042\u308A\u307E\u3059")
      };
    }
    record.specimenPrefix = normalizeSpecimenPrefix(_parsed.prefix);
    record.specimenSerial = compactNoSpaceValue(_parsed.serial);
    record.specimenNo = nextSpecimenNo;
    record.category = categoryFromPrefix(record.specimenPrefix);
    if (record.specimenPrefix !== "a") {
      record.analysisType = "";
    }
    return {
      ok: true
    };
  }
  if (editKey === "category") {
    var nextPrefix = normalizeSpecimenPrefix(formData.get("specimenPrefix"));
    var currentSpecimen = parseSpecimenNo(formData.get("specimenNo"), nextPrefix, "");
    var nextSerial = compactNoSpaceValue(currentSpecimen.serial) || findSmallestUnusedSpecimenSerial(getRecordKuwaku(record), nextPrefix, record.id);
    if (!nextSerial) {
      record.specimenPrefix = nextPrefix;
      record.specimenSerial = "";
      record.specimenNo = "";
      record.category = categoryFromPrefix(nextPrefix);
      record.analysisType = nextPrefix === "a" ? normalizeAnalysisType(formData.get("analysisType")) : "";
      return {
        ok: true
      };
    }
    var _nextSpecimenNo = buildSpecimenNo(nextPrefix, nextSerial);
    var _duplicate2 = findDuplicateRecordByKuwakuAndSpecimen(getRecordKuwaku(record), _nextSpecimenNo, record.id);
    if (_duplicate2) {
      return {
        ok: false,
        message: "\u3053\u306E\u533A\u753B\u306B\u306F ".concat(_nextSpecimenNo, " \u304C\u3059\u3067\u306B\u3042\u308A\u307E\u3059")
      };
    }
    record.specimenPrefix = nextPrefix;
    record.specimenSerial = nextSerial;
    record.specimenNo = _nextSpecimenNo;
    record.category = categoryFromPrefix(nextPrefix);
    record.analysisType = nextPrefix === "a" ? normalizeAnalysisType(formData.get("analysisType")) : "";
    return {
      ok: true
    };
  }
  if (editKey === "nameMemo") {
    record.nameMemo = value(formData.get("nameMemo"));
    return {
      ok: true
    };
  }
  if (editKey === "importantFlag") {
    record.importantFlag = normalizeHasFlag(formData.get("importantFlag")) || "無";
    return {
      ok: true
    };
  }
  if (editKey === "unit") {
    record.unit = compactNoSpaceValue(formData.get("unit"));
    return {
      ok: true
    };
  }
  if (editKey === "detail") {
    record.detail = compactNoSpaceValue(formData.get("detail"));
    record.detailSub = value(formData.get("detailSub"));
    return {
      ok: true
    };
  }
  if (editKey === "discoverer") {
    record.discoverer = value(formData.get("discoverer"));
    return {
      ok: true
    };
  }
  if (editKey === "identifier") {
    record.identifier = value(formData.get("identifier"));
    return {
      ok: true
    };
  }
  if (editKey === "levelRead") {
    record.levelUpperCm = value(formData.get("levelUpperCm"));
    record.levelLowerCm = value(formData.get("levelLowerCm"));
    return {
      ok: true
    };
  }
  if (editKey === "altitudeM") {
    record.levelHeight = value(formData.get("levelHeight"));
    var useDirectAltitude = normalizeToggleFlag(formData.get("altitudeInputEnabled")) === "1";
    record.altitudeInputEnabled = useDirectAltitude ? "1" : "";
    record.altitudeDirectM = useDirectAltitude ? value(formData.get("altitudeDirectM")) : "";
    return {
      ok: true
    };
  }
  if (editKey === "occurrenceSection") {
    record.occurrenceSection = normalizeNeedFlag(formData.get("occurrenceSection"));
    return {
      ok: true
    };
  }
  if (editKey === "occurrenceSketch") {
    record.occurrenceSketch = normalizeNeedFlag(formData.get("occurrenceSketch"));
    return {
      ok: true
    };
  }
  if (editKey === "position") {
    record.nsDir = normalizeNsDir(formData.get("nsDir"));
    record.nsCm = value(formData.get("nsCm"));
    record.ewDir = normalizeEwDir(formData.get("ewDir"));
    record.ewCm = value(formData.get("ewCm"));
    return {
      ok: true
    };
  }
  if (editKey === "notes") {
    record.notes = value(formData.get("notes"));
    return {
      ok: true
    };
  }
  return {
    ok: false,
    message: "このセルは直接編集に対応していません"
  };
}
function findSmallestUnusedSpecimenSerialForActiveEdit(prefixRaw) {
  if (!activeOutputCellEdit) {
    return "1";
  }
  var record = findRecordByEditContext(activeOutputCellEdit.recordId, activeOutputCellEdit.recordIndex, activeOutputCellEdit.recordSnapshot);
  if (!record) {
    return "1";
  }
  return findSmallestUnusedSpecimenSerial(getRecordKuwaku(record), prefixRaw, record.id);
}
function findSmallestUnusedSpecimenSerial(kuwakuRaw, prefixRaw, excludeRecordIdRaw) {
  var kuwaku = normalizeKuwakuText(kuwakuRaw);
  var prefix = normalizeSpecimenPrefix(prefixRaw);
  var excludeRecordId = value(excludeRecordIdRaw);
  var used = new Set();
  state.records.forEach(function (item) {
    if (!item || value(item.id) === excludeRecordId || normalizeKuwakuText(getRecordKuwaku(item)) !== kuwaku) {
      return;
    }
    var specimen = parseSpecimenNo(item.specimenNo, item.specimenPrefix, item.specimenSerial);
    if (normalizeSpecimenPrefix(specimen.prefix) !== prefix) {
      return;
    }
    var serial = compactNoSpaceValue(specimen.serial);
    if (/^\d+$/.test(serial)) {
      used.add(Number(serial));
    }
  });
  var next = 1;
  while (used.has(next)) {
    next += 1;
  }
  return String(next);
}
function buildOutputCellEditPrefixOptionsHtml(selectedPrefixRaw) {
  var selectedPrefix = normalizeSpecimenPrefix(selectedPrefixRaw);
  var order = ["m", "b", "l", "s", "i", "g", "h", "a"];
  return order.map(function (prefix) {
    var label = "".concat(prefix, ": ").concat(SPECIMEN_CATEGORY_MAP[prefix] || "");
    return "<option value=\"".concat(prefix, "\" ").concat(prefix === selectedPrefix ? "selected" : "", ">").concat(escapeHtml(label), "</option>");
  }).join("");
}
function buildOutputCellEditAnalysisOptionsHtml(selectedTypeRaw) {
  var selectedType = normalizeAnalysisType(selectedTypeRaw);
  var order = ["A", "C", "M", "F", "P", "B", "I", "D", "R", "S", "H", "MG"];
  var options = ["<option value=\"\" ".concat(selectedType ? "" : "selected", ">\u672A\u8A2D\u5B9A</option>")];
  order.forEach(function (code) {
    var displayCode = code === "MG" ? "Mg" : code;
    var optionValue = "".concat(displayCode, ": ").concat(ANALYSIS_TYPE_MAP[code]);
    options.push("<option value=\"".concat(escapeHtml(optionValue), "\" ").concat(optionValue === selectedType ? "selected" : "", ">").concat(escapeHtml(optionValue), "</option>"));
  });
  return options.join("");
}
function isOutputListColumnVisible(columnKeyRaw) {
  var columnKey = value(columnKeyRaw);
  if (!columnKey) {
    return true;
  }
  return outputListColumnVisibility[columnKey] !== false;
}
function getOutputListVisibleColumnCount() {
  var visibleCount = OUTPUT_LIST_COLUMN_DEFS.reduce(function (count, column) {
    return count + (isOutputListColumnVisible(column.key) ? 1 : 0);
  }, 0);
  return visibleCount > 0 ? visibleCount : 1;
}
function renderOutputColumnToggleRow() {
  if (!outputColumnToggleRow) {
    return;
  }
  outputColumnToggleRow.innerHTML = OUTPUT_LIST_COLUMN_DEFS.map(function (column) {
    var visible = isOutputListColumnVisible(column.key);
    return "<button type=\"button\" class=\"output-column-toggle-btn ".concat(visible ? "active" : "", "\" data-col-key=\"").concat(escapeHtml(column.key), "\" aria-pressed=\"").concat(visible ? "true" : "false", "\">").concat(escapeHtml(column.label), "</button>");
  }).join("");
}
function applyOutputListColumnVisibility() {
  if (!outputListTable) {
    return;
  }
  var allCells = outputListTable.querySelectorAll("[data-col-key]");
  allCells.forEach(function (cell) {
    var columnKey = value(cell.dataset.colKey);
    if (!columnKey) {
      return;
    }
    cell.hidden = !isOutputListColumnVisible(columnKey);
  });
}
function toggleOutputListColumnVisibility(columnKeyRaw) {
  var columnKey = value(columnKeyRaw);
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
    outputListBody.innerHTML = "<tr><td colspan=\"".concat(getOutputListVisibleColumnCount(), "\">\u51FA\u529B\u5BFE\u8C61\u30C7\u30FC\u30BF\u304C\u3042\u308A\u307E\u305B\u3093\u3002</td></tr>");
    applyOutputListColumnVisibility();
    return;
  }
  var filteredRecords = getFilteredOutputRecords();
  if (!filteredRecords.length) {
    outputListBody.innerHTML = "<tr><td colspan=\"".concat(getOutputListVisibleColumnCount(), "\">\u6761\u4EF6\u306B\u4E00\u81F4\u3059\u308B\u30C7\u30FC\u30BF\u304C\u3042\u308A\u307E\u305B\u3093\u3002</td></tr>");
    applyOutputListColumnVisibility();
    return;
  }
  var recordIndexMap = new Map();
  state.records.forEach(function (record, index) {
    if (record && _typeof(record) === "object") {
      recordIndexMap.set(record, index);
    }
  });
  outputListBody.innerHTML = sortOutputRecordsForList(filteredRecords).map(function (record) {
    var kuwakuText = getRecordKuwaku(record);
    var kuwakuStyle = getKuwakuCellStyle(kuwakuText);
    var categoryColor = getRecordSpecimenColor(record);
    var categoryBackground = toRgbaColor(categoryColor, 0.2);
    var categoryBorderColor = toRgbaColor(categoryColor, 0.45);
    var unitStyle = getUnitCellStyle(record.unit);
    var missingRequiredKeys = getMissingRequiredKeys(record);
    var dataComplete = missingRequiredKeys.size === 0;
    var missingTitle = formatMissingRequiredTooltip(missingRequiredKeys);
    var kuwakuMissing = hasAnyMissingRequiredKey(missingRequiredKeys, ["kuwakuHeadA", "kuwakuHeadB", "kuwakuBlock", "kuwakuNo"]);
    var teamMissing = hasAnyMissingRequiredKey(missingRequiredKeys, ["team", "teamOther"]);
    var specimenMissing = hasAnyMissingRequiredKey(missingRequiredKeys, ["specimenSerial"]);
    var categoryMissing = hasAnyMissingRequiredKey(missingRequiredKeys, ["analysisType"]);
    var nameMissing = hasAnyMissingRequiredKey(missingRequiredKeys, ["nameMemo"]);
    var importantMissing = hasAnyMissingRequiredKey(missingRequiredKeys, ["importantFlag"]);
    var unitMissing = hasAnyMissingRequiredKey(missingRequiredKeys, ["unit"]);
    var discovererMissing = hasAnyMissingRequiredKey(missingRequiredKeys, ["discoverer"]);
    var identifierMissing = hasAnyMissingRequiredKey(missingRequiredKeys, ["identifier"]);
    var levelHeightMissing = hasAnyMissingRequiredKey(missingRequiredKeys, ["levelHeight"]);
    var levelReadMissing = hasAnyMissingRequiredKey(missingRequiredKeys, ["levelUpperCm", "levelLowerCm"]);
    var sectionMissing = hasAnyMissingRequiredKey(missingRequiredKeys, ["occurrenceSection", "sectionDiagrams"]);
    var sectionChecklistMissing = hasAnyMissingRequiredKey(missingRequiredKeys, ["sectionDiagramDistanceChecked", "sectionDiagramHorizonChecked", "sectionDiagramLayerFaciesChecked"]);
    var sketchMissing = hasAnyMissingRequiredKey(missingRequiredKeys, ["occurrenceSketch"]);
    var positionMissing = hasAnyMissingRequiredKey(missingRequiredKeys, ["nsDir", "nsCm", "ewDir", "ewCm", "largeShapeType", "largeAxisDirection", "lineLengthCm", "rectSide1Cm", "rectSide2Cm", "ellipseLongRadiusCm", "ellipseShortRadiusCm", "imgRotateDeg", "imgFrameWidthCm", "imgFrameHeightCm", "customLargeImageName", "customLargeImageDataUrl"]);
    var recordIndex = Number(recordIndexMap.get(record));
    return "\n      <tr data-record-id=\"".concat(escapeHtml(value(record.id)), "\" data-record-index=\"").concat(Number.isInteger(recordIndex) && recordIndex >= 0 ? recordIndex : "", "\">\n        <td data-col-key=\"kuwaku\" data-cell-edit-key=\"kuwaku\" class=\"").concat(listCellClass("kuwaku-color-cell", kuwakuMissing), "\" style=\"background:").concat(kuwakuStyle.background, ";color:").concat(kuwakuStyle.color, ";border-color:").concat(kuwakuStyle.border, ";\" ").concat(missingTitle, ">").concat(escapeHtml(kuwakuText), "</td>\n        <td data-col-key=\"team\" data-cell-edit-key=\"team\" class=\"").concat(listCellClass("", teamMissing), "\" ").concat(missingTitle, ">").concat(escapeHtml(getRecordTeamValue(record)), "</td>\n        <td data-col-key=\"date\" data-cell-edit-key=\"date\">").concat(escapeHtml(getRecordDate(record)), "</td>\n        <td data-col-key=\"dataStatus\" class=\"").concat(listCellClass(dataComplete ? "data-status-complete" : "data-status-incomplete", !dataComplete), "\" ").concat(missingTitle, ">").concat(dataComplete ? "○" : "未記入", "</td>\n        <td data-col-key=\"specimenNo\" data-cell-edit-key=\"specimenNo\" class=\"").concat(listCellClass("", specimenMissing), "\" ").concat(missingTitle, ">").concat(escapeHtml(record.specimenNo), "</td>\n        <td data-col-key=\"category\" data-cell-edit-key=\"category\" class=\"").concat(listCellClass("category-color-cell", categoryMissing), "\" style=\"background:").concat(categoryBackground, ";color:#111827;border-color:").concat(categoryBorderColor, ";\" ").concat(missingTitle, ">").concat(escapeHtml(formatCategoryForRecord(record)), "</td>\n        <td data-col-key=\"nameMemo\" data-cell-edit-key=\"nameMemo\" class=\"").concat(listCellClass("", nameMissing), "\" ").concat(missingTitle, ">").concat(escapeHtml(record.nameMemo || ""), "</td>\n        <td data-col-key=\"importantFlag\" data-cell-edit-key=\"importantFlag\" class=\"").concat(listCellClass(record.importantFlag === "有" ? "important-cell-important" : "", importantMissing), "\" ").concat(missingTitle, ">").concat(escapeHtml(record.importantFlag || ""), "</td>\n        <td data-col-key=\"unit\" data-cell-edit-key=\"unit\" class=\"").concat(listCellClass("unit-color-cell", unitMissing), "\" style=\"background:").concat(unitStyle.background, ";color:").concat(unitStyle.color, ";border-color:").concat(unitStyle.border, ";\" ").concat(missingTitle, ">").concat(escapeHtml(record.unit || ""), "</td>\n        <td data-col-key=\"detail\" data-cell-edit-key=\"detail\">").concat(escapeHtml(formatDetailForRecord(record)), "</td>\n        <td data-col-key=\"discoverer\" data-cell-edit-key=\"discoverer\" class=\"").concat(listCellClass("", discovererMissing), "\" ").concat(missingTitle, ">").concat(escapeHtml(record.discoverer || ""), "</td>\n        <td data-col-key=\"identifier\" data-cell-edit-key=\"identifier\" class=\"").concat(listCellClass("", identifierMissing), "\" ").concat(missingTitle, ">").concat(escapeHtml(record.identifier || ""), "</td>\n        <td data-col-key=\"levelRead\" data-cell-edit-key=\"levelRead\" class=\"").concat(listCellClass("", levelReadMissing), "\" ").concat(missingTitle, ">").concat(escapeHtml(formatLevelRead(record)), "</td>\n        <td data-col-key=\"altitudeM\" data-cell-edit-key=\"altitudeM\" class=\"").concat(listCellClass("", levelHeightMissing || levelReadMissing), "\" ").concat(missingTitle, ">").concat(escapeHtml(formatRecordAltitudeM(record)), "</td>\n        <td data-col-key=\"occurrenceSection\" data-cell-edit-key=\"occurrenceSection\" class=\"").concat(listCellClass("", sectionMissing || sectionChecklistMissing), "\" ").concat(missingTitle, ">").concat(escapeHtml(record.occurrenceSection || ""), "</td>\n        <td data-col-key=\"occurrenceSketch\" data-cell-edit-key=\"occurrenceSketch\" class=\"").concat(listCellClass("", sketchMissing), "\" ").concat(missingTitle, ">").concat(escapeHtml(record.occurrenceSketch || ""), "</td>\n        <td data-col-key=\"position\" data-cell-edit-key=\"position\" class=\"").concat(listCellClass("", positionMissing), "\" ").concat(missingTitle, ">").concat(escapeHtml(formatPlanPosition(record)), "</td>\n        <td data-col-key=\"notes\" data-cell-edit-key=\"notes\">").concat(escapeHtml(record.notes || ""), "</td>\n        <td data-col-key=\"actions\">\n          <div class=\"row-actions\">\n            <button type=\"button\" data-action=\"insert-row\" data-id=\"").concat(record.id, "\" data-kuwaku=\"").concat(escapeHtml(getRecordKuwaku(record)), "\" data-record-index=\"").concat(Number.isInteger(recordIndex) && recordIndex >= 0 ? recordIndex : "", "\">\u884C\u633F\u5165</button>\n            <button type=\"button\" data-action=\"copy-to-input\" data-id=\"").concat(record.id, "\" data-kuwaku=\"").concat(escapeHtml(getRecordKuwaku(record)), "\" data-record-index=\"").concat(Number.isInteger(recordIndex) && recordIndex >= 0 ? recordIndex : "", "\">\u30B3\u30D4\u30FC\u3057\u3066\u65B0\u898F\u5165\u529B</button>\n            <button type=\"button\" data-action=\"edit\" data-id=\"").concat(record.id, "\" data-kuwaku=\"").concat(escapeHtml(getRecordKuwaku(record)), "\" data-record-index=\"").concat(Number.isInteger(recordIndex) && recordIndex >= 0 ? recordIndex : "", "\">\u7DE8\u96C6</button>\n            <button class=\"danger\" type=\"button\" data-action=\"delete\" data-id=\"").concat(record.id, "\" data-record-index=\"").concat(Number.isInteger(recordIndex) && recordIndex >= 0 ? recordIndex : "", "\">\u524A\u9664</button>\n          </div>\n        </td>\n      </tr>\n      ");
  }).join("");
  applyOutputListColumnVisibility();
}
function listCellClass(baseClass, isMissing) {
  var classes = [];
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
  return keys.some(function (key) {
    return missingKeys.has(key);
  });
}
function formatMissingRequiredTooltip(missingKeys) {
  if (!(missingKeys instanceof Set) || missingKeys.size === 0) {
    return "";
  }
  var labels = Array.from(missingKeys).map(function (key) {
    return REQUIRED_FIELD_LABELS[key] || key;
  });
  return "title=\"".concat(escapeHtml("\u672A\u8A18\u5165: ".concat(labels.join(" / "))), "\"");
}
function updateOutputListSortHeader() {
  if (!outputListTable) {
    return;
  }
  var headers = outputListTable.querySelectorAll("th[data-sort-key]");
  headers.forEach(function (header) {
    var sortKey = value(header.dataset.sortKey);
    var isActive = sortKey === outputListSortKey;
    header.classList.add("sortable-header");
    header.classList.toggle("sort-asc", isActive && outputListSortDirection === "asc");
    header.classList.toggle("sort-desc", isActive && outputListSortDirection === "desc");
    header.setAttribute("aria-sort", isActive ? outputListSortDirection === "asc" ? "ascending" : "descending" : "none");
  });
}
function sortOutputRecordsForList(records) {
  var list = _toConsumableArray(records);
  list.sort(compareOutputRecordsForList);
  return list;
}
function compareOutputRecordsForList(a, b) {
  var compared = 0;
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
      compared = compareSortText(a === null || a === void 0 ? void 0 : a.nameMemo, b === null || b === void 0 ? void 0 : b.nameMemo);
      break;
    case "importantFlag":
      compared = compareSortText(a === null || a === void 0 ? void 0 : a.importantFlag, b === null || b === void 0 ? void 0 : b.importantFlag);
      break;
    case "discoverer":
      compared = compareSortText(a === null || a === void 0 ? void 0 : a.discoverer, b === null || b === void 0 ? void 0 : b.discoverer);
      break;
    case "identifier":
      compared = compareSortText(a === null || a === void 0 ? void 0 : a.identifier, b === null || b === void 0 ? void 0 : b.identifier);
      break;
    case "levelRead":
      compared = compareSortText(formatLevelRead(a), formatLevelRead(b));
      break;
    case "altitudeM":
      compared = compareNullableNumber(getRecordAltitudeMValue(a), getRecordAltitudeMValue(b));
      break;
    case "occurrenceSection":
      compared = compareSortText(a === null || a === void 0 ? void 0 : a.occurrenceSection, b === null || b === void 0 ? void 0 : b.occurrenceSection);
      break;
    case "occurrenceSketch":
      compared = compareSortText(a === null || a === void 0 ? void 0 : a.occurrenceSketch, b === null || b === void 0 ? void 0 : b.occurrenceSketch);
      break;
    case "position":
      compared = compareSortText(formatPlanPosition(a), formatPlanPosition(b));
      break;
    case "unit":
      compared = compareSortText(a === null || a === void 0 ? void 0 : a.unit, b === null || b === void 0 ? void 0 : b.unit);
      break;
    case "detail":
      compared = compareSortText(formatDetailForRecord(a), formatDetailForRecord(b));
      break;
    case "notes":
      compared = compareSortText(a === null || a === void 0 ? void 0 : a.notes, b === null || b === void 0 ? void 0 : b.notes);
      break;
    default:
      compared = compareRecordsByKuwakuThenSpecimen(a, b);
      break;
  }
  var fallback = compareRecordsByKuwakuThenSpecimen(a, b);
  var direction = outputListSortDirection === "desc" ? -1 : 1;
  return (compared || fallback) * direction;
}
function compareSortText(a, b) {
  return value(a).localeCompare(value(b), "ja", {
    numeric: true,
    sensitivity: "base"
  });
}
function compareNullableNumber(a, b) {
  var aValid = Number.isFinite(a);
  var bValid = Number.isFinite(b);
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
  var a = normalizeDateForExportRange(aRaw);
  var b = normalizeDateForExportRange(bRaw);
  if (a && b) {
    return a.localeCompare(b, "ja", {
      numeric: true,
      sensitivity: "base"
    });
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
  var nsDir = value(record === null || record === void 0 ? void 0 : record.nsDir);
  var nsCm = formatCmValue(record === null || record === void 0 ? void 0 : record.nsCm);
  var ewDir = value(record === null || record === void 0 ? void 0 : record.ewDir);
  var ewCm = formatCmValue(record === null || record === void 0 ? void 0 : record.ewCm);
  var nsPart = "".concat(nsDir).concat(nsCm);
  var ewPart = "".concat(ewDir).concat(ewCm);
  var base = "";
  if (nsPart && ewPart) {
    base = "".concat(nsPart, " / ").concat(ewPart);
  } else {
    base = nsPart || ewPart;
  }
  var planSizeMode = normalizePlanSizeMode(record === null || record === void 0 ? void 0 : record.planSizeMode);
  if (planSizeMode === "複数点") {
    var multiPointCount = collectPlanMultiPointCoords(record).length;
    if (multiPointCount > 0) {
      return base ? "\u8907\u6570\u70B9(".concat(multiPointCount, "\u70B9) / ").concat(base) : "\u8907\u6570\u70B9(".concat(multiPointCount, "\u70B9)");
    }
  }
  if (planSizeMode !== "大きなもの") {
    return base;
  }
  var axisDirection = normalizeLargeAxisDirection(record === null || record === void 0 ? void 0 : record.largeAxisDirection);
  var plungeDeg = normalizeLargeAxisPlungeDeg(record === null || record === void 0 ? void 0 : record.largeAxisPlungeDeg);
  var plungeDir8 = normalizeCompass8Direction(record === null || record === void 0 ? void 0 : record.largeAxisPlungeDir8);
  var plungeText = plungeDeg ? "\u30D7\u30E9\u30F3\u30B8:".concat(plungeDeg).concat(plungeDir8 ? "(".concat(plungeDir8, ")") : "") : "";
  var planeStrike = normalizePlaneStrikeDirection(record === null || record === void 0 ? void 0 : record.planeStrikeDirection);
  var planeDip = normalizePlaneDipDeg(record === null || record === void 0 ? void 0 : record.planeDipDeg);
  var planeDipDir8 = normalizeCompass8Direction(record === null || record === void 0 ? void 0 : record.planeDipDir8);
  var planeAttitudeText = planeStrike && planeDip ? "\u8D70\u5411\u50BE\u659C:".concat(planeStrike, "/").concat(planeDip).concat(planeDipDir8 ? "(".concat(planeDipDir8, ")") : "") : "";
  var shapeType = normalizeLargeShapeType(record === null || record === void 0 ? void 0 : record.largeShapeType);
  var lineLength = shapeType === "直線状" ? formatCmValue(record === null || record === void 0 ? void 0 : record.lineLengthCm) : "";
  if (!base) {
    if (shapeType === "直線状") {
      if (axisDirection && lineLength) {
        return ["\u65B9\u4F4D:".concat(axisDirection), "\u9577\u3055:".concat(lineLength), plungeText].filter(Boolean).join(" / ");
      }
      return [axisDirection ? "\u65B9\u4F4D:".concat(axisDirection) : "", lineLength ? "\u9577\u3055:".concat(lineLength) : "", plungeText].filter(Boolean).join(" / ");
    }
    return [planeAttitudeText].filter(Boolean).join(" / ");
  }
  if (shapeType === "直線状") {
    if (axisDirection && lineLength) {
      return [base, "\u65B9\u4F4D:".concat(axisDirection), "\u9577\u3055:".concat(lineLength), plungeText].filter(Boolean).join(" / ");
    }
    if (axisDirection) {
      return [base, "\u65B9\u4F4D:".concat(axisDirection), plungeText].filter(Boolean).join(" / ");
    }
    if (lineLength) {
      return [base, "\u9577\u3055:".concat(lineLength), plungeText].filter(Boolean).join(" / ");
    }
  }
  if (shapeType !== "直線状" && planeAttitudeText) {
    return [base, planeAttitudeText].filter(Boolean).join(" / ");
  }
  if (axisDirection || plungeText) {
    return [base, axisDirection ? "\u65B9\u4F4D:".concat(axisDirection) : "", plungeText].filter(Boolean).join(" / ");
  }
  return base;
}
function renderCardOutput() {
  if (!state.records.length) {
    selectedCardRecordId = "";
    cardOutputList.innerHTML = "";
    return;
  }
  var filteredRecords = getFilteredOutputRecords();
  if (!filteredRecords.length) {
    selectedCardRecordId = "";
    cardOutputList.innerHTML = "";
    return;
  }
  var hasSelected = filteredRecords.some(function (item) {
    return item.id === selectedCardRecordId;
  });
  if (!hasSelected) {
    selectedCardRecordId = "";
    cardOutputList.innerHTML = "";
    return;
  }
  var selectedRecord = filteredRecords.find(function (item) {
    return item.id === selectedCardRecordId;
  });
  if (!selectedRecord) {
    cardOutputList.innerHTML = "";
    return;
  }
  var sectionDiagramsHtml = (selectedRecord.sectionDiagrams || []).length ? "<div class=\"card-photo-grid\">".concat(selectedRecord.sectionDiagrams.map(function (item) {
    return "<figure><img src=\"".concat(item.dataUrl, "\" alt=\"").concat(escapeHtml(item.name || "diagram"), "\" /><figcaption>").concat(escapeHtml(item.caption || ""), "</figcaption></figure>");
  }).join(""), "</div>") : "<p class=\"muted\">断面図なし</p>";
  var photosHtml = (selectedRecord.photos || []).length ? "<div class=\"card-photo-grid\">".concat(selectedRecord.photos.map(function (photo) {
    return "<figure><img src=\"".concat(photo.dataUrl, "\" alt=\"").concat(escapeHtml(photo.name || "photo"), "\" /><figcaption>").concat(escapeHtml(photo.caption || ""), "</figcaption></figure>");
  }).join(""), "</div>") : "<p class=\"muted\">写真なし</p>";
  cardOutputList.innerHTML = "\n    <article class=\"card-output-item\">\n      <h3>".concat(escapeHtml(selectedRecord.specimenNo), " / ").concat(escapeHtml(selectedRecord.nameMemo || ""), "</h3>\n      <div class=\"kv-grid\">\n        <div><span>\u5206\u985E</span><strong>").concat(escapeHtml(formatCategoryForRecord(selectedRecord)), "</strong></div>\n        <div><span>\u91CD\u8981\u54C1\u6307\u5B9A</span><strong>").concat(escapeHtml(selectedRecord.importantFlag || ""), "</strong></div>\n        <div><span>\u7C21\u6613\u8A18\u8F09</span><strong>").concat(escapeHtml(selectedRecord.simpleRecordFlag || "-"), "</strong></div>\n        <div><span>\u5730\u5C64\u540D</span><strong>").concat(escapeHtml(selectedRecord.layerName || ""), "</strong></div>\n        <div><span>\u30E6\u30CB\u30C3\u30C8</span><strong>").concat(escapeHtml(selectedRecord.unit || ""), "</strong></div>\n        <div><span>\u30B5\u30D6\u30E6\u30CB\u30C3\u30C8</span><strong>").concat(escapeHtml(formatDetailForRecord(selectedRecord)), "</strong></div>\n        <div><span>\u5C64\u76F8</span><strong>").concat(escapeHtml(selectedRecord.layerFacies || ""), "</strong></div>\n        <div><span>\u5730\u5C64\u4E2D\u306E\u4F4D\u7F6E</span><strong>").concat(escapeHtml(formatLayerPosition(selectedRecord)), "</strong></div>\n        <div><span>\u767A\u898B\u8005</span><strong>").concat(escapeHtml(selectedRecord.discoverer || ""), "</strong></div>\n        <div><span>\u5224\u5B9A\u8005</span><strong>").concat(escapeHtml(selectedRecord.identifier || ""), "</strong></div>\n        <div><span>\u30EC\u30D9\u30EB\u8AAD\u5024(\u4E0A\u9762/\u4E0B\u5E95)</span><strong>").concat(escapeHtml(formatLevelRead(selectedRecord)), "</strong></div>\n        <div><span>\u7523\u51FA\u72B6\u6CC1\u65AD\u9762</span><strong>").concat(escapeHtml(selectedRecord.occurrenceSection || ""), "</strong></div>\n        <div><span>\u7523\u72B6\u30B9\u30B1\u30C3\u30C1</span><strong>").concat(escapeHtml(selectedRecord.occurrenceSketch || ""), "</strong></div>\n        <div><span>\u5E73\u9762\u4F4D\u7F6E</span><strong>").concat(escapeHtml(formatPlanPosition(selectedRecord)), "</strong></div>\n      </div>\n      <p><strong>\u5099\u8003\uFF08\u89B3\u5BDF\u4E8B\u9805\u306A\u3069\uFF09:</strong> ").concat(escapeHtml(selectedRecord.notes || ""), "</p>\n      <p><strong>\u7523\u51FA\u72B6\u6CC1\u65AD\u9762\u56F3:</strong></p>\n      ").concat(sectionDiagramsHtml, "\n      <p><strong>\u5199\u771F:</strong></p>\n      ").concat(photosHtml, "\n    </article>\n  ");
}
function getFilteredOutputRecords() {
  var sortedRecords = _toConsumableArray(state.records).sort(compareRecordsByKuwakuThenSpecimen);
  syncOutputKuwakuSelect(sortedRecords);
  var kuwakuScopedRecords = selectedOutputKuwaku === ALL_GRIDS_VALUE ? sortedRecords : sortedRecords.filter(function (record) {
    return kuwakuValueForSelect(getRecordKuwaku(record)) === selectedOutputKuwaku;
  });
  syncOutputCategorySelect(kuwakuScopedRecords);
  syncOutputStatusSelect();
  var filteredRecords = kuwakuScopedRecords;
  if (selectedOutputCategory && selectedOutputCategory !== EXPORT_CATEGORY_ALL_VALUE) {
    filteredRecords = filteredRecords.filter(function (record) {
      var specimen = parseSpecimenNo(record.specimenNo, record.specimenPrefix, record.specimenSerial);
      return normalizeSpecimenPrefix(specimen.prefix) === selectedOutputCategory;
    });
  }
  if (selectedOutputStatus === "complete") {
    filteredRecords = filteredRecords.filter(function (record) {
      return isRecordDataComplete(record);
    });
  } else if (selectedOutputStatus === "incomplete") {
    filteredRecords = filteredRecords.filter(function (record) {
      return !isRecordDataComplete(record);
    });
  }
  syncOutputDateSelect(filteredRecords);
  if (selectedOutputDate) {
    filteredRecords = filteredRecords.filter(function (record) {
      return normalizeDateForExportRange(getRecordDate(record)) === selectedOutputDate;
    });
  }
  syncOutputSearchInput();
  var searchText = value(outputSearchText).toLowerCase();
  if (searchText) {
    filteredRecords = filteredRecords.filter(function (record) {
      return buildOutputFilterSearchText(record).includes(searchText);
    });
  }
  updateOutputFilterSummary(filteredRecords.length, sortedRecords.length);
  return filteredRecords;
}
function syncOutputKuwakuSelect(records) {
  if (!outputKuwakuSelect) {
    return;
  }
  var options = collectOutputKuwakuOptions(records);
  if (!options.some(function (item) {
    return item.value === selectedOutputKuwaku;
  })) {
    selectedOutputKuwaku = ALL_GRIDS_VALUE;
  }
  outputKuwakuSelect.innerHTML = options.map(function (item) {
    return "<option value=\"".concat(escapeHtml(item.value), "\" ").concat(item.value === selectedOutputKuwaku ? "selected" : "", ">").concat(escapeHtml(item.label), "</option>");
  }).join("");
}
function syncOutputCategorySelect(records) {
  if (!outputCategorySelect) {
    return;
  }
  var options = collectExportCategoryOptions(records);
  if (!options.some(function (item) {
    return item.value === selectedOutputCategory;
  })) {
    selectedOutputCategory = EXPORT_CATEGORY_ALL_VALUE;
  }
  outputCategorySelect.innerHTML = options.map(function (item) {
    return "<option value=\"".concat(escapeHtml(item.value), "\" ").concat(item.value === selectedOutputCategory ? "selected" : "", ">").concat(escapeHtml(item.label), "</option>");
  }).join("");
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
  var options = collectOutputDateOptions(records);
  if (!options.some(function (item) {
    return item.value === selectedOutputDate;
  })) {
    selectedOutputDate = "";
  }
  outputDateSelect.innerHTML = options.map(function (item) {
    return "<option value=\"".concat(escapeHtml(item.value), "\" ").concat(item.value === selectedOutputDate ? "selected" : "", ">").concat(escapeHtml(item.label), "</option>");
  }).join("");
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
  return [getRecordKuwaku(record), record.specimenNo, formatCategoryForRecord(record), record.nameMemo, record.unit, formatDetailForRecord(record), getRecordTeamValue(record), record.discoverer, record.identifier, formatPlanPosition(record), record.notes].map(function (item) {
    return value(item).toLowerCase();
  }).join(" ");
}
function updateOutputFilterSummary(filteredCount, totalCount) {
  if (!outputFilterSummary) {
    return;
  }
  var filtered = Number.isFinite(filteredCount) ? filteredCount : 0;
  var total = Number.isFinite(totalCount) ? totalCount : 0;
  outputFilterSummary.textContent = "\u8868\u793A: ".concat(filtered, "\u4EF6 / \u5168\u4F53: ").concat(total, "\u4EF6");
}
function collectOutputKuwakuOptions(records) {
  if (!records.length) {
    return [{
      value: ALL_GRIDS_VALUE,
      label: "全グリッド"
    }];
  }
  var kuwakuSet = new Set(records.map(function (record) {
    return kuwakuValueForSelect(getRecordKuwaku(record));
  }));
  var options = Array.from(kuwakuSet).sort(function (a, b) {
    return kuwakuLabelForSelect(a).localeCompare(kuwakuLabelForSelect(b), "ja", {
      numeric: true,
      sensitivity: "base"
    });
  }).map(function (kuwakuValue) {
    return {
      value: kuwakuValue,
      label: kuwakuLabelForSelect(kuwakuValue)
    };
  });
  return [{
    value: ALL_GRIDS_VALUE,
    label: "全グリッド"
  }].concat(_toConsumableArray(options));
}
function collectOutputDateOptions(records) {
  if (!records.length) {
    return [{
      value: "",
      label: "すべて"
    }];
  }
  var dateSet = new Set();
  records.forEach(function (record) {
    var normalizedDate = normalizeDateForExportRange(getRecordDate(record));
    if (normalizedDate) {
      dateSet.add(normalizedDate);
    }
  });
  var options = Array.from(dateSet).sort(function (a, b) {
    return b.localeCompare(a, "ja", {
      numeric: true,
      sensitivity: "base"
    });
  }).map(function (dateValue) {
    return {
      value: dateValue,
      label: dateValue
    };
  });
  return [{
    value: "",
    label: "すべて"
  }].concat(_toConsumableArray(options));
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
  var kuwakuFilteredRecords = getFilteredPlanRecords();
  syncPlanCategorySelect(kuwakuFilteredRecords);
  var categoryRecords = filterRecordsByCategory(kuwakuFilteredRecords, selectedPlanCategory);
  if (planKuwakuInfo) {
    var kuwakuLabel = selectedPlanKuwaku ? kuwakuLabelForSelect(selectedPlanKuwaku) : "-";
    planKuwakuInfo.textContent = "\u533A\u753B: ".concat(kuwakuLabel);
  }
  if (!categoryRecords.length) {
    selectedPlanUnit = "";
    selectedPlanDetail = ALL_DETAILS_VALUE;
    selectedPlanDetailSub = ALL_DETAIL_SUBS_VALUE;
    planUnitSelect.innerHTML = "";
    planDetailSelect.innerHTML = "";
    planDetailSubSelect.innerHTML = "";
    var _kuwakuLabelForMeta = selectedPlanKuwaku ? kuwakuLabelForSelect(selectedPlanKuwaku) : "-";
    var _categoryLabelForMeta = selectedPlanCategory === EXPORT_CATEGORY_ALL_VALUE ? "全分類" : "".concat(selectedPlanCategory, ": ").concat(SPECIMEN_CATEGORY_MAP[selectedPlanCategory] || "");
    planMapWrap.innerHTML = "\n      <div class=\"plan-map-meta\">\n        <span>\u533A\u753B\uFF08\u30B0\u30EA\u30C3\u30C9\uFF09: ".concat(escapeHtml(_kuwakuLabelForMeta), "</span>\n        <span>\u5206\u985E: ").concat(escapeHtml(_categoryLabelForMeta), "</span>\n        <span>\u51FA\u529B\u968E\u5C64: -</span>\n        <span>\u8868\u793A\u4EF6\u6570: 0\u4EF6</span>\n      </div>\n      <p class=\"muted\">\u3053\u306E\u6761\u4EF6\u306B\u306F\u8868\u793A\u5BFE\u8C61\u30C7\u30FC\u30BF\u304C\u3042\u308A\u307E\u305B\u3093\u3002</p>\n    ");
    return;
  }
  var units = collectPlanUnits(categoryRecords);
  if (!units.some(function (unit) {
    return unit.value === selectedPlanUnit;
  })) {
    selectedPlanUnit = units[0].value;
  }
  planUnitSelect.innerHTML = units.map(function (unit) {
    return "<option value=\"".concat(escapeHtml(unit.value), "\" ").concat(unit.value === selectedPlanUnit ? "selected" : "", ">").concat(escapeHtml(unit.label), "</option>");
  }).join("");
  var unitRecords = selectedPlanUnit === ALL_UNITS_VALUE ? categoryRecords : categoryRecords.filter(function (record) {
    return unitValueForSelect(record.unit) === selectedPlanUnit;
  });
  var details = collectPlanDetails(unitRecords);
  if (!details.some(function (detail) {
    return detail.value === selectedPlanDetail;
  })) {
    selectedPlanDetail = details[0].value;
  }
  planDetailSelect.innerHTML = details.map(function (detail) {
    return "<option value=\"".concat(escapeHtml(detail.value), "\" ").concat(detail.value === selectedPlanDetail ? "selected" : "", ">").concat(escapeHtml(detail.label), "</option>");
  }).join("");
  var detailRecords = selectedPlanDetail === ALL_DETAILS_VALUE ? unitRecords : unitRecords.filter(function (record) {
    return detailValueForSelect(record.detail) === selectedPlanDetail;
  });
  var detailSubs = collectPlanDetailSubs(detailRecords);
  if (!detailSubs.some(function (detailSub) {
    return detailSub.value === selectedPlanDetailSub;
  })) {
    selectedPlanDetailSub = detailSubs[0].value;
  }
  planDetailSubSelect.innerHTML = detailSubs.map(function (detailSub) {
    return "<option value=\"".concat(escapeHtml(detailSub.value), "\" ").concat(detailSub.value === selectedPlanDetailSub ? "selected" : "", ">").concat(escapeHtml(detailSub.label), "</option>");
  }).join("");
  var detailSubRecords = selectedPlanDetailSub === ALL_DETAIL_SUBS_VALUE ? detailRecords : detailRecords.filter(function (record) {
    return detailSubValueForSelect(record.detailSub) === selectedPlanDetailSub;
  });
  var drawables = detailSubRecords.map(function (record) {
    return buildPlanDrawable(record);
  }).filter(Boolean);
  var kuwakuLabelForMeta = selectedPlanKuwaku ? kuwakuLabelForSelect(selectedPlanKuwaku) : "-";
  var categoryLabelForMeta = selectedPlanCategory === EXPORT_CATEGORY_ALL_VALUE ? "全分類" : "".concat(selectedPlanCategory, ": ").concat(SPECIMEN_CATEGORY_MAP[selectedPlanCategory] || "");
  var unitLabelForMeta = selectedPlanUnit === ALL_UNITS_VALUE ? "全ユニット" : unitLabelForSelect(selectedPlanUnit);
  var detailLabelForMeta = selectedPlanDetail === ALL_DETAILS_VALUE ? "全サブユニット" : detailLabelForSelect(selectedPlanDetail);
  var detailSubLabelForMeta = selectedPlanDetailSub === ALL_DETAIL_SUBS_VALUE ? "全細分" : detailSubLabelForSelect(selectedPlanDetailSub);
  var hierarchyLabelForMeta = "".concat(unitLabelForMeta, " > ").concat(detailLabelForMeta, " > ").concat(detailSubLabelForMeta);
  var mapMetaHtml = "\n    <div class=\"plan-map-meta\">\n      <span>\u533A\u753B\uFF08\u30B0\u30EA\u30C3\u30C9\uFF09: ".concat(escapeHtml(kuwakuLabelForMeta), "</span>\n      <span>\u5206\u985E: ").concat(escapeHtml(categoryLabelForMeta), "</span>\n      <span>\u51FA\u529B\u968E\u5C64: ").concat(escapeHtml(hierarchyLabelForMeta), "</span>\n      <span>\u8868\u793A\u4EF6\u6570: ").concat(detailSubRecords.length, "\u4EF6</span>\n    </div>\n  ");
  if (!drawables.length) {
    planMapWrap.innerHTML = "\n      ".concat(mapMetaHtml, "\n      <p class=\"muted\">\u3053\u306E\u30E6\u30CB\u30C3\u30C8/\u30B5\u30D6\u30E6\u30CB\u30C3\u30C8/\u7D30\u5206\u306F\u3001\u5E73\u9762\u4F4D\u7F6E\u306E\u5165\u529B\u304C\u4E0D\u8DB3\u3057\u3066\u3044\u308B\u305F\u3081\u8868\u793A\u3067\u304D\u307E\u305B\u3093\u3002</p>\n    ");
    return;
  }
  var verticalGrid = [100, 200, 300].map(function (x) {
    return "<line class=\"plan-grid-line\" x1=\"".concat(x, "\" y1=\"0\" x2=\"").concat(x, "\" y2=\"").concat(PLAN_SIZE_CM, "\" />");
  }).join("");
  var horizontalGrid = [100, 200, 300].map(function (y) {
    return "<line class=\"plan-grid-line\" x1=\"0\" y1=\"".concat(y, "\" x2=\"").concat(PLAN_SIZE_CM, "\" y2=\"").concat(y, "\" />");
  }).join("");
  var pointsSvg = drawables.map(function (drawable, index) {
    return renderPlanDrawableSvg(drawable, index);
  }).join("");
  var cornerLabels = buildPlanCornerLabels(selectedPlanKuwaku);
  var cornerLabelsSvg = buildPlanCornerLabelsSvg(cornerLabels);
  var viewBox = computePlanSvgViewBox(drawables);
  planMapWrap.innerHTML = "\n    ".concat(mapMetaHtml, "\n    <div class=\"plan-map-shell\">\n      <div class=\"plan-axis north\">\u5317</div>\n      <div class=\"plan-axis east\">\u6771</div>\n      <div class=\"plan-axis south\">\u5357</div>\n      <div class=\"plan-axis west\">\u897F</div>\n      <svg class=\"plan-map-svg\" viewBox=\"").concat(viewBox.minX, " ").concat(viewBox.minY, " ").concat(viewBox.width, " ").concat(viewBox.height, "\" aria-label=\"\u30E6\u30CB\u30C3\u30C8\u5225\u5E73\u9762\u56F3\" xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\">\n        <rect class=\"plan-frame\" x=\"0\" y=\"0\" width=\"").concat(PLAN_SIZE_CM, "\" height=\"").concat(PLAN_SIZE_CM, "\" />\n        ").concat(verticalGrid, "\n        ").concat(horizontalGrid, "\n        ").concat(cornerLabelsSvg, "\n        ").concat(pointsSvg, "\n      </svg>\n      <div class=\"plan-map-tooltip\" hidden></div>\n    </div>\n  ");
  attachPlanMapTooltips();
}
function buildCurrentRecordDraftForPositionPreview() {
  var _siteForm$elements9, _siteForm$elements0, _siteForm$elements1, _siteForm$elements10;
  if (!recordForm) {
    return null;
  }
  var formData = new FormData(recordForm);
  var isEditTab = getActiveTabId() === "edit-tab";
  var kuwaku = isEditTab ? buildKuwaku(normalizeKuwakuHeadA(editKuwakuHeadAInput === null || editKuwakuHeadAInput === void 0 ? void 0 : editKuwakuHeadAInput.value), normalizeKuwakuHeadB(editKuwakuHeadBInput === null || editKuwakuHeadBInput === void 0 ? void 0 : editKuwakuHeadBInput.value), normalizeKuwakuBlock(editKuwakuBlockInput === null || editKuwakuBlockInput === void 0 ? void 0 : editKuwakuBlockInput.value), normalizeKuwakuNo(editKuwakuNoInput === null || editKuwakuNoInput === void 0 ? void 0 : editKuwakuNoInput.value)) : buildKuwaku(normalizeKuwakuHeadA(siteForm === null || siteForm === void 0 || (_siteForm$elements9 = siteForm.elements) === null || _siteForm$elements9 === void 0 || (_siteForm$elements9 = _siteForm$elements9.kuwakuHeadA) === null || _siteForm$elements9 === void 0 ? void 0 : _siteForm$elements9.value), normalizeKuwakuHeadB(siteForm === null || siteForm === void 0 || (_siteForm$elements0 = siteForm.elements) === null || _siteForm$elements0 === void 0 || (_siteForm$elements0 = _siteForm$elements0.kuwakuHeadB) === null || _siteForm$elements0 === void 0 ? void 0 : _siteForm$elements0.value), normalizeKuwakuBlock(siteForm === null || siteForm === void 0 || (_siteForm$elements1 = siteForm.elements) === null || _siteForm$elements1 === void 0 || (_siteForm$elements1 = _siteForm$elements1.kuwakuBlock) === null || _siteForm$elements1 === void 0 ? void 0 : _siteForm$elements1.value), normalizeKuwakuNo(siteForm === null || siteForm === void 0 || (_siteForm$elements10 = siteForm.elements) === null || _siteForm$elements10 === void 0 || (_siteForm$elements10 = _siteForm$elements10.kuwakuNo) === null || _siteForm$elements10 === void 0 ? void 0 : _siteForm$elements10.value));
  var specimenPrefix = normalizeSpecimenPrefix(value(formData.get("specimenPrefix")));
  var specimenSerial = compactNoSpaceValue(formData.get("specimenSerial"));
  var planSizeMode = normalizePlanSizeMode(value(formData.get("planSizeMode")));
  var draftMultiPoints = planSizeMode === "複数点" ? readMultiPointRowsFromForm() : [];
  var rawLargeShapeType = value(formData.get("largeShapeType"));
  var largeShapeType = planSizeMode === "大きなもの" ? normalizeLargeShapeType(rawLargeShapeType) || normalizeLargeShapeLabel(rawLargeShapeType) : "";
  var imageCornerFields = extractImageCornerFieldsFromFormData(formData);
  var imageTransformFields = extractImageTransformFieldsFromFormData(formData);
  var isImageShape = isLargeShapeImageType(largeShapeType);
  var isCustomImageShape = isImageShape && isCustomLargeShapeType(largeShapeType);
  return {
    id: value(editingRecordId || (recordIdInput === null || recordIdInput === void 0 ? void 0 : recordIdInput.value)),
    kuwaku: kuwaku,
    specimenPrefix: specimenPrefix,
    specimenSerial: specimenSerial,
    specimenNo: buildSpecimenNo(specimenPrefix, specimenSerial),
    nameMemo: value(formData.get("nameMemo")),
    unit: compactNoSpaceValue(formData.get("unit")),
    detail: compactNoSpaceValue(formData.get("detail")),
    detailSub: value(formData.get("detailSub")),
    nsDir: normalizeNsDir(value(formData.get("nsDir"))),
    nsCm: value(formData.get("nsCm")),
    ewDir: normalizeEwDir(value(formData.get("ewDir"))),
    ewCm: value(formData.get("ewCm")),
    multiPoints: draftMultiPoints,
    planSizeMode: planSizeMode,
    largeShapeType: largeShapeType,
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
    customLargeImageDataUrl: isCustomImageShape ? normalizeCustomLargeImageDataUrl(value(formData.get("customLargeImageDataUrl"))) : ""
  };
}
function renderPositionPreviewModalContent() {
  if (!positionPreviewMeta || !positionPreviewMap) {
    return false;
  }
  var draftRecord = buildCurrentRecordDraftForPositionPreview();
  if (!draftRecord) {
    return false;
  }
  var currentDrawableRaw = buildPlanDrawable(draftRecord);
  if (!currentDrawableRaw) {
    showToast("平面位置の入力値を確認してください");
    return false;
  }
  var kuwakuValue = kuwakuValueForSelect(getRecordKuwaku(draftRecord));
  var savedDrawables = state.records.filter(function (record) {
    return kuwakuValueForSelect(getRecordKuwaku(record)) === kuwakuValue && value(record.id) !== value(draftRecord.id);
  }).map(function (record) {
    return buildPlanDrawable(record);
  }).filter(Boolean);
  var currentDrawable = _objectSpread(_objectSpread({}, currentDrawableRaw), {}, {
    color: "#dc2626",
    label: value(draftRecord.specimenNo) || "入力中"
  });
  var drawables = [].concat(_toConsumableArray(savedDrawables), [currentDrawable]);
  var verticalGrid = [100, 200, 300].map(function (x) {
    return "<line class=\"plan-grid-line\" x1=\"".concat(x, "\" y1=\"0\" x2=\"").concat(x, "\" y2=\"").concat(PLAN_SIZE_CM, "\" />");
  }).join("");
  var horizontalGrid = [100, 200, 300].map(function (y) {
    return "<line class=\"plan-grid-line\" x1=\"0\" y1=\"".concat(y, "\" x2=\"").concat(PLAN_SIZE_CM, "\" y2=\"").concat(y, "\" />");
  }).join("");
  var pointsSvg = drawables.map(function (drawable, index) {
    return renderPlanDrawableSvg(drawable, index);
  }).join("");
  var cornerLabels = buildPlanCornerLabels(kuwakuValue);
  var cornerLabelsSvg = buildPlanCornerLabelsSvg(cornerLabels);
  var viewBox = computePlanSvgViewBox(drawables);
  var kuwakuLabel = kuwakuValue === EMPTY_KUWAKU_VALUE ? "（未設定）" : kuwakuLabelForSelect(kuwakuValue);
  positionPreviewMeta.innerHTML = "\n    <span>\u533A\u753B\uFF08\u30B0\u30EA\u30C3\u30C9\uFF09: ".concat(escapeHtml(kuwakuLabel), "</span>\n    <span>\u8868\u793A\u4EF6\u6570: ").concat(drawables.length, "\u4EF6</span>\n  ");
  positionPreviewMap.innerHTML = "\n    <div class=\"plan-map-shell position-preview-shell\">\n      <div class=\"plan-axis north\">\u5317</div>\n      <div class=\"plan-axis east\">\u6771</div>\n      <div class=\"plan-axis south\">\u5357</div>\n      <div class=\"plan-axis west\">\u897F</div>\n      <svg class=\"plan-map-svg\" viewBox=\"".concat(viewBox.minX, " ").concat(viewBox.minY, " ").concat(viewBox.width, " ").concat(viewBox.height, "\" aria-label=\"\u5E73\u9762\u4F4D\u7F6E\u78BA\u8A8D\" xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\">\n        <rect class=\"plan-frame\" x=\"0\" y=\"0\" width=\"").concat(PLAN_SIZE_CM, "\" height=\"").concat(PLAN_SIZE_CM, "\" />\n        ").concat(verticalGrid, "\n        ").concat(horizontalGrid, "\n        ").concat(cornerLabelsSvg, "\n        ").concat(pointsSvg, "\n      </svg>\n    </div>\n  ");
  return true;
}
function openPositionPreviewModal() {
  if (!positionPreviewModal) {
    return;
  }
  var rendered = renderPositionPreviewModalContent();
  if (!rendered) {
    return;
  }
  positionPreviewModal.classList.remove("hidden");
}
function closePositionPreviewModal() {
  if (!positionPreviewModal) {
    return;
  }
  positionPreviewModal.classList.add("hidden");
}
function getFilteredPlanRecords() {
  var sortedRecords = _toConsumableArray(state.records).sort(compareRecordsByKuwakuThenSpecimen);
  syncPlanKuwakuSelect(sortedRecords);
  if (!selectedPlanKuwaku) {
    return [];
  }
  return sortedRecords.filter(function (record) {
    return kuwakuValueForSelect(getRecordKuwaku(record)) === selectedPlanKuwaku;
  });
}
function syncPlanKuwakuSelect(records) {
  if (!planKuwakuSelect) {
    return;
  }
  var options = collectOutputKuwakuOptions(records).filter(function (item) {
    return item.value !== ALL_GRIDS_VALUE;
  });
  if (!options.length) {
    selectedPlanKuwaku = "";
    planKuwakuSelect.innerHTML = "";
    return;
  }
  if (!options.some(function (item) {
    return item.value === selectedPlanKuwaku;
  })) {
    selectedPlanKuwaku = options[0].value;
  }
  planKuwakuSelect.innerHTML = options.map(function (item) {
    return "<option value=\"".concat(escapeHtml(item.value), "\" ").concat(item.value === selectedPlanKuwaku ? "selected" : "", ">").concat(escapeHtml(item.label), "</option>");
  }).join("");
}
function collectPlanUnits(records) {
  var unitSet = new Set(records.map(function (record) {
    return unitValueForSelect(record.unit);
  }));
  var unitOptions = Array.from(unitSet).sort(function (a, b) {
    return unitLabelForSelect(a).localeCompare(unitLabelForSelect(b), "ja", {
      numeric: true,
      sensitivity: "base"
    });
  }).map(function (unitValue) {
    return {
      value: unitValue,
      label: unitLabelForSelect(unitValue)
    };
  });
  return [{
    value: ALL_UNITS_VALUE,
    label: "全ユニット"
  }].concat(_toConsumableArray(unitOptions));
}
function unitValueForSelect(unitRaw) {
  var unit = value(unitRaw);
  return unit || EMPTY_UNIT_VALUE;
}
function unitLabelForSelect(unitValue) {
  return unitValue === EMPTY_UNIT_VALUE ? "（未設定）" : unitValue;
}
function collectPlanDetails(records) {
  var detailSet = new Set(records.map(function (record) {
    return detailValueForSelect(record.detail);
  }));
  var detailOptions = Array.from(detailSet).sort(function (a, b) {
    return detailLabelForSelect(a).localeCompare(detailLabelForSelect(b), "ja", {
      numeric: true,
      sensitivity: "base"
    });
  }).map(function (detailValue) {
    return {
      value: detailValue,
      label: detailLabelForSelect(detailValue)
    };
  });
  return [{
    value: ALL_DETAILS_VALUE,
    label: "全サブユニット"
  }].concat(_toConsumableArray(detailOptions));
}
function collectPlanDetailSubs(records) {
  var detailSubSet = new Set(records.map(function (record) {
    return detailSubValueForSelect(record.detailSub);
  }));
  var detailSubOptions = Array.from(detailSubSet).sort(function (a, b) {
    return detailSubLabelForSelect(a).localeCompare(detailSubLabelForSelect(b), "ja", {
      numeric: true,
      sensitivity: "base"
    });
  }).map(function (detailSubValue) {
    return {
      value: detailSubValue,
      label: detailSubLabelForSelect(detailSubValue)
    };
  });
  return [{
    value: ALL_DETAIL_SUBS_VALUE,
    label: "全細分"
  }].concat(_toConsumableArray(detailSubOptions));
}
function normalizeViewerVerticalScale(scaleRaw) {
  var num = Number(value(scaleRaw));
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
    viewerZScaleValue.textContent = "".concat(viewerVerticalScale.toFixed(1), "x");
  }
}
function applyViewerVerticalScale(zRaw, baseZRaw) {
  var z = Number(zRaw);
  var baseZ = Number(baseZRaw);
  if (!Number.isFinite(z) || !Number.isFinite(baseZ)) {
    return z;
  }
  return baseZ + (z - baseZ) * viewerVerticalScale;
}
function hideViewerFallbackPanel() {
  if (!viewerCanvasWrap) {
    return;
  }
  var panel = viewerCanvasWrap.querySelector(".viewer-fallback-panel");
  if (panel) {
    panel.remove();
  }
}
function showViewerFallbackPanel(reasonText) {
  var context = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : {};
  if (!viewerCanvasWrap) {
    return;
  }
  var reason = value(reasonText) || "この端末では3D表示が利用できません。";
  var kuwakuLabel = value(context.kuwakuLabel) || "-";
  var totalCount = Number.isFinite(Number(context.totalCount)) ? Number(context.totalCount) : 0;
  var drawableCount = Number.isFinite(Number(context.drawableCount)) ? Number(context.drawableCount) : 0;
  var panel = viewerCanvasWrap.querySelector(".viewer-fallback-panel");
  if (!panel) {
    panel = document.createElement("div");
    panel.className = "viewer-fallback-panel";
    viewerCanvasWrap.appendChild(panel);
  }
  panel.innerHTML = "\n    <div class=\"viewer-fallback-card\">\n      <p class=\"viewer-fallback-title\">\u3053\u306E\u7AEF\u672B\u3067\u306F3D\u8868\u793A\u304C\u4F7F\u3048\u307E\u305B\u3093</p>\n      <p class=\"viewer-fallback-reason\">".concat(escapeHtml(reason), "</p>\n      <p class=\"viewer-fallback-meta\">\u533A\u753B: ").concat(escapeHtml(kuwakuLabel), " / \u5BFE\u8C61 ").concat(totalCount, "\u4EF6 / \u63CF\u753B\u53EF\u80FD ").concat(drawableCount, "\u4EF6</p>\n      <button type=\"button\" class=\"viewer-fallback-open-plan-btn\" data-action=\"viewer-open-plan\">\u5E73\u9762\u56F3\u30BF\u30D6\u3067\u78BA\u8A8D</button>\n    </div>\n  ");
}
function renderViewerOutput() {
  var options = arguments.length > 0 && arguments[0] !== undefined ? arguments[0] : {};
  var preserveCamera = Boolean(options === null || options === void 0 ? void 0 : options.preserveCamera);
  if (!viewerKuwakuSelect || !viewerUnitSelect || !viewerDetailSelect || !viewerDetailSubSelect || !viewerMapLegend || !viewerStatus) {
    return;
  }
  hideViewerFallbackPanel();
  viewerMapLegend.innerHTML = buildPlanLegendHtml();
  syncViewerVerticalScaleUi();
  syncViewerViewButtons();
  var sortedRecords = _toConsumableArray(state.records).sort(compareRecordsByKuwakuThenSpecimen);
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
  var kuwakuRecords = selectedViewerKuwaku === ALL_GRIDS_VALUE ? sortedRecords : sortedRecords.filter(function (record) {
    return kuwakuValueForSelect(getRecordKuwaku(record)) === selectedViewerKuwaku;
  });
  syncViewerCategorySelect(kuwakuRecords);
  var categoryRecords = filterRecordsByCategory(kuwakuRecords, selectedViewerCategory);
  var kuwakuLabel = selectedViewerKuwaku === ALL_GRIDS_VALUE ? "全グリッド" : kuwakuLabelForSelect(selectedViewerKuwaku);
  if (viewerKuwakuInfo) {
    viewerKuwakuInfo.textContent = "\u533A\u753B: ".concat(kuwakuLabel);
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
  var units = collectPlanUnits(categoryRecords);
  if (!units.some(function (unit) {
    return unit.value === selectedViewerUnit;
  })) {
    var _units$;
    selectedViewerUnit = ((_units$ = units[0]) === null || _units$ === void 0 ? void 0 : _units$.value) || ALL_UNITS_VALUE;
  }
  viewerUnitSelect.innerHTML = units.map(function (unit) {
    return "<option value=\"".concat(escapeHtml(unit.value), "\" ").concat(unit.value === selectedViewerUnit ? "selected" : "", ">").concat(escapeHtml(unit.label), "</option>");
  }).join("");
  var unitRecords = selectedViewerUnit === ALL_UNITS_VALUE ? categoryRecords : categoryRecords.filter(function (record) {
    return unitValueForSelect(record.unit) === selectedViewerUnit;
  });
  var details = collectPlanDetails(unitRecords);
  if (!details.some(function (detail) {
    return detail.value === selectedViewerDetail;
  })) {
    var _details$;
    selectedViewerDetail = ((_details$ = details[0]) === null || _details$ === void 0 ? void 0 : _details$.value) || ALL_DETAILS_VALUE;
  }
  viewerDetailSelect.innerHTML = details.map(function (detail) {
    return "<option value=\"".concat(escapeHtml(detail.value), "\" ").concat(detail.value === selectedViewerDetail ? "selected" : "", ">").concat(escapeHtml(detail.label), "</option>");
  }).join("");
  var detailRecords = selectedViewerDetail === ALL_DETAILS_VALUE ? unitRecords : unitRecords.filter(function (record) {
    return detailValueForSelect(record.detail) === selectedViewerDetail;
  });
  var detailSubs = collectPlanDetailSubs(detailRecords);
  if (!detailSubs.some(function (detailSub) {
    return detailSub.value === selectedViewerDetailSub;
  })) {
    var _detailSubs$;
    selectedViewerDetailSub = ((_detailSubs$ = detailSubs[0]) === null || _detailSubs$ === void 0 ? void 0 : _detailSubs$.value) || ALL_DETAIL_SUBS_VALUE;
  }
  viewerDetailSubSelect.innerHTML = detailSubs.map(function (detailSub) {
    return "<option value=\"".concat(escapeHtml(detailSub.value), "\" ").concat(detailSub.value === selectedViewerDetailSub ? "selected" : "", ">").concat(escapeHtml(detailSub.label), "</option>");
  }).join("");
  var detailSubRecords = selectedViewerDetailSub === ALL_DETAIL_SUBS_VALUE ? detailRecords : detailRecords.filter(function (record) {
    return detailSubValueForSelect(record.detailSub) === selectedViewerDetailSub;
  });
  var viewerCandidates = [];
  var missingPositionCount = 0;
  var missingAltitudeCount = 0;
  var _iterator4 = _createForOfIteratorHelper(detailSubRecords),
    _step6;
  try {
    for (_iterator4.s(); !(_step6 = _iterator4.n()).done;) {
      var record = _step6.value;
      var drawable = buildPlanDrawable(record);
      if (!drawable) {
        missingPositionCount += 1;
        continue;
      }
      var altitudeM = getRecordAltitudeMValue(record);
      var altitudeEstimated = false;
      if (altitudeM == null) {
        missingAltitudeCount += 1;
        altitudeM = VIEWER_ALTITUDE_BASE_M;
        altitudeEstimated = true;
      }
      var kuwaku = parseKuwaku(getRecordKuwaku(record));
      var xIndex = kuwakuToViewerXIndex(kuwaku);
      var noIndex = parseGridNoToIndex(kuwaku.no);
      viewerCandidates.push({
        record: record,
        drawable: drawable,
        altitudeM: altitudeM,
        altitudeEstimated: altitudeEstimated,
        grid: {
          kuwaku: buildKuwaku(kuwaku.headA, kuwaku.headB, kuwaku.block, kuwaku.no),
          headB: kuwaku.headB,
          block: kuwaku.block,
          no: kuwaku.no,
          xIndex: xIndex,
          noIndex: noIndex
        }
      });
    }
  } catch (err) {
    _iterator4.e(err);
  } finally {
    _iterator4.f();
  }
  viewerStatus.textContent = "\u5BFE\u8C61 ".concat(detailSubRecords.length, "\u4EF6 / 3D\u8868\u793A ").concat(viewerCandidates.length, "\u4EF6 / \u5E73\u9762\u4F4D\u7F6E\u672A\u8A18\u5165 ").concat(missingPositionCount, "\u4EF6 / \u6A19\u9AD8\u672A\u8A18\u5165 ").concat(missingAltitudeCount, "\u4EF6\uFF08655m\u3067\u4EEE\u8868\u793A\uFF09 / \u7E26\u30B9\u30B1\u30FC\u30EB ").concat(viewerVerticalScale.toFixed(1), "x");
  if (!viewerCandidates.length) {
    clearViewerScene();
    hideViewerFallbackPanel();
    return;
  }
  if (!isViewerTabActive() && !viewer3d.initialized) {
    return;
  }
  if (!ensureViewerInitialized()) {
    showViewerFallbackPanel(viewerStatus === null || viewerStatus === void 0 ? void 0 : viewerStatus.textContent, {
      kuwakuLabel: kuwakuLabel,
      totalCount: detailSubRecords.length,
      drawableCount: viewerCandidates.length
    });
    return;
  }
  hideViewerFallbackPanel();
  var metrics = buildViewerGridMetrics(viewerCandidates);
  var shapes = viewerCandidates.map(function (candidate) {
    return buildViewerShapeFromCandidate(candidate, metrics);
  }).filter(Boolean);
  renderViewerScene(shapes, metrics, {
    preserveCamera: preserveCamera
  });
}
function isViewerTabActive() {
  return getActiveTabId() === "viewer-tab";
}
function syncViewerKuwakuSelect(records) {
  if (!viewerKuwakuSelect) {
    return;
  }
  var options = collectOutputKuwakuOptions(records);
  if (!options.some(function (item) {
    return item.value === selectedViewerKuwaku;
  })) {
    selectedViewerKuwaku = ALL_GRIDS_VALUE;
  }
  viewerKuwakuSelect.innerHTML = options.map(function (item) {
    return "<option value=\"".concat(escapeHtml(item.value), "\" ").concat(item.value === selectedViewerKuwaku ? "selected" : "", ">").concat(escapeHtml(item.label), "</option>");
  }).join("");
}
function parseGridNoToIndex(noRaw) {
  var no = value(noRaw);
  if (/^-?\d+$/.test(no)) {
    return Number(no);
  }
  if (!no) {
    return 0;
  }
  return hashText(no) % 100;
}
function buildViewerGridMetrics(candidates) {
  var xIndexes = candidates.map(function (item) {
    return item.grid.xIndex;
  }).filter(function (num) {
    return Number.isFinite(num);
  });
  var noIndexes = candidates.map(function (item) {
    return item.grid.noIndex;
  }).filter(function (num) {
    return Number.isFinite(num);
  });
  var altitudes = candidates.map(function (item) {
    return item.altitudeM;
  }).filter(function (num) {
    return Number.isFinite(num);
  });
  var minXIndex = xIndexes.length ? Math.min.apply(Math, _toConsumableArray(xIndexes)) : 0;
  var maxXIndex = xIndexes.length ? Math.max.apply(Math, _toConsumableArray(xIndexes)) : minXIndex;
  var presentHeads = new Set(candidates.map(function (item) {
    var _item$grid;
    return normalizeViewerHead(item === null || item === void 0 || (_item$grid = item.grid) === null || _item$grid === void 0 ? void 0 : _item$grid.headB);
  }));
  presentHeads.forEach(function (head) {
    var headIndex = VIEWER_HEAD_INDEX_MAP.get(head);
    if (!Number.isFinite(headIndex)) {
      return;
    }
    var fIndex = headIndex * 26 + 5;
    if (fIndex > maxXIndex) {
      maxXIndex = fIndex;
    }
  });
  var minNo = noIndexes.length ? Math.min.apply(Math, _toConsumableArray(noIndexes)) : 0;
  var maxNo = noIndexes.length ? Math.max.apply(Math, _toConsumableArray(noIndexes)) : minNo;
  var minZ = VIEWER_ALTITUDE_BASE_M;
  var observedMaxZ = altitudes.length ? Math.max.apply(Math, _toConsumableArray(altitudes)) : minZ + 1;
  var maxZ = Math.max(minZ + 1, Math.ceil(observedMaxZ));
  return {
    minXIndex: minXIndex,
    maxXIndex: maxXIndex,
    minNo: minNo,
    maxNo: maxNo,
    minZ: minZ,
    maxZ: maxZ,
    gridWidthM: Math.max(4, (maxXIndex - minXIndex + 1) * 4),
    gridHeightM: Math.max(4, (maxNo - minNo + 1) * 4)
  };
}
function normalizeAzimuth360(valueRaw) {
  var valueNum = Number(valueRaw);
  if (!Number.isFinite(valueNum)) {
    return null;
  }
  return (valueNum % 360 + 360) % 360;
}
function angularDistanceDeg(aRaw, bRaw) {
  var a = normalizeAzimuth360(aRaw);
  var b = normalizeAzimuth360(bRaw);
  if (!Number.isFinite(a) || !Number.isFinite(b)) {
    return Infinity;
  }
  var diff = Math.abs(a - b) % 360;
  return diff > 180 ? 360 - diff : diff;
}
function azimuthToViewerHorizontalUnit(azimuthRaw) {
  var azimuth = Number(azimuthRaw);
  if (!Number.isFinite(azimuth)) {
    return {
      x: 0,
      y: 1
    };
  }
  var rad = azimuth * Math.PI / 180;
  return {
    x: Math.sin(rad),
    y: Math.cos(rad)
  };
}
function resolvePlaneTiltParameters(record) {
  var strikeAzimuth = parseLargeAxisAzimuth(record === null || record === void 0 ? void 0 : record.planeStrikeDirection);
  var dipDeg = parseLargeAxisPlungeDeg(record === null || record === void 0 ? void 0 : record.planeDipDeg);
  var dipDirRaw = parseCompass8Azimuth(record === null || record === void 0 ? void 0 : record.planeDipDir8);
  if (!Number.isFinite(dipDeg) || dipDeg <= 0) {
    return {
      strikeAzimuth: Number.isFinite(strikeAzimuth) ? strikeAzimuth : null,
      dipAzimuth: null,
      dipDeg: null
    };
  }
  if (Number.isFinite(strikeAzimuth)) {
    var rightDip = (strikeAzimuth + 90) % 360;
    var leftDip = (strikeAzimuth + 270) % 360;
    if (!Number.isFinite(dipDirRaw)) {
      return {
        strikeAzimuth: strikeAzimuth,
        dipAzimuth: rightDip,
        dipDeg: dipDeg
      };
    }
    var rightDist = angularDistanceDeg(dipDirRaw, rightDip);
    var leftDist = angularDistanceDeg(dipDirRaw, leftDip);
    return {
      strikeAzimuth: strikeAzimuth,
      dipAzimuth: rightDist <= leftDist ? rightDip : leftDip,
      dipDeg: dipDeg
    };
  }
  return {
    strikeAzimuth: null,
    dipAzimuth: Number.isFinite(dipDirRaw) ? dipDirRaw : null,
    dipDeg: dipDeg
  };
}
function buildViewerShapeFromCandidate(candidate, metrics) {
  var _candidate$record4;
  var drawable = candidate.drawable;
  var altitudeM = candidate.altitudeM;
  var centerPlanPoint = {
    x: drawable.x,
    y: drawable.y
  };
  var tiltAzimuth = null;
  var tiltDeg = null;
  var lineDirectionAzimuth = null;
  var linePlungeDeg = null;
  var lineDownwardAzimuth = null;
  var planeStrikeAzimuth = null;
  var planeDipAzimuth = null;
  var planeDipDeg = null;
  if (drawable.type === "line") {
    var _candidate$record, _candidate$record2, _candidate$record3;
    lineDirectionAzimuth = parseLargeAxisAzimuth(candidate === null || candidate === void 0 || (_candidate$record = candidate.record) === null || _candidate$record === void 0 ? void 0 : _candidate$record.largeAxisDirection);
    linePlungeDeg = parseLargeAxisPlungeDeg(candidate === null || candidate === void 0 || (_candidate$record2 = candidate.record) === null || _candidate$record2 === void 0 ? void 0 : _candidate$record2.largeAxisPlungeDeg);
    lineDownwardAzimuth = parseCompass8Azimuth(candidate === null || candidate === void 0 || (_candidate$record3 = candidate.record) === null || _candidate$record3 === void 0 ? void 0 : _candidate$record3.largeAxisPlungeDir8);
    tiltAzimuth = lineDownwardAzimuth !== null && lineDownwardAzimuth !== void 0 ? lineDownwardAzimuth : lineDirectionAzimuth;
    tiltDeg = linePlungeDeg;
  } else if (drawable.type === "rect" || drawable.type === "ellipse" || drawable.type === "imageQuad") {
    var planeTilt = resolvePlaneTiltParameters(candidate === null || candidate === void 0 ? void 0 : candidate.record);
    planeStrikeAzimuth = planeTilt.strikeAzimuth;
    planeDipAzimuth = planeTilt.dipAzimuth;
    planeDipDeg = planeTilt.dipDeg;
    tiltAzimuth = planeDipAzimuth;
    tiltDeg = planeDipDeg;
  }
  var getViewerZForPlanPoint = function getViewerZForPlanPoint(planPointRaw) {
    var planPoint = planPointRaw || centerPlanPoint;
    var deltaM = computeViewerPlungeDeltaM(planPoint, centerPlanPoint, tiltAzimuth, tiltDeg);
    return applyViewerVerticalScale(altitudeM + deltaM, metrics.minZ);
  };
  var getViewerZForImagePlanPoint = function getViewerZForImagePlanPoint(planPointRaw) {
    var planPoint = planPointRaw || centerPlanPoint;
    var rawDeltaM = computeViewerPlungeDeltaM(planPoint, centerPlanPoint, tiltAzimuth, tiltDeg);
    var deltaM = Number.isFinite(rawDeltaM) ? clamp(rawDeltaM * IMAGE_QUAD_TILT_Z_SCALE, -IMAGE_QUAD_TILT_Z_LIMIT_M, IMAGE_QUAD_TILT_Z_LIMIT_M) : 0;
    var baseZ = applyViewerVerticalScale(altitudeM, metrics.minZ);
    return baseZ + deltaM;
  };
  var altitudeZ = getViewerZForPlanPoint(centerPlanPoint);
  var directBottomAltitudeEnabled = normalizeToggleFlag(candidate === null || candidate === void 0 || (_candidate$record4 = candidate.record) === null || _candidate$record4 === void 0 ? void 0 : _candidate$record4.altitudeInputEnabled) === "1";
  var bottomTargetZ = applyViewerVerticalScale(altitudeM, metrics.minZ);
  var worldCenter = convertViewerPointCmToWorld(drawable.x, drawable.y, candidate.grid, metrics);
  var meta = {
    id: value(candidate.record.id),
    label: value(candidate.record.specimenNo),
    nameMemo: value(candidate.record.nameMemo),
    unit: value(candidate.record.unit),
    detail: buildDetailText(candidate.record.detail, candidate.record.detailSub),
    kuwaku: value(candidate.grid.kuwaku),
    altitudeM: altitudeM,
    altitudeEstimated: Boolean(candidate.altitudeEstimated),
    color: drawable.color
  };
  if (drawable.type === "point") {
    return _objectSpread({
      type: "point",
      x: worldCenter.x,
      y: worldCenter.y,
      z: altitudeZ
    }, meta);
  }
  if (drawable.type === "multipoint") {
    var _points$;
    var points = (drawable.points || []).map(function (point) {
      var world = convertViewerPointCmToWorld(point.x, point.y, candidate.grid, metrics);
      return {
        x: world.x,
        y: world.y,
        z: getViewerZForPlanPoint(point)
      };
    }).filter(function (point) {
      return Number.isFinite(point.x) && Number.isFinite(point.y) && Number.isFinite(point.z);
    });
    if (!points.length) {
      return null;
    }
    if (directBottomAltitudeEnabled) {
      var anchored = anchorViewerPointsToBottomAltitude(points, bottomTargetZ);
      points = anchored.points;
    }
    var hull = buildHullPointsFromSource(points, (_points$ = points[0]) === null || _points$ === void 0 ? void 0 : _points$.z);
    var centerZ = points.reduce(function (sum, point) {
      return sum + point.z;
    }, 0) / points.length;
    return _objectSpread({
      type: "multipoint",
      points: points,
      hull: hull,
      x: worldCenter.x,
      y: worldCenter.y,
      z: centerZ
    }, meta);
  }
  if (drawable.type === "line") {
    var _parseDistanceToCm, _candidate$record5;
    var viewerScale = normalizeViewerVerticalScale(viewerVerticalScale);
    var lineLengthCm = (_parseDistanceToCm = parseDistanceToCm(candidate === null || candidate === void 0 || (_candidate$record5 = candidate.record) === null || _candidate$record5 === void 0 ? void 0 : _candidate$record5.lineLengthCm)) !== null && _parseDistanceToCm !== void 0 ? _parseDistanceToCm : Math.hypot(drawable.x2 - drawable.x1, drawable.y2 - drawable.y1);
    var halfLengthM = Math.max(0, lineLengthCm / 200);
    var axisAzimuth = normalizeAzimuth360(lineDirectionAzimuth !== null && lineDirectionAzimuth !== void 0 ? lineDirectionAzimuth : tiltAzimuth);
    if (!Number.isFinite(axisAzimuth)) {
      axisAzimuth = 0;
    }
    var downwardAzimuth = normalizeAzimuth360(lineDownwardAzimuth);
    if (Number.isFinite(downwardAzimuth)) {
      var forwardDist = angularDistanceDeg(downwardAzimuth, axisAzimuth);
      if (forwardDist > 90) {
        axisAzimuth = (axisAzimuth + 180) % 360;
      }
    }
    var plunge = Number.isFinite(linePlungeDeg) ? clamp(linePlungeDeg, 0, 90) : 0;
    var plungeRad = plunge * Math.PI / 180;
    var horizFactor = Math.cos(plungeRad);
    var verticalFactor = Math.sin(plungeRad);
    var axisUnit = azimuthToViewerHorizontalUnit(axisAzimuth);
    var vx = axisUnit.x * horizFactor;
    var vy = axisUnit.y * horizFactor;
    var vz = -verticalFactor * viewerScale;
    var linePoints = [{
      x: worldCenter.x - vx * halfLengthM,
      y: worldCenter.y - vy * halfLengthM,
      z: altitudeZ - vz * halfLengthM
    }, {
      x: worldCenter.x + vx * halfLengthM,
      y: worldCenter.y + vy * halfLengthM,
      z: altitudeZ + vz * halfLengthM
    }];
    var lineCenterZ = altitudeZ;
    if (directBottomAltitudeEnabled) {
      var _anchored = anchorViewerPointsToBottomAltitude(linePoints, bottomTargetZ);
      linePoints = _anchored.points;
      if (_anchored.minZ != null && _anchored.maxZ != null) {
        lineCenterZ = (_anchored.minZ + _anchored.maxZ) / 2;
      }
    }
    return _objectSpread({
      type: "line",
      points: linePoints,
      x: worldCenter.x,
      y: worldCenter.y,
      z: lineCenterZ
    }, meta);
  }
  if (drawable.type === "rect") {
    var halfW = drawable.width / 2;
    var halfH = drawable.height / 2;
    var localCorners = [{
      x: drawable.x - halfW,
      y: drawable.y - halfH
    }, {
      x: drawable.x + halfW,
      y: drawable.y - halfH
    }, {
      x: drawable.x + halfW,
      y: drawable.y + halfH
    }, {
      x: drawable.x - halfW,
      y: drawable.y + halfH
    }].map(function (point) {
      return rotatePlanPoint(point, {
        x: drawable.x,
        y: drawable.y
      }, drawable.rotationDeg);
    });
    var _points = localCorners.map(function (point) {
      var world = convertViewerPointCmToWorld(point.x, point.y, candidate.grid, metrics);
      return {
        x: world.x,
        y: world.y,
        z: getViewerZForPlanPoint(point)
      };
    });
    if (directBottomAltitudeEnabled) {
      var _anchored2 = anchorViewerPointsToBottomAltitude(_points, bottomTargetZ);
      _points = _anchored2.points;
    }
    _points.push(_points[0]);
    return _objectSpread({
      type: "polyline",
      points: _points,
      x: worldCenter.x,
      y: worldCenter.y,
      z: directBottomAltitudeEnabled && _points.length ? _points.reduce(function (sum, p) {
        return sum + p.z;
      }, 0) / _points.length : altitudeZ
    }, meta);
  }
  if (drawable.type === "ellipse") {
    var _points2 = [];
    var segmentCount = 48;
    for (var i = 0; i <= segmentCount; i += 1) {
      var theta = i / segmentCount * Math.PI * 2;
      var local = {
        x: drawable.x + Math.cos(theta) * drawable.rx,
        y: drawable.y + Math.sin(theta) * drawable.ry
      };
      var rotated = rotatePlanPoint(local, {
        x: drawable.x,
        y: drawable.y
      }, drawable.rotationDeg);
      var world = convertViewerPointCmToWorld(rotated.x, rotated.y, candidate.grid, metrics);
      _points2.push({
        x: world.x,
        y: world.y,
        z: getViewerZForPlanPoint(rotated)
      });
    }
    if (directBottomAltitudeEnabled) {
      var _anchored3 = anchorViewerPointsToBottomAltitude(_points2, bottomTargetZ);
      _points2 = _anchored3.points;
    }
    return _objectSpread({
      type: "polyline",
      points: _points2,
      x: worldCenter.x,
      y: worldCenter.y,
      z: directBottomAltitudeEnabled && _points2.length ? _points2.reduce(function (sum, p) {
        return sum + p.z;
      }, 0) / _points2.length : altitudeZ
    }, meta);
  }
  if (drawable.type === "imageQuad") {
    var _viewerScale = normalizeViewerVerticalScale(viewerVerticalScale);
    var dipDeg = Number.isFinite(planeDipDeg) ? clamp(planeDipDeg, 0, 90) : 0;
    var dipAzimuth = normalizeAzimuth360(planeDipAzimuth !== null && planeDipAzimuth !== void 0 ? planeDipAzimuth : Number.isFinite(planeStrikeAzimuth) ? (planeStrikeAzimuth + 90) % 360 : null);
    var strikeAzimuth = normalizeAzimuth360(planeStrikeAzimuth !== null && planeStrikeAzimuth !== void 0 ? planeStrikeAzimuth : Number.isFinite(dipAzimuth) ? (dipAzimuth + 270) % 360 : null);
    var canUsePlaneProjection = Number.isFinite(dipAzimuth) && Number.isFinite(strikeAzimuth) && dipDeg > 0;
    var strikeUnit = azimuthToViewerHorizontalUnit(strikeAzimuth);
    var dipUnit = azimuthToViewerHorizontalUnit(dipAzimuth);
    var dipRad = dipDeg * Math.PI / 180;
    var dipHorizFactor = Math.cos(dipRad);
    var dipVerticalFactor = Math.sin(dipRad);
    var _points3 = (drawable.points || []).map(function (point) {
      var world = convertViewerPointCmToWorld(point.x, point.y, candidate.grid, metrics);
      if (!canUsePlaneProjection) {
        return {
          x: world.x,
          y: world.y,
          z: getViewerZForImagePlanPoint(point)
        };
      }
      var relX = world.x - worldCenter.x;
      var relY = world.y - worldCenter.y;
      var strikeComp = relX * strikeUnit.x + relY * strikeUnit.y;
      var dipComp = relX * dipUnit.x + relY * dipUnit.y;
      return {
        x: worldCenter.x + strikeUnit.x * strikeComp + dipUnit.x * dipComp * dipHorizFactor,
        y: worldCenter.y + strikeUnit.y * strikeComp + dipUnit.y * dipComp * dipHorizFactor,
        z: altitudeZ - dipComp * dipVerticalFactor * _viewerScale
      };
    });
    if (directBottomAltitudeEnabled) {
      var _anchored4 = anchorViewerPointsToBottomAltitude(_points3, bottomTargetZ);
      _points3 = _anchored4.points;
    }
    return _objectSpread({
      type: "imageQuad",
      points: _points3,
      imageType: value(drawable.imageType),
      imagePath: value(drawable.imagePath),
      useOriginalImageColor: Boolean(drawable.useOriginalImageColor),
      x: worldCenter.x,
      y: worldCenter.y,
      z: directBottomAltitudeEnabled && _points3.length ? _points3.reduce(function (sum, p) {
        return sum + p.z;
      }, 0) / _points3.length : altitudeZ
    }, meta);
  }
  return null;
}
function anchorViewerPointsToBottomAltitude(pointsRaw, targetBottomZRaw) {
  var points = Array.isArray(pointsRaw) ? pointsRaw : [];
  var targetBottomZ = Number(targetBottomZRaw);
  if (!points.length || !Number.isFinite(targetBottomZ)) {
    return {
      points: points,
      minZ: null,
      maxZ: null
    };
  }
  var minZ = Infinity;
  var maxZ = -Infinity;
  points.forEach(function (point) {
    var z = Number(point === null || point === void 0 ? void 0 : point.z);
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
    return {
      points: points,
      minZ: null,
      maxZ: null
    };
  }
  var deltaZ = targetBottomZ - minZ;
  if (Math.abs(deltaZ) < 1e-9) {
    return {
      points: points,
      minZ: minZ,
      maxZ: maxZ
    };
  }
  var shifted = points.map(function (point) {
    return {
      x: point.x,
      y: point.y,
      z: Number.isFinite(Number(point === null || point === void 0 ? void 0 : point.z)) ? Number(point.z) + deltaZ : point.z
    };
  });
  return {
    points: shifted,
    minZ: minZ + deltaZ,
    maxZ: maxZ + deltaZ
  };
}
function rotatePlanPoint(point, center, rotationDegRaw) {
  var rotationDeg = Number(rotationDegRaw);
  if (!Number.isFinite(rotationDeg) || Math.abs(rotationDeg) < 1e-6) {
    return {
      x: point.x,
      y: point.y
    };
  }
  var rad = rotationDeg * Math.PI / 180;
  var cos = Math.cos(rad);
  var sin = Math.sin(rad);
  var dx = point.x - center.x;
  var dy = point.y - center.y;
  return {
    x: center.x + dx * cos - dy * sin,
    y: center.y + dx * sin + dy * cos
  };
}
function computeViewerPlungeDeltaM(point, center, axisAzimuthRaw, plungeDegRaw) {
  if (!point || !center) {
    return 0;
  }
  var axisAzimuth = Number(axisAzimuthRaw);
  var plungeDeg = Number(plungeDegRaw);
  if (!Number.isFinite(axisAzimuth) || !Number.isFinite(plungeDeg) || plungeDeg <= 0) {
    return 0;
  }
  var unit = azimuthToPlanUnitVector(axisAzimuth);
  var alongAxisCm = (point.x - center.x) * unit.dx + (point.y - center.y) * unit.dy;
  var alongAxisM = alongAxisCm / 100;
  var tilt = Math.tan(plungeDeg * Math.PI / 180);
  if (!Number.isFinite(tilt) || tilt === 0) {
    return 0;
  }
  return -alongAxisM * tilt;
}
function convertViewerPointCmToWorld(xCmRaw, yCmRaw, grid, metrics) {
  var xCm = Number(xCmRaw);
  var yCm = Number(yCmRaw);
  var xIndex = Number(grid === null || grid === void 0 ? void 0 : grid.xIndex);
  var noIndex = Number(grid === null || grid === void 0 ? void 0 : grid.noIndex);
  var baseEast = (xIndex - metrics.minXIndex) * 4;
  var baseNorth = (metrics.maxNo - noIndex) * 4;
  return {
    x: baseEast + xCm / 100,
    y: baseNorth + (PLAN_SIZE_CM - yCm) / 100
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
    return {
      renderer: null,
      error: new Error("THREE未読込")
    };
  }
  var createRendererWithContext = function createRendererWithContext(contextType, contextAttrs) {
    var canvas = document.createElement("canvas");
    var gl = canvas.getContext(contextType, contextAttrs) || (contextType === "webgl" ? canvas.getContext("experimental-webgl", contextAttrs) : null);
    if (!gl) {
      throw new Error("".concat(contextType, " context unavailable"));
    }
    return new THREE.WebGLRenderer({
      canvas: canvas,
      context: gl,
      antialias: false,
      alpha: Boolean(contextAttrs === null || contextAttrs === void 0 ? void 0 : contextAttrs.alpha),
      powerPreference: "low-power",
      precision: "mediump"
    });
  };
  var attempts = [function () {
    return new THREE.WebGLRenderer({
      antialias: false,
      alpha: false,
      stencil: false,
      powerPreference: "low-power",
      precision: "mediump"
    });
  }, function () {
    return createRendererWithContext("webgl2", {
      antialias: false,
      alpha: false,
      depth: true,
      stencil: false
    });
  }, function () {
    return createRendererWithContext("webgl", {
      antialias: false,
      alpha: false,
      depth: true,
      stencil: false
    });
  }, function () {
    return createRendererWithContext("webgl", {
      antialias: false,
      alpha: true,
      depth: true,
      stencil: false
    });
  }];
  if (typeof THREE.WebGL1Renderer === "function") {
    attempts.push(function () {
      return new THREE.WebGL1Renderer({
        antialias: false,
        alpha: false,
        stencil: false
      });
    });
  }
  var lastError = null;
  for (var _i = 0, _attempts = attempts; _i < _attempts.length; _i++) {
    var makeRenderer = _attempts[_i];
    try {
      var renderer = makeRenderer();
      if (renderer) {
        return {
          renderer: renderer,
          error: null
        };
      }
    } catch (error) {
      lastError = error;
    }
  }
  var probeCanvas = document.createElement("canvas");
  var hasWebGl = Boolean(probeCanvas.getContext("webgl")) || Boolean(probeCanvas.getContext("experimental-webgl")) || Boolean(probeCanvas.getContext("webgl2"));
  if (!hasWebGl) {
    lastError = new Error("この端末・ブラウザでWebGLが利用できません");
  }
  return {
    renderer: null,
    error: lastError || new Error("WebGLRenderer生成失敗")
  };
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
    var _viewer3d$renderer;
    viewer3d.scene = new THREE.Scene();
    viewer3d.scene.background = new THREE.Color(0xf8fafc);
    viewer3d.camera = new THREE.PerspectiveCamera(34, 1, 0.1, 10000);
    var rendererResult = createViewerRendererWithFallback();
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
    viewer3d.controls = typeof THREE.OrbitControls === "function" ? new THREE.OrbitControls(viewer3d.camera, viewer3d.renderer.domElement) : null;
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
          RIGHT: null
        };
        viewer3d.defaultLeftMouseAction = THREE.MOUSE.ROTATE;
      }
      viewer3d.controls.target.set(0, 0, 0);
    }
    var ambient = new THREE.AmbientLight(0xffffff, 0.88);
    var directional = new THREE.DirectionalLight(0xffffff, 0.55);
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
    if ((_viewer3d$renderer = viewer3d.renderer) !== null && _viewer3d$renderer !== void 0 && _viewer3d$renderer.domElement) {
      viewer3d.renderer.domElement.addEventListener("pointerdown", handleViewerControlPointerDown);
      viewer3d.renderer.domElement.addEventListener("pointerup", handleViewerControlPointerUp);
      viewer3d.renderer.domElement.addEventListener("pointercancel", handleViewerControlPointerUp);
    }
    window.addEventListener("pointerup", handleViewerControlPointerUp);
    window.addEventListener("blur", handleViewerControlPointerUp);
    if (typeof ResizeObserver === "function") {
      viewer3d.resizeObserver = new ResizeObserver(function () {
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
      } catch (_disposeError) {}
      var canvas = viewer3d.renderer.domElement;
      if (canvas !== null && canvas !== void 0 && canvas.parentNode) {
        canvas.parentNode.removeChild(canvas);
      }
      viewer3d.renderer = null;
    }
    if (viewerStatus) {
      var reason = value(error === null || error === void 0 ? void 0 : error.message);
      viewerStatus.textContent = reason ? "3D\u8868\u793A\u306E\u521D\u671F\u5316\u306B\u5931\u6557\u3057\u307E\u3057\u305F\uFF08".concat(reason, "\uFF09\u3002") : "3D表示の初期化に失敗しました。";
    }
    return false;
  }
}
function ensureViewerCanvasSize() {
  if (!viewer3d.initialized || !viewerCanvasWrap || !viewer3d.renderer || !viewer3d.camera) {
    return;
  }
  var rect = viewerCanvasWrap.getBoundingClientRect();
  var width = Math.max(16, Math.floor(rect.width));
  var height = Math.max(16, Math.floor(rect.height));
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
  viewer3d.dataGroup.children.forEach(function (child) {
    return disposeViewerObject3D(child);
  });
  viewer3d.labelGroup.children.forEach(function (child) {
    return disposeViewerObject3D(child);
  });
  viewer3d.gridGroup.children.forEach(function (child) {
    return disposeViewerObject3D(child);
  });
  viewer3d.dataGroup.clear();
  viewer3d.labelGroup.clear();
  viewer3d.gridGroup.clear();
  viewer3d.pickMeshes = [];
  viewer3d.meshesByRecordId = new Map();
  hideViewerTooltip();
}
function disposeViewerObject3D(object) {
  var _object$traverse;
  if (!object) {
    return;
  }
  (_object$traverse = object.traverse) === null || _object$traverse === void 0 || _object$traverse.call(object, function (child) {
    var _child$geometry;
    if ((_child$geometry = child.geometry) !== null && _child$geometry !== void 0 && _child$geometry.dispose) {
      child.geometry.dispose();
    }
    if (child.material) {
      var materials = Array.isArray(child.material) ? child.material : [child.material];
      materials.forEach(function (material) {
        var _material$map;
        if ((_material$map = material.map) !== null && _material$map !== void 0 && _material$map.dispose) {
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
  var _camera$position, _camera$position2, _camera$position3, _camera$up, _camera$up2, _camera$up3, _viewer3d$controls;
  if (!viewer3d.initialized || !viewer3d.camera || !viewer3d.bounds) {
    return null;
  }
  var camera = viewer3d.camera;
  var position = {
    x: Number((_camera$position = camera.position) === null || _camera$position === void 0 ? void 0 : _camera$position.x),
    y: Number((_camera$position2 = camera.position) === null || _camera$position2 === void 0 ? void 0 : _camera$position2.y),
    z: Number((_camera$position3 = camera.position) === null || _camera$position3 === void 0 ? void 0 : _camera$position3.z)
  };
  if (!Number.isFinite(position.x) || !Number.isFinite(position.y) || !Number.isFinite(position.z)) {
    return null;
  }
  var up = {
    x: Number((_camera$up = camera.up) === null || _camera$up === void 0 ? void 0 : _camera$up.x),
    y: Number((_camera$up2 = camera.up) === null || _camera$up2 === void 0 ? void 0 : _camera$up2.y),
    z: Number((_camera$up3 = camera.up) === null || _camera$up3 === void 0 ? void 0 : _camera$up3.z)
  };
  var targetVec = (_viewer3d$controls = viewer3d.controls) === null || _viewer3d$controls === void 0 ? void 0 : _viewer3d$controls.target;
  var target = {
    x: Number(targetVec === null || targetVec === void 0 ? void 0 : targetVec.x),
    y: Number(targetVec === null || targetVec === void 0 ? void 0 : targetVec.y),
    z: Number(targetVec === null || targetVec === void 0 ? void 0 : targetVec.z)
  };
  if (!Number.isFinite(target.x) || !Number.isFinite(target.y) || !Number.isFinite(target.z)) {
    return null;
  }
  var dx = position.x - target.x;
  var dy = position.y - target.y;
  var dz = position.z - target.z;
  var distance = Math.hypot(dx, dy, dz);
  if (!Number.isFinite(distance) || distance < 0.5) {
    return null;
  }
  return {
    position: position,
    up: Number.isFinite(up.x) && Number.isFinite(up.y) && Number.isFinite(up.z) ? up : {
      x: 0,
      y: 0,
      z: 1
    },
    target: target
  };
}
function restoreViewerCameraState(viewState) {
  if (!viewState || !viewer3d.initialized || !viewer3d.camera) {
    return false;
  }
  var position = viewState.position || {};
  var up = viewState.up || {};
  var target = viewState.target || {};
  if (!Number.isFinite(position.x) || !Number.isFinite(position.y) || !Number.isFinite(position.z) || !Number.isFinite(target.x) || !Number.isFinite(target.y) || !Number.isFinite(target.z)) {
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
  var target = viewState.target || {};
  var position = viewState.position || {};
  if (!Number.isFinite(target.x) || !Number.isFinite(target.y) || !Number.isFinite(target.z) || !Number.isFinite(position.x) || !Number.isFinite(position.y) || !Number.isFinite(position.z)) {
    return false;
  }
  var spanX = Math.max(1, Number(bounds.maxX) - Number(bounds.minX));
  var spanY = Math.max(1, Number(bounds.maxY) - Number(bounds.minY));
  var spanZ = Math.max(1, Number(bounds.maxZ) - Number(bounds.minZ));
  var marginFactor = 3.5;
  var within = function within(valueNum, minNum, maxNum, span) {
    return valueNum >= minNum - span * marginFactor && valueNum <= maxNum + span * marginFactor;
  };
  if (!within(target.x, bounds.minX, bounds.maxX, spanX) || !within(target.y, bounds.minY, bounds.maxY, spanY) || !within(target.z, bounds.minZ, bounds.maxZ, spanZ)) {
    return false;
  }
  if (!within(position.x, bounds.minX, bounds.maxX, spanX) || !within(position.y, bounds.minY, bounds.maxY, spanY) || !within(position.z, bounds.minZ, bounds.maxZ, spanZ)) {
    return false;
  }
  return true;
}
function renderViewerScene(shapes, metrics) {
  var options = arguments.length > 2 && arguments[2] !== undefined ? arguments[2] : {};
  if (!viewer3d.initialized || !viewer3d.dataGroup || !viewer3d.labelGroup || !viewer3d.gridGroup) {
    return;
  }
  var preserveCamera = Boolean(options === null || options === void 0 ? void 0 : options.preserveCamera);
  var previousViewState = preserveCamera ? captureViewerCameraState() : null;
  clearViewerScene();
  renderViewerGrid(metrics);
  viewer3d.renderNonce += 1;
  var renderNonce = viewer3d.renderNonce;
  shapes.forEach(function (shape) {
    var color = new THREE.Color(shape.color || "#6b7280");
    if (shape.type === "point") {
      var pointMesh = new THREE.Mesh(new THREE.SphereGeometry(0.11, 14, 14), new THREE.MeshBasicMaterial({
        color: color
      }));
      pointMesh.position.set(shape.x, shape.y, shape.z);
      viewer3d.dataGroup.add(pointMesh);
    } else if (shape.type === "multipoint") {
      (shape.points || []).forEach(function (point) {
        var pointMesh = new THREE.Mesh(new THREE.SphereGeometry(0.085, 12, 12), new THREE.MeshBasicMaterial({
          color: color
        }));
        pointMesh.position.set(point.x, point.y, point.z);
        viewer3d.dataGroup.add(pointMesh);
      });
      var hull = Array.isArray(shape.hull) ? shape.hull : [];
      if (hull.length === 2) {
        renderViewerSegment(hull[0], hull[1], color, 0.025);
      } else if (hull.length >= 3) {
        for (var i = 0; i < hull.length; i += 1) {
          var start = hull[i];
          var end = hull[(i + 1) % hull.length];
          renderViewerSegment(start, end, color, 0.025);
        }
      }
    } else if (shape.type === "line") {
      renderViewerSegment(shape.points[0], shape.points[1], color, 0.05);
    } else if (shape.type === "polyline") {
      for (var _i2 = 0; _i2 < shape.points.length - 1; _i2 += 1) {
        renderViewerSegment(shape.points[_i2], shape.points[_i2 + 1], color, 0.04);
      }
    } else if (shape.type === "imageQuad") {
      renderViewerImageQuad(shape, renderNonce);
    }
    var label = createViewerTextSprite(shape.label || "-", shape.color);
    label.position.set(shape.x, shape.y, shape.z + 0.16);
    viewer3d.labelGroup.add(label);
    var pickTargets = shape.type === "multipoint" && Array.isArray(shape.points) && shape.points.length ? shape.points : [{
      x: shape.x,
      y: shape.y,
      z: shape.z
    }];
    pickTargets.forEach(function (targetPoint) {
      var pickMesh = new THREE.Mesh(new THREE.SphereGeometry(0.2, 10, 10), new THREE.MeshBasicMaterial({
        transparent: true,
        opacity: 0.001,
        depthWrite: false
      }));
      pickMesh.position.set(targetPoint.x, targetPoint.y, targetPoint.z);
      pickMesh.userData = {
        id: shape.id,
        label: shape.label,
        nameMemo: shape.nameMemo,
        unit: shape.unit,
        detail: shape.detail,
        kuwaku: shape.kuwaku,
        altitudeM: shape.altitudeM,
        altitudeEstimated: Boolean(shape.altitudeEstimated)
      };
      viewer3d.pickMeshes.push(pickMesh);
      viewer3d.dataGroup.add(pickMesh);
    });
    viewer3d.meshesByRecordId.set(shape.id, shape);
  });
  viewer3d.bounds = computeViewerBounds(shapes, metrics);
  var shouldRestoreCamera = preserveCamera && isViewerCameraStateCompatible(previousViewState, viewer3d.bounds) && restoreViewerCameraState(previousViewState);
  if (!shouldRestoreCamera) {
    applyViewerPerspective();
  }
}
function renderViewerGrid(metrics) {
  if (!viewer3d.gridGroup) {
    return;
  }
  var axisMinAltitude = Math.floor(metrics.minZ);
  var axisMaxAltitude = Math.max(axisMinAltitude + 1, Math.ceil(metrics.maxZ));
  var zBase = applyViewerVerticalScale(axisMinAltitude, metrics.minZ) - 0.05;
  var width = metrics.gridWidthM;
  var height = metrics.gridHeightM;
  var gridMat = new THREE.LineBasicMaterial({
    color: 0xcbd5e1
  });
  for (var x = 0; x <= width + 0.001; x += 4) {
    var geometry = new THREE.BufferGeometry().setFromPoints([new THREE.Vector3(x, 0, zBase), new THREE.Vector3(x, height, zBase)]);
    viewer3d.gridGroup.add(new THREE.Line(geometry, gridMat));
  }
  for (var y = 0; y <= height + 0.001; y += 4) {
    var _geometry = new THREE.BufferGeometry().setFromPoints([new THREE.Vector3(0, y, zBase), new THREE.Vector3(width, y, zBase)]);
    viewer3d.gridGroup.add(new THREE.Line(_geometry, gridMat));
  }
  var cardinalOffset = 1.18;
  var cardinalZ = zBase + 0.06;
  var cardinalLabels = [{
    text: "北",
    x: width / 2,
    y: height + cardinalOffset
  }, {
    text: "南",
    x: width / 2,
    y: -cardinalOffset
  }, {
    text: "東",
    x: width + cardinalOffset,
    y: height / 2
  }, {
    text: "西",
    x: -cardinalOffset,
    y: height / 2
  }];
  cardinalLabels.forEach(function (item) {
    var sprite = createViewerTextSprite(item.text, "#0f172a", {
      fontPx: 120,
      scaleX: 1.8,
      scaleY: 0.45
    });
    sprite.position.set(item.x, item.y, cardinalZ);
    viewer3d.gridGroup.add(sprite);
  });
  var colNorthY = height + 0.62;
  var colSouthY = -0.62;
  var rowWestX = -0.62;
  var rowEastX = width + 0.62;
  var colCount = Math.max(1, metrics.maxXIndex - metrics.minXIndex + 1);
  var rowCount = Math.max(1, metrics.maxNo - metrics.minNo + 1);
  var denseFactor = Math.max(colCount, rowCount);
  var frameLabelWidth = clamp(2.002 - denseFactor * 0.0208, 1.17, 1.664);
  var frameLabelHeight = clamp(0.598 - denseFactor * 0.00624, 0.312, 0.468);
  var stakeLabelWidth = clamp(1.3824 - denseFactor * 0.01296, 0.72, 1.0944);
  var stakeLabelHeight = clamp(0.432 - denseFactor * 0.004032, 0.2016, 0.3168);
  var labelPlaneZ = zBase + 0.004;
  for (var xIndex = metrics.minXIndex; xIndex <= metrics.maxXIndex; xIndex += 1) {
    var block = viewerXIndexToHeadBlock(xIndex).block;
    var centerX = (xIndex - metrics.minXIndex) * 4 + 2;
    var northLabel = createViewerTextPlane(block, "#1f3b35", {
      fontPx: 127,
      width: frameLabelWidth,
      height: frameLabelHeight
    });
    northLabel.position.set(centerX, colNorthY, labelPlaneZ);
    viewer3d.gridGroup.add(northLabel);
    var southLabel = createViewerTextPlane(block, "#1f3b35", {
      fontPx: 127,
      width: frameLabelWidth,
      height: frameLabelHeight
    });
    southLabel.position.set(centerX, colSouthY, labelPlaneZ);
    viewer3d.gridGroup.add(southLabel);
  }
  for (var no = metrics.minNo; no <= metrics.maxNo; no += 1) {
    var rowText = String(no);
    var centerY = (metrics.maxNo - no) * 4 + 2;
    var westLabel = createViewerTextPlane(rowText, "#1f3b35", {
      fontPx: 127,
      width: frameLabelWidth,
      height: frameLabelHeight
    });
    westLabel.position.set(rowWestX, centerY, labelPlaneZ);
    viewer3d.gridGroup.add(westLabel);
    var eastLabel = createViewerTextPlane(rowText, "#1f3b35", {
      fontPx: 127,
      width: frameLabelWidth,
      height: frameLabelHeight
    });
    eastLabel.position.set(rowEastX, centerY, labelPlaneZ);
    viewer3d.gridGroup.add(eastLabel);
  }
  var allCornerStakes = buildViewerAllGridCornerStakeLabels(metrics);
  allCornerStakes.forEach(function (corner) {
    var label = createViewerTextPlane(corner.label, "#97a7bc", {
      fontPx: 98,
      width: stakeLabelWidth,
      height: stakeLabelHeight
    });
    label.position.set(corner.x, corner.y, labelPlaneZ);
    viewer3d.gridGroup.add(label);
  });
  var zStart = applyViewerVerticalScale(axisMinAltitude, metrics.minZ);
  var zEnd = applyViewerVerticalScale(axisMaxAltitude, metrics.minZ);
  var zAxisMat = new THREE.LineBasicMaterial({
    color: 0x334155
  });
  var cornerAxes = [{
    x: 0,
    y: 0,
    dirX: -1,
    dirY: -1
  }, {
    x: width,
    y: 0,
    dirX: 1,
    dirY: -1
  }, {
    x: 0,
    y: height,
    dirX: -1,
    dirY: 1
  }, {
    x: width,
    y: height,
    dirX: 1,
    dirY: 1
  }];
  cornerAxes.forEach(function (corner) {
    var zAxisGeom = new THREE.BufferGeometry().setFromPoints([new THREE.Vector3(corner.x, corner.y, zStart), new THREE.Vector3(corner.x, corner.y, zEnd)]);
    viewer3d.gridGroup.add(new THREE.Line(zAxisGeom, zAxisMat));
    for (var altitude = axisMinAltitude; altitude <= axisMaxAltitude; altitude += 1) {
      var z = applyViewerVerticalScale(altitude, metrics.minZ);
      var tickGeom = new THREE.BufferGeometry().setFromPoints([new THREE.Vector3(corner.x, corner.y, z), new THREE.Vector3(corner.x + corner.dirX * 0.14, corner.y + corner.dirY * 0.14, z)]);
      viewer3d.gridGroup.add(new THREE.Line(tickGeom, zAxisMat));
      var label = createViewerTextSprite("".concat(altitude, "m"), "#334155", {
        fontPx: 87,
        scaleX: 1.57,
        scaleY: 0.39
      });
      label.position.set(corner.x + corner.dirX * 0.95, corner.y + corner.dirY * 0.95, z + 0.12);
      viewer3d.gridGroup.add(label);
    }
    var axisLabel = createViewerTextSprite("標高(m)", "#1e293b", {
      fontPx: 87,
      scaleX: 1.57,
      scaleY: 0.39
    });
    axisLabel.position.set(corner.x + corner.dirX * 1.5, corner.y + corner.dirY * 1.5, zEnd + 0.85);
    viewer3d.gridGroup.add(axisLabel);
  });
}
function buildViewerAllGridCornerStakeLabels(metrics) {
  var cols = Math.max(1, metrics.maxXIndex - metrics.minXIndex + 1);
  var rows = Math.max(1, metrics.maxNo - metrics.minNo + 1);
  var labels = [];
  var cellSize = 4;
  for (var xLine = 0; xLine <= cols; xLine += 1) {
    var block = incrementGridBlock(viewerXIndexToHeadBlock(metrics.minXIndex).block, xLine);
    var x = xLine * cellSize;
    for (var yLine = 0; yLine <= rows; yLine += 1) {
      var no = String(metrics.minNo + yLine);
      var y = metrics.gridHeightM - yLine * cellSize;
      labels.push({
        label: "".concat(block, "-").concat(no),
        x: x,
        y: y
      });
    }
  }
  return labels;
}
function renderViewerSegment(start, end, color) {
  var radius = arguments.length > 3 && arguments[3] !== undefined ? arguments[3] : 0.04;
  if (!viewer3d.dataGroup) {
    return;
  }
  var startVec = new THREE.Vector3(start.x, start.y, start.z);
  var endVec = new THREE.Vector3(end.x, end.y, end.z);
  var diff = new THREE.Vector3().subVectors(endVec, startVec);
  var length = diff.length();
  if (!Number.isFinite(length) || length <= 0.0001) {
    return;
  }
  var geometry = new THREE.CylinderGeometry(radius, radius, length, 8);
  var material = new THREE.MeshBasicMaterial({
    color: color
  });
  var mesh = new THREE.Mesh(geometry, material);
  var mid = new THREE.Vector3().addVectors(startVec, endVec).multiplyScalar(0.5);
  mesh.position.copy(mid);
  var axis = new THREE.Vector3(0, 1, 0);
  mesh.quaternion.setFromUnitVectors(axis, diff.clone().normalize());
  viewer3d.dataGroup.add(mesh);
}
function renderViewerImageQuad(shape, renderNonce) {
  if (!viewer3d.dataGroup) {
    return;
  }
  var targetGroup = viewer3d.dataGroup;
  var points = Array.isArray(shape === null || shape === void 0 ? void 0 : shape.points) ? shape.points : [];
  var imagePath = value(shape === null || shape === void 0 ? void 0 : shape.imagePath) || getLargeShapeImagePath(shape === null || shape === void 0 ? void 0 : shape.imageType);
  var useOriginalImageColor = Boolean(shape === null || shape === void 0 ? void 0 : shape.useOriginalImageColor);
  if (points.length !== 4) {
    return;
  }
  if (!useOriginalImageColor) {
    for (var i = 0; i < points.length; i += 1) {
      var start = points[i];
      var end = points[(i + 1) % points.length];
      renderViewerSegment(start, end, new THREE.Color((shape === null || shape === void 0 ? void 0 : shape.color) || "#6b7280"), 0.012);
    }
  }
  if (!imagePath) {
    return;
  }
  var buildGeometry = function buildGeometry() {
    var segments = 28;
    var geometry = new THREE.BufferGeometry();
    var positions = [];
    var uvs = [];
    var indices = [];
    for (var y = 0; y <= segments; y += 1) {
      var v = 1 - y / segments;
      for (var x = 0; x <= segments; x += 1) {
        var u = x / segments;
        var p = interpolateViewerQuadPoint(points, u, v);
        positions.push(p.x, p.y, p.z);
        uvs.push(u, v);
      }
    }
    for (var _y = 0; _y < segments; _y += 1) {
      for (var _x8 = 0; _x8 < segments; _x8 += 1) {
        var row = _y * (segments + 1);
        var nextRow = (_y + 1) * (segments + 1);
        var a = row + _x8;
        var b = row + _x8 + 1;
        var c = nextRow + _x8 + 1;
        var d = nextRow + _x8;
        indices.push(a, b, d, b, c, d);
      }
    }
    var vertices = new Float32Array(positions);
    var uvArray = new Float32Array(uvs);
    geometry.setAttribute("position", new THREE.BufferAttribute(vertices, 3));
    geometry.setAttribute("uv", new THREE.BufferAttribute(uvArray, 2));
    geometry.setIndex(indices);
    geometry.computeVertexNormals();
    return geometry;
  };
  var addTexturedMesh = function addTexturedMesh(texture) {
    var _window$THREE;
    var tintHex = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : "#ffffff";
    if (!viewer3d.initialized || !targetGroup || !targetGroup.parent) {
      return;
    }
    texture.needsUpdate = true;
    texture.minFilter = THREE.LinearFilter;
    texture.magFilter = THREE.LinearFilter;
    texture.generateMipmaps = false;
    texture.flipY = false;
    if ("colorSpace" in texture && (_window$THREE = window.THREE) !== null && _window$THREE !== void 0 && _window$THREE.SRGBColorSpace) {
      texture.colorSpace = window.THREE.SRGBColorSpace;
    }
    var geometry = buildGeometry();
    var material = new THREE.MeshBasicMaterial({
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
      polygonOffsetUnits: -1
    });
    var mesh = new THREE.Mesh(geometry, material);
    mesh.renderOrder = 6;
    targetGroup.add(mesh);
  };
  var dilateIterations = getImageShapeDilateIterations(shape === null || shape === void 0 ? void 0 : shape.imageType);
  var loadFromImagePath = function loadFromImagePath() {
    if (useOriginalImageColor) {
      return getOrLoadPlanLargeShapeImage(imagePath, shape === null || shape === void 0 ? void 0 : shape.imageType).then(function (image) {
        var texture = new THREE.Texture(image);
        addTexturedMesh(texture, "#ffffff");
        return true;
      })["catch"](function () {
        return false;
      });
    }
    return getOrLoadPlanLargeShapeTintedCanvas(imagePath, shape === null || shape === void 0 ? void 0 : shape.color, shape === null || shape === void 0 ? void 0 : shape.imageType, {
      dilateIterations: dilateIterations
    }).then(function (tintedCanvas) {
      var texture = new THREE.CanvasTexture(tintedCanvas);
      addTexturedMesh(texture, "#ffffff");
      return true;
    })["catch"](function () {
      return getOrLoadPlanLargeShapeImage(imagePath, shape === null || shape === void 0 ? void 0 : shape.imageType).then(function (image) {
        var texture = new THREE.Texture(image);
        addTexturedMesh(texture, "#ffffff");
        return true;
      });
    })["catch"](function () {
      return false;
    });
  };
  void loadFromImagePath();
}
function addViewerImageStrokeOverlay(points, imageSource, colorRaw, targetGroup, renderNonce) {
  if (!Array.isArray(points) || points.length !== 4 || !targetGroup || !targetGroup.parent || !viewer3d.initialized) {
    return;
  }
  var canvas = ensureCanvasFromImageSource(imageSource);
  if (!canvas) {
    return;
  }
  var width = Number(canvas.width);
  var height = Number(canvas.height);
  if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 1 || height <= 1) {
    return;
  }
  var ctx = canvas.getContext("2d", {
    willReadFrequently: true
  });
  if (!ctx) {
    return;
  }
  var imageData;
  try {
    imageData = ctx.getImageData(0, 0, width, height);
  } catch (_error) {
    return;
  }
  var data = imageData.data;
  var minDim = Math.max(1, Math.min(width, height));
  var stride = Math.max(1, Math.floor(minDim / 180));
  var maxPoints = 18000;
  var positions = [];
  var zOffset = 0.004;
  var useLineMask = false;
  for (var y = 0; y < height && !useLineMask; y += Math.max(1, stride * 2)) {
    for (var x = 0; x < width; x += Math.max(1, stride * 2)) {
      var idx = (y * width + x) * 4;
      var alpha = data[idx + 3];
      if (alpha < 1) {
        continue;
      }
      var r = data[idx];
      var g = data[idx + 1];
      var b = data[idx + 2];
      var isNearWhite = r >= 244 && g >= 244 && b >= 244;
      if (!isNearWhite) {
        useLineMask = true;
        break;
      }
    }
  }
  for (var _y2 = 0; _y2 < height; _y2 += stride) {
    for (var _x9 = 0; _x9 < width; _x9 += stride) {
      var index = (_y2 * width + _x9) * 4;
      var _alpha = data[index + 3];
      if (_alpha < 1) {
        continue;
      }
      if (useLineMask) {
        var _r = data[index];
        var _g = data[index + 1];
        var _b = data[index + 2];
        var _isNearWhite = _r >= 244 && _g >= 244 && _b >= 244;
        if (_isNearWhite) {
          continue;
        }
      }
      var u = width <= 1 ? 0 : _x9 / (width - 1);
      var v = height <= 1 ? 0 : _y2 / (height - 1);
      var world = interpolateViewerQuadPoint(points, u, v);
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
  var geometry = new THREE.BufferGeometry();
  geometry.setAttribute("position", new THREE.Float32BufferAttribute(positions, 3));
  var material = new THREE.PointsMaterial({
    color: new THREE.Color(parseHexColor(colorRaw).hex),
    size: 3.2,
    sizeAttenuation: false,
    transparent: true,
    opacity: 0.98,
    depthTest: false,
    depthWrite: false
  });
  var cloud = new THREE.Points(geometry, material);
  cloud.renderOrder = 7;
  targetGroup.add(cloud);
}
function ensureCanvasFromImageSource(imageSource) {
  if (!imageSource) {
    return null;
  }
  var isCanvas = typeof HTMLCanvasElement !== "undefined" && imageSource instanceof HTMLCanvasElement && Number(imageSource.width) > 0;
  if (isCanvas) {
    return imageSource;
  }
  var width = Math.max(1, Number(imageSource.naturalWidth || imageSource.width) || 0);
  var height = Math.max(1, Number(imageSource.naturalHeight || imageSource.height) || 0);
  if (!width || !height) {
    return null;
  }
  var canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  var ctx = canvas.getContext("2d", {
    willReadFrequently: true
  });
  if (!ctx) {
    return null;
  }
  ctx.drawImage(imageSource, 0, 0, width, height);
  return canvas;
}
function buildTintedCanvasFromImageSource(imageSource, colorRaw) {
  var sourceCanvas = ensureCanvasFromImageSource(imageSource);
  if (!sourceCanvas) {
    return null;
  }
  var width = Math.max(1, Number(sourceCanvas.width) || 0);
  var height = Math.max(1, Number(sourceCanvas.height) || 0);
  if (!width || !height) {
    return null;
  }
  var canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  var ctx = canvas.getContext("2d", {
    willReadFrequently: true
  });
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
  var u = clampNumber(Number(uRaw), 0, 1);
  var v = clampNumber(Number(vRaw), 0, 1);
  var p0 = points[0];
  var p1 = points[1];
  var p2 = points[2];
  var p3 = points[3];
  var w0 = (1 - u) * (1 - v);
  var w1 = u * (1 - v);
  var w2 = u * v;
  var w3 = (1 - u) * v;
  return {
    x: p0.x * w0 + p1.x * w1 + p2.x * w2 + p3.x * w3,
    y: p0.y * w0 + p1.y * w1 + p2.y * w2 + p3.y * w3,
    z: p0.z * w0 + p1.z * w1 + p2.z * w2 + p3.z * w3
  };
}
function clampNumber(valueRaw, min, max) {
  var valueNum = Number(valueRaw);
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
function getOrLoadPlanLargeShapeImage(imagePathRaw) {
  var shapeTypeRaw = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : "";
  var candidates = getLargeShapeImagePathCandidates(shapeTypeRaw, imagePathRaw);
  if (!candidates.length) {
    return Promise.reject(new Error("imagePath is empty"));
  }
  var cacheKey = candidates.join("|");
  var cached = planLargeShapeImageCache.get(cacheKey);
  if (cached) {
    return cached;
  }
  var promise = new Promise(function (resolve, reject) {
    var _tryLoad = function tryLoad(index) {
      if (index >= candidates.length) {
        reject(new Error("image load failed"));
        return;
      }
      var image = new Image();
      image.onload = function () {
        return resolve(image);
      };
      image.onerror = function () {
        return _tryLoad(index + 1);
      };
      image.src = candidates[index];
    };
    _tryLoad(0);
  });
  promise["catch"](function () {
    if (planLargeShapeImageCache.get(cacheKey) === promise) {
      planLargeShapeImageCache["delete"](cacheKey);
    }
  });
  planLargeShapeImageCache.set(cacheKey, promise);
  return promise;
}
function dilateAlphaMask(alphaRaw, width, height) {
  var iterations = arguments.length > 3 && arguments[3] !== undefined ? arguments[3] : 1;
  var w = Math.max(1, Math.floor(width));
  var h = Math.max(1, Math.floor(height));
  var src = alphaRaw instanceof Uint8ClampedArray ? alphaRaw : new Uint8ClampedArray(w * h);
  var iterCount = Math.max(0, Math.floor(iterations));
  for (var iter = 0; iter < iterCount; iter += 1) {
    var dst = new Uint8ClampedArray(src.length);
    for (var y = 0; y < h; y += 1) {
      for (var x = 0; x < w; x += 1) {
        var maxAlpha = 0;
        for (var dy = -1; dy <= 1; dy += 1) {
          var ny = y + dy;
          if (ny < 0 || ny >= h) {
            continue;
          }
          for (var dx = -1; dx <= 1; dx += 1) {
            var nx = x + dx;
            if (nx < 0 || nx >= w) {
              continue;
            }
            var alpha = src[ny * w + nx];
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
function getOrLoadPlanLargeShapeTintedCanvas(imagePathRaw, colorRaw) {
  var shapeTypeRaw = arguments.length > 2 && arguments[2] !== undefined ? arguments[2] : "";
  var options = arguments.length > 3 && arguments[3] !== undefined ? arguments[3] : {};
  var candidates = getLargeShapeImagePathCandidates(shapeTypeRaw, imagePathRaw);
  if (!candidates.length) {
    return Promise.reject(new Error("imagePath is empty"));
  }
  var tint = parseHexColor(colorRaw);
  var requestedIterations = Number(options === null || options === void 0 ? void 0 : options.dilateIterations);
  var dilateIterations = Number.isFinite(requestedIterations) ? clamp(Math.round(requestedIterations), 0, 8) : getImageShapeDilateIterations(shapeTypeRaw);
  var key = "".concat(candidates.join("|"), "::").concat(tint.hex, "::d").concat(dilateIterations);
  var cached = planLargeShapeTintedCanvasCache.get(key);
  if (cached) {
    return cached;
  }
  var promise = getOrLoadPlanLargeShapeImage(candidates[0], shapeTypeRaw).then(function (image) {
    var width = Math.max(1, Number(image.naturalWidth || image.width) || 1);
    var height = Math.max(1, Number(image.naturalHeight || image.height) || 1);
    var canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    var ctx = canvas.getContext("2d", {
      willReadFrequently: true
    });
    if (!ctx) {
      throw new Error("2d context unavailable");
    }
    ctx.clearRect(0, 0, width, height);
    ctx.drawImage(image, 0, 0, width, height);
    var imageData;
    try {
      imageData = ctx.getImageData(0, 0, width, height);
    } catch (_error) {
      ctx.globalCompositeOperation = "source-in";
      ctx.fillStyle = tint.hex;
      ctx.fillRect(0, 0, width, height);
      ctx.globalCompositeOperation = "source-over";
      return canvas;
    }
    var data = imageData.data;
    var alphaMask = new Uint8ClampedArray(width * height);
    var lineMask = new Uint8ClampedArray(width * height);
    for (var i = 0, p = 0; i < data.length; i += 4, p += 1) {
      var alpha = data[i + 3];
      alphaMask[p] = alpha;
      if (alpha === 0) {
        continue;
      }
      var r = data[i];
      var g = data[i + 1];
      var b = data[i + 2];
      var isNearWhite = r >= 244 && g >= 244 && b >= 244;
      if (!isNearWhite) {
        lineMask[p] = alpha;
      }
    }
    var useLineMask = false;
    for (var _i3 = 0; _i3 < lineMask.length; _i3 += 1) {
      if (lineMask[_i3] > 0) {
        useLineMask = true;
        break;
      }
    }
    var sourceMask = useLineMask ? lineMask : alphaMask;
    var expandedAlpha = dilateAlphaMask(sourceMask, width, height, dilateIterations);
    for (var _i4 = 0; _i4 < data.length; _i4 += 4) {
      var _alpha2 = expandedAlpha[_i4 / 4];
      if (_alpha2 === 0) {
        data[_i4 + 3] = 0;
        continue;
      }
      data[_i4] = tint.r;
      data[_i4 + 1] = tint.g;
      data[_i4 + 2] = tint.b;
      data[_i4 + 3] = _alpha2;
    }
    ctx.putImageData(imageData, 0, 0);
    return canvas;
  });
  promise["catch"](function () {
    if (planLargeShapeTintedCanvasCache.get(key) === promise) {
      planLargeShapeTintedCanvasCache["delete"](key);
    }
  });
  planLargeShapeTintedCanvasCache.set(key, promise);
  return promise;
}
function getPlanLargeShapeTintedDataUrl(imagePathRaw, colorRaw) {
  var shapeTypeRaw = arguments.length > 2 && arguments[2] !== undefined ? arguments[2] : "";
  var options = arguments.length > 3 && arguments[3] !== undefined ? arguments[3] : {};
  var candidates = getLargeShapeImagePathCandidates(shapeTypeRaw, imagePathRaw);
  if (!candidates.length) {
    return "";
  }
  var tint = parseHexColor(colorRaw);
  var requestedIterations = Number(options === null || options === void 0 ? void 0 : options.dilateIterations);
  var dilateIterations = Number.isFinite(requestedIterations) ? clamp(Math.round(requestedIterations), 0, 8) : getImageShapeDilateIterations(shapeTypeRaw);
  var key = "".concat(candidates.join("|"), "::").concat(tint.hex, "::d").concat(dilateIterations);
  var cached = planLargeShapeTintedDataUrlCache.get(key);
  if (cached === "loading") {
    return "";
  }
  if (typeof cached === "string") {
    return cached;
  }
  var preferredSource = candidates.find(function (candidate) {
    return !String(candidate).startsWith("data:");
  }) || candidates[0];
  planLargeShapeTintedDataUrlCache.set(key, "loading");
  getOrLoadPlanLargeShapeTintedCanvas(preferredSource, tint.hex, shapeTypeRaw, {
    dilateIterations: dilateIterations
  }).then(function (canvas) {
    var dataUrl = canvas.toDataURL("image/png");
    if (dataUrl.length > PLAN_IMAGE_TINTED_DATA_URL_MAX_LENGTH) {
      planLargeShapeTintedDataUrlCache.set(key, "");
      renderOutputs();
      return;
    }
    planLargeShapeTintedDataUrlCache.set(key, dataUrl);
    renderOutputs();
  })["catch"](function () {
    planLargeShapeTintedDataUrlCache["delete"](key);
  });
  return "";
}
function createViewerTextSprite(textRaw, colorRaw) {
  var options = arguments.length > 2 && arguments[2] !== undefined ? arguments[2] : {};
  var fontPx = Math.max(12, Number(options === null || options === void 0 ? void 0 : options.fontPx) || 44);
  var scaleX = Math.max(0.05, Number(options === null || options === void 0 ? void 0 : options.scaleX) || 0.95);
  var scaleY = Math.max(0.05, Number(options === null || options === void 0 ? void 0 : options.scaleY) || 0.24);
  var resolutionScale = clamp(Math.round(Number(options === null || options === void 0 ? void 0 : options.resolutionScale) || 2), 1, 4);
  var sizeAttenuation = (options === null || options === void 0 ? void 0 : options.sizeAttenuation) !== false;
  var canvas = document.createElement("canvas");
  canvas.width = 512 * resolutionScale;
  canvas.height = 128 * resolutionScale;
  var ctx = canvas.getContext("2d");
  ctx.clearRect(0, 0, canvas.width, canvas.height);
  ctx.font = "700 ".concat(fontPx * resolutionScale, "px 'Yu Gothic', sans-serif");
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.fillStyle = value(colorRaw) || "#111827";
  ctx.strokeStyle = "rgba(255,255,255,0.96)";
  ctx.lineWidth = 8 * resolutionScale;
  ctx.strokeText(value(textRaw) || "-", canvas.width / 2, canvas.height / 2);
  ctx.fillText(value(textRaw) || "-", canvas.width / 2, canvas.height / 2);
  var texture = new THREE.CanvasTexture(canvas);
  texture.generateMipmaps = false;
  texture.minFilter = THREE.LinearFilter;
  texture.magFilter = THREE.LinearFilter;
  texture.needsUpdate = true;
  var material = new THREE.SpriteMaterial({
    map: texture,
    transparent: true,
    depthTest: false,
    sizeAttenuation: sizeAttenuation
  });
  var sprite = new THREE.Sprite(material);
  sprite.scale.set(scaleX, scaleY, 1);
  return sprite;
}
function createViewerTextPlane(textRaw, colorRaw) {
  var options = arguments.length > 2 && arguments[2] !== undefined ? arguments[2] : {};
  var fontPx = Math.max(12, Number(options === null || options === void 0 ? void 0 : options.fontPx) || 44);
  var width = Math.max(0.05, Number(options === null || options === void 0 ? void 0 : options.width) || 0.8);
  var height = Math.max(0.05, Number(options === null || options === void 0 ? void 0 : options.height) || 0.22);
  var resolutionScale = clamp(Math.round(Number(options === null || options === void 0 ? void 0 : options.resolutionScale) || 2), 1, 4);
  var canvas = document.createElement("canvas");
  canvas.width = 512 * resolutionScale;
  canvas.height = 128 * resolutionScale;
  var ctx = canvas.getContext("2d");
  ctx.clearRect(0, 0, canvas.width, canvas.height);
  ctx.font = "700 ".concat(fontPx * resolutionScale, "px 'Yu Gothic', sans-serif");
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.fillStyle = value(colorRaw) || "#111827";
  ctx.strokeStyle = "rgba(255,255,255,0.96)";
  ctx.lineWidth = 8 * resolutionScale;
  ctx.strokeText(value(textRaw) || "-", canvas.width / 2, canvas.height / 2);
  ctx.fillText(value(textRaw) || "-", canvas.width / 2, canvas.height / 2);
  var texture = new THREE.CanvasTexture(canvas);
  texture.generateMipmaps = false;
  texture.minFilter = THREE.LinearFilter;
  texture.magFilter = THREE.LinearFilter;
  texture.needsUpdate = true;
  var material = new THREE.MeshBasicMaterial({
    map: texture,
    transparent: true,
    depthTest: false,
    side: THREE.DoubleSide
  });
  var geometry = new THREE.PlaneGeometry(width, height);
  var mesh = new THREE.Mesh(geometry, material);
  mesh.renderOrder = 5;
  return mesh;
}
function computeViewerBounds(shapes, metrics) {
  var xs = [];
  var ys = [];
  var zs = [];
  var axisMinZ = applyViewerVerticalScale(metrics.minZ, metrics.minZ);
  var axisMaxZ = applyViewerVerticalScale(metrics.maxZ, metrics.minZ);
  zs.push(axisMinZ, axisMaxZ);
  shapes.forEach(function (shape) {
    if (shape.type === "point") {
      xs.push(shape.x);
      ys.push(shape.y);
      zs.push(shape.z);
      return;
    }
    (shape.points || []).forEach(function (point) {
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
      maxZ: axisMaxZ
    };
  }
  return {
    minX: Math.min.apply(Math, xs),
    maxX: Math.max.apply(Math, xs),
    minY: Math.min.apply(Math, ys),
    maxY: Math.max.apply(Math, ys),
    minZ: Math.min.apply(Math, zs),
    maxZ: Math.max.apply(Math, zs)
  };
}
function applyViewerPerspective() {
  if (!viewer3d.initialized || !viewer3d.camera) {
    return;
  }
  var bounds = viewer3d.bounds;
  if (!bounds) {
    return;
  }
  var center = new THREE.Vector3((bounds.minX + bounds.maxX) / 2, (bounds.minY + bounds.maxY) / 2, (bounds.minZ + bounds.maxZ) / 2);
  var spanX = Math.max(1, bounds.maxX - bounds.minX);
  var spanY = Math.max(1, bounds.maxY - bounds.minY);
  var spanZ = Math.max(1, bounds.maxZ - bounds.minZ);
  var scale = normalizeViewerVerticalScale(viewerVerticalScale);
  var unscaledSpanZ = Math.max(1, spanZ / scale);
  var focus = new THREE.Vector3(center.x, center.y, bounds.minZ + spanZ * 0.16);
  var baseDist = Math.max(spanX, spanY) * 0.68 + Math.max(3.2, unscaledSpanZ * 2.2);
  var zoomInFactor = Math.pow(scale, -0.7);
  var dist = baseDist * zoomInFactor * 0.82;
  var perspective = selectedViewerPerspective === "iso" ? "se" : selectedViewerPerspective;
  if (perspective === "top") {
    viewer3d.camera.up.set(0, 1, 0);
    viewer3d.camera.position.set(focus.x, focus.y, focus.z + dist);
  } else {
    viewer3d.camera.up.set(0, 0, 1);
    var sideElev = dist * 0.14;
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
  var _window$THREE2;
  if (!viewer3d.controls || !((_window$THREE2 = window.THREE) !== null && _window$THREE2 !== void 0 && _window$THREE2.MOUSE)) {
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
  viewer3d.controls.mouseButtons.LEFT = viewer3d.defaultLeftMouseAction != null ? viewer3d.defaultLeftMouseAction : THREE.MOUSE.ROTATE;
  viewer3d.shiftPanActive = false;
}
function handleViewerControlPointerUp() {
  var _window$THREE3;
  if (!viewer3d.controls || !((_window$THREE3 = window.THREE) !== null && _window$THREE3 !== void 0 && _window$THREE3.MOUSE)) {
    return;
  }
  if (!viewer3d.shiftPanActive) {
    return;
  }
  viewer3d.controls.mouseButtons.LEFT = viewer3d.defaultLeftMouseAction != null ? viewer3d.defaultLeftMouseAction : THREE.MOUSE.ROTATE;
  viewer3d.shiftPanActive = false;
}
function handleViewerPointerMove(event) {
  if (isTouchLikePointerEvent(event)) {
    updateViewerTouchLongPressByMove(event);
    hideViewerTooltip();
    return;
  }
  var picked = pickViewerDataAtEvent(event);
  if (!picked) {
    hideViewerTooltip();
    return;
  }
  showViewerTooltip(event, picked);
}
function pickViewerDataAtEvent(event) {
  return pickViewerDataAtClient(event === null || event === void 0 ? void 0 : event.clientX, event === null || event === void 0 ? void 0 : event.clientY);
}
function pickViewerDataAtClient(clientXRaw, clientYRaw) {
  var _intersects$0$object;
  if (!viewer3d.initialized || !viewer3d.raycaster || !viewer3d.camera || !viewer3d.pickMeshes.length || !viewerCanvasWrap) {
    return null;
  }
  var clientX = Number(clientXRaw);
  var clientY = Number(clientYRaw);
  if (!Number.isFinite(clientX) || !Number.isFinite(clientY)) {
    return null;
  }
  var rect = viewerCanvasWrap.getBoundingClientRect();
  if (rect.width <= 0 || rect.height <= 0) {
    return null;
  }
  var x = (clientX - rect.left) / rect.width * 2 - 1;
  var y = -((clientY - rect.top) / rect.height) * 2 + 1;
  viewer3d.pointer.set(x, y);
  viewer3d.raycaster.setFromCamera(viewer3d.pointer, viewer3d.camera);
  var intersects = viewer3d.raycaster.intersectObjects(viewer3d.pickMeshes, false);
  if (!intersects.length) {
    return null;
  }
  return ((_intersects$0$object = intersects[0].object) === null || _intersects$0$object === void 0 ? void 0 : _intersects$0$object.userData) || null;
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
  viewerTouchLongPressState.timer = window.setTimeout(function () {
    if (viewerTouchLongPressState.pointerId == null) {
      return;
    }
    viewerTouchLongPressState.timer = 0;
    var picked = pickViewerDataAtClient(viewerTouchLongPressState.startX, viewerTouchLongPressState.startY);
    if (!picked || !value(picked.id)) {
      return;
    }
    hideViewerTooltip();
    showHoverEditMenu(viewerTouchLongPressState.startX, viewerTouchLongPressState.startY, picked.id, picked.kuwaku, picked.label);
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
  var moved = pointerMovedBeyondThreshold(event.clientX, event.clientY, viewerTouchLongPressState.startX, viewerTouchLongPressState.startY, TOUCH_LONG_PRESS_MOVE_THRESHOLD_PX);
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
  var picked = pickViewerDataAtEvent(event);
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
  var hasAltitude = Number.isFinite(Number(data.altitudeM));
  var altitudeTextBase = hasAltitude ? Number(data.altitudeM).toFixed(3).replace(/\.?0+$/, "") : "-";
  var altitudeText = data.altitudeEstimated ? "".concat(altitudeTextBase, "\uFF08\u4EEE\uFF09") : altitudeTextBase;
  viewerTooltip.innerHTML = "\n    <div><strong>\u6A19\u672C\u756A\u53F7:</strong> ".concat(escapeHtml(value(data.label) || "-"), "</div>\n    <div><strong>\u540D\u79F0:</strong> ").concat(escapeHtml(value(data.nameMemo) || "-"), "</div>\n    <div><strong>\u30E6\u30CB\u30C3\u30C8:</strong> ").concat(escapeHtml(value(data.unit) || "-"), "</div>\n    <div><strong>\u30B5\u30D6\u30E6\u30CB\u30C3\u30C8:</strong> ").concat(escapeHtml(value(data.detail) || "-"), "</div>\n    <div><strong>\u533A\u753B:</strong> ").concat(escapeHtml(value(data.kuwaku) || "-"), "</div>\n    <div><strong>\u6A19\u9AD8(m):</strong> ").concat(escapeHtml(altitudeText), "</div>\n  ");
  viewerTooltip.hidden = false;
  var rect = viewerCanvasWrap.getBoundingClientRect();
  var maxX = Math.max(8, rect.width - 240);
  var maxY = Math.max(8, rect.height - 132);
  var x = clamp(event.clientX - rect.left + 14, 8, maxX);
  var y = clamp(event.clientY - rect.top + 12, 8, maxY);
  viewerTooltip.style.left = "".concat(x, "px");
  viewerTooltip.style.top = "".concat(y, "px");
}
function hideViewerTooltip() {
  if (!viewerTooltip) {
    return;
  }
  viewerTooltip.hidden = true;
}
function blockIndexToLabel(indexRaw) {
  var index = Number(indexRaw);
  if (!Number.isFinite(index) || index < 1) {
    return "A";
  }
  var label = "";
  while (index > 0) {
    var remainder = (index - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    index = Math.floor((index - 1) / 26);
  }
  return label || "A";
}
function normalizeViewerHead(headRaw) {
  var head = value(headRaw).toUpperCase();
  if (head === "Ⅲ" || head === "3" || head === "III") {
    return "Ⅲ";
  }
  if (head === "Ⅱ" || head === "2" || head === "II") {
    return "Ⅱ";
  }
  return "Ⅰ";
}
function kuwakuToViewerXIndex(kuwakuParts) {
  var parts = kuwakuParts || {};
  var head = normalizeViewerHead(parts.headB);
  var headIndex = VIEWER_HEAD_INDEX_MAP.get(head);
  var block = normalizeKuwakuBlock(parts.block);
  var baseLetterIndex = block ? block.charCodeAt(0) - 65 : 0;
  var letterIndex = Number.isFinite(baseLetterIndex) && baseLetterIndex >= 0 && baseLetterIndex < 26 ? baseLetterIndex : 0;
  if (!Number.isFinite(headIndex)) {
    return blockLabelToIndex(block) - 1;
  }
  return headIndex * 26 + letterIndex;
}
function viewerXIndexToHeadBlock(indexRaw) {
  var index = Number(indexRaw);
  if (!Number.isFinite(index)) {
    return {
      head: "Ⅰ",
      block: "A"
    };
  }
  var seqLength = VIEWER_HEAD_SEQUENCE.length;
  var normalized = (Math.floor(index) % (26 * seqLength) + 26 * seqLength) % (26 * seqLength);
  var headIndex = Math.floor(normalized / 26) % seqLength;
  var letterIndex = normalized % 26;
  return {
    head: VIEWER_HEAD_SEQUENCE[headIndex] || "Ⅰ",
    block: String.fromCharCode(65 + letterIndex)
  };
}
function buildDetailText(detailRaw) {
  var detailSubRaw = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : "";
  var detail = value(detailRaw);
  var detailSub = value(detailSubRaw);
  if (detail && detailSub) {
    return "".concat(detail, " ").concat(detailSub);
  }
  return detail || detailSub;
}
function formatDetailForRecord(record) {
  return buildDetailText(record === null || record === void 0 ? void 0 : record.detail, record === null || record === void 0 ? void 0 : record.detailSub);
}
function detailValueForSelect(detailRaw) {
  var detail = value(detailRaw);
  return detail || EMPTY_DETAIL_VALUE;
}
function detailLabelForSelect(detailValue) {
  return detailValue === EMPTY_DETAIL_VALUE ? "（未設定）" : detailValue;
}
function detailSubValueForSelect(detailSubRaw) {
  var detailSub = value(detailSubRaw);
  return detailSub || EMPTY_DETAIL_SUB_VALUE;
}
function detailSubLabelForSelect(detailSubValue) {
  return detailSubValue === EMPTY_DETAIL_SUB_VALUE ? "（未設定）" : detailSubValue;
}
function getRecordKuwaku(record) {
  return normalizeKuwakuText(record === null || record === void 0 ? void 0 : record.kuwaku);
}
function getRecordTeamValue(record) {
  var teamState = normalizeTeamState(value(record === null || record === void 0 ? void 0 : record.team), value(record === null || record === void 0 ? void 0 : record.teamOther));
  if (teamState.team) {
    return formatTeamValue(teamState);
  }
  return formatTeamValue(state.site);
}
function getRecordLevelHeight(record) {
  var _state$site8;
  return value(record === null || record === void 0 ? void 0 : record.levelHeight) || value((_state$site8 = state.site) === null || _state$site8 === void 0 ? void 0 : _state$site8.levelHeight);
}
function getRecordDate(record) {
  var _state$site9;
  return value(record === null || record === void 0 ? void 0 : record.date) || value((_state$site9 = state.site) === null || _state$site9 === void 0 ? void 0 : _state$site9.date);
}
function getRecordTeamLead(record) {
  var _state$site0;
  return value(record === null || record === void 0 ? void 0 : record.teamLead) || value((_state$site0 = state.site) === null || _state$site0 === void 0 ? void 0 : _state$site0.teamLead);
}
function getRecordRecorder(record) {
  var _state$site1;
  return value(record === null || record === void 0 ? void 0 : record.recorder) || value((_state$site1 = state.site) === null || _state$site1 === void 0 ? void 0 : _state$site1.recorder);
}
function kuwakuValueForSelect(kuwakuRaw) {
  var kuwaku = normalizeKuwakuText(kuwakuRaw);
  return kuwaku || EMPTY_KUWAKU_VALUE;
}
function kuwakuLabelForSelect(kuwakuValue) {
  return kuwakuValue === EMPTY_KUWAKU_VALUE ? "（未設定）" : kuwakuValue;
}
function isDefaultKuwaku(kuwakuRaw) {
  return normalizeKuwakuText(kuwakuRaw) === DEFAULT_KUWAKU;
}
function buildPlanDrawableMeta(record) {
  var specimen = parseSpecimenNo(record.specimenNo, record.specimenPrefix, record.specimenSerial);
  var prefix = normalizeSpecimenPrefix(specimen.prefix);
  return {
    id: value(record.id),
    kuwaku: value(record.kuwaku) || getRecordKuwaku(record),
    color: getSpecimenPrefixColor(prefix),
    label: record.specimenNo || "",
    nameMemo: value(record.nameMemo),
    unit: value(record.unit),
    detail: buildDetailText(record.detail, record.detailSub)
  };
}
function convertPositionToPlanCoords(nsDirRaw, nsCmRaw, ewDirRaw, ewCmRaw) {
  var nsCm = parseDistanceToCm(nsCmRaw);
  var ewCm = parseDistanceToCm(ewCmRaw);
  if (nsCm == null || ewCm == null) {
    return null;
  }
  var nsDir = normalizeNsDir(nsDirRaw);
  var ewDir = normalizeEwDir(ewDirRaw);
  var yRaw = nsDir === "北から" ? nsCm : PLAN_SIZE_CM - nsCm;
  var xRaw = ewDir === "西から" ? ewCm : PLAN_SIZE_CM - ewCm;
  return {
    x: xRaw,
    y: yRaw
  };
}
function parseLargeAxisAzimuth(valueRaw) {
  var text = normalizeLargeAxisDirection(valueRaw);
  if (text === "NS") {
    return 0;
  }
  if (text === "EW") {
    return 90;
  }
  var matched = text.match(/^([NS])(\d+(?:\.\d+)?)([EW])$/);
  if (!matched) {
    return null;
  }
  var _matched = _slicedToArray(matched, 4),
    ns = _matched[1],
    degreeRaw = _matched[2],
    ew = _matched[3];
  var degree = Number(degreeRaw);
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
  var text = value(valueRaw).replace(/[°度]/g, "");
  if (!text) {
    return null;
  }
  var matched = text.match(/-?\d+(?:\.\d+)?/);
  if (!matched) {
    return null;
  }
  var num = Number(matched[0]);
  if (!Number.isFinite(num)) {
    return null;
  }
  return clamp(Math.abs(num), 0, 90);
}
function azimuthToPlanUnitVector(azimuthDegRaw) {
  var azimuthDeg = Number(azimuthDegRaw);
  if (!Number.isFinite(azimuthDeg)) {
    return {
      dx: 0,
      dy: -1
    };
  }
  var rad = azimuthDeg * Math.PI / 180;
  return {
    dx: Math.sin(rad),
    dy: -Math.cos(rad)
  };
}
function azimuthToSvgRotationDeg(azimuthDegRaw) {
  var azimuthDeg = Number(azimuthDegRaw);
  if (!Number.isFinite(azimuthDeg)) {
    return 0;
  }
  return azimuthDeg - 90;
}
function pointsToAzimuthDeg(pointA, pointB) {
  if (!pointA || !pointB) {
    return null;
  }
  var dx = pointB.x - pointA.x;
  var dy = pointB.y - pointA.y;
  var distance = Math.hypot(dx, dy);
  if (!Number.isFinite(distance) || distance <= 0) {
    return null;
  }
  var rad = Math.atan2(dx, -dy);
  var deg = rad * 180 / Math.PI;
  return (deg + 360) % 360;
}
function buildConvexHull2d(pointsRaw) {
  var points = Array.isArray(pointsRaw) ? pointsRaw : [];
  var uniqueMap = new Map();
  points.forEach(function (point) {
    var x = Number(point === null || point === void 0 ? void 0 : point.x);
    var y = Number(point === null || point === void 0 ? void 0 : point.y);
    if (!Number.isFinite(x) || !Number.isFinite(y)) {
      return;
    }
    var key = "".concat(x.toFixed(6), "|").concat(y.toFixed(6));
    if (!uniqueMap.has(key)) {
      uniqueMap.set(key, {
        x: x,
        y: y
      });
    }
  });
  var unique = Array.from(uniqueMap.values());
  if (unique.length <= 2) {
    return unique;
  }
  unique.sort(function (a, b) {
    return a.x === b.x ? a.y - b.y : a.x - b.x;
  });
  var cross = function cross(origin, a, b) {
    return (a.x - origin.x) * (b.y - origin.y) - (a.y - origin.y) * (b.x - origin.x);
  };
  var lower = [];
  unique.forEach(function (point) {
    while (lower.length >= 2 && cross(lower[lower.length - 2], lower[lower.length - 1], point) <= 0) {
      lower.pop();
    }
    lower.push(point);
  });
  var upper = [];
  for (var i = unique.length - 1; i >= 0; i -= 1) {
    var point = unique[i];
    while (upper.length >= 2 && cross(upper[upper.length - 2], upper[upper.length - 1], point) <= 0) {
      upper.pop();
    }
    upper.push(point);
  }
  lower.pop();
  upper.pop();
  return [].concat(lower, upper);
}
function buildHullPointsFromSource(pointsRaw) {
  var zFallback = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : null;
  var points = Array.isArray(pointsRaw) ? pointsRaw : [];
  if (points.length <= 2) {
    return points.filter(function (point) {
      return Number.isFinite(Number(point === null || point === void 0 ? void 0 : point.x)) && Number.isFinite(Number(point === null || point === void 0 ? void 0 : point.y));
    });
  }
  var hull2d = buildConvexHull2d(points);
  if (!hull2d.length) {
    return [];
  }
  var keyOf = function keyOf(xRaw, yRaw) {
    return "".concat(Number(xRaw).toFixed(6), "|").concat(Number(yRaw).toFixed(6));
  };
  var pointMap = new Map();
  points.forEach(function (point) {
    var x = Number(point === null || point === void 0 ? void 0 : point.x);
    var y = Number(point === null || point === void 0 ? void 0 : point.y);
    if (!Number.isFinite(x) || !Number.isFinite(y)) {
      return;
    }
    pointMap.set(keyOf(x, y), point);
  });
  var fallbackZ = Number.isFinite(Number(zFallback)) ? Number(zFallback) : 0;
  return hull2d.map(function (point) {
    var original = pointMap.get(keyOf(point.x, point.y));
    if (original) {
      return original;
    }
    return {
      x: point.x,
      y: point.y,
      z: fallbackZ
    };
  });
}
function parseImageQuadPlanPoints(record) {
  var center = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : null;
  var centerPointRaw = center ? {
    x: Number(center.x),
    y: Number(center.y)
  } : convertPositionToPlanCoords(record === null || record === void 0 ? void 0 : record.nsDir, record === null || record === void 0 ? void 0 : record.nsCm, record === null || record === void 0 ? void 0 : record.ewDir, record === null || record === void 0 ? void 0 : record.ewCm);
  var centerPoint = centerPointRaw && Number.isFinite(centerPointRaw.x) && Number.isFinite(centerPointRaw.y) ? centerPointRaw : null;
  var widthCm = parseDistanceToCm(record === null || record === void 0 ? void 0 : record.imgFrameWidthCm);
  var heightCm = parseDistanceToCm(record === null || record === void 0 ? void 0 : record.imgFrameHeightCm);
  var rotationText = normalizeImageRotationDeg(record === null || record === void 0 ? void 0 : record.imgRotateDeg);
  if (centerPoint && widthCm != null && widthCm > 0 && heightCm != null && heightCm > 0) {
    var rotationDeg = Number(rotationText === "" ? "0" : rotationText);
    var skewXDeg = Number(normalizeImageSkewDeg(record === null || record === void 0 ? void 0 : record.imgSkewXDeg) || "0");
    var skewYDeg = Number(normalizeImageSkewDeg(record === null || record === void 0 ? void 0 : record.imgSkewYDeg) || "0");
    var flipH = normalizeToggleFlag(record === null || record === void 0 ? void 0 : record.imgFlipH) === "1";
    var flipV = normalizeToggleFlag(record === null || record === void 0 ? void 0 : record.imgFlipV) === "1";
    var up = azimuthToPlanUnitVector(rotationDeg);
    var down = {
      dx: -up.dx,
      dy: -up.dy
    };
    var right = azimuthToPlanUnitVector((rotationDeg + 90) % 360);
    var tanSkewX = Math.tan(skewXDeg * Math.PI / 180);
    var tanSkewY = Math.tan(skewYDeg * Math.PI / 180);
    var baseLocalCorners = [{
      x: -widthCm / 2,
      y: -heightCm / 2
    }, {
      x: widthCm / 2,
      y: -heightCm / 2
    }, {
      x: widthCm / 2,
      y: heightCm / 2
    }, {
      x: -widthCm / 2,
      y: heightCm / 2
    }];
    var shearedWorldCorners = baseLocalCorners.map(function (local) {
      var xSheared = local.x + tanSkewX * local.y;
      var ySheared = local.y + tanSkewY * local.x;
      return {
        x: centerPoint.x + right.dx * xSheared + down.dx * ySheared,
        y: centerPoint.y + right.dy * xSheared + down.dy * ySheared
      };
    });
    var order = [0, 1, 2, 3];
    if (flipH && flipV) {
      order = [2, 3, 0, 1];
    } else if (flipH) {
      order = [1, 0, 3, 2];
    } else if (flipV) {
      order = [3, 2, 1, 0];
    }
    return order.map(function (index) {
      return shearedWorldCorners[index];
    });
  }
  var corner1 = convertPositionToPlanCoords(record === null || record === void 0 ? void 0 : record.imgP1NsDir, record === null || record === void 0 ? void 0 : record.imgP1NsCm, record === null || record === void 0 ? void 0 : record.imgP1EwDir, record === null || record === void 0 ? void 0 : record.imgP1EwCm);
  var corner2 = convertPositionToPlanCoords(record === null || record === void 0 ? void 0 : record.imgP2NsDir, record === null || record === void 0 ? void 0 : record.imgP2NsCm, record === null || record === void 0 ? void 0 : record.imgP2EwDir, record === null || record === void 0 ? void 0 : record.imgP2EwCm);
  var corner3 = convertPositionToPlanCoords(record === null || record === void 0 ? void 0 : record.imgP3NsDir, record === null || record === void 0 ? void 0 : record.imgP3NsCm, record === null || record === void 0 ? void 0 : record.imgP3EwDir, record === null || record === void 0 ? void 0 : record.imgP3EwCm);
  var corner4 = convertPositionToPlanCoords(record === null || record === void 0 ? void 0 : record.imgP4NsDir, record === null || record === void 0 ? void 0 : record.imgP4NsCm, record === null || record === void 0 ? void 0 : record.imgP4EwDir, record === null || record === void 0 ? void 0 : record.imgP4EwCm);
  if (!corner1 || !corner2 || !corner3 || !corner4) {
    if (!centerPoint) {
      return null;
    }
    var candidates = [corner1, corner2, corner3, corner4, centerPoint].filter(Boolean);
    if (!candidates.length) {
      return null;
    }
    var minX = Math.min.apply(Math, _toConsumableArray(candidates.map(function (point) {
      return point.x;
    })));
    var maxX = Math.max.apply(Math, _toConsumableArray(candidates.map(function (point) {
      return point.x;
    })));
    var minY = Math.min.apply(Math, _toConsumableArray(candidates.map(function (point) {
      return point.y;
    })));
    var maxY = Math.max.apply(Math, _toConsumableArray(candidates.map(function (point) {
      return point.y;
    })));
    var fallbackHalfSize = 20;
    if (Math.abs(maxX - minX) < 1) {
      minX = centerPoint.x - fallbackHalfSize;
      maxX = centerPoint.x + fallbackHalfSize;
    }
    if (Math.abs(maxY - minY) < 1) {
      minY = centerPoint.y - fallbackHalfSize;
      maxY = centerPoint.y + fallbackHalfSize;
    }
    return [{
      x: minX,
      y: minY
    }, {
      x: maxX,
      y: minY
    }, {
      x: maxX,
      y: maxY
    }, {
      x: minX,
      y: maxY
    }];
  }
  return [corner1, corner2, corner3, corner4];
}
function collectImageCornerPoints(record) {
  var corners = [convertPositionToPlanCoords(record === null || record === void 0 ? void 0 : record.imgP1NsDir, record === null || record === void 0 ? void 0 : record.imgP1NsCm, record === null || record === void 0 ? void 0 : record.imgP1EwDir, record === null || record === void 0 ? void 0 : record.imgP1EwCm), convertPositionToPlanCoords(record === null || record === void 0 ? void 0 : record.imgP2NsDir, record === null || record === void 0 ? void 0 : record.imgP2NsCm, record === null || record === void 0 ? void 0 : record.imgP2EwDir, record === null || record === void 0 ? void 0 : record.imgP2EwCm), convertPositionToPlanCoords(record === null || record === void 0 ? void 0 : record.imgP3NsDir, record === null || record === void 0 ? void 0 : record.imgP3NsCm, record === null || record === void 0 ? void 0 : record.imgP3EwDir, record === null || record === void 0 ? void 0 : record.imgP3EwCm), convertPositionToPlanCoords(record === null || record === void 0 ? void 0 : record.imgP4NsDir, record === null || record === void 0 ? void 0 : record.imgP4NsCm, record === null || record === void 0 ? void 0 : record.imgP4EwDir, record === null || record === void 0 ? void 0 : record.imgP4EwCm)].filter(Boolean);
  return corners;
}
function buildPlanDrawable(record) {
  var _parseLargeAxisAzimut;
  var meta = buildPlanDrawableMeta(record);
  var planSizeMode = normalizePlanSizeMode(record.planSizeMode);
  if (planSizeMode === "複数点") {
    var multiPoints = collectPlanMultiPointCoords(record);
    var fallbackCenter = convertPositionToPlanCoords(record === null || record === void 0 ? void 0 : record.nsDir, record === null || record === void 0 ? void 0 : record.nsCm, record === null || record === void 0 ? void 0 : record.ewDir, record === null || record === void 0 ? void 0 : record.ewCm);
    var points = multiPoints.length ? multiPoints : fallbackCenter ? [fallbackCenter] : [];
    if (!points.length) {
      return null;
    }
    var centroid = points.reduce(function (acc, point) {
      return {
        x: acc.x + point.x / points.length,
        y: acc.y + point.y / points.length
      };
    }, {
      x: 0,
      y: 0
    });
    return _objectSpread({
      type: "multipoint",
      points: points,
      hull: buildHullPointsFromSource(points),
      x: centroid.x,
      y: centroid.y
    }, meta);
  }
  var shapeType = normalizeLargeShapeType(record.largeShapeType);
  var normalizedShapeLabel = normalizeLargeShapeLabel(record.largeShapeType);
  var customImagePath = isCustomLargeShapeType(shapeType) ? normalizeCustomLargeImageDataUrl(record.customLargeImageDataUrl) : "";
  var isImageShape = isLargeShapeImageType(shapeType);
  var hasMappedImageType = largeShapeImagePathMap.has(normalizedShapeLabel);
  var rawImageCorners = collectImageCornerPoints(record);
  var center = convertPositionToPlanCoords(record.nsDir, record.nsCm, record.ewDir, record.ewCm);
  if (!center && (isImageShape || rawImageCorners.length)) {
    if (rawImageCorners.length) {
      center = rawImageCorners.reduce(function (acc, point) {
        return {
          x: acc.x + point.x / rawImageCorners.length,
          y: acc.y + point.y / rawImageCorners.length
        };
      }, {
        x: 0,
        y: 0
      });
    }
  }
  if (!center) {
    return null;
  }
  var resolvedImageType = isImageShape ? shapeType : hasMappedImageType ? normalizedShapeLabel : "";
  var shouldUseImageQuad = planSizeMode === "大きなもの" && (isImageShape || resolvedImageType && rawImageCorners.length > 0);
  var orientationAzimuth = shapeType === "直線状" ? parseLargeAxisAzimuth(record.largeAxisDirection) : (_parseLargeAxisAzimut = parseLargeAxisAzimuth(record.planeStrikeDirection)) !== null && _parseLargeAxisAzimut !== void 0 ? _parseLargeAxisAzimut : parseLargeAxisAzimuth(record.largeAxisDirection);
  var axisAzimuth = shouldUseImageQuad ? null : orientationAzimuth;
  if (planSizeMode !== "大きなもの" || !shapeType) {
    return _objectSpread({
      type: "point",
      x: center.x,
      y: center.y
    }, meta);
  }
  if (shapeType === "直線状") {
    var lineLength = parseDistanceToCm(record.lineLengthCm);
    if (lineLength == null || lineLength <= 0) {
      return null;
    }
    if (axisAzimuth == null) {
      return null;
    }
    var unit = azimuthToPlanUnitVector(axisAzimuth);
    var halfLength = lineLength / 2;
    var x1 = center.x - unit.dx * halfLength;
    var y1 = center.y - unit.dy * halfLength;
    var x2 = center.x + unit.dx * halfLength;
    var y2 = center.y + unit.dy * halfLength;
    return _objectSpread({
      type: "line",
      x1: x1,
      y1: y1,
      x2: x2,
      y2: y2,
      x: center.x,
      y: center.y
    }, meta);
  }
  if (shapeType === "長方形") {
    var side1 = parseDistanceToCm(record.rectSide1Cm);
    var side2 = parseDistanceToCm(record.rectSide2Cm);
    if (side1 == null || side2 == null) {
      return null;
    }
    var longSide = Math.max(side1, side2);
    var shortSide = Math.min(side1, side2);
    var width = Math.max(1, longSide);
    var height = Math.max(1, shortSide);
    return _objectSpread({
      type: "rect",
      x: center.x,
      y: center.y,
      left: center.x - width / 2,
      top: center.y - height / 2,
      width: width,
      height: height,
      rotationDeg: azimuthToSvgRotationDeg(axisAzimuth !== null && axisAzimuth !== void 0 ? axisAzimuth : 90)
    }, meta);
  }
  if (shapeType === "楕円") {
    var rx = parseDistanceToCm(record.ellipseLongRadiusCm);
    var ry = parseDistanceToCm(record.ellipseShortRadiusCm);
    if (rx == null || ry == null) {
      return null;
    }
    var longRadius = Math.max(rx, ry);
    var shortRadius = Math.min(rx, ry);
    return _objectSpread({
      type: "ellipse",
      x: center.x,
      y: center.y,
      rx: Math.min(Math.max(1, longRadius), PLAN_SIZE_CM),
      ry: Math.min(Math.max(1, shortRadius), PLAN_SIZE_CM),
      rotationDeg: azimuthToSvgRotationDeg(axisAzimuth !== null && axisAzimuth !== void 0 ? axisAzimuth : 90)
    }, meta);
  }
  if (shouldUseImageQuad) {
    var _points4 = parseImageQuadPlanPoints(record, center);
    if (!_points4) {
      return null;
    }
    var useOriginalImageColor = normalizeToggleFlag(record === null || record === void 0 ? void 0 : record.imgUseOriginalColor) === "1";
    var _centroid = _points4.reduce(function (acc, point) {
      return {
        x: acc.x + point.x / _points4.length,
        y: acc.y + point.y / _points4.length
      };
    }, {
      x: 0,
      y: 0
    });
    return _objectSpread({
      type: "imageQuad",
      points: _points4,
      imageType: resolvedImageType,
      imagePath: customImagePath || getLargeShapeImagePath(resolvedImageType),
      useOriginalImageColor: useOriginalImageColor,
      x: _centroid.x,
      y: _centroid.y
    }, meta);
  }
  return null;
}
function renderPlanDrawableSvg(drawable) {
  var index = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : 0;
  var labelX = Math.min(PLAN_SIZE_CM - 2, drawable.x + 6);
  var labelY = Math.max(8, drawable.y - 6);
  var ariaLabel = ["\u6A19\u672C\u756A\u53F7 ".concat(drawable.label || "未設定"), "\u5316\u77F3\u30FB\u907A\u7269\u540D\u79F0 ".concat(drawable.nameMemo || "未設定"), "\u30E6\u30CB\u30C3\u30C8 ".concat(drawable.unit || "未設定"), "\u30B5\u30D6\u30E6\u30CB\u30C3\u30C8 ".concat(drawable.detail || "未設定")].join(" / ");
  var shapeSvg = "";
  if (drawable.type === "line") {
    shapeSvg = "<line class=\"plan-shape-line\" x1=\"".concat(drawable.x1, "\" y1=\"").concat(drawable.y1, "\" x2=\"").concat(drawable.x2, "\" y2=\"").concat(drawable.y2, "\" stroke=\"").concat(drawable.color, "\" />");
  } else if (drawable.type === "multipoint") {
    var points = Array.isArray(drawable.points) ? drawable.points : [];
    var hull = Array.isArray(drawable.hull) ? drawable.hull : [];
    var hullPointsText = hull.map(function (point) {
      return "".concat(point.x, ",").concat(point.y);
    }).join(" ");
    var hullSvg = "";
    if (hull.length >= 3) {
      hullSvg = "<polygon class=\"plan-shape-multipoint-hull\" points=\"".concat(hullPointsText, "\" stroke=\"").concat(drawable.color, "\" />");
    } else if (hull.length === 2) {
      hullSvg = "<line class=\"plan-shape-multipoint-hull\" x1=\"".concat(hull[0].x, "\" y1=\"").concat(hull[0].y, "\" x2=\"").concat(hull[1].x, "\" y2=\"").concat(hull[1].y, "\" stroke=\"").concat(drawable.color, "\" />");
    }
    var pointsSvg = points.map(function (point) {
      return "<circle class=\"plan-shape-multipoint-dot\" cx=\"".concat(point.x, "\" cy=\"").concat(point.y, "\" r=\"4.5\" fill=\"").concat(drawable.color, "\" />");
    }).join("");
    shapeSvg = "".concat(hullSvg).concat(pointsSvg);
  } else if (drawable.type === "rect") {
    var transform = Number.isFinite(drawable.rotationDeg) ? " transform=\"rotate(".concat(drawable.rotationDeg, " ").concat(drawable.x, " ").concat(drawable.y, ")\"") : "";
    shapeSvg = "<rect class=\"plan-shape-rect\" x=\"".concat(drawable.left, "\" y=\"").concat(drawable.top, "\" width=\"").concat(drawable.width, "\" height=\"").concat(drawable.height, "\" stroke=\"").concat(drawable.color, "\"").concat(transform, " />");
  } else if (drawable.type === "ellipse") {
    var _transform = Number.isFinite(drawable.rotationDeg) ? " transform=\"rotate(".concat(drawable.rotationDeg, " ").concat(drawable.x, " ").concat(drawable.y, ")\"") : "";
    shapeSvg = "<ellipse class=\"plan-shape-ellipse\" cx=\"".concat(drawable.x, "\" cy=\"").concat(drawable.y, "\" rx=\"").concat(drawable.rx, "\" ry=\"").concat(drawable.ry, "\" stroke=\"").concat(drawable.color, "\"").concat(_transform, " />");
  } else if (drawable.type === "imageQuad") {
    var imageSvg = buildPlanImageWarpSvg(drawable, index);
    if (drawable.useOriginalImageColor) {
      shapeSvg = imageSvg;
    } else {
      var polygonPoints = (drawable.points || []).map(function (point) {
        return "".concat(point.x, ",").concat(point.y);
      }).join(" ");
      var outlineSvg = "<polygon class=\"plan-shape-image-outline\" points=\"".concat(polygonPoints, "\" fill=\"none\" stroke=\"").concat(drawable.color, "\" />");
      shapeSvg = imageSvg ? "".concat(imageSvg).concat(outlineSvg) : outlineSvg;
    }
  } else {
    shapeSvg = "<circle class=\"plan-point-hit\" cx=\"".concat(drawable.x, "\" cy=\"").concat(drawable.y, "\" r=\"5\" fill=\"").concat(drawable.color, "\" />");
  }
  var hotspotSvg = "<circle class=\"plan-point-hotspot\" cx=\"".concat(drawable.x, "\" cy=\"").concat(drawable.y, "\" r=\"12\" fill=\"transparent\" />");
  if (drawable.type === "imageQuad") {
    hotspotSvg = "<polygon class=\"plan-point-hotspot plan-image-hotspot\" points=\"".concat((drawable.points || []).map(function (point) {
      return "".concat(point.x, ",").concat(point.y);
    }).join(" "), "\" fill=\"transparent\" />");
  } else if (drawable.type === "multipoint") {
    var _points5 = Array.isArray(drawable.points) ? drawable.points : [];
    var _hull = Array.isArray(drawable.hull) ? drawable.hull : [];
    var hotspotPoints = _points5.map(function (point) {
      return "<circle class=\"plan-point-hotspot\" cx=\"".concat(point.x, "\" cy=\"").concat(point.y, "\" r=\"10\" fill=\"transparent\" />");
    }).join("");
    if (_hull.length >= 3) {
      hotspotSvg = "<polygon class=\"plan-point-hotspot\" points=\"".concat(_hull.map(function (point) {
        return "".concat(point.x, ",").concat(point.y);
      }).join(" "), "\" fill=\"transparent\" />").concat(hotspotPoints);
    } else {
      hotspotSvg = hotspotPoints || hotspotSvg;
    }
  }
  return "\n      <g\n        class=\"plan-point-group\"\n        data-id=\"".concat(escapeHtml(value(drawable.id) || ""), "\"\n        data-kuwaku=\"").concat(escapeHtml(value(drawable.kuwaku) || ""), "\"\n        data-label=\"").concat(escapeHtml(drawable.label || ""), "\"\n        data-name-memo=\"").concat(escapeHtml(drawable.nameMemo || ""), "\"\n        data-unit=\"").concat(escapeHtml(drawable.unit || ""), "\"\n        data-detail=\"").concat(escapeHtml(drawable.detail || ""), "\"\n        data-x=\"").concat(drawable.x, "\"\n        data-y=\"").concat(drawable.y, "\"\n        tabindex=\"0\"\n        role=\"button\"\n        aria-label=\"").concat(escapeHtml(ariaLabel), "\"\n      >\n        ").concat(shapeSvg, "\n        ").concat(hotspotSvg, "\n        <text x=\"").concat(labelX, "\" y=\"").concat(labelY, "\">").concat(escapeHtml(drawable.label || ""), "</text>\n      </g>\n    ");
}
function parseHexColor(colorRaw) {
  var fallback = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : "#6b7280";
  var text = value(colorRaw);
  var fallbackText = value(fallback) || "#6b7280";
  var normalized = /^#[0-9a-f]{3}$/i.test(text) ? "#".concat(text[1]).concat(text[1]).concat(text[2]).concat(text[2]).concat(text[3]).concat(text[3]) : /^#[0-9a-f]{6}$/i.test(text) ? text : fallbackText;
  var matched = normalized.match(/^#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i);
  if (!matched) {
    return {
      hex: "#6b7280",
      r: 107,
      g: 114,
      b: 128
    };
  }
  return {
    hex: "#".concat(matched[1]).concat(matched[2]).concat(matched[3]).toLowerCase(),
    r: Number.parseInt(matched[1], 16),
    g: Number.parseInt(matched[2], 16),
    b: Number.parseInt(matched[3], 16)
  };
}
function buildPlanImageWarpSvg(drawable) {
  var index = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : 0;
  var points = Array.isArray(drawable === null || drawable === void 0 ? void 0 : drawable.points) ? drawable.points : [];
  var imagePath = value(drawable === null || drawable === void 0 ? void 0 : drawable.imagePath);
  var useOriginalImageColor = Boolean(drawable === null || drawable === void 0 ? void 0 : drawable.useOriginalImageColor);
  var imageCandidates = getLargeShapeImagePathCandidates(drawable === null || drawable === void 0 ? void 0 : drawable.imageType, imagePath);
  var imageFallback = imageCandidates.find(function (candidate) {
    return !String(candidate).startsWith("data:");
  }) || imageCandidates[0] || imagePath;
  var imageRef = imageFallback;
  if (!useOriginalImageColor) {
    var dilateIterations = getImageShapeDilateIterations(drawable === null || drawable === void 0 ? void 0 : drawable.imageType);
    var tintedDataUrlRaw = getPlanLargeShapeTintedDataUrl(imageFallback, drawable === null || drawable === void 0 ? void 0 : drawable.color, drawable === null || drawable === void 0 ? void 0 : drawable.imageType, {
      dilateIterations: dilateIterations
    });
    var tintedDataUrl = value(tintedDataUrlRaw);
    var canUseTintedDataUrl = tintedDataUrl && tintedDataUrl.startsWith("data:") && tintedDataUrl.length <= PLAN_IMAGE_TINTED_DATA_URL_MAX_LENGTH;
    imageRef = canUseTintedDataUrl ? tintedDataUrl : imageFallback;
  }
  if (points.length !== 4 || !imageRef) {
    return "";
  }
  var _points6 = _slicedToArray(points, 4),
    p1 = _points6[0],
    p2 = _points6[1],
    p3 = _points6[2],
    p4 = _points6[3];
  var isParallelogram = Number.isFinite(p1.x) && Number.isFinite(p1.y) && Number.isFinite(p2.x) && Number.isFinite(p2.y) && Number.isFinite(p3.x) && Number.isFinite(p3.y) && Number.isFinite(p4.x) && Number.isFinite(p4.y) && Math.abs(p1.x + p3.x - (p2.x + p4.x)) <= 0.8 && Math.abs(p1.y + p3.y - (p2.y + p4.y)) <= 0.8;
  if (isParallelogram) {
    var matrix = [p2.x - p1.x, p2.y - p1.y, p4.x - p1.x, p4.y - p1.y, p1.x, p1.y];
    var matrixText = matrix.map(function (num) {
      return Number.isFinite(num) ? Number(num).toFixed(4) : "0";
    }).join(" ");
    return "<image href=\"".concat(escapeHtml(imageRef), "\" xlink:href=\"").concat(escapeHtml(imageRef), "\" x=\"0\" y=\"0\" width=\"1\" height=\"1\" preserveAspectRatio=\"none\" transform=\"matrix(").concat(matrixText, ")\" />");
  }
  var labelKey = value(drawable.label || "x").replace(/[^a-zA-Z0-9_-]/g, "");
  var clipIdA = "plan-img-clip-a-".concat(index, "-").concat(labelKey || "x");
  var clipIdB = "plan-img-clip-b-".concat(index, "-").concat(labelKey || "x");
  var matrixA = [p2.x - p1.x, p2.y - p1.y, p3.x - p2.x, p3.y - p2.y, p1.x, p1.y];
  var matrixB = [p3.x - p4.x, p3.y - p4.y, p4.x - p1.x, p4.y - p1.y, p1.x, p1.y];
  var triA = "".concat(p1.x, ",").concat(p1.y, " ").concat(p2.x, ",").concat(p2.y, " ").concat(p3.x, ",").concat(p3.y);
  var triB = "".concat(p1.x, ",").concat(p1.y, " ").concat(p3.x, ",").concat(p3.y, " ").concat(p4.x, ",").concat(p4.y);
  var matrixAText = matrixA.map(function (num) {
    return Number.isFinite(num) ? Number(num).toFixed(4) : "0";
  }).join(" ");
  var matrixBText = matrixB.map(function (num) {
    return Number.isFinite(num) ? Number(num).toFixed(4) : "0";
  }).join(" ");
  return "\n    <defs>\n      <clipPath id=\"".concat(clipIdA, "\">\n        <polygon points=\"").concat(triA, "\" />\n      </clipPath>\n      <clipPath id=\"").concat(clipIdB, "\">\n        <polygon points=\"").concat(triB, "\" />\n      </clipPath>\n    </defs>\n    <image href=\"").concat(escapeHtml(imageRef), "\" xlink:href=\"").concat(escapeHtml(imageRef), "\" x=\"0\" y=\"0\" width=\"1\" height=\"1\" preserveAspectRatio=\"none\" transform=\"matrix(").concat(matrixAText, ")\" clip-path=\"url(#").concat(clipIdA, ")\" />\n    <image href=\"").concat(escapeHtml(imageRef), "\" xlink:href=\"").concat(escapeHtml(imageRef), "\" x=\"0\" y=\"0\" width=\"1\" height=\"1\" preserveAspectRatio=\"none\" transform=\"matrix(").concat(matrixBText, ")\" clip-path=\"url(#").concat(clipIdB, ")\" />\n  ");
}
function getPlanDrawableExtent(drawable) {
  if (!drawable || _typeof(drawable) !== "object") {
    return null;
  }
  if (drawable.type === "line") {
    return {
      minX: Math.min(Number(drawable.x1), Number(drawable.x2)),
      maxX: Math.max(Number(drawable.x1), Number(drawable.x2)),
      minY: Math.min(Number(drawable.y1), Number(drawable.y2)),
      maxY: Math.max(Number(drawable.y1), Number(drawable.y2))
    };
  }
  if (drawable.type === "rect") {
    var _cx = Number(drawable.x);
    var _cy = Number(drawable.y);
    var w = Number(drawable.width);
    var h = Number(drawable.height);
    var halfDiag = Math.hypot(w / 2, h / 2);
    if (!Number.isFinite(_cx) || !Number.isFinite(_cy) || !Number.isFinite(halfDiag)) {
      return null;
    }
    return {
      minX: _cx - halfDiag,
      maxX: _cx + halfDiag,
      minY: _cy - halfDiag,
      maxY: _cy + halfDiag
    };
  }
  if (drawable.type === "ellipse") {
    var _cx2 = Number(drawable.x);
    var _cy2 = Number(drawable.y);
    var radius = Math.max(Number(drawable.rx), Number(drawable.ry));
    if (!Number.isFinite(_cx2) || !Number.isFinite(_cy2) || !Number.isFinite(radius)) {
      return null;
    }
    return {
      minX: _cx2 - radius,
      maxX: _cx2 + radius,
      minY: _cy2 - radius,
      maxY: _cy2 + radius
    };
  }
  if (drawable.type === "imageQuad" && Array.isArray(drawable.points) && drawable.points.length) {
    var xs = drawable.points.map(function (point) {
      return Number(point === null || point === void 0 ? void 0 : point.x);
    }).filter(function (num) {
      return Number.isFinite(num);
    });
    var ys = drawable.points.map(function (point) {
      return Number(point === null || point === void 0 ? void 0 : point.y);
    }).filter(function (num) {
      return Number.isFinite(num);
    });
    if (!xs.length || !ys.length) {
      return null;
    }
    return {
      minX: Math.min.apply(Math, _toConsumableArray(xs)),
      maxX: Math.max.apply(Math, _toConsumableArray(xs)),
      minY: Math.min.apply(Math, _toConsumableArray(ys)),
      maxY: Math.max.apply(Math, _toConsumableArray(ys))
    };
  }
  if (drawable.type === "multipoint") {
    var points = [].concat(_toConsumableArray(Array.isArray(drawable.points) ? drawable.points : []), _toConsumableArray(Array.isArray(drawable.hull) ? drawable.hull : []));
    var _xs = points.map(function (point) {
      return Number(point === null || point === void 0 ? void 0 : point.x);
    }).filter(function (num) {
      return Number.isFinite(num);
    });
    var _ys = points.map(function (point) {
      return Number(point === null || point === void 0 ? void 0 : point.y);
    }).filter(function (num) {
      return Number.isFinite(num);
    });
    if (!_xs.length || !_ys.length) {
      return null;
    }
    return {
      minX: Math.min.apply(Math, _toConsumableArray(_xs)),
      maxX: Math.max.apply(Math, _toConsumableArray(_xs)),
      minY: Math.min.apply(Math, _toConsumableArray(_ys)),
      maxY: Math.max.apply(Math, _toConsumableArray(_ys))
    };
  }
  var cx = Number(drawable.x);
  var cy = Number(drawable.y);
  if (!Number.isFinite(cx) || !Number.isFinite(cy)) {
    return null;
  }
  return {
    minX: cx,
    maxX: cx,
    minY: cy,
    maxY: cy
  };
}
function computePlanSvgViewBox() {
  var drawables = arguments.length > 0 && arguments[0] !== undefined ? arguments[0] : [];
  var base = {
    minX: 0,
    maxX: PLAN_SIZE_CM,
    minY: 0,
    maxY: PLAN_SIZE_CM
  };
  var extents = (Array.isArray(drawables) ? drawables : []).map(function (drawable) {
    return getPlanDrawableExtent(drawable);
  }).filter(Boolean);
  if (extents.length) {
    extents.forEach(function (extent) {
      base.minX = Math.min(base.minX, extent.minX);
      base.maxX = Math.max(base.maxX, extent.maxX);
      base.minY = Math.min(base.minY, extent.minY);
      base.maxY = Math.max(base.maxY, extent.maxY);
    });
  }
  var leftPad = 40;
  var rightPad = 24;
  var topPad = 24;
  var bottomPad = 24;
  var minX = Math.floor(base.minX - leftPad);
  var minY = Math.floor(base.minY - topPad);
  var width = Math.ceil(base.maxX + rightPad - minX);
  var height = Math.ceil(base.maxY + bottomPad - minY);
  return {
    minX: minX,
    minY: minY,
    width: Math.max(PLAN_SIZE_CM, width),
    height: Math.max(PLAN_SIZE_CM, height)
  };
}
function parseDistanceToCm(distanceRaw) {
  var text = value(distanceRaw).replace(",", ".");
  if (!text) {
    return null;
  }
  var matched = text.match(/-?\d+(?:\.\d+)?/);
  if (!matched) {
    return null;
  }
  var num = Number(matched[0]);
  return Number.isFinite(num) ? num : null;
}
function buildPlanCornerLabels(kuwakuRaw) {
  var parts = parseKuwaku(kuwakuLabelForSelect(kuwakuRaw));
  var block = value(parts.block).toUpperCase();
  var no = value(parts.no);
  if (!block || !no) {
    return {
      topLeft: "-",
      topRight: "-",
      bottomLeft: "-",
      bottomRight: "-"
    };
  }
  var rightBlock = incrementGridBlock(block, 1);
  var lowerNo = incrementGridNo(no, 1);
  return {
    topLeft: "".concat(block, "-").concat(no),
    topRight: "".concat(rightBlock, "-").concat(no),
    bottomLeft: "".concat(block, "-").concat(lowerNo),
    bottomRight: "".concat(rightBlock, "-").concat(lowerNo)
  };
}
function buildPlanCornerLabelsSvg(cornerLabels) {
  var labels = cornerLabels || {};
  var tl = escapeHtml(value(labels.topLeft) || "-");
  var tr = escapeHtml(value(labels.topRight) || "-");
  var bl = escapeHtml(value(labels.bottomLeft) || "-");
  var br = escapeHtml(value(labels.bottomRight) || "-");
  return "\n    <g class=\"plan-grid-corner-svg\">\n      <text x=\"4\" y=\"4\" text-anchor=\"start\" dominant-baseline=\"hanging\">".concat(tl, "</text>\n      <text x=\"").concat(PLAN_SIZE_CM - 4, "\" y=\"4\" text-anchor=\"end\" dominant-baseline=\"hanging\">").concat(tr, "</text>\n      <text x=\"4\" y=\"").concat(PLAN_SIZE_CM - 4, "\" text-anchor=\"start\" dominant-baseline=\"ideographic\">").concat(bl, "</text>\n      <text x=\"").concat(PLAN_SIZE_CM - 4, "\" y=\"").concat(PLAN_SIZE_CM - 4, "\" text-anchor=\"end\" dominant-baseline=\"ideographic\">").concat(br, "</text>\n    </g>\n  ");
}
function incrementGridBlock(blockRaw, step) {
  var block = value(blockRaw).toUpperCase();
  if (!/^[A-Z]+$/.test(block)) {
    return block;
  }
  if (/^[A-Z]$/.test(block)) {
    var base = block.charCodeAt(0) - 65;
    var _next2 = ((base + step) % 26 + 26) % 26;
    return String.fromCharCode(65 + _next2);
  }
  var colNumber = 0;
  var _iterator5 = _createForOfIteratorHelper(block),
    _step7;
  try {
    for (_iterator5.s(); !(_step7 = _iterator5.n()).done;) {
      var _char = _step7.value;
      colNumber = colNumber * 26 + (_char.charCodeAt(0) - 64);
    }
  } catch (err) {
    _iterator5.e(err);
  } finally {
    _iterator5.f();
  }
  colNumber += step;
  if (colNumber <= 0) {
    return block;
  }
  var next = "";
  var current = colNumber;
  while (current > 0) {
    var remainder = (current - 1) % 26;
    next = String.fromCharCode(65 + remainder) + next;
    current = Math.floor((current - 1) / 26);
  }
  return next;
}
function incrementGridNo(noRaw, step) {
  var raw = value(noRaw);
  if (!/^-?\d+$/.test(raw)) {
    return raw;
  }
  return String(Number(raw) + step);
}
function buildPlanLegendHtml() {
  var order = ["m", "b", "l", "s", "i", "g", "h", "a"];
  return order.map(function (prefix) {
    var color = getSpecimenPrefixColor(prefix);
    var label = SPECIMEN_CATEGORY_MAP[prefix] || "";
    return "<span class=\"plan-legend-item\"><span class=\"plan-legend-dot\" style=\"background:".concat(color, "\"></span>").concat(prefix, ": ").concat(label, "</span>");
  }).join("");
}
function getSpecimenPrefixColor(prefixRaw) {
  var prefix = normalizeSpecimenPrefix(prefixRaw);
  return SPECIMEN_POINT_COLORS[prefix] || SPECIMEN_POINT_COLORS.h;
}
function getRecordSpecimenColor(record) {
  var specimen = parseSpecimenNo(record === null || record === void 0 ? void 0 : record.specimenNo, record === null || record === void 0 ? void 0 : record.specimenPrefix, record === null || record === void 0 ? void 0 : record.specimenSerial);
  return getSpecimenPrefixColor(specimen.prefix);
}
function getKuwakuCellStyle(kuwakuRaw) {
  var kuwaku = normalizeKuwakuText(kuwakuRaw);
  if (!kuwaku) {
    return {
      background: "#f3f4f6",
      border: "#d1d5db",
      color: "#111827"
    };
  }
  var parts = parseKuwaku(kuwaku);
  var blockIndex = blockLabelToIndex(parts.block);
  var noSeed = /^-?\d+$/.test(parts.no) ? Number(parts.no) : hashText(parts.no || kuwaku) % 97 + 1;
  var headSeed = hashText("".concat(parts.headA, "-").concat(parts.headB)) % 360;
  var hue = ((blockIndex * 41 + noSeed * 17 + headSeed) % 360 + 360) % 360;
  var sat = 66;
  var bgLightness = 93;
  var borderLightness = 82;
  return {
    background: "hsl(".concat(hue, ", ").concat(sat, "%, ").concat(bgLightness, "%)"),
    border: "hsl(".concat(hue, ", 48%, ").concat(borderLightness, "%)"),
    color: "#111827"
  };
}
function getUnitCellStyle(unitRaw) {
  var normalized = compactNoSpaceValue(unitRaw).toUpperCase();
  if (UNIT_CELL_COLOR_MAP[normalized]) {
    return UNIT_CELL_COLOR_MAP[normalized];
  }
  if (!normalized) {
    return {
      background: "#f3f4f6",
      border: "#d1d5db",
      color: "#111827"
    };
  }
  var hue = hashText(normalized) % 360;
  return {
    background: "hsl(".concat(hue, ", 58%, 93%)"),
    border: "hsl(".concat(hue, ", 38%, 80%)"),
    color: "#111827"
  };
}
function blockLabelToIndex(blockRaw) {
  var block = normalizeKuwakuBlock(blockRaw);
  if (!block || !/^[A-Z]+$/.test(block)) {
    return hashText(block) % 26 + 1;
  }
  var index = 0;
  var _iterator6 = _createForOfIteratorHelper(block),
    _step8;
  try {
    for (_iterator6.s(); !(_step8 = _iterator6.n()).done;) {
      var _char2 = _step8.value;
      index = index * 26 + (_char2.charCodeAt(0) - 64);
    }
  } catch (err) {
    _iterator6.e(err);
  } finally {
    _iterator6.f();
  }
  return index;
}
function hashText(textRaw) {
  var text = value(textRaw);
  var hash = 0;
  for (var i = 0; i < text.length; i += 1) {
    hash = hash * 31 + text.charCodeAt(i) >>> 0;
  }
  return hash;
}
function toRgbaColor(hexColorRaw, alphaRaw) {
  var hexColor = value(hexColorRaw).replace("#", "");
  var alpha = clamp(Number(alphaRaw), 0, 1);
  if (!/^[0-9a-fA-F]{6}$/.test(hexColor)) {
    return "rgba(107, 114, 128, ".concat(alpha, ")");
  }
  var r = Number.parseInt(hexColor.slice(0, 2), 16);
  var g = Number.parseInt(hexColor.slice(2, 4), 16);
  var b = Number.parseInt(hexColor.slice(4, 6), 16);
  return "rgba(".concat(r, ", ").concat(g, ", ").concat(b, ", ").concat(alpha, ")");
}
function attachPlanMapTooltips() {
  if (!planMapWrap) {
    return;
  }
  var shell = planMapWrap.querySelector(".plan-map-shell");
  var svg = shell === null || shell === void 0 ? void 0 : shell.querySelector(".plan-map-svg");
  var tooltip = shell === null || shell === void 0 ? void 0 : shell.querySelector(".plan-map-tooltip");
  if (!shell || !svg || !tooltip) {
    return;
  }
  var hide = function hide() {
    tooltip.hidden = true;
  };
  var show = function show(pointEl) {
    var mouseEvent = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : null;
    var specimenNo = value(pointEl.dataset.label) || "未設定";
    var nameMemo = value(pointEl.dataset.nameMemo) || "未設定";
    var unit = value(pointEl.dataset.unit) || "未設定";
    var detail = value(pointEl.dataset.detail) || "未設定";
    tooltip.innerHTML = "\n      <div><strong>\u6A19\u672C\u756A\u53F7:</strong> ".concat(escapeHtml(specimenNo), "</div>\n      <div><strong>\u5316\u77F3\u30FB\u907A\u7269\u540D\u79F0:</strong> ").concat(escapeHtml(nameMemo), "</div>\n      <div><strong>\u30E6\u30CB\u30C3\u30C8:</strong> ").concat(escapeHtml(unit), "</div>\n      <div><strong>\u30B5\u30D6\u30E6\u30CB\u30C3\u30C8:</strong> ").concat(escapeHtml(detail), "</div>\n    ");
    tooltip.hidden = false;
    positionTooltip(pointEl, mouseEvent, shell, svg, tooltip);
  };
  var points = shell.querySelectorAll(".plan-point-group");
  points.forEach(function (pointEl) {
    var touchLongPressState = {
      pointerId: null,
      startX: 0,
      startY: 0,
      timer: 0
    };
    var clearTouchLongPress = function clearTouchLongPress() {
      if (touchLongPressState.timer) {
        window.clearTimeout(touchLongPressState.timer);
      }
      touchLongPressState.timer = 0;
      touchLongPressState.pointerId = null;
    };
    pointEl.addEventListener("mouseenter", function (event) {
      return show(pointEl, event);
    });
    pointEl.addEventListener("mousemove", function (event) {
      if (!tooltip.hidden) {
        positionTooltip(pointEl, event, shell, svg, tooltip);
      }
    });
    pointEl.addEventListener("mouseleave", hide);
    pointEl.addEventListener("focus", function () {
      return show(pointEl);
    });
    pointEl.addEventListener("blur", hide);
    pointEl.addEventListener("click", function (event) {
      event.stopPropagation();
      show(pointEl, event);
    });
    pointEl.addEventListener("contextmenu", function (event) {
      event.preventDefault();
      event.stopPropagation();
      var recordId = value(pointEl.dataset.id);
      if (!recordId) {
        hideHoverEditMenu();
        return;
      }
      show(pointEl, event);
      showHoverEditMenu(event.clientX, event.clientY, recordId, value(pointEl.dataset.kuwaku), value(pointEl.dataset.label));
    });
    pointEl.addEventListener("pointerdown", function (event) {
      if (!isTouchLikePointerEvent(event)) {
        return;
      }
      var recordId = value(pointEl.dataset.id);
      if (!recordId) {
        return;
      }
      clearTouchLongPress();
      touchLongPressState.pointerId = Number.isFinite(Number(event.pointerId)) ? event.pointerId : null;
      touchLongPressState.startX = Number(event.clientX) || 0;
      touchLongPressState.startY = Number(event.clientY) || 0;
      touchLongPressState.timer = window.setTimeout(function () {
        if (touchLongPressState.pointerId == null) {
          return;
        }
        touchLongPressState.timer = 0;
        show(pointEl, {
          clientX: touchLongPressState.startX,
          clientY: touchLongPressState.startY
        });
        showHoverEditMenu(touchLongPressState.startX, touchLongPressState.startY, recordId, value(pointEl.dataset.kuwaku), value(pointEl.dataset.label));
      }, TOUCH_LONG_PRESS_MS);
    });
    pointEl.addEventListener("pointermove", function (event) {
      if (!isTouchLikePointerEvent(event) || touchLongPressState.pointerId == null) {
        return;
      }
      if (Number(event.pointerId) !== Number(touchLongPressState.pointerId)) {
        return;
      }
      var moved = pointerMovedBeyondThreshold(event.clientX, event.clientY, touchLongPressState.startX, touchLongPressState.startY, TOUCH_LONG_PRESS_MOVE_THRESHOLD_PX);
      if (moved) {
        clearTouchLongPress();
      }
    });
    pointEl.addEventListener("pointerup", function (event) {
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
  shell.addEventListener("click", function (event) {
    if (event.target.closest(".plan-point-group")) {
      return;
    }
    hide();
    hideHoverEditMenu();
  });
  shell.addEventListener("contextmenu", function (event) {
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
  var menu = document.createElement("div");
  menu.className = "hover-edit-menu";
  menu.hidden = true;
  menu.innerHTML = "\n    <div class=\"hover-edit-menu-title\">\u3053\u306E\u30C7\u30FC\u30BF\u3092\u7DE8\u96C6</div>\n    <button type=\"button\" class=\"hover-edit-menu-button\">\u7DE8\u96C6</button>\n  ";
  var button = menu.querySelector(".hover-edit-menu-button");
  if (button) {
    button.addEventListener("click", function (event) {
      event.preventDefault();
      event.stopPropagation();
      var recordId = value(hoverEditMenuRecordId);
      if (!recordId) {
        hideHoverEditMenu();
        return;
      }
      var preferredKuwaku = value(hoverEditMenuKuwaku);
      hideHoverEditMenu();
      openRecordForEdit(recordId, preferredKuwaku);
    });
  }
  menu.addEventListener("click", function (event) {
    event.stopPropagation();
  });
  document.body.appendChild(menu);
  hoverEditMenuEl = menu;
  return hoverEditMenuEl;
}
function showHoverEditMenu(clientXRaw, clientYRaw, recordIdRaw) {
  var kuwakuRaw = arguments.length > 3 && arguments[3] !== undefined ? arguments[3] : "";
  var labelRaw = arguments.length > 4 && arguments[4] !== undefined ? arguments[4] : "";
  var recordId = value(recordIdRaw);
  if (!recordId) {
    hideHoverEditMenu();
    return;
  }
  var menu = ensureHoverEditMenu();
  hoverEditMenuRecordId = recordId;
  hoverEditMenuKuwaku = value(kuwakuRaw);
  var title = menu.querySelector(".hover-edit-menu-title");
  var labelText = value(labelRaw);
  if (title) {
    title.textContent = labelText ? "".concat(labelText, " \u3092\u7DE8\u96C6") : "このデータを編集";
  }
  menu.hidden = false;
  var clientX = Number(clientXRaw);
  var clientY = Number(clientYRaw);
  var fallbackX = Math.max(0, Math.floor(window.innerWidth / 2));
  var fallbackY = Math.max(0, Math.floor(window.innerHeight / 2));
  var anchorX = Number.isFinite(clientX) ? clientX : fallbackX;
  var anchorY = Number.isFinite(clientY) ? clientY : fallbackY;
  var margin = 8;
  var offset = 12;
  var rect = menu.getBoundingClientRect();
  var maxLeft = Math.max(margin, window.innerWidth - rect.width - margin);
  var maxTop = Math.max(margin, window.innerHeight - rect.height - margin);
  var left = clamp(anchorX + offset, margin, maxLeft);
  var top = clamp(anchorY + offset, margin, maxTop);
  menu.style.left = "".concat(left, "px");
  menu.style.top = "".concat(top, "px");
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
  var pointerType = value(event === null || event === void 0 ? void 0 : event.pointerType).toLowerCase();
  return pointerType === "touch" || pointerType === "pen";
}
function pointerMovedBeyondThreshold(clientXRaw, clientYRaw, startXRaw, startYRaw) {
  var thresholdRaw = arguments.length > 4 && arguments[4] !== undefined ? arguments[4] : 12;
  var clientX = Number(clientXRaw);
  var clientY = Number(clientYRaw);
  var startX = Number(startXRaw);
  var startY = Number(startYRaw);
  if (!Number.isFinite(clientX) || !Number.isFinite(clientY) || !Number.isFinite(startX) || !Number.isFinite(startY)) {
    return false;
  }
  var threshold = Math.max(1, Number(thresholdRaw) || 1);
  return Math.hypot(clientX - startX, clientY - startY) > threshold;
}
function positionTooltip(pointEl, mouseEvent, shell, svg, tooltip) {
  var shellRect = shell.getBoundingClientRect();
  var xLocal = 0;
  var yLocal = 0;
  if (mouseEvent && typeof mouseEvent.clientX === "number" && typeof mouseEvent.clientY === "number") {
    xLocal = mouseEvent.clientX - shellRect.left;
    yLocal = mouseEvent.clientY - shellRect.top;
  } else {
    var pointRect = pointEl.getBoundingClientRect();
    xLocal = pointRect.left - shellRect.left + pointRect.width / 2;
    yLocal = pointRect.top - shellRect.top + pointRect.height / 2;
  }
  var offset = 14;
  var desiredLeft = xLocal + offset;
  var desiredTop = yLocal + offset;
  var maxLeft = Math.max(8, shellRect.width - tooltip.offsetWidth - 8);
  var maxTop = Math.max(8, shellRect.height - tooltip.offsetHeight - 8);
  tooltip.style.left = "".concat(clamp(desiredLeft, 8, maxLeft), "px");
  tooltip.style.top = "".concat(clamp(desiredTop, 8, maxTop), "px");
}
function renderPhotoList() {
  if (!currentPhotos.length) {
    photoList.innerHTML = "<p>写真はまだありません。</p>";
    return;
  }
  photoList.innerHTML = currentPhotos.map(function (photo) {
    return "\n      <article class=\"photo-card\">\n        <img src=\"".concat(photo.dataUrl, "\" alt=\"").concat(escapeHtml(photo.name || "photo"), "\" />\n        <input\n          class=\"caption\"\n          data-photo-id=\"").concat(photo.id, "\"\n          type=\"text\"\n          placeholder=\"\u5199\u771F\u30AD\u30E3\u30D7\u30B7\u30E7\u30F3\"\n          value=\"").concat(escapeHtml(photo.caption || ""), "\"\n        />\n        <div class=\"panel-actions\">\n          <button class=\"danger\" type=\"button\" data-remove-photo-id=\"").concat(photo.id, "\">\u5199\u771F\u524A\u9664</button>\n        </div>\n      </article>\n      ");
  }).join("");
}
function renderSectionDiagramList() {
  if (!currentSectionDiagrams.length) {
    sectionDiagramList.innerHTML = "<p>断面図はまだありません。</p>";
    return;
  }
  sectionDiagramList.innerHTML = currentSectionDiagrams.map(function (item) {
    return "\n      <article class=\"photo-card\">\n        <img src=\"".concat(item.dataUrl, "\" alt=\"").concat(escapeHtml(item.name || "diagram"), "\" />\n        <input\n          class=\"caption\"\n          data-diagram-id=\"").concat(item.id, "\"\n          type=\"text\"\n          placeholder=\"\u65AD\u9762\u56F3\u30AD\u30E3\u30D7\u30B7\u30E7\u30F3\"\n          value=\"").concat(escapeHtml(item.caption || ""), "\"\n        />\n        <div class=\"panel-actions\">\n          <button class=\"danger\" type=\"button\" data-remove-diagram-id=\"").concat(item.id, "\">\u65AD\u9762\u56F3\u524A\u9664</button>\n        </div>\n      </article>\n      ");
  }).join("");
}
function loadState() {
  var raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) {
    stateNeedsRewriteAfterLoad = false;
    return createInitialState();
  }
  try {
    var parsed = JSON.parse(raw);
    var _normalized = normalizeState(parsed);
    stateNeedsRewriteAfterLoad = hasSpacingNormalizationDiff(parsed, _normalized);
    return _normalized;
  } catch (_error) {
    stateNeedsRewriteAfterLoad = false;
    return createInitialState();
  }
}
function buildSpacingNormalizationFingerprint(candidateState) {
  var stateCandidate = candidateState && _typeof(candidateState) === "object" ? candidateState : {};
  var site = stateCandidate.site && _typeof(stateCandidate.site) === "object" ? stateCandidate.site : {};
  var records = Array.isArray(stateCandidate.records) ? stateCandidate.records : [];
  return {
    site: {
      kuwaku: value(site.kuwaku),
      kuwakuHeadA: value(site.kuwakuHeadA),
      kuwakuHeadB: value(site.kuwakuHeadB),
      kuwakuBlock: value(site.kuwakuBlock),
      kuwakuNo: value(site.kuwakuNo)
    },
    records: records.map(function (record) {
      return {
        id: value(record === null || record === void 0 ? void 0 : record.id),
        kuwaku: value(record === null || record === void 0 ? void 0 : record.kuwaku),
        specimenNo: value(record === null || record === void 0 ? void 0 : record.specimenNo),
        specimenPrefix: value(record === null || record === void 0 ? void 0 : record.specimenPrefix),
        specimenSerial: value(record === null || record === void 0 ? void 0 : record.specimenSerial),
        unit: value(record === null || record === void 0 ? void 0 : record.unit),
        detail: value(record === null || record === void 0 ? void 0 : record.detail)
      };
    })
  };
}
function hasSpacingNormalizationDiff(beforeState, afterState) {
  try {
    return JSON.stringify(buildSpacingNormalizationFingerprint(beforeState)) !== JSON.stringify(buildSpacingNormalizationFingerprint(afterState));
  } catch (_error) {
    return false;
  }
}
function ensureUniqueRecordIds(recordsRaw) {
  if (!Array.isArray(recordsRaw)) {
    return [];
  }
  var seenIds = new Set();
  return recordsRaw.map(function (recordRaw) {
    if (!recordRaw || _typeof(recordRaw) !== "object") {
      return recordRaw;
    }
    var record = _objectSpread({}, recordRaw);
    var currentId = value(record.id);
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
  var _candidate$site, _candidate$site2, _candidate$site3, _candidate$site4, _candidate$site5, _candidate$site6, _candidate$site7, _candidate$site8, _candidate$site9, _candidate$site0, _candidate$site1, _candidate$site10;
  var safe = createInitialState();
  if (!candidate || _typeof(candidate) !== "object") {
    return safe;
  }
  var kuwakuParts = parseKuwaku(value((_candidate$site = candidate.site) === null || _candidate$site === void 0 ? void 0 : _candidate$site.kuwaku));
  var kuwakuHeadA = normalizeKuwakuHeadA(value((_candidate$site2 = candidate.site) === null || _candidate$site2 === void 0 ? void 0 : _candidate$site2.kuwakuHeadA) || kuwakuParts.headA || DEFAULT_KUWAKU_HEAD_A);
  var kuwakuHeadB = normalizeKuwakuHeadB(value((_candidate$site3 = candidate.site) === null || _candidate$site3 === void 0 ? void 0 : _candidate$site3.kuwakuHeadB) || kuwakuParts.headB || DEFAULT_KUWAKU_HEAD_B);
  var kuwakuBlock = normalizeKuwakuBlock(value((_candidate$site4 = candidate.site) === null || _candidate$site4 === void 0 ? void 0 : _candidate$site4.kuwakuBlock) || kuwakuParts.block);
  var kuwakuNo = normalizeKuwakuNo(value((_candidate$site5 = candidate.site) === null || _candidate$site5 === void 0 ? void 0 : _candidate$site5.kuwakuNo) || kuwakuParts.no);
  var teamState = normalizeTeamState(value((_candidate$site6 = candidate.site) === null || _candidate$site6 === void 0 ? void 0 : _candidate$site6.team), value((_candidate$site7 = candidate.site) === null || _candidate$site7 === void 0 ? void 0 : _candidate$site7.teamOther));
  var candidateSite = candidate.site && _typeof(candidate.site) === "object" ? candidate.site : {};
  var hasSavedRecords = Array.isArray(candidate.records) && candidate.records.length > 0;
  var isLegacyEmptyDefault = kuwakuHeadA === "24" && !kuwakuBlock && !kuwakuNo && !value(candidateSite.levelHeight) && !value(candidateSite.date) && !value(candidateSite.team) && !value(candidateSite.teamLead) && !value(candidateSite.recorder) && !value(candidateSite.scribe) && !hasSavedRecords;
  if (isLegacyEmptyDefault) {
    kuwakuHeadA = DEFAULT_KUWAKU_HEAD_A;
  }
  safe.site = {
    kuwaku: buildKuwaku(kuwakuHeadA, kuwakuHeadB, kuwakuBlock, kuwakuNo),
    kuwakuHeadA: kuwakuHeadA,
    kuwakuHeadB: kuwakuHeadB,
    kuwakuBlock: kuwakuBlock,
    kuwakuNo: kuwakuNo,
    levelHeight: value((_candidate$site8 = candidate.site) === null || _candidate$site8 === void 0 ? void 0 : _candidate$site8.levelHeight),
    date: value((_candidate$site9 = candidate.site) === null || _candidate$site9 === void 0 ? void 0 : _candidate$site9.date),
    team: teamState.team,
    teamOther: teamState.teamOther,
    teamLead: value((_candidate$site0 = candidate.site) === null || _candidate$site0 === void 0 ? void 0 : _candidate$site0.teamLead),
    recorder: value((_candidate$site1 = candidate.site) === null || _candidate$site1 === void 0 ? void 0 : _candidate$site1.recorder),
    scribe: value((_candidate$site10 = candidate.site) === null || _candidate$site10 === void 0 ? void 0 : _candidate$site10.scribe)
  };
  if (Array.isArray(candidate.records)) {
    safe.records = ensureUniqueRecordIds(candidate.records.map(function (item) {
      return normalizeRecord(item, safe.site);
    }).filter(Boolean));
    return safe;
  }
  var artifacts = Array.isArray(candidate.artifacts) ? candidate.artifacts : [];
  var cards = candidate.cards && _typeof(candidate.cards) === "object" ? candidate.cards : {};
  var photos = candidate.photos && _typeof(candidate.photos) === "object" ? candidate.photos : {};
  safe.records = ensureUniqueRecordIds(artifacts.map(function (artifact) {
    var _candidate$site11, _artifact$categories, _candidate$site12, _candidate$site13, _candidate$site14, _candidate$site15, _candidate$site16, _candidate$site17;
    if (!artifact || _typeof(artifact) !== "object") {
      return null;
    }
    var id = value(artifact.id);
    if (!id) {
      return null;
    }
    var card = cards[id] && _typeof(cards[id]) === "object" ? cards[id] : {};
    var recordPhotos = Array.isArray(photos[id]) ? photos[id] : [];
    return normalizeRecord({
      id: id,
      kuwaku: value(artifact.kuwaku) || value((_candidate$site11 = candidate.site) === null || _candidate$site11 === void 0 ? void 0 : _candidate$site11.kuwaku),
      specimenNo: value(artifact.specimenNo),
      specimenPrefix: value(artifact.specimenPrefix),
      specimenSerial: value(artifact.specimenSerial),
      category: value(artifact.category) || value((_artifact$categories = artifact.categories) === null || _artifact$categories === void 0 ? void 0 : _artifact$categories[0]) || categoryFromPrefix(value(artifact.specimenPrefix)),
      analysisType: value(artifact.analysisType) || value(card.analysisType) || extractAnalysisTypeFromCategory(value(artifact.category)),
      levelHeight: value(artifact.levelHeight) || value((_candidate$site12 = candidate.site) === null || _candidate$site12 === void 0 ? void 0 : _candidate$site12.levelHeight),
      date: value(artifact.date) || value((_candidate$site13 = candidate.site) === null || _candidate$site13 === void 0 ? void 0 : _candidate$site13.date),
      team: value(artifact.team) || value((_candidate$site14 = candidate.site) === null || _candidate$site14 === void 0 ? void 0 : _candidate$site14.team),
      teamOther: value(artifact.teamOther) || value((_candidate$site15 = candidate.site) === null || _candidate$site15 === void 0 ? void 0 : _candidate$site15.teamOther),
      teamLead: value(artifact.teamLead) || value((_candidate$site16 = candidate.site) === null || _candidate$site16 === void 0 ? void 0 : _candidate$site16.teamLead),
      recorder: value(artifact.recorder) || value((_candidate$site17 = candidate.site) === null || _candidate$site17 === void 0 ? void 0 : _candidate$site17.recorder),
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
      layerFacies: value(card.layerFacies),
      layerRef: value(card.layerRef) || value(card.layerPosition),
      layerFromCm: value(card.layerFromCm),
      layerRelative: value(card.layerRelative),
      notes: mergeLegacyNotes({
        notes: value(artifact.notes),
        occurrenceNote: value(card.occurrenceNote),
        sketchNote: value(card.sketchNote)
      }),
      sectionDiagrams: normalizePhotos(card.sectionDiagrams),
      photos: recordPhotos,
      createdAt: value(artifact.createdAt),
      updatedAt: value(artifact.updatedAt)
    }, safe.site);
  }).filter(Boolean));
  return safe;
}
function normalizeRecord(item) {
  var fallbackSiteRaw = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : null;
  if (!item || _typeof(item) !== "object") {
    return null;
  }
  var fallbackSite = fallbackSiteRaw && _typeof(fallbackSiteRaw) === "object" ? fallbackSiteRaw : {};
  var id = value(item.id) || newId("record");
  var parsedSpecimen = parseSpecimenNo(value(item.specimenNo), value(item.specimenPrefix), value(item.specimenSerial));
  var category = normalizeCategory(value(item.category), parsedSpecimen.prefix);
  var analysisType = normalizeAnalysisType(value(item.analysisType) || extractAnalysisTypeFromCategory(value(item.category)));
  var teamState = normalizeTeamState(value(item.team) || value(fallbackSite.team), value(item.teamOther) || value(fallbackSite.teamOther));
  var rawKuwaku = normalizeKuwakuText(item.kuwaku);
  var fallbackKuwaku = normalizeKuwakuText(value(fallbackSite.kuwaku) || buildKuwaku(fallbackSite.kuwakuHeadA, fallbackSite.kuwakuHeadB, fallbackSite.kuwakuBlock, fallbackSite.kuwakuNo));
  var kuwaku = !rawKuwaku || isDefaultKuwaku(rawKuwaku) ? fallbackKuwaku : rawKuwaku;
  var rawLargeShapeType = value(item.largeShapeType);
  var normalizedLargeShapeType = normalizeLargeShapeType(rawLargeShapeType) || normalizeLargeShapeLabel(rawLargeShapeType);
  return {
    id: id,
    kuwaku: kuwaku,
    specimenPrefix: parsedSpecimen.prefix,
    specimenSerial: parsedSpecimen.serial,
    specimenNo: parsedSpecimen.specimenNo,
    category: category,
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
    multiPoints: normalizePlanMultiPoints(item.multiPoints),
    planSizeMode: normalizePlanSizeMode(value(item.planSizeMode)),
    largeShapeType: normalizedLargeShapeType,
    largeAxisDirection: normalizeLargeAxisDirection(value(item.largeAxisDirection)),
    largeAxisPlungeDeg: normalizeLargeAxisPlungeDeg(value(item.largeAxisPlungeDeg)),
    largeAxisPlungeDir8: normalizeCompass8Direction(value(item.largeAxisPlungeDir8) || value(item.largeAxisPlungeDirection) || value(item.plungeDir8)),
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
    layerFacies: value(item.layerFacies),
    layerRef: value(item.layerRef) || value(item.layerPosition),
    layerFromCm: value(item.layerFromCm),
    layerRelative: value(item.layerRelative),
    notes: mergeLegacyNotes({
      notes: value(item.notes),
      occurrenceNote: value(item.occurrenceNote),
      sketchNote: value(item.sketchNote)
    }),
    sectionDiagrams: normalizePhotos(item.sectionDiagrams),
    photos: normalizePhotos(item.photos),
    history: normalizeRecordHistory(item.history),
    createdAt: value(item.createdAt) || new Date().toISOString(),
    updatedAt: value(item.updatedAt) || new Date().toISOString()
  };
  var isLineShape = normalized.largeShapeType === "直線状";
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
    var p1 = convertPositionToPlanCoords(normalized.line1NsDir, normalized.line1NsCm, normalized.line1EwDir, normalized.line1EwCm);
    var p2 = convertPositionToPlanCoords(normalized.line2NsDir, normalized.line2NsCm, normalized.line2EwDir, normalized.line2EwCm);
    if (p1 && p2) {
      var distance = Math.hypot(p2.x - p1.x, p2.y - p1.y);
      if (Number.isFinite(distance) && distance > 0) {
        normalized.lineLengthCm = trimNumericText(distance.toFixed(1));
      }
    }
  }
  return normalized;
}
function buildNextRecordHistory(previousRecord, nextRecord, actionRaw) {
  var previousHistory = normalizeRecordHistory(previousRecord === null || previousRecord === void 0 ? void 0 : previousRecord.history);
  var previousSnapshot = previousRecord ? createHistorySnapshot(previousRecord) : null;
  var snapshot = createHistorySnapshot(nextRecord);
  var entry = {
    id: newId("history"),
    action: value(actionRaw) || "保存",
    content: buildHistoryContent(nextRecord, snapshot),
    snapshot: snapshot,
    changedKeys: getHistoryChangedKeys(previousSnapshot, snapshot),
    at: nowIso()
  };
  return [].concat(_toConsumableArray(previousHistory), [entry]).slice(-50);
}
function createHistorySnapshot(record) {
  return {
    specimenNo: value(record === null || record === void 0 ? void 0 : record.specimenNo),
    nameMemo: value(record === null || record === void 0 ? void 0 : record.nameMemo),
    category: formatCategoryForRecord(record),
    layerName: value(record === null || record === void 0 ? void 0 : record.layerName),
    unit: value(record === null || record === void 0 ? void 0 : record.unit),
    detail: formatDetailForRecord(record),
    layerFacies: value(record === null || record === void 0 ? void 0 : record.layerFacies),
    layerPosition: formatLayerPosition(record)
  };
}
function buildHistoryContent(record) {
  var snapshotRaw = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : null;
  var snapshot = snapshotRaw || createHistorySnapshot(record);
  var summaryParts = HISTORY_SNAPSHOT_FIELDS.map(function (field) {
    var fieldValue = value(snapshot === null || snapshot === void 0 ? void 0 : snapshot[field.key]) || "-";
    return "".concat(field.label, " ").concat(fieldValue);
  });
  return summaryParts.join(" / ");
}
function normalizeRecordHistory(historyRaw) {
  if (!Array.isArray(historyRaw)) {
    return [];
  }
  return historyRaw.filter(function (entry) {
    return entry && _typeof(entry) === "object";
  }).map(function (entry) {
    return {
      id: value(entry.id) || newId("history"),
      action: value(entry.action) || "保存",
      content: value(entry.content),
      snapshot: normalizeHistorySnapshot(entry.snapshot) || extractHistorySnapshotFromContent(value(entry.content)),
      changedKeys: normalizeHistoryChangedKeys(entry.changedKeys),
      at: value(entry.at) || nowIso()
    };
  }).filter(function (entry) {
    return entry.content || entry.snapshot;
  });
}
function normalizeHistorySnapshot(snapshotRaw) {
  if (!snapshotRaw || _typeof(snapshotRaw) !== "object") {
    return null;
  }
  var snapshot = {};
  HISTORY_SNAPSHOT_FIELDS.forEach(function (field) {
    snapshot[field.key] = value(snapshotRaw[field.key]);
  });
  return snapshot;
}
function normalizeHistoryChangedKeys(changedKeysRaw) {
  if (!Array.isArray(changedKeysRaw)) {
    return [];
  }
  var seen = new Set();
  return changedKeysRaw.map(function (key) {
    return value(key);
  }).filter(function (key) {
    return HISTORY_SNAPSHOT_FIELD_KEYS.has(key);
  }).filter(function (key) {
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
  return HISTORY_SNAPSHOT_FIELDS.map(function (field) {
    return field.key;
  }).filter(function (key) {
    return historyComparableValue(currentSnapshot === null || currentSnapshot === void 0 ? void 0 : currentSnapshot[key]) !== historyComparableValue(previousSnapshot === null || previousSnapshot === void 0 ? void 0 : previousSnapshot[key]);
  });
}
function extractHistorySnapshotFromContent(contentRaw) {
  var content = value(contentRaw);
  if (!content) {
    return null;
  }
  var parts = content.split(/\s*\/\s*/);
  var snapshot = {};
  HISTORY_SNAPSHOT_FIELDS.forEach(function (field) {
    var part = parts.find(function (item) {
      return value(item).startsWith(field.label);
    });
    if (!part) {
      snapshot[field.key] = "";
      return;
    }
    var stripped = value(part).slice(field.label.length).replace(/^[:：]?\s*/, "");
    snapshot[field.key] = value(stripped);
  });
  var hasAny = HISTORY_SNAPSHOT_FIELDS.some(function (field) {
    return value(snapshot[field.key]);
  });
  return hasAny ? snapshot : null;
}
function renderHistoryContentHtml(entry, prevEntry) {
  var snapshot = (entry === null || entry === void 0 ? void 0 : entry.snapshot) || extractHistorySnapshotFromContent(value(entry === null || entry === void 0 ? void 0 : entry.content));
  if (!snapshot) {
    return escapeHtml((entry === null || entry === void 0 ? void 0 : entry.content) || "");
  }
  var changedKeys = normalizeHistoryChangedKeys(entry === null || entry === void 0 ? void 0 : entry.changedKeys);
  var changedKeySet = new Set(changedKeys);
  var hasExplicitChangedKeys = changedKeySet.size > 0;
  var prevSnapshot = (prevEntry === null || prevEntry === void 0 ? void 0 : prevEntry.snapshot) || extractHistorySnapshotFromContent(value(prevEntry === null || prevEntry === void 0 ? void 0 : prevEntry.content));
  return HISTORY_SNAPSHOT_FIELDS.map(function (field) {
    var currentValueRaw = value(snapshot[field.key]);
    var currentValue = currentValueRaw || "-";
    var isChanged = hasExplicitChangedKeys ? changedKeySet.has(field.key) : Boolean(prevSnapshot) && historyComparableValue(currentValueRaw) !== historyComparableValue(prevSnapshot === null || prevSnapshot === void 0 ? void 0 : prevSnapshot[field.key]);
    var className = isChanged ? "edit-history-value changed" : "edit-history-value";
    return "<span class=\"".concat(className, "\">").concat(escapeHtml(field.label), ": ").concat(escapeHtml(currentValue), "</span>");
  }).join(" / ");
}
function formatHistoryDateTime(isoRaw) {
  var iso = value(isoRaw);
  if (!iso) {
    return "-";
  }
  var date = new Date(iso);
  if (Number.isNaN(date.getTime())) {
    return iso;
  }
  var y = date.getFullYear();
  var m = String(date.getMonth() + 1).padStart(2, "0");
  var d = String(date.getDate()).padStart(2, "0");
  var hh = String(date.getHours()).padStart(2, "0");
  var mm = String(date.getMinutes()).padStart(2, "0");
  return "".concat(y, "/").concat(m, "/").concat(d, " ").concat(hh, ":").concat(mm);
}
function mergeLegacyNotes(_ref13) {
  var _ref13$notes = _ref13.notes,
    notes = _ref13$notes === void 0 ? "" : _ref13$notes,
    _ref13$occurrenceNote = _ref13.occurrenceNote,
    occurrenceNote = _ref13$occurrenceNote === void 0 ? "" : _ref13$occurrenceNote,
    _ref13$sketchNote = _ref13.sketchNote,
    sketchNote = _ref13$sketchNote === void 0 ? "" : _ref13$sketchNote;
  var base = value(notes);
  var occ = value(occurrenceNote);
  var sketch = value(sketchNote);
  var merged = [];
  if (base) {
    merged.push(base);
  }
  if (occ && !base.includes(occ)) {
    merged.push("\u7523\u51FA\u72B6\u6CC1\u30E1\u30E2: ".concat(occ));
  }
  if (sketch && !base.includes(sketch)) {
    merged.push("\u30B9\u30B1\u30C3\u30C1\u30FB\u89B3\u5BDF\u4E8B\u9805\u30E1\u30E2: ".concat(sketch));
  }
  return merged.join("\n\n");
}
function normalizePhotos(photosRaw) {
  if (!Array.isArray(photosRaw)) {
    return [];
  }
  return photosRaw.filter(function (photo) {
    return photo && _typeof(photo) === "object";
  }).map(function (photo) {
    return {
      id: value(photo.id) || newId("photo"),
      name: value(photo.name),
      dataUrl: value(photo.dataUrl),
      caption: value(photo.caption),
      createdAt: value(photo.createdAt) || new Date().toISOString()
    };
  }).filter(function (photo) {
    return photo.dataUrl;
  });
}
function findRecord(recordId) {
  return state.records.find(function (item) {
    return item.id === recordId;
  });
}
function findDuplicateRecordByKuwakuAndSpecimen(kuwakuRaw, specimenNoRaw) {
  var excludeRecordIdRaw = arguments.length > 2 && arguments[2] !== undefined ? arguments[2] : "";
  var kuwaku = normalizeKuwakuText(kuwakuRaw);
  var specimenNo = parseSpecimenNo(specimenNoRaw).specimenNo;
  var excludeRecordId = value(excludeRecordIdRaw);
  if (!kuwaku || !specimenNo) {
    return null;
  }
  return state.records.find(function (item) {
    if (!item || value(item.id) === excludeRecordId) {
      return false;
    }
    var itemKuwaku = normalizeKuwakuText(getRecordKuwaku(item));
    if (itemKuwaku !== kuwaku) {
      return false;
    }
    var itemSpecimenNo = parseSpecimenNo(item.specimenNo, item.specimenPrefix, item.specimenSerial).specimenNo;
    return itemSpecimenNo === specimenNo;
  }) || null;
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
function bootstrapCloudSync() {
  return _bootstrapCloudSync.apply(this, arguments);
}
function _bootstrapCloudSync() {
  _bootstrapCloudSync = _asyncToGenerator(_regenerator().m(function _callee10() {
    var response, remoteState, remoteHasData, localHasData, activeTabId, canApplyRemote, _t7;
    return _regenerator().w(function (_context10) {
      while (1) switch (_context10.p = _context10.n) {
        case 0:
          if (cloudEndpoint) {
            _context10.n = 1;
            break;
          }
          updateCloudStatus();
          return _context10.a(2);
        case 1:
          _context10.p = 1;
          _context10.n = 2;
          return requestCloud("load");
        case 2:
          response = _context10.v;
          remoteState = normalizeState(response.state);
          remoteHasData = hasAnyStateData(remoteState);
          localHasData = hasAnyStateData(state);
          activeTabId = getActiveTabId();
          canApplyRemote = activeTabId === "output-tab" || activeTabId === "plan-tab" || activeTabId === "viewer-tab" || activeTabId === "export-tab";
          if (!remoteHasData) {
            _context10.n = 3;
            break;
          }
          cloudLastPulledAt = value(response.updatedAt) || nowIso();
          if (canApplyRemote) {
            applyStateSnapshot(remoteState);
          }
          _context10.n = 4;
          break;
        case 3:
          if (!localHasData) {
            _context10.n = 4;
            break;
          }
          _context10.n = 4;
          return pushStateToCloud({
            showToastOnSuccess: false,
            silentOnError: true
          });
        case 4:
          _context10.n = 6;
          break;
        case 5:
          _context10.p = 5;
          _t7 = _context10.v;
          updateCloudStatus("同期エラー");
        case 6:
          startCloudPullTimer();
          updateCloudStatus();
        case 7:
          return _context10.a(2);
      }
    }, _callee10, null, [[1, 5]]);
  }));
  return _bootstrapCloudSync.apply(this, arguments);
}
function handleCloudConnect() {
  return _handleCloudConnect.apply(this, arguments);
}
function _handleCloudConnect() {
  _handleCloudConnect = _asyncToGenerator(_regenerator().m(function _callee11() {
    var nextEndpoint, response, remoteState, remoteHasData, localHasData, remoteUpdatedAt, localUpdatedAt, remoteMs, localMs, overwriteCloud, applied, _applied, _t8;
    return _regenerator().w(function (_context11) {
      while (1) switch (_context11.p = _context11.n) {
        case 0:
          nextEndpoint = normalizeCloudEndpoint(value(cloudEndpointInput === null || cloudEndpointInput === void 0 ? void 0 : cloudEndpointInput.value));
          if (nextEndpoint) {
            _context11.n = 1;
            break;
          }
          showToast("Google Apps Script のWebアプリURLを入力してください");
          return _context11.a(2);
        case 1:
          cloudEndpoint = nextEndpoint;
          saveCloudEndpoint(nextEndpoint);
          if (cloudEndpointInput) {
            cloudEndpointInput.value = nextEndpoint;
          }
          updateCloudStatus("接続確認中");
          _context11.p = 2;
          _context11.n = 3;
          return requestCloud("load");
        case 3:
          response = _context11.v;
          remoteState = normalizeState(response.state);
          remoteHasData = hasAnyStateData(remoteState);
          localHasData = hasAnyStateData(state);
          remoteUpdatedAt = value(response.updatedAt) || getStateUpdatedAt(remoteState);
          localUpdatedAt = getStateUpdatedAt(state);
          remoteMs = Number.parseInt(String(Date.parse(remoteUpdatedAt || "")), 10);
          localMs = Number.parseInt(String(Date.parse(localUpdatedAt || "")), 10);
          if (!(remoteHasData && localHasData && Number.isFinite(remoteMs) && Number.isFinite(localMs) && localMs > remoteMs)) {
            _context11.n = 7;
            break;
          }
          overwriteCloud = window.confirm("端末側のデータのほうが新しい可能性があります。\nOK: 端末データでクラウドを上書き\nキャンセル: クラウドデータを読み込み");
          if (!overwriteCloud) {
            _context11.n = 5;
            break;
          }
          _context11.n = 4;
          return pushStateToCloud({
            showToastOnSuccess: false
          });
        case 4:
          showToast("共有保存を有効化し、端末データをクラウドへ保存しました");
          _context11.n = 6;
          break;
        case 5:
          applied = applyStateSnapshot(remoteState);
          cloudLastPulledAt = remoteUpdatedAt || nowIso();
          showToast(applied ? "共有保存を有効化し、クラウドデータを読み込みました" : "共有保存を有効化しました");
        case 6:
          _context11.n = 10;
          break;
        case 7:
          if (!remoteHasData) {
            _context11.n = 8;
            break;
          }
          _applied = applyStateSnapshot(remoteState);
          cloudLastPulledAt = remoteUpdatedAt || nowIso();
          showToast(_applied ? "共有保存を有効化し、クラウドデータを読み込みました" : "共有保存を有効化しました");
          _context11.n = 10;
          break;
        case 8:
          _context11.n = 9;
          return pushStateToCloud({
            showToastOnSuccess: false
          });
        case 9:
          showToast("共有保存を有効化しました");
        case 10:
          startCloudPullTimer();
          updateCloudStatus();
          _context11.n = 12;
          break;
        case 11:
          _context11.p = 11;
          _t8 = _context11.v;
          disableCloudSync({
            showToastMessage: false
          });
          showToast("共有保存の接続に失敗しました。URLと公開設定を確認してください");
        case 12:
          return _context11.a(2);
      }
    }, _callee11, null, [[2, 11]]);
  }));
  return _handleCloudConnect.apply(this, arguments);
}
function handleCloudManualReload() {
  return _handleCloudManualReload.apply(this, arguments);
}
function _handleCloudManualReload() {
  _handleCloudManualReload = _asyncToGenerator(_regenerator().m(function _callee12() {
    var activeTabId, answer;
    return _regenerator().w(function (_context12) {
      while (1) switch (_context12.n) {
        case 0:
          if (cloudEndpoint) {
            _context12.n = 1;
            break;
          }
          showToast("共有保存は未設定です");
          return _context12.a(2);
        case 1:
          activeTabId = getActiveTabId();
          if (!(activeTabId === "input-tab" || activeTabId === "edit-tab")) {
            _context12.n = 2;
            break;
          }
          answer = window.confirm("入力途中の内容は失われる場合があります。クラウドを再読込しますか？");
          if (answer) {
            _context12.n = 2;
            break;
          }
          return _context12.a(2);
        case 2:
          _context12.n = 3;
          return pullStateFromCloud({
            force: true,
            showToastOnSuccess: true
          });
        case 3:
          return _context12.a(2);
      }
    }, _callee12);
  }));
  return _handleCloudManualReload.apply(this, arguments);
}
function disableCloudSync() {
  var _ref14 = arguments.length > 0 && arguments[0] !== undefined ? arguments[0] : {
      showToastMessage: false
    },
    showToastMessage = _ref14.showToastMessage;
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
  cloudPullTimer = window.setInterval(function () {
    var activeTabId = getActiveTabId();
    if (activeTabId === "output-tab" || activeTabId === "plan-tab" || activeTabId === "viewer-tab" || activeTabId === "export-tab") {
      void pullStateFromCloud({
        force: false,
        showToastOnSuccess: false,
        silentOnError: true
      });
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
  cloudSaveTimer = window.setTimeout(function () {
    cloudSaveTimer = null;
    void pushStateToCloud({
      showToastOnSuccess: false,
      silentOnError: true
    });
  }, CLOUD_SAVE_DEBOUNCE_MS);
}
function pullStateFromCloud() {
  return _pullStateFromCloud.apply(this, arguments);
}
function _pullStateFromCloud() {
  _pullStateFromCloud = _asyncToGenerator(_regenerator().m(function _callee13() {
    var _ref17,
      _ref17$force,
      force,
      _ref17$showToastOnSuc,
      showToastOnSuccess,
      _ref17$silentOnError,
      silentOnError,
      tabAtRequest,
      response,
      remoteState,
      remoteHasData,
      localHasData,
      remoteUpdatedAt,
      localUpdatedAt,
      remoteMs,
      localMs,
      tabBeforeApply,
      applied,
      _args13 = arguments,
      _t9;
    return _regenerator().w(function (_context13) {
      while (1) switch (_context13.p = _context13.n) {
        case 0:
          _ref17 = _args13.length > 0 && _args13[0] !== undefined ? _args13[0] : {}, _ref17$force = _ref17.force, force = _ref17$force === void 0 ? false : _ref17$force, _ref17$showToastOnSuc = _ref17.showToastOnSuccess, showToastOnSuccess = _ref17$showToastOnSuc === void 0 ? false : _ref17$showToastOnSuc, _ref17$silentOnError = _ref17.silentOnError, silentOnError = _ref17$silentOnError === void 0 ? false : _ref17$silentOnError;
          if (!(!cloudEndpoint || cloudPullInProgress)) {
            _context13.n = 1;
            break;
          }
          return _context13.a(2, false);
        case 1:
          tabAtRequest = getActiveTabId();
          if (!(!force && tabAtRequest !== "output-tab" && tabAtRequest !== "plan-tab" && tabAtRequest !== "viewer-tab" && tabAtRequest !== "export-tab")) {
            _context13.n = 2;
            break;
          }
          return _context13.a(2, false);
        case 2:
          cloudPullInProgress = true;
          _context13.p = 3;
          _context13.n = 4;
          return requestCloud("load");
        case 4:
          response = _context13.v;
          remoteState = normalizeState(response.state);
          remoteHasData = hasAnyStateData(remoteState);
          localHasData = hasAnyStateData(state);
          remoteUpdatedAt = value(response.updatedAt) || getStateUpdatedAt(remoteState);
          localUpdatedAt = getStateUpdatedAt(state);
          remoteMs = Date.parse(remoteUpdatedAt || "");
          localMs = Date.parse(localUpdatedAt || "");
          if (!(!force && !remoteHasData && localHasData)) {
            _context13.n = 5;
            break;
          }
          return _context13.a(2, false);
        case 5:
          if (!(!force && Number.isFinite(remoteMs) && Number.isFinite(localMs) && remoteMs <= localMs)) {
            _context13.n = 6;
            break;
          }
          cloudLastPulledAt = remoteUpdatedAt || cloudLastPulledAt;
          updateCloudStatus();
          return _context13.a(2, false);
        case 6:
          tabBeforeApply = getActiveTabId();
          if (!(!force && (tabBeforeApply === "input-tab" || tabBeforeApply === "edit-tab"))) {
            _context13.n = 7;
            break;
          }
          return _context13.a(2, false);
        case 7:
          applied = applyStateSnapshot(remoteState, {
            force: force
          });
          if (applied) {
            _context13.n = 8;
            break;
          }
          return _context13.a(2, false);
        case 8:
          cloudLastPulledAt = remoteUpdatedAt || nowIso();
          updateCloudStatus();
          if (showToastOnSuccess) {
            showToast("クラウドから最新データを読み込みました");
          }
          return _context13.a(2, true);
        case 9:
          _context13.p = 9;
          _t9 = _context13.v;
          if (!silentOnError) {
            notifyCloudError("クラウド読込に失敗しました");
          }
          updateCloudStatus("同期エラー");
          return _context13.a(2, false);
        case 10:
          _context13.p = 10;
          cloudPullInProgress = false;
          return _context13.f(10);
        case 11:
          return _context13.a(2);
      }
    }, _callee13, null, [[3, 9, 10, 11]]);
  }));
  return _pullStateFromCloud.apply(this, arguments);
}
function pushStateToCloud() {
  return _pushStateToCloud.apply(this, arguments);
}
function _pushStateToCloud() {
  _pushStateToCloud = _asyncToGenerator(_regenerator().m(function _callee14() {
    var _ref18,
      _ref18$showToastOnSuc,
      showToastOnSuccess,
      _ref18$silentOnError,
      silentOnError,
      mergedStateForSave,
      remoteResponse,
      remoteState,
      remoteUpdatedAt,
      payload,
      response,
      _args14 = arguments,
      _t0,
      _t1;
    return _regenerator().w(function (_context14) {
      while (1) switch (_context14.p = _context14.n) {
        case 0:
          _ref18 = _args14.length > 0 && _args14[0] !== undefined ? _args14[0] : {}, _ref18$showToastOnSuc = _ref18.showToastOnSuccess, showToastOnSuccess = _ref18$showToastOnSuc === void 0 ? false : _ref18$showToastOnSuc, _ref18$silentOnError = _ref18.silentOnError, silentOnError = _ref18$silentOnError === void 0 ? false : _ref18$silentOnError;
          if (!(!cloudEndpoint || cloudPushInProgress)) {
            _context14.n = 1;
            break;
          }
          return _context14.a(2, false);
        case 1:
          cloudPushInProgress = true;
          _context14.p = 2;
          mergedStateForSave = normalizeState(state);
          _context14.p = 3;
          _context14.n = 4;
          return requestCloud("load");
        case 4:
          remoteResponse = _context14.v;
          remoteState = normalizeState(remoteResponse === null || remoteResponse === void 0 ? void 0 : remoteResponse.state);
          mergedStateForSave = mergeStatesForCloud(remoteState, mergedStateForSave);
          remoteUpdatedAt = value(remoteResponse === null || remoteResponse === void 0 ? void 0 : remoteResponse.updatedAt) || getStateUpdatedAt(remoteState);
          if (remoteUpdatedAt) {
            cloudLastPulledAt = remoteUpdatedAt;
          }
          _context14.n = 6;
          break;
        case 5:
          _context14.p = 5;
          _t0 = _context14.v;
          mergedStateForSave = normalizeState(state);
        case 6:
          payload = {
            clientId: cloudClientId,
            updatedAt: getStateUpdatedAt(mergedStateForSave) || nowIso(),
            state: mergedStateForSave
          };
          _context14.n = 7;
          return requestCloud("save", payload);
        case 7:
          response = _context14.v;
          state = normalizeState(mergedStateForSave);
          try {
            localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
          } catch (_error) {}
          cloudLastSyncedAt = value(response.updatedAt) || payload.updatedAt;
          updateCloudStatus();
          if (showToastOnSuccess) {
            showToast("クラウドへ保存しました");
          }
          return _context14.a(2, true);
        case 8:
          _context14.p = 8;
          _t1 = _context14.v;
          if (!silentOnError) {
            notifyCloudError("クラウド保存に失敗しました（端末には保存済み）");
          }
          updateCloudStatus("同期エラー");
          return _context14.a(2, false);
        case 9:
          _context14.p = 9;
          cloudPushInProgress = false;
          return _context14.f(9);
        case 10:
          return _context14.a(2);
      }
    }, _callee14, null, [[3, 5], [2, 8, 9, 10]]);
  }));
  return _pushStateToCloud.apply(this, arguments);
}
function requestCloud(_x0) {
  return _requestCloud.apply(this, arguments);
}
function _requestCloud() {
  _requestCloud = _asyncToGenerator(_regenerator().m(function _callee15(action) {
    var payload,
      response,
      separator,
      url,
      form,
      text,
      body,
      _body,
      _args15 = arguments,
      _t10;
    return _regenerator().w(function (_context15) {
      while (1) switch (_context15.p = _context15.n) {
        case 0:
          payload = _args15.length > 1 && _args15[1] !== undefined ? _args15[1] : null;
          if (cloudEndpoint) {
            _context15.n = 1;
            break;
          }
          throw new Error("Cloud endpoint is not configured");
        case 1:
          if (!(action === "load")) {
            _context15.n = 3;
            break;
          }
          separator = cloudEndpoint.includes("?") ? "&" : "?";
          url = "".concat(cloudEndpoint).concat(separator, "action=load&t=").concat(Date.now());
          _context15.n = 2;
          return fetch(url, {
            method: "GET",
            cache: "no-store"
          });
        case 2:
          response = _context15.v;
          _context15.n = 5;
          break;
        case 3:
          form = new URLSearchParams();
          form.set("action", action);
          form.set("payload", JSON.stringify(payload || {}));
          _context15.n = 4;
          return fetch(cloudEndpoint, {
            method: "POST",
            body: form
          });
        case 4:
          response = _context15.v;
        case 5:
          _context15.n = 6;
          return response.text();
        case 6:
          text = _context15.v;
          body = null;
          _context15.p = 7;
          body = JSON.parse(text);
          _context15.n = 9;
          break;
        case 8:
          _context15.p = 8;
          _t10 = _context15.v;
          throw new Error("Invalid cloud response");
        case 9:
          if (!(!response.ok || !body || body.ok === false)) {
            _context15.n = 10;
            break;
          }
          throw new Error(value((_body = body) === null || _body === void 0 ? void 0 : _body.error) || "HTTP ".concat(response.status));
        case 10:
          return _context15.a(2, body);
      }
    }, _callee15, null, [[7, 8]]);
  }));
  return _requestCloud.apply(this, arguments);
}
function applyStateSnapshot(nextStateRaw) {
  var _ref15 = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : {},
    _ref15$force = _ref15.force,
    force = _ref15$force === void 0 ? false : _ref15$force;
  var activeTabId = getActiveTabId();
  if (!force && (activeTabId === "input-tab" || activeTabId === "edit-tab")) {
    return false;
  }
  state = normalizeState(nextStateRaw);
  hydrateSiteForm();
  resetRecordForm({
    showMessage: false
  });
  renderRecordTable();
  renderOutputs();
  return true;
}
function updateCloudStatus() {
  var statusNote = arguments.length > 0 && arguments[0] !== undefined ? arguments[0] : "";
  if (!cloudStatusEl) {
    return;
  }
  if (!cloudEndpoint) {
    cloudStatusEl.textContent = "保存先: 端末内のみ";
    cloudStatusEl.classList.remove("online");
    return;
  }
  var latest = cloudLastSyncedAt || cloudLastPulledAt;
  var latestText = latest ? " / \u6700\u7D42\u540C\u671F ".concat(formatStatusTime(latest)) : "";
  var note = statusNote ? " / ".concat(statusNote) : "";
  cloudStatusEl.textContent = "\u4FDD\u5B58\u5148: \u5171\u6709Google\u30C9\u30E9\u30A4\u30D6".concat(latestText).concat(note);
  cloudStatusEl.classList.add("online");
}
function notifyCloudError(message) {
  var now = Date.now();
  if (now - cloudLastErrorAt < 5000) {
    return;
  }
  cloudLastErrorAt = now;
  showToast(message);
}
function normalizeCloudEndpoint(urlRaw) {
  var url = value(urlRaw);
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
    var saved = normalizeCloudEndpoint(value(localStorage.getItem(CLOUD_ENDPOINT_KEY)));
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
    var endpoint = value(url);
    if (endpoint) {
      localStorage.setItem(CLOUD_ENDPOINT_KEY, endpoint);
    } else {
      localStorage.removeItem(CLOUD_ENDPOINT_KEY);
    }
  } catch (_error) {}
}
function loadOrCreateCloudClientId() {
  try {
    var existing = value(localStorage.getItem(CLOUD_CLIENT_ID_KEY));
    if (existing) {
      return existing;
    }
    var clientId = newId("client");
    localStorage.setItem(CLOUD_CLIENT_ID_KEY, clientId);
    return clientId;
  } catch (_error) {
    return newId("client");
  }
}
function hasAnyStateData(candidateState) {
  var normalized = normalizeState(candidateState);
  if (normalized.records.length) {
    return true;
  }
  var site = normalized.site || {};
  return Boolean(value(site.kuwakuHeadA) !== DEFAULT_KUWAKU_HEAD_A || value(site.kuwakuHeadB) !== DEFAULT_KUWAKU_HEAD_B || value(site.kuwakuBlock) || value(site.kuwakuNo) || value(site.levelHeight) || value(site.date) || value(site.team) || value(site.teamLead) || value(site.recorder));
}
function getStateUpdatedAt(candidateState) {
  var _candidateState$site;
  if (!candidateState || _typeof(candidateState) !== "object") {
    return "";
  }
  var latestMs = Number.NEGATIVE_INFINITY;
  var pushDate = function pushDate(raw) {
    var ms = Date.parse(value(raw));
    if (Number.isFinite(ms)) {
      latestMs = Math.max(latestMs, ms);
    }
  };
  pushDate((_candidateState$site = candidateState.site) === null || _candidateState$site === void 0 ? void 0 : _candidateState$site.updatedAt);
  if (Array.isArray(candidateState.records)) {
    candidateState.records.forEach(function (record) {
      pushDate(record === null || record === void 0 ? void 0 : record.updatedAt);
      pushDate(record === null || record === void 0 ? void 0 : record.createdAt);
    });
  }
  return Number.isFinite(latestMs) ? new Date(latestMs).toISOString() : "";
}
function mergeStatesForCloud(remoteStateRaw, localStateRaw) {
  var remoteState = normalizeState(remoteStateRaw);
  var localState = normalizeState(localStateRaw);
  var mergedSite = mergeSiteForCloud(remoteState.site, localState.site);
  var recordById = new Map();
  var upsertRecord = function upsertRecord(recordRaw, preferIncomingOnTie) {
    var record = normalizeRecord(recordRaw, mergedSite);
    var recordId = value(record === null || record === void 0 ? void 0 : record.id);
    if (!recordId) {
      return;
    }
    var existing = recordById.get(recordId);
    if (!existing) {
      recordById.set(recordId, record);
      return;
    }
    recordById.set(recordId, chooseNewerRecordForCloud(existing, record, preferIncomingOnTie));
  };
  remoteState.records.forEach(function (record) {
    return upsertRecord(record, false);
  });
  localState.records.forEach(function (record) {
    return upsertRecord(record, true);
  });
  return normalizeState({
    site: mergedSite,
    records: Array.from(recordById.values())
  });
}
function mergeSiteForCloud(remoteSiteRaw, localSiteRaw) {
  var remoteSite = remoteSiteRaw && _typeof(remoteSiteRaw) === "object" ? remoteSiteRaw : {};
  var localSite = localSiteRaw && _typeof(localSiteRaw) === "object" ? localSiteRaw : {};
  var preferLocal = isIsoTimestampGreaterOrEqual(localSite.updatedAt, remoteSite.updatedAt);
  var primary = preferLocal ? localSite : remoteSite;
  var secondary = preferLocal ? remoteSite : localSite;
  var mergedUpdatedAt = pickLatestIsoTimestamp(localSite.updatedAt, remoteSite.updatedAt);
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
    updatedAt: mergedUpdatedAt || value(primary.updatedAt) || value(secondary.updatedAt)
  };
}
function chooseNewerRecordForCloud(existingRecordRaw, incomingRecordRaw) {
  var _winner;
  var preferIncomingOnTie = arguments.length > 2 && arguments[2] !== undefined ? arguments[2] : true;
  var existingRecord = normalizeRecord(existingRecordRaw);
  var incomingRecord = normalizeRecord(incomingRecordRaw);
  var existingMs = getRecordUpdatedAtMs(existingRecord);
  var incomingMs = getRecordUpdatedAtMs(incomingRecord);
  var winner = existingRecord;
  var loser = incomingRecord;
  if (Number.isFinite(incomingMs) && Number.isFinite(existingMs)) {
    if (incomingMs > existingMs || incomingMs === existingMs && preferIncomingOnTie) {
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
  var mergedHistory = mergeRecordHistoryEntries(existingRecord === null || existingRecord === void 0 ? void 0 : existingRecord.history, incomingRecord === null || incomingRecord === void 0 ? void 0 : incomingRecord.history);
  return _objectSpread(_objectSpread(_objectSpread({}, loser), winner), {}, {
    history: mergedHistory.length ? mergedHistory : normalizeRecordHistory((_winner = winner) === null || _winner === void 0 ? void 0 : _winner.history)
  });
}
function mergeRecordHistoryEntries(historyA, historyB) {
  var mergedMap = new Map();
  var upsert = function upsert(entryRaw) {
    var entry = normalizeRecordHistory([entryRaw])[0];
    if (!entry) {
      return;
    }
    var key = value(entry.id) || "".concat(value(entry.at), "::").concat(value(entry.action), "::").concat(value(entry.content));
    if (!key) {
      return;
    }
    var existing = mergedMap.get(key);
    if (!existing) {
      mergedMap.set(key, entry);
      return;
    }
    var existingMs = Date.parse(value(existing.at));
    var incomingMs = Date.parse(value(entry.at));
    if (Number.isFinite(incomingMs) && (!Number.isFinite(existingMs) || incomingMs >= existingMs)) {
      mergedMap.set(key, entry);
    }
  };
  normalizeRecordHistory(historyA).forEach(upsert);
  normalizeRecordHistory(historyB).forEach(upsert);
  return Array.from(mergedMap.values()).sort(function (a, b) {
    var aMs = Date.parse(value(a.at));
    var bMs = Date.parse(value(b.at));
    if (Number.isFinite(aMs) && Number.isFinite(bMs) && aMs !== bMs) {
      return aMs - bMs;
    }
    return value(a.id).localeCompare(value(b.id));
  }).slice(-50);
}
function getRecordUpdatedAtMs(record) {
  var updatedMs = Date.parse(value(record === null || record === void 0 ? void 0 : record.updatedAt));
  if (Number.isFinite(updatedMs)) {
    return updatedMs;
  }
  var createdMs = Date.parse(value(record === null || record === void 0 ? void 0 : record.createdAt));
  return Number.isFinite(createdMs) ? createdMs : Number.NaN;
}
function pickLatestIsoTimestamp(tsA, tsB) {
  var aMs = Date.parse(value(tsA));
  var bMs = Date.parse(value(tsB));
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
  var aMs = Date.parse(value(tsA));
  var bMs = Date.parse(value(tsB));
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
  var date = new Date(isoString);
  if (!Number.isFinite(date.getTime())) {
    return "-";
  }
  var hh = String(date.getHours()).padStart(2, "0");
  var mm = String(date.getMinutes()).padStart(2, "0");
  var ss = String(date.getSeconds()).padStart(2, "0");
  return "".concat(hh, ":").concat(mm, ":").concat(ss);
}
function buildListCsv() {
  var header = ["区画", "レベル高(m)", "日付", "発掘班", "班長", "記載係", "標本番号", "分類", "名称", "ユニット", "発見者", "判定者", "レベル読値_上面(cm)", "レベル読値_下底(cm)", "標高(m)", "産出状況断面", "産状スケッチ", "平面位置_NS", "平面位置_NS_cm", "平面位置_EW", "平面位置_EW_cm", "備考（観察事項など）"];
  var rows = getListExportRecords().map(function (record) {
    return [getRecordKuwaku(record), getRecordLevelHeight(record), getRecordDate(record), getRecordTeamValue(record), getRecordTeamLead(record), getRecordRecorder(record), record.specimenNo, formatCategoryForRecord(record), record.nameMemo, record.unit, record.discoverer, record.identifier, formatCmValue(record.levelUpperCm), formatCmValue(record.levelLowerCm), formatRecordAltitudeM(record), record.occurrenceSection, record.occurrenceSketch, record.nsDir, formatCmValue(record.nsCm), record.ewDir, formatCmValue(record.ewCm), record.notes];
  });
  return [header].concat(_toConsumableArray(rows)).map(function (row) {
    return row.map(csvCell).join(",");
  }).join("\n");
}
function buildCardCsv() {
  var records = getCardExportRecords();
  var header = ["区画", "標本番号", "分類", "化石・遺物名称", "重要品指定", "簡易記載", "地層名", "ユニット", "サブユニット", "細分", "層相", "地層中の位置", "レベル読値", "平面位置", "発見者", "判定者", "産出状況断面", "産状スケッチ", "備考（観察事項など）", "産出状況断面図枚数", "写真枚数"];
  var rows = records.map(function (record) {
    return [getRecordKuwaku(record), record.specimenNo, formatCategoryForRecord(record), record.nameMemo, record.importantFlag, record.simpleRecordFlag, record.layerName, record.unit, record.detail, record.detailSub, record.layerFacies, formatLayerPosition(record), formatLevelRead(record), formatPlanPosition(record), record.discoverer, record.identifier, record.occurrenceSection, record.occurrenceSketch, record.notes, String((record.sectionDiagrams || []).length), String((record.photos || []).length)];
  });
  return [header].concat(_toConsumableArray(rows)).map(function (row) {
    return row.map(csvCell).join(",");
  }).join("\n");
}
function exportListPdf() {
  var records = getListExportRecords();
  if (!records.length) {
    showToast("PDF出力対象データがありません");
    return;
  }
  var selectedGridLabel = exportListRangeKuwaku === ALL_GRIDS_VALUE ? "全グリッド" : kuwakuLabelForSelect(exportListRangeKuwaku);
  var categoryLabel = exportListRangeCategory === EXPORT_CATEGORY_ALL_VALUE ? "全分類" : "".concat(exportListRangeCategory, ": ").concat(SPECIMEN_CATEGORY_MAP[exportListRangeCategory] || "");
  var statusLabel = exportListRangeStatus === "complete" ? "必須完了のみ" : exportListRangeStatus === "incomplete" ? "未記入ありのみ" : "すべて";
  var specimenRangeLabel = exportListRangeSpecimenFrom || exportListRangeSpecimenTo ? "".concat(value(exportListRangeSpecimenFrom) || "-", " \u301C ").concat(value(exportListRangeSpecimenTo) || "-") : "指定なし";
  var html = buildListPdfHtml(records, {
    selectedGridLabel: selectedGridLabel,
    categoryLabel: categoryLabel,
    statusLabel: statusLabel,
    specimenRangeLabel: specimenRangeLabel
  });
  var opened = openPdfPrintWindow({
    title: "化石遺物リストout＿出力.pdf",
    pageSize: "A3 landscape",
    bodyHtml: html
  });
  if (opened) {
    showToast("リストPDFの印刷画面を開きました（保存先でPDFを選択）");
  }
}
function exportCardPdf() {
  var records = getCardExportRecords();
  if (!records.length) {
    showToast("カードPDFの出力対象データがありません");
    return;
  }
  var html = buildCardPdfHtml(records);
  var opened = openPdfPrintWindow({
    title: "化石遺物カードout＿出力.pdf",
    pageSize: "A4 portrait",
    bodyHtml: html
  });
  if (opened) {
    showToast("カードPDFの印刷画面を開きました（保存先でPDFを選択）");
  }
}
function exportPlanPdf() {
  var groups = buildPlanPdfGroupsForExport({
    kuwakuValue: exportPlanKuwaku,
    categoryValue: exportPlanCategory,
    dateFromRaw: exportPlanDateFrom,
    dateToRaw: exportPlanDateTo,
    modeSelections: getExportPlanModeSelections()
  });
  if (!groups.length) {
    showToast("平面図PDFの出力対象がありません（平面位置を確認してください）");
    return;
  }
  var html = buildPlanPdfHtml(groups);
  var opened = openPdfPrintWindow({
    title: "層準別平面図out＿出力.pdf",
    pageSize: "A4 portrait",
    bodyHtml: html
  });
  if (opened) {
    showToast("平面図PDFの印刷画面を開きました（保存先でPDFを選択）");
  }
}
function buildListPdfHtml(records) {
  var options = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : {};
  var generatedAt = formatPdfGeneratedAt(new Date());
  var selectedGrid = value(options.selectedGridLabel) || "全グリッド";
  var categoryLabel = value(options.categoryLabel);
  var statusLabel = value(options.statusLabel);
  var specimenRangeLabel = value(options.specimenRangeLabel);
  var rows = records.map(function (record) {
    var complete = isRecordDataComplete(record);
    return "\n        <tr>\n          <td>".concat(escapeHtml(getRecordKuwaku(record)), "</td>\n          <td>").concat(escapeHtml(getRecordTeamValue(record)), "</td>\n          <td>").concat(escapeHtml(record.specimenNo || ""), "</td>\n          <td>").concat(escapeHtml(formatCategoryForRecord(record)), "</td>\n          <td>").concat(escapeHtml(record.nameMemo || ""), "</td>\n          <td>").concat(escapeHtml(record.importantFlag || ""), "</td>\n          <td>").concat(escapeHtml(record.unit || ""), "</td>\n          <td>").concat(escapeHtml(formatDetailForRecord(record)), "</td>\n          <td>").concat(escapeHtml(record.discoverer || ""), "</td>\n          <td>").concat(escapeHtml(record.identifier || ""), "</td>\n          <td>").concat(escapeHtml(formatLevelRead(record)), "</td>\n          <td>").concat(escapeHtml(formatRecordAltitudeM(record)), "</td>\n          <td>").concat(escapeHtml(formatPlanPosition(record)), "</td>\n          <td class=\"").concat(complete ? "pdf-status-complete" : "pdf-status-incomplete", "\">").concat(complete ? "○" : "未記入", "</td>\n        </tr>\n      ");
  }).join("");
  return "\n    <section class=\"pdf-page\">\n      <header class=\"pdf-header\">\n        <h1>\u5316\u77F3\u907A\u7269\u30EA\u30B9\u30C8</h1>\n        <div class=\"pdf-meta\">\n          <span>\u533A\u753B: ".concat(escapeHtml(selectedGrid), "</span>\n          <span>\u5206\u985E: ").concat(escapeHtml(categoryLabel || "全分類"), "</span>\n          <span>\u30C7\u30FC\u30BF\u72B6\u614B: ").concat(escapeHtml(statusLabel || "すべて"), "</span>\n          <span>\u6A19\u672C\u756A\u53F7: ").concat(escapeHtml(specimenRangeLabel || "指定なし"), "</span>\n          <span>\u51FA\u529B\u65E5\u6642: ").concat(escapeHtml(generatedAt), "</span>\n          <span>\u4EF6\u6570: ").concat(records.length, "</span>\n        </div>\n      </header>\n      <table class=\"pdf-table pdf-list-table\">\n        <thead>\n          <tr>\n            <th>\u533A\u753B</th>\n            <th>\u767A\u6398\u73ED</th>\n            <th>\u6A19\u672C\u756A\u53F7</th>\n            <th>\u5206\u985E</th>\n            <th>\u540D\u79F0</th>\n            <th>\u91CD\u8981\u54C1</th>\n            <th>\u30E6\u30CB\u30C3\u30C8</th>\n            <th>\u30B5\u30D6\u30E6\u30CB\u30C3\u30C8</th>\n            <th>\u767A\u898B\u8005</th>\n            <th>\u5224\u5B9A\u8005</th>\n            <th>\u30EC\u30D9\u30EB\u8AAD\u5024</th>\n            <th>\u6A19\u9AD8(m)</th>\n            <th>\u5E73\u9762\u4F4D\u7F6E</th>\n            <th>\u30C7\u30FC\u30BF</th>\n          </tr>\n        </thead>\n        <tbody>").concat(rows, "</tbody>\n      </table>\n    </section>\n  ");
}
function buildCardPdfHtml(records) {
  var generatedAt = formatPdfGeneratedAt(new Date());
  return records.map(function (record, index) {
    var sectionDiagramsHtml = buildPdfImageGrid(record.sectionDiagrams, "断面図なし", "pdf-card-image-grid");
    var photosHtml = buildPdfImageGrid(record.photos, "写真なし", "pdf-card-image-grid");
    return "\n        <section class=\"pdf-page pdf-card-page ".concat(index < records.length - 1 ? "pdf-page-break" : "", "\">\n          <header class=\"pdf-header\">\n            <h1>\u5316\u77F3\u907A\u7269\u30AB\u30FC\u30C9</h1>\n            <div class=\"pdf-meta\">\n              <span>\u533A\u753B: ").concat(escapeHtml(getRecordKuwaku(record)), "</span>\n              <span>\u6A19\u672C\u756A\u53F7: ").concat(escapeHtml(record.specimenNo || ""), "</span>\n              <span>\u51FA\u529B\u65E5\u6642: ").concat(escapeHtml(generatedAt), "</span>\n            </div>\n          </header>\n          <table class=\"pdf-table pdf-card-table\">\n            <tbody>\n              <tr><th>\u5206\u985E</th><td>").concat(escapeHtml(formatCategoryForRecord(record)), "</td><th>\u91CD\u8981\u54C1\u6307\u5B9A</th><td>").concat(escapeHtml(record.importantFlag || ""), "</td></tr>\n              <tr><th>\u5316\u77F3\u30FB\u907A\u7269\u540D\u79F0</th><td>").concat(escapeHtml(record.nameMemo || ""), "</td><th>\u7C21\u6613\u8A18\u8F09</th><td>").concat(escapeHtml(record.simpleRecordFlag || "-"), "</td></tr>\n              <tr><th>\u5730\u5C64\u540D</th><td>").concat(escapeHtml(record.layerName || ""), "</td><th>\u30E6\u30CB\u30C3\u30C8</th><td>").concat(escapeHtml(record.unit || ""), "</td></tr>\n              <tr><th>\u30B5\u30D6\u30E6\u30CB\u30C3\u30C8</th><td>").concat(escapeHtml(record.detail || ""), "</td><th>\u7D30\u5206</th><td>").concat(escapeHtml(record.detailSub || ""), "</td></tr>\n              <tr><th>\u5C64\u76F8</th><td colspan=\"3\">").concat(escapeHtml(record.layerFacies || ""), "</td></tr>\n              <tr><th>\u5730\u5C64\u4E2D\u306E\u4F4D\u7F6E</th><td colspan=\"3\">").concat(escapeHtml(formatLayerPosition(record)), "</td></tr>\n              <tr><th>\u767A\u898B\u8005\u6C0F\u540D</th><td>").concat(escapeHtml(record.discoverer || ""), "</td><th>\u5224\u5B9A\u8005\u6C0F\u540D</th><td>").concat(escapeHtml(record.identifier || ""), "</td></tr>\n              <tr><th>\u30EC\u30D9\u30EB\u8AAD\u5024</th><td>").concat(escapeHtml(formatLevelRead(record)), "</td><th>\u5E73\u9762\u4F4D\u7F6E</th><td>").concat(escapeHtml(formatPlanPosition(record)), "</td></tr>\n              <tr><th>\u7523\u51FA\u72B6\u6CC1\u65AD\u9762</th><td>").concat(escapeHtml(record.occurrenceSection || ""), "</td><th>\u7523\u72B6\u30B9\u30B1\u30C3\u30C1</th><td>").concat(escapeHtml(record.occurrenceSketch || ""), "</td></tr>\n              <tr><th>\u5099\u8003\uFF08\u89B3\u5BDF\u4E8B\u9805\u306A\u3069\uFF09</th><td colspan=\"3\">").concat(escapeHtml(record.notes || ""), "</td></tr>\n            </tbody>\n          </table>\n          <section class=\"pdf-image-section\">\n            <h2>\u7523\u51FA\u72B6\u6CC1\u65AD\u9762\u56F3</h2>\n            ").concat(sectionDiagramsHtml, "\n          </section>\n          <section class=\"pdf-image-section\">\n            <h2>\u5199\u771F</h2>\n            ").concat(photosHtml, "\n          </section>\n        </section>\n      ");
  }).join("");
}
function buildPlanPdfGroups() {
  var kuwakuRecords = getFilteredPlanRecords();
  if (!kuwakuRecords.length) {
    return [];
  }
  var unitValues = selectedPlanUnit === ALL_UNITS_VALUE ? collectPlanUnits(kuwakuRecords).map(function (item) {
    return item.value;
  }).filter(function (valueRaw) {
    return valueRaw !== ALL_UNITS_VALUE;
  }) : [selectedPlanUnit];
  var groups = [];
  unitValues.forEach(function (unitValue) {
    var unitRecords = unitValue === ALL_UNITS_VALUE ? kuwakuRecords : kuwakuRecords.filter(function (record) {
      return unitValueForSelect(record.unit) === unitValue;
    });
    if (!unitRecords.length) {
      return;
    }
    var detailValues = selectedPlanDetail === ALL_DETAILS_VALUE ? collectPlanDetails(unitRecords).map(function (item) {
      return item.value;
    }).filter(function (valueRaw) {
      return valueRaw !== ALL_DETAILS_VALUE;
    }) : [selectedPlanDetail];
    detailValues.forEach(function (detailValue) {
      var detailRecords = detailValue === ALL_DETAILS_VALUE ? unitRecords : unitRecords.filter(function (record) {
        return detailValueForSelect(record.detail) === detailValue;
      });
      if (!detailRecords.length) {
        return;
      }
      var detailSubValues = selectedPlanDetailSub === ALL_DETAIL_SUBS_VALUE ? collectPlanDetailSubs(detailRecords).map(function (item) {
        return item.value;
      }).filter(function (valueRaw) {
        return valueRaw !== ALL_DETAIL_SUBS_VALUE;
      }) : [selectedPlanDetailSub];
      detailSubValues.forEach(function (detailSubValue) {
        var scopedRecords = detailSubValue === ALL_DETAIL_SUBS_VALUE ? detailRecords : detailRecords.filter(function (record) {
          return detailSubValueForSelect(record.detailSub) === detailSubValue;
        });
        var drawables = scopedRecords.map(function (record) {
          return buildPlanDrawable(record);
        }).filter(Boolean);
        if (!drawables.length) {
          return;
        }
        groups.push({
          unitValue: unitValue,
          detailValue: detailValue,
          detailSubValue: detailSubValue,
          drawables: drawables,
          count: scopedRecords.length
        });
      });
    });
  });
  return groups;
}
function buildPlanPdfGroupsForExport() {
  var filters = arguments.length > 0 && arguments[0] !== undefined ? arguments[0] : {};
  var allRecords = _toConsumableArray(state.records).sort(compareRecordsByKuwakuThenSpecimen);
  if (!allRecords.length) {
    return [];
  }
  var kuwakuValue = value(filters.kuwakuValue);
  if (!kuwakuValue || kuwakuValue === ALL_GRIDS_VALUE) {
    return [];
  }
  var categoryValue = value(filters.categoryValue) || EXPORT_CATEGORY_ALL_VALUE;
  var scopedRecords = getRecordsByExportRangeFilters({
    kuwakuValue: kuwakuValue,
    categoryValue: categoryValue,
    statusValue: "all",
    dateFromRaw: filters.dateFromRaw,
    dateToRaw: filters.dateToRaw
  });
  if (!scopedRecords.length) {
    return [];
  }
  var modeSelections = filters.modeSelections || {};
  var groups = [];
  var unitSelection = modeSelections.unit || {};
  var detailSelection = modeSelections.detail || {};
  var detailSubSelection = modeSelections.detailSub || {};
  if (unitSelection.enabled) {
    var selectedUnitValues = Array.from(normalizeSelectionSet(unitSelection.unitValues)).sort(function (a, b) {
      return unitLabelForSelect(a).localeCompare(unitLabelForSelect(b), "ja", {
        numeric: true,
        sensitivity: "base"
      });
    });
    selectedUnitValues.forEach(function (unitValue) {
      var records = filterPlanRecordsForMode(scopedRecords, {
        unitValues: [unitValue]
      });
      if (!records.length) {
        return;
      }
      var drawables = records.map(function (record) {
        return buildPlanDrawable(record);
      }).filter(Boolean);
      if (!drawables.length) {
        return;
      }
      groups.push({
        kuwakuValue: kuwakuValue,
        modeLabel: "ユニット別",
        unitValue: unitValue,
        detailValue: "",
        detailSubValue: "",
        unitLabel: unitLabelForSelect(unitValue),
        detailLabel: "-",
        detailSubLabel: "-",
        drawables: drawables,
        count: records.length
      });
    });
  }
  if (detailSelection.enabled) {
    var selectedUnitValue = value(detailSelection.unitValue);
    if (selectedUnitValue) {
      var selectedDetailValues = Array.from(normalizeSelectionSet(detailSelection.detailValues));
      var unitRecords = filterPlanRecordsForMode(scopedRecords, {
        unitValues: [selectedUnitValue]
      });
      selectedDetailValues.sort(function (a, b) {
        return detailLabelForSelect(a).localeCompare(detailLabelForSelect(b), "ja", {
          numeric: true,
          sensitivity: "base"
        });
      }).forEach(function (detailValue) {
        var records = filterPlanRecordsForMode(unitRecords, {
          detailValues: [detailValue]
        });
        if (!records.length) {
          return;
        }
        var drawables = records.map(function (record) {
          return buildPlanDrawable(record);
        }).filter(Boolean);
        if (!drawables.length) {
          return;
        }
        groups.push({
          kuwakuValue: kuwakuValue,
          modeLabel: "サブユニット別",
          unitValue: selectedUnitValue,
          detailValue: detailValue,
          detailSubValue: "",
          unitLabel: unitLabelForSelect(selectedUnitValue),
          detailLabel: detailLabelForSelect(detailValue),
          detailSubLabel: "-",
          drawables: drawables,
          count: records.length
        });
      });
    }
  }
  if (detailSubSelection.enabled) {
    var _selectedUnitValue = value(detailSubSelection.unitValue);
    var selectedDetailValue = value(detailSubSelection.detailValue);
    if (_selectedUnitValue && selectedDetailValue) {
      var selectedDetailSubValues = Array.from(normalizeSelectionSet(detailSubSelection.detailSubValues));
      var _unitRecords = filterPlanRecordsForMode(scopedRecords, {
        unitValues: [_selectedUnitValue]
      });
      var detailRecords = filterPlanRecordsForMode(_unitRecords, {
        detailValues: [selectedDetailValue]
      });
      selectedDetailSubValues.sort(function (a, b) {
        return detailSubLabelForSelect(a).localeCompare(detailSubLabelForSelect(b), "ja", {
          numeric: true,
          sensitivity: "base"
        });
      }).forEach(function (detailSubValue) {
        var records = filterPlanRecordsForMode(detailRecords, {
          detailSubValues: [detailSubValue]
        });
        if (!records.length) {
          return;
        }
        var drawables = records.map(function (record) {
          return buildPlanDrawable(record);
        }).filter(Boolean);
        if (!drawables.length) {
          return;
        }
        groups.push({
          kuwakuValue: kuwakuValue,
          modeLabel: "サブユニット細分別",
          unitValue: _selectedUnitValue,
          detailValue: selectedDetailValue,
          detailSubValue: detailSubValue,
          unitLabel: unitLabelForSelect(_selectedUnitValue),
          detailLabel: detailLabelForSelect(selectedDetailValue),
          detailSubLabel: detailSubLabelForSelect(detailSubValue),
          drawables: drawables,
          count: records.length
        });
      });
    }
  }
  return groups;
}
function buildPlanPdfHtml(groups) {
  var generatedAt = formatPdfGeneratedAt(new Date());
  return groups.map(function (group, index) {
    var kuwakuValue = value(group.kuwakuValue) || selectedPlanKuwaku;
    var kuwakuLabel = kuwakuValue ? kuwakuLabelForSelect(kuwakuValue) : "-";
    var unitLabel = value(group.unitLabel) || (group.unitValue === ALL_UNITS_VALUE ? "全ユニット" : unitLabelForSelect(group.unitValue));
    var detailLabel = value(group.detailLabel) || (group.detailValue === ALL_DETAILS_VALUE ? "全サブユニット" : detailLabelForSelect(group.detailValue));
    var detailSubLabel = value(group.detailSubLabel) || (group.detailSubValue === ALL_DETAIL_SUBS_VALUE ? "全細分" : detailSubLabelForSelect(group.detailSubValue));
    var mapSvg = buildPlanPdfMapSvg(group.drawables, kuwakuValue);
    return "\n        <section class=\"pdf-page ".concat(index < groups.length - 1 ? "pdf-page-break" : "", "\">\n          <header class=\"pdf-header\">\n            <h1>\u5C64\u6E96\u5225\u5E73\u9762\u56F3</h1>\n            <div class=\"pdf-meta\">\n              <span>\u533A\u753B: ").concat(escapeHtml(kuwakuLabel), "</span>\n              <span>\u51FA\u529B\u30E2\u30FC\u30C9: ").concat(escapeHtml(value(group.modeLabel) || "-"), "</span>\n              <span>\u30E6\u30CB\u30C3\u30C8: ").concat(escapeHtml(unitLabel), "</span>\n              <span>\u30B5\u30D6\u30E6\u30CB\u30C3\u30C8: ").concat(escapeHtml(detailLabel), "</span>\n              <span>\u7D30\u5206: ").concat(escapeHtml(detailSubLabel), "</span>\n              <span>\u4EF6\u6570: ").concat(group.count, "</span>\n              <span>\u51FA\u529B\u65E5\u6642: ").concat(escapeHtml(generatedAt), "</span>\n            </div>\n          </header>\n          <div class=\"pdf-plan-wrap\">\n            ").concat(mapSvg, "\n          </div>\n          <div class=\"pdf-plan-legend\">").concat(buildPlanLegendHtml(), "</div>\n        </section>\n      ");
  }).join("");
}
function buildPlanPdfMapSvg(drawables, kuwakuRaw) {
  var verticalGrid = [100, 200, 300].map(function (x) {
    return "<line class=\"pdf-plan-grid-line\" x1=\"".concat(x, "\" y1=\"0\" x2=\"").concat(x, "\" y2=\"").concat(PLAN_SIZE_CM, "\" />");
  }).join("");
  var horizontalGrid = [100, 200, 300].map(function (y) {
    return "<line class=\"pdf-plan-grid-line\" x1=\"0\" y1=\"".concat(y, "\" x2=\"").concat(PLAN_SIZE_CM, "\" y2=\"").concat(y, "\" />");
  }).join("");
  var pointsSvg = drawables.map(function (drawable, index) {
    return renderPlanPdfDrawableSvg(drawable, index);
  }).join("");
  var cornerLabels = buildPlanCornerLabels(kuwakuRaw);
  return "\n    <div class=\"pdf-plan-shell\">\n      <div class=\"pdf-plan-axis north\">\u5317</div>\n      <div class=\"pdf-plan-axis east\">\u6771</div>\n      <div class=\"pdf-plan-axis south\">\u5357</div>\n      <div class=\"pdf-plan-axis west\">\u897F</div>\n      <div class=\"pdf-plan-corner top-left\">".concat(escapeHtml(cornerLabels.topLeft), "</div>\n      <div class=\"pdf-plan-corner top-right\">").concat(escapeHtml(cornerLabels.topRight), "</div>\n      <div class=\"pdf-plan-corner bottom-left\">").concat(escapeHtml(cornerLabels.bottomLeft), "</div>\n      <div class=\"pdf-plan-corner bottom-right\">").concat(escapeHtml(cornerLabels.bottomRight), "</div>\n      <svg class=\"pdf-plan-svg\" viewBox=\"0 0 ").concat(PLAN_SIZE_CM, " ").concat(PLAN_SIZE_CM, "\" aria-label=\"\u5C64\u6E96\u5225\u5E73\u9762\u56F3\" xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\">\n        <rect class=\"pdf-plan-frame\" x=\"0\" y=\"0\" width=\"").concat(PLAN_SIZE_CM, "\" height=\"").concat(PLAN_SIZE_CM, "\" />\n        ").concat(verticalGrid, "\n        ").concat(horizontalGrid, "\n        ").concat(pointsSvg, "\n      </svg>\n    </div>\n  ");
}
function renderPlanPdfDrawableSvg(drawable) {
  var index = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : 0;
  var labelX = Math.min(PLAN_SIZE_CM - 2, drawable.x + 6);
  var labelY = Math.max(8, drawable.y - 6);
  var shapeSvg = "";
  if (drawable.type === "line") {
    shapeSvg = "<line class=\"pdf-plan-shape-line\" x1=\"".concat(drawable.x1, "\" y1=\"").concat(drawable.y1, "\" x2=\"").concat(drawable.x2, "\" y2=\"").concat(drawable.y2, "\" stroke=\"").concat(drawable.color, "\" />");
  } else if (drawable.type === "multipoint") {
    var points = Array.isArray(drawable.points) ? drawable.points : [];
    var hull = Array.isArray(drawable.hull) ? drawable.hull : [];
    var hullSvg = "";
    if (hull.length >= 3) {
      hullSvg = "<polygon points=\"".concat(hull.map(function (point) {
        return "".concat(point.x, ",").concat(point.y);
      }).join(" "), "\" fill=\"none\" stroke=\"").concat(drawable.color, "\" stroke-width=\"1.8\" />");
    } else if (hull.length === 2) {
      hullSvg = "<line x1=\"".concat(hull[0].x, "\" y1=\"").concat(hull[0].y, "\" x2=\"").concat(hull[1].x, "\" y2=\"").concat(hull[1].y, "\" stroke=\"").concat(drawable.color, "\" stroke-width=\"1.8\" />");
    }
    var pointsSvg = points.map(function (point) {
      return "<circle cx=\"".concat(point.x, "\" cy=\"").concat(point.y, "\" r=\"3.8\" fill=\"").concat(drawable.color, "\" stroke=\"#ffffff\" stroke-width=\"0.9\" />");
    }).join("");
    shapeSvg = "".concat(hullSvg).concat(pointsSvg);
  } else if (drawable.type === "rect") {
    var transform = Number.isFinite(drawable.rotationDeg) ? " transform=\"rotate(".concat(drawable.rotationDeg, " ").concat(drawable.x, " ").concat(drawable.y, ")\"") : "";
    shapeSvg = "<rect class=\"pdf-plan-shape-rect\" x=\"".concat(drawable.left, "\" y=\"").concat(drawable.top, "\" width=\"").concat(drawable.width, "\" height=\"").concat(drawable.height, "\" stroke=\"").concat(drawable.color, "\"").concat(transform, " />");
  } else if (drawable.type === "ellipse") {
    var _transform2 = Number.isFinite(drawable.rotationDeg) ? " transform=\"rotate(".concat(drawable.rotationDeg, " ").concat(drawable.x, " ").concat(drawable.y, ")\"") : "";
    shapeSvg = "<ellipse class=\"pdf-plan-shape-ellipse\" cx=\"".concat(drawable.x, "\" cy=\"").concat(drawable.y, "\" rx=\"").concat(drawable.rx, "\" ry=\"").concat(drawable.ry, "\" stroke=\"").concat(drawable.color, "\"").concat(_transform2, " />");
  } else if (drawable.type === "imageQuad") {
    shapeSvg = buildPlanImageWarpSvg(drawable, "pdf-".concat(index));
  } else {
    shapeSvg = "<circle cx=\"".concat(drawable.x, "\" cy=\"").concat(drawable.y, "\" r=\"5\" fill=\"").concat(drawable.color, "\" />");
  }
  return "\n    <g>\n      ".concat(shapeSvg, "\n      <text x=\"").concat(labelX, "\" y=\"").concat(labelY, "\">").concat(escapeHtml(drawable.label || ""), "</text>\n    </g>\n  ");
}
function buildPdfImageGrid(itemsRaw, emptyText) {
  var gridClassRaw = arguments.length > 2 && arguments[2] !== undefined ? arguments[2] : "";
  var items = Array.isArray(itemsRaw) ? itemsRaw : [];
  if (!items.length) {
    return "<p class=\"pdf-muted\">".concat(escapeHtml(emptyText), "</p>");
  }
  var gridClass = value(gridClassRaw).replace(/[^a-zA-Z0-9_-]/g, "");
  var classNames = gridClass ? "pdf-image-grid ".concat(gridClass) : "pdf-image-grid";
  return "\n    <div class=\"".concat(classNames, "\">\n      ").concat(items.map(function (item) {
    return "\n            <figure>\n              <img src=\"".concat(item.dataUrl, "\" alt=\"").concat(escapeHtml(item.name || "image"), "\" />\n              <figcaption>").concat(escapeHtml(item.caption || ""), "</figcaption>\n            </figure>\n          ");
  }).join(""), "\n    </div>\n  ");
}
function openPdfPrintWindow(_ref16) {
  var title = _ref16.title,
    pageSize = _ref16.pageSize,
    bodyHtml = _ref16.bodyHtml;
  var safeTitle = escapeHtml(title || "出力");
  var safeBody = bodyHtml || "<p>出力データがありません。</p>";
  var htmlText = "\n    <!doctype html>\n    <html lang=\"ja\">\n      <head>\n        <meta charset=\"utf-8\" />\n        <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\" />\n        <title>".concat(safeTitle, "</title>\n        <style>").concat(buildPdfPrintStyles(pageSize), "</style>\n      </head>\n      <body>").concat(safeBody, "</body>\n    </html>\n  ");
  var printWindow = window.open("", "_blank");
  if (printWindow) {
    printWindow.document.open();
    printWindow.document.write(htmlText);
    printWindow.document.close();
    var hasPrinted = false;
    var triggerPrint = function triggerPrint() {
      if (hasPrinted) {
        return;
      }
      hasPrinted = true;
      try {
        printWindow.focus();
        printWindow.print();
      } catch (_error) {}
    };
    printWindow.onload = function () {
      window.setTimeout(triggerPrint, 220);
    };
    window.setTimeout(triggerPrint, 900);
    return true;
  }
  try {
    var frame = document.createElement("iframe");
    frame.style.position = "fixed";
    frame.style.right = "0";
    frame.style.bottom = "0";
    frame.style.width = "0";
    frame.style.height = "0";
    frame.style.border = "0";
    frame.setAttribute("aria-hidden", "true");
    frame.srcdoc = htmlText;
    document.body.appendChild(frame);
    frame.onload = function () {
      var frameWindow = frame.contentWindow;
      if (frameWindow) {
        try {
          frameWindow.focus();
          frameWindow.print();
        } catch (_error) {}
      }
      window.setTimeout(function () {
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
  var pageSize = value(pageSizeRaw) || "A4 portrait";
  return "\n    @page {\n      size: ".concat(pageSize, ";\n      margin: 10mm;\n    }\n    * { box-sizing: border-box; }\n    body {\n      margin: 0;\n      font-family: \"Hiragino Kaku Gothic ProN\", \"Yu Gothic\", sans-serif;\n      color: #111827;\n      font-size: 11px;\n      line-height: 1.45;\n    }\n    .pdf-page {\n      width: 100%;\n    }\n    .pdf-page-break {\n      page-break-after: always;\n    }\n    .pdf-header h1 {\n      margin: 0;\n      font-size: 18px;\n    }\n    .pdf-meta {\n      margin-top: 6px;\n      display: flex;\n      flex-wrap: wrap;\n      gap: 8px 12px;\n      color: #334155;\n      font-size: 12px;\n    }\n    .pdf-table {\n      width: 100%;\n      border-collapse: collapse;\n      margin-top: 8px;\n    }\n    .pdf-table th, .pdf-table td {\n      border: 1px solid #9ca3af;\n      padding: 4px 6px;\n      vertical-align: top;\n    }\n    .pdf-table th {\n      background: #f3f4f6;\n      white-space: nowrap;\n    }\n    .pdf-list-table td:nth-child(14) {\n      text-align: center;\n      font-weight: 700;\n    }\n    .pdf-status-complete {\n      color: #111827;\n    }\n    .pdf-status-incomplete {\n      color: #b42318;\n    }\n    .pdf-image-section h2 {\n      margin: 10px 0 5px;\n      font-size: 12px;\n    }\n    .pdf-image-grid {\n      display: grid;\n      grid-template-columns: repeat(auto-fill, minmax(42mm, 1fr));\n      gap: 6px;\n    }\n    .pdf-image-grid figure {\n      margin: 0;\n      border: 1px solid #d1d5db;\n      border-radius: 4px;\n      padding: 4px;\n    }\n    .pdf-image-grid img {\n      width: 100%;\n      height: auto;\n      max-height: 38mm;\n      object-fit: cover;\n      display: block;\n    }\n    .pdf-image-grid figcaption {\n      margin-top: 3px;\n      color: #475569;\n      font-size: 10px;\n    }\n    .pdf-card-image-grid {\n      grid-template-columns: 1fr;\n      gap: 8px;\n    }\n    .pdf-card-image-grid img {\n      max-height: 72mm;\n      object-fit: contain;\n      background: #f8fafc;\n    }\n    .pdf-muted {\n      color: #6b7280;\n      margin: 4px 0;\n    }\n    .pdf-plan-wrap {\n      margin-top: 8px;\n      display: flex;\n      justify-content: center;\n    }\n    .pdf-plan-shell {\n      position: relative;\n      width: 162mm;\n      height: 162mm;\n      border: 1px solid #9ca3af;\n      background: #fff;\n      padding: 12mm;\n    }\n    .pdf-plan-svg {\n      width: 100%;\n      height: 100%;\n      display: block;\n      overflow: visible;\n    }\n    .pdf-plan-svg .pdf-plan-frame {\n      fill: #fff;\n      stroke: #334155;\n      stroke-width: 2;\n    }\n    .pdf-plan-svg .pdf-plan-grid-line {\n      stroke: #94a3b8;\n      stroke-width: 0.9;\n      stroke-dasharray: 4 3;\n    }\n    .pdf-plan-svg text {\n      font-size: 12px;\n      font-weight: 700;\n      fill: #0f172a;\n    }\n    .pdf-plan-svg .pdf-plan-shape-line {\n      stroke-width: 4;\n      fill: none;\n      stroke-linecap: round;\n      stroke-dasharray: none;\n    }\n    .pdf-plan-svg .pdf-plan-shape-rect,\n    .pdf-plan-svg .pdf-plan-shape-ellipse {\n      stroke-width: 3.2;\n      fill: none;\n    }\n    .pdf-plan-axis {\n      position: absolute;\n      font-size: 13px;\n      font-weight: 700;\n      color: #1f2937;\n    }\n    .pdf-plan-axis.north { top: 5px; left: 50%; transform: translateX(-50%); }\n    .pdf-plan-axis.south { bottom: 5px; left: 50%; transform: translateX(-50%); }\n    .pdf-plan-axis.east { right: 5px; top: 50%; transform: translateY(-50%); }\n    .pdf-plan-axis.west { left: 5px; top: 50%; transform: translateY(-50%); }\n    .pdf-plan-corner {\n      position: absolute;\n      font-size: 12px;\n      font-weight: 700;\n      background: #eff6ff;\n      border: 1px solid #93c5fd;\n      border-radius: 4px;\n      padding: 1px 5px;\n    }\n    .pdf-plan-corner.top-left { top: 10mm; left: 10mm; }\n    .pdf-plan-corner.top-right { top: 10mm; right: 10mm; }\n    .pdf-plan-corner.bottom-left { bottom: 10mm; left: 10mm; }\n    .pdf-plan-corner.bottom-right { bottom: 10mm; right: 10mm; }\n    .pdf-plan-legend {\n      margin-top: 6px;\n      display: flex;\n      flex-wrap: wrap;\n      gap: 6px;\n      justify-content: center;\n    }\n    .plan-legend-item {\n      display: inline-flex;\n      align-items: center;\n      gap: 4px;\n      border: 1px solid #cbd5e1;\n      border-radius: 999px;\n      padding: 2px 8px;\n      font-size: 11px;\n      color: #334155;\n      background: #f8fafc;\n    }\n    .plan-legend-dot {\n      width: 10px;\n      height: 10px;\n      border-radius: 50%;\n      border: 1px solid rgba(0, 0, 0, 0.25);\n      display: inline-block;\n    }\n  ");
}
function formatPdfGeneratedAt(dateRaw) {
  var date = dateRaw instanceof Date ? dateRaw : new Date();
  if (!Number.isFinite(date.getTime())) {
    return "-";
  }
  var yyyy = date.getFullYear();
  var mm = String(date.getMonth() + 1).padStart(2, "0");
  var dd = String(date.getDate()).padStart(2, "0");
  var hh = String(date.getHours()).padStart(2, "0");
  var min = String(date.getMinutes()).padStart(2, "0");
  return "".concat(yyyy, "/").concat(mm, "/").concat(dd, " ").concat(hh, ":").concat(min);
}
function csvCell(valueRaw) {
  var valueText = value(valueRaw).replaceAll("\r\n", "\n").replaceAll("\r", "\n");
  var escaped = valueText.replaceAll('"', '""');
  return "\"".concat(escaped, "\"");
}
function compactNoSpaceValue(inputRaw) {
  return value(inputRaw).replace(/\s+/g, "");
}
function normalizeKuwakuHeadA(headARaw) {
  return compactNoSpaceValue(headARaw);
}
function normalizeKuwakuHeadB(headBRaw) {
  var headB = compactNoSpaceValue(headBRaw);
  var upper = headB.toUpperCase();
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
  var kuwaku = value(kuwakuRaw);
  if (!kuwaku) {
    return "";
  }
  var parts = parseKuwaku(kuwaku);
  return buildKuwaku(parts.headA, parts.headB, parts.block, parts.no);
}
function buildKuwaku(headARaw, headBRaw, blockRaw, noRaw) {
  var headA = normalizeKuwakuHeadA(headARaw) || DEFAULT_KUWAKU_HEAD_A;
  var headB = normalizeKuwakuHeadB(headBRaw) || DEFAULT_KUWAKU_HEAD_B;
  var block = normalizeKuwakuBlock(blockRaw);
  var no = normalizeKuwakuNo(noRaw);
  return "".concat(headA, "-").concat(headB, "-").concat(block, "-").concat(no);
}
function parseKuwaku(kuwakuText) {
  var text = value(kuwakuText).replaceAll("－", "-").replaceAll("―", "-").replaceAll("ー", "-").replaceAll("−", "-");
  var parts = text.split("-").map(function (part) {
    return compactNoSpaceValue(part);
  });
  if (parts.length >= 4) {
    return {
      headA: normalizeKuwakuHeadA(parts[0]) || DEFAULT_KUWAKU_HEAD_A,
      headB: normalizeKuwakuHeadB(parts[1]) || DEFAULT_KUWAKU_HEAD_B,
      block: normalizeKuwakuBlock(parts[2]),
      no: normalizeKuwakuNo(parts[3])
    };
  }
  return {
    headA: DEFAULT_KUWAKU_HEAD_A,
    headB: DEFAULT_KUWAKU_HEAD_B,
    block: "",
    no: ""
  };
}
function normalizeSpecimenPrefix(prefixRaw) {
  var prefix = compactNoSpaceValue(prefixRaw).toLowerCase();
  if (prefix === "ii") {
    prefix = "i";
  }
  return VALID_SPECIMEN_PREFIXES.has(prefix) ? prefix : DEFAULT_SPECIMEN_PREFIX;
}
function buildSpecimenNo(prefixRaw, serialRaw) {
  var prefix = normalizeSpecimenPrefix(prefixRaw);
  var serial = compactNoSpaceValue(serialRaw);
  return serial ? "".concat(prefix, "-").concat(serial) : "";
}
function parseSpecimenNo(specimenNoRaw) {
  var prefixRaw = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : "";
  var serialRaw = arguments.length > 2 && arguments[2] !== undefined ? arguments[2] : "";
  var directPrefix = normalizeSpecimenPrefix(compactNoSpaceValue(prefixRaw));
  var directSerial = compactNoSpaceValue(serialRaw);
  if (directSerial) {
    return {
      prefix: directPrefix,
      serial: directSerial,
      specimenNo: buildSpecimenNo(directPrefix, directSerial)
    };
  }
  var specimenNo = compactNoSpaceValue(specimenNoRaw);
  var hyphenMatched = specimenNo.match(/^([A-Za-z]{1,2})-(.+)$/);
  if (hyphenMatched) {
    var prefix = normalizeSpecimenPrefix(hyphenMatched[1]);
    var serial = compactNoSpaceValue(hyphenMatched[2]);
    return {
      prefix: prefix,
      serial: serial,
      specimenNo: buildSpecimenNo(prefix, serial)
    };
  }
  var compactMatched = specimenNo.match(/^([A-Za-z]{1,2})(\d.+)$/);
  if (compactMatched) {
    var _prefix = normalizeSpecimenPrefix(compactMatched[1]);
    var _serial = compactNoSpaceValue(compactMatched[2]);
    return {
      prefix: _prefix,
      serial: _serial,
      specimenNo: buildSpecimenNo(_prefix, _serial)
    };
  }
  var fallbackPrefix = normalizeSpecimenPrefix(compactNoSpaceValue(prefixRaw));
  return {
    prefix: fallbackPrefix,
    serial: specimenNo,
    specimenNo: buildSpecimenNo(fallbackPrefix, specimenNo)
  };
}
function compareRecordsBySpecimenNo(a, b) {
  var aSpecimen = parseSpecimenNo(a === null || a === void 0 ? void 0 : a.specimenNo, a === null || a === void 0 ? void 0 : a.specimenPrefix, a === null || a === void 0 ? void 0 : a.specimenSerial);
  var bSpecimen = parseSpecimenNo(b === null || b === void 0 ? void 0 : b.specimenNo, b === null || b === void 0 ? void 0 : b.specimenPrefix, b === null || b === void 0 ? void 0 : b.specimenSerial);
  var prefixCompared = aSpecimen.prefix.localeCompare(bSpecimen.prefix, "ja", {
    sensitivity: "base"
  });
  if (prefixCompared !== 0) {
    return prefixCompared;
  }
  var aSerial = value(aSpecimen.serial);
  var bSerial = value(bSpecimen.serial);
  var aIsNumber = /^\d+$/.test(aSerial);
  var bIsNumber = /^\d+$/.test(bSerial);
  if (aIsNumber && bIsNumber) {
    var diff = Number(aSerial) - Number(bSerial);
    if (diff !== 0) {
      return diff;
    }
  } else {
    var serialCompared = aSerial.localeCompare(bSerial, "ja", {
      numeric: true,
      sensitivity: "base"
    });
    if (serialCompared !== 0) {
      return serialCompared;
    }
  }
  return value(a === null || a === void 0 ? void 0 : a.id).localeCompare(value(b === null || b === void 0 ? void 0 : b.id), "ja", {
    sensitivity: "base"
  });
}
function compareRecordsByKuwakuThenSpecimen(a, b) {
  var aKuwaku = kuwakuLabelForSelect(kuwakuValueForSelect(getRecordKuwaku(a)));
  var bKuwaku = kuwakuLabelForSelect(kuwakuValueForSelect(getRecordKuwaku(b)));
  var kuwakuCompared = aKuwaku.localeCompare(bKuwaku, "ja", {
    numeric: true,
    sensitivity: "base"
  });
  if (kuwakuCompared !== 0) {
    return kuwakuCompared;
  }
  return compareRecordsBySpecimenNo(a, b);
}
function categoryFromPrefix(prefixRaw) {
  var prefix = normalizeSpecimenPrefix(prefixRaw);
  return "".concat(prefix, ": ").concat(SPECIMEN_CATEGORY_MAP[prefix] || "");
}
function normalizeCategory(categoryRaw, prefixRaw) {
  var categoryText = value(categoryRaw);
  var matched = categoryText.match(/^([A-Za-z]{1,2})\s*:\s*(.*)$/);
  if (matched) {
    var prefix = normalizeSpecimenPrefix(matched[1]);
    return "".concat(prefix, ": ").concat(SPECIMEN_CATEGORY_MAP[prefix] || value(matched[2]));
  }
  if (categoryText) {
    return categoryText;
  }
  return categoryFromPrefix(prefixRaw);
}
function normalizeAnalysisType(typeRaw) {
  var text = value(typeRaw);
  if (!text) {
    return "";
  }
  var matched = text.match(/^([A-Za-z]{1,2})\s*:/);
  var code = (matched ? matched[1] : text).replaceAll(" ", "").toUpperCase();
  if (!ANALYSIS_TYPE_MAP[code]) {
    return "";
  }
  var displayCode = code === "MG" ? "Mg" : code;
  return "".concat(displayCode, ": ").concat(ANALYSIS_TYPE_MAP[code]);
}
function extractAnalysisTypeFromCategory(categoryRaw) {
  var text = value(categoryRaw);
  if (!text) {
    return "";
  }
  var slashIndex = text.indexOf("/");
  if (slashIndex < 0 && /^a\s*:/i.test(text)) {
    return "";
  }
  var candidate = slashIndex >= 0 ? text.slice(slashIndex + 1) : text;
  var matched = candidate.match(/([A-Za-z]{1,2})\s*:/);
  if (!matched) {
    return "";
  }
  return normalizeAnalysisType(matched[1]);
}
function formatCategoryForRecord(record) {
  var base = normalizeCategory(value(record === null || record === void 0 ? void 0 : record.category), value(record === null || record === void 0 ? void 0 : record.specimenPrefix));
  var prefix = normalizeSpecimenPrefix(value(record === null || record === void 0 ? void 0 : record.specimenPrefix));
  if (prefix !== "a") {
    return base;
  }
  var analysisType = normalizeAnalysisType(value(record === null || record === void 0 ? void 0 : record.analysisType));
  if (!analysisType) {
    return base;
  }
  return "".concat(base, " / ").concat(analysisType);
}
function normalizeLayerName(layerRaw) {
  var layer = value(layerRaw);
  if (!layer) {
    return "";
  }
  if (LEGACY_LAYER_NAME_ALIASES[layer]) {
    return LEGACY_LAYER_NAME_ALIASES[layer];
  }
  var legacyEntries = Object.entries(LEGACY_LAYER_NAME_ALIASES);
  for (var _i5 = 0, _legacyEntries = legacyEntries; _i5 < _legacyEntries.length; _i5++) {
    var _legacyEntries$_i = _slicedToArray(_legacyEntries[_i5], 2),
      legacy = _legacyEntries$_i[0],
      latest = _legacyEntries$_i[1];
    if (layer.startsWith("".concat(legacy, ":"))) {
      return "".concat(latest, ":").concat(layer.slice("".concat(legacy, ":").length));
    }
    if (layer.startsWith("".concat(legacy, "\uFF1A"))) {
      return "".concat(latest, "\uFF1A").concat(layer.slice("".concat(legacy, "\uFF1A").length));
    }
    if (layer.startsWith("".concat(legacy, "(")) && layer.endsWith(")")) {
      return "".concat(latest).concat(layer.slice(legacy.length));
    }
  }
  return layer;
}
function extractOtherLayerText(layerRaw) {
  var layer = normalizeLayerName(layerRaw);
  if (!layer) {
    return "";
  }
  if (layer.startsWith("".concat(OTHER_LAYER_NAME, ":"))) {
    return layer.slice("".concat(OTHER_LAYER_NAME, ":").length).trim();
  }
  if (layer.startsWith("".concat(OTHER_LAYER_NAME, "\uFF1A"))) {
    return layer.slice("".concat(OTHER_LAYER_NAME, "\uFF1A").length).trim();
  }
  if (layer.startsWith("".concat(OTHER_LAYER_NAME, "(")) && layer.endsWith(")")) {
    return layer.slice(OTHER_LAYER_NAME.length + 1, -1).trim();
  }
  if (layer === OTHER_LAYER_NAME) {
    return "";
  }
  return layer;
}
function extractOtherTeamText(teamRaw) {
  var team = value(teamRaw);
  if (!team) {
    return "";
  }
  if (team.startsWith("".concat(OTHER_TEAM_NAME, ":"))) {
    return team.slice("".concat(OTHER_TEAM_NAME, ":").length).trim();
  }
  if (team.startsWith("".concat(OTHER_TEAM_NAME, "\uFF1A"))) {
    return team.slice("".concat(OTHER_TEAM_NAME, "\uFF1A").length).trim();
  }
  if (team === OTHER_TEAM_NAME) {
    return "";
  }
  return team;
}
function normalizeTeamState(teamRaw) {
  var teamOtherRaw = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : "";
  var team = value(teamRaw);
  var teamOther = value(teamOtherRaw);
  if (!team) {
    return {
      team: "",
      teamOther: ""
    };
  }
  if (PRESET_TEAMS.includes(team) && team !== OTHER_TEAM_NAME) {
    return {
      team: team,
      teamOther: ""
    };
  }
  if (team === OTHER_TEAM_NAME) {
    return {
      team: OTHER_TEAM_NAME,
      teamOther: teamOther
    };
  }
  return {
    team: OTHER_TEAM_NAME,
    teamOther: teamOther || extractOtherTeamText(team)
  };
}
function syncTeamOtherInput(teamRaw) {
  var team = value(teamRaw);
  var isOther = team === OTHER_TEAM_NAME;
  teamOtherInput.classList.toggle("hidden", !isOther);
  if (!isOther) {
    teamOtherInput.value = "";
  }
}
function syncEditTeamOtherInput(teamRaw) {
  if (!editTeamOtherInput) {
    return;
  }
  var team = value(teamRaw);
  var isOther = team === OTHER_TEAM_NAME;
  editTeamOtherInput.classList.toggle("hidden", !isOther);
  if (!isOther) {
    editTeamOtherInput.value = "";
  }
}
function formatTeamValue(site) {
  var teamState = normalizeTeamState(site === null || site === void 0 ? void 0 : site.team, site === null || site === void 0 ? void 0 : site.teamOther);
  if (teamState.team !== OTHER_TEAM_NAME) {
    return teamState.team;
  }
  return teamState.teamOther ? "".concat(OTHER_TEAM_NAME, ":").concat(teamState.teamOther) : OTHER_TEAM_NAME;
}
function validateInputRequiredFields(siteSnapshot, recordFormData) {
  if (!siteSnapshot) {
    return "入力情報を取得できませんでした";
  }
  var siteRequiredFields = [["区画（グリッド）名の1番目", siteSnapshot.kuwakuHeadA], ["区画（グリッド）名の2番目", siteSnapshot.kuwakuHeadB], ["区画（グリッド）の英字", siteSnapshot.kuwakuBlock], ["区画（グリッド）の番号", siteSnapshot.kuwakuNo], ["レベル高", siteSnapshot.levelHeight], ["日付", siteSnapshot.date], ["発掘班", siteSnapshot.team], ["記載者", siteSnapshot.scribe]];
  for (var _i6 = 0, _siteRequiredFields = siteRequiredFields; _i6 < _siteRequiredFields.length; _i6++) {
    var _siteRequiredFields$_ = _slicedToArray(_siteRequiredFields[_i6], 2),
      label = _siteRequiredFields$_[0],
      fieldValue = _siteRequiredFields$_[1];
    if (!value(fieldValue)) {
      return "".concat(label, "\u3092\u5165\u529B\u3057\u3066\u304F\u3060\u3055\u3044");
    }
  }
  if (siteSnapshot.team === OTHER_TEAM_NAME && !value(siteSnapshot.teamOther)) {
    return "発掘班が「その他」の場合は内容を入力してください";
  }
  var selectedLayerName = getSelectedLayerName();
  var recordRequiredFields = [["標本番号", recordFormData.get("specimenSerial")], ["化石・遺物名称", recordFormData.get("nameMemo")], ["重要品指定", recordFormData.get("importantFlag")], ["簡易記載（専門班の指示による）", recordFormData.get("simpleRecordFlag")], ["発見者氏名", recordFormData.get("discoverer")], ["判定者氏名", recordFormData.get("identifier")], ["レベル読値（上面）", recordFormData.get("levelUpperCm")], ["レベル読値（下底）", recordFormData.get("levelLowerCm")], ["産出状況断面", recordFormData.get("occurrenceSection")], ["産状スケッチ", recordFormData.get("occurrenceSketch")], ["平面位置（北から/南から）", recordFormData.get("nsDir")], ["平面位置（北から/南からの距離）", recordFormData.get("nsCm")], ["平面位置（東から/西から）", recordFormData.get("ewDir")], ["平面位置（東から/西からの距離）", recordFormData.get("ewCm")], ["地層名", selectedLayerName], ["ユニット", recordFormData.get("unit")], ["層理面や鍵層名", recordFormData.get("layerRef")], ["地層中の位置（上/下）", recordFormData.get("layerRelative")], ["地層中の位置（cm）", recordFormData.get("layerFromCm")]];
  for (var _i7 = 0, _recordRequiredFields = recordRequiredFields; _i7 < _recordRequiredFields.length; _i7++) {
    var _recordRequiredFields2 = _slicedToArray(_recordRequiredFields[_i7], 2),
      _label = _recordRequiredFields2[0],
      _fieldValue = _recordRequiredFields2[1];
    if (!value(_fieldValue)) {
      return "".concat(_label, "\u3092\u5165\u529B\u3057\u3066\u304F\u3060\u3055\u3044");
    }
  }
  if (selectedLayerName === OTHER_LAYER_NAME && !value(layerOtherInput.value)) {
    return "地層名が「4.その他」の場合は内容を入力してください";
  }
  var specimenPrefix = normalizeSpecimenPrefix(value(recordFormData.get("specimenPrefix")));
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
  var missing = new Set();
  if (!record) {
    return missing;
  }
  var kuwaku = parseKuwaku(value(record.kuwaku));
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
  var specimen = parseSpecimenNo(record.specimenNo, record.specimenPrefix, record.specimenSerial);
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
  var planSizeMode = normalizePlanSizeMode(value(record.planSizeMode));
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
    var largeShapeType = normalizeLargeShapeType(value(record.largeShapeType));
    var isImageShape = isLargeShapeImageType(largeShapeType);
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
      var plungeDeg = parseLargeAxisPlungeDeg(value(record.largeAxisPlungeDeg));
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
      var dipDeg = parseLargeAxisPlungeDeg(value(record.planeDipDeg));
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
      var _dipDeg = parseLargeAxisPlungeDeg(value(record.planeDipDeg));
      if (Number.isFinite(_dipDeg) && _dipDeg > 0 && !normalizeCompass8Direction(value(record.planeDipDir8))) {
        missing.add("planeDipDir8");
      }
    } else if (isImageShape) {
      var hasFrameSpec = parseDistanceToCm(record.imgFrameWidthCm) != null && parseDistanceToCm(record.imgFrameHeightCm) != null && normalizeImageRotationDeg(record.imgRotateDeg) !== "";
      if (!hasFrameSpec) {
        ["imgRotateDeg", "imgFrameWidthCm", "imgFrameHeightCm"].forEach(function (key) {
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
      var _dipDeg2 = parseLargeAxisPlungeDeg(value(record.planeDipDeg));
      if (Number.isFinite(_dipDeg2) && _dipDeg2 > 0 && !normalizeCompass8Direction(value(record.planeDipDir8))) {
        missing.add("planeDipDir8");
      }
    }
  }
  var layerName = normalizeLayerName(value(record.layerName));
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
  var specimenPrefix = normalizeSpecimenPrefix(value(record.specimenPrefix));
  if (specimenPrefix === "a" && !normalizeAnalysisType(value(record.analysisType))) {
    missing.add("analysisType");
  }
  var sectionDiagrams = Array.isArray(record.sectionDiagrams) ? record.sectionDiagrams : [];
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
  var photos = Array.isArray(record.photos) ? record.photos : [];
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
  var text = value(valueRaw);
  if (text === "有" || text === "無") {
    return text;
  }
  return "";
}
function normalizeCircleDashFlag(valueRaw) {
  var text = value(valueRaw);
  if (text === "○" || text === "◯") {
    return "○";
  }
  return "-";
}
function normalizeChecklistChecked(valueRaw) {
  var text = value(valueRaw).toLowerCase();
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
  return normalizeChecklistChecked(formData.get("sectionDiagramDistanceChecked")) === "1" && normalizeChecklistChecked(formData.get("sectionDiagramHorizonChecked")) === "1" && normalizeChecklistChecked(formData.get("sectionDiagramLayerFaciesChecked")) === "1";
}
function arePhotoChecklistComplete(formData) {
  return normalizeChecklistChecked(formData.get("photoClinometerChecked")) === "1" && normalizeChecklistChecked(formData.get("photoRulerChecked")) === "1";
}
function normalizeNsDir(valueRaw) {
  var dir = value(valueRaw);
  if (dir === "南" || dir === "南から") {
    return "南から";
  }
  return "北から";
}
function normalizeEwDir(valueRaw) {
  var dir = value(valueRaw);
  if (dir === "西" || dir === "西から") {
    return "西から";
  }
  return "東から";
}
function normalizeCompass8Direction(valueRaw) {
  var raw = value(valueRaw);
  if (!raw) {
    return "";
  }
  var text = raw.toUpperCase().replace(/\s+/g, "");
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
  var direction = normalizeCompass8Direction(valueRaw);
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
  var mode = value(valueRaw);
  if (mode === "複数点" || mode === "複数") {
    return "複数点";
  }
  if (mode === "大きなもの" || mode === "大きいもの") {
    return "大きなもの";
  }
  return "通常";
}
function normalizeLargeShapeType(valueRaw) {
  var raw = value(valueRaw);
  var normalizedRaw = typeof raw.normalize === "function" ? raw.normalize("NFC") : raw;
  var text = normalizeLargeShapeLabel(normalizedRaw);
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
  var shapeType = normalizeLargeShapeType(shapeTypeRaw);
  if (!shapeType) {
    return false;
  }
  return shapeType === CUSTOM_LARGE_SHAPE_TYPE || largeShapeImagePathMap.has(shapeType);
}
function getLargeShapeImagePath(shapeTypeRaw) {
  var shapeType = normalizeLargeShapeType(shapeTypeRaw);
  if (!shapeType) {
    return "";
  }
  var fallbackList = Array.isArray(LARGE_SHAPE_IMAGE_FALLBACK_PATHS[shapeType]) ? LARGE_SHAPE_IMAGE_FALLBACK_PATHS[shapeType] : [];
  var _iterator7 = _createForOfIteratorHelper(fallbackList),
    _step9;
  try {
    for (_iterator7.s(); !(_step9 = _iterator7.n()).done;) {
      var fallback = _step9.value;
      var safe = toSafeAssetUrl(fallback);
      if (safe) {
        return safe;
      }
    }
  } catch (err) {
    _iterator7.e(err);
  } finally {
    _iterator7.f();
  }
  return toSafeAssetUrl(largeShapeImagePathMap.get(shapeType) || "");
}
function getLargeShapeImagePathCandidates(shapeTypeRaw) {
  var imagePathRaw = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : "";
  var candidates = [];
  var pushCandidate = function pushCandidate(pathRaw) {
    var safe = toSafeAssetUrl(pathRaw);
    if (!safe || candidates.includes(safe)) {
      return;
    }
    candidates.push(safe);
  };
  var pushInlineCandidate = function pushInlineCandidate(pathRaw) {
    var inline = getInlineLargeShapeDataUrl(pathRaw);
    if (!inline) {
      return;
    }
    pushCandidate(inline);
  };
  var shapeType = normalizeLargeShapeType(shapeTypeRaw);
  var explicitPath = toSafeAssetUrl(imagePathRaw);
  var hasExplicitPath = Boolean(explicitPath);
  if (shapeType) {
    var fallbackList = LARGE_SHAPE_IMAGE_FALLBACK_PATHS[shapeType] || [];
    var mappedPath = largeShapeImagePathMap.get(shapeType) || "";
    pushInlineCandidate(mappedPath);
    pushCandidate(mappedPath);
    fallbackList.forEach(function (pathRaw) {
      pushInlineCandidate(pathRaw);
      pushCandidate(pathRaw);
    });
  }
  if (hasExplicitPath) {
    pushInlineCandidate(explicitPath);
    pushCandidate(explicitPath);
  }
  return candidates;
}
function normalizeLargeAxisDirection(valueRaw) {
  var text = value(valueRaw).toUpperCase().replace(/\s+/g, "").replace(/[°度]/g, "").replace(/[-_]/g, "");
  if (!text) {
    return "";
  }
  if (text === "NS" || text === "SN") {
    return "NS";
  }
  if (text === "EW" || text === "WE") {
    return "EW";
  }
  var matched = text.match(/^([NS])(\d+(?:\.\d+)?)([EW])$/);
  if (!matched) {
    return text;
  }
  var _matched2 = _slicedToArray(matched, 4),
    ns = _matched2[1],
    degreeRaw = _matched2[2],
    ew = _matched2[3];
  var degree = Number(degreeRaw);
  if (!Number.isFinite(degree)) {
    return text;
  }
  var degreeText = Number.isInteger(degree) ? String(degree) : String(degree).replace(/\.?0+$/, "");
  return "".concat(ns).concat(degreeText).concat(ew);
}
function normalizeLargeAxisPlungeDeg(valueRaw) {
  var plunge = parseLargeAxisPlungeDeg(valueRaw);
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
  var text = value(valueRaw);
  if (text === "上") {
    return "上";
  }
  if (text === "下") {
    return "下";
  }
  var hasUpper = text.includes("上");
  var hasLower = text.includes("下");
  if (hasUpper && !hasLower) {
    return "上";
  }
  if (hasLower && !hasUpper) {
    return "下";
  }
  return "";
}
function formatLevelRead(record) {
  var upper = value(record === null || record === void 0 ? void 0 : record.levelUpperCm);
  var lower = value(record === null || record === void 0 ? void 0 : record.levelLowerCm);
  if (!upper && !lower) {
    return "";
  }
  return "".concat(formatCmValue(upper, "-"), " / ").concat(formatCmValue(lower, "-"));
}
function getRecordAltitudeMValue(record) {
  var useDirectAltitude = normalizeToggleFlag(record === null || record === void 0 ? void 0 : record.altitudeInputEnabled) === "1";
  if (useDirectAltitude) {
    var directAltitudeM = parseDistanceToCm(record === null || record === void 0 ? void 0 : record.altitudeDirectM);
    return directAltitudeM != null ? directAltitudeM : null;
  }
  var levelHeightM = parseDistanceToCm(getRecordLevelHeight(record));
  var lowerCm = parseDistanceToCm(record === null || record === void 0 ? void 0 : record.levelLowerCm);
  if (levelHeightM == null || lowerCm == null) {
    return null;
  }
  var altitudeM = levelHeightM - lowerCm / 100;
  return Number.isFinite(altitudeM) ? altitudeM : null;
}
function formatRecordAltitudeM(record) {
  var altitudeM = getRecordAltitudeMValue(record);
  if (altitudeM == null) {
    return "";
  }
  return altitudeM.toFixed(3).replace(/\.?0+$/, "");
}
function formatLayerPosition(record) {
  var ref = value(record === null || record === void 0 ? void 0 : record.layerRef);
  var fromCm = formatCmValue(record === null || record === void 0 ? void 0 : record.layerFromCm);
  var relative = value(record === null || record === void 0 ? void 0 : record.layerRelative);
  var line1 = ref;
  var line2 = "";
  if (relative && fromCm) {
    line2 = "".concat(relative, " \u306B ").concat(fromCm);
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
  return "".concat(line1, " / ").concat(line2);
}
function formatCmValue(cmRaw) {
  var fallback = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : "";
  var text = value(cmRaw);
  if (!text) {
    return fallback;
  }
  if (/^[-ー－]+$/.test(text)) {
    return text;
  }
  var normalized = text.replace(/\s*(cm|㎝)$/i, "");
  if (!normalized) {
    return fallback;
  }
  return "".concat(normalized, "cm");
}
function clonePhotos(photos) {
  return normalizePhotos(photos).map(function (photo) {
    return _objectSpread({}, photo);
  });
}
function normalizeAsciiWidth(inputText) {
  return String(inputText).replace(/\u3000/g, " ").replace(/[！-～]/g, function (_char3) {
    return String.fromCharCode(_char3.charCodeAt(0) - 0xfee0);
  });
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
  var _window$crypto;
  if ((_window$crypto = window.crypto) !== null && _window$crypto !== void 0 && _window$crypto.randomUUID) {
    return "".concat(prefix, "-").concat(window.crypto.randomUUID());
  }
  return "".concat(prefix, "-").concat(Date.now(), "-").concat(Math.random().toString(36).slice(2, 9));
}
function showToast(message) {
  toastEl.textContent = message;
  toastEl.classList.add("show");
  if (toastTimer) {
    window.clearTimeout(toastTimer);
  }
  toastTimer = window.setTimeout(function () {
    toastEl.classList.remove("show");
  }, 2200);
}
function downloadFile(fileName, text, mimeType) {
  var blob = new Blob([text], {
    type: mimeType
  });
  var url = URL.createObjectURL(blob);
  var anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}
function escapeHtml(input) {
  return String(input || "").replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;").replaceAll("'", "&#39;");
}
function timestamp() {
  var now = new Date();
  var yyyy = String(now.getFullYear());
  var mm = String(now.getMonth() + 1).padStart(2, "0");
  var dd = String(now.getDate()).padStart(2, "0");
  var hh = String(now.getHours()).padStart(2, "0");
  var mi = String(now.getMinutes()).padStart(2, "0");
  return "".concat(yyyy).concat(mm).concat(dd, "-").concat(hh).concat(mi);
}
function isLikelyQuotaExceededError(error) {
  if (!error) {
    return false;
  }
  var name = value(error.name).toLowerCase();
  var message = value(error.message).toLowerCase();
  return name.includes("quota") || name.includes("ns_error_dom_quota_reached") || message.includes("quota") || message.includes("storage");
}
function recoverFromQuotaError(_x1) {
  return _recoverFromQuotaError.apply(this, arguments);
}
function _recoverFromQuotaError() {
  _recoverFromQuotaError = _asyncToGenerator(_regenerator().m(function _callee16(successMessage) {
    var causeError,
      quotaExceeded,
      _iterator8,
      _step0,
      step,
      changed,
      pushed,
      _args16 = arguments,
      _t11,
      _t12;
    return _regenerator().w(function (_context16) {
      while (1) switch (_context16.p = _context16.n) {
        case 0:
          causeError = _args16.length > 1 && _args16[1] !== undefined ? _args16[1] : null;
          if (!quotaRecoveryInProgress) {
            _context16.n = 1;
            break;
          }
          showToast("写真データを圧縮中です。数秒待ってから再試行してください");
          return _context16.a(2);
        case 1:
          quotaExceeded = isLikelyQuotaExceededError(causeError);
          quotaRecoveryInProgress = true;
          showToast("保存容量を超えたため写真を圧縮しています…");
          _context16.p = 2;
          _iterator8 = _createForOfIteratorHelper(PHOTO_COMPRESSION_STEPS);
          _context16.p = 3;
          _iterator8.s();
        case 4:
          if ((_step0 = _iterator8.n()).done) {
            _context16.n = 9;
            break;
          }
          step = _step0.value;
          _context16.n = 5;
          return recompressAllPhotos(step.maxLength, step.quality);
        case 5:
          changed = _context16.v;
          if (changed) {
            _context16.n = 6;
            break;
          }
          return _context16.a(3, 8);
        case 6:
          _context16.p = 6;
          localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
          renderSectionDiagramList();
          renderPhotoList();
          renderOutputs();
          if (successMessage) {
            showToast("".concat(successMessage, "\uFF08\u5199\u771F\u3092\u5727\u7E2E\u3057\u3066\u4FDD\u5B58\uFF09"));
          } else {
            showToast("写真を圧縮して保存しました");
          }
          scheduleCloudSave();
          return _context16.a(2);
        case 7:
          _context16.p = 7;
          _t11 = _context16.v;
        case 8:
          _context16.n = 4;
          break;
        case 9:
          _context16.n = 11;
          break;
        case 10:
          _context16.p = 10;
          _t12 = _context16.v;
          _iterator8.e(_t12);
        case 11:
          _context16.p = 11;
          _iterator8.f();
          return _context16.f(11);
        case 12:
          if (!cloudEndpoint) {
            _context16.n = 14;
            break;
          }
          _context16.n = 13;
          return pushStateToCloud({
            showToastOnSuccess: false,
            silentOnError: true
          });
        case 13:
          pushed = _context16.v;
          if (!pushed) {
            _context16.n = 14;
            break;
          }
          if (successMessage) {
            showToast("".concat(successMessage, "\uFF08\u7AEF\u672B\u5BB9\u91CF\u8D85\u904E\u306E\u305F\u3081\u30AF\u30E9\u30A6\u30C9\u4FDD\u5B58\uFF09"));
          } else {
            showToast("端末容量超過のためクラウド保存しました");
          }
          return _context16.a(2);
        case 14:
          if (quotaExceeded) {
            showToast("端末保存の容量を超えました。画像を減らすか圧縮して再試行してください");
          } else {
            showToast("保存に失敗しました。写真を一部削除して再試行してください");
          }
        case 15:
          _context16.p = 15;
          quotaRecoveryInProgress = false;
          return _context16.f(15);
        case 16:
          return _context16.a(2);
      }
    }, _callee16, null, [[6, 7], [3, 10, 11, 12], [2,, 15, 16]]);
  }));
  return _recoverFromQuotaError.apply(this, arguments);
}
function recompressAllPhotos(_x10, _x11) {
  return _recompressAllPhotos.apply(this, arguments);
}
function _recompressAllPhotos() {
  _recompressAllPhotos = _asyncToGenerator(_regenerator().m(function _callee17(maxLength, quality) {
    var changed, _iterator9, _step1, record, sectionResult, photoResult, imageResult, currentSectionResult, currentPhotoResult, currentCustomImageResult, _t13;
    return _regenerator().w(function (_context17) {
      while (1) switch (_context17.p = _context17.n) {
        case 0:
          changed = false;
          _iterator9 = _createForOfIteratorHelper(state.records);
          _context17.p = 1;
          _iterator9.s();
        case 2:
          if ((_step1 = _iterator9.n()).done) {
            _context17.n = 7;
            break;
          }
          record = _step1.value;
          _context17.n = 3;
          return recompressPhotoArray(record.sectionDiagrams, maxLength, quality);
        case 3:
          sectionResult = _context17.v;
          if (sectionResult.changed) {
            record.sectionDiagrams = sectionResult.photos;
            changed = true;
          }
          _context17.n = 4;
          return recompressPhotoArray(record.photos, maxLength, quality);
        case 4:
          photoResult = _context17.v;
          if (photoResult.changed) {
            record.photos = photoResult.photos;
            changed = true;
          }
          if (!isCustomLargeShapeType(record.largeShapeType)) {
            _context17.n = 6;
            break;
          }
          _context17.n = 5;
          return recompressImageDataUrl(record.customLargeImageDataUrl, maxLength, quality);
        case 5:
          imageResult = _context17.v;
          if (imageResult.changed) {
            record.customLargeImageDataUrl = imageResult.dataUrl;
            changed = true;
          }
        case 6:
          _context17.n = 2;
          break;
        case 7:
          _context17.n = 9;
          break;
        case 8:
          _context17.p = 8;
          _t13 = _context17.v;
          _iterator9.e(_t13);
        case 9:
          _context17.p = 9;
          _iterator9.f();
          return _context17.f(9);
        case 10:
          _context17.n = 11;
          return recompressPhotoArray(currentSectionDiagrams, maxLength, quality);
        case 11:
          currentSectionResult = _context17.v;
          if (currentSectionResult.changed) {
            currentSectionDiagrams = currentSectionResult.photos;
            changed = true;
          }
          _context17.n = 12;
          return recompressPhotoArray(currentPhotos, maxLength, quality);
        case 12:
          currentPhotoResult = _context17.v;
          if (currentPhotoResult.changed) {
            currentPhotos = currentPhotoResult.photos;
            changed = true;
          }
          if (!customLargeImageDataUrlInput) {
            _context17.n = 15;
            break;
          }
          _context17.n = 13;
          return recompressImageDataUrl(customLargeImageDataUrlInput.value, maxLength, quality);
        case 13:
          currentCustomImageResult = _context17.v;
          if (!currentCustomImageResult.changed) {
            _context17.n = 15;
            break;
          }
          customLargeImageDataUrlInput.value = currentCustomImageResult.dataUrl;
          _context17.n = 14;
          return updateCustomLargeImageAspectFromDataUrl(currentCustomImageResult.dataUrl);
        case 14:
          syncLargeShapeImagePreviewForCurrentForm();
          changed = true;
        case 15:
          return _context17.a(2, changed);
      }
    }, _callee17, null, [[1, 8, 9, 10]]);
  }));
  return _recompressAllPhotos.apply(this, arguments);
}
function recompressPhotoArray(_x12, _x13, _x14) {
  return _recompressPhotoArray.apply(this, arguments);
}
function _recompressPhotoArray() {
  _recompressPhotoArray = _asyncToGenerator(_regenerator().m(function _callee18(photosRaw, maxLength, quality) {
    var nextPhotos, changed, _iterator0, _step10, photo, dataUrl, nextDataUrl, recompressed, _t14, _t15;
    return _regenerator().w(function (_context18) {
      while (1) switch (_context18.p = _context18.n) {
        case 0:
          if (!(!Array.isArray(photosRaw) || !photosRaw.length)) {
            _context18.n = 1;
            break;
          }
          return _context18.a(2, {
            photos: Array.isArray(photosRaw) ? photosRaw : [],
            changed: false
          });
        case 1:
          nextPhotos = [];
          changed = false;
          _iterator0 = _createForOfIteratorHelper(photosRaw);
          _context18.p = 2;
          _iterator0.s();
        case 3:
          if ((_step10 = _iterator0.n()).done) {
            _context18.n = 11;
            break;
          }
          photo = _step10.value;
          if (!(!photo || _typeof(photo) !== "object")) {
            _context18.n = 4;
            break;
          }
          return _context18.a(3, 10);
        case 4:
          dataUrl = value(photo.dataUrl);
          if (dataUrl) {
            _context18.n = 5;
            break;
          }
          return _context18.a(3, 10);
        case 5:
          nextDataUrl = dataUrl;
          _context18.p = 6;
          _context18.n = 7;
          return resizeDataUrlImage(dataUrl, maxLength, quality);
        case 7:
          recompressed = _context18.v;
          if (recompressed && recompressed.length < dataUrl.length) {
            nextDataUrl = recompressed;
            changed = true;
          }
          _context18.n = 9;
          break;
        case 8:
          _context18.p = 8;
          _t14 = _context18.v;
        case 9:
          nextPhotos.push(_objectSpread(_objectSpread({}, photo), {}, {
            dataUrl: nextDataUrl
          }));
        case 10:
          _context18.n = 3;
          break;
        case 11:
          _context18.n = 13;
          break;
        case 12:
          _context18.p = 12;
          _t15 = _context18.v;
          _iterator0.e(_t15);
        case 13:
          _context18.p = 13;
          _iterator0.f();
          return _context18.f(13);
        case 14:
          return _context18.a(2, {
            photos: nextPhotos,
            changed: changed
          });
      }
    }, _callee18, null, [[6, 8], [2, 12, 13, 14]]);
  }));
  return _recompressPhotoArray.apply(this, arguments);
}
function recompressImageDataUrl(_x15, _x16, _x17) {
  return _recompressImageDataUrl.apply(this, arguments);
}
function _recompressImageDataUrl() {
  _recompressImageDataUrl = _asyncToGenerator(_regenerator().m(function _callee19(dataUrlRaw, maxLength, quality) {
    var dataUrl, matchedMimeType, sourceMimeType, outputMimeType, recompressed, _t16;
    return _regenerator().w(function (_context19) {
      while (1) switch (_context19.p = _context19.n) {
        case 0:
          dataUrl = normalizeCustomLargeImageDataUrl(dataUrlRaw);
          if (dataUrl) {
            _context19.n = 1;
            break;
          }
          return _context19.a(2, {
            dataUrl: dataUrl,
            changed: false
          });
        case 1:
          matchedMimeType = dataUrl.match(/^data:([^;,]+)/i);
          sourceMimeType = normalizeImageMimeType(matchedMimeType === null || matchedMimeType === void 0 ? void 0 : matchedMimeType[1]);
          outputMimeType = sourceMimeType === "image/png" ? "image/png" : "image/jpeg";
          _context19.p = 2;
          _context19.n = 3;
          return resizeDataUrlImage(dataUrl, maxLength, quality, outputMimeType);
        case 3:
          recompressed = _context19.v;
          if (!(recompressed && recompressed.length < dataUrl.length)) {
            _context19.n = 4;
            break;
          }
          return _context19.a(2, {
            dataUrl: recompressed,
            changed: true
          });
        case 4:
          _context19.n = 6;
          break;
        case 5:
          _context19.p = 5;
          _t16 = _context19.v;
        case 6:
          return _context19.a(2, {
            dataUrl: dataUrl,
            changed: false
          });
      }
    }, _callee19, null, [[2, 5]]);
  }));
  return _recompressImageDataUrl.apply(this, arguments);
}
var SUPPORTED_UPLOAD_IMAGE_MIME_TYPES = new Set(["image/jpeg", "image/png"]);
function normalizeImageMimeType(mimeTypeRaw) {
  var mimeType = value(mimeTypeRaw).toLowerCase();
  if (!mimeType) {
    return "";
  }
  if (mimeType === "image/jpg") {
    return "image/jpeg";
  }
  return mimeType;
}
function inferImageMimeTypeFromFileName(fileNameRaw) {
  var fileName = value(fileNameRaw).toLowerCase();
  if (fileName.endsWith(".jpg") || fileName.endsWith(".jpeg")) {
    return "image/jpeg";
  }
  if (fileName.endsWith(".png")) {
    return "image/png";
  }
  return "";
}
function resolveImageMimeType(file) {
  var preferredMimeTypeRaw = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : "";
  var preferredMimeType = normalizeImageMimeType(preferredMimeTypeRaw);
  if (SUPPORTED_UPLOAD_IMAGE_MIME_TYPES.has(preferredMimeType)) {
    return preferredMimeType;
  }
  var fileMimeType = normalizeImageMimeType(file === null || file === void 0 ? void 0 : file.type);
  if (SUPPORTED_UPLOAD_IMAGE_MIME_TYPES.has(fileMimeType)) {
    return fileMimeType;
  }
  var fromName = inferImageMimeTypeFromFileName(file === null || file === void 0 ? void 0 : file.name);
  if (SUPPORTED_UPLOAD_IMAGE_MIME_TYPES.has(fromName)) {
    return fromName;
  }
  return "";
}
function normalizeImageDataUrlForUpload(dataUrlRaw, file) {
  var preferredMimeTypeRaw = arguments.length > 2 && arguments[2] !== undefined ? arguments[2] : "";
  var dataUrl = value(dataUrlRaw);
  if (!dataUrl || !dataUrl.startsWith("data:")) {
    return "";
  }
  if (dataUrl.startsWith("data:image/")) {
    return dataUrl;
  }
  var resolvedMimeType = resolveImageMimeType(file, preferredMimeTypeRaw);
  if (!resolvedMimeType) {
    return "";
  }
  if (/^data:;base64,/i.test(dataUrl)) {
    return dataUrl.replace(/^data:;base64,/i, "data:".concat(resolvedMimeType, ";base64,"));
  }
  return dataUrl.replace(/^data:[^;,]+/i, "data:".concat(resolvedMimeType));
}
function resizeDataUrlImage(dataUrl, maxLength) {
  var quality = arguments.length > 2 && arguments[2] !== undefined ? arguments[2] : 0.72;
  var mimeType = arguments.length > 3 && arguments[3] !== undefined ? arguments[3] : "image/jpeg";
  return new Promise(function (resolve, reject) {
    var image = new Image();
    image.onload = function () {
      var width = image.width;
      var height = image.height;
      if (width > maxLength || height > maxLength) {
        var scale = Math.min(maxLength / width, maxLength / height);
        width = Math.max(1, Math.round(width * scale));
        height = Math.max(1, Math.round(height * scale));
      }
      var canvas = document.createElement("canvas");
      canvas.width = width;
      canvas.height = height;
      var context = canvas.getContext("2d");
      if (!context) {
        reject(new Error("Canvas context unavailable"));
        return;
      }
      context.drawImage(image, 0, 0, width, height);
      resolve(canvas.toDataURL(mimeType, quality));
    };
    image.onerror = function () {
      return reject(new Error("Failed to load image"));
    };
    image.src = dataUrl;
  });
}
function resizeImage(file, maxLength) {
  var quality = arguments.length > 2 && arguments[2] !== undefined ? arguments[2] : 0.72;
  var mimeType = arguments.length > 3 && arguments[3] !== undefined ? arguments[3] : "image/jpeg";
  return new Promise(function (resolve, reject) {
    var outputMimeType = resolveImageMimeType(file, mimeType) || "image/jpeg";
    var reader = new FileReader();
    reader.onload = function () {
      var normalizedDataUrl = normalizeImageDataUrlForUpload(reader.result, file, outputMimeType);
      if (!normalizedDataUrl) {
        reject(new Error("Unsupported image type"));
        return;
      }
      resizeDataUrlImage(normalizedDataUrl, maxLength, quality, outputMimeType).then(resolve)["catch"](reject);
    };
    reader.onerror = function () {
      return reject(new Error("Failed to read file"));
    };
    reader.readAsDataURL(file);
  });
}
function readFileAsDataUrl(file) {
  return new Promise(function (resolve, reject) {
    var sourceFile = file instanceof File ? file : null;
    if (!sourceFile) {
      reject(new Error("Invalid file"));
      return;
    }
    var reader = new FileReader();
    reader.onload = function () {
      resolve(String(reader.result || ""));
    };
    reader.onerror = function () {
      return reject(new Error("Failed to read file"));
    };
    reader.readAsDataURL(sourceFile);
  });
}
function loadImageFileDataUrlWithFallback(_x18) {
  return _loadImageFileDataUrlWithFallback.apply(this, arguments);
}
function _loadImageFileDataUrlWithFallback() {
  _loadImageFileDataUrlWithFallback = _asyncToGenerator(_regenerator().m(function _callee20(file) {
    var _ref19,
      _ref19$maxLength,
      maxLength,
      _ref19$quality,
      quality,
      _ref19$mimeType,
      mimeType,
      _ref19$allowOriginalF,
      allowOriginalFallback,
      sourceFile,
      outputMimeType,
      originalDataUrl,
      _args20 = arguments,
      _t17,
      _t18;
    return _regenerator().w(function (_context20) {
      while (1) switch (_context20.p = _context20.n) {
        case 0:
          _ref19 = _args20.length > 1 && _args20[1] !== undefined ? _args20[1] : {}, _ref19$maxLength = _ref19.maxLength, maxLength = _ref19$maxLength === void 0 ? 1280 : _ref19$maxLength, _ref19$quality = _ref19.quality, quality = _ref19$quality === void 0 ? 0.72 : _ref19$quality, _ref19$mimeType = _ref19.mimeType, mimeType = _ref19$mimeType === void 0 ? "image/jpeg" : _ref19$mimeType, _ref19$allowOriginalF = _ref19.allowOriginalFallback, allowOriginalFallback = _ref19$allowOriginalF === void 0 ? true : _ref19$allowOriginalF;
          sourceFile = file instanceof File ? file : null;
          if (sourceFile) {
            _context20.n = 1;
            break;
          }
          throw new Error("Invalid file");
        case 1:
          outputMimeType = resolveImageMimeType(sourceFile, mimeType) || normalizeImageMimeType(mimeType) || "image/jpeg";
          _context20.p = 2;
          _context20.n = 3;
          return resizeImage(sourceFile, maxLength, quality, outputMimeType);
        case 3:
          return _context20.a(2, _context20.v);
        case 4:
          _context20.p = 4;
          _t17 = _context20.v;
          if (allowOriginalFallback) {
            _context20.n = 5;
            break;
          }
          throw _t17;
        case 5:
          _t18 = normalizeImageDataUrlForUpload;
          _context20.n = 6;
          return readFileAsDataUrl(sourceFile);
        case 6:
          originalDataUrl = _t18(_context20.v, sourceFile, outputMimeType);
          if (originalDataUrl) {
            _context20.n = 7;
            break;
          }
          throw _t17;
        case 7:
          return _context20.a(2, originalDataUrl);
      }
    }, _callee20, null, [[2, 4]]);
  }));
  return _loadImageFileDataUrlWithFallback.apply(this, arguments);
}
