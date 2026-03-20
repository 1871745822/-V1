/* global XLSX, echarts */

const $ = (id) => document.getElementById(id);

const state = {
  files: [],
  rows: [],
  cols: [],
  mapping: null,
  periodCache: new Map(), // yyyymm -> { rows, cols, mapping, fileName }
  periodMeta: {
    // parsed from file names like YYYYMM
    yyyymm: new Set(),
    years: [],
    maxYear: new Date().getFullYear(),
  },
  uiDefaults: {
    year: null,
    month: null,
    monthMode: "month",
  },
  tableSort: {
    key: null, // bill|std|rate|cost|rev|costRate|otDaily
    dir: null, // asc|desc|null
  },
  // 明细树表点击联动：{ scope:'dept1'|'dept2'|'emp', dept1, dept2, emp, selKey }
  tableSelection: null,
  charts: {
    billTrend: null,
    otTrend: null,
    yoymom: null,
  },
  latestAlerts: [],
  alertView: {
    page: 1,
    pageSize: 5,
  },
  titleState: null,
};

const TITLE_STORAGE_KEY = "dashboardTitles.v1";
const THEME_STORAGE_KEY = "dashboardTheme.v1";
const COLLAPSE_STORAGE_KEY = "dashboardCollapse.v1";

function loadCollapseState() {
  try {
    const raw = localStorage.getItem(COLLAPSE_STORAGE_KEY);
    return raw ? JSON.parse(raw) : {};
  } catch {
    return {};
  }
}

function saveCollapseState(obj) {
  try {
    localStorage.setItem(COLLAPSE_STORAGE_KEY, JSON.stringify(obj || {}));
  } catch {
    // ignore
  }
}

function applyCollapseState(collapseState) {
  const state = collapseState || {};
  document.querySelectorAll(".collapse-toggle").forEach((btn) => {
    const key = btn?.dataset?.collapseKey;
    if (!key) return;
    const card = btn.closest(".card");
    if (!card) return;
    const collapsed = !!state[key];
    card.classList.toggle("collapsed", collapsed);
    btn.textContent = collapsed ? "▸" : "▾";
  });
}

function applyTheme(theme) {
  const v = theme === "light" ? "light" : "dark";
  document.documentElement.dataset.theme = v;
  try {
    localStorage.setItem(THEME_STORAGE_KEY, v);
  } catch {
    // ignore storage errors
  }

  const btnDark = $("btnThemeDark");
  const btnLight = $("btnThemeLight");
  if (btnDark) btnDark.classList.toggle("active", v === "dark");
  if (btnLight) btnLight.classList.toggle("active", v === "light");
}

function getChartTheme() {
  const cs = getComputedStyle(document.documentElement);
  return {
    axis: cs.getPropertyValue("--chartAxisColor").trim() || "rgba(255,255,255,0.65)",
    split: cs.getPropertyValue("--chartSplitLineColor").trim() || "rgba(255,255,255,0.08)",
    legend: cs.getPropertyValue("--chartLegendColor").trim() || "rgba(255,255,255,0.75)",
  };
}

const SHARE_STORAGE_VERSION = 1;
const SHARE_PARAM_KEY = "share";

function utf8ToBytes(s) {
  return new TextEncoder().encode(String(s));
}

function bytesToUtf8(bytes) {
  return new TextDecoder().decode(bytes);
}

function bytesToBase64(bytes) {
  // bytes: Uint8Array
  let binary = "";
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
  }
  // base64url: URL safe and no '=' padding (fits better in short links)
  const b64 = btoa(binary);
  return b64.replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
}

function base64ToBytes(b64) {
  // convert base64url back to normal base64
  const s = String(b64).replace(/-/g, "+").replace(/_/g, "/");
  const pad = s.length % 4 ? "=".repeat(4 - (s.length % 4)) : "";
  const bin = atob(s + pad);
  const len = bin.length;
  const bytes = new Uint8Array(len);
  for (let i = 0; i < len; i++) bytes[i] = bin.charCodeAt(i);
  return bytes;
}

async function encodeShareData(obj) {
  const json = JSON.stringify(obj);
  const plainBytes = utf8ToBytes(json);

  // 优先 gzip 压缩以减小链接长度
  if (typeof CompressionStream !== "undefined") {
    const blob = new Blob([plainBytes], { type: "application/json" });
    const compressedStream = blob.stream().pipeThrough(new CompressionStream("gzip"));
    const arrBuf = await new Response(compressedStream).arrayBuffer();
    const bytes = new Uint8Array(arrBuf);
    return `gz.${bytesToBase64(bytes)}`;
  }

  return `plain.${bytesToBase64(plainBytes)}`;
}

async function decodeShareData(encoded) {
  if (!encoded) return null;
  const str = String(encoded);
  const idx = str.indexOf(".");
  if (idx < 0) return null;
  const kind = str.slice(0, idx);
  const payload = str.slice(idx + 1);

  if (kind === "gz") {
    if (typeof DecompressionStream === "undefined") return null;
    const bytes = base64ToBytes(payload);
    const blob = new Blob([bytes], { type: "application/octet-stream" });
    const decompressedStream = blob.stream().pipeThrough(new DecompressionStream("gzip"));
    const text = await new Response(decompressedStream).text();
    return JSON.parse(text);
  }

  if (kind === "plain") {
    const bytes = base64ToBytes(payload);
    const text = bytesToUtf8(bytes);
    return JSON.parse(text);
  }

  return null;
}

function getShareEncodedFromUrl() {
  const hash = location.hash || "";
  const m = hash.match(new RegExp(`${SHARE_PARAM_KEY}=([^&]+)`));
  return m ? decodeURIComponent(m[1]) : null;
}

function encodeCompactPeriodRows(rows, m, dailyOt) {
  // dictionary: de-duplicate repeated dimension strings to shorten payload
  const dict = [];
  const dictIdx = new Map();
  const idOf = (s) => {
    const k = String(s ?? "");
    if (dictIdx.has(k)) return dictIdx.get(k);
    const id = dict.length;
    dict.push(k);
    dictIdx.set(k, id);
    return id;
  };

  const group = new Map(); // key => agg
  for (const r of rows || []) {
    const d1 = (r?.[m.dept1Col] ?? "").toString().trim();
    const d2 = (r?.[m.dept2Col] ?? "").toString().trim();
    const emp = (r?.[m.empCol] ?? "").toString().trim();
    const key = `${d1}|||${d2}|||${emp}`;
    const cur =
      group.get(key) ?? { d1, d2, emp, bill: 0, std: 0, ot: 0, cost: 0, rev: 0, cnt: 0, daySet: new Set() };
    cur.bill += toNumber(r?.[m.billCol]) ?? 0;
    cur.std += toNumber(r?.[m.stdCol]) ?? 0;
    cur.ot += toNumber(r?.[m.otCol]) ?? 0;
    cur.cost += toNumber(r?.[m.costCol]) ?? 0;
    cur.rev += toNumber(r?.[m.revCol]) ?? 0;
    cur.cnt += 1;
    if (!dailyOt) {
      const d = toDate(r?.[m.dateCol]);
      if (d) cur.daySet.add(d.toISOString().slice(0, 10));
    }
    group.set(key, cur);
  }

  const round2 = (x) => {
    const n = Number(x);
    if (!Number.isFinite(n)) return 0;
    return Math.round(n * 100) / 100;
  };

  // row tuple: [d1Id,d2Id,empId,bill,std,ot,cost,rev,cnt,dayCount]
  const packedRows = [];
  for (const cur of group.values()) {
    packedRows.push([
      idOf(cur.d1),
      idOf(cur.d2),
      idOf(cur.emp),
      round2(cur.bill),
      round2(cur.std),
      round2(cur.ot),
      round2(cur.cost),
      round2(cur.rev),
      cur.cnt,
      dailyOt ? 0 : cur.daySet.size,
    ]);
  }

  return { d: dict, r: packedRows };
}

function decodeCompactPeriodRows(compact) {
  const dict = Array.isArray(compact?.d) ? compact.d : [];
  const rows = Array.isArray(compact?.r) ? compact.r : [];
  const safe = (id) => (Number.isInteger(id) && id >= 0 && id < dict.length ? dict[id] : "");
  return rows.map((x) => ({
    d1: safe(x?.[0]),
    d2: safe(x?.[1]),
    emp: safe(x?.[2]),
    bill: Number(x?.[3] ?? 0) || 0,
    std: Number(x?.[4] ?? 0) || 0,
    ot: Number(x?.[5] ?? 0) || 0,
    cost: Number(x?.[6] ?? 0) || 0,
    rev: Number(x?.[7] ?? 0) || 0,
    __shareCnt: Number(x?.[8] ?? 0) || 0,
    __shareDayCount: Number(x?.[9] ?? 0) || 0,
  }));
}

/** 全局字典：所有月份共用一份字符串表，比每月份各自 d/r 小得多 */
function encodeGlobalCompactSnapshot(periodKeys, m0, dailyOt) {
  const D = [];
  const dictIdx = new Map();
  const idOf = (s) => {
    const k = String(s ?? "");
    if (dictIdx.has(k)) return dictIdx.get(k);
    const id = D.length;
    D.push(k);
    dictIdx.set(k, id);
    return id;
  };

  const round2 = (x) => {
    const n = Number(x);
    if (!Number.isFinite(n)) return 0;
    return Math.round(n * 100) / 100;
  };

  const P = [];
  for (const p of periodKeys) {
    const rows = getPeriodData(p).rows || [];
    const group = new Map();
    for (const r of rows) {
      const d1 = (r?.[m0.dept1Col] ?? "").toString().trim();
      const d2 = (r?.[m0.dept2Col] ?? "").toString().trim();
      const emp = (r?.[m0.empCol] ?? "").toString().trim();
      const key = `${d1}|||${d2}|||${emp}`;
      const cur =
        group.get(key) ?? { d1, d2, emp, bill: 0, std: 0, ot: 0, cost: 0, rev: 0, cnt: 0, daySet: new Set() };
      cur.bill += toNumber(r?.[m0.billCol]) ?? 0;
      cur.std += toNumber(r?.[m0.stdCol]) ?? 0;
      cur.ot += toNumber(r?.[m0.otCol]) ?? 0;
      cur.cost += toNumber(r?.[m0.costCol]) ?? 0;
      cur.rev += toNumber(r?.[m0.revCol]) ?? 0;
      cur.cnt += 1;
      if (!dailyOt) {
        const d = toDate(r?.[m0.dateCol]);
        if (d) cur.daySet.add(d.toISOString().slice(0, 10));
      }
      group.set(key, cur);
    }
    const packed = [];
    for (const cur of group.values()) {
      packed.push([
        idOf(cur.d1),
        idOf(cur.d2),
        idOf(cur.emp),
        round2(cur.bill),
        round2(cur.std),
        round2(cur.ot),
        round2(cur.cost),
        round2(cur.rev),
        cur.cnt,
        dailyOt ? 0 : cur.daySet.size,
      ]);
    }
    P.push([p, packed]);
  }

  return { g: D, p: P };
}

function decodeGlobalCompactSnapshot(snap) {
  const G = Array.isArray(snap?.g) ? snap.g : [];
  const safe = (i) =>
    Number.isFinite(i) && i >= 0 && i < G.length ? G[Math.floor(Number(i))] : "";
  const P = Array.isArray(snap?.p) ? snap.p : [];
  return P.map(([period, tuples]) => [
    period,
    {
      rows: (tuples || []).map((t) => ({
        d1: safe(t?.[0]),
        d2: safe(t?.[1]),
        emp: safe(t?.[2]),
        bill: Number(t?.[3] ?? 0) || 0,
        std: Number(t?.[4] ?? 0) || 0,
        ot: Number(t?.[5] ?? 0) || 0,
        cost: Number(t?.[6] ?? 0) || 0,
        rev: Number(t?.[7] ?? 0) || 0,
        __shareCnt: Number(t?.[8] ?? 0) || 0,
        __shareDayCount: Number(t?.[9] ?? 0) || 0,
      })),
    },
  ]);
}

function showShareQrModal(url) {
  const modal = $("shareQrModal");
  if (!modal) return;
  const box = $("shareQrBox");
  const input = $("shareQrUrlInput");
  if (input) input.value = url || "";

  // QRCode 由 qrcodejs 提供
  if (typeof window.QRCode !== "function") {
    if (box) box.textContent = "未加载二维码库，无法生成二维码。";
    modal.style.display = "flex";
    return;
  }

  if (box) box.innerHTML = "";
  try {
    // qrcodejs: new QRCode(domEl, options)
    if (box) {
      // eslint-disable-next-line no-new
      new QRCode(box, {
        text: url || "",
        width: 256,
        height: 256,
        correctLevel: QRCode.CorrectLevel.H,
      });
    }
  } catch {
    if (box) box.textContent = "二维码生成失败。";
  }

  modal.style.display = "flex";
}

const defaultTitles = {
  dashboardTitle: "各团队数据看板",
  summaryTitle: "部门 / 人员明细（可展开）",
  billTrendTitle: "账单小时趋势",
  otTrendTitle: "加班小时趋势",
  yoyTitle: "同比 / 环比（完成率 & 加班小时）",
  alertTitle: "数据异常预警",
  kpiBillLabel: "账单小时",
  kpiStdLabel: "额定账单小时",
  kpiRateLabel: "小时完成率",
  kpiOtDailyLabel: "日均加班小时",
};

function fmtNumber(x) {
  if (x == null || Number.isNaN(x)) return "-";
  const n = Number(x);
  if (!Number.isFinite(n)) return "-";
  return n.toLocaleString("zh-CN", { maximumFractionDigits: 2 });
}

function toDate(v) {
  if (v == null || v === "") return null;
  if (v instanceof Date && !Number.isNaN(v.getTime())) return v;
  if (typeof v === "number") {
    // Excel serial date
    const d = XLSX.SSF.parse_date_code(v);
    if (!d) return null;
    return new Date(Date.UTC(d.y, d.m - 1, d.d));
  }
  if (typeof v === "string") {
    const s = v.trim();
    if (!s) return null;
    // normalize common formats: 2026/03/17, 2026-03-17, 2026.03.17, 20260317
    const s2 = s.replace(/[.\/]/g, "-");
    const iso = /^\d{4}-\d{1,2}-\d{1,2}/.test(s2) ? s2 : null;
    if (iso) {
      const d = new Date(iso);
      if (!Number.isNaN(d.getTime())) return d;
    }
    if (/^\d{8}$/.test(s)) {
      const d = new Date(`${s.slice(0, 4)}-${s.slice(4, 6)}-${s.slice(6, 8)}`);
      if (!Number.isNaN(d.getTime())) return d;
    }
    const d = new Date(s);
    if (!Number.isNaN(d.getTime())) return d;
  }
  return null;
}

function toNumber(v) {
  if (v == null || v === "") return null;
  if (typeof v === "number") return Number.isFinite(v) ? v : null;
  if (typeof v === "string") {
    const s = v.trim().replace(/,/g, "");
    if (!s) return null;
    const n = Number(s);
    return Number.isFinite(n) ? n : null;
  }
  return null;
}

function uniq(arr) {
  return Array.from(new Set(arr));
}

function normColName(s) {
  return String(s ?? "")
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/[()（）\[\]【】\-_]/g, "")
    .trim();
}

function toYYYYMM(year, month) {
  const y = String(year);
  const m = String(month).padStart(2, "0");
  return `${y}${m}`;
}

function parseNameYYYYMM(name) {
  const s = String(name).trim();
  if (!/^\d{6}$/.test(s)) return null;
  const year = Number(s.slice(0, 4));
  const month = Number(s.slice(4, 6));
  if (!Number.isFinite(year) || !Number.isFinite(month)) return null;
  if (year < 1900 || month < 1 || month > 12) return null;
  return { year, month, yyyymm: s };
}

function parseFileNameYYYYMM(fileName) {
  const s = String(fileName || "");
  // find first 6-digit token like 202602 in filename
  const m = s.match(/(\d{6})/);
  if (!m) return null;
  return parseNameYYYYMM(m[1]);
}

function computePeriodMeta(periodKeys) {
  const yyyymm = new Set();
  const years = [];
  const monthsByYear = new Map(); // year -> Set(month)
  const maxMonthByYear = new Map(); // year -> maxMonth
  for (const k of periodKeys) {
    const p = parseNameYYYYMM(k);
    if (!p) continue;
    yyyymm.add(p.yyyymm);
    years.push(p.year);
    if (!monthsByYear.has(p.year)) monthsByYear.set(p.year, new Set());
    monthsByYear.get(p.year).add(p.month);
    const curMax = maxMonthByYear.get(p.year) ?? 0;
    if (p.month > curMax) maxMonthByYear.set(p.year, p.month);
  }
  years.sort((a, b) => a - b);
  const maxYear = years.length ? years[years.length - 1] : new Date().getFullYear();
  return { yyyymm, years: uniq(years), maxYear, monthsByYear, maxMonthByYear };
}

function parseLeadingNumber(s) {
  const m = String(s ?? "").trim().match(/^(\d{1,4})/);
  return m ? Number(m[1]) : null;
}

function sortByDeptCode(a, b) {
  const na = parseLeadingNumber(a);
  const nb = parseLeadingNumber(b);
  if (na != null && nb != null && na !== nb) return na - nb;
  if (na != null && nb == null) return -1;
  if (na == null && nb != null) return 1;
  return String(a).localeCompare(String(b), "zh-CN");
}

function guessMapping(cols) {
  const pick = (candidates) =>
    cols.find((c) => {
      const nc = normColName(c);
      return candidates.some((k) => nc.includes(normColName(k)));
    });
  const dateCol =
    pick(["所属日期", "日期", "date", "time", "时间", "月份", "month", "期间"]) || null;
  const dept1Col = pick(["一级部门", "部门1", "一级", "事业部", "bg"]) || null;
  const dept2Col = pick(["二级部门", "部门2", "二级", "组", "team"]) || null;
  const empCol = pick(["员工", "姓名", "人员", "员工姓名", "人员姓名", "name"]) || null;
  const billCol = pick(["账单小时", "计费小时", "bill", "billing"]) || null;
  const stdCol = pick(["额定账单小时", "标准账单小时", "额定", "标准工时", "std"]) || null;
  // 加班字段在不同表里命名差异很大，这里尽量覆盖常见写法
  const otCol =
    pick([
      "加班小时",
      "加班时长",
      "加班工时",
      "加班小时数",
      "ot小时",
      "othours",
      "overtimehours",
      "overtime",
      "加班",
      "ot",
      "日均加班小时",
      "日均加班时长",
    ]) || null;
  const costCol = pick(["人力成本", "成本", "labor cost"]) || null;
  const revCol = pick(["收入", "营收", "revenue"]) || null;
  return { dateCol, dept1Col, dept2Col, empCol, billCol, stdCol, otCol, costCol, revCol };
}

function setMappingHint(m) {
  const el = $("mappingHint");
  if (!el) return;
  const fmt = (k, v) => `${k}=${v || "（未识别）"}`;
  el.textContent =
    "字段映射：" +
    [
      fmt("日期", m?.dateCol),
      fmt("一级部门", m?.dept1Col),
      fmt("二级部门", m?.dept2Col),
      fmt("员工", m?.empCol),
      fmt("账单小时", m?.billCol),
      fmt("额定账单小时", m?.stdCol),
      fmt("加班小时", m?.otCol),
      fmt("人力成本", m?.costCol),
      fmt("收入", m?.revCol),
    ].join("，");
}

function inferDateColFromValues(rows, cols, targetYear, targetMonth) {
  if (!Array.isArray(rows) || !rows.length) return null;
  if (!Array.isArray(cols) || !cols.length) return null;

  let bestCol = null;
  let bestScore = -1;
  const maxSampleRows = Math.min(rows.length, 200);

  for (const c of cols) {
    // 排除明显不是日期的列（例如：所有“小时/成本/收入”类列）
    const cn = normColName(c);
    const excluded =
      cn.includes("小时") ||
      cn.includes("工时") ||
      cn.includes("成本") ||
      cn.includes("收入") ||
      cn.includes("加班") ||
      cn.includes("额定") ||
      cn.includes("标准工时") ||
      cn.includes("标准");
    if (excluded) continue;

    let ok = 0; // toDate(v) 成功数
    let match = 0; // toDate(v) 且命中目标 yyyymm 的数
    let seen = 0;
    for (let i = 0; i < maxSampleRows; i++) {
      const v = rows[i]?.[c];
      if (v == null || v === "") continue;
      seen++;
      const d = toDate(v);
      if (d) {
        ok++;
        if (
          Number.isFinite(targetYear) &&
          Number.isFinite(targetMonth) &&
          d.getUTCFullYear() === targetYear &&
          d.getUTCMonth() + 1 === targetMonth
        ) {
          match++;
        }
      }
    }

    if (seen < 5) continue;

    const okRatio = ok / seen;
    const matchRatio = targetYear && targetMonth ? match / Math.max(1, seen) : 0;
    // 目标约束更重要：先看 match 命中，其次才看 ok 解析比例
    const score = (targetYear && targetMonth ? matchRatio * 30 : 0) + okRatio * 10 + (seen >= 50 ? 2 : 0);

    if (score > bestScore && ok >= 5 && (!targetYear || !targetMonth || match >= 3)) {
      bestScore = score;
      bestCol = c;
    }
  }

  return bestCol;
}

function colMatchesTargetYYYYMM(rows, col, targetYear, targetMonth) {
  if (!Array.isArray(rows) || !rows.length) return { match: 0, seen: 0, ok: 0, ratio: 0 };
  if (!col) return { match: 0, seen: 0, ok: 0, ratio: 0 };
  const maxSampleRows = Math.min(rows.length, 200);
  let ok = 0;
  let match = 0;
  let seen = 0;
  for (let i = 0; i < maxSampleRows; i++) {
    const v = rows[i]?.[col];
    if (v == null || v === "") continue;
    seen++;
    const d = toDate(v);
    if (!d) continue;
    ok++;
    if (
      Number.isFinite(targetYear) &&
      Number.isFinite(targetMonth) &&
      d.getUTCFullYear() === targetYear &&
      d.getUTCMonth() + 1 === targetMonth
    ) {
      match++;
    }
  }
  const ratio = seen ? match / seen : 0;
  return { match, seen, ok, ratio };
}

function fillSelect(selectEl, options, selected) {
  selectEl.innerHTML = "";
  for (const opt of options) {
    const o = document.createElement("option");
    o.value = opt;
    o.textContent = opt;
    selectEl.appendChild(o);
  }
  if (selected && options.includes(selected)) selectEl.value = selected;
}

function formatKeyByGrain(date, grain) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  if (grain === "year") return `${y}`;
  if (grain === "month") return `${y}-${m}`;
  return `${y}-${m}-${d}`;
}

function getSelectedMulti(selectEl) {
  return Array.from(selectEl.selectedOptions).map((o) => o.value);
}

function refreshDept2Options() {
  if (!state.mapping || !state.rows.length) return;
  const m = state.mapping;
  const d1Sel = $("dept1Select");
  const d2Sel = $("dept2Select");
  if (!d1Sel || !d2Sel) return;
  const selectedD1 = getSelectedMulti(d1Sel);
  const d1Set = selectedD1.length ? new Set(selectedD1) : null;
  const candidates = state.rows.filter((r) => {
    const d1 = (r?.[m.dept1Col] ?? "").toString().trim();
    if (d1Set && !d1Set.has(d1)) return false;
    return true;
  });
  const d2s = uniq(
    candidates
      .map((r) => (r?.[m.dept2Col] ?? "").toString().trim())
      .filter(Boolean)
  ).sort(sortByDeptCode);
  const prevSelected = new Set(getSelectedMulti(d2Sel));
  d2Sel.innerHTML = "";
  for (const v of d2s) {
    const o = document.createElement("option");
    o.value = v;
    o.textContent = v;
    if (prevSelected.has(v)) o.selected = true;
    d2Sel.appendChild(o);
  }
}

function setNoDataHint(year, month, mode) {
  const el = $("mappingHint");
  if (!el) return;
  const y = Number.isFinite(year) ? year : "-";
  const m = Number.isFinite(month) ? month : "-";
  const modeName = mode === "month" ? "当月" : mode === "ytd" ? "截至目前" : "全年";
  el.textContent = `无数据：未上传 ${y} 年 ${modeName}${mode === "year" ? "" : `（月=${m}）`} 对应的Excel文件。`;
}

function refreshYearMonthOptions() {
  const years = (state.periodMeta.years || []).slice().sort((a, b) => a - b);
  const yearSelect = $("yearSelect");
  const monthSelect = $("monthSelect");
  if (!yearSelect || !monthSelect) return;

  const prevYear = Number(yearSelect.value) || null;
  const defaultYear = years.length ? years[years.length - 1] : new Date().getFullYear();
  fillSelect(yearSelect, years.map(String), String(prevYear && years.includes(prevYear) ? prevYear : defaultYear));
  state.uiDefaults.year = String(yearSelect.value);

  const y = Number(yearSelect.value);
  const monthsSet = state.periodMeta.monthsByYear.get(y) || new Set();
  const months = Array.from(monthsSet).sort((a, b) => a - b);
  const prevMonth = Number(monthSelect.value) || null;
  const defaultMonth = months.length ? months[months.length - 1] : new Date().getMonth() + 1;
  fillSelect(monthSelect, months.map(String), String(prevMonth && months.includes(prevMonth) ? prevMonth : defaultMonth));
  state.uiDefaults.month = String(monthSelect.value);
}

function ensureCharts() {
  if (!state.charts.billTrend) state.charts.billTrend = echarts.init($("chartBillTrend"));
  if (!state.charts.otTrend) state.charts.otTrend = echarts.init($("chartOtTrend"));
  if (!state.charts.yoymom) state.charts.yoymom = echarts.init($("chartYoYMoM"));
  window.addEventListener("resize", () => {
    state.charts.billTrend?.resize();
    state.charts.otTrend?.resize();
    state.charts.yoymom?.resize();
  });
}

function loadTitles() {
  let stored = null;
  try {
    stored = JSON.parse(localStorage.getItem(TITLE_STORAGE_KEY) || "null");
  } catch {
    stored = null;
  }
  state.titleState = { ...defaultTitles, ...(stored || {}) };
}

function saveTitles() {
  if (!state.titleState) return;
  try {
    localStorage.setItem(TITLE_STORAGE_KEY, JSON.stringify(state.titleState));
  } catch {
    // ignore quota errors etc.
  }
}

function applyTitles() {
  if (!state.titleState) loadTitles();
  const els = document.querySelectorAll(".editable-title");
  els.forEach((el) => {
    const key = el.dataset.titleKey;
    if (!key) return;
    const txt = state.titleState?.[key] || defaultTitles[key] || el.textContent;
    el.textContent = txt;
  });
}

function pct(x) {
  if (x == null || Number.isNaN(x)) return "-";
  if (!Number.isFinite(x)) return "-";
  return `${(x * 100).toFixed(2)}%`;
}

function isDailyOtCol(m) {
  if (m && m.__dailyOt != null) return !!m.__dailyOt;
  const n = normColName(m?.otCol);
  return n.includes("日均") || n.includes("daily");
}

function setKpis(agg) {
  $("kpiBill").textContent = fmtNumber(Number(agg.bill).toFixed(2));
  $("kpiStd").textContent = fmtNumber(Number(agg.std).toFixed(2));
  $("kpiRate").textContent = pct(agg.std > 0 ? agg.bill / agg.std : null);
  $("kpiOtDaily").textContent = fmtNumber(Number(agg.otDaily).toFixed(2));
}

function renderLine(chart, keys, series, name) {
  const ct = getChartTheme();
  chart.setOption({
    backgroundColor: "transparent",
    tooltip: {
      trigger: "axis",
      formatter: (params) => {
        const p = Array.isArray(params) ? params[0] : params;
        const v = Number(p?.data ?? 0);
        const axis = p?.axisValue ?? "";
        const isRate = String(p?.seriesName ?? "").includes("完成率");
        return `${axis}<br/>${p?.seriesName ?? name}: ${isRate ? v.toFixed(2) + "%" : fmtNumber(v)}`;
      },
    },
    grid: { left: 46, right: 18, top: 22, bottom: 30 },
    xAxis: { type: "category", data: keys, axisLabel: { color: ct.axis } },
    yAxis: {
      type: "value",
      axisLabel: { color: ct.axis },
      splitLine: { lineStyle: { color: ct.split } },
    },
    series: [
      {
        name,
        type: "line",
        data: series,
        smooth: true,
        showSymbol: false,
        areaStyle: { opacity: 0.14 },
      },
    ],
  });
}

function renderTrends(periodKeys, m, dept1Sel, dept2Sel, empSel) {
  // 趋势图按“文件名 YYYYMM”逐月汇总，不依赖日期列映射可靠性
  ensureCharts();

  const keys = Array.from(periodKeys || [])
    .map((p) => String(p))
    .filter((p) => state.periodMeta.yyyymm.has(p))
    .sort((a, b) => Number(a) - Number(b));

  const billSeries = [];
  const billRateSeries = [];
  const otSeries = [];

  for (const p of keys) {
    const rows = getPeriodData(p).rows || [];
    let billSum = 0;
    let stdSum = 0;
    let otSum = 0;
    for (const r of rows) {
      const d1 = (r?.[m.dept1Col] ?? "").toString().trim();
      const d2 = (r?.[m.dept2Col] ?? "").toString().trim();
      if (dept1Sel.size && !dept1Sel.has(d1)) continue;
      if (dept2Sel.size && !dept2Sel.has(d2)) continue;
      if (empSel) {
        const emp = (r?.[m.empCol] ?? "").toString().trim();
        if (emp !== empSel) continue;
      }
      billSum += toNumber(r?.[m.billCol]) ?? 0;
      stdSum += toNumber(r?.[m.stdCol]) ?? 0;
      otSum += toNumber(r?.[m.otCol]) ?? 0;
    }
    // 账单小时变化趋势：按完成率（账单小时 / 额定账单小时）展示
    const rate = stdSum > 0 ? (billSum / stdSum) * 100 : 0;
    billRateSeries.push(rate);
    otSeries.push(otSum);
  }

  const labels = keys.map((p) => `${p.slice(0, 4)}-${p.slice(4, 6)}`);
  renderLine(state.charts.billTrend, labels, billRateSeries, "账单小时完成率");
  renderLine(state.charts.otTrend, labels, otSeries, "加班小时");
}

function renderYoYMoM(monthAgg, selectedYear, selectedMonth) {
  ensureCharts();
  const ct = getChartTheme();
  const chart = state.charts.yoymom;
  const months = Array.from({ length: 12 }, (_, i) => `${i + 1}`.padStart(2, "0"));
  const y = String(selectedYear);
  const prevY = String(Number(selectedYear) - 1);

  const get = (yy, mm) => monthAgg.get(`${yy}-${mm}`) ?? { bill: 0, std: 0, ot: 0 };
  const rate = (x) => (x.std > 0 ? x.bill / x.std : 0);

  const curRate = months.map((mm) => rate(get(y, mm)));
  const curOt = months.map((mm) => get(y, mm).ot);
  const prevRate = months.map((mm) => rate(get(prevY, mm)));
  const prevOt = months.map((mm) => get(prevY, mm).ot);

  // MoM at selectedMonth (or last available)
  const mm = String(selectedMonth).padStart(2, "0");
  const cur = get(y, mm);
  const lastMm = String(Math.max(1, Number(mm) - 1)).padStart(2, "0");
  const last = get(y, lastMm);
  const momRate = last.std > 0 ? rate(cur) - rate(last) : null;
  const momOt = curOt[Number(mm) - 1] - (get(y, lastMm).ot ?? 0);

  chart.setOption({
    backgroundColor: "transparent",
    tooltip: { trigger: "axis" },
    legend: { textStyle: { color: ct.legend } },
    grid: { left: 50, right: 50, top: 40, bottom: 30 },
    xAxis: { type: "category", data: months, axisLabel: { color: ct.axis } },
    yAxis: [
      {
        type: "value",
        axisLabel: { color: ct.axis, formatter: (v) => `${Number(v).toFixed(0)}%` },
        splitLine: { lineStyle: { color: ct.split } },
        min: 0,
        max: 200,
      },
      {
        type: "value",
        axisLabel: { color: ct.axis },
        splitLine: { show: false },
      },
    ],
    title: {
      text: `选中：${y}-${mm}  |  环比完成率：${momRate == null ? "-" : (momRate * 100).toFixed(2) + "%"}  |  环比加班小时：${fmtNumber(momOt)}`,
      left: 10,
      top: 6,
      textStyle: { color: "rgba(255,255,255,0.72)", fontSize: 12, fontWeight: 600 },
    },
    series: [
      {
        name: `${y} 完成率`,
        type: "bar",
        yAxisIndex: 0,
        data: curRate.map((x) => x * 100),
        itemStyle: { opacity: 0.6 },
      },
      {
        name: `${prevY} 完成率`,
        type: "bar",
        yAxisIndex: 0,
        data: prevRate.map((x) => x * 100),
        itemStyle: { opacity: 0.25 },
      },
      {
        name: `${y} 加班小时`,
        type: "line",
        yAxisIndex: 1,
        data: curOt,
        smooth: true,
        showSymbol: false,
      },
      {
        name: `${prevY} 加班小时`,
        type: "line",
        yAxisIndex: 1,
        data: prevOt,
        smooth: true,
        showSymbol: false,
        lineStyle: { opacity: 0.35 },
      },
    ],
  });
}

function sumAgg(rows, m) {
  let bill = 0;
  let std = 0;
  let ot = 0;
  let cost = 0;
  let rev = 0;
  const dailyOt = isDailyOtCol(m);
  const empMap = new Map(); // empKey => {ot, days:Set, cnt, dayCount}
  for (const r of rows) {
    const d = toDate(r?.[m.dateCol]);
    bill += toNumber(r?.[m.billCol]) ?? 0;
    std += toNumber(r?.[m.stdCol]) ?? 0;
    ot += toNumber(r?.[m.otCol]) ?? 0;
    cost += toNumber(r?.[m.costCol]) ?? 0;
    rev += toNumber(r?.[m.revCol]) ?? 0;
    const emp = (r?.[m.empCol] ?? "").toString().trim();
    if (emp) {
      const key = emp;
      const cur =
        empMap.get(key) ?? { ot: 0, days: new Set(), cnt: 0, dayCount: 0 };
      cur.ot += toNumber(r?.[m.otCol]) ?? 0;
      const shareCnt = toNumber(r?.__shareCnt);
      cur.cnt += Number.isFinite(shareCnt) ? shareCnt : 1;
      if (!dailyOt) {
        const shareDayCount = toNumber(r?.__shareDayCount);
        if (Number.isFinite(shareDayCount) && shareDayCount > 0) {
          cur.dayCount += shareDayCount;
        } else if (d) {
          cur.days.add(d.toISOString().slice(0, 10));
        }
      }
      empMap.set(key, cur);
    }
  }
  // 口径：日均加班小时模块 = 所有人“日均加班小时”合计 / 人员数量
  let sumEmpOtDaily = 0;
  for (const v of empMap.values()) {
    if (dailyOt) {
      const cnt = v.cnt || 0;
      sumEmpOtDaily += cnt > 0 ? v.ot / cnt : 0;
    } else {
      const days = Number.isFinite(v.dayCount) && v.dayCount > 0 ? v.dayCount : v.days.size || 0;
      sumEmpOtDaily += days > 0 ? v.ot / days : 0;
    }
  }
  const headcount = empMap.size || 0;
  const otDaily = headcount > 0 ? sumEmpOtDaily / headcount : 0;
  return { bill, std, ot, cost, rev, otDaily, headcount };
}

function buildMonthAgg(rows, m) {
  const map = new Map(); // YYYY-MM => agg
  for (const r of rows) {
    const d = toDate(r?.[m.dateCol]);
    if (!d) continue;
    const k = formatKeyByGrain(d, "month");
    const cur = map.get(k) ?? { bill: 0, std: 0, ot: 0 };
    cur.bill += toNumber(r?.[m.billCol]) ?? 0;
    cur.std += toNumber(r?.[m.stdCol]) ?? 0;
    cur.ot += toNumber(r?.[m.otCol]) ?? 0;
    map.set(k, cur);
  }
  return map;
}

function buildMonthAggFromPeriods(periodKeys, m, dept1Sel, dept2Sel, empSel) {
  const map = new Map(); // YYYY-MM => { bill, std, ot }
  const keys = Array.from(periodKeys || [])
    .map((p) => String(p))
    .filter((p) => /^\d{6}$/.test(p))
    .sort((a, b) => Number(a) - Number(b));

  for (const p of keys) {
    const ymKey = `${p.slice(0, 4)}-${p.slice(4, 6)}`;
    let cur = map.get(ymKey) || { bill: 0, std: 0, ot: 0 };
    const rows = getPeriodData(p).rows || [];
    for (const r of rows) {
      const d1 = (r?.[m.dept1Col] ?? "").toString().trim();
      const d2 = (r?.[m.dept2Col] ?? "").toString().trim();
      if (dept1Sel?.size && !dept1Sel.has(d1)) continue;
      if (dept2Sel?.size && !dept2Sel.has(d2)) continue;
      if (empSel) {
        const emp = (r?.[m.empCol] ?? "").toString().trim();
        if (emp !== empSel) continue;
      }
      cur.bill += toNumber(r?.[m.billCol]) ?? 0;
      cur.std += toNumber(r?.[m.stdCol]) ?? 0;
      cur.ot += toNumber(r?.[m.otCol]) ?? 0;
    }
    map.set(ymKey, cur);
  }
  return map;
}

function computeAnomalies(rows, m) {
  // 条件（按员工聚合后判断）
  const empMap = new Map(); // key => agg
  const dailyOt = isDailyOtCol(m);
  for (const r of rows) {
    const emp = (r?.[m.empCol] ?? "").toString().trim();
    if (!emp) continue;
    const key = [
      (r?.[m.dept1Col] ?? "").toString().trim(),
      (r?.[m.dept2Col] ?? "").toString().trim(),
      emp,
    ].join(" / ");
    const cur =
      empMap.get(key) ?? { dept1: "", dept2: "", emp, bill: 0, std: 0, otDaily: 0, ot: 0, days: new Set(), cnt: 0 };
    cur.dept1 = (r?.[m.dept1Col] ?? "").toString().trim();
    cur.dept2 = (r?.[m.dept2Col] ?? "").toString().trim();
    cur.bill += toNumber(r?.[m.billCol]) ?? 0;
    cur.std += toNumber(r?.[m.stdCol]) ?? 0;
    cur.ot += toNumber(r?.[m.otCol]) ?? 0;
    const shareCnt = toNumber(r?.__shareCnt);
    cur.cnt += Number.isFinite(shareCnt) ? shareCnt : 1;
    if (!dailyOt) {
      const shareDayCount = toNumber(r?.__shareDayCount);
      if (Number.isFinite(shareDayCount) && shareDayCount > 0) {
        cur.dayCount = (cur.dayCount ?? 0) + shareDayCount;
      } else {
        const d = toDate(r?.[m.dateCol]);
        if (d) cur.days.add(d.toISOString().slice(0, 10));
      }
    }
    empMap.set(key, cur);
  }

  const alerts = [];
  for (const [key, v] of empMap.entries()) {
    const rate = v.std > 0 ? v.bill / v.std : null;
    const days =
      dailyOt
        ? 0
        : Number.isFinite(v.dayCount) && v.dayCount > 0
          ? v.dayCount
          : v.days.size || 0;
    const cnt = v.cnt || 0;
    const otDaily = dailyOt ? (cnt > 0 ? v.ot / cnt : 0) : days > 0 ? v.ot / days : 0;
    const billPct = rate == null ? null : rate * 100;
    const msgs = [];
    const reasons = [];
    if (billPct != null && billPct > 150) reasons.push({ code: "rate_high_150", text: `账单小时完成率高于150%（${billPct.toFixed(2)}%）` });
    if (billPct != null && billPct < 50) reasons.push({ code: "rate_low_50", text: `账单小时完成率低于50%（${billPct.toFixed(2)}%）` });
    if (billPct != null && billPct > 150 && otDaily > 2.5) reasons.push({ code: "rate_high_ot_high", text: `完成率>150% 且 日均加班>2.5（${otDaily.toFixed(2)}）` });
    if (billPct != null && billPct > 150 && otDaily < 1) reasons.push({ code: "rate_high_ot_low", text: `完成率>150% 但 日均加班<1（${otDaily.toFixed(2)}）` });
    if (billPct != null && billPct < 50 && otDaily > 2) reasons.push({ code: "rate_low_ot_high", text: `完成率<50% 但 日均加班>2（${otDaily.toFixed(2)}）` });
    for (const r0 of reasons) msgs.push(r0.text);
    if (msgs.length) {
      alerts.push({
        key,
        dept1: v.dept1,
        dept2: v.dept2,
        emp: v.emp,
        rate,
        bill: v.bill,
        std: v.std,
        otDaily,
        msgs,
        reasons,
      });
    }
  }
  alerts.sort((a, b) => (b.rate ?? 0) - (a.rate ?? 0));
  return alerts;
}

function renderAlerts(alerts) {
  const reasonCodeLabel = {
    rate_high_150: "完成率高于150%",
    rate_low_50: "完成率低于50%",
    rate_high_ot_high: "完成率>150%且加班>2.5",
    rate_high_ot_low: "完成率>150%但加班<1",
    rate_low_ot_high: "完成率<50%但加班>2",
  };
  const table = $("alertsTable");
  const pageInfoEl = $("alertPageInfo");
  if (!table) return;
  // 自动根据 alerts 填充筛选下拉选项（单选，下拉交互与“所属年份”类似）
  const empSel = $("alertEmpFilter");
  const d1Sel = $("alertDept1Filter");
  const d2Sel = $("alertDept2Filter");
  if (empSel && !empSel.dataset.filled) {
    const emps = uniq(alerts.map((a) => a.emp).filter(Boolean)).sort((a, b) => a.localeCompare(b, "zh-CN"));
    empSel.innerHTML = "";
    const all = document.createElement("option");
    all.value = "";
    all.textContent = "全部";
    empSel.appendChild(all);
    for (const v of emps) {
      const o = document.createElement("option");
      o.value = v;
      o.textContent = v;
      empSel.appendChild(o);
    }
    empSel.dataset.filled = "1";
  }
  if (d1Sel && !d1Sel.dataset.filled) {
    const d1s = uniq(alerts.map((a) => a.dept1).filter(Boolean)).sort(sortByDeptCode);
    d1Sel.innerHTML = "";
    const all = document.createElement("option");
    all.value = "";
    all.textContent = "全部";
    d1Sel.appendChild(all);
    for (const v of d1s) {
      const o = document.createElement("option");
      o.value = v;
      o.textContent = v;
      d1Sel.appendChild(o);
    }
    d1Sel.dataset.filled = "1";
  }
  if (d2Sel && !d2Sel.dataset.filled) {
    const d2s = uniq(alerts.map((a) => a.dept2 || "未分配二级部门")).sort(sortByDeptCode);
    d2Sel.innerHTML = "";
    const all = document.createElement("option");
    all.value = "";
    all.textContent = "全部";
    d2Sel.appendChild(all);
    for (const v of d2s) {
      const o = document.createElement("option");
      o.value = v;
      o.textContent = v;
      d2Sel.appendChild(o);
    }
    d2Sel.dataset.filled = "1";
  }

  let empVal = empSel ? empSel.value : "";
  let d1Val = d1Sel ? d1Sel.value : "";
  let d2Val = d2Sel ? d2Sel.value : "";

  // 明细表联动：点击部门/人员后，异常模块也按同一范围过滤展示
  const ts = state.tableSelection;
  if (ts && ts.selKey) {
    if (ts.scope === "emp") {
      empVal = ts.emp ? ts.emp.toString().trim() : "";
      d1Val = ts.dept1 || "";
      d2Val = ts.dept2 || "";
    } else if (ts.scope === "dept2") {
      empVal = "";
      d1Val = ts.dept1 || "";
      d2Val = ts.dept2 || "";
    } else if (ts.scope === "dept1") {
      empVal = "";
      d1Val = ts.dept1 || "";
      d2Val = "";
    }
  }

  const filteredAlerts = alerts.filter((a) => {
    if (empVal && a.emp !== empVal) return false;
    if (d1Val && a.dept1 !== d1Val) return false;
    const d2 = a.dept2 || "未分配二级部门";
    if (d2Val && d2 !== d2Val) return false;
    return true;
  });
  // 构建层级结构：dept1 -> dept2 -> [employees]
  const byDept1 = new Map();
  for (const a of filteredAlerts) {
    const d1 = a.dept1 || "未分配一级部门";
    const d2 = a.dept2 || "未分配二级部门";
    if (!byDept1.has(d1)) byDept1.set(d1, new Map());
    const m2 = byDept1.get(d1);
    if (!m2.has(d2)) m2.set(d2, []);
    m2.get(d2).push(a);
  }

  const dept1List = Array.from(byDept1.keys()).sort(sortByDeptCode);
  if (!dept1List.length) {
    table.innerHTML = "";
    const thead = document.createElement("thead");
    const trh = document.createElement("tr");
    ["所属部门（一级 / 二级 / 员工）", "预警提示"].forEach((t) => {
      const th = document.createElement("th");
      th.textContent = t;
      trh.appendChild(th);
    });
    thead.appendChild(trh);
    const tbody = document.createElement("tbody");
    const tr = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = 2;
    td.textContent = "按当前异常筛选条件，未发现异常记录。";
    tr.appendChild(td);
    tbody.appendChild(tr);
    table.appendChild(thead);
    table.appendChild(tbody);
    if (pageInfoEl) pageInfoEl.textContent = "第 0 / 0 页";
    return;
  }

  // 分页：按一级部门分页，每页展示 N 个一级部门树
  const view = state.alertView;
  const pageSize = Math.max(1, Math.min(20, view.pageSize || 5));
  const totalTop = dept1List.length;
  const totalPages = Math.max(1, Math.ceil(totalTop / pageSize));
  if (view.page > totalPages) view.page = totalPages;
  const startTop = (view.page - 1) * pageSize;
  const pageDept1 = dept1List.slice(startTop, startTop + pageSize);

  // 画表格
  table.innerHTML = "";
  const thead = document.createElement("thead");
  const trh = document.createElement("tr");
  ["所属部门（一级 / 二级 / 员工）", "预警提示"].forEach((t) => {
    const th = document.createElement("th");
    th.textContent = t;
    trh.appendChild(th);
  });
  thead.appendChild(trh);
  const tbody = document.createElement("tbody");

  const makeRow = (label, msg, level, hasChildren, isLeaf) => {
    const indent = level * 14;
    const lamp = isLeaf ? '<span class="lamp" title="异常"></span>' : "";
    const arrow = hasChildren ? `<button class="tree-btn" data-toggle="1">▸</button>` : `<span class="tree-sp"></span>`;
    return `
      <tr class="tree-row" data-level="${level}" ${level > 0 ? 'data-hidden="1"' : ""}>
        <td style="padding-left:${10 + indent}px">
          <span class="tree">${arrow}${lamp}<span class="tree-label">${label}</span></span>
        </td>
        <td>${msg}</td>
      </tr>
    `;
  };

  for (const d1 of pageDept1) {
    const m2 = byDept1.get(d1);
    const d2List = Array.from(m2.keys()).sort(sortByDeptCode);
    const allAlerts = Array.from(m2.values()).flat();
    const empCount = uniq(allAlerts.map((a) => `${a.dept1}|||${a.dept2}|||${a.emp}`)).length;
    tbody.insertAdjacentHTML("beforeend", makeRow(d1, `异常 ${empCount} 人`, 0, d2List.length > 0, false));

    for (const d2 of d2List) {
      const list = m2.get(d2);
      const empKeys = uniq(list.map((a) => `${a.dept1}|||${a.dept2}|||${a.emp}`));
      const stats = new Map();
      for (const a of list) for (const r of a.reasons ?? []) stats.set(r.code, (stats.get(r.code) ?? 0) + 1);
      const summary = Array.from(stats.entries())
        .sort((a, b) => b[1] - a[1])
        .map(([code, cnt]) => `${reasonCodeLabel[code] || code}: ${cnt}人`)
        .join("，");
      const msg = summary || `异常 ${empKeys.length} 人`;
      tbody.insertAdjacentHTML("beforeend", makeRow(d2, msg, 1, true, false));

      const byEmp = new Map();
      for (const a of list) {
        const key = a.emp || "未命名";
        if (!byEmp.has(key)) byEmp.set(key, []);
        byEmp.get(key).push(a);
      }
      const empList = Array.from(byEmp.keys()).sort((a, b) => a.localeCompare(b, "zh-CN"));
      for (const emp of empList) {
        const arr = byEmp.get(emp);
        const reasons = uniq(arr.flatMap((x) => x.msgs || []));
        tbody.insertAdjacentHTML("beforeend", makeRow(emp, reasons.join("；"), 2, false, true));
      }
    }
  }

  table.appendChild(thead);
  table.appendChild(tbody);

  if (pageInfoEl) pageInfoEl.textContent = `第 ${view.page} / ${totalPages} 页（共 ${dept1List.length} 个一级部门）`;

  if (table.dataset.treeBound !== "1") {
    table.dataset.treeBound = "1";
    table.addEventListener("click", (e) => {
      const btn = e.target?.closest?.(".tree-btn");
      if (!btn) return;
      const tr = btn.closest("tr");
      const level = Number(tr.getAttribute("data-level") || 0);
      const isOpen = tr.getAttribute("data-open") === "1";
      tr.setAttribute("data-open", isOpen ? "0" : "1");
      btn.textContent = isOpen ? "▸" : "▾";
      let next = tr.nextElementSibling;
      while (next) {
        const nl = Number(next.getAttribute("data-level") || 0);
        if (nl <= level) break;
        if (!isOpen) {
          if (nl === level + 1) next.removeAttribute("data-hidden");
        } else {
          next.setAttribute("data-hidden", "1");
          next.setAttribute("data-open", "0");
          const b = next.querySelector?.(".tree-btn");
          if (b) b.textContent = "▸";
        }
        next = next.nextElementSibling;
      }
    });
  }
}

function renderSummaryTable(rows, m, alerts) {
  const alertEmp = new Set(alerts.map((a) => `${a.dept1}|||${a.dept2}|||${a.emp}`));
  const alertDept2 = new Set(alerts.map((a) => `${a.dept1}|||${a.dept2}`));
  const alertDept1 = new Set(alerts.map((a) => `${a.dept1}`));
  const table = $("summaryTable");
  table.innerHTML = "";
  const dailyOt = isDailyOtCol(m);
  const escapeAttr = (s) =>
    String(s ?? "")
      .replace(/&/g, "&amp;")
      .replace(/"/g, "&quot;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");

  const cols = [
    { title: "一级部门", key: "name", sortable: false },
    { title: "账单小时", key: "bill", sortable: true },
    { title: "额定账单小时", key: "std", sortable: true },
    { title: "小时完成率", key: "rate", sortable: true },
    { title: "人力成本", key: "cost", sortable: true },
    { title: "收入", key: "rev", sortable: true },
    { title: "人力成本率", key: "costRate", sortable: true },
    { title: "日均加班小时", key: "otDaily", sortable: true },
  ];

  const thead = document.createElement("thead");
  const trh = document.createElement("tr");
  for (const c of cols) {
    const th = document.createElement("th");
    if (!c.sortable) {
      th.textContent = c.title;
    } else {
      th.innerHTML = `
        <div class="th-wrap">
          <span>${c.title}</span>
          <span class="th-actions">
            <button class="th-btn" data-sort-key="${c.key}" data-sort-dir="asc" title="升序">↑</button>
            <button class="th-btn" data-sort-key="${c.key}" data-sort-dir="desc" title="降序">↓</button>
            <button class="th-btn" data-sort-key="${c.key}" data-sort-dir="reset" title="重置">⟲</button>
          </span>
        </div>
      `;
    }
    trh.appendChild(th);
  }
  thead.appendChild(trh);
  table.appendChild(thead);

  // aggregate: dept1 -> dept2 -> emp
  const node = () => ({
    bill: 0,
    std: 0,
    cost: 0,
    rev: 0,
    ot: 0,
    days: new Set(),
    dayCount: 0, // share 模式聚合后可直接使用天数，避免依赖日期列
    cnt: 0,
    children: new Map(),
  });
  const root = new Map();

  for (const r of rows) {
    const d1 = (r?.[m.dept1Col] ?? "未分配").toString().trim() || "未分配";
    const d2 = (r?.[m.dept2Col] ?? "未分配").toString().trim() || "未分配";
    const emp = (r?.[m.empCol] ?? "未命名").toString().trim() || "未命名";
    const date = toDate(r?.[m.dateCol]);
    const shareCnt = toNumber(r?.__shareCnt);
    const shareDayCount = toNumber(r?.__shareDayCount);

    const getOr = (map, k) => {
      if (!map.has(k)) map.set(k, node());
      return map.get(k);
    };

    const n1 = getOr(root, d1);
    const n2 = getOr(n1.children, d2);
    const n3 = getOr(n2.children, emp);

    const add = (n) => {
      n.bill += toNumber(r?.[m.billCol]) ?? 0;
      n.std += toNumber(r?.[m.stdCol]) ?? 0;
      n.cost += toNumber(r?.[m.costCol]) ?? 0;
      n.rev += toNumber(r?.[m.revCol]) ?? 0;
      n.ot += toNumber(r?.[m.otCol]) ?? 0;
      n.cnt += Number.isFinite(shareCnt) ? shareCnt : 1;
      if (!dailyOt) {
        if (Number.isFinite(shareDayCount) && shareDayCount > 0) {
          n.dayCount += shareDayCount;
        } else if (date) {
          n.days.add(date.toISOString().slice(0, 10));
        }
      }
    };
    add(n1);
    add(n2);
    add(n3);
  }

  const tbody = document.createElement("tbody");

  const sortChildrenKeys = (map, getNode) => {
    const keys = Array.from(map.keys());
    const { key, dir } = state.tableSort;
    if (!key || !dir) return keys.sort(sortByDeptCode);
    const factor = dir === "asc" ? 1 : -1;
    const valueOf = (n) => {
      const rate = n.std > 0 ? n.bill / n.std : null;
      const costRate = n.rev > 0 ? n.cost / n.rev : null;
      const days = dailyOt
        ? 0
        : Number.isFinite(n.dayCount) && n.dayCount > 0
          ? n.dayCount
          : n.days.size || 0;
      const cnt = n.cnt || 0;
      const otDaily = dailyOt ? (cnt > 0 ? n.ot / cnt : 0) : days > 0 ? n.ot / days : 0;
      const dict = { bill: n.bill, std: n.std, rate: rate ?? -Infinity, cost: n.cost, rev: n.rev, costRate: costRate ?? -Infinity, otDaily };
      return dict[key] ?? 0;
    };
    keys.sort((a, b) => {
      const va = valueOf(getNode(a));
      const vb = valueOf(getNode(b));
      if (va === vb) return sortByDeptCode(a, b);
      return (va - vb) * factor;
    });
    return keys;
  };

  const rowHtml = (label, n, level, hasChildren, isAlertRow, sel) => {
    const rate = n.std > 0 ? n.bill / n.std : null;
    const costRate = n.rev > 0 ? n.cost / n.rev : null;
    const days = dailyOt
      ? 0
      : Number.isFinite(n.dayCount) && n.dayCount > 0
        ? n.dayCount
        : n.days.size || 0;
    const cnt = n.cnt || 0;
    const otDaily = dailyOt ? (cnt > 0 ? n.ot / cnt : 0) : days > 0 ? n.ot / days : 0;
    const indent = level * 14;
    const lamp = isAlertRow ? `<span class="lamp" title="异常"></span>` : "";
    const arrow = hasChildren ? `<button class="tree-btn" data-toggle="1">▸</button>` : `<span class="tree-sp"></span>`;
    const scope = sel?.scope || "";
    const selKey = sel?.selKey || "";
    const isSelected = state.tableSelection && selKey && state.tableSelection.selKey === selKey;
    const cls = isSelected ? "tree-row row-selected" : "tree-row";
    const dept1 = sel?.dept1 || "";
    const dept2 = sel?.dept2 || "";
    const emp = sel?.emp || "";
    return `
      <tr
        class="${cls}"
        data-level="${level}"
        ${level > 0 ? 'data-hidden="1"' : ""}
        data-scope="${escapeAttr(scope)}"
        data-selkey="${escapeAttr(selKey)}"
        data-dept1="${escapeAttr(dept1)}"
        data-dept2="${escapeAttr(dept2)}"
        data-emp="${escapeAttr(emp)}"
      >
        <td style="padding-left:${10 + indent}px">
          <span class="tree">${arrow}${lamp}<span class="tree-label">${label}</span></span>
        </td>
        <td>${fmtNumber(Number(n.bill).toFixed(2))}</td>
        <td>${fmtNumber(Number(n.std).toFixed(2))}</td>
        <td>${pct(rate)}</td>
        <td>${fmtNumber(Number(n.cost).toFixed(2))}</td>
        <td>${fmtNumber(Number(n.rev).toFixed(2))}</td>
        <td>${pct(costRate)}</td>
        <td>${fmtNumber(Number(otDaily).toFixed(2))}</td>
      </tr>
    `;
  };

  const appendRows = () => {
    const d1s = sortChildrenKeys(root, (k) => root.get(k));
    for (const d1 of d1s) {
      const n1 = root.get(d1);
      const sel1 = { scope: "dept1", dept1: d1, selKey: `dept1|||${d1}` };
      tbody.insertAdjacentHTML(
        "beforeend",
        rowHtml(d1, n1, 0, n1.children.size > 0, alertDept1.has(d1), sel1)
      );
      const d2s = sortChildrenKeys(n1.children, (k) => n1.children.get(k));
      for (const d2 of d2s) {
        const n2 = n1.children.get(d2);
        const sel2 = { scope: "dept2", dept1: d1, dept2: d2, selKey: `dept2|||${d1}|||${d2}` };
        tbody.insertAdjacentHTML(
          "beforeend",
          rowHtml(d2, n2, 1, n2.children.size > 0, alertDept2.has(`${d1}|||${d2}`), sel2)
        );
        const emps = sortChildrenKeys(n2.children, (k) => n2.children.get(k));
        for (const emp of emps) {
          const n3 = n2.children.get(emp);
          const key = `${d1}|||${d2}|||${emp}`;
          const sel3 = {
            scope: "emp",
            dept1: d1,
            dept2: d2,
            emp,
            selKey: `emp|||${d1}|||${d2}|||${emp}`,
          };
          tbody.insertAdjacentHTML(
            "beforeend",
            rowHtml(emp, n3, 2, false, alertEmp.has(key), sel3)
          );
        }
      }
    }
  };
  appendRows();
  table.appendChild(tbody);

  // expand/collapse: simple DOM toggle by levels (toggle next rows until level <= current)
  if (table.dataset.treeBound !== "1") {
    table.dataset.treeBound = "1";
    table.addEventListener("click", (e) => {
    const btn = e.target?.closest?.(".tree-btn");
    if (!btn) return;
    const tr = btn.closest("tr");
    const level = Number(tr.getAttribute("data-level") || 0);
    const isOpen = tr.getAttribute("data-open") === "1";
    tr.setAttribute("data-open", isOpen ? "0" : "1");
    btn.textContent = isOpen ? "▸" : "▾";
    let next = tr.nextElementSibling;
    while (next) {
      const nl = Number(next.getAttribute("data-level") || 0);
      if (nl <= level) break;
      if (!isOpen) {
        // opening: show only direct children, deeper remain hidden unless their parents open
        if (nl === level + 1) next.removeAttribute("data-hidden");
      } else {
        // closing: hide all descendants
        next.setAttribute("data-hidden", "1");
        next.setAttribute("data-open", "0");
        const b = next.querySelector?.(".tree-btn");
        if (b) b.textContent = "▸";
      }
      next = next.nextElementSibling;
    }
    });
  }

  if (table.dataset.sortBound !== "1") {
    table.dataset.sortBound = "1";
    table.addEventListener("click", (e) => {
      const btn = e.target?.closest?.(".th-btn");
      if (!btn) return;
      const key = btn.getAttribute("data-sort-key");
      const dir = btn.getAttribute("data-sort-dir");
      if (dir === "reset") {
        state.tableSort.key = null;
        state.tableSort.dir = null;
      } else {
        state.tableSort.key = key;
        state.tableSort.dir = dir;
      }
      applyFiltersAndRender();
    });
  }

  // 明细树表：点击行 -> 全局联动（展开按钮不触发）
  if (table.dataset.selectBound !== "1") {
    table.dataset.selectBound = "1";
    table.addEventListener("click", (e) => {
      if (e.target?.closest?.(".tree-btn")) return; // 折叠/展开按钮
      const tr = e.target?.closest?.("tr.tree-row");
      if (!tr) return;
      const selKey = tr.getAttribute("data-selkey") || "";
      const scope = tr.getAttribute("data-scope") || "";
      if (!selKey || !scope) return;

      const same = state.tableSelection && state.tableSelection.selKey === selKey;
      if (same) {
        state.tableSelection = null;
      } else {
        state.tableSelection = {
          selKey,
          scope,
          dept1: tr.getAttribute("data-dept1") || "",
          dept2: tr.getAttribute("data-dept2") || "",
          emp: tr.getAttribute("data-emp") || "",
        };
      }
      applyFiltersAndRender();
    });
  }
}

function applyFiltersAndRender() {
  let dept1Sel = new Set(getSelectedMulti($("dept1Select")));
  let dept2Sel = new Set(getSelectedMulti($("dept2Select")));
  const year = Number($("yearSelect").value);
  const month = Number($("monthSelect").value);
  const mode = document.querySelector('input[name="monthMode"]:checked')?.value ?? "month";

  // 注意：部门下拉选项会在后续根据 rows 刷新（避免 state.rows 尚未更新导致选项为空）

  // 1) 按“文件名 YYYYMM”取数。当月=单月；截至目前=1..选中月；全年=1..最新有数据的月
  const collectPeriods = () => {
    const names = [];
    if (mode === "month") {
      names.push(toYYYYMM(year, month));
    } else if (mode === "ytd") {
      for (let mm = 1; mm <= month; mm++) names.push(toYYYYMM(year, mm));
    } else {
      const maxM = state.periodMeta.maxMonthByYear.get(year) ?? 0;
      for (let mm = 1; mm <= maxM; mm++) names.push(toYYYYMM(year, mm));
    }
    return names;
  };
  const targetPeriods = collectPeriods();
  const existingTargets = targetPeriods.filter((p) => state.periodMeta.yyyymm.has(p));
  if (!existingTargets.length) {
    setNoDataHint(year, month, mode);
    setKpis({ bill: 0, std: 0, otDaily: 0 });
    state.latestAlerts = [];
    renderAlerts([]);
    renderSummaryTable([], state.mapping || {}, []);
    ensureCharts();
    state.charts.billTrend?.clear();
    state.charts.otTrend?.clear();
    state.charts.yoymom?.clear();
    return;
  }

  const rows = existingTargets.flatMap((p) => getPeriodData(p).rows);
  const m = getPeriodData(existingTargets[0]).mapping || state.mapping;
  if (!m) return;
  setMappingHint(m);

  // 关键：根据当前时间窗口 rows，刷新部门下拉选项（否则你会看到二级/一级为空）
  state.rows = rows;
  state.mapping = m;
  const d1s = uniq(
    state.rows.map((r) => (r?.[m.dept1Col] ?? "").toString().trim()).filter(Boolean)
  ).sort(sortByDeptCode);
  const dept1Select = $("dept1Select");
  const dept2Select = $("dept2Select");
  if (dept1Select && dept2Select) {
    const prevDept1 = dept1Sel;
    dept1Select.innerHTML = "";
    for (const v of d1s) {
      const o = document.createElement("option");
      o.value = v;
      o.textContent = v;
      if (prevDept1.has(v)) o.selected = true;
      dept1Select.appendChild(o);
    }
    // 二级部门联动一级部门
    refreshDept2Options();
    // 刷新后重新读取当前选择集合（可能由于可选项变化而被清空）
    dept1Sel = new Set(getSelectedMulti(dept1Select));
    dept2Sel = new Set(getSelectedMulti(dept2Select));
  }

  // 明细树表联动：点击行后，覆盖“部门/人员”过滤范围
  let empSel = null;
  if (state.tableSelection && state.tableSelection.selKey) {
    const ts = state.tableSelection;
    if (ts.scope === "dept1") {
      dept1Sel = ts.dept1 ? new Set([ts.dept1]) : new Set();
      dept2Sel = new Set();
      empSel = null;
    } else if (ts.scope === "dept2") {
      dept1Sel = ts.dept1 ? new Set([ts.dept1]) : new Set();
      dept2Sel = ts.dept2 ? new Set([ts.dept2]) : new Set();
      empSel = null;
    } else if (ts.scope === "emp") {
      dept1Sel = ts.dept1 ? new Set([ts.dept1]) : new Set();
      dept2Sel = ts.dept2 ? new Set([ts.dept2]) : new Set();
      empSel = ts.emp ? ts.emp.toString().trim() : null;
    }
  }

  // 2) 部门筛选（基于合并后的 rows）
  // 重要：当数据已通过“文件名 YYYYMM”选中对应月份时，不再依赖日期列做过滤，
  // 否则一旦日期列解析失败会导致整表空数据。
  const filtered = rows.filter((r) => {
    const d1 = (r?.[m.dept1Col] ?? "").toString().trim();
    const d2 = (r?.[m.dept2Col] ?? "").toString().trim();
    if (dept1Sel.size && !dept1Sel.has(d1)) return false;
    if (dept2Sel.size && !dept2Sel.has(d2)) return false;
    if (empSel) {
      const emp = (r?.[m.empCol] ?? "").toString().trim();
      if (emp !== empSel) return false;
    }
    return true;
  });

  const agg = sumAgg(filtered, m);
  setKpis(agg);

  const alerts = computeAnomalies(filtered, m);
  state.latestAlerts = alerts;
  renderAlerts(alerts);
  renderSummaryTable(filtered, m, alerts);

  // trends by month (labels driven by YYYYMM file names)
  renderTrends(existingTargets, m, dept1Sel, dept2Sel, empSel);

  // YoY/MoM：按“文件名 YYYYMM”逐月汇总（避免依赖日期列）
  const allPeriodKeys = state.periodMeta.yyyymm.size ? Array.from(state.periodMeta.yyyymm) : [];
  const monthAgg = buildMonthAggFromPeriods(allPeriodKeys, m, dept1Sel, dept2Sel, empSel);
  renderYoYMoM(monthAgg, year, month);
}

function populateControls() {
  refreshYearMonthOptions();
  applyFiltersAndRender();
}

function parseSheetToJson(ws, opts = {}) {
  const maxDataRows = opts.maxDataRows ?? null;
  // 列头 = 工作表中第一行“非空有效行”（自动跳过空行）
  // 这个口径对你的数据源更稳定：避免“表头并非最佳匹配行”导致 cols/rows错位。
  const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, blankrows: false });
  const isEmptyRow = (row) => {
    if (!row || !row.length) return true;
    return row.every((c) => c == null || String(c).trim() === "");
  };
  if (!matrix || !matrix.length) return { rows: [], cols: [] };

  const headerIdx = (() => {
    const idx = matrix.findIndex((r) => !isEmptyRow(r));
    return idx < 0 ? 0 : idx;
  })();

  const headerRaw = matrix[headerIdx] ?? [];
  const header = headerRaw.map((h, i) => {
    const s = h == null ? "" : String(h).trim();
    return s || `列${i + 1}`;
  });

  let dataRows = matrix.slice(headerIdx + 1).filter((r) => !isEmptyRow(r));
  if (typeof maxDataRows === "number" && maxDataRows > 0) {
    dataRows = dataRows.slice(0, maxDataRows);
  }

  const json = dataRows.map((r) => {
    const obj = {};
    for (let i = 0; i < header.length; i++) obj[header[i]] = r?.[i] ?? null;
    return obj;
  });
  return { rows: json, cols: header };
}

function getPeriodData(yyyymm) {
  const key = String(yyyymm);
  return state.periodCache.get(key) || { rows: [], cols: [], mapping: state.mapping, fileName: "" };
}

function loadPeriod(yyyymm) {
  const data = getPeriodData(yyyymm);
  state.rows = data.rows;
  state.cols = data.cols;
  state.mapping = data.mapping || guessMapping(state.cols);
  // 如果日期列未能通过列名识别，尝试根据值自动推断（避免“日期=（未识别）导致无数据”）
  const target = parseNameYYYYMM(String(yyyymm));
  const needInfer =
    !state.mapping?.dateCol ||
    (target &&
      (() => {
        const res = colMatchesTargetYYYYMM(state.rows, state.mapping.dateCol, target.year, target.month);
        // 如果大多数解析不命中目标年月，就认为日期列推断错误并重新推断
        if (!res.seen) return true;
        const bad = res.ratio < 0.2 && res.match < 3;
        return bad;
      })());
  if (needInfer) {
    const inferred = inferDateColFromValues(state.rows, state.cols, target?.year, target?.month);
    if (inferred) state.mapping.dateCol = inferred;
  }
  setMappingHint(state.mapping);

  // build dept options
  const m = state.mapping;
  const d1s = uniq(state.rows.map((r) => (r?.[m.dept1Col] ?? "").toString().trim()).filter(Boolean)).sort(sortByDeptCode);
  const d2s = uniq(state.rows.map((r) => (r?.[m.dept2Col] ?? "").toString().trim()).filter(Boolean)).sort(sortByDeptCode);
  $("dept1Select").innerHTML = "";
  for (const v of d1s) {
    const o = document.createElement("option");
    o.value = v;
    o.textContent = v;
    $("dept1Select").appendChild(o);
  }
  $("dept2Select").innerHTML = "";
  for (const v of d2s) {
    const o = document.createElement("option");
    o.value = v;
    o.textContent = v;
    $("dept2Select").appendChild(o);
  }

  // years: 优先从 sheet 名（YYYYMM）推断，再退化为日期列推断；固定近三年
  const metaMax = state.periodMeta?.maxYear ?? new Date().getFullYear();
  const fallbackYears = uniq(
    state.rows
      .map((r) => toDate(r?.[m.dateCol]))
      .filter(Boolean)
      .map((d) => d.getFullYear())
  ).sort((a, b) => a - b);
  const maxYear = metaMax || (fallbackYears.length ? fallbackYears[fallbackYears.length - 1] : new Date().getFullYear());
  const near3 = [maxYear - 2, maxYear - 1, maxYear].filter((y) => y > 1900);
  fillSelect($("yearSelect"), near3.map(String), String(maxYear));
  state.uiDefaults.year = String(maxYear);

  // months 1..12
  fillSelect(
    $("monthSelect"),
    Array.from({ length: 12 }, (_, i) => String(i + 1)),
    String(new Date().getMonth() + 1)
  );
  state.uiDefaults.month = String(new Date().getMonth() + 1);

  $("btnApply").disabled = false;
  $("btnReset").disabled = false;
  applyFiltersAndRender();
}

async function parseFiles(fileList) {
  const all = Array.from(fileList || []);
  const hintEl = $("mappingHint");
  if (hintEl) {
    const sample = all.slice(0, 8).map((f) => f.name).join(", ");
    hintEl.textContent = `已选择 ${all.length} 个条目：${sample || "（无）"}`;
  }

  const files = all.filter((f) => /\.(xlsx|xls|xlsm)$/i.test(f.name));
  state.files = files;
  state.periodCache = new Map();

  if (!files.length) {
    const el = $("mappingHint");
    if (el) el.textContent = "未选择到 Excel 文件条目（请在文件选择器中选择包含 .xlsx/.xls 的文件夹）。";
    $("btnApply").disabled = true;
    return;
  }

  for (const f of files) {
    const p = parseFileNameYYYYMM(f.name);
    if (!p) continue;
    const ab = await f.arrayBuffer();
    const wb = XLSX.read(ab, { type: "array", cellDates: true });
    let best = { score: -1, parsed: null, mapping: null, sheetName: null };
    const sheets = wb.SheetNames || [];

    for (const sheetName of sheets) {
      const ws = wb.Sheets[sheetName];
      if (!ws) continue;

      // 先用采样数据解析表头/列，打分选最像的数据 sheet
      const sampleParsed = parseSheetToJson(ws, { maxDataRows: 200 });
      const mapping = guessMapping(sampleParsed.cols);
      const fields = [
        mapping.dateCol,
        mapping.dept1Col,
        mapping.dept2Col,
        mapping.empCol,
        mapping.billCol,
        mapping.stdCol,
        mapping.otCol,
        mapping.costCol,
        mapping.revCol,
      ];
      const score = fields.filter(Boolean).length + (mapping.billCol && mapping.stdCol ? 3 : 0);

      if (score > best.score) {
        best = { score, parsed: sampleParsed, mapping, sheetName };
      }
    }

    if (!best.parsed || !best.mapping) continue;

    // 选中最优 sheet 后再完整解析一次（保证计算用到全部数据）
    const wsBest = best.sheetName ? wb.Sheets[best.sheetName] : null;
    if (!wsBest) continue;
    const parsedFull = parseSheetToJson(wsBest, { maxDataRows: null });
    const mappingFull = guessMapping(parsedFull.cols);
    state.periodCache.set(p.yyyymm, { ...parsedFull, mapping: mappingFull, fileName: f.name });
  }

  state.periodMeta = computePeriodMeta(Array.from(state.periodCache.keys()));

  if (!state.periodMeta.yyyymm || !state.periodMeta.yyyymm.size) {
    const el = $("mappingHint");
    if (el) el.textContent = "未识别到可用数据：请确保文件名包含 YYYYMM（如 202601）。";
    $("btnApply").disabled = true;
    return;
  }

  // UI 提示：显示识别到的月份
  const periods = Array.from(state.periodMeta.yyyymm).sort();
  const el = $("mappingHint");
  if (el) el.textContent = `已识别 ${periods.length} 个月份：${periods.join(", ")}`;

  $("btnReload").disabled = false;
  $("btnApply").disabled = true;
  populateControls();
}

async function init() {
  // theme
  const savedTheme = (() => {
    try {
      return localStorage.getItem(THEME_STORAGE_KEY);
    } catch {
      return null;
    }
  })();
  applyTheme(savedTheme === "light" ? "light" : "dark");
  $("btnThemeDark")?.addEventListener("click", () => applyTheme("dark"));
  $("btnThemeLight")?.addEventListener("click", () => applyTheme("light"));

  // view-only mode: 通过链接参数进入
  const view = new URLSearchParams(location.search).get("view") === "1";
  if (view) document.documentElement.dataset.view = "1";

  if (view) {
    const params = new URLSearchParams(location.search);
    let enc = getShareEncodedFromUrl();
    if (!enc) {
      // 支持从线上快照文件加载：?view=1&snapUrl=https://.../latest.txt
      const snapUrl = params.get("snapUrl");
      if (snapUrl) {
        try {
          const resp = await fetch(snapUrl, { cache: "no-store" });
          const text = await resp.text();
          enc = text?.trim();
        } catch {
          enc = null;
        }
      }
    }
    if (!enc) {
      const el = $("mappingHint");
      if (el) el.textContent = "该分享链接缺少快照数据（share=... 或 snapUrl=...）。";
      return;
    }

    const snap = await decodeShareData(enc);
    if (!snap || snap.version !== SHARE_STORAGE_VERSION) {
      const el = $("mappingHint");
      if (el) el.textContent = "分享快照数据无法解析。";
      return;
    }
    // 兼容快照格式：
    // 1) 全局紧凑 g+p（推荐，体积最小）
    // 2) c: [[yyyymm,{d,r}], ...]（每 period 各有一份字典）
    // 3) periodCache: 老格式
    let cacheEntries = [];
    if (Array.isArray(snap.g) && Array.isArray(snap.p)) {
      cacheEntries = decodeGlobalCompactSnapshot(snap);
    } else if (Array.isArray(snap.c)) {
      cacheEntries = snap.c.map((x) => [x[0], { rows: decodeCompactPeriodRows(x[1]) }]);
    } else {
      cacheEntries = (snap.periodCache || []).map((x) => [x[0], x[1]]);
    }
    state.periodCache = new Map(cacheEntries);
    const keys = Array.from(state.periodCache.keys());
    state.periodMeta = computePeriodMeta(keys);
    state.mapping = snap.mapping || state.periodCache.get(keys[0])?.mapping || null;
    // 保证每个 periodCache 都有 mapping 引用（避免个别 period 缺失）
    for (const k of keys) {
      const v = state.periodCache.get(k);
      if (v && !v.mapping) v.mapping = state.mapping;
      state.periodCache.set(k, v);
    }
    state.titleState = snap.titleState || state.titleState;

    // 载入好快照后，正常渲染
    refreshYearMonthOptions();
    applyTitles();
    applyFiltersAndRender();
  }

  if (!view) {
    // 分享按钮（生成“只读查看链接”，快照数据内嵌到链接中）
    const shareBtn = $("btnShare");
    if (shareBtn) {
      shareBtn.addEventListener("click", async () => {
        if (!state.mapping) {
          alert("尚未完成字段映射，请先上传并解析 Excel。");
          return;
        }

        const allKeys = Array.from(state.periodCache.keys());
        if (!allKeys.length) {
          alert("请先上传并解析 Excel 数据后再分享。");
          return;
        }

        // 共享快照：每期间按员工聚合 + 全局字符串字典 g，体积极小；链接长度不再强制 700（该限制对真实数据不可行）。
        const m0 = state.mapping;
        const dailyOt = isDailyOtCol(m0);

        const shareMapping = {
          dateCol: "__date",
          dept1Col: "d1",
          dept2Col: "d2",
          empCol: "emp",
          billCol: "bill",
          stdCol: "std",
          otCol: "ot",
          costCol: "cost",
          revCol: "rev",
          __dailyOt: dailyOt,
        };

        const collectPeriodsForUI = () => {
          const year = Number($("yearSelect")?.value);
          const month = Number($("monthSelect")?.value);
          const mode = document.querySelector('input[name="monthMode"]:checked')?.value ?? "month";
          const names = [];
          if (!Number.isFinite(year) || !Number.isFinite(month)) return [];
          if (mode === "month") {
            names.push(toYYYYMM(year, month));
          } else if (mode === "ytd") {
            for (let mm = 1; mm <= month; mm++) names.push(toYYYYMM(year, mm));
          } else {
            const maxM = state.periodMeta.maxMonthByYear.get(year) ?? 0;
            for (let mm = 1; mm <= maxM; mm++) names.push(toYYYYMM(year, mm));
          }
          return names.filter((p) => state.periodMeta.yyyymm.has(p));
        };

        const buildSnapForKeys = async (periodKeys) => {
          const filtered = (periodKeys || []).filter(Boolean);
          const { g, p } = encodeGlobalCompactSnapshot(filtered, m0, dailyOt);
          return {
            version: SHARE_STORAGE_VERSION,
            mapping: shareMapping,
            g,
            p,
          };
        };

        const base = String(location.href).split("#")[0];
        const join = base.includes("?") ? (base.endsWith("?") || base.endsWith("&") ? "" : "&") : "?";
        const viewBase = `${base}${join}view=1`;

        // 700 字符无法承载“多月份 × 多名员工”的 gzip+base64 快照（信息论上不可行）。
        // 使用 hash 承载数据时，现代浏览器通常可支持较长片段；仍设上限以防极端情况。
        const MAX_SHARE_URL = 150000;

        const currentMonthOnlyKeys = () => {
          const y = Number($("yearSelect")?.value);
          const mo = Number($("monthSelect")?.value);
          if (!Number.isFinite(y) || !Number.isFinite(mo)) return [];
          const p = toYYYYMM(y, mo);
          return state.periodMeta.yyyymm.has(p) ? [p] : [];
        };

        let shareUrl = "";
        const tryKeySets = [allKeys, collectPeriodsForUI(), currentMonthOnlyKeys()].filter(
          (k) => k && k.length
        );
        for (const periodKeys of tryKeySets) {
          const snap = await buildSnapForKeys(periodKeys);
          const encoded = await encodeShareData(snap);
          const u = `${viewBase}#${SHARE_PARAM_KEY}=${encoded}`;
          if (u.length <= MAX_SHARE_URL) {
            shareUrl = u;
            break;
          }
        }

        if (!shareUrl) {
          alert(
            "数据量过大，无法在浏览器中生成分享链接（已超过安全长度上限）。请减少上传月份或精简数据后重试。"
          );
          return;
        }

        // 展示二维码，便于对方扫码查看
        showShareQrModal(shareUrl);

        try {
          await navigator.clipboard.writeText(shareUrl);
          alert("分享链接已复制到剪贴板。");
        } catch {
          window.prompt("复制分享链接：", shareUrl);
        }
      });
    }

    const fileInput = $("fileInputFiles");
    if (fileInput) {
      fileInput.addEventListener("change", async (e) => {
        const files = e.target.files;
        if (!files || !files.length) return;
        await parseFiles(files);
      });
    }

    $("btnReload").addEventListener("click", async () => {
      if (!state.files || !state.files.length) return;
      await parseFiles(state.files);
    });
  }

  // QR modal actions
  const closeBtn = $("btnCloseShareQr");
  if (closeBtn) {
    closeBtn.addEventListener("click", () => {
      const modal = $("shareQrModal");
      if (modal) modal.style.display = "none";
    });
  }
  const copyBtn = $("btnCopyShareQrUrl");
  if (copyBtn) {
    copyBtn.addEventListener("click", async () => {
      const input = $("shareQrUrlInput");
      const url = input?.value || "";
      if (!url) return;
      try {
        await navigator.clipboard.writeText(url);
        alert("链接已复制。");
      } catch {
        window.prompt("复制分享链接：", url);
      }
    });
  }

  // 筛选器交互：不按 Ctrl 时，点一次选中；再点同一项取消（清空该筛选）
  // 按住 Ctrl 时保持浏览器原生多选行为。
  let suppressDept1Change = false;
  let suppressDept2Change = false;

  const setupToggleNoCtrl = (selEl, setSuppress) => {
    if (!selEl) return;
    selEl.addEventListener("mousedown", (e) => {
      const opt = e.target;
      if (!opt || opt.tagName !== "OPTION") return;
      if (e.ctrlKey || e.metaKey) return; // Ctrl/Mac Command：允许原生行为
      e.preventDefault();

      const val = opt.value;
      const already = Array.from(selEl.options).some((o) => o.value === val && o.selected);
      Array.from(selEl.options).forEach((o) => (o.selected = false));
      if (!already) opt.selected = true;

      setSuppress(true);
      applyFiltersAndRender();
    });
  };

  setupToggleNoCtrl($("dept1Select"), (v) => {
    suppressDept1Change = v;
    setTimeout(() => (suppressDept1Change = false), 0);
  });
  setupToggleNoCtrl($("dept2Select"), (v) => {
    suppressDept2Change = v;
    setTimeout(() => (suppressDept2Change = false), 0);
  });

  $("dept1Select").addEventListener("change", () => {
    if (suppressDept1Change) return;
    applyFiltersAndRender();
  });
  $("dept2Select").addEventListener("change", () => {
    if (suppressDept2Change) return;
    applyFiltersAndRender();
  });
  $("yearSelect").addEventListener("change", () => {
    refreshYearMonthOptions();
    applyFiltersAndRender();
  });
  $("monthSelect").addEventListener("change", () => applyFiltersAndRender());
  document.querySelectorAll('input[name="monthMode"]').forEach((el) => el.addEventListener("change", () => applyFiltersAndRender()));

  $("btnApply").addEventListener("click", () => applyFiltersAndRender());

  $("btnReset").addEventListener("click", () => {
    // 清空部门筛选
    Array.from($("dept1Select").options).forEach((o) => (o.selected = false));
    Array.from($("dept2Select").options).forEach((o) => (o.selected = false));
    // 年/月恢复默认
    if (state.uiDefaults.year) $("yearSelect").value = state.uiDefaults.year;
    if (state.uiDefaults.month) $("monthSelect").value = state.uiDefaults.month;
    // 模式恢复当月
    const r = document.querySelector('input[name="monthMode"][value="month"]');
    if (r) r.checked = true;
    // 清空排序
    state.tableSort.key = null;
    state.tableSort.dir = null;
    applyFiltersAndRender();
  });

  const alertEmpSel = $("alertEmpFilter");
  const alertD1Sel = $("alertDept1Filter");
  const alertD2Sel = $("alertDept2Filter");
  let suppressAlertChange = false;
  const setupToggleSingleNoCtrl = (selEl) => {
    if (!selEl) return;
    selEl.addEventListener("change", () => {
      if (suppressAlertChange) return;
      state.alertView.page = 1;
      renderAlerts(state.latestAlerts || []);
    });

    selEl.addEventListener("mousedown", (e) => {
      const opt = e.target;
      if (!opt || opt.tagName !== "OPTION") return;
      if (e.ctrlKey || e.metaKey) return; // Ctrl：不做“二次点击取消”的交互
      e.preventDefault();

      const clickedVal = opt.value;
      const curVal = selEl.value || "";
      const nextVal = clickedVal === curVal ? "" : clickedVal;

      suppressAlertChange = true;
      selEl.value = nextVal;
      state.alertView.page = 1;
      renderAlerts(state.latestAlerts || []);
      setTimeout(() => (suppressAlertChange = false), 0);
    });
  };

  setupToggleSingleNoCtrl(alertEmpSel);
  setupToggleSingleNoCtrl(alertD1Sel);
  setupToggleSingleNoCtrl(alertD2Sel);

  const pageSizeInput = $("alertPageSize");
  if (pageSizeInput) {
    pageSizeInput.addEventListener("change", () => {
      const v = Number(pageSizeInput.value);
      if (!Number.isFinite(v) || v <= 0) {
        state.alertView.pageSize = 5;
        pageSizeInput.value = "5";
      } else {
        state.alertView.pageSize = Math.min(20, Math.max(1, Math.round(v)));
        pageSizeInput.value = String(state.alertView.pageSize);
      }
      state.alertView.page = 1;
      renderAlerts(state.latestAlerts || []);
    });
  }

  const prevBtn = $("alertPrev");
  const nextBtn = $("alertNext");
  if (prevBtn) {
    prevBtn.addEventListener("click", () => {
      if (state.alertView.page > 1) {
        state.alertView.page -= 1;
        renderAlerts(state.latestAlerts || []);
      }
    });
  }
  if (nextBtn) {
    nextBtn.addEventListener("click", () => {
      const flat = state.latestAlerts || [];
      const pageSize = Math.max(1, Math.min(20, state.alertView.pageSize || 5));
      const totalPages = Math.max(1, Math.ceil(flat.length / pageSize));
      if (state.alertView.page < totalPages) {
        state.alertView.page += 1;
        renderAlerts(state.latestAlerts || []);
      }
    });
  }

  // 折叠/展开：允许用户手动折叠各模块
  const collapseState = loadCollapseState();
  applyCollapseState(collapseState);
  document.querySelectorAll(".collapse-toggle").forEach((btn) => {
    btn.addEventListener("click", (e) => {
      e.preventDefault();
      e.stopPropagation();
      const key = btn?.dataset?.collapseKey;
      const card = btn?.closest?.(".card");
      if (!key || !card) return;
      const collapsed = !card.classList.contains("collapsed");
      card.classList.toggle("collapsed", collapsed);
      btn.textContent = collapsed ? "▸" : "▾";
      collapseState[key] = collapsed;
      saveCollapseState(collapseState);
    });
  });

  if (!view) {
    // 标题自定义：加载并应用初始值
    loadTitles();
    applyTitles();

    // 编辑按钮
    document.querySelectorAll(".title-edit").forEach((btn) => {
      btn.addEventListener("click", () => {
        const key = btn.dataset.titleKey;
        if (!key) return;
        if (!state.titleState) loadTitles();
        const current = state.titleState?.[key] || defaultTitles[key] || "";
        // 简单使用浏览器 prompt 进行编辑
        const next = window.prompt("请输入新的标题：", current);
        if (next == null) return;
        const value = next.trim();
        if (!value) return;
        state.titleState[key] = value;
        applyTitles();
        saveTitles();
      });
    });

    // 重置标题按钮
    const resetTitlesBtn = $("btnResetTitles");
    if (resetTitlesBtn) {
      resetTitlesBtn.addEventListener("click", () => {
        if (!window.confirm("确定要将所有标题恢复为默认吗？")) return;
        state.titleState = { ...defaultTitles };
        try {
          localStorage.removeItem(TITLE_STORAGE_KEY);
        } catch {
          // ignore
        }
        applyTitles();
      });
    }

    // hint: try auto-load if user drags file onto page
    document.addEventListener("dragover", (e) => e.preventDefault());
    document.addEventListener("drop", async (e) => {
      e.preventDefault();
      const file = e.dataTransfer?.files?.[0];
      if (!file) return;
      if (!/\.xlsx?$/.test(file.name.toLowerCase())) return;
      await parseFile(file);
    });
  }
}

init();

