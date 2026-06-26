/* S低L3-3 数据处理工具 - app.js */
(function () {
  "use strict";

  /* ========== 常量 ========== */
  const CLASS_KEEP = ["班级ID", "课导ID", "课导花名", "学期ID", "课导组名", "课导组长名称", "学期名称", "班级名称"];
  const USER_KEEP = ["用户ID", "孩子姓名", "学期ID", "班级ID", "微信名称"]; // 开课日衍生
  const USER_DEL = ["用户真实姓名", "微信昵称"];
  const USER_HIDE = ["学期ID"]; // 隐藏该列（导出 Excel 时保留但不展示在预览）
  const CACHE_KEY = "s_low_l3_cache_v1";

  /* ========== 状态 ========== */
  const state = {
    raw: { class: null, user: null, "activity-class": null, "activity-user": null },
    fileName: { class: "", user: "", "activity-class": "", "activity-user": "" },
    sheets: {},     // 生成后的 sheet -> { headers:[], rows:[[]] }
    workbook: null, // 生成后的 workbook 供下载
    fileNameBase: "",
    activityClass: null, // 原始 activity-class 处理结果，供看板筛选重算
    selectedDays: [],    // 看板筛选选中的开课日（字符串数组）
    dashboardName: "",
    lastMode: "",    // "base" | "full"
  };

  /* ========== 工具函数 ========== */
  function toast(msg) {
    const el = document.getElementById("toast");
    el.textContent = msg;
    el.classList.add("show");
    clearTimeout(toast._t);
    toast._t = setTimeout(() => el.classList.remove("show"), 2000);
  }

  function pickField(row, names) {
    const keys = Object.keys(row);
    for (const name of names) {
      // 精确 or 模糊匹配，处理列名带空格/换行
      const hit = keys.find((k) => k.replace(/\s|　/g, "").includes(name.replace(/\s|　/g, "")));
      if (hit) return row[hit];
    }
    return undefined;
  }

  function normalizeField(key) {
    return String(key || "").replace(/\s|　|\n|\r/g, "");
  }

  function readFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const wb = XLSX.read(data, { type: "array" });
          const sheetName = wb.SheetNames[0];
          const ws = wb.Sheets[sheetName];
          const rows = XLSX.utils.sheet_to_json(ws, { defval: "", raw: false });
          resolve(rows);
        } catch (err) {
          reject(err);
        }
      };
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    });
  }

  function extractWeekday(className) {
    if (!className) return "";
    const s = String(className);
    // 尝试匹配 "周X" 或 "周XX" 或 "星期X"
    const m1 = s.match(/周[一二三四五六日天一二三四五六日]/);
    if (m1) {
      const ch = m1[0].charAt(1);
      if (ch === "天" || ch === "日") return "周日";
      return "周" + ch;
    }
    const m2 = s.match(/星期[一二三四五六日天]/);
    if (m2) {
      const ch = m2[0].charAt(2);
      if (ch === "天" || ch === "日") return "周日";
      return "周" + ch;
    }
    // 回退：取前两个字符
    return s.slice(0, 2);
  }

  // 将任意格式的开课日规范化为 "周x"（如 星期四/周4/周四 => 周四）
  function normalizeOpenDay(v) {
    if (v === undefined || v === null) return "";
    const s = String(v).trim();
    if (!s) return "";
    // 处理 "周4" / "星期4" 这种数字写法
    let m = s.match(/星期?\s*(\d)/);
    if (m) {
      const map = { "1": "周一", "2": "周二", "3": "周三", "4": "周四", "5": "周五", "6": "周六", "7": "周日", "0": "周日" };
      return map[m[1]] || s;
    }
    // 处理 星期x / 周x 中文写法
    m = s.match(/星期([一二三四五六日天])/);
    if (m) return "周" + (m[1] === "天" ? "日" : m[1]);
    m = s.match(/^周([一二三四五六日天])$/);
    if (m) return "周" + (m[1] === "天" ? "日" : m[1]);
    return s;
  }

  // 从 activityClass 中收集去重开课日（规范化后），按 周一..周日 的顺序返回
  function collectOpenDays(activityClass) {
    if (!activityClass || !activityClass.rows || !activityClass.headers) return [];
    const idx = activityClass.headers.findIndex((h) => normalizeField(h) === "开课日");
    if (idx < 0) return [];
    const set = new Set();
    for (const r of activityClass.rows) {
      const d = normalizeOpenDay(r[idx]);
      if (d) set.add(d);
    }
    const order = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"];
    return Array.from(set).sort((a, b) => {
      const ia = order.indexOf(a);
      const ib = order.indexOf(b);
      return (ia === -1 ? 99 : ia) - (ib === -1 ? 99 : ib);
    });
  }

  // 根据所选开课日自动生成看板名称：
  //   1 个 => 周四看板
  //   N 个 => 周四&周五&周六看板
  //   0 个 => 全部看板
  function autoDashboardName(openDays) {
    if (!openDays || openDays.length === 0) return "全部看板";
    if (openDays.length === 1) return openDays[0] + "看板";
    return openDays.join("&") + "看板";
  }
  function toNum(v) {
    if (v === null || v === undefined || v === "") return 0;
    const n = Number(String(v).replace(/,/g, "").replace(/%$/, ""));
    return isNaN(n) ? 0 : n;
  }

  function emptyToZero(v) {
    return v === "" || v === null || v === undefined ? 0 : v;
  }

  // 3.1.4 需求: 根据好友体验UV / 是否分享 计算核桃币
  // UV>=10 -> 4000;  UV>=3且<10 -> 200; UV<3且分享=1 -> 50; 其余为空
  function calcCoin(uv, shared) {
    const n = toNum(uv);
    if (n >= 10) return 4000;
    if (n >= 3) return 200;
    if (toNum(shared) >= 1) return 50;
    return "";
  }

  // 核桃币数值（用于排序降序，空值=0）
  function coinSortVal(v) {
    if (v === "" || v === null || v === undefined) return 0;
    const n = Number(String(v).replace(/,/g, ""));
    return isNaN(n) ? 0 : n;
  }

  /* ========== 学期名称缩写 ========== */
  // 示例: "2026-春季图形化编程第3期-1周2课专用（乐高测试）"
  // -> "S低思维-26春3期" / 底表: "S低思维-26春3期底表"
  function shortenTerm(term) {
    if (!term) return "";
    const s = String(term);
    const yearMatch = s.match(/(\d{4})/);
    const year = yearMatch ? yearMatch[1].slice(-2) : "";
    // 季节
    let season = "";
    if (s.includes("春")) season = "春";
    else if (s.includes("夏")) season = "夏";
    else if (s.includes("秋")) season = "秋";
    else if (s.includes("冬")) season = "冬";
    // 第N期
    const periodMatch = s.match(/第\s*(\d+)\s*期/);
    const period = periodMatch ? periodMatch[1] : "";
    if (!year && !season && !period) return s.slice(0, 20);
    return `S低思维-${year}${season}${period}期`;
  }

  /* ========== 数据处理：基础班级 ========== */
  function processBaseClass(rows) {
    const headers = CLASS_KEEP.slice();
    headers.push("开课日");
    const out = [];
    const classIds = [];
    for (const r of rows) {
      const row = [];
      for (const f of CLASS_KEEP) row.push(pickField(r, [f]) ?? "");
      const classId = row[0];
      if (classId !== "" && classId !== null && classId !== undefined) classIds.push(String(classId));
      const className = row[7];
      row.push(extractWeekday(className));
      out.push(row);
    }
    return { headers, rows: out, classIds };
  }

  /* ========== 数据处理：基础用户 ========== */
  function processBaseUser(rows, baseClassLookup) {
    const headers = USER_KEEP.slice();
    headers.push("开课日");
    const out = [];
    for (const r of rows) {
      const row = [];
      for (const f of USER_KEEP) row.push(pickField(r, [f]) ?? "");
      const classId = row[3];
      row.push(baseClassLookup[String(classId)] || "");
      out.push(row);
    }
    return { headers, rows: out };
  }

  /* ========== 数据处理：活动班级（班级信息） ========== */
  function processActivityClass(rows, originalHeaders, baseClassLookup) {
    const headers = originalHeaders.slice();
    if (!headers.includes("开课日")) headers.push("开课日");
    const out = [];
    for (const r of rows) {
      const row = [];
      for (const f of originalHeaders) row.push(r[f] !== undefined ? r[f] : "");
      const classIdRaw = pickField(r, ["班级ID"]) ?? "";
      const classId = String(classIdRaw);
      // 保证开课日位置在最后
      if (!originalHeaders.includes("开课日")) {
        row.push(baseClassLookup[classId] || "");
      } else {
        const idx = originalHeaders.indexOf("开课日");
        row[idx] = baseClassLookup[classId] || row[idx];
      }
      out.push(row);
    }
    return { headers, rows: out };
  }

  /* ========== 数据处理：活动用户（用户明细） ========== */
  function processActivityUser(rows, originalHeaders, baseUserLookup, baseClassLookup) {
    // 剔除 用户真实姓名、微信昵称
    let headers = originalHeaders.filter((h) => {
      const nh = normalizeField(h);
      if (USER_DEL.some((d) => nh === normalizeField(d))) return false;
      return true;
    });

    const out = [];
    for (const r of rows) {
      const row = [];
      let uvVal = 0; // 记录空白填充后的 好友体验UV
      let sharedVal = 0;
      let hasSharedCol = false;
      for (const f of headers) {
        let v = r[f] !== undefined ? r[f] : "";
        const nh = normalizeField(f);
        // 严格按「好友体验UV」精确匹配；空值填 0
        if (nh === "好友体验UV") {
          if (v === "" || v === null || v === undefined) v = 0;
          uvVal = toNum(v);
        }
        // 严格按「是否分享」精确匹配
        if (nh === "是否分享") {
          hasSharedCol = true;
          sharedVal = toNum(v);
        }
        row.push(v);
      }
      // 衍生字段
      const userId = String(pickField(r, ["用户ID"]) ?? "");
      const classId = String(pickField(r, ["班级ID"]) ?? "");
      const lastShare = pickField(r, ["末次分享时间"]);
      const childName = baseUserLookup[userId] || "";
      const openDay = baseClassLookup[classId] || "";
      // 若原数据没有「是否分享」列，则退回到用末次分享时间判断
      if (!hasSharedCol) {
        sharedVal = (lastShare === "" || lastShare === null || lastShare === undefined) ? 0 : 1;
      }
      const coin = calcCoin(uvVal, sharedVal);
      row.push(childName, openDay, sharedVal, coin);
      out.push(row);
    }

    // 新增 headers
    if (!headers.includes("孩子姓名")) headers.push("孩子姓名");
    if (!headers.includes("开课日")) headers.push("开课日");
    if (!headers.includes("是否分享")) headers.push("是否分享");
    if (!headers.includes("核桃币")) headers.push("核桃币");

    // 排序: 按核桃币数值降序（3.1.4）
    const coinIdx = headers.indexOf("核桃币");
    if (coinIdx >= 0) {
      out.sort((a, b) => coinSortVal(b[coinIdx]) - coinSortVal(a[coinIdx]));
    }

    return { headers, rows: out };
  }

  /* ========== 3.2.4 班级信息透视 -> 周x看板 ========== */
  // selectedDays: 规范后的开课日数组，如 ["周四","周五"]。
  //   - 空数组 / null => 不过滤（等价于"全部"）
  // dashboardName: 可为空，空时根据 selectedDays 自动生成
  // 返回: { title, subtitle, headers, rows, autoName }
  function buildDashboard(activityClass, selectedDays, dashboardName) {
    const src = activityClass.rows;
    const rawHeaders = activityClass.headers;
    const norm = rawHeaders.map((h) => normalizeField(h));
    const useFilter = Array.isArray(selectedDays) && selectedDays.length > 0;
    const filterSet = new Set(selectedDays || []);
    const dayIdx = norm.indexOf("开课日");

    function findCol(primary, alts) {
      let i = norm.indexOf(primary);
      if (i >= 0) return i;
      if (alts) {
        for (const a of alts) {
          i = norm.findIndex((h) => h.includes(a));
          if (i >= 0) return i;
        }
      }
      return -1;
    }

    const idxLeader = findCol("组长");
    const idxTeacher = findCol("老师花名", ["课导花名", "老师姓名", "老师"]);
    const idxInClass = findCol("在班人数", ["学生数", "班级人数"]);
    const idxFish = findCol("完课人数", ["完成课人数", "完课数"]);
    const idxTouch = findCol("触达人数");
    const idxEnter = findCol("进入活动人数", ["进入活动"]);
    const idxExp = findCol("体验作品人数", ["体验作品"]);
    const idxShare = findCol("分享作品人数", ["分享作品"]);

    const groups = new Map();
    for (let r = 0; r < src.length; r++) {
      const row = src[r];
      if (useFilter) {
        const rowDay = dayIdx >= 0 ? normalizeOpenDay(row[dayIdx]) : "";
        if (rowDay && !filterSet.has(rowDay)) continue;
      }
      const leader = (idxLeader >= 0 && row[idxLeader] !== undefined && row[idxLeader] !== null && String(row[idxLeader]).trim() !== "")
        ? String(row[idxLeader]).trim() : "(无组长)";
      const teacher = (idxTeacher >= 0 && row[idxTeacher] !== undefined && row[idxTeacher] !== null && String(row[idxTeacher]).trim() !== "")
        ? String(row[idxTeacher]).trim() : "(无老师)";
      const key = leader + "||" + teacher;
      let g = groups.get(key);
      if (!g) {
        g = { leader, teacher, inClass: 0, finish: 0, touch: 0, enter: 0, exp: 0, share: 0 };
        groups.set(key, g);
      }
      g.inClass += toNum(idxInClass >= 0 ? row[idxInClass] : 0);
      g.finish += toNum(idxFish >= 0 ? row[idxFish] : 0);
      g.touch += toNum(idxTouch >= 0 ? row[idxTouch] : 0);
      g.enter += toNum(idxEnter >= 0 ? row[idxEnter] : 0);
      g.exp += toNum(idxExp >= 0 ? row[idxExp] : 0);
      g.share += toNum(idxShare >= 0 ? row[idxShare] : 0);
    }

    // 排序：按组长的中文排序，相同组长按老师花名
    const sorted = Array.from(groups.values()).sort((a, b) => {
      const cl = String(a.leader).localeCompare(String(b.leader), "zh");
      if (cl !== 0) return cl;
      return String(a.teacher).localeCompare(String(b.teacher), "zh");
    });

    function pct(n, d) {
      if (!d || d === 0) return "0.00%";
      return (n / d * 100).toFixed(2) + "%";
    }

    const outHeaders = ["组长", "老师花名", "在班人数", "完课人数", "完课率",
      "触达人数", "完课触达率", "进入活动人数", "完课进入率",
      "体验作品人数", "完课体验率", "分享作品人数"];

    const dataRows = sorted.map((g) => [
      g.leader, g.teacher,
      g.inClass,
      g.finish,
      pct(g.finish, g.inClass),
      g.touch,
      pct(g.touch, g.finish),
      g.enter,
      pct(g.enter, g.finish),
      g.exp,
      pct(g.exp, g.finish),
      g.share,
    ]);

    // 总计行：所有数值求和后重新计算百分比
    const total = dataRows.reduce(
      (acc, r) => {
        acc.inClass += toNum(r[2]);
        acc.finish += toNum(r[3]);
        acc.touch += toNum(r[5]);
        acc.enter += toNum(r[7]);
        acc.exp += toNum(r[9]);
        acc.share += toNum(r[11]);
        return acc;
      },
      { inClass: 0, finish: 0, touch: 0, enter: 0, exp: 0, share: 0 }
    );

    const totalRow = [
      "总计", "",
      total.inClass,
      total.finish,
      pct(total.finish, total.inClass),
      total.touch,
      pct(total.touch, total.finish),
      total.enter,
      pct(total.enter, total.finish),
      total.exp,
      pct(total.exp, total.finish),
      total.share,
    ];

    const allRows = dataRows.concat([totalRow]);
    const autoName = autoDashboardName(selectedDays);
    return {
      title: "开课日",
      subtitle: dashboardName || autoName,
      headers: outHeaders,
      rows: allRows,
      autoName: autoName,
    };
  }

  /* ========== 3.2.1 获奖名单: 按 10 个指定字段重构造 ========== */
  function buildAwardList(activityUser) {
    const src = activityUser.headers;
    // 目标顺序（文档 3.1.4 / 3.2.1 要求：以 "好友体验UV" 为准）
    const targets = [
      "用户ID",
      "班级ID",
      "老师ID",
      "老师花名",
      "好友体验UV",
      "奖项",
      "孩子姓名",
      "开课日",
      "是否分享",
      "核桃币",
    ];
    // 在 src 中按规范化字段名查找
    const srcIdx = {};
    src.forEach((h, i) => { srcIdx[normalizeField(h)] = i; });

    function lookup(name) {
      const key = normalizeField(name);
      // 精确匹配
      if (srcIdx[key] !== undefined) return srcIdx[key];
      // 部分匹配（例如 "奖项/获奖/老师ID/老师花名/好友体验UV/新用户好友体验UV"）
      for (const k of Object.keys(srcIdx)) {
        if (k.includes(key) || key.includes(k)) return srcIdx[k];
      }
      return -1;
    }

    // 对 UV 字段做降级: 先按"好友体验UV"精确，再用包含"好友体验UV"的列（如"新用户好友体验UV"）
    let uvIdx = lookup("好友体验UV");
    if (uvIdx < 0) {
      const altKey = Object.keys(srcIdx).find((k) => k.includes("好友体验UV") && k !== normalizeField("好友体验UV"));
      if (altKey !== undefined) uvIdx = srcIdx[altKey];
    }

    const mapping = [
      lookup("用户ID"),
      lookup("班级ID"),
      lookup("老师ID"),
      lookup("老师花名"),
      uvIdx,
      lookup("奖项"),
      lookup("孩子姓名"),
      lookup("开课日"),
      lookup("是否分享"),
      lookup("核桃币"),
    ];

    const awardSrc = lookup("奖项");

    const filteredRows = [];
    for (const r of activityUser.rows) {
      if (awardSrc >= 0) {
        const v = String(r[awardSrc] || "");
        if (v === "" || v.includes("未获奖")) continue;
      }
      const row = targets.map((_, i) => {
        const idx = mapping[i];
        if (idx < 0) return "";
        const val = r[idx];
        return (val === undefined || val === null) ? "" : val;
      });
      filteredRows.push(row);
    }

    // 按核桃币降序（同 3.1.4 排序规则）
    const coinCol = targets.indexOf("核桃币");
    if (coinCol >= 0) {
      filteredRows.sort((a, b) => coinSortVal(b[coinCol]) - coinSortVal(a[coinCol]));
    }

    return { headers: targets.slice(), rows: filteredRows };
  }

  function filterFinishedExp(activityUser, wantExp) {
    const headers = activityUser.headers;
    const finishIdx = headers.findIndex((h) => normalizeField(h).includes("完课时间"));
    const expIdx = headers.findIndex((h) => normalizeField(h).includes("末次体验时间"));
    if (finishIdx < 0 || expIdx < 0) return { headers, rows: activityUser.rows.slice() };
    const rows = activityUser.rows.filter((r) => {
      const hasFinish = String(r[finishIdx] || "") !== "";
      const hasExp = String(r[expIdx] || "") !== "";
      if (!hasFinish) return false;
      return wantExp ? hasExp : !hasExp;
    });
    return { headers, rows };
  }

  /* ========== 渲染 ========== */
  function isDashboardTab(name) {
    // 简单判定: 包含"看板"字样的 tab 视为看板 tab
    return /看板/.test(name);
  }

  function renderSheetTabs(sheetNames) {
    const wrap = document.getElementById("sheetTabs");
    wrap.innerHTML = "";
    sheetNames.forEach((name, i) => {
      const row = document.createElement("div");
      row.className = "flex items-center gap-1";

      const b = document.createElement("button");
      b.className = "sheet-tab px-3 py-1.5 text-xs border border-slate-300 rounded-md bg-white hover:bg-slate-50";
      b.textContent = name;
      b.addEventListener("click", () => {
        wrap.querySelectorAll(".sheet-tab").forEach((x) => x.classList.remove("active"));
        b.classList.add("active");
        renderSheetTable(name);
      });
      row.appendChild(b);

      // 仅对「周x看板」tab 追加编辑按钮
      if (isDashboardTab(name)) {
        const editBtn = document.createElement("button");
        editBtn.className = "px-2 py-1 text-xs border border-indigo-300 text-indigo-700 rounded-md bg-indigo-50 hover:bg-indigo-100";
        editBtn.textContent = "✎ 编辑";
        editBtn.title = "编辑本 Sheet 名称";
        editBtn.addEventListener("click", () => renameDashboardTab(name, row));
        row.appendChild(editBtn);
      }

      wrap.appendChild(row);
    });
    if (sheetNames.length) {
      const firstTab = wrap.querySelector(".sheet-tab");
      if (firstTab) firstTab.classList.add("active");
      renderSheetTable(sheetNames[0]);
    }
  }

  function renameDashboardTab(oldName, rowEl) {
    const newName = prompt("请输入新的 Sheet 名称：", oldName);
    if (newName === null) return;
    const trimmed = String(newName).trim();
    if (!trimmed || trimmed === oldName) return;
    // 避免与已有 sheet 冲突
    if (state.sheets[trimmed] && trimmed !== oldName) {
      alert("已存在同名 Sheet，请换一个名称。");
      return;
    }
    // 1. 重命名 state.sheets key 并同步更新 subtitle
    const newSheets = {};
    Object.keys(state.sheets).forEach((k) => {
      if (k === oldName) {
        const s = state.sheets[k];
        newSheets[trimmed] = { ...s, subtitle: trimmed };
      } else {
        newSheets[k] = state.sheets[k];
      }
    });
    state.sheets = newSheets;
    // 2. 更新 state.dashboardName
    if (state.dashboardName === oldName) state.dashboardName = trimmed;
    // 3. 重渲染全部 tab
    renderSheetTabs(Object.keys(state.sheets));
    // 4. 选中新 tab
    const wrap = document.getElementById("sheetTabs");
    wrap.querySelectorAll(".sheet-tab").forEach((x) => {
      if (x.textContent === trimmed) x.classList.add("active");
      else x.classList.remove("active");
    });
    renderSheetTable(trimmed);
    toast("Sheet 名称已更新：" + trimmed);
  }

  // 用户修改筛选器时重算看板（同时重命名 tab/sheet）
  function recalcDashboard(newSelected) {
    if (!state.activityClass) return;
    const oldName = state.dashboardName;
    const filtered = Array.isArray(newSelected) ? newSelected.slice() : [];
    state.selectedDays = filtered;
    const dash = buildDashboard(state.activityClass, state.selectedDays);
    const newName = dash.autoName;
    // 重命名 state.sheets key（若旧名存在且不同）
    if (oldName && oldName !== newName && state.sheets[oldName]) {
      const newSheets = {};
      Object.keys(state.sheets).forEach((k) => {
        if (k === oldName) newSheets[newName] = dash;
        else newSheets[k] = state.sheets[k];
      });
      state.sheets = newSheets;
    } else if (state.sheets[oldName]) {
      state.sheets[oldName] = dash;
    } else {
      state.sheets[newName] = dash;
    }
    state.dashboardName = newName;
    renderSheetTabs(Object.keys(state.sheets));
    const wrap = document.getElementById("sheetTabs");
    wrap.querySelectorAll(".sheet-tab").forEach((x) => {
      if (x.textContent === newName) x.classList.add("active");
      else x.classList.remove("active");
    });
    renderSheetTable(newName);
  }

  function renderSheetTable(name) {
    const container = document.getElementById("sheetContainer");
    container.innerHTML = "";
    const sheet = state.sheets[name];
    if (!sheet) return;

    // 看板: 使用 title(开课日) + subtitle(看板名称) 的双行标题结构
    // 并提供开课日多选筛选器，切换时自动重算与重命名
    if (isDashboardTab(name)) {
      const wrap = document.createElement("div");
      wrap.id = "snapshot-" + name;
      wrap.style.display = "block";
      wrap.style.padding = "12px 16px";
      wrap.style.background = "#ffffff";
      wrap.style.minWidth = "100%";

      const title = document.createElement("div");
      title.textContent = sheet.title || "开课日";
      title.style.fontSize = "13px";
      title.style.color = "#374151";
      title.style.fontWeight = "600";
      title.style.padding = "2px 0";
      wrap.appendChild(title);

      const sub = document.createElement("div");
      sub.textContent = sheet.subtitle || name;
      sub.style.fontSize = "18px";
      sub.style.color = "#111827";
      sub.style.fontWeight = "700";
      sub.style.padding = "2px 0 10px";
      wrap.appendChild(sub);

      // 开课日筛选器区域
      const allDays = collectOpenDays(state.activityClass);
      if (allDays.length) {
        const filterWrap = document.createElement("div");
        filterWrap.style.padding = "8px 0 14px";
        filterWrap.style.display = "flex";
        filterWrap.style.flexWrap = "wrap";
        filterWrap.style.alignItems = "center";
        filterWrap.style.gap = "8px";
        filterWrap.style.borderBottom = "1px dashed #e5e7eb";
        filterWrap.style.marginBottom = "10px";

        const label = document.createElement("span");
        label.textContent = "开课日筛选：";
        label.style.fontSize = "13px";
        label.style.color = "#374151";
        label.style.fontWeight = "500";
        filterWrap.appendChild(label);

        // 全选 / 清空
        const allBtn = document.createElement("button");
        allBtn.className = "px-2 py-0.5 text-xs border border-slate-300 rounded bg-white hover:bg-slate-50";
        allBtn.textContent = "全选";
        allBtn.addEventListener("click", () => recalcDashboard(allDays.slice()));
        filterWrap.appendChild(allBtn);

        const noneBtn = document.createElement("button");
        noneBtn.className = "px-2 py-0.5 text-xs border border-slate-300 rounded bg-white hover:bg-slate-50";
        noneBtn.textContent = "清空";
        noneBtn.addEventListener("click", () => recalcDashboard([]));
        filterWrap.appendChild(noneBtn);

        // 多选复选框
        const current = new Set(state.selectedDays || []);
        allDays.forEach((d) => {
          const box = document.createElement("label");
          box.style.fontSize = "13px";
          box.style.color = "#111827";
          box.style.display = "inline-flex";
          box.style.alignItems = "center";
          box.style.gap = "4px";
          box.style.cursor = "pointer";
          box.style.userSelect = "none";
          const cb = document.createElement("input");
          cb.type = "checkbox";
          cb.checked = current.has(d);
          cb.addEventListener("change", () => {
            const next = new Set(state.selectedDays || []);
            if (cb.checked) next.add(d);
            else next.delete(d);
            recalcDashboard(Array.from(next));
          });
          box.appendChild(cb);
          const span = document.createElement("span");
          span.textContent = d;
          box.appendChild(span);
          filterWrap.appendChild(box);
        });
        wrap.appendChild(filterWrap);
      }

      wrap.appendChild(buildTable(name, sheet, true));
      container.appendChild(wrap);
      return;
    }

    if (name === "获奖名单") {
      const wrap = document.createElement("div");
      wrap.id = "snapshot-" + name;
      wrap.style.background = "#ffffff";
      wrap.style.padding = "12px 16px";
      wrap.style.minWidth = "100%";
      wrap.appendChild(buildTable(name, sheet, false));
      container.appendChild(wrap);
      return;
    }

    // 普通 sheet
    container.appendChild(buildTable(name, sheet, false));
  }

  function buildTable(sheetName, sheet, isDashboard) {
    const table = document.createElement("table");
    table.className = "preview-table";
    table.style.borderCollapse = "collapse";

    const thead = document.createElement("thead");
    const trh = document.createElement("tr");
    sheet.headers.forEach((h) => {
      // 仅「用户明细」sheet 下的学期ID列隐藏
      if (sheetName === "用户明细" && USER_HIDE.some((x) => x === h)) return;
      const th = document.createElement("th");
      th.textContent = h;
      th.style.background = "#f3f4f6";
      th.style.fontWeight = "600";
      th.style.border = "1px solid #d1d5db";
      th.style.padding = "6px 10px";
      th.style.whiteSpace = "nowrap";
      if (isDashboard) th.style.textAlign = "center";
      trh.appendChild(th);
    });
    thead.appendChild(trh);
    table.appendChild(thead);

    const tbody = document.createElement("tbody");

    // —— 看板：按列计算「完课进入率 / 完课体验率」的动态色阶
    //    参考图逻辑：按列数值的 min / mid / max 分三段
    //    高 → 绿色 (#d1fae5 / #065f46)
    //    中 → 黄色 (#fef3c7 / #92400e)
    //    低 → 红色 (#fee2e2 / #991b1b)
    const scales = {};
    if (isDashboard) {
      const dataRows = sheet.rows.filter((r) => !r.some((c) => String(c) === "总计"));
      ["完课进入率", "完课体验率"].forEach((h) => {
        const idx = sheet.headers.indexOf(h);
        if (idx < 0) return;
        const vals = dataRows.map((r) => toNum(r[idx])).filter((n) => !Number.isNaN(n));
        if (!vals.length) return;
        const lo = Math.min.apply(null, vals);
        const hi = Math.max.apply(null, vals);
        const mid = (lo + hi) / 2;
        scales[h] = { idx, lo, mid, hi };
      });
    }

    sheet.rows.forEach((row) => {
      const tr = document.createElement("tr");
      const isTotal = row.some((c) => String(c) === "总计");
      if (isTotal) tr.classList.add("total-row");
      row.forEach((cell, cIdx) => {
        const headerName = sheet.headers[cIdx];
        if (sheetName === "用户明细" && USER_HIDE.some((x) => x === headerName)) return;
        const td = document.createElement("td");
        td.textContent = cell === undefined || cell === null ? "" : cell;
        td.style.border = "1px solid #d1d5db";
        td.style.padding = "6px 10px";
        td.style.whiteSpace = "nowrap";
        if (isTotal) {
          td.style.background = "#fef3c7";
          td.style.fontWeight = "600";
        }
        if (isDashboard) {
          const hn = String(headerName || "");
          const scale = scales[hn];
          if (scale && !isTotal) {
            const v = toNum(cell);
            // 三色：高(绿) / 中(黄) / 低(红)
            if (v >= scale.mid) {
              // 上半段再按与 max 的距离分绿 vs 黄；若只有两级则最大值为绿，最小值为红
              if (scale.hi === scale.lo) { td.style.background = "#d1fae5"; td.style.color = "#065f46"; }
              else {
                const upperMid = (scale.mid + scale.hi) / 2;
                if (v >= upperMid) { td.style.background = "#d1fae5"; td.style.color = "#065f46"; }
                else { td.style.background = "#fef3c7"; td.style.color = "#92400e"; }
              }
            } else {
              td.style.background = "#fee2e2"; td.style.color = "#991b1b";
            }
            td.style.textAlign = "center";
          } else if (hn === "完课率" || hn === "完课触达率") {
            td.style.textAlign = "center";
          } else if (hn === "组长" || hn === "老师花名") {
            td.style.textAlign = "left";
          } else {
            td.style.textAlign = "right";
          }
        }
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    return table;
  }

  /* ========== 生成 workbook (带样式提示) ========== */
  function makeSheetAOA(sheet) {
    return [sheet.headers.slice()].concat(sheet.rows.map((r) => r.slice()));
  }

  function buildWorkbook(sheets) {
    const wb = XLSX.utils.book_new();
    for (const name of Object.keys(sheets)) {
      const aoa = makeSheetAOA(sheets[name]);
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      // 自动列宽
      const cols = sheets[name].headers.map((h, i) => {
        const maxLen = Math.max(
          String(h).length,
          ...sheets[name].rows.slice(0, 500).map((r) => String(r[i] ?? "").length)
        );
        return { wch: Math.min(40, Math.max(10, maxLen + 2)) };
      });
      ws["!cols"] = cols;
      XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 31));
    }
    return wb;
  }

  /* ========== 本地缓存 ========== */
  function saveCache() {
    try {
      const cache = {
        fileNames: state.fileName,
        // 原始 rows 按原样缓存（可丢失二进制 blob，但已处理过）
        t: Date.now(),
      };
      // 直接缓存已处理的 baseClass / baseUser / acClass / acUser 的原始行数据
      // 为了避免体积大，缓存 state.raw
      cache.raw = state.raw;
      localStorage.setItem(CACHE_KEY, JSON.stringify(cache));
    } catch (e) {
      console.warn("缓存写入失败", e);
    }
  }
  function loadCache() {
    try {
      const s = localStorage.getItem(CACHE_KEY);
      if (!s) return null;
      return JSON.parse(s);
    } catch { return null; }
  }
  function clearCache() { localStorage.removeItem(CACHE_KEY); }

  /* ========== 上传卡片交互 ========== */
  function bindUploadCards() {
    const cards = document.querySelectorAll(".upload-card");
    cards.forEach((card) => {
      const slot = card.dataset.slot;
      const input = card.querySelector(".file-input");
      const statusBadge = card.querySelector(".status-badge");
      const nameEl = card.querySelector(".file-name");

      input.addEventListener("change", async (e) => {
        const file = e.target.files[0];
        if (!file) return;
        try {
          const rows = await readFile(file);
          if (!rows.length) { toast("文件为空"); return; }
          state.raw[slot] = rows;
          state.fileName[slot] = file.name;
          card.classList.add("loaded");
          statusBadge.textContent = "已上传";
          statusBadge.classList.remove("text-slate-400");
          statusBadge.classList.add("text-emerald-600");
          nameEl.textContent = file.name + "（" + rows.length + "行）";

          // 基础班级：提取班级ID
          if (slot === "class") {
            const ids = [];
            for (const r of rows) {
              const v = pickField(r, ["班级ID"]);
              if (v !== "" && v !== undefined && v !== null) ids.push(String(v));
            }
            renderClassIds(ids);
          }
          updateButtons();
          saveCache();
          toast("上传成功");
        } catch (err) {
          console.error(err);
          toast("文件解析失败");
        }
      });
    });

    document.querySelector(".copy-ids-btn")?.addEventListener("click", () => {
      const chips = document.querySelectorAll(".class-ids-list .chip");
      const ids = Array.from(chips).map((c) => c.textContent.trim()).join("\n");
      if (!ids) return;
      copyText(ids).then(() => toast("班级ID已复制到剪贴板"));
    });
  }

  function renderClassIds(ids) {
    const box = document.querySelector(".class-ids-box");
    const list = document.querySelector(".class-ids-list");
    const count = document.querySelector(".class-ids-box .count");
    if (!box) return;
    list.innerHTML = "";
    ids.forEach((id) => {
      const span = document.createElement("span");
      span.className = "chip";
      span.textContent = id;
      list.appendChild(span);
    });
    count.textContent = ids.length;
    box.classList.toggle("hidden", ids.length === 0);
  }

  function updateButtons() {
    const baseOk = !!(state.raw["class"] && state.raw["user"]);
    const fullOk = baseOk && !!state.raw["activity-class"] && !!state.raw["activity-user"];
    document.getElementById("btnBase").disabled = !baseOk;
    document.getElementById("btnFull").disabled = !fullOk;
    const hint = document.getElementById("hintText");
    if (fullOk) hint.textContent = "4 份数据已就绪，可生成完整结果并导出图片与文件。";
    else if (baseOk) hint.textContent = "基础数据已就绪，可生成底表；如需完整数据，请上传活动班级与活动用户。";
    else hint.textContent = "请先上传基础班级 + 基础用户数据，即可生成底表。";
  }

  /* ========== 文本复制 ========== */
  async function copyText(text) {
    try {
      if (navigator.clipboard && navigator.clipboard.writeText) {
        await navigator.clipboard.writeText(text);
        return;
      }
    } catch {}
    const ta = document.createElement("textarea");
    ta.value = text;
    ta.style.position = "fixed";
    ta.style.opacity = "0";
    document.body.appendChild(ta);
    ta.select();
    document.execCommand("copy");
    document.body.removeChild(ta);
  }

  /**
   * 将目标 sheet 的表格渲染为图片并复制到剪贴板
   * @param {string} sheetName - sheet 名称（如"获奖名单" / "周四看板"）
   * @param {string} mode - "award" 或 "dashboard"，仅用于 toast 文案
   */
  async function copySheetAsImage(sheetName, mode) {
    if (!state.sheets || !state.sheets[sheetName]) {
      toast("未找到对应数据，请先点击文件生成");
      return;
    }

    // 1) 切 tab，确保快照 DOM 已渲染在 sheetContainer 中
    const tabs = document.querySelectorAll(".sheet-tab");
    let targetTab = null;
    tabs.forEach((t) => { if (t.textContent === sheetName) targetTab = t; });
    if (targetTab) {
      tabs.forEach((x) => x.classList.remove("active"));
      targetTab.classList.add("active");
    }
    // 强制渲染目标 sheet（即便是其他 tab 也会让 snapshot-xxx 容器存在）
    renderSheetTable(sheetName);

    // 2) 查找快照 DOM
    let target = document.getElementById("snapshot-" + sheetName);
    if (!target) {
      // 兜底：直接使用 sheetContainer 下的整段渲染结果
      target = document.getElementById("sheetContainer");
    }
    if (!target) {
      toast("未找到内容");
      return;
    }

    // 3) 临时将目标容器从 overflow/滚动中取出，保证完整可见
    const container = document.getElementById("sheetContainer");
    const origOverflow = container ? container.style.overflow : "";
    const origMaxH = container ? container.style.maxHeight : "";
    if (container) {
      container.style.overflow = "visible";
      container.style.maxHeight = "none";
    }

    try {
      if (typeof html2canvas !== "function") {
        toast("html2canvas 未加载，请检查网络");
        return;
      }
      const canvas = await html2canvas(target, {
        backgroundColor: "#ffffff",
        scale: 2,
        useCORS: true,
        logging: false,
        windowWidth: Math.max(target.scrollWidth, 800),
        windowHeight: target.scrollHeight + 200,
      });
      const dataUrl = canvas.toDataURL("image/png");

      // 4) 写入剪贴板
      let copiedToClipboard = false;
      try {
        if (navigator.clipboard && window.ClipboardItem && canvas.toBlob) {
          await new Promise((resolve, reject) => {
            canvas.toBlob(async (blob) => {
              if (!blob) { reject(new Error("blob 失败")); return; }
              try {
                await navigator.clipboard.write([new ClipboardItem({ "image/png": blob })]);
                copiedToClipboard = true;
                resolve();
              } catch (e) { reject(e); }
            }, "image/png");
          });
        }
      } catch (e) {
        console.warn("剪贴板写入失败，降级为图片下载", e);
      }

      // 5) 在按钮旁渲染图片预览（点击可下载）
      renderImagePreview(sheetName, dataUrl, mode);

      if (copiedToClipboard) {
        toast((mode === "dashboard" ? "周x看板" : "获奖名单") + "图片已复制到剪贴板");
      } else {
        toast((mode === "dashboard" ? "周x看板" : "获奖名单") + "图片已生成，点击预览图片可下载");
      }
    } catch (e) {
      console.error(e);
      toast("图片生成失败：" + (e && e.message ? e.message : e));
    } finally {
      if (container) {
        container.style.overflow = origOverflow || "";
        container.style.maxHeight = origMaxH || "";
      }
    }
  }

  /**
   * 在按钮旁渲染一张小图预览（点击下载）
   */
  function renderImagePreview(sheetName, dataUrl, mode) {
    const hostId = "image-preview-" + mode;
    let host = document.getElementById(hostId);
    if (!host) {
      host = document.createElement("div");
      host.id = hostId;
      host.style.padding = "8px";
      host.style.border = "1px dashed #94a3b8";
      host.style.borderRadius = "6px";
      host.style.background = "#f8fafc";
      host.style.fontSize = "12px";
      host.style.color = "#475569";
      const parent = document.getElementById("imagePreviewArea");
      if (parent) parent.appendChild(host);
      else document.getElementById("resultSection").appendChild(host);
    }
    host.innerHTML = "";
    const title = document.createElement("div");
    title.textContent = (mode === "dashboard" ? "📊 周x看板" : "🏆 获奖名单") + " · 图片预览";
    title.style.fontWeight = "600";
    title.style.marginBottom = "6px";
    title.style.color = "#0f172a";
    title.style.fontSize = "13px";
    host.appendChild(title);

    const imgWrap = document.createElement("div");
    imgWrap.style.overflow = "auto";
    imgWrap.style.maxHeight = "360px";
    imgWrap.style.background = "#ffffff";
    imgWrap.style.border = "1px solid #e2e8f0";
    imgWrap.style.borderRadius = "4px";
    imgWrap.style.padding = "8px";

    const img = document.createElement("img");
    img.src = dataUrl;
    img.style.maxWidth = "100%";
    img.style.display = "block";
    img.style.cursor = "pointer";
    img.title = "点击图片下载 PNG 文件";
    img.addEventListener("click", () => {
      const a = document.createElement("a");
      a.href = dataUrl;
      a.download = sheetName + ".png";
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
    });
    imgWrap.appendChild(img);
    host.appendChild(imgWrap);

    const tip = document.createElement("div");
    tip.style.marginTop = "6px";
    tip.textContent = "✅ 已复制到剪贴板，可直接粘贴到聊天/文档；点击图片可下载 PNG 原文件。";
    host.appendChild(tip);
  }

  /* ========== 主处理流程 ========== */
  function processBase() {
    if (!state.raw.class || !state.raw.user) { toast("请先上传基础班级与基础用户"); return; }
    const baseClass = processBaseClass(state.raw.class);
    // 班级ID -> 开课日 lookup
    const classLookup = {};
    baseClass.rows.forEach((r) => { classLookup[String(r[0])] = r[r.length - 1]; });
    // 用户ID -> 孩子姓名
    const userLookup = {};
    const baseUser = processBaseUser(state.raw.user, classLookup);
    baseUser.rows.forEach((r) => { userLookup[String(r[0])] = r[1] || ""; });

    state.sheets = {
      "底表不动（班级）": baseClass,
      "底表不动（用户）": baseUser,
    };
    state.fileNameBase = shortenTerm(baseClass.rows[0] ? baseClass.rows[0][6] : "");
    state.lastMode = "base";
    renderSheetTabs(Object.keys(state.sheets));
    document.getElementById("resultSection").classList.remove("hidden");
    document.getElementById("btnCopyAward").classList.add("hidden");
    document.getElementById("btnCopyDashboard").classList.add("hidden");
    toast("底表生成成功");
  }

  function processFull() {
    if (!state.raw.class || !state.raw.user || !state.raw["activity-class"] || !state.raw["activity-user"]) {
      toast("请上传全部 4 份文件"); return;
    }
    const baseClass = processBaseClass(state.raw.class);
    const classLookup = {};
    baseClass.rows.forEach((r) => { classLookup[String(r[0])] = r[r.length - 1]; });

    const baseUser = processBaseUser(state.raw.user, classLookup);
    const userLookup = {};
    baseUser.rows.forEach((r) => { userLookup[String(r[0])] = r[1] || ""; });

    // 活动班级：取原始字段名
    const acHeadersRaw = state.raw["activity-class"].length ? Object.keys(state.raw["activity-class"][0]) : [];
    const activityClass = processActivityClass(state.raw["activity-class"], acHeadersRaw, classLookup);
    state.activityClass = activityClass;

    const auHeadersRaw = state.raw["activity-user"].length ? Object.keys(state.raw["activity-user"][0]) : [];
    const activityUser = processActivityUser(state.raw["activity-user"], auHeadersRaw, userLookup, classLookup);

    const award = buildAwardList(activityUser);
    const finishedExp = filterFinishedExp(activityUser, true);
    const finishedNoExp = filterFinishedExp(activityUser, false);

    // 默认全选所有出现的开课日（不再依赖系统时间）
    const allDays = collectOpenDays(activityClass);
    state.selectedDays = allDays.slice();
    const dash = buildDashboard(activityClass, state.selectedDays);
    const dashName = dash.autoName;
    state.dashboardName = dashName;

    state.sheets = {
      "底表不动（班级）": baseClass,
      "底表不动（用户）": baseUser,
      "班级信息": activityClass,
      "用户明细": activityUser,
      "获奖名单": award,
      "完课已体验用户": finishedExp,
      "完课未体验用户": finishedNoExp,
      [dashName]: dash,
    };
    state.fileNameBase = shortenTerm(baseClass.rows[0] ? baseClass.rows[0][6] : "");
    state.lastMode = "full";
    renderSheetTabs(Object.keys(state.sheets));
    document.getElementById("resultSection").classList.remove("hidden");
    document.getElementById("btnCopyAward").classList.remove("hidden");
    document.getElementById("btnCopyDashboard").classList.remove("hidden");
    toast("完整数据生成成功");
  }

  function downloadFile() {
    if (!Object.keys(state.sheets).length) { toast("请先生成数据"); return; }
    const wb = buildWorkbook(state.sheets);
    const base = state.fileNameBase || "S低思维数据";
    const name = state.lastMode === "base" ? `${base}底表.xlsx` : `${base}.xlsx`;
    XLSX.writeFile(wb, name);
    toast("文件已下载：" + name);
  }

  /* ========== 缓存恢复弹窗 ========== */
  function bindCacheModal() {
    const cache = loadCache();
    if (!cache || !cache.raw || !cache.raw.class) return;
    const modal = document.getElementById("cacheModal");
    modal.classList.remove("hidden");
    modal.classList.add("flex");
    document.getElementById("cacheRestoreBtn").addEventListener("click", () => {
      // 恢复 raw 数据并刷新状态与班级ID
      state.raw = cache.raw;
      state.fileName = cache.fileNames || {};
      const cards = document.querySelectorAll(".upload-card");
      cards.forEach((card) => {
        const slot = card.dataset.slot;
        if (state.raw[slot] && state.raw[slot].length) {
          card.classList.add("loaded");
          const sb = card.querySelector(".status-badge");
          sb.textContent = "已上传";
          sb.classList.remove("text-slate-400");
          sb.classList.add("text-emerald-600");
          const nm = card.querySelector(".file-name");
          nm.textContent = (state.fileName[slot] || "已恢复文件") + "（" + state.raw[slot].length + "行）";
        }
      });
      if (state.raw.class && state.raw.class.length) {
        const ids = [];
        for (const r of state.raw.class) {
          const v = pickField(r, ["班级ID"]);
          if (v !== "" && v !== undefined && v !== null) ids.push(String(v));
        }
        renderClassIds(ids);
      }
      updateButtons();
      modal.classList.add("hidden");
      toast("已恢复上次缓存");
    });
    document.getElementById("cacheIgnoreBtn").addEventListener("click", () => {
      clearCache();
      modal.classList.add("hidden");
    });
  }

  /* ========== 主入口 ========== */
  function init() {
    bindUploadCards();
    document.getElementById("btnBase").addEventListener("click", processBase);
    document.getElementById("btnFull").addEventListener("click", processFull);
    document.getElementById("btnDownload").addEventListener("click", downloadFile);
    document.getElementById("btnCopyAward").addEventListener("click", () => copySheetAsImage("获奖名单", "award"));
    document.getElementById("btnCopyDashboard").addEventListener("click", () => {
      const dashName = state.dashboardName || Object.keys(state.sheets).find((k) => /看板/.test(k));
      if (dashName) copySheetAsImage(dashName, "dashboard");
      else toast("未找到看板数据");
    });
    document.getElementById("btnClear").addEventListener("click", () => {
      if (!confirm("确认清空所有上传数据与缓存？")) return;
      state.raw = { class: null, user: null, "activity-class": null, "activity-user": null };
      state.fileName = {};
      state.sheets = {};
      state.workbook = null;
      state.dashboardName = "";
      clearCache();
      document.querySelectorAll(".upload-card").forEach((c) => {
        c.classList.remove("loaded");
        c.querySelector(".status-badge").textContent = "未上传";
        c.querySelector(".status-badge").classList.add("text-slate-400");
        c.querySelector(".status-badge").classList.remove("text-emerald-600");
        c.querySelector(".file-name").textContent = "";
      });
      document.querySelectorAll(".file-input").forEach((i) => (i.value = ""));
      document.querySelector(".class-ids-box")?.classList.add("hidden");
      document.getElementById("resultSection").classList.add("hidden");
      updateButtons();
      toast("已清空");
    });

    bindCacheModal();
    updateButtons();
  }

  document.addEventListener("DOMContentLoaded", init);
})();
