// 数据结构：{ teachers: [...], classes: [...], timetableStats: {...}, gradeDayStats: [...] }
const STORAGE_KEY = "teacherSalarySystem";

let state = {
  teachers: [],
  classes: [],
  // 由Excel课表自动统计：{ teacherStats: [{ name, expected, actual }], sourceFileName, updatedAt }
  timetableStats: {
    teacherStats: [],
    sourceFileName: "",
    updatedAt: ""
  },
  // 按年级、星期、上午/下午统计课时：[{ grade, teacher, weekday, am, pm }]
  gradeDayStats: []
};

const teacherForm = document.getElementById("teacherForm");
const teacherTableBody = document.getElementById("teacherTableBody");

const classForm = document.getElementById("classForm");
const classTeacherSelect = document.getElementById("classTeacher");
const classDateInput = document.getElementById("classDate");
const classCourseInput = document.getElementById("classCourse");
const classHoursInput = document.getElementById("classHours");
const classRateInput = document.getElementById("classRate");
const classRemarkInput = document.getElementById("classRemark");
const classTableBody = document.getElementById("classTableBody");

const filterTeacherSelect = document.getElementById("filterTeacher");
const filterMonthInput = document.getElementById("filterMonth");

const salaryTeacherSelect = document.getElementById("salaryTeacher");
const salaryMonthInput = document.getElementById("salaryMonth");
const calcSalaryBtn = document.getElementById("calcSalaryBtn");
const salaryResult = document.getElementById("salaryResult");

// Excel 课表相关 DOM
const timetableFileInput = document.getElementById("timetableFile");
const parseTimetableBtn = document.getElementById("parseTimetableBtn");
const exportStatsBtn = document.getElementById("exportStatsBtn");
const timetableStatsBody = document.getElementById("timetableStatsBody");

// 分年级按天统计相关 DOM
const gradeStatsFilesInput = document.getElementById("gradeStatsFiles");
const importGradeStatsBtn = document.getElementById("importGradeStatsBtn");
const gradeSelectForStats = document.getElementById("gradeSelectForStats");
const calcGradeStatsBtn = document.getElementById("calcGradeStatsBtn");
const gradeStatsResult = document.getElementById("gradeStatsResult");
const gradeStatsTableBody = document.getElementById("gradeStatsTableBody");
const clearGradeStatsBtn = document.getElementById("clearGradeStatsBtn");
const exportGradeStatsBtn = document.getElementById("exportGradeStatsBtn");
const gradeStatsTeacherSearch = document.getElementById(
  "gradeStatsTeacherSearch"
);
const gradeStatsTeacherClear = document.getElementById(
  "gradeStatsTeacherClear"
);
const gradeStatsTeacherTip = document.getElementById("gradeStatsTeacherTip");
const teacherHoursTemplateInput = document.getElementById(
  "teacherHoursTemplate"
);
const applyHoursToTemplateBtn = document.getElementById(
  "applyHoursToTemplateBtn"
);

// 最近一次“分年级按天统计”的结果缓存（用于导出）
let lastGradeStatsCache = null;

// 初始化
init();

function init() {
  loadState();
  // 如果相关DOM存在才渲染（支持删除对应板块后的页面）
  if (teacherTableBody) {
    renderTeachers();
    renderTeacherOptions();
  }
  if (classTableBody) {
    renderClasses();
  }

  // 默认上课日期为今天
  if (classDateInput) {
    const today = new Date().toISOString().slice(0, 10);
    classDateInput.value = today;
  }

  bindEvents();
}

function bindEvents() {
  if (teacherForm) {
    teacherForm.addEventListener("submit", onAddTeacher);
  }
  if (classForm) {
    classForm.addEventListener("submit", onAddClass);
  }

  if (filterTeacherSelect) {
    filterTeacherSelect.addEventListener("change", renderClasses);
  }
  if (filterMonthInput) {
    filterMonthInput.addEventListener("change", renderClasses);
  }

  if (classTeacherSelect) {
    classTeacherSelect.addEventListener("change", syncClassRateWithTeacher);
  }

  if (calcSalaryBtn) {
    calcSalaryBtn.addEventListener("click", onCalcSalary);
  }

  // Excel 课表相关事件
  if (parseTimetableBtn && timetableFileInput) {
    parseTimetableBtn.addEventListener("click", onParseTimetable);
  }
  if (exportStatsBtn) {
    exportStatsBtn.addEventListener("click", onExportStats);
  }

  // 页面初始化时渲染一次Excel统计表（如果本地已有数据）
  renderTimetableStats();

  // 分年级按天统计事件
  if (importGradeStatsBtn && gradeStatsFilesInput) {
    importGradeStatsBtn.addEventListener("click", onImportGradeStats);
  }
  if (calcGradeStatsBtn) {
    calcGradeStatsBtn.addEventListener("click", onCalcGradeStats);
  }
  if (clearGradeStatsBtn) {
    clearGradeStatsBtn.addEventListener("click", onClearGradeStats);
  }
  if (exportGradeStatsBtn) {
    exportGradeStatsBtn.addEventListener("click", onExportGradeStats);
  }
  if (gradeStatsTeacherSearch) {
    gradeStatsTeacherSearch.addEventListener("input", applyGradeStatsSearch);
  }
  if (gradeStatsTeacherClear) {
    gradeStatsTeacherClear.addEventListener("click", () => {
      if (!gradeStatsTeacherSearch) return;
      gradeStatsTeacherSearch.value = "";
      applyGradeStatsSearch();
      gradeStatsTeacherSearch.focus();
    });
  }
  if (applyHoursToTemplateBtn) {
    applyHoursToTemplateBtn.addEventListener("click", onApplyHoursToTemplate);
  }
}

function loadState() {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (raw) {
    try {
      const parsed = JSON.parse(raw);
      if (parsed && typeof parsed === "object") {
        state = {
          teachers: parsed.teachers || [],
          classes: parsed.classes || [],
          timetableStats:
            parsed.timetableStats || {
              teacherStats: [],
              sourceFileName: "",
              updatedAt: ""
            },
          gradeDayStats: parsed.gradeDayStats || []
        };
      }
    } catch (e) {
      console.error("加载本地数据失败：", e);
    }
  }
}

function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

/* 工具函数 */

function createId() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
}

function findTeacherById(id) {
  return state.teachers.find(t => t.id === id) || null;
}

function formatCurrency(n) {
  return (Math.round(n * 100) / 100).toFixed(2);
}

// 统一的年级枚举
const GRADE_LABELS = ["七年级", "八年级", "九年级", "高一", "高二", "高三"];

// 将Excel中的年级文本归一化到上面的枚举
function normalizeGrade(text) {
  if (!text) return "";
  const t = text.toString().replace(/\s/g, "");
  if (/初一|七年级|初1/.test(t)) return "初一";
  if (/初二|八年级|初2/.test(t)) return "初二";
  if (/初三|九年级|初3/.test(t)) return "初三";
  if (/高一|高1/.test(t)) return "高一";
  if (/高二|高2/.test(t)) return "高二";
  if (/高三|高3/.test(t)) return "高三";
  // 已经是目标枚举之一
  if (GRADE_LABELS.includes(t)) return t;
  return "";
}

// 将“星期一/周一/礼拜一”等转成 1~5
function weekdayTextToNumber(text) {
  if (!text) return null;
  const t = text.toString().replace(/\s/g, "");
  if (/一/.test(t) && /周|星期|礼拜/.test(t)) return 1;
  if (/二/.test(t) && /周|星期|礼拜/.test(t)) return 2;
  if (/三/.test(t) && /周|星期|礼拜/.test(t)) return 3;
  if (/四/.test(t) && /周|星期|礼拜/.test(t)) return 4;
  if (/五/.test(t) && /周|星期|礼拜/.test(t)) return 5;
  return null;
}

// 判断时段：上午 / 下午 / 全天
function parsePeriodType(text) {
  if (!text) return "";
  const t = text.toString().replace(/\s/g, "");
  if (/全天|一整天/.test(t)) return "all";
  if (/上午|早上|早晨/.test(t)) return "am";
  if (/下午/.test(t)) return "pm";
  return "";
}

// 判断表头是否为周一至周五的某一天
function isWeekdayHeader(text) {
  if (!text) return false;
  const t = text.toString().replace(/\s/g, "");
  // 只匹配周一至周五，排除周六周日
  return /周[一二三四五]|星期[一二三四五]/.test(t);
}

// 从单元格内容中提取教师姓名，支持：
// "语文-张三"、"数学：李四"、"物理(王五)"、"英语（赵六）"、"张三/李四"、"张三、李四"
function extractTeacherNamesFromCell(text) {
  if (!text) return [];
  const clean = text.toString().replace(/\s+/g, "");
  if (!clean) return [];

  const segments = clean.split(/[、，,\/]/);
  const names = [];

  segments.forEach(seg => {
    if (!seg) return;
    let name = seg;

    // 优先从括号中取
    const parenMatch = seg.match(/[（(]([^）)]+)[）)]/);
    if (parenMatch && parenMatch[1]) {
      name = parenMatch[1];
    } else {
      // 从 - ： 之后截取
      const m = seg.match(/[-－—:：]([^-\-—:：]+)$/);
      if (m && m[1]) {
        name = m[1];
      }
    }

    name = name.trim();
    if (name) {
      names.push(name);
    }
  });

  return names;
}

function normalizeTeacherName(name) {
  return (name || "").toString().replace(/\s/g, "");
}

function getSheetMatrix(sheet) {
  const range = XLSX.utils.decode_range(sheet["!ref"]);
  const rows = [];
  for (let r = range.s.r; r <= range.e.r; r++) {
    const row = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = sheet[XLSX.utils.encode_cell({ r, c })];
      let v = cell ? cell.v : "";
      if (v === null || typeof v === "undefined") v = "";
      row.push(v.toString ? v.toString().trim() : v);
    }
    rows.push(row);
  }
  return { rows, range };
}

function findHeaderCell(matrix, pattern) {
  const { rows, range } = matrix;
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r] || [];
    for (let c = 0; c < row.length; c++) {
      const text = row[c];
      if (text && pattern.test(text)) {
        return { row: range.s.r + r, col: range.s.c + c, text };
      }
    }
  }
  return null;
}

function findNameHeaderCell(matrix) {
  let header = findHeaderCell(
    matrix,
    /^(姓名|老师姓名|教师姓名|老师|教师)$/
  );
  if (header) return header;

  const { rows, range } = matrix;
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r] || [];
    const rowText = row.join("");
    const hasHeaderContext = /科目|任课|班级|年级|上课/.test(rowText);
    if (!hasHeaderContext) continue;
    for (let c = 0; c < row.length; c++) {
      const text = row[c];
      if (text && /姓名|老师姓名|教师姓名|老师|教师/.test(text)) {
        return { row: range.s.r + r, col: range.s.c + c, text };
      }
    }
  }

  return null;
}

function findTotalActualHoursColumn(sheet, matrix) {
  const { rows, range } = matrix;
  const merges = sheet["!merges"] || [];
  let totalMerge = null;

  merges.forEach(m => {
    if (totalMerge) return;
    const cell = sheet[XLSX.utils.encode_cell(m.s)];
    const text = cell ? cell.v : "";
    if (text && /总课时/.test(text.toString())) {
      totalMerge = m;
    }
  });

  if (totalMerge) {
    const startRow = totalMerge.s.r - range.s.r;
    const searchRows = [startRow + 1, startRow + 2, startRow + 3];
    for (let c = totalMerge.s.c; c <= totalMerge.e.c; c++) {
      for (let i = 0; i < searchRows.length; i++) {
        const r = searchRows[i];
        if (r < 0 || r >= rows.length) continue;
        const text = (rows[r] || [])[c - range.s.c];
        if (text && /实上/.test(text)) {
          return c;
        }
      }
    }
  }

  // 兜底：同一行包含“总课时”和“实上”，取距离“总课时”最近的“实上”
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r] || [];
    let totalIdx = -1;
    row.forEach((text, idx) => {
      if (text && /总课时/.test(text)) totalIdx = idx;
    });
    if (totalIdx === -1) continue;
    let candidate = -1;
    row.forEach((text, idx) => {
      if (text && /实上/.test(text)) {
        if (candidate === -1) candidate = idx;
        else if (Math.abs(idx - totalIdx) < Math.abs(candidate - totalIdx)) {
          candidate = idx;
        }
      }
    });
    if (candidate !== -1) {
      return range.s.c + candidate;
    }
  }

  // 最后兜底：找到所有“实上”列，优先选择最靠右的一列（通常为总课时实上）
  let fallbackCol = null;
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r] || [];
    row.forEach((text, idx) => {
      if (text && /实上/.test(text)) {
        fallbackCol = range.s.c + idx;
      }
    });
  }
  return fallbackCol;
}

/* 教师管理 */

function onAddTeacher(e) {
  e.preventDefault();

  const name = document.getElementById("teacherName").value.trim();
  const title = document.getElementById("teacherTitle").value.trim();
  const rate = parseFloat(document.getElementById("teacherRate").value);

  if (!name) {
    alert("教师姓名不能为空");
    return;
  }
  if (isNaN(rate) || rate < 0) {
    alert("课时单价必须为非负数字");
    return;
  }

  const teacher = {
    id: createId(),
    name,
    title,
    rate
  };

  state.teachers.push(teacher);
  saveState();
  renderTeachers();
  renderTeacherOptions();

  teacherForm.reset();
}

function renderTeachers() {
  if (!teacherTableBody) return;

  teacherTableBody.innerHTML = "";

  if (state.teachers.length === 0) {
    teacherTableBody.innerHTML =
      '<tr><td colspan="4" style="text-align:center;color:#9ca3af;">暂时没有教师，请先添加。</td></tr>';
    return;
  }

  state.teachers.forEach(teacher => {
    const tr = document.createElement("tr");

    const titleText = teacher.title || "-";

    tr.innerHTML = `
      <td>${teacher.name}</td>
      <td>${titleText}</td>
      <td>${formatCurrency(teacher.rate)}</td>
      <td>
        <button class="btn danger btn-small" data-id="${teacher.id}">删除</button>
      </td>
    `;

    const deleteBtn = tr.querySelector("button");
    deleteBtn.addEventListener("click", () => onDeleteTeacher(teacher.id));

    teacherTableBody.appendChild(tr);
  });
}

function onDeleteTeacher(id) {
  const teacher = findTeacherById(id);
  if (!teacher) return;

  const hasClasses = state.classes.some(c => c.teacherId === id);
  if (hasClasses) {
    const confirmMsg =
      "该教师已经存在上课记录，删除后不会影响已记录的上课情况，但无法再选择该教师进行新记录。\n\n确定要删除吗？";
    if (!window.confirm(confirmMsg)) {
      return;
    }
  } else {
    if (!window.confirm("确定要删除该教师吗？")) {
      return;
    }
  }

  state.teachers = state.teachers.filter(t => t.id !== id);
  saveState();
  renderTeachers();
  renderTeacherOptions();
  renderClasses();
}

function renderTeacherOptions() {
  if (!classTeacherSelect && !filterTeacherSelect && !salaryTeacherSelect) {
    return;
  }

  const selects = [classTeacherSelect, filterTeacherSelect, salaryTeacherSelect];

  selects.forEach(select => {
    if (!select) return;

    const currentValue = select.value;

    // 保留第一个“全部教师”或空选项
    const firstOption = select.querySelector("option:first-child");
    const placeholder = firstOption
      ? { value: firstOption.value, text: firstOption.textContent }
      : null;

    select.innerHTML = "";

    if (placeholder) {
      const opt = document.createElement("option");
      opt.value = placeholder.value;
      opt.textContent = placeholder.text;
      select.appendChild(opt);
    }

    state.teachers.forEach(teacher => {
      const opt = document.createElement("option");
      opt.value = teacher.id;
      opt.textContent = teacher.name;
      select.appendChild(opt);
    });

    // 尝试恢复选中
    if (currentValue) {
      select.value = currentValue;
    }
  });

  // 如果上课教师下拉没有选中，默认选第一个教师
  if (classTeacherSelect && !classTeacherSelect.value && state.teachers[0]) {
    classTeacherSelect.value = state.teachers[0].id;
  }

  // 同步课时单价
  syncClassRateWithTeacher();
}

/* Excel 课表导入与周课时统计 */

function onParseTimetable() {
  if (!timetableFileInput || !timetableFileInput.files.length) {
    alert("请先选择要上传的课表 Excel 文件（.xls 或 .xlsx）。");
    return;
  }
  if (typeof XLSX === "undefined") {
    alert("Excel 解析库未加载成功，请检查网络后刷新页面重试。");
    return;
  }

  const file = timetableFileInput.files[0];
  const reader = new FileReader();

  reader.onload = function (e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });

      const teacherMap = {}; // { name: { name, expected, actual } }

      workbook.SheetNames.forEach(sheetName => {
        const sheet = workbook.Sheets[sheetName];
        if (!sheet) return;

        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });
        if (!rows || rows.length < 2) return;

        // ① 尝试“表头为周一至周五”的横向结构
        const headerRow = rows[0].map(v => (v || "").toString().trim());
        const weekdayIndices = [];

        headerRow.forEach((h, idx) => {
          if (isWeekdayHeader(h)) {
            weekdayIndices.push(idx);
          }
        });

        if (weekdayIndices.length > 0) {
          for (let i = 1; i < rows.length; i++) {
            const row = rows[i] || [];
            weekdayIndices.forEach(colIdx => {
              const cell = (row[colIdx] || "").toString().trim();
              if (!cell) return;

              const names = extractTeacherNamesFromCell(cell);
              names.forEach(name => {
                if (!name) return;
                if (!teacherMap[name]) {
                  teacherMap[name] = {
                    name,
                    expected: 0,
                    actual: 0
                  };
                }
                // 周一至周五每出现一次，记 1 节
                teacherMap[name].expected += 1;
              });
            });
          }
          return; // 本工作表已按该结构统计完毕
        }

        // ② 兼容你提供的“行上写星期几，列为班级”的课表结构
        let currentWeekday = null;

        for (let i = 0; i < rows.length; i++) {
          const row = rows[i] || [];
          const firstCell = (row[0] || "").toString().trim();

          // 如果本行首列是“星期一/周一”等，则记录当前星期
          if (isWeekdayHeader(firstCell)) {
            currentWeekday = firstCell;
            continue;
          }

          // 只有处在“星期一~星期五”区块内，才会统计
          if (!currentWeekday) continue;

          // 第二列一般是“节次/节数”，为空时跳过
          const periodCell = (row[1] || "").toString().trim();
          if (!periodCell) continue;

          // 从第3列开始是各个班级，对应的任课老师/科目
          for (let colIdx = 2; colIdx < row.length; colIdx++) {
            const cell = (row[colIdx] || "").toString().trim();
            if (!cell) continue;

            const names = extractTeacherNamesFromCell(cell);
            names.forEach(name => {
              if (!name) return;
              if (!teacherMap[name]) {
                teacherMap[name] = {
                  name,
                  expected: 0,
                  actual: 0
                };
              }
              teacherMap[name].expected += 1;
            });
          }
        }
      });

      const teacherStats = Object.values(teacherMap);
      teacherStats.forEach(t => {
        t.actual = t.expected; // 初始“实上”=“应上”，后续可手动调整
      });

      state.timetableStats = {
        teacherStats,
        sourceFileName: file.name,
        updatedAt: new Date().toISOString()
      };
      saveState();
      renderTimetableStats();

      if (!teacherStats.length) {
        alert(
          "没有在 Excel 中识别到周一至周五的课表结构，请检查模板是否符合提示要求。"
        );
      } else {
        alert("课表解析完成，已根据周一至周五统计出每位教师每周的“应上”课时。");
      }
    } catch (err) {
      console.error("解析 Excel 失败：", err);
      alert("解析 Excel 失败，请确认文件为有效的课表并重试。");
    }
  };

  reader.onerror = function () {
    alert("读取文件失败，请重试。");
  };

  reader.readAsArrayBuffer(file);
}

/* 初中 / 高中 老师课程统计表导入（按时间段、年级分行） */

// 适配你截图中的结构：
// 表头示例：老师姓名 | 时间段 | 年级 | 周一 | 周二 | 周三 | 周四 | 周五
// 每一行：某位老师 + 某个时间段（上午/下午） + 某个年级 + 各天节数
// 示例：马欢 | 上午 | 九年级 | 0 | 1 | 0 | 0 | 1

function onImportGradeStats() {
  if (!gradeStatsFilesInput || !gradeStatsFilesInput.files.length) {
    alert("请先选择“初中老师课程统计_按时间段年级分行.xlsx”和“高中老师课程统计_按时间段年级分行.xlsx”文件（可一次选择两个）。");
    return;
  }
  if (typeof XLSX === "undefined") {
    alert("Excel 解析库未加载成功，请检查网络后刷新页面重试。");
    return;
  }

  const files = Array.from(gradeStatsFilesInput.files);
  const gradeMap = {}; // { grade: { teacherName: { grade, teacher, weekdayStats: {1:{am,pm},...} } } }

  let remaining = files.length;
  files.forEach(file => {
    const reader = new FileReader();
    reader.onload = function (e) {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });

        workbook.SheetNames.forEach(sheetName => {
          const sheet = workbook.Sheets[sheetName];
          if (!sheet) return;

          const rows = XLSX.utils.sheet_to_json(sheet, {
            header: 1,
            raw: false
          });
          if (!rows || rows.length < 2) return;

          const header = rows[0].map(v => (v || "").toString().trim());

          // 定位关键列：老师姓名 / 时间段 / 年级 / 周一~周五
          let gradeCol = -1;
          let periodCol = -1;
          let teacherCol = -1;
          const weekdayCols = { 1: -1, 2: -1, 3: -1, 4: -1, 5: -1 };

          header.forEach((h, idx) => {
            if (teacherCol === -1 && /老师姓名|教师姓名|老师|教师/.test(h)) {
              teacherCol = idx;
            }
            if (periodCol === -1 && /时间段|时段|上午|下午|全天/.test(h)) {
              periodCol = idx;
            }
            if (gradeCol === -1 && /年级|七年级|八年级|九年级|高一|高二|高三/.test(h)) {
              gradeCol = idx;
            }
            if (weekdayCols[1] === -1 && /周一|星期一/.test(h)) weekdayCols[1] = idx;
            if (weekdayCols[2] === -1 && /周二|星期二/.test(h)) weekdayCols[2] = idx;
            if (weekdayCols[3] === -1 && /周三|星期三/.test(h)) weekdayCols[3] = idx;
            if (weekdayCols[4] === -1 && /周四|星期四/.test(h)) weekdayCols[4] = idx;
            if (weekdayCols[5] === -1 && /周五|星期五/.test(h)) weekdayCols[5] = idx;
          });

          if (gradeCol === -1 || periodCol === -1 || teacherCol === -1) {
            // 当前工作表结构不符合预期，跳过
            return;
          }

          for (let i = 1; i < rows.length; i++) {
            const row = rows[i] || [];

            const gradeRaw = row[gradeCol];
            const periodRaw = row[periodCol];
            const teacherRaw = row[teacherCol];

            const grade = normalizeGrade(gradeRaw);
            const periodType = parsePeriodType(periodRaw);
            if (!grade || !periodType) continue;

            const names = extractTeacherNamesFromCell(teacherRaw || "");
            if (!names.length) continue;

            names.forEach(name => {
              if (!name) return;

              if (!gradeMap[grade]) gradeMap[grade] = {};
              if (!gradeMap[grade][name]) {
                gradeMap[grade][name] = {
                  grade,
                  teacher: name,
                  weekdayStats: {}
                };
              }

              const teacherStat = gradeMap[grade][name];

              // 遍历周一~周五列，每列的数字就是该时间段的节数
              for (let w = 1; w <= 5; w++) {
                const colIdx = weekdayCols[w];
                if (colIdx === -1) continue;
                let val = row[colIdx];
                let count = parseFloat(val);
                if (isNaN(count) || count <= 0) continue;

                if (!teacherStat.weekdayStats[w]) {
                  teacherStat.weekdayStats[w] = { am: 0, pm: 0 };
                }
                const ws = teacherStat.weekdayStats[w];

                if (periodType === "am") {
                  ws.am += count;
                } else if (periodType === "pm") {
                  ws.pm += count;
                } else if (periodType === "all") {
                  ws.am += count;
                  ws.pm += count;
                }
              }
            });
          }
        });
      } catch (err) {
        console.error("解析年级课程统计表失败：", err);
      } finally {
        remaining -= 1;
        if (remaining === 0) {
          // 全部文件处理完，写入全局 state
          const results = [];
          Object.values(gradeMap).forEach(byTeacher => {
            Object.values(byTeacher).forEach(t => {
              const { grade, teacher, weekdayStats } = t;
              for (let w = 1; w <= 5; w++) {
                const ws = weekdayStats[w];
                if (!ws) continue;
                results.push({
                  grade,
                  teacher,
                  weekday: w,
                  am: ws.am || 0,
                  pm: ws.pm || 0
                });
              }
            });
          });

          state.gradeDayStats = results;
          saveState();
          alert("老师课程统计表导入完成，可以按年级和日期时段进行统计了。");
        }
      }
    };

    reader.onerror = function () {
      remaining -= 1;
      alert("读取文件失败，请重试。");
    };

    reader.readAsArrayBuffer(file);
  });
}

function onCalcGradeStats(e) {
  e.preventDefault();

  if (!state.gradeDayStats || !state.gradeDayStats.length) {
    alert("还没有导入老师课程统计表，请先导入后再统计。");
    return;
  }

  const gradeSelection = gradeSelectForStats ? gradeSelectForStats.value : "";
  if (!gradeSelection) {
    alert("请先选择要统计的年级（或初中/高中/全年级总计）。");
    return;
  }

  // 读取每一天选择的时段
  const dayModes = {}; // { weekdayNumber: 'am' | 'pm' | 'all' | '' }
  for (let w = 1; w <= 5; w++) {
    const checked = document.querySelector(
      `input[name="day-${w}-mode"]:checked`
    );
    dayModes[w] = checked ? checked.value : "";
  }

  const selectedWeekdays = Object.keys(dayModes).filter(
    w => dayModes[w] && dayModes[w] !== ""
  );
  if (!selectedWeekdays.length) {
    alert("请至少为周一至周五中的某一天选择上午/下午/全天之一。");
    return;
  }

  let allowedGrades = [];
  if (gradeSelection === "__JUNIOR__") {
    allowedGrades = ["初一", "初二", "初三"];
  } else if (gradeSelection === "__SENIOR__") {
    allowedGrades = ["高一", "高二", "高三"];
  } else if (gradeSelection === "__ALL__") {
    allowedGrades = ["初一", "初二", "初三", "高一", "高二", "高三"];
  } else {
    allowedGrades = [gradeSelection];
  }

  const filtered = state.gradeDayStats.filter(
    item => allowedGrades.includes(item.grade) && dayModes[item.weekday]
  );

  if (!filtered.length) {
    gradeStatsResult.textContent = `范围【${getGradeSelectionLabel(
      gradeSelection
    )}】在所选日期和时段内没有任何老师的课时记录。`;
    gradeStatsTableBody.innerHTML =
      '<tr><td colspan="7" style="text-align:center;color:#9ca3af;">暂无统计结果。</td></tr>';
    setGradeStatsSearchEnabled(false);
    return;
  }

  // 汇总：每个老师在每个工作日的总节数，以及总节数
  const teacherMap = {}; // { teacher: { teacher, byWeekday:{1:n,...}, total:n } }

  filtered.forEach(item => {
    const mode = dayModes[item.weekday];
    if (!mode) return;

    if (!teacherMap[item.teacher]) {
      teacherMap[item.teacher] = {
        teacher: item.teacher,
        byWeekday: { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0 },
        total: 0
      };
    }

    let add = 0;
    if (mode === "am") add = item.am;
    else if (mode === "pm") add = item.pm;
    else if (mode === "all") add = item.am + item.pm;

    teacherMap[item.teacher].byWeekday[item.weekday] += add;
    teacherMap[item.teacher].total += add;
  });

  const rows = Object.values(teacherMap).filter(t => t.total > 0);
  if (!rows.length) {
    gradeStatsResult.textContent = `范围【${getGradeSelectionLabel(
      gradeSelection
    )}】在所选日期和时段内所有老师的课时为 0 节。`;
    gradeStatsTableBody.innerHTML =
      '<tr><td colspan="7" style="text-align:center;color:#9ca3af;">暂无统计结果。</td></tr>';
    setGradeStatsSearchEnabled(false);
    return;
  }

  rows.sort((a, b) => b.total - a.total);

  // 渲染结果表
  gradeStatsTableBody.innerHTML = "";
  rows.forEach(row => {
    const tr = document.createElement("tr");
    tr.dataset.teacher = row.teacher;
    tr.innerHTML = `
      <td>${row.teacher}</td>
      <td>
        <input
          type="number"
          min="0"
          step="0.5"
          value="${row.total}"
          data-teacher="${row.teacher}"
          data-weekday="0"
          class="result-edit"
          style="width:78px;"
          title="可手动修改总节数（会同步分配到周一至周五按比例调整）"
        />
      </td>
      <td>
        <input
          type="number"
          min="0"
          step="0.5"
          value="${row.byWeekday[1] || 0}"
          data-teacher="${row.teacher}"
          data-weekday="1"
          class="result-edit"
          style="width:66px;"
        />
      </td>
      <td>
        <input
          type="number"
          min="0"
          step="0.5"
          value="${row.byWeekday[2] || 0}"
          data-teacher="${row.teacher}"
          data-weekday="2"
          class="result-edit"
          style="width:66px;"
        />
      </td>
      <td>
        <input
          type="number"
          min="0"
          step="0.5"
          value="${row.byWeekday[3] || 0}"
          data-teacher="${row.teacher}"
          data-weekday="3"
          class="result-edit"
          style="width:66px;"
        />
      </td>
      <td>
        <input
          type="number"
          min="0"
          step="0.5"
          value="${row.byWeekday[4] || 0}"
          data-teacher="${row.teacher}"
          data-weekday="4"
          class="result-edit"
          style="width:66px;"
        />
      </td>
      <td>
        <input
          type="number"
          min="0"
          step="0.5"
          value="${row.byWeekday[5] || 0}"
          data-teacher="${row.teacher}"
          data-weekday="5"
          class="result-edit"
          style="width:66px;"
        />
      </td>
    `;
    gradeStatsTableBody.appendChild(tr);
  });

  // 缓存本次统计结果，方便一键导出
  lastGradeStatsCache = {
    gradeSelection,
    dayModes,
    rows
  };

  // 绑定编辑事件（修改后即时刷新总计）
  gradeStatsTableBody.querySelectorAll("input.result-edit").forEach(input => {
    input.addEventListener("change", onGradeStatsCellEdit);
  });

  // 初次渲染后刷新一次汇总文本
  refreshGradeStatsSummaryText();
  setGradeStatsSearchEnabled(true);
  applyGradeStatsSearch();
}

function onGradeStatsCellEdit(e) {
  if (!lastGradeStatsCache) return;
  const input = e.target;
  const teacher = input.dataset.teacher;
  const weekday = parseInt(input.dataset.weekday, 10); // 0 表示总节数

  let value = parseFloat(input.value);
  if (isNaN(value) || value < 0) value = 0;

  const row = lastGradeStatsCache.rows.find(r => r.teacher === teacher);
  if (!row) return;

  if (weekday === 0) {
    // 手改总节数：按当前周一~周五的占比重新分配（若全为0，则全部放到周一）
    const currentSum =
      (row.byWeekday[1] || 0) +
      (row.byWeekday[2] || 0) +
      (row.byWeekday[3] || 0) +
      (row.byWeekday[4] || 0) +
      (row.byWeekday[5] || 0);

    if (currentSum <= 0) {
      row.byWeekday = { 1: value, 2: 0, 3: 0, 4: 0, 5: 0 };
    } else {
      const ratio = value / currentSum;
      for (let w = 1; w <= 5; w++) {
        const next = (row.byWeekday[w] || 0) * ratio;
        row.byWeekday[w] = Math.round(next * 2) / 2; // 以0.5为最小粒度
      }
    }
  } else {
    row.byWeekday[weekday] = value;
  }

  // 重算该老师总节数
  row.total =
    (row.byWeekday[1] || 0) +
    (row.byWeekday[2] || 0) +
    (row.byWeekday[3] || 0) +
    (row.byWeekday[4] || 0) +
    (row.byWeekday[5] || 0);

  // 回写到表格（同步总节数与各天）
  const tr = input.closest("tr");
  if (tr) {
    const totalInput = tr.querySelector('input.result-edit[data-weekday="0"]');
    if (totalInput) totalInput.value = row.total;
    for (let w = 1; w <= 5; w++) {
      const dayInput = tr.querySelector(
        `input.result-edit[data-weekday="${w}"]`
      );
      if (dayInput) dayInput.value = row.byWeekday[w] || 0;
    }
  }

  refreshGradeStatsSummaryText();
}

function refreshGradeStatsSummaryText() {
  if (!gradeStatsResult || !lastGradeStatsCache) return;
  const { gradeSelection, rows } = lastGradeStatsCache;
  const teacherCount = rows.length;
  const totalLessons = rows.reduce((sum, r) => sum + (r.total || 0), 0);
  gradeStatsResult.textContent = `范围【${getGradeSelectionLabel(
    gradeSelection
  )}】在所选日期与时段内共有 ${teacherCount} 位老师，共 ${totalLessons} 节课。（表格数值可修改，导出以修改后的为准）`;
}

function setGradeStatsSearchEnabled(enabled) {
  if (!gradeStatsTeacherSearch) return;
  gradeStatsTeacherSearch.disabled = !enabled;
  if (gradeStatsTeacherClear) {
    gradeStatsTeacherClear.disabled = !enabled;
  }
  if (!enabled) {
    gradeStatsTeacherSearch.value = "";
    if (gradeStatsTeacherTip) gradeStatsTeacherTip.textContent = "";
  }
  updateGradeStatsSearchControls();
}

function updateGradeStatsSearchControls() {
  if (!gradeStatsTeacherSearch || !gradeStatsTeacherClear) return;
  if (gradeStatsTeacherSearch.disabled) {
    gradeStatsTeacherClear.disabled = true;
    return;
  }
  gradeStatsTeacherClear.disabled =
    gradeStatsTeacherSearch.value.trim().length === 0;
}

function applyGradeStatsSearch() {
  if (!gradeStatsTableBody || !lastGradeStatsCache) return;
  const keyword = (gradeStatsTeacherSearch?.value || "")
    .trim()
    .toLowerCase();

  let total = 0;
  let visible = 0;

  gradeStatsTableBody.querySelectorAll("tr").forEach(tr => {
    const teacher = tr.dataset.teacher || "";
    if (!teacher) return;
    total += 1;
    const match = !keyword || teacher.toLowerCase().includes(keyword);
    tr.style.display = match ? "" : "none";
    if (match) visible += 1;
  });

  if (gradeStatsTeacherTip) {
    if (!keyword) {
      gradeStatsTeacherTip.textContent = "";
    } else if (total === 0) {
      gradeStatsTeacherTip.textContent = "暂无统计结果。";
    } else if (visible === 0) {
      gradeStatsTeacherTip.textContent = "未找到匹配老师。";
    } else {
      gradeStatsTeacherTip.textContent = `已筛选 ${visible}/${total} 位老师`;
    }
  }

  updateGradeStatsSearchControls();
}

function onExportGradeStats() {
  if (typeof XLSX === "undefined") {
    alert("Excel 导出库未加载成功，请检查网络后刷新页面重试。");
    return;
  }

  if (!lastGradeStatsCache || !lastGradeStatsCache.rows.length) {
    alert("请先点击“统计该年级老师课时”生成结果后，再导出Excel。");
    return;
  }

  const { gradeSelection, rows } = lastGradeStatsCache;
  const scopeLabel = getGradeSelectionLabel(gradeSelection);

  const header = ["教师", "总节数", "周一", "周二", "周三", "周四", "周五"];
  const dataRows = rows.map(r => [
    r.teacher,
    r.total,
    r.byWeekday[1] || 0,
    r.byWeekday[2] || 0,
    r.byWeekday[3] || 0,
    r.byWeekday[4] || 0,
    r.byWeekday[5] || 0
  ]);

  const sheet = XLSX.utils.aoa_to_sheet([header, ...dataRows]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, sheet, "老师课时统计");

  const fileName = `${scopeLabel}_老师课时统计.xlsx`;
  XLSX.writeFile(wb, fileName);
}

function onApplyHoursToTemplate() {
  if (typeof XLSX === "undefined") {
    alert("Excel 解析库未加载成功，请检查网络后刷新页面重试。");
    return;
  }

  if (!lastGradeStatsCache || !lastGradeStatsCache.rows.length) {
    alert("请先完成“统计该年级老师课时”，再提交模板 Excel。");
    return;
  }

  if (!teacherHoursTemplateInput || !teacherHoursTemplateInput.files.length) {
    alert("请先选择要上传的模板 Excel 文件（.xls 或 .xlsx）。");
    return;
  }

  const file = teacherHoursTemplateInput.files[0];
  const teacherTotalMap = {};
  lastGradeStatsCache.rows.forEach(row => {
    const key = normalizeTeacherName(row.teacher);
    if (!key) return;
    teacherTotalMap[key] = row.total || 0;
  });

  const reader = new FileReader();
  reader.onload = function (e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array", cellStyles: true });

      let updatedCount = 0;
      let hitCount = 0;

      workbook.SheetNames.forEach(sheetName => {
        const sheet = workbook.Sheets[sheetName];
        if (!sheet || !sheet["!ref"]) return;

        const matrix = getSheetMatrix(sheet);
        const nameHeader = findNameHeaderCell(matrix);
        if (!nameHeader) return;

        const actualCol = findTotalActualHoursColumn(sheet, matrix);
        if (actualCol === null || typeof actualCol === "undefined") return;

        const startRow = nameHeader.row + 1;
        const endRow = XLSX.utils.decode_range(sheet["!ref"]).e.r;

        for (let r = startRow; r <= endRow; r++) {
          const nameCellAddr = XLSX.utils.encode_cell({
            r,
            c: nameHeader.col
          });
          const nameCell = sheet[nameCellAddr];
          const rawName = nameCell ? nameCell.v : "";
          const normalized = normalizeTeacherName(rawName);
          if (!normalized) continue;
          if (!teacherTotalMap.hasOwnProperty(normalized)) continue;

          hitCount += 1;
          const addValue = teacherTotalMap[normalized] || 0;

          const actualAddr = XLSX.utils.encode_cell({ r, c: actualCol });
          const actualCell = sheet[actualAddr];
          let base = 0;
          if (actualCell && actualCell.v !== null && actualCell.v !== "") {
            const parsed = parseFloat(actualCell.v);
            if (!isNaN(parsed)) base = parsed;
          }
          const next = base + addValue;
          if (actualCell) {
            actualCell.v = next;
            actualCell.t = "n";
          } else {
            sheet[actualAddr] = { t: "n", v: next };
          }
          updatedCount += 1;
        }
      });

      if (!hitCount) {
        alert(
          "未在模板中找到匹配的老师姓名或未识别到“总课时-实上”列，请检查模板格式。"
        );
        return;
      }

      const baseName = file.name.replace(/(\.xlsx?|\.xlsm)$/i, "");
      const fileName = `${baseName}_已填课时.xlsx`;
      XLSX.writeFile(workbook, fileName);
      alert(`已更新 ${updatedCount} 行老师课时，文件已生成。`);
    } catch (err) {
      console.error("模板写入失败：", err);
      alert("模板写入失败，请检查Excel格式后重试。");
    }
  };

  reader.onerror = function () {
    alert("读取文件失败，请重试。");
  };

  reader.readAsArrayBuffer(file);
}

function getGradeSelectionLabel(value) {
  if (value === "__JUNIOR__") return "初中总计";
  if (value === "__SENIOR__") return "高中总计";
  if (value === "__ALL__") return "全年级总计";
  return value || "未选择";
}

function onClearGradeStats() {
  if (
    !state.gradeDayStats ||
    !state.gradeDayStats.length
  ) {
    alert("当前没有已导入的老师课程统计数据。");
    return;
  }

  if (!window.confirm("确认要清空所有已导入的老师课程统计数据吗？此操作只影响本系统中的统计结果，不会修改原Excel文件。")) {
    return;
  }

  state.gradeDayStats = [];
  saveState();

  // 重置界面显示
  if (gradeStatsTableBody) {
    gradeStatsTableBody.innerHTML =
      '<tr><td colspan="7" style="text-align:center;color:#9ca3af;">暂无统计结果。</td></tr>';
  }
  if (gradeStatsResult) {
    gradeStatsResult.textContent =
      "已清空导入的数据，请重新导入统计表后再进行统计。";
  }
  setGradeStatsSearchEnabled(false);

  if (gradeSelectForStats) {
    gradeSelectForStats.value = "";
  }

  // 恢复周一到周五的“不统计”默认选项
  for (let w = 1; w <= 5; w++) {
    const radio = document.querySelector(
      `input[name="day-${w}-mode"][value=""]`
    );
    if (radio) {
      radio.checked = true;
    }
  }
}

function renderTimetableStats() {
  if (!timetableStatsBody) return;

  timetableStatsBody.innerHTML = "";

  const stats =
    (state.timetableStats && state.timetableStats.teacherStats) || [];

  if (!stats.length) {
    timetableStatsBody.innerHTML =
      '<tr><td colspan="5" style="text-align:center;color:#9ca3af;">尚未导入课表 Excel 或未识别到教师课表数据。</td></tr>';
    return;
  }

  // 按“应上”从高到低排序
  const sorted = [...stats].sort((a, b) => b.expected - a.expected);

  sorted.forEach(item => {
    const diff = (item.actual || 0) - (item.expected || 0);
    const diffText = (diff >= 0 ? "+" : "") + diff;

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${item.name}</td>
      <td>${item.expected}</td>
      <td>
        <input
          type="number"
          min="0"
          step="0.5"
          value="${item.actual}"
          data-name="${item.name}"
          class="actual-input"
          style="width:80px;"
        />
      </td>
      <td>${diffText}</td>
      <td>
        <button class="btn btn-xs" data-action="minus" data-name="${
          item.name
        }">-1节</button>
        <button class="btn btn-xs" data-action="plus" data-name="${
          item.name
        }">+1节</button>
      </td>
    `;

    timetableStatsBody.appendChild(tr);
  });

  // 绑定“实上”直接编辑
  timetableStatsBody
    .querySelectorAll("input.actual-input")
    .forEach(input => {
      input.addEventListener("change", onActualInputChange);
    });

  // 绑定加减按钮
  timetableStatsBody.querySelectorAll("button[data-action]").forEach(btn => {
    btn.addEventListener("click", onAdjustActualClick);
  });
}

function findTeacherStatByName(name) {
  if (!state.timetableStats || !state.timetableStats.teacherStats) return null;
  return (
    state.timetableStats.teacherStats.find(item => item.name === name) || null
  );
}

function onActualInputChange(e) {
  const input = e.target;
  const name = input.dataset.name;
  if (!name) return;

  let value = parseFloat(input.value);
  if (isNaN(value) || value < 0) {
    value = 0;
  }

  const stat = findTeacherStatByName(name);
  if (!stat) return;

  stat.actual = value;
  saveState();
  renderTimetableStats();
}

function onAdjustActualClick(e) {
  const btn = e.currentTarget;
  const name = btn.dataset.name;
  const action = btn.dataset.action;
  if (!name || !action) return;

  const stat = findTeacherStatByName(name);
  if (!stat) return;

  const delta = action === "plus" ? 1 : -1;
  const next = Math.max(0, (stat.actual || 0) + delta);
  stat.actual = next;
  saveState();
  renderTimetableStats();
}

function onExportStats() {
  if (typeof XLSX === "undefined") {
    alert("Excel 导出库未加载成功，请检查网络后刷新页面重试。");
    return;
  }

  const stats =
    (state.timetableStats && state.timetableStats.teacherStats) || [];
  if (!stats.length) {
    alert("暂无可导出的统计数据，请先导入并解析课表。");
    return;
  }

  const header = ["教师", "应上（节/周）", "实上（节/周）", "差额（实上-应上）"];
  const rows = stats.map(item => {
    const diff = (item.actual || 0) - (item.expected || 0);
    return [item.name, item.expected, item.actual, diff];
  });

  const sheet = XLSX.utils.aoa_to_sheet([header, ...rows]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, sheet, "教师课时统计");

  let baseName =
    (state.timetableStats && state.timetableStats.sourceFileName) || "课表";
  baseName = baseName.replace(/(\.xlsx?|\.xlsm)$/i, "");

  const fileName = `${baseName}_教师课时统计.xlsx`;
  XLSX.writeFile(wb, fileName);
}

/* 上课记录 */

function syncClassRateWithTeacher() {
  const teacherId = classTeacherSelect.value;
  const teacher = findTeacherById(teacherId);
  if (teacher) {
    classRateInput.value = teacher.rate;
  }
}

function onAddClass(e) {
  e.preventDefault();

  const teacherId = classTeacherSelect.value;
  const teacher = findTeacherById(teacherId);
  if (!teacher) {
    alert("请选择合法的教师");
    return;
  }

  const date = classDateInput.value;
  const course = classCourseInput.value.trim();
  const hours = parseFloat(classHoursInput.value);
  const rate = parseFloat(classRateInput.value);
  const remark = classRemarkInput.value.trim();

  if (!date) {
    alert("上课日期不能为空");
    return;
  }
  if (!course) {
    alert("课程名称不能为空");
    return;
  }
  if (isNaN(hours) || hours <= 0) {
    alert("课时数必须为大于 0 的数字");
    return;
  }
  if (isNaN(rate) || rate < 0) {
    alert("单价必须为非负数字");
    return;
  }

  const record = {
    id: createId(),
    teacherId,
    date,
    course,
    hours,
    rate,
    remark
  };

  state.classes.push(record);
  saveState();
  renderClasses();

  classForm.reset();
  // 重置后再同步默认的教师和单价
  if (state.teachers[0]) {
    classTeacherSelect.value = state.teachers[0].id;
    syncClassRateWithTeacher();
  }
  classDateInput.value = new Date().toISOString().slice(0, 10);
}

function renderClasses() {
  if (!classTableBody) return;

  classTableBody.innerHTML = "";

  if (state.classes.length === 0) {
    classTableBody.innerHTML =
      '<tr><td colspan="7" style="text-align:center;color:#9ca3af;">暂时没有上课记录。</td></tr>';
    return;
  }

  const teacherFilter = filterTeacherSelect.value;
  const monthFilter = filterMonthInput.value; // 形如 2026-03

  let filtered = [...state.classes];

  if (teacherFilter) {
    filtered = filtered.filter(c => c.teacherId === teacherFilter);
  }

  if (monthFilter) {
    filtered = filtered.filter(c => c.date && c.date.startsWith(monthFilter));
  }

  // 按日期倒序显示
  filtered.sort((a, b) => (a.date < b.date ? 1 : -1));

  if (filtered.length === 0) {
    classTableBody.innerHTML =
      '<tr><td colspan="7" style="text-align:center;color:#9ca3af;">符合条件的记录为空。</td></tr>';
    return;
  }

  filtered.forEach(record => {
    const teacher = findTeacherById(record.teacherId);
    const teacherName = teacher ? teacher.name : "（已删除教师）";

    const subtotal = record.hours * record.rate;

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${record.date}</td>
      <td>${teacherName}</td>
      <td>${record.course}</td>
      <td>${record.hours}</td>
      <td>${formatCurrency(record.rate)}</td>
      <td>${formatCurrency(subtotal)}</td>
      <td>${record.remark || "-"}</td>
    `;
    classTableBody.appendChild(tr);
  });
}

/* 工资统计 */

function onCalcSalary() {
  const teacherId = salaryTeacherSelect.value;
  const month = salaryMonthInput.value; // 形如 2026-03

  if (!teacherId) {
    alert("请先选择教师");
    return;
  }
  if (!month) {
    alert("请先选择月份");
    return;
  }

  const teacher = findTeacherById(teacherId);
  const teacherName = teacher ? teacher.name : "（已删除教师）";

  // 过滤出该教师、该月份的所有记录
  const records = state.classes.filter(
    c => c.teacherId === teacherId && c.date && c.date.startsWith(month)
  );

  if (records.length === 0) {
    salaryResult.textContent = `${month}，教师【${teacherName}】没有上课记录。`;
    return;
  }

  let totalHours = 0;
  let totalPay = 0;

  records.forEach(r => {
    totalHours += r.hours;
    totalPay += r.hours * r.rate;
  });

  const details = records
    .map(
      r =>
        `${r.date}「${r.course}」${r.hours} 课时 × ${formatCurrency(
          r.rate
        )} 元 = ${formatCurrency(r.hours * r.rate)} 元`
    )
    .join("<br>");

  salaryResult.innerHTML = `
    <div><strong>${month} 教师【${teacherName}】工资统计：</strong></div>
    <div>上课次数：${records.length} 次</div>
    <div>课时总数：${totalHours} 课时</div>
    <div>应发工资：<strong>${formatCurrency(totalPay)} 元</strong></div>
    <hr style="border:none;border-top:1px dashed #d1d5db;margin:6px 0;" />
    <div style="font-size:13px;color:#4b5563;">明细：</div>
    <div style="font-size:13px;color:#4b5563;margin-top:2px;">${details}</div>
  `;
}
