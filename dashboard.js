/* ==============================================================================
   PIVC BYPASS EXECUTIVE DASHBOARD - CORE JAVASCRIPT ENGINE
   Implements exact Power BI DAX & Calculated Column Logic client-side
   ============================================================================== */

// Baseline production dataset generator matching PIVC_bypass_500_records.csv (500 records)
function generateBase500Records() {
  const records = [];
  for (let i = 0; i < 500; i++) {
    const isBypass = (i % 5 === 0);
    const reqNo = String(15961850 + i * 3);
    const appNo = String(6119374000 + i * 17);
    const remarks = isBypass ? `Bypass request-AS-${715441 + Math.floor(i / 5) * 5}` : "";
    const substatus = isBypass ? "Exceptional" : "DPIVC DONE";
    records.push({
      REQUEST_NO: reqNo,
      REMARKS: remarks,
      USER_ID: "WEO_SYS_CONNECT",
      APPROVAL_SUBSTATUS: substatus,
      APPLICATION_NO: appNo
    });
  }
  return records;
}

const RAW_DATA = generateBase500Records();

// Months array for chronological sorting
const MONTH_NAMES = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

// Active processed dataset (initialized with RAW_DATA, replaced when user uploads Excel)
let processedData = [];

// Helper: robust date parser for Excel dates, serials, and strings
function parseCustomDate(val) {
  if (!val) return new Date();
  if (val instanceof Date) return val;

  // Handle Excel numeric serial dates (e.g. 45321)
  if (typeof val === "number" || (!isNaN(val) && !isNaN(parseFloat(val)) && !String(val).includes("-") && !String(val).includes("/"))) {
    const serial = parseFloat(val);
    // Excel base date Dec 30, 1899
    const utcDays = Math.floor(serial - 25569);
    const utcValue = utcDays * 86400;
    const dateInfo = new Date(utcValue * 1000);
    return isNaN(dateInfo.getTime()) ? new Date() : dateInfo;
  }

  // Handle String dates
  const str = String(val).trim();
  // Try direct parse
  let d = new Date(str);
  if (!isNaN(d.getTime())) return d;

  // Try DD-MM-YYYY or DD/MM/YYYY
  const parts = str.split(/[-/]/);
  if (parts.length === 3) {
    if (parts[0].length === 2 && parts[2].length === 4) {
      // DD-MM-YYYY
      d = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
      if (!isNaN(d.getTime())) return d;
    } else if (parts[0].length === 4) {
      // YYYY-MM-DD
      d = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
      if (!isNaN(d.getTime())) return d;
    }
  }

  return new Date();
}

// Format date to YYYY-MM-DD for table display
function formatDateDisplay(d) {
  if (!d || isNaN(d.getTime())) return "";
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

// Normalize key search for flexible matching (e.g. "Application No", "application_no", "APPLICATION_NO")
function findField(row, candidates) {
  const keys = Object.keys(row);
  for (const c of candidates) {
    const cleanCand = c.toUpperCase().replace(/[\s_]+/g, "");
    for (const k of keys) {
      const cleanK = k.toUpperCase().replace(/[\s_]+/g, "");
      if (cleanK === cleanCand) {
        return row[k];
      }
    }
  }
  return "";
}

// Convert raw JSON rows (from Excel or initial array) into normalized records
function transformRows(rawArray) {
  const total = Math.max(rawArray.length, 1);
  const DEFAULT_ASSIGNEES = ["David Kumar", "Sarah Jenkins", "Robert Chen", "Elena Rostova", "Aisha Patel", "Priya Sharma"];

  return rawArray.map((row, idx) => {
    const appNo = String(findField(row, ["APPLICATION_NO", "APPLICATION NO", "APPLICATIONNO", "APP_NO", "APPLICATION_NUMBER", "REQUEST_NO", "REQUESTNO"]) || row.APPLICATION_NO || row.REQUEST_NO || `APP-${1000 + idx}`).trim();
    const reqNo = String(findField(row, ["REQUEST_NO", "REQUEST NO", "REQUESTNO", "REQ_NO"]) || row.REQUEST_NO || "").trim();
    const rawDate = findField(row, ["APPROVAL_DATE", "APPROVAL DATE", "APPROVALDATE", "DATE", "DATE_OF_APPROVAL", "REQUEST_DATE", "DATE_TIME", "CREATED_ON"]);
    
    // Flexible Approval Status discovery & Null handling (supports APPROVAL_SUBSTATUS, APPROVAL_STATUS, etc.)
    let rawStatus = findField(row, ["APPROVAL_SUBSTATUS", "APPROVAL SUBSTATUS", "SUBSTATUS", "APPROVAL_STATUS", "APPROVAL STATUS", "APPROVALSTATUS", "APPROVAL", "STATUS", "APPROVAL_STATE", "APPROVAL_EXCEPTION"]);
    if (rawStatus === undefined || rawStatus === null) {
      rawStatus = row.APPROVAL_SUBSTATUS || row.APPROVAL_STATUS;
    }
    const statusStr = (rawStatus !== undefined && rawStatus !== null) ? String(rawStatus).trim() : "";
    const isStatusNull = !statusStr || statusStr.toLowerCase() === "null" || statusStr.toLowerCase() === "undefined";
    const status = isStatusNull ? "" : statusStr;

    const userId = String(findField(row, ["USER_ID", "USER ID", "USERID", "AGENT_ID", "USER"]) || row.USER_ID || "USR_Unknown").trim();
    
    // Assignee mapping: use assigned field if present, or analyst rotation
    let assignee = String(findField(row, ["ASSIGNED_TO", "ASSIGNED TO", "ASSIGNEDTO", "ASSIGNEE", "ANALYST"]) || row.ASSIGNED_TO || "").trim();
    if (!assignee) {
      assignee = DEFAULT_ASSIGNEES[idx % DEFAULT_ASSIGNEES.length];
    }

    const remarks = String(findField(row, ["REMARKS", "REMARK", "COMMENTS", "COMMENT", "NOTES"]) || row.REMARKS || "").trim();

    // Dual-Condition Business Rules for Bypass Identification:
    // Condition 1: APPROVAL_SUBSTATUS / APPROVAL_STATUS contains "except" (matches "Exceptional" or "Exception")
    // Condition 2: REMARKS contains "bypass" (matches "Bypass request-AS-...", "Bypass request", etc.)
    // When status is "DPIVC DONE" or null/blank and remarks do not contain bypass -> "Regular" (System Done)
    const isExceptionStatus = !isStatusNull && status.toLowerCase().includes("except");
    const isBypassRemark = remarks.toLowerCase().includes("bypass");
    const isBypass = isExceptionStatus || isBypassRemark;

    let bypassTrigger = "System Done (Regular)";
    if (isExceptionStatus && isBypassRemark) {
      bypassTrigger = "Status: Exceptional & Remarks";
    } else if (isExceptionStatus) {
      bypassTrigger = `Status: ${status}`;
    } else if (isBypassRemark) {
      bypassTrigger = "Remarks: Bypass Request";
    }

    // Date Resolution: if input file doesn't have dates, distribute smoothly across Jan - Sep 2024
    let dateObj;
    if (rawDate && String(rawDate).trim() !== "" && String(rawDate).trim().toLowerCase() !== "undefined") {
      dateObj = parseCustomDate(rawDate);
    } else {
      const startDate = new Date(2024, 0, 3);
      const totalDays = 270;
      const dayOffset = Math.floor((idx / total) * totalDays);
      dateObj = new Date(startDate.getTime() + dayOffset * 86400 * 1000);
      if (dateObj.getDay() === 0) dateObj.setDate(dateObj.getDate() + 1);
      if (dateObj.getDay() === 6) dateObj.setDate(dateObj.getDate() + 2);
    }

    const year = dateObj.getFullYear();
    const monthIdx = dateObj.getMonth();
    const monthStr = MONTH_NAMES[monthIdx];
    const monthYear = `${monthStr} ${year}`;
    const monthYearSort = year * 100 + (monthIdx + 1);

    return {
      APPLICATION_NO: appNo,
      REQUEST_NO: reqNo,
      APPROVAL_DATE: formatDateDisplay(dateObj),
      APPROVAL_STATUS: status, // e.g. "Exceptional", "DPIVC DONE", or ""
      IS_STATUS_NULL: isStatusNull,
      USER_ID: userId,
      ASSIGNED_TO: assignee,
      REMARKS: remarks,
      "Bypass Flag": isBypass ? "Bypass" : "Regular",
      "Bypass Trigger": bypassTrigger,
      "Month-Year": monthYear,
      "Month-Year Sort": monthYearSort,
      "DateObj": dateObj
    };
  });
}

// Initial transformation of mock data
processedData = transformRows(RAW_DATA);

// State for active slicers
let activeFilters = {
  monthYear: "All",
  approvalStatus: "All",
  userId: "All",
  assignedTo: "All"
};

// Interactive Chart Mode States
let currentLineChartMode = "pct"; // "pct" | "dual" | "cumulative"
let currentStackedBarMode = "stacked"; // "stacked" | "100" | "clustered"
let currentPieViewMode = "donut"; // "donut" | "solid" | "reason"
let currentUserMetricMode = "bypass"; // "bypass" | "total" | "rate"
let currentAssigneeMetricMode = "bypass"; // "bypass" | "total" | "rate"


// Pagination State
let currentPage = 1;
let pageSize = 50;

// DOM Elements
const monthSelect = document.getElementById("slicerMonth");
const statusSelect = document.getElementById("slicerStatus");
const userSelect = document.getElementById("slicerUser");
const assigneeSelect = document.getElementById("slicerAssignee");
const resetFiltersBtn = document.getElementById("resetFiltersBtn");
const searchInput = document.getElementById("searchTable");

// KPI Elements
const kpiTotalApps = document.getElementById("kpiTotalApps");
const kpiBypassApps = document.getElementById("kpiBypassApps");
const kpiBypassPct = document.getElementById("kpiBypassPct");
const kpiRegularApps = document.getElementById("kpiRegularApps");
const kpiBadgePct = document.getElementById("kpiBadgePct");

// Chart Containers
const barChartMonthly = document.getElementById("barChartMonthly");
const lineChartTrend = document.getElementById("lineChartTrend");
const userBarList = document.getElementById("userBarList");
const assigneeBarList = document.getElementById("assigneeBarList");
const donutSvg = document.getElementById("donutSvg");
const donutPctText = document.getElementById("donutPctText");
const donutBypassLegend = document.getElementById("donutBypassLegend");
const donutRegularLegend = document.getElementById("donutRegularLegend");
const matrixTableBody = document.getElementById("matrixTableBody");
const drilldownTableBody = document.getElementById("drilldownTableBody");
const recordCountEl = document.getElementById("recordCount");

// Dynamic Slicer Re-population
function populateSlicers() {
  // Clear existing options (keep "All")
  monthSelect.innerHTML = `<option value="All">All Months</option>`;
  statusSelect.innerHTML = `<option value="All">All Statuses</option>`;
  userSelect.innerHTML = `<option value="All">All Users</option>`;
  assigneeSelect.innerHTML = `<option value="All">All Assignees</option>`;

  // 1. Month-Year slicer (Chronologically sorted by sort key)
  const uniqueMonths = [...new Map(processedData.map(d => [d["Month-Year Sort"], d["Month-Year"]])).entries()]
    .sort((a, b) => a[0] - b[0])
    .map(entry => entry[1]);

  uniqueMonths.forEach(m => {
    const opt = document.createElement("option");
    opt.value = m;
    opt.textContent = m;
    monthSelect.appendChild(opt);
  });

  populateV1MonthPicker();

  // 2. Approval Status
  const statusSet = new Set();
  let hasNullStatus = false;
  processedData.forEach(d => {
    if (!d.APPROVAL_STATUS || d.APPROVAL_STATUS.trim() === "" || d.IS_STATUS_NULL) {
      hasNullStatus = true;
    } else {
      statusSet.add(d.APPROVAL_STATUS);
    }
  });

  const uniqueStatuses = [...statusSet].sort();
  // Pin Exception / Exceptional at the top if present
  const excIdx = uniqueStatuses.findIndex(s => s.toLowerCase().includes("except"));
  if (excIdx > -1) {
    const [exc] = uniqueStatuses.splice(excIdx, 1);
    uniqueStatuses.unshift(exc);
  }

  uniqueStatuses.forEach(s => {
    const opt = document.createElement("option");
    opt.value = s;
    if (s.toLowerCase().includes("except")) {
      opt.textContent = `⚠️ ${s} (Bypass)`;
    } else if (s.toUpperCase().includes("DPIVC") || s.toUpperCase().includes("DONE")) {
      opt.textContent = `✓ ${s} (System Done)`;
    } else {
      opt.textContent = s;
    }
    statusSelect.appendChild(opt);
  });

  if (hasNullStatus) {
    const opt = document.createElement("option");
    opt.value = "__NULL__";
    opt.textContent = "Null / Blank (System PIVC Done)";
    statusSelect.appendChild(opt);
  }

  // 3. User ID
  const uniqueUsers = [...new Set(processedData.map(d => d.USER_ID))].filter(Boolean).sort();
  uniqueUsers.forEach(u => {
    const opt = document.createElement("option");
    opt.value = u;
    opt.textContent = u;
    userSelect.appendChild(opt);
  });

  // 4. Assigned To
  const uniqueAssignees = [...new Set(processedData.map(d => d.ASSIGNED_TO))].filter(Boolean).sort();
  uniqueAssignees.forEach(a => {
    const opt = document.createElement("option");
    opt.value = a;
    opt.textContent = a;
    assigneeSelect.appendChild(opt);
  });
}

function initSlicers() {
  populateSlicers();

  // Event Listeners
  monthSelect.addEventListener("change", e => { 
    activeFilters.monthYear = e.target.value; 
    v1SelectedMonth = e.target.value;
    const v1Picker = document.getElementById("v1MonthPicker");
    if (v1Picker) v1Picker.value = e.target.value;
    currentPage = 1; 
    updateDashboard(); 
  });
  statusSelect.addEventListener("change", e => { activeFilters.approvalStatus = e.target.value; currentPage = 1; updateDashboard(); });
  userSelect.addEventListener("change", e => { activeFilters.userId = e.target.value; currentPage = 1; updateDashboard(); });
  assigneeSelect.addEventListener("change", e => { activeFilters.assignedTo = e.target.value; currentPage = 1; updateDashboard(); });

  resetFiltersBtn.addEventListener("click", () => {
    activeFilters = { monthYear: "All", approvalStatus: "All", userId: "All", assignedTo: "All" };
    v1SelectedMonth = "All";
    const v1Picker = document.getElementById("v1MonthPicker");
    if (v1Picker) v1Picker.value = "All";
    monthSelect.value = "All";
    statusSelect.value = "All";
    userSelect.value = "All";
    assigneeSelect.value = "All";
    currentPage = 1;
    if (searchInput) searchInput.value = "";
    updateDashboard();
  });

  if (searchInput) {
    searchInput.addEventListener("input", () => {
      currentPage = 1;
      renderDrilldownTable(getFilteredData());
    });
  }
}

// Filter dataset based on active slicers
function getFilteredData() {
  return processedData.filter(item => {
    if (activeFilters.monthYear !== "All" && item["Month-Year"] !== activeFilters.monthYear) return false;
    if (activeFilters.approvalStatus !== "All") {
      if (activeFilters.approvalStatus === "__NULL__") {
        if (item.APPROVAL_STATUS && item.APPROVAL_STATUS.trim() !== "" && !item.IS_STATUS_NULL) return false;
      } else if (item.APPROVAL_STATUS !== activeFilters.approvalStatus) {
        return false;
      }
    }
    if (activeFilters.userId !== "All" && item.USER_ID !== activeFilters.userId) return false;
    if (activeFilters.assignedTo !== "All" && item.ASSIGNED_TO !== activeFilters.assignedTo) return false;
    return true;
  });
}

// Compute DAX Measures
function computeMeasures(data) {
  // Total Applications = Distinct count of APPLICATION_NO
  const totalApps = new Set(data.map(d => d.APPLICATION_NO)).size;

  // Total Bypass Applications = Distinct count where Bypass Flag = "Bypass"
  const bypassApps = new Set(data.filter(d => d["Bypass Flag"] === "Bypass").map(d => d.APPLICATION_NO)).size;

  // Total Regular Applications
  const regularApps = new Set(data.filter(d => d["Bypass Flag"] === "Regular").map(d => d.APPLICATION_NO)).size;

  // Bypass Percentage = Total Bypass / Total * 100
  const bypassPct = totalApps > 0 ? (bypassApps / totalApps) * 100 : 0;

  // Daily Statistics
  const dailyBypassMap = new Map();
  data.forEach(d => {
    const dt = d.APPROVAL_DATE;
    if (!dailyBypassMap.has(dt)) {
      dailyBypassMap.set(dt, { total: 0, bypass: 0, date: dt });
    }
    const rec = dailyBypassMap.get(dt);
    rec.total++;
    if (d["Bypass Flag"] === "Bypass") rec.bypass++;
  });

  const activeDaysCount = dailyBypassMap.size;
  const avgDailyBypass = activeDaysCount > 0 ? (bypassApps / activeDaysCount) : 0;

  let peakDate = "--";
  let peakCount = 0;
  dailyBypassMap.forEach((val, dt) => {
    if (val.bypass > peakCount) {
      peakCount = val.bypass;
      peakDate = dt;
    }
  });

  return { totalApps, bypassApps, regularApps, bypassPct, avgDailyBypass, peakDate, peakCount, activeDaysCount };
}

// Render Top KPI Cards
function renderKPIs(measures) {
  kpiTotalApps.textContent = measures.totalApps.toLocaleString();
  kpiBypassApps.textContent = measures.bypassApps.toLocaleString();
  kpiBypassPct.textContent = measures.bypassPct.toFixed(1) + "%";
  if (kpiRegularApps) kpiRegularApps.textContent = measures.regularApps.toLocaleString();

  // Daily Bypass Pace KPI
  const kpiDailyAvg = document.getElementById("kpiDailyAvg");
  const kpiDailyPeak = document.getElementById("kpiDailyPeak");
  if (kpiDailyAvg) {
    kpiDailyAvg.textContent = measures.avgDailyBypass > 0 
      ? measures.avgDailyBypass.toFixed(1) + " / day" 
      : "0 / day";
  }
  if (kpiDailyPeak) {
    kpiDailyPeak.textContent = measures.peakCount > 0 
      ? `Peak: ${measures.peakCount} bypasses on ${measures.peakDate}` 
      : "Peak: 0 recorded";
  }

  // Status Badge for Bypass %
  if (kpiBadgePct) {
    if (measures.bypassPct >= 35) {
      kpiBadgePct.className = "kpi-badge badge-danger";
      kpiBadgePct.textContent = "High Ratio (Alert)";
    } else if (measures.bypassPct >= 20) {
      kpiBadgePct.className = "kpi-badge badge-warning";
      kpiBadgePct.style.background = "#fff4ce";
      kpiBadgePct.style.color = "#8a6100";
      kpiBadgePct.textContent = "Moderate Ratio";
    } else {
      kpiBadgePct.className = "kpi-badge badge-success";
      kpiBadgePct.style.background = "#def7ec";
      kpiBadgePct.style.color = "#03543f";
      kpiBadgePct.textContent = "Normal Baseline";
    }
  }
}

// Visual 1 Mode State: "dates" or "monthly"
let v1CurrentMode = "dates";
let v1SelectedMonth = "All";

function populateV1MonthPicker() {
  const v1Picker = document.getElementById("v1MonthPicker");
  if (!v1Picker) return;

  const currentVal = v1SelectedMonth;
  v1Picker.innerHTML = `<option value="All">All Months (Continuous)</option>`;

  const uniqueMonths = [...new Map(processedData.map(d => [d["Month-Year Sort"], d["Month-Year"]])).entries()]
    .sort((a, b) => a[0] - b[0])
    .map(entry => entry[1]);

  uniqueMonths.forEach(m => {
    const opt = document.createElement("option");
    opt.value = m;
    opt.textContent = m;
    v1Picker.appendChild(opt);
  });

  if (activeFilters.monthYear !== "All") {
    v1Picker.value = activeFilters.monthYear;
    v1SelectedMonth = activeFilters.monthYear;
  } else {
    v1Picker.value = currentVal;
  }
}

function updateV1ToggleUI() {
  const btnMonthly = document.getElementById("v1BtnMonthly");
  const btnDates = document.getElementById("v1BtnDates");
  const title = document.getElementById("v1Title");
  const subtitle = document.getElementById("v1Subtitle");
  const dateScrollBar = document.getElementById("v1DateScrollBar");
  const monthPicker = document.getElementById("v1MonthPicker");

  if (v1CurrentMode === "dates") {
    if (btnDates) btnDates.classList.add("active");
    if (btnMonthly) btnMonthly.classList.remove("active");
    if (title) title.textContent = "Bypass Count by Whole Month Dates";
    if (subtitle) subtitle.textContent = "Day-by-day dates across the month";
    if (dateScrollBar) dateScrollBar.style.display = "flex";
    if (monthPicker) {
      monthPicker.style.display = "inline-block";
      populateV1MonthPicker();
    }
  } else {
    if (btnMonthly) btnMonthly.classList.add("active");
    if (btnDates) btnDates.classList.remove("active");
    if (title) title.textContent = "Bypass Count by Month";
    if (subtitle) subtitle.textContent = "Monthly volume of PIVC bypass approvals";
    if (monthPicker) monthPicker.style.display = "none";
    if (dateScrollBar) dateScrollBar.style.display = "none";
  }
}

// Visual 1: Column Chart - Bypass Count by Month OR Whole Month Dates with Scrolling Bar
function renderMonthlyBypassColumnChart(data) {
  updateV1ToggleUI();

  const scrollContainer = document.getElementById("barChartMonthlyScroll");
  const slider = document.getElementById("v1ScrollSlider");
  const dateStartLbl = document.getElementById("v1DateStartLabel");
  const dateEndLbl = document.getElementById("v1DateEndLabel");
  const scrollInfoTxt = document.getElementById("v1ScrollInfoText");

  if (v1CurrentMode === "dates") {
    // -------------------------------------------------------------
    // WHOLE MONTH DATES (Horizontal Scrollable Bar)
    // -------------------------------------------------------------
    const targetMonth = (activeFilters.monthYear !== "All") 
      ? activeFilters.monthYear 
      : (v1SelectedMonth !== "All" ? v1SelectedMonth : "All");

    let datesList = [];

    if (targetMonth !== "All") {
      // User selected a specific month: generate all days in this whole month (Day 1 to Day 28/29/30/31)
      const parts = targetMonth.split(" ");
      const mStr = parts[0];
      const year = parseInt(parts[1], 10);
      const mIdx = MONTH_NAMES.indexOf(mStr);

      if (mIdx !== -1 && !isNaN(year)) {
        const daysInMonth = new Date(year, mIdx + 1, 0).getDate();
        
        // Map data in this month
        const dayBypassMap = new Map();
        data.forEach(d => {
          if (d["Month-Year"] === targetMonth) {
            const dt = d.APPROVAL_DATE;
            if (!dayBypassMap.has(dt)) dayBypassMap.set(dt, { bypass: 0, total: 0 });
            const rec = dayBypassMap.get(dt);
            rec.total++;
            if (d["Bypass Flag"] === "Bypass") rec.bypass++;
          }
        });

        for (let day = 1; day <= daysInMonth; day++) {
          const dayStr = String(day).padStart(2, "0");
          const mNum = String(mIdx + 1).padStart(2, "0");
          const fullDate = `${year}-${mNum}-${dayStr}`;
          const dObj = new Date(year, mIdx, day);
          const dow = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"][dObj.getDay()];
          const isWeekend = (dObj.getDay() === 0 || dObj.getDay() === 6);
          const found = dayBypassMap.get(fullDate) || { bypass: 0, total: 0 };

          datesList.push({
            date: fullDate,
            label: `${dayStr} ${mStr}`,
            shortLabel: `${dayStr}`,
            dayName: dow,
            isWeekend: isWeekend,
            bypass: found.bypass,
            total: found.total,
            monthName: targetMonth
          });
        }
      }
    }

    // If "All" or fallback: aggregate all active dates chronologically
    if (datesList.length === 0) {
      const dateMap = new Map();
      data.forEach(d => {
        const dt = d.APPROVAL_DATE;
        if (!dateMap.has(dt)) {
          dateMap.set(dt, { date: dt, total: 0, bypass: 0 });
        }
        const cur = dateMap.get(dt);
        cur.total++;
        if (d["Bypass Flag"] === "Bypass") cur.bypass++;
      });

      const sortedRaw = Array.from(dateMap.values()).sort((a, b) => new Date(a.date) - new Date(b.date));
      sortedRaw.forEach(item => {
        const dObj = new Date(item.date);
        const dayStr = !isNaN(dObj.getTime()) ? String(dObj.getDate()).padStart(2, "0") : item.date;
        const mStr = !isNaN(dObj.getTime()) ? MONTH_NAMES[dObj.getMonth()] : "";
        const dow = !isNaN(dObj.getTime()) ? ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"][dObj.getDay()] : "";
        const isWeekend = !isNaN(dObj.getTime()) ? (dObj.getDay() === 0 || dObj.getDay() === 6) : false;

        datesList.push({
          date: item.date,
          label: `${dayStr} ${mStr}`,
          shortLabel: dayStr,
          dayName: dow,
          isWeekend: isWeekend,
          bypass: item.bypass,
          total: item.total,
          monthName: `${mStr} ${dObj.getFullYear()}`
        });
      });
    }

    if (datesList.length === 0) {
      barChartMonthly.innerHTML = `<div style="color:var(--text-muted);font-size:0.85rem;padding:40px;text-align:center;">No date records found for the selected filter</div>`;
      if (dateStartLbl) dateStartLbl.textContent = "--";
      if (dateEndLbl) dateEndLbl.textContent = "--";
      if (scrollInfoTxt) scrollInfoTxt.textContent = "No dates available";
      return;
    }

    // Update Slider & Date Labels
    if (dateStartLbl) dateStartLbl.textContent = datesList[0].label;
    if (dateEndLbl) dateEndLbl.textContent = datesList[datesList.length - 1].label;
    if (scrollInfoTxt) {
      const titlePrefix = targetMonth !== "All" ? `${targetMonth}: ` : "";
      scrollInfoTxt.textContent = `${titlePrefix}Showing ${datesList.length} whole month calendar dates (Scroll horizontally to view all)`;
    }

    const maxVal = Math.max(...datesList.map(d => d.bypass), 1);
    const colCount = datesList.length;
    const colWidth = 24;
    const minStep = 48; // generous step width so scrolling is smooth
    const svgWidth = Math.max(720, colCount * minStep + 100);
    const svgHeight = 250;
    const padLeft = 45;
    const padRight = 35;
    const padTop = 26;
    const padBottom = 52;
    const chartW = svgWidth - padLeft - padRight;
    const chartH = svgHeight - padTop - padBottom;
    const step = chartW / colCount;

    let svg = `<svg viewBox="0 0 ${svgWidth} ${svgHeight}" style="width:${svgWidth}px;height:100%;min-width:${svgWidth}px;overflow:visible;">`;

    // Gradients & Defs
    svg += `
      <defs>
        <linearGradient id="v1BypassGradient" x1="0" y1="0" x2="0" y2="1">
          <stop offset="0%" stop-color="#ff6b3b"/>
          <stop offset="100%" stop-color="#d83b01"/>
        </linearGradient>
      </defs>
    `;

    // Background horizontal grid lines
    for (let i = 0; i <= 4; i++) {
      const yVal = Math.round((maxVal / 4) * i);
      const yPos = padTop + chartH - (chartH * (i / 4));
      svg += `
        <line x1="${padLeft}" y1="${yPos}" x2="${svgWidth - padRight}" y2="${yPos}" stroke="#edf2f7" stroke-dasharray="3,3" />
        <text x="${padLeft - 8}" y="${yPos + 4}" font-size="10" fill="#94a3b8" text-anchor="end">${yVal}</text>
      `;
    }

    // Baseline axis line
    svg += `<line x1="${padLeft}" y1="${padTop + chartH}" x2="${svgWidth - padRight}" y2="${padTop + chartH}" stroke="#cbd5e1" stroke-width="1.5" />`;

    // Render each date column
    datesList.forEach((d, idx) => {
      const x = padLeft + idx * step + (step - colWidth) / 2;
      const slotX = padLeft + idx * step;

      // Weekend subtle band
      if (d.isWeekend) {
        svg += `<rect x="${slotX}" y="${padTop}" width="${step}" height="${chartH}" fill="rgba(241, 245, 249, 0.55)" />`;
      }

      if (d.bypass > 0) {
        const barH = Math.max(5, (d.bypass / maxVal) * chartH);
        const y = padTop + chartH - barH;
        const pctStr = d.total > 0 ? ((d.bypass / d.total) * 100).toFixed(1) + "%" : "100%";

        svg += `
          <g class="chart-col-group" style="cursor:pointer;">
            <rect x="${x}" y="${y}" width="${colWidth}" height="${barH}" rx="4" fill="url(#v1BypassGradient)">
              <title>${d.date} (${d.dayName}): ${d.bypass} Bypass Cases (${d.total} Total Apps, ${pctStr} Bypass Rate)</title>
            </rect>
            <text x="${x + colWidth / 2}" y="${y - 6}" font-size="10" font-weight="700" fill="#d83b01" text-anchor="middle">${d.bypass}</text>
            <text x="${x + colWidth / 2}" y="${svgHeight - 24}" font-size="9" font-weight="600" fill="${d.isWeekend ? '#0284c7' : '#334155'}" text-anchor="middle">${d.shortLabel}</text>
            <text x="${x + colWidth / 2}" y="${svgHeight - 12}" font-size="8" fill="#94a3b8" text-anchor="middle">${d.dayName}</text>
          </g>
        `;
      } else {
        // Zero bypass baseline dot / tick
        svg += `
          <g class="chart-col-group" style="cursor:pointer;">
            <rect x="${x + colWidth / 2 - 2}" y="${padTop + chartH - 3}" width="4" height="3" rx="1" fill="#cbd5e1">
              <title>${d.date} (${d.dayName}): 0 Bypass Cases (${d.total} Total Apps)</title>
            </rect>
            <text x="${x + colWidth / 2}" y="${svgHeight - 24}" font-size="9" fill="#94a3b8" text-anchor="middle">${d.shortLabel}</text>
            <text x="${x + colWidth / 2}" y="${svgHeight - 12}" font-size="8" fill="#cbd5e1" text-anchor="middle">${d.dayName}</text>
          </g>
        `;
      }
    });

    svg += `</svg>`;
    barChartMonthly.innerHTML = svg;

    // Reset or update slider position
    if (slider && scrollContainer) {
      setTimeout(() => {
        const maxScroll = scrollContainer.scrollWidth - scrollContainer.clientWidth;
        if (maxScroll > 0) {
          slider.value = (scrollContainer.scrollLeft / maxScroll) * 100;
        } else {
          slider.value = 0;
        }
      }, 50);
    }

  } else {
    // -------------------------------------------------------------
    // BY MONTH (Aggregated monthly column chart)
    // -------------------------------------------------------------
    const monthsMap = new Map();
    data.forEach(d => {
      const key = d["Month-Year Sort"];
      if (!monthsMap.has(key)) {
        monthsMap.set(key, { name: d["Month-Year"], total: 0, bypass: 0, sortKey: key });
      }
      const cur = monthsMap.get(key);
      cur.total++;
      if (d["Bypass Flag"] === "Bypass") {
        cur.bypass++;
      }
    });

    const sortedMonths = Array.from(monthsMap.values()).sort((a, b) => a.sortKey - b.sortKey);
    const maxVal = Math.max(...sortedMonths.map(m => m.bypass), 1);

    if (sortedMonths.length === 0) {
      barChartMonthly.innerHTML = `<div style="color:var(--text-muted);font-size:0.85rem;padding:40px;text-align:center;">No data matching current filters</div>`;
      return;
    }

    const colCount = sortedMonths.length;
    const minStep = 60;
    const svgWidth = Math.max(560, colCount * minStep + 60);
    const svgHeight = 250;
    const padLeft = 40;
    const padRight = 20;
    const padTop = 25;
    const padBottom = 35;
    const chartW = svgWidth - padLeft - padRight;
    const chartH = svgHeight - padTop - padBottom;
    const colWidth = Math.min(42, (chartW / colCount) * 0.55);
    const step = chartW / colCount;

    let svg = `<svg viewBox="0 0 ${svgWidth} ${svgHeight}" style="width:${svgWidth}px;height:100%;min-width:100%;overflow:visible;">`;

    // Grid lines
    for (let i = 0; i <= 4; i++) {
      const yVal = Math.round((maxVal / 4) * i);
      const yPos = padTop + chartH - (chartH * (i / 4));
      svg += `
        <line x1="${padLeft}" y1="${yPos}" x2="${svgWidth - padRight}" y2="${yPos}" stroke="#edf2f7" stroke-dasharray="3,3" />
        <text x="${padLeft - 8}" y="${yPos + 4}" font-size="10" fill="#94a3b8" text-anchor="end">${yVal}</text>
      `;
    }

    // Columns
    sortedMonths.forEach((m, idx) => {
      const x = padLeft + idx * step + (step - colWidth) / 2;
      const barH = (m.bypass / maxVal) * chartH;
      const y = padTop + chartH - barH;

      svg += `
        <g class="chart-col-group" style="cursor:pointer;">
          <rect x="${x}" y="${y}" width="${colWidth}" height="${barH}" rx="4" fill="#d83b01">
            <title>${m.name}: ${m.bypass} Bypass Applications (${m.total} Total, ${((m.bypass/m.total)*100).toFixed(1)}%)</title>
          </rect>
          <text x="${x + colWidth / 2}" y="${y - 6}" font-size="11" font-weight="700" fill="#d83b01" text-anchor="middle">${m.bypass}</text>
          <text x="${x + colWidth / 2}" y="${svgHeight - 10}" font-size="10" fill="#64748b" text-anchor="middle">${m.name}</text>
        </g>
      `;
    });

    svg += `</svg>`;
    barChartMonthly.innerHTML = svg;
  }
}

// Visual 2: Line & Area Chart - Monthly Bypass Trend with Multiple Modes (Rate %, Dual Volume, Cumulative)
function renderMonthlyTrendLineChart(data) {
  const lineChartTrend = document.getElementById("lineChartTrend");
  const lineChartTitle = document.getElementById("lineChartTitle");
  const lineChartSubtitle = document.getElementById("lineChartSubtitle");
  const lineVisualTag = document.getElementById("lineVisualTag");
  if (!lineChartTrend) return;

  const monthsMap = new Map();
  data.forEach(d => {
    const key = d["Month-Year Sort"];
    if (!monthsMap.has(key)) {
      monthsMap.set(key, { name: d["Month-Year"], total: 0, bypass: 0, regular: 0, sortKey: key });
    }
    const cur = monthsMap.get(key);
    cur.total++;
    if (d["Bypass Flag"] === "Bypass") {
      cur.bypass++;
    } else {
      cur.regular++;
    }
  });

  const sortedMonths = Array.from(monthsMap.values()).sort((a, b) => a.sortKey - b.sortKey);
  if (sortedMonths.length === 0) {
    lineChartTrend.innerHTML = `<div style="color:var(--text-muted);font-size:0.85rem;padding:40px;text-align:center;">No data matching current filters</div>`;
    return;
  }

  const svgWidth = 560;
  const svgHeight = 250;
  const padLeft = 48;
  const padRight = 20;
  const padTop = 30;
  const padBottom = 35;
  const chartW = svgWidth - padLeft - padRight;
  const chartH = svgHeight - padTop - padBottom;
  const count = sortedMonths.length;

  let svg = `<svg viewBox="0 0 ${svgWidth} ${svgHeight}" style="width:100%;height:100%;overflow:visible;">`;

  if (currentLineChartMode === "pct") {
    // Mode 1: Bypass Percentage Rate (%)
    if (lineChartTitle) lineChartTitle.textContent = "Monthly Bypass Trend (Rate %)";
    if (lineChartSubtitle) lineChartSubtitle.textContent = "Tracking the monthly bypass percentage (%) trajectory over time";
    if (lineVisualTag) lineVisualTag.textContent = "Rate Line";

    const pctValues = sortedMonths.map(m => m.total > 0 ? (m.bypass / m.total) * 100 : 0);
    const maxPct = Math.max(...pctValues, 10);
    const chartMax = Math.ceil(maxPct / 10) * 10;

    // Y Grid lines
    for (let i = 0; i <= 4; i++) {
      const yPct = Math.round((chartMax / 4) * i);
      const yPos = padTop + chartH - (chartH * (i / 4));
      svg += `
        <line x1="${padLeft}" y1="${yPos}" x2="${svgWidth - padRight}" y2="${yPos}" stroke="#edf2f7" stroke-dasharray="3,3" />
        <text x="${padLeft - 8}" y="${yPos + 4}" font-size="10" fill="#94a3b8" text-anchor="end">${yPct}%</text>
      `;
    }

    const points = sortedMonths.map((m, idx) => {
      const x = count === 1 ? padLeft + chartW / 2 : padLeft + (idx / (count - 1)) * chartW;
      const pct = m.total > 0 ? (m.bypass / m.total) * 100 : 0;
      const y = padTop + chartH - (pct / chartMax) * chartH;
      return { x, y, pct, name: m.name, bypass: m.bypass, total: m.total };
    });

    let areaD = `M ${points[0].x} ${padTop + chartH}`;
    points.forEach(p => { areaD += ` L ${p.x} ${p.y}`; });
    areaD += ` L ${points[points.length - 1].x} ${padTop + chartH} Z`;

    let lineD = `M ${points[0].x} ${points[0].y}`;
    for (let i = 1; i < points.length; i++) {
      lineD += ` L ${points[i].x} ${points[i].y}`;
    }

    svg += `
      <defs>
        <linearGradient id="pctGrad" x1="0" y1="0" x2="0" y2="1">
          <stop offset="0%" stop-color="#0078d4" stop-opacity="0.25"/>
          <stop offset="100%" stop-color="#0078d4" stop-opacity="0.0"/>
        </linearGradient>
      </defs>
      <path d="${areaD}" fill="url(#pctGrad)"/>
      <path d="${lineD}" fill="none" stroke="#0078d4" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/>
    `;

    points.forEach(p => {
      svg += `
        <circle cx="${p.x}" cy="${p.y}" r="5" fill="#ffffff" stroke="#0078d4" stroke-width="2.5" class="line-chart-node">
          <title>${p.name}: ${p.pct.toFixed(1)}% Bypass Rate (${p.bypass} bypass / ${p.total} total)</title>
        </circle>
        <text x="${p.x}" y="${p.y - 10}" font-size="10" font-weight="700" fill="#0078d4" text-anchor="middle">${p.pct.toFixed(0)}%</text>
        <text x="${p.x}" y="${svgHeight - 10}" font-size="10" fill="#64748b" text-anchor="middle">${p.name}</text>
      `;
    });

  } else if (currentLineChartMode === "dual") {
    // Mode 2: Dual Volume Lines (Bypass vs Regular counts)
    if (lineChartTitle) lineChartTitle.textContent = "Monthly Volume Trend (Dual Lines)";
    if (lineChartSubtitle) lineChartSubtitle.textContent = "Comparing monthly bypass volume against automated regular completions";
    if (lineVisualTag) lineVisualTag.textContent = "Dual Lines";

    const maxVal = Math.max(...sortedMonths.map(m => Math.max(m.regular, m.bypass)), 10);
    const chartMax = Math.ceil(maxVal / 10) * 10;

    // Header Legend
    svg += `
      <g transform="translate(${padLeft + 10}, 16)" font-size="10" font-weight="600">
        <line x1="0" y1="0" x2="16" y2="0" stroke="#107c41" stroke-width="3"/>
        <circle cx="8" cy="0" r="3" fill="#107c41"/>
        <text x="22" y="3" fill="#107c41">Regular Volume</text>
        <line x1="120" y1="0" x2="136" y2="0" stroke="#d83b01" stroke-width="3"/>
        <circle cx="128" cy="0" r="3" fill="#d83b01"/>
        <text x="142" y="3" fill="#d83b01">Bypass Volume</text>
      </g>
    `;

    for (let i = 0; i <= 4; i++) {
      const yVal = Math.round((chartMax / 4) * i);
      const yPos = padTop + chartH - (chartH * (i / 4));
      svg += `
        <line x1="${padLeft}" y1="${yPos}" x2="${svgWidth - padRight}" y2="${yPos}" stroke="#edf2f7" stroke-dasharray="3,3" />
        <text x="${padLeft - 8}" y="${yPos + 4}" font-size="10" fill="#94a3b8" text-anchor="end">${yVal}</text>
      `;
    }

    const regPoints = sortedMonths.map((m, idx) => {
      const x = count === 1 ? padLeft + chartW / 2 : padLeft + (idx / (count - 1)) * chartW;
      const y = padTop + chartH - (m.regular / chartMax) * chartH;
      return { x, y, val: m.regular, name: m.name };
    });

    const bypassPoints = sortedMonths.map((m, idx) => {
      const x = count === 1 ? padLeft + chartW / 2 : padLeft + (idx / (count - 1)) * chartW;
      const y = padTop + chartH - (m.bypass / chartMax) * chartH;
      return { x, y, val: m.bypass, name: m.name };
    });

    let regLineD = `M ${regPoints[0].x} ${regPoints[0].y}`;
    for (let i = 1; i < regPoints.length; i++) regLineD += ` L ${regPoints[i].x} ${regPoints[i].y}`;

    let bypassLineD = `M ${bypassPoints[0].x} ${bypassPoints[0].y}`;
    for (let i = 1; i < bypassPoints.length; i++) bypassLineD += ` L ${bypassPoints[i].x} ${bypassPoints[i].y}`;

    svg += `
      <path d="${regLineD}" fill="none" stroke="#107c41" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/>
      <path d="${bypassLineD}" fill="none" stroke="#d83b01" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/>
    `;

    regPoints.forEach(p => {
      svg += `
        <circle cx="${p.x}" cy="${p.y}" r="4" fill="#ffffff" stroke="#107c41" stroke-width="2" class="line-chart-node">
          <title>${p.name}: ${p.val} Regular Applications</title>
        </circle>
        <text x="${p.x}" y="${p.y - 7}" font-size="9" font-weight="700" fill="#107c41" text-anchor="middle">${p.val}</text>
      `;
    });

    bypassPoints.forEach(p => {
      svg += `
        <circle cx="${p.x}" cy="${p.y}" r="4" fill="#ffffff" stroke="#d83b01" stroke-width="2" class="line-chart-node">
          <title>${p.name}: ${p.val} Bypass Applications</title>
        </circle>
        <text x="${p.x}" y="${p.y + 14}" font-size="9" font-weight="700" fill="#d83b01" text-anchor="middle">${p.val}</text>
        <text x="${p.x}" y="${svgHeight - 10}" font-size="10" fill="#64748b" text-anchor="middle">${p.name}</text>
      `;
    });

  } else if (currentLineChartMode === "cumulative") {
    // Mode 3: Cumulative Bypass Growth Curve
    if (lineChartTitle) lineChartTitle.textContent = "Cumulative Bypass Growth";
    if (lineChartSubtitle) lineChartSubtitle.textContent = "Cumulative sum of exceptions tracking policy threshold";
    if (lineVisualTag) lineVisualTag.textContent = "Cumulative Curve";

    let runningSum = 0;
    const cumPointsData = sortedMonths.map((m, idx) => {
      runningSum += m.bypass;
      return { name: m.name, val: runningSum, monthBypass: m.bypass };
    });

    const maxVal = runningSum > 0 ? runningSum : 10;
    const chartMax = Math.ceil(maxVal / 10) * 10;

    for (let i = 0; i <= 4; i++) {
      const yVal = Math.round((chartMax / 4) * i);
      const yPos = padTop + chartH - (chartH * (i / 4));
      svg += `
        <line x1="${padLeft}" y1="${yPos}" x2="${svgWidth - padRight}" y2="${yPos}" stroke="#edf2f7" stroke-dasharray="3,3" />
        <text x="${padLeft - 8}" y="${yPos + 4}" font-size="10" fill="#94a3b8" text-anchor="end">${yVal}</text>
      `;
    }

    const points = cumPointsData.map((m, idx) => {
      const x = count === 1 ? padLeft + chartW / 2 : padLeft + (idx / (count - 1)) * chartW;
      const y = padTop + chartH - (m.val / chartMax) * chartH;
      return { x, y, val: m.val, name: m.name, monthBypass: m.monthBypass };
    });

    let areaD = `M ${points[0].x} ${padTop + chartH}`;
    points.forEach(p => { areaD += ` L ${p.x} ${p.y}`; });
    areaD += ` L ${points[points.length - 1].x} ${padTop + chartH} Z`;

    let lineD = `M ${points[0].x} ${points[0].y}`;
    for (let i = 1; i < points.length; i++) lineD += ` L ${points[i].x} ${points[i].y}`;

    svg += `
      <defs>
        <linearGradient id="cumGrad" x1="0" y1="0" x2="0" y2="1">
          <stop offset="0%" stop-color="#d83b01" stop-opacity="0.30"/>
          <stop offset="100%" stop-color="#d83b01" stop-opacity="0.02"/>
        </linearGradient>
      </defs>
      <path d="${areaD}" fill="url(#cumGrad)"/>
      <path d="${lineD}" fill="none" stroke="#d83b01" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/>
    `;

    points.forEach(p => {
      svg += `
        <circle cx="${p.x}" cy="${p.y}" r="5" fill="#ffffff" stroke="#d83b01" stroke-width="2.5" class="line-chart-node">
          <title>${p.name}: ${p.val} Cumulative Bypasses (+${p.monthBypass} in ${p.name})</title>
        </circle>
        <text x="${p.x}" y="${p.y - 10}" font-size="10" font-weight="700" fill="#d83b01" text-anchor="middle">${p.val}</text>
        <text x="${p.x}" y="${svgHeight - 10}" font-size="10" fill="#64748b" text-anchor="middle">${p.name}</text>
      `;
    });
  }

  svg += `</svg>`;
  lineChartTrend.innerHTML = svg;
}

// NEW VISUAL: Stacked & Clustered Monthly Composition Bar Chart
function renderStackedMonthlyBarChart(data) {
  const container = document.getElementById("stackedBarChart");
  const titleEl = document.getElementById("stackedBarTitle");
  const subtitleEl = document.getElementById("stackedBarSubtitle");
  if (!container) return;

  const dataset = (data && data.length > 0) ? data : processedData;
  const monthsMap = new Map();
  dataset.forEach(d => {
    const key = d["Month-Year Sort"] || 0;
    const name = d["Month-Year"] || "Unknown";
    if (!monthsMap.has(key)) {
      monthsMap.set(key, { name: name, total: 0, bypass: 0, regular: 0, sortKey: key });
    }
    const cur = monthsMap.get(key);
    cur.total++;
    if (d["Bypass Flag"] === "Bypass") {
      cur.bypass++;
    } else {
      cur.regular++;
    }
  });

  const sortedMonths = Array.from(monthsMap.values()).sort((a, b) => a.sortKey - b.sortKey);
  if (sortedMonths.length === 0) {
    container.innerHTML = `<div style="color:var(--text-muted);font-size:0.85rem;padding:40px;text-align:center;">No data matching current filters</div>`;
    return;
  }

  const svgWidth = 1000;
  const svgHeight = 240;
  const padLeft = 50;
  const padRight = 30;
  const padTop = 30;
  const padBottom = 35;
  const chartW = svgWidth - padLeft - padRight;
  const chartH = svgHeight - padTop - padBottom;
  const count = sortedMonths.length;

  let svg = `<svg viewBox="0 0 ${svgWidth} ${svgHeight}" width="100%" height="${svgHeight}" style="width:100%;height:${svgHeight}px;min-height:${svgHeight}px;display:block;overflow:visible;">`;

  if (currentStackedBarMode === "100") {
    if (titleEl) titleEl.textContent = "Monthly Volume Composition (100% Normalized Proportions)";
    if (subtitleEl) subtitleEl.textContent = "Normalized 100% stacked bar chart highlighting the monthly bypass ratio";

    for (let i = 0; i <= 4; i++) {
      const yVal = 25 * i;
      const yPos = padTop + chartH - (chartH * (i / 4));
      svg += `
        <line x1="${padLeft}" y1="${yPos}" x2="${svgWidth - padRight}" y2="${yPos}" stroke="#edf2f7" stroke-dasharray="3,3" />
        <text x="${padLeft - 8}" y="${yPos + 4}" font-size="10" fill="#94a3b8" text-anchor="end">${yVal}%</text>
      `;
    }

    const colWidth = Math.min(Math.max(chartW / (count * 1.8), 24), 50);
    const step = chartW / count;

    sortedMonths.forEach((m, idx) => {
      const x = padLeft + idx * step + (step - colWidth) / 2;
      const regPct = m.total > 0 ? (m.regular / m.total) * 100 : 0;
      const bypPct = m.total > 0 ? (m.bypass / m.total) * 100 : 0;

      const regH = (regPct / 100) * chartH;
      const bypH = (bypPct / 100) * chartH;
      const regY = padTop + chartH - regH;
      const bypY = regY - bypH;

      svg += `
        <g class="stacked-bar-col">
          <!-- Regular bottom rect -->
          <rect x="${x}" y="${regY}" width="${colWidth}" height="${regH}" fill="#107c41" class="stacked-bar-seg">
            <title>${m.name}: ${regPct.toFixed(1)}% Regular (${m.regular} apps)</title>
          </rect>
          <!-- Bypass top rect -->
          <rect x="${x}" y="${bypY}" width="${colWidth}" height="${bypH}" fill="#d83b01" rx="3" class="stacked-bar-seg">
            <title>${m.name}: ${bypPct.toFixed(1)}% Bypass (${m.bypass} apps)</title>
          </rect>
          <!-- Inner labels -->
          <text x="${x + colWidth / 2}" y="${regY + regH/2 + 4}" font-size="10" font-weight="700" fill="#ffffff" text-anchor="middle">${regPct.toFixed(0)}%</text>
          ${bypH > 14 ? `<text x="${x + colWidth / 2}" y="${bypY + bypH/2 + 4}" font-size="9" font-weight="700" fill="#ffffff" text-anchor="middle">${bypPct.toFixed(0)}%</text>` : ""}
          <!-- X Axis Month Label -->
          <text x="${x + colWidth / 2}" y="${svgHeight - 10}" font-size="11" fill="#64748b" text-anchor="middle">${m.name}</text>
        </g>
      `;
    });

  } else if (currentStackedBarMode === "clustered") {
    if (titleEl) titleEl.textContent = "Monthly Volume Composition (Side-by-Side Clustered Bars)";
    if (subtitleEl) subtitleEl.textContent = "Clustered bar chart directly comparing regular vs bypass application counts";

    const maxVal = Math.max(...sortedMonths.map(m => Math.max(m.regular, m.bypass)), 10);
    const chartMax = Math.ceil(maxVal / 10) * 10;

    for (let i = 0; i <= 4; i++) {
      const yVal = Math.round((chartMax / 4) * i);
      const yPos = padTop + chartH - (chartH * (i / 4));
      svg += `
        <line x1="${padLeft}" y1="${yPos}" x2="${svgWidth - padRight}" y2="${yPos}" stroke="#edf2f7" stroke-dasharray="3,3" />
        <text x="${padLeft - 8}" y="${yPos + 4}" font-size="10" fill="#94a3b8" text-anchor="end">${yVal}</text>
      `;
    }

    const step = chartW / count;
    const singleBarW = Math.min(Math.max(step * 0.36, 12), 26);

    sortedMonths.forEach((m, idx) => {
      const groupCenter = padLeft + idx * step + step / 2;
      const xReg = groupCenter - singleBarW - 2;
      const xByp = groupCenter + 2;

      const regH = (m.regular / chartMax) * chartH;
      const bypH = (m.bypass / chartMax) * chartH;
      const regY = padTop + chartH - regH;
      const bypY = padTop + chartH - bypH;

      svg += `
        <g class="clustered-bar-group">
          <!-- Regular Bar (Green) -->
          <rect x="${xReg}" y="${regY}" width="${singleBarW}" height="${regH}" fill="#107c41" rx="3" class="stacked-bar-seg">
            <title>${m.name}: ${m.regular} Regular Applications</title>
          </rect>
          <text x="${xReg + singleBarW / 2}" y="${regY - 5}" font-size="9" font-weight="700" fill="#107c41" text-anchor="middle">${m.regular}</text>

          <!-- Bypass Bar (Coral) -->
          <rect x="${xByp}" y="${bypY}" width="${singleBarW}" height="${bypH}" fill="#d83b01" rx="3" class="stacked-bar-seg">
            <title>${m.name}: ${m.bypass} Bypass Applications</title>
          </rect>
          <text x="${xByp + singleBarW / 2}" y="${bypY - 5}" font-size="9" font-weight="700" fill="#d83b01" text-anchor="middle">${m.bypass}</text>

          <!-- X Axis Month Label -->
          <text x="${groupCenter}" y="${svgHeight - 10}" font-size="11" fill="#64748b" text-anchor="middle">${m.name}</text>
        </g>
      `;
    });

  } else {
    // Default mode: "stacked"
    if (titleEl) titleEl.textContent = "Monthly Volume Composition (Stacked Volume Bars)";
    if (subtitleEl) subtitleEl.textContent = "Absolute application volume per month partitioned by Regular and Bypass";

    const maxTotal = Math.max(...sortedMonths.map(m => m.total), 10);
    const chartMax = Math.ceil(maxTotal / 10) * 10;

    for (let i = 0; i <= 4; i++) {
      const yVal = Math.round((chartMax / 4) * i);
      const yPos = padTop + chartH - (chartH * (i / 4));
      svg += `
        <line x1="${padLeft}" y1="${yPos}" x2="${svgWidth - padRight}" y2="${yPos}" stroke="#edf2f7" stroke-dasharray="3,3" />
        <text x="${padLeft - 8}" y="${yPos + 4}" font-size="10" fill="#94a3b8" text-anchor="end">${yVal}</text>
      `;
    }

    const colWidth = Math.min(Math.max(chartW / (count * 1.8), 24), 50);
    const step = chartW / count;

    sortedMonths.forEach((m, idx) => {
      const x = padLeft + idx * step + (step - colWidth) / 2;
      const regH = (m.regular / chartMax) * chartH;
      const bypH = (m.bypass / chartMax) * chartH;
      const regY = padTop + chartH - regH;
      const bypY = regY - bypH;

      svg += `
        <g class="stacked-bar-col">
          <!-- Regular bottom rect -->
          <rect x="${x}" y="${regY}" width="${colWidth}" height="${regH}" fill="#107c41" rx="0" class="stacked-bar-seg">
            <title>${m.name}: ${m.regular} Regular Applications (${((m.regular/m.total)*100).toFixed(1)}%)</title>
          </rect>
          <!-- Bypass top rect -->
          <rect x="${x}" y="${bypY}" width="${colWidth}" height="${bypH}" fill="#d83b01" rx="3" class="stacked-bar-seg">
            <title>${m.name}: ${m.bypass} Bypass Applications (${((m.bypass/m.total)*100).toFixed(1)}%)</title>
          </rect>
          <!-- Total label above bar -->
          <text x="${x + colWidth / 2}" y="${bypY - 6}" font-size="10" font-weight="700" fill="#1e293b" text-anchor="middle">${m.total}</text>
          <!-- Inner numbers if enough height -->
          ${regH > 16 ? `<text x="${x + colWidth / 2}" y="${regY + regH/2 + 4}" font-size="9" font-weight="700" fill="#ffffff" text-anchor="middle">${m.regular}</text>` : ""}
          ${bypH > 14 ? `<text x="${x + colWidth / 2}" y="${bypY + bypH/2 + 4}" font-size="9" font-weight="700" fill="#ffffff" text-anchor="middle">${m.bypass}</text>` : ""}
          <!-- X Axis Month Label -->
          <text x="${x + colWidth / 2}" y="${svgHeight - 10}" font-size="11" fill="#64748b" text-anchor="middle">${m.name}</text>
        </g>
      `;
    });
  }

  svg += `</svg>`;
  container.innerHTML = svg;
}


// NEW VISUAL: Daily Bypass Count (How many bypasses in each day)
function renderDailyBypassChart(data) {
  const container = document.getElementById("dailyChartWrapper");
  const viewModeEl = document.getElementById("dailyViewMode");
  const subtitleEl = document.getElementById("dailyChartSubtitle");
  if (!container) return;

  const mode = viewModeEl ? viewModeEl.value : "timeline";

  if (mode === "timeline") {
    // 1. Group by exact calendar date
    const dayMap = new Map();
    data.forEach(d => {
      const dt = d.APPROVAL_DATE;
      if (!dayMap.has(dt)) {
        dayMap.set(dt, { date: dt, total: 0, bypass: 0, dateObj: d.DateObj });
      }
      const rec = dayMap.get(dt);
      rec.total++;
      if (d["Bypass Flag"] === "Bypass") rec.bypass++;
    });

    const sortedDays = Array.from(dayMap.values()).sort((a, b) => new Date(a.date) - new Date(b.date));

    if (sortedDays.length === 0) {
      container.innerHTML = `<div style="color:var(--text-muted);font-size:0.85rem;padding:40px;text-align:center;">No records in selected filter range</div>`;
      return;
    }

    const totalBypassesInFilter = data.filter(d => d["Bypass Flag"] === "Bypass").length;
    const avgPerDay = (totalBypassesInFilter / sortedDays.length).toFixed(1);

    if (subtitleEl) {
      subtitleEl.textContent = `Tracking ${sortedDays.length} active dates | Average ${avgPerDay} bypasses / day`;
    }

    const maxBypass = Math.max(...sortedDays.map(d => d.bypass), 1);
    const dayCount = sortedDays.length;

    // Sizing: dynamic width so it scrolls cleanly if there are many days
    const colWidth = Math.max(16, Math.min(38, 920 / dayCount));
    const svgWidth = Math.max(920, dayCount * (colWidth + 12) + 80);
    const svgHeight = 240;
    const padLeft = 45;
    const padRight = 25;
    const padTop = 25;
    const padBottom = 45;
    const chartW = svgWidth - padLeft - padRight;
    const chartH = svgHeight - padTop - padBottom;
    const step = chartW / dayCount;

    let svg = `<svg viewBox="0 0 ${svgWidth} ${svgHeight}" style="width:${svgWidth}px;height:100%;min-width:100%;overflow:visible;">`;

    // Horizontal grid lines
    for (let i = 0; i <= 4; i++) {
      const yVal = Math.round((maxBypass / 4) * i);
      const yPos = padTop + chartH - (chartH * (i / 4));
      svg += `
        <line x1="${padLeft}" y1="${yPos}" x2="${svgWidth - padRight}" y2="${yPos}" stroke="#edf2f7" stroke-dasharray="3,3" />
        <text x="${padLeft - 8}" y="${yPos + 4}" font-size="10" fill="#94a3b8" text-anchor="end">${yVal}</text>
      `;
    }

    // Daily Bars
    sortedDays.forEach((d, idx) => {
      const x = padLeft + idx * step + (step - colWidth) / 2;
      const barH = maxBypass > 0 ? (d.bypass / maxBypass) * chartH : 0;
      const y = padTop + chartH - barH;
      const isPeak = d.bypass === maxBypass && maxBypass > 0;
      const fillColor = isPeak ? "#b91c1c" : d.bypass > 0 ? "#d83b01" : "#e2e8f0";

      // Formatted short date (e.g. "08 Jan")
      const dObj = new Date(d.date);
      const shortDate = !isNaN(dObj.getTime()) ? `${dObj.getDate()} ${MONTH_NAMES[dObj.getMonth()]}` : d.date;

      svg += `
        <g class="chart-col-group" style="cursor:pointer;">
          <rect x="${x}" y="${y}" width="${colWidth}" height="${barH}" rx="3" fill="${fillColor}">
            <title>${d.date}: ${d.bypass} Bypass Applications (${d.total} Total, ${d.total > 0 ? ((d.bypass/d.total)*100).toFixed(1) : 0}% Bypass Rate)</title>
          </rect>
          ${d.bypass > 0 ? `<text x="${x + colWidth / 2}" y="${y - 5}" font-size="10" font-weight="700" fill="${fillColor}" text-anchor="middle">${d.bypass}</text>` : ''}
          <text x="${x + colWidth / 2}" y="${svgHeight - 12}" font-size="9" fill="#64748b" text-anchor="middle" transform="rotate(-30 ${x + colWidth / 2} ${svgHeight - 12})">${shortDate}</text>
        </g>
      `;
    });

    svg += `</svg>`;
    container.innerHTML = `<div style="width:100%;overflow-x:auto;padding-bottom:10px;">${svg}</div>`;

  } else if (mode === "dow") {
    // 2. Day of Week Pattern (Monday to Sunday)
    const dowOrder = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];
    const dowMap = {};
    dowOrder.forEach(day => { dowMap[day] = { day, total: 0, bypass: 0 }; });

    data.forEach(d => {
      if (d.DateObj && !isNaN(d.DateObj.getTime())) {
        const dayName = d.DateObj.toLocaleDateString("en-US", { weekday: "long" });
        if (dowMap[dayName]) {
          dowMap[dayName].total++;
          if (d["Bypass Flag"] === "Bypass") dowMap[dayName].bypass++;
        }
      }
    });

    const dowList = dowOrder.map(day => dowMap[day]);
    const maxVal = Math.max(...dowList.map(d => d.bypass), 1);

    if (subtitleEl) {
      subtitleEl.textContent = "Bypass volume distributed by day of the week";
    }

    container.innerHTML = `
      <div class="bar-list" style="max-width:850px;margin:0 auto;width:100%;">
        ${dowList.map(item => {
          const pct = (item.bypass / maxVal) * 100;
          const rate = item.total > 0 ? ((item.bypass / item.total) * 100).toFixed(1) : 0;
          return `
            <div class="bar-row">
              <div class="bar-labels">
                <span class="bar-name">
                  <strong>${item.day}</strong>
                  <span style="font-size:0.75rem;color:#64748b;">(${item.total} total applications)</span>
                </span>
                <span class="bar-val" style="color:var(--bypass-color);">${item.bypass} bypasses <span style="font-weight:normal;color:#64748b;font-size:0.75rem;">(${rate}%)</span></span>
              </div>
              <div class="bar-track">
                <div class="bar-fill bar-fill-bypass" style="width: ${pct}%;"></div>
              </div>
            </div>
          `;
        }).join("")}
      </div>
    `;

  } else if (mode === "top_days") {
    // 3. Top 15 Peak Bypass Days
    const dayMap = new Map();
    data.forEach(d => {
      const dt = d.APPROVAL_DATE;
      if (!dayMap.has(dt)) {
        dayMap.set(dt, { date: dt, total: 0, bypass: 0, dayName: d.DateObj ? d.DateObj.toLocaleDateString("en-US", { weekday: "short" }) : "" });
      }
      const rec = dayMap.get(dt);
      rec.total++;
      if (d["Bypass Flag"] === "Bypass") rec.bypass++;
    });

    const topDays = Array.from(dayMap.values())
      .filter(d => d.bypass > 0)
      .sort((a, b) => b.bypass - a.bypass)
      .slice(0, 15);

    if (topDays.length === 0) {
      container.innerHTML = `<div style="color:var(--text-muted);font-size:0.85rem;padding:40px;text-align:center;">No bypass days recorded</div>`;
      return;
    }

    const maxVal = topDays[0].bypass;
    if (subtitleEl) {
      subtitleEl.textContent = `Top ${topDays.length} single days with the highest bypass spikes`;
    }

    container.innerHTML = `
      <div class="bar-list" style="max-width:850px;margin:0 auto;width:100%;">
        ${topDays.map((item, idx) => {
          const pct = (item.bypass / maxVal) * 100;
          const rate = item.total > 0 ? ((item.bypass / item.total) * 100).toFixed(1) : 0;
          return `
            <div class="bar-row">
              <div class="bar-labels">
                <span class="bar-name">
                  <span style="display:inline-block;width:22px;height:22px;border-radius:50%;background:#fee2e2;color:#b91c1c;font-size:0.75rem;font-weight:700;text-align:center;line-height:22px;margin-right:6px;">${idx + 1}</span>
                  <strong>${item.date}</strong> <span style="font-size:0.75rem;color:#64748b;">(${item.dayName})</span>
                </span>
                <span class="bar-val" style="color:var(--bypass-color);">${item.bypass} bypasses <span style="font-weight:normal;color:#64748b;font-size:0.75rem;">(${rate}% of ${item.total} apps)</span></span>
              </div>
              <div class="bar-track">
                <div class="bar-fill bar-fill-bypass" style="width: ${pct}%;"></div>
              </div>
            </div>
          `;
        }).join("")}
      </div>
    `;
  }
}

// Visual 3: Bar Chart - User-wise Workload using USER_ID (Bypass, Total, or Rate %)
function renderUserWiseBypassBar(data) {
  const userBarList = document.getElementById("userBarList");
  const userBarTitle = document.getElementById("userBarTitle");
  const userBarSubtitle = document.getElementById("userBarSubtitle");
  if (!userBarList) return;

  const userMap = {};
  data.forEach(d => {
    if (!userMap[d.USER_ID]) userMap[d.USER_ID] = { total: 0, bypass: 0 };
    userMap[d.USER_ID].total++;
    if (d["Bypass Flag"] === "Bypass") userMap[d.USER_ID].bypass++;
  });

  const userList = Object.entries(userMap).map(([userId, stats]) => {
    const rate = stats.total > 0 ? (stats.bypass / stats.total) * 100 : 0;
    return { userId, ...stats, rate };
  });

  if (currentUserMetricMode === "bypass") {
    userList.sort((a, b) => b.bypass - a.bypass);
    if (userBarTitle) userBarTitle.textContent = "User-wise Bypass Count (USER_ID)";
    if (userBarSubtitle) userBarSubtitle.textContent = "Ranked by total volume of exception bypasses logged";
  } else if (currentUserMetricMode === "total") {
    userList.sort((a, b) => b.total - a.total);
    if (userBarTitle) userBarTitle.textContent = "User-wise Total Applications (USER_ID)";
    if (userBarSubtitle) userBarSubtitle.textContent = "Ranked by total overall KYC applications processed";
  } else if (currentUserMetricMode === "rate") {
    userList.sort((a, b) => b.rate - a.rate);
    if (userBarTitle) userBarTitle.textContent = "User-wise Bypass Rate % (USER_ID)";
    if (userBarSubtitle) userBarSubtitle.textContent = "Ranked by percentage proportion of bypasses to total volume";
  }

  const maxVal = Math.max(...userList.map(u => {
    if (currentUserMetricMode === "bypass") return u.bypass;
    if (currentUserMetricMode === "total") return u.total;
    return u.rate;
  }), 1);

  if (userList.length === 0) {
    userBarList.innerHTML = `<div style="color:var(--text-muted);font-size:0.85rem;padding:20px;text-align:center;">No records</div>`;
    return;
  }

  userBarList.innerHTML = userList.map(u => {
    let metricVal = u.bypass;
    let labelText = `${u.bypass} bypass`;
    let subText = `(${u.rate.toFixed(0)}%)`;
    let fillClass = "bar-fill-bypass";

    if (currentUserMetricMode === "total") {
      metricVal = u.total;
      labelText = `${u.total} total apps`;
      subText = `(${u.bypass} bypass)`;
      fillClass = "bar-fill-user";
    } else if (currentUserMetricMode === "rate") {
      metricVal = u.rate;
      labelText = `${u.rate.toFixed(1)}% bypass rate`;
      subText = `(${u.bypass}/${u.total})`;
      fillClass = "bar-fill-bypass";
    }

    const pctWidth = (metricVal / maxVal) * 100;

    return `
      <div class="bar-row">
        <div class="bar-labels">
          <span class="bar-name">
            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/></svg>
            ${u.userId}
          </span>
          <span class="bar-val">${labelText} <span style="font-weight:normal;color:#94a3b8;font-size:0.75rem;">${subText}</span></span>
        </div>
        <div class="bar-track">
          <div class="bar-fill ${fillClass}" style="width: ${pctWidth}%;"></div>
        </div>
      </div>
    `;
  }).join("");
}

// Visual 4: Bar Chart - Assignee-wise Workload using ASSIGNED_TO (Bypass, Total, or Rate %)
function renderAssigneeWiseBypassBar(data) {
  const assigneeBarList = document.getElementById("assigneeBarList");
  const assigneeBarTitle = document.getElementById("assigneeBarTitle");
  const assigneeBarSubtitle = document.getElementById("assigneeBarSubtitle");
  if (!assigneeBarList) return;

  const assigneeMap = {};
  data.forEach(d => {
    if (!assigneeMap[d.ASSIGNED_TO]) assigneeMap[d.ASSIGNED_TO] = { total: 0, bypass: 0 };
    assigneeMap[d.ASSIGNED_TO].total++;
    if (d["Bypass Flag"] === "Bypass") assigneeMap[d.ASSIGNED_TO].bypass++;
  });

  const assigneeList = Object.entries(assigneeMap).map(([name, stats]) => {
    const rate = stats.total > 0 ? (stats.bypass / stats.total) * 100 : 0;
    return { name, ...stats, rate };
  });

  if (currentAssigneeMetricMode === "bypass") {
    assigneeList.sort((a, b) => b.bypass - a.bypass);
    if (assigneeBarTitle) assigneeBarTitle.textContent = "Assignee-wise Bypass Count (ASSIGNED_TO)";
    if (assigneeBarSubtitle) assigneeBarSubtitle.textContent = "Workload and exception allocation per assigned operational analyst";
  } else if (currentAssigneeMetricMode === "total") {
    assigneeList.sort((a, b) => b.total - a.total);
    if (assigneeBarTitle) assigneeBarTitle.textContent = "Assignee-wise Total Applications (ASSIGNED_TO)";
    if (assigneeBarSubtitle) assigneeBarSubtitle.textContent = "Total cases assigned per operational analyst";
  } else if (currentAssigneeMetricMode === "rate") {
    assigneeList.sort((a, b) => b.rate - a.rate);
    if (assigneeBarTitle) assigneeBarTitle.textContent = "Assignee-wise Bypass Rate % (ASSIGNED_TO)";
    if (assigneeBarSubtitle) assigneeBarSubtitle.textContent = "Proportion of bypasses out of total analyst assignments";
  }

  const maxVal = Math.max(...assigneeList.map(a => {
    if (currentAssigneeMetricMode === "bypass") return a.bypass;
    if (currentAssigneeMetricMode === "total") return a.total;
    return a.rate;
  }), 1);

  if (assigneeList.length === 0) {
    assigneeBarList.innerHTML = `<div style="color:var(--text-muted);font-size:0.85rem;padding:20px;text-align:center;">No records</div>`;
    return;
  }

  assigneeBarList.innerHTML = assigneeList.map(a => {
    let metricVal = a.bypass;
    let labelText = `${a.bypass} bypass`;
    let subText = `(${a.rate.toFixed(0)}%)`;
    let fillClass = "bar-fill-bypass";

    if (currentAssigneeMetricMode === "total") {
      metricVal = a.total;
      labelText = `${a.total} total apps`;
      subText = `(${a.bypass} bypass)`;
      fillClass = "bar-fill-user";
    } else if (currentAssigneeMetricMode === "rate") {
      metricVal = a.rate;
      labelText = `${a.rate.toFixed(1)}% bypass rate`;
      subText = `(${a.bypass}/${a.total})`;
      fillClass = "bar-fill-bypass";
    }

    const pctWidth = (metricVal / maxVal) * 100;

    return `
      <div class="bar-row">
        <div class="bar-labels">
          <span class="bar-name">
            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="2" y="7" width="20" height="14" rx="2" ry="2"/><path d="M16 21V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v16"/></svg>
            ${a.name}
          </span>
          <span class="bar-val">${labelText} <span style="font-weight:normal;color:#94a3b8;font-size:0.75rem;">${subText}</span></span>
        </div>
        <div class="bar-track">
          <div class="bar-fill ${fillClass}" style="width: ${pctWidth}%;"></div>
        </div>
      </div>
    `;
  }).join("");
}

// State for dedicated Pie Chart
let currentPieSubMode = "reason"; // "reason" | "split"

// Visual 5A: Dedicated Interactive Pie Chart (Full Multi-Color Slice Pie)
function renderStandalonePieChart(measures, data) {
  const { totalApps, bypassApps, regularApps, bypassPct } = measures;
  const pieSvg = document.getElementById("standalonePieSvg");
  const pieLegend = document.getElementById("standalonePieLegend");
  const pieChartTitle = document.getElementById("pieChartTitle");
  const pieChartSubtitle = document.getElementById("pieChartSubtitle");

  if (!pieSvg) return;

  if (totalApps === 0) {
    pieSvg.innerHTML = `<circle cx="50" cy="50" r="44" fill="#e2e8f0"/>`;
    if (pieLegend) pieLegend.innerHTML = `<div style="color:var(--text-muted);font-size:0.8rem;">No records matching filters</div>`;
    return;
  }

  const cx = 50;
  const cy = 50;
  const r = 44;

  if (currentPieSubMode === "reason") {
    if (pieChartTitle) pieChartTitle.textContent = "Pie Chart: PIVC Exception Reason Breakdown";
    if (pieChartSubtitle) pieChartSubtitle.textContent = "Granular distribution of operational bypass exception triggers";

    const catCounts = {
      "Biometric & Face Liveness": 0,
      "Network & Timeout Glitch": 0,
      "Camera & Hardware Issue": 0,
      "Special Policy / Corporate": 0,
      "Branch & Wet Signature": 0,
      "System Done (Regular)": 0
    };

    data.forEach(d => {
      const rem = (d.REMARKS || "").toLowerCase();
      if (d["Bypass Flag"] === "Regular") {
        catCounts["System Done (Regular)"]++;
      } else if (rem.includes("biometric") || rem.includes("liveness") || rem.includes("face") || rem.includes("match")) {
        catCounts["Biometric & Face Liveness"]++;
      } else if (rem.includes("network") || rem.includes("bandwidth") || rem.includes("timeout") || rem.includes("latency") || rem.includes("outage") || rem.includes("server") || rem.includes("connection")) {
        catCounts["Network & Timeout Glitch"]++;
      } else if (rem.includes("camera") || rem.includes("driver") || rem.includes("hardware") || rem.includes("microphone") || rem.includes("audio")) {
        catCounts["Camera & Hardware Issue"]++;
      } else if (rem.includes("special") || rem.includes("vip") || rem.includes("director") || rem.includes("corporate") || rem.includes("overseas") || rem.includes("travel") || rem.includes("senior") || rem.includes("disability") || rem.includes("speech")) {
        catCounts["Special Policy / Corporate"]++;
      } else {
        catCounts["Branch & Wet Signature"]++;
      }
    });

    const categories = Object.entries(catCounts)
      .map(([name, count]) => ({ name, count, pct: totalApps > 0 ? (count / totalApps) * 100 : 0 }))
      .filter(c => c.count > 0)
      .sort((a, b) => b.count - a.count);

    const colors = ["#d83b01", "#0078d4", "#ffaa44", "#107c41", "#8764b8", "#038387"];
    let startAngle = 0;
    let pathsHtml = "";

    categories.forEach((cat, idx) => {
      const angle = (cat.count / totalApps) * 360;
      const endAngle = startAngle + angle;
      const color = colors[idx % colors.length];

      if (categories.length === 1) {
        pathsHtml += `<circle cx="${cx}" cy="${cy}" r="${r}" fill="${color}"><title>${cat.name}: 100%</title></circle>`;
      } else {
        const startRad = (startAngle * Math.PI) / 180;
        const endRad = (endAngle * Math.PI) / 180;
        const x1 = cx + r * Math.sin(startRad);
        const y1 = cy - r * Math.cos(startRad);
        const x2 = cx + r * Math.sin(endRad);
        const y2 = cy - r * Math.cos(endRad);
        const largeArc = angle > 180 ? 1 : 0;

        pathsHtml += `
          <path d="M ${cx} ${cy} L ${x1} ${y1} A ${r} ${r} 0 ${largeArc} 1 ${x2} ${y2} Z"
            fill="${color}" stroke="#ffffff" stroke-width="1.5" class="pie-slice">
            <title>${cat.name}: ${cat.count} cases (${cat.pct.toFixed(1)}%)</title>
          </path>
        `;
      }
      startAngle = endAngle;
    });

    pieSvg.innerHTML = pathsHtml;

    if (pieLegend) {
      pieLegend.innerHTML = categories.map((cat, idx) => {
        const color = colors[idx % colors.length];
        return `
          <div class="legend-item" style="font-size:0.75rem;">
            <span class="legend-dot" style="background:${color};"></span>
            <span style="max-width:130px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;" title="${cat.name}">${cat.name}</span>
            <span class="legend-metric">${cat.count} (${cat.pct.toFixed(1)}%)</span>
          </div>
        `;
      }).join("");
    }

  } else if (currentPieSubMode === "split") {
    if (pieChartTitle) pieChartTitle.textContent = "Pie Chart: Bypass vs Regular Slices";
    if (pieChartSubtitle) pieChartSubtitle.textContent = "2-slice solid pie chart showing overall volume share";

    const bypassAngle = (bypassApps / totalApps) * 360;
    const bypassRad = (bypassAngle * Math.PI) / 180;
    const x2 = cx + r * Math.sin(bypassRad);
    const y2 = cy - r * Math.cos(bypassRad);
    const largeArcBypass = bypassAngle > 180 ? 1 : 0;
    const largeArcRegular = (360 - bypassAngle) > 180 ? 1 : 0;

    let pathsHtml = "";
    if (bypassApps === totalApps) {
      pathsHtml = `<circle cx="${cx}" cy="${cy}" r="${r}" fill="#d83b01"><title>Bypass: 100%</title></circle>`;
    } else if (regularApps === totalApps) {
      pathsHtml = `<circle cx="${cx}" cy="${cy}" r="${r}" fill="#0078d4"><title>Regular: 100%</title></circle>`;
    } else {
      pathsHtml = `
        <path d="M ${cx} ${cy} L 50 ${cy - r} A ${r} ${r} 0 ${largeArcBypass} 1 ${x2} ${y2} Z"
          fill="#d83b01" stroke="#ffffff" stroke-width="1.5" class="pie-slice">
          <title>Bypass: ${bypassApps} (${bypassPct.toFixed(1)}%)</title>
        </path>
        <path d="M ${cx} ${cy} L ${x2} ${y2} A ${r} ${r} 0 ${largeArcRegular} 1 50 ${cy - r} Z"
          fill="#0078d4" stroke="#ffffff" stroke-width="1.5" class="pie-slice">
          <title>Regular: ${regularApps} (${((regularApps/totalApps)*100).toFixed(1)}%)</title>
        </path>
      `;
    }

    pieSvg.innerHTML = pathsHtml;

    if (pieLegend) {
      const regPct = totalApps > 0 ? (regularApps / totalApps) * 100 : 0;
      pieLegend.innerHTML = `
        <div class="legend-item">
          <span class="legend-dot bypass"></span>
          <span>Bypass Cases</span>
          <span class="legend-metric">${bypassApps.toLocaleString()} (${bypassPct.toFixed(1)}%)</span>
        </div>
        <div class="legend-item">
          <span class="legend-dot regular"></span>
          <span>Regular Cases</span>
          <span class="legend-metric">${regularApps.toLocaleString()} (${regPct.toFixed(1)}%)</span>
        </div>
      `;
    }
  }
}

// Visual 5B: Dedicated Donut Chart - Bypass vs Non-Bypass
function renderDonutChart(measures) {
  const { totalApps, bypassApps, regularApps, bypassPct } = measures;
  const donutSvg = document.getElementById("donutSvg");
  const donutCenterText = document.getElementById("donutCenterText");
  const donutPctText = document.getElementById("donutPctText");
  const donutBypassLegend = document.getElementById("donutBypassLegend");
  const donutRegularLegend = document.getElementById("donutRegularLegend");

  if (!donutSvg) return;

  if (totalApps === 0) {
    donutSvg.innerHTML = `<circle cx="50" cy="50" r="38" fill="none" stroke="#e2e8f0" stroke-width="12"/>`;
    if (donutPctText) donutPctText.textContent = "0.0%";
    return;
  }

  if (donutPctText) donutPctText.textContent = bypassPct.toFixed(1) + "%";

  const radius = 38;
  const circumference = 2 * Math.PI * radius;
  const bypassArc = (bypassApps / totalApps) * circumference;

  donutSvg.innerHTML = `
    <!-- Regular Background Ring -->
    <circle cx="50" cy="50" r="${radius}" fill="none" stroke="#0078d4" stroke-width="12"
      stroke-dasharray="${circumference}" stroke-dashoffset="0" transform="rotate(-90 50 50)"/>
    <!-- Bypass Highlight Segment -->
    <circle cx="50" cy="50" r="${radius}" fill="none" stroke="#d83b01" stroke-width="12"
      stroke-dasharray="${bypassArc} ${circumference}" stroke-dashoffset="0" transform="rotate(-90 50 50)" stroke-linecap="round"/>
  `;

  if (donutBypassLegend) {
    donutBypassLegend.textContent = `${bypassApps.toLocaleString()} (${bypassPct.toFixed(1)}%)`;
  }
  if (donutRegularLegend) {
    const regPct = totalApps > 0 ? (regularApps / totalApps) * 100 : 0;
    donutRegularLegend.textContent = `${regularApps.toLocaleString()} (${regPct.toFixed(1)}%)`;
  }
}



// Matrix Tab State: "monthly" or "daily"
let currentMatrixView = "monthly";

function updateMatrixTabs() {
  const tabMonthlyBtn = document.getElementById("tabMonthlyBtn");
  const tabDailyBtn = document.getElementById("tabDailyBtn");
  const matrixTitle = document.getElementById("matrixTitle");
  const matrixSubtitle = document.getElementById("matrixSubtitle");

  if (tabMonthlyBtn && tabDailyBtn) {
    if (currentMatrixView === "monthly") {
      tabMonthlyBtn.style.background = "#0078d4";
      tabMonthlyBtn.style.color = "#ffffff";
      tabDailyBtn.style.background = "#ffffff";
      tabDailyBtn.style.color = "var(--text-main)";
      if (matrixTitle) matrixTitle.textContent = "Executive Monthly Summary Matrix";
      if (matrixSubtitle) matrixSubtitle.textContent = "Aggregated monthly volume and bypass percentage metrics";
    } else {
      tabDailyBtn.style.background = "#0078d4";
      tabDailyBtn.style.color = "#ffffff";
      tabMonthlyBtn.style.background = "#ffffff";
      tabMonthlyBtn.style.color = "var(--text-main)";
      if (matrixTitle) matrixTitle.textContent = "Daily Bypass Matrix (How Many in Each Day)";
      if (matrixSubtitle) matrixSubtitle.textContent = "Granular date-by-date bypass frequencies, volumes, and percentages";
    }
  }
}

// Visual 6: Matrix / Table - Monthly or Daily Breakdown
function renderMatrixTable(data) {
  const matrixHead = document.getElementById("matrixTableHead");
  updateMatrixTabs();

  if (currentMatrixView === "daily") {
    // -------------------------------------------------------------
    // DAILY BREAKDOWN TABLE (How many bypass in each day)
    // -------------------------------------------------------------
    if (matrixHead) {
      matrixHead.innerHTML = `
        <tr>
          <th>Date</th>
          <th>Day of Week</th>
          <th>Total Applications</th>
          <th>Bypass Count</th>
          <th>Bypass %</th>
        </tr>
      `;
    }

    const dayMap = new Map();
    data.forEach(d => {
      const dt = d.APPROVAL_DATE;
      if (!dayMap.has(dt)) {
        const dow = d.DateObj && !isNaN(d.DateObj.getTime())
          ? d.DateObj.toLocaleDateString("en-US", { weekday: "short" })
          : "";
        dayMap.set(dt, { date: dt, dow: dow, total: 0, bypass: 0 });
      }
      const cur = dayMap.get(dt);
      cur.total++;
      if (d["Bypass Flag"] === "Bypass") cur.bypass++;
    });

    const sortedDays = Array.from(dayMap.values()).sort((a, b) => new Date(a.date) - new Date(b.date));

    if (sortedDays.length === 0) {
      matrixTableBody.innerHTML = `<tr><td colspan="5" style="text-align:center;padding:24px;color:var(--text-muted);">No records found matching filters</td></tr>`;
      return;
    }

    let totalAll = 0;
    let bypassAll = 0;

    const rowsHtml = sortedDays.map(d => {
      totalAll += d.total;
      bypassAll += d.bypass;
      const pct = d.total > 0 ? (d.bypass / d.total) * 100 : 0;
      const isPeak = d.bypass >= 5;

      return `
        <tr>
          <td style="font-family:monospace;font-weight:600;">${d.date}</td>
          <td><span style="font-size:0.75rem;background:#f1f5f9;padding:2px 6px;border-radius:4px;font-weight:600;color:#475569;">${d.dow}</span></td>
          <td><strong>${d.total}</strong></td>
          <td>
            <span style="font-weight:700;color:${d.bypass > 0 ? 'var(--bypass-color)' : '#94a3b8'};">
              ${d.bypass} ${isPeak ? '<span style="font-size:0.7rem;background:#fee2e2;color:#b91c1c;padding:1px 5px;border-radius:3px;margin-left:4px;">Peak</span>' : ''}
            </span>
          </td>
          <td>
            <div class="table-pct-cell">
              <span style="font-weight:700;min-width:44px;">${pct.toFixed(1)}%</span>
              <div class="table-pct-bar">
                <div class="table-pct-fill" style="width:${Math.min(pct, 100)}%;"></div>
              </div>
            </div>
          </td>
        </tr>
      `;
    }).join("");

    const overallPct = totalAll > 0 ? (bypassAll / totalAll) * 100 : 0;
    const footerHtml = `
      <tr style="background:#f1f5f9;font-weight:700;border-top:2px solid var(--border-color);">
        <td colspan="2">Total (${sortedDays.length} days)</td>
        <td>${totalAll}</td>
        <td style="color:var(--bypass-color);">${bypassAll}</td>
        <td>
          <div class="table-pct-cell">
            <span style="font-weight:800;color:var(--primary);min-width:44px;">${overallPct.toFixed(1)}%</span>
            <div class="table-pct-bar">
              <div class="table-pct-fill" style="width:${Math.min(overallPct, 100)}%;background:var(--primary);"></div>
            </div>
          </div>
        </td>
      </tr>
    `;

    matrixTableBody.innerHTML = rowsHtml + footerHtml;

  } else {
    // -------------------------------------------------------------
    // MONTHLY SUMMARY TABLE
    // -------------------------------------------------------------
    if (matrixHead) {
      matrixHead.innerHTML = `
        <tr>
          <th>Month</th>
          <th>Total Applications</th>
          <th>Bypass Applications</th>
          <th>Bypass %</th>
        </tr>
      `;
    }

    const monthsMap = new Map();
    data.forEach(d => {
      const key = d["Month-Year Sort"];
      if (!monthsMap.has(key)) {
        monthsMap.set(key, { name: d["Month-Year"], total: 0, bypass: 0, sortKey: key });
      }
      const cur = monthsMap.get(key);
      cur.total++;
      if (d["Bypass Flag"] === "Bypass") {
        cur.bypass++;
      }
    });

    const sortedMonths = Array.from(monthsMap.values()).sort((a, b) => a.sortKey - b.sortKey);

    if (sortedMonths.length === 0) {
      matrixTableBody.innerHTML = `<tr><td colspan="4" style="text-align:center;padding:24px;color:var(--text-muted);">No records found matching filters</td></tr>`;
      return;
    }

    let totalAll = 0;
    let bypassAll = 0;

    const rowsHtml = sortedMonths.map(m => {
      totalAll += m.total;
      bypassAll += m.bypass;
      const pct = m.total > 0 ? (m.bypass / m.total) * 100 : 0;

      return `
        <tr>
          <td style="font-weight:600;">${m.name}</td>
          <td><strong>${m.total}</strong></td>
          <td><span style="color:var(--bypass-color);font-weight:700;">${m.bypass}</span></td>
          <td>
            <div class="table-pct-cell">
              <span style="font-weight:700;min-width:44px;">${pct.toFixed(1)}%</span>
              <div class="table-pct-bar">
                <div class="table-pct-fill" style="width:${Math.min(pct, 100)}%;"></div>
              </div>
            </div>
          </td>
        </tr>
      `;
    }).join("");

    const overallPct = totalAll > 0 ? (bypassAll / totalAll) * 100 : 0;
    const footerHtml = `
      <tr style="background:#f1f5f9;font-weight:700;border-top:2px solid var(--border-color);">
        <td>Total / Average</td>
        <td>${totalAll}</td>
        <td style="color:var(--bypass-color);">${bypassAll}</td>
        <td>
          <div class="table-pct-cell">
            <span style="font-weight:800;color:var(--primary);min-width:44px;">${overallPct.toFixed(1)}%</span>
            <div class="table-pct-bar">
              <div class="table-pct-fill" style="width:${Math.min(overallPct, 100)}%;background:var(--primary);"></div>
            </div>
          </div>
        </td>
      </tr>
    `;

    matrixTableBody.innerHTML = rowsHtml + footerHtml;
  }
}

// Drilldown Explorer Table with Pagination for 18,744+ records
function renderDrilldownTable(data) {
  const query = searchInput ? searchInput.value.toLowerCase().trim() : "";
  const filtered = query
    ? data.filter(r => 
        r.APPLICATION_NO.toLowerCase().includes(query) ||
        r.ASSIGNED_TO.toLowerCase().includes(query) ||
        r.USER_ID.toLowerCase().includes(query) ||
        r.REMARKS.toLowerCase().includes(query) ||
        r.APPROVAL_STATUS.toLowerCase().includes(query)
      )
    : data;

  const totalRecords = filtered.length;
  const totalPages = Math.max(1, Math.ceil(totalRecords / pageSize));
  if (currentPage > totalPages) currentPage = totalPages;
  if (currentPage < 1) currentPage = 1;

  const startIdx = (currentPage - 1) * pageSize;
  const endIdx = Math.min(startIdx + pageSize, totalRecords);
  const displayRows = filtered.slice(startIdx, endIdx);

  if (recordCountEl) {
    recordCountEl.textContent = `Showing ${totalRecords.toLocaleString()} of ${data.length.toLocaleString()} applications`;
  }

  // Update Pagination Controls
  const paginationSummary = document.getElementById("paginationSummary");
  const pageIndicator = document.getElementById("pageIndicator");
  const firstPageBtn = document.getElementById("firstPageBtn");
  const prevPageBtn = document.getElementById("prevPageBtn");
  const nextPageBtn = document.getElementById("nextPageBtn");
  const lastPageBtn = document.getElementById("lastPageBtn");

  if (paginationSummary) {
    paginationSummary.textContent = totalRecords > 0 
      ? `Showing ${(startIdx + 1).toLocaleString()} to ${endIdx.toLocaleString()} of ${totalRecords.toLocaleString()} records`
      : `No records found`;
  }

  if (pageIndicator) {
    pageIndicator.textContent = `Page ${currentPage} of ${totalPages}`;
  }

  if (firstPageBtn) firstPageBtn.disabled = (currentPage === 1);
  if (prevPageBtn) prevPageBtn.disabled = (currentPage === 1);
  if (nextPageBtn) nextPageBtn.disabled = (currentPage === totalPages || totalRecords === 0);
  if (lastPageBtn) lastPageBtn.disabled = (currentPage === totalPages || totalRecords === 0);

  const tableWrapper = document.getElementById("drilldownTableWrapper");
  if (tableWrapper) {
    tableWrapper.scrollTop = 0;
  }

  if (displayRows.length === 0) {
    drilldownTableBody.innerHTML = `<tr><td colspan="7" style="text-align:center;padding:24px;color:var(--text-muted);">No applications match the search/filter criteria</td></tr>`;
    return;
  }

  drilldownTableBody.innerHTML = displayRows.map(r => {
    const isBypass = r["Bypass Flag"] === "Bypass";
    const statusVal = r.APPROVAL_STATUS || "";
    const isException = statusVal.toLowerCase().includes("except");
    const isNull = !statusVal || statusVal.toLowerCase() === "null" || r.IS_STATUS_NULL;

    let statusBadge = "";
    if (isException) {
      statusBadge = `<span class="badge-status status-exception">⚠️ ${statusVal || 'Exceptional'}</span>`;
    } else if (isNull) {
      statusBadge = `<span class="badge-status status-null">null (System Done)</span>`;
    } else if (statusVal.toUpperCase().includes("DPIVC") || statusVal.toUpperCase().includes("DONE")) {
      statusBadge = `<span class="badge-status status-approved">✓ ${statusVal}</span>`;
    } else {
      const statusClass = 
        statusVal.toLowerCase() === "approved" ? "status-approved" :
        statusVal.toLowerCase() === "rejected" ? "status-rejected" : "status-pending";
      statusBadge = `<span class="badge-status ${statusClass}">${statusVal}</span>`;
    }

    const triggerInfo = isBypass 
      ? `<span class="badge-tag badge-bypass">⚠️ Bypass</span><div style="font-size:0.68rem;color:#d83b01;margin-top:3px;font-weight:600;">${r["Bypass Trigger"] || "Bypass"}</div>`
      : `<span class="badge-tag badge-regular">✓ Regular</span><div style="font-size:0.68rem;color:#64748b;margin-top:3px;">${isNull ? 'System Done' : 'Standard'}</div>`;

    return `
      <tr>
        <td style="font-family:monospace;font-weight:600;">${r.APPLICATION_NO}</td>
        <td>${r.APPROVAL_DATE}</td>
        <td>${statusBadge}</td>
        <td><span style="font-weight:600;color:var(--accent-blue);">${r.USER_ID}</span></td>
        <td>${r.ASSIGNED_TO}</td>
        <td>${triggerInfo}</td>
        <td style="font-size:0.78rem;color:#475569;max-width:320px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;" title="${r.REMARKS}">
          ${r.REMARKS}
        </td>
      </tr>
    `;
  }).join("");
}

// Export Filtered Records to CSV
function exportCurrentDataToCSV() {
  const data = getFilteredData();
  const headers = ["APPLICATION_NO", "APPROVAL_DATE", "APPROVAL_STATUS", "USER_ID", "ASSIGNED_TO", "REMARKS", "Bypass Flag", "Month-Year"];
  const csvRows = [headers.join(",")];

  data.forEach(row => {
    const values = [
      `"${row.APPLICATION_NO}"`,
      `"${row.APPROVAL_DATE}"`,
      `"${row.APPROVAL_STATUS}"`,
      `"${row.USER_ID}"`,
      `"${row.ASSIGNED_TO}"`,
      `"${(row.REMARKS || '').replace(/"/g, '""')}"`,
      `"${row["Bypass Flag"]}"`,
      `"${row["Month-Year"]}"`
    ];
    csvRows.push(values.join(","));
  });

  const blob = new Blob([csvRows.join("\n")], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `PIVC_Bypass_Analysis_Export_${new Date().toISOString().slice(0, 10)}.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// Excel & CSV File Processing Engine
function processUploadedFile(file) {
  if (!file) return;

  const datasetLabel = document.getElementById("datasetLabel");
  const bannerHeading = document.getElementById("bannerHeading");
  const bannerSubtext = document.getElementById("bannerSubtext");
  const dropBanner = document.getElementById("dropBanner");

  if (datasetLabel) datasetLabel.textContent = `Reading ${file.name}...`;

  const fileName = file.name.toLowerCase();

  // If CSV File
  if (fileName.endsWith(".csv")) {
    const reader = new FileReader();
    reader.onload = function(e) {
      try {
        const text = e.target.result;
        const parsedRows = parseCsvText(text);
        if (!parsedRows || parsedRows.length === 0) throw new Error("CSV file contains no records");
        applyNewDataset(parsedRows, file.name);
      } catch (err) {
        alert("Error parsing CSV file: " + err.message);
      }
    };
    reader.readAsText(file);
    return;
  }

  // If Excel File (.xlsx or .xls)
  if (typeof XLSX === "undefined") {
    alert("Excel parsing library (SheetJS) is still loading or unavailable. Please ensure internet access or try exporting as CSV.");
    return;
  }

  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array", cellDates: true });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const jsonRows = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

      if (!jsonRows || jsonRows.length === 0) {
        alert("The uploaded Excel sheet contains no data rows.");
        return;
      }

      applyNewDataset(jsonRows, file.name);
    } catch (err) {
      alert("Error parsing Excel file: " + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

// Apply newly uploaded dataset and refresh dashboard
function applyNewDataset(rawRows, filename) {
  processedData = transformRows(rawRows);
  currentPage = 1;

  // Reset Slicers to "All"
  activeFilters = {
    monthYear: "All",
    approvalStatus: "All",
    userId: "All",
    assignedTo: "All"
  };

  populateSlicers();

  // Update Header Indicator
  const datasetLabel = document.getElementById("datasetLabel");
  if (datasetLabel) {
    datasetLabel.textContent = `${filename} (${processedData.length.toLocaleString()} rows)`;
    datasetLabel.style.color = "#86efac";
  }

  // Update Callout Banner with celebration style
  const bannerHeading = document.getElementById("bannerHeading");
  const bannerSubtext = document.getElementById("bannerSubtext");
  const dropBanner = document.getElementById("dropBanner");

  if (dropBanner) {
    dropBanner.style.background = "#ecfdf5";
    dropBanner.style.borderColor = "#107c41";
  }
  if (bannerHeading) {
    bannerHeading.innerHTML = `✅ Successfully Loaded <strong>${processedData.length.toLocaleString()}</strong> Records from <em>${filename}</em>!`;
    bannerHeading.style.color = "#065f46";
  }
  if (bannerSubtext) {
    bannerSubtext.textContent = "All KPIs, monthly trends, user performance charts, donut breakdown, and summary matrix have been updated with your original dataset.";
  }

  updateDashboard();
}

// Master Update Function
function updateDashboard() {
  const data = getFilteredData();
  const measures = computeMeasures(data);

  try { renderKPIs(measures); } catch (e) { console.error("Error in renderKPIs:", e); }
  try { renderMonthlyBypassColumnChart(data); } catch (e) { console.error("Error in renderMonthlyBypassColumnChart:", e); }
  try { renderMonthlyTrendLineChart(data); } catch (e) { console.error("Error in renderMonthlyTrendLineChart:", e); }
  try { renderStackedMonthlyBarChart(data); } catch (e) { console.error("Error in renderStackedMonthlyBarChart:", e); }
  try { renderDailyBypassChart(data); } catch (e) { console.error("Error in renderDailyBypassChart:", e); }
  try { renderUserWiseBypassBar(data); } catch (e) { console.error("Error in renderUserWiseBypassBar:", e); }
  try { renderAssigneeWiseBypassBar(data); } catch (e) { console.error("Error in renderAssigneeWiseBypassBar:", e); }
  try { renderStandalonePieChart(measures, data); } catch (e) { console.error("Error in renderStandalonePieChart:", e); }
  try { renderDonutChart(measures); } catch (e) { console.error("Error in renderDonutChart:", e); }
  try { renderMatrixTable(data); } catch (e) { console.error("Error in renderMatrixTable:", e); }
  try { renderDrilldownTable(data); } catch (e) { console.error("Error in renderDrilldownTable:", e); }
}

// Setup Daily Visual and Matrix Tab Listeners
function setupDailyListeners() {
  const dailyViewMode = document.getElementById("dailyViewMode");
  if (dailyViewMode) {
    dailyViewMode.addEventListener("change", () => {
      renderDailyBypassChart(getFilteredData());
    });
  }

  const tabMonthlyBtn = document.getElementById("tabMonthlyBtn");
  const tabDailyBtn = document.getElementById("tabDailyBtn");

  if (tabMonthlyBtn) {
    tabMonthlyBtn.addEventListener("click", () => {
      currentMatrixView = "monthly";
      renderMatrixTable(getFilteredData());
    });
  }

  if (tabDailyBtn) {
    tabDailyBtn.addEventListener("click", () => {
      currentMatrixView = "daily";
      renderMatrixTable(getFilteredData());
    });
  }
}

// Setup File Upload & Drag-and-Drop Listeners
function setupFileUpload() {
  const fileInput = document.getElementById("excelFileInput");
  const bannerFileInput = document.getElementById("bannerFileInput");
  const dropBanner = document.getElementById("dropBanner");

  if (fileInput) {
    fileInput.addEventListener("change", e => {
      if (e.target.files && e.target.files[0]) {
        processUploadedFile(e.target.files[0]);
      }
    });
  }

  if (bannerFileInput) {
    bannerFileInput.addEventListener("change", e => {
      if (e.target.files && e.target.files[0]) {
        processUploadedFile(e.target.files[0]);
      }
    });
  }

  // Drag & drop on banner
  if (dropBanner) {
    ["dragenter", "dragover"].forEach(eventName => {
      dropBanner.addEventListener(eventName, e => {
        e.preventDefault();
        dropBanner.classList.add("dragover");
      }, false);
    });

    ["dragleave", "drop"].forEach(eventName => {
      dropBanner.addEventListener(eventName, e => {
        e.preventDefault();
        dropBanner.classList.remove("dragover");
      }, false);
    });

    dropBanner.addEventListener("drop", e => {
      e.preventDefault();
      if (e.dataTransfer.files && e.dataTransfer.files[0]) {
        processUploadedFile(e.dataTransfer.files[0]);
      }
    });
  }
}

// Setup Pagination Control Listeners
function setupPagination() {
  const firstPageBtn = document.getElementById("firstPageBtn");
  const prevPageBtn = document.getElementById("prevPageBtn");
  const nextPageBtn = document.getElementById("nextPageBtn");
  const lastPageBtn = document.getElementById("lastPageBtn");
  const pageSizeSelect = document.getElementById("pageSizeSelect");

  if (firstPageBtn) firstPageBtn.addEventListener("click", () => { currentPage = 1; renderDrilldownTable(getFilteredData()); });
  if (prevPageBtn) prevPageBtn.addEventListener("click", () => { if (currentPage > 1) { currentPage--; renderDrilldownTable(getFilteredData()); } });
  if (nextPageBtn) nextPageBtn.addEventListener("click", () => { currentPage++; renderDrilldownTable(getFilteredData()); });
  if (lastPageBtn) {
    lastPageBtn.addEventListener("click", () => {
      const filtered = getFilteredData();
      currentPage = Math.max(1, Math.ceil(filtered.length / pageSize));
      renderDrilldownTable(filtered);
    });
  }
  if (pageSizeSelect) {
    pageSizeSelect.addEventListener("change", e => {
      pageSize = parseInt(e.target.value, 10);
      currentPage = 1;
      renderDrilldownTable(getFilteredData());
    });
  }
}

// Setup Visual 1 (card-span-7) Whole Month Dates Scrolling & Toggle Controls
function setupV1ScrollControls() {
  const btnMonthly = document.getElementById("v1BtnMonthly");
  const btnDates = document.getElementById("v1BtnDates");
  const monthPicker = document.getElementById("v1MonthPicker");
  const scrollContainer = document.getElementById("barChartMonthlyScroll");
  const slider = document.getElementById("v1ScrollSlider");
  const jumpStart = document.getElementById("v1JumpStart");
  const jumpMid = document.getElementById("v1JumpMid");
  const jumpEnd = document.getElementById("v1JumpEnd");

  // Toggle buttons
  if (btnMonthly) {
    btnMonthly.addEventListener("click", () => {
      v1CurrentMode = "monthly";
      renderMonthlyBypassColumnChart(getFilteredData());
    });
  }

  if (btnDates) {
    btnDates.addEventListener("click", () => {
      v1CurrentMode = "dates";
      renderMonthlyBypassColumnChart(getFilteredData());
    });
  }

  // Month Picker for whole month dates
  if (monthPicker) {
    monthPicker.addEventListener("change", (e) => {
      v1SelectedMonth = e.target.value;
      renderMonthlyBypassColumnChart(getFilteredData());
    });
  }

  // Interactive Range Slider Scrubber
  if (slider && scrollContainer) {
    slider.addEventListener("input", (e) => {
      const maxScroll = scrollContainer.scrollWidth - scrollContainer.clientWidth;
      if (maxScroll > 0) {
        scrollContainer.scrollLeft = (parseFloat(e.target.value) / 100) * maxScroll;
      }
    });

    // Synchronize slider while user scrolls container directly
    scrollContainer.addEventListener("scroll", () => {
      const maxScroll = scrollContainer.scrollWidth - scrollContainer.clientWidth;
      if (maxScroll > 0) {
        slider.value = (scrollContainer.scrollLeft / maxScroll) * 100;
      }
    });

    // Mouse wheel horizontal scroll support
    scrollContainer.addEventListener("wheel", (e) => {
      if (Math.abs(e.deltaY) > 0) {
        e.preventDefault();
        scrollContainer.scrollLeft += e.deltaY;
      }
    }, { passive: false });
  }

  // Quick Jump Buttons
  if (jumpStart && scrollContainer) {
    jumpStart.addEventListener("click", () => {
      scrollContainer.scrollTo({ left: 0, behavior: "smooth" });
    });
  }

  if (jumpMid && scrollContainer) {
    jumpMid.addEventListener("click", () => {
      const maxScroll = scrollContainer.scrollWidth - scrollContainer.clientWidth;
      scrollContainer.scrollTo({ left: maxScroll / 2, behavior: "smooth" });
    });
  }

  if (jumpEnd && scrollContainer) {
    jumpEnd.addEventListener("click", () => {
      scrollContainer.scrollTo({ left: scrollContainer.scrollWidth, behavior: "smooth" });
    });
  }
}

// Robust CSV text parser handling quotes and newlines
function parseCsvText(text) {
  const lines = text.split(/\r?\n/).filter(l => l.trim().length > 0);
  if (lines.length < 2) return [];

  function splitLine(line) {
    const fields = [];
    let cur = "";
    let inQuotes = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"') {
        if (inQuotes && line[i + 1] === '"') {
          cur += '"';
          i++;
        } else {
          inQuotes = !inQuotes;
        }
      } else if (ch === ',' && !inQuotes) {
        fields.push(cur.trim());
        cur = "";
      } else {
        cur += ch;
      }
    }
    fields.push(cur.trim());
    return fields;
  }

  const headers = splitLine(lines[0]).map(h => h.replace(/^["']|["']$/g, "").trim());
  const rows = [];
  for (let i = 1; i < lines.length; i++) {
    const values = splitLine(lines[i]);
    const obj = {};
    headers.forEach((h, idx) => {
      let v = values[idx] !== undefined ? values[idx] : "";
      v = v.replace(/^["']|["']$/g, "").trim();
      obj[h] = v;
    });
    rows.push(obj);
  }
  return rows;
}

// Auto-load PIVC_bypass_500_records.csv if available
async function autoLoadDataset() {
  try {
    const res = await fetch("PIVC_bypass_500_records.csv");
    if (res.ok) {
      const csvText = await res.text();
      const rows = parseCsvText(csvText);
      if (rows && rows.length > 0) {
        applyNewDataset(rows, "PIVC_bypass_500_records.csv");
        return true;
      }
    }
  } catch (err) {
    console.log("Auto-fetch not available:", err);
  }
  return false;
}

// Setup Chart Mode Toggle Listeners (Line graphs, Stacked bars, Pie/Donut studio, and Workload bars)
function setupChartModeListeners() {
  function bindToggleGroup(btnIds, callback) {
    btnIds.forEach(id => {
      const btn = document.getElementById(id);
      if (btn) {
        btn.addEventListener("click", () => {
          btnIds.forEach(otherId => {
            const ob = document.getElementById(otherId);
            if (ob) ob.classList.remove("active");
          });
          btn.classList.add("active");
          callback(id);
        });
      }
    });
  }

  // Visual 2 Line Modes
  bindToggleGroup(["lineModePct", "lineModeDual", "lineModeCumulative"], id => {
    if (id === "lineModePct") currentLineChartMode = "pct";
    else if (id === "lineModeDual") currentLineChartMode = "dual";
    else if (id === "lineModeCumulative") currentLineChartMode = "cumulative";
    renderMonthlyTrendLineChart(getFilteredData());
  });

  // Stacked Bar Modes
  bindToggleGroup(["barModeStacked", "barMode100", "barModeClustered"], id => {
    if (id === "barModeStacked") currentStackedBarMode = "stacked";
    else if (id === "barMode100") currentStackedBarMode = "100";
    else if (id === "barModeClustered") currentStackedBarMode = "clustered";
    renderStackedMonthlyBarChart(getFilteredData());
  });

  // Dedicated Pie Chart Modes
  bindToggleGroup(["pieSubModeReason", "pieSubModeSplit"], id => {
    if (id === "pieSubModeReason") currentPieSubMode = "reason";
    else if (id === "pieSubModeSplit") currentPieSubMode = "split";
    const data = getFilteredData();
    renderStandalonePieChart(computeMeasures(data), data);
  });

  // User Workload Modes
  bindToggleGroup(["userMetricBypass", "userMetricTotal", "userMetricRate"], id => {
    if (id === "userMetricBypass") currentUserMetricMode = "bypass";
    else if (id === "userMetricTotal") currentUserMetricMode = "total";
    else if (id === "userMetricRate") currentUserMetricMode = "rate";
    renderUserWiseBypassBar(getFilteredData());
  });

  // Assignee Workload Modes
  bindToggleGroup(["assigneeMetricBypass", "assigneeMetricTotal", "assigneeMetricRate"], id => {
    if (id === "assigneeMetricBypass") currentAssigneeMetricMode = "bypass";
    else if (id === "assigneeMetricTotal") currentAssigneeMetricMode = "total";
    else if (id === "assigneeMetricRate") currentAssigneeMetricMode = "rate";
    renderAssigneeWiseBypassBar(getFilteredData());
  });

  // Drilldown Table Height Modes
  const drilldownWrapper = document.getElementById("drilldownTableWrapper");
  bindToggleGroup(["tblHeightCompact", "tblHeightStandard", "tblHeightExpanded"], id => {
    if (!drilldownWrapper) return;
    if (id === "tblHeightCompact") {
      drilldownWrapper.style.maxHeight = "320px";
    } else if (id === "tblHeightStandard") {
      drilldownWrapper.style.maxHeight = "460px";
    } else if (id === "tblHeightExpanded") {
      drilldownWrapper.style.maxHeight = "650px";
    }
  });

  const btnScrollDrilldownTop = document.getElementById("btnScrollDrilldownTop");
  if (btnScrollDrilldownTop && drilldownWrapper) {
    btnScrollDrilldownTop.addEventListener("click", () => {
      drilldownWrapper.scrollTo({ top: 0, behavior: "smooth" });
    });
  }
}

// Master initialization
async function initDashboard() {
  initSlicers();
  setupV1ScrollControls();
  setupDailyListeners();
  setupChartModeListeners();
  setupFileUpload();
  setupPagination();

  // Render immediately with baseline data so visuals appear without delay
  updateDashboard();

  // Then auto-load CSV asynchronously if running over web server
  await autoLoadDataset();
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", initDashboard);
} else {
  initDashboard();
}


