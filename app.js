let chartObj = null;
let capexChartObj = null;
let criticalChartAcc = null;
let criticalChartBoiler = null;
let criticalChartPcp = null;
let deptChartObj = null;
let deptChartOpen = false;
let rawTaskRows = [];
let tasksLoaded = false;

let selectedDept = '';
let selectedMonth = '';

const DATA_FILE_PATH = 'Data-file.xlsx';
const PDF_ROOT = 'Communications';
const PDF_SUBFOLDERS = ['CEO Communication', 'GM Project Communication', 'EIC Communications', 'BHEL Communications','Meeting record notes','Hop Visits'];
let currentCommFolder = 'CEO Communication';
const TASK_FILE_PATH = 'Calendar-tasks.xlsx';
const UPDATES_FILE_PATH = 'Daily-updates.xlsx';
const KPI_FILE_PATH = 'KPI-details.xlsx';
const SLIDES_PDF_PATH = 'Unit 3 Update as on 23-09-26.pdf';
const ACC_PDF_PATH = 'acc_unit3_progress.pdf';
let accPdfDoc = null;

async function loadWorkbook(path){
    const response = await fetch(path + '?t=' + Date.now());
    if (!response.ok) {
        throw new Error(`Failed to load ${path}. HTTP Status: ${response.status}`);
    }
    const data = await response.arrayBuffer();
    return XLSX.read(data, {type:'array'});
}

function parseAnyDate(rawDate){
    if(rawDate instanceof Date){
        return new Date(rawDate.getFullYear(), rawDate.getMonth(), rawDate.getDate());
    }
    if(typeof rawDate === 'number'){
        const parsed = XLSX.SSF.parse_date_code(rawDate);
        return new Date(parsed.y, parsed.m - 1, parsed.d);
    }
    if(typeof rawDate === 'string'){
        const trimmed = rawDate.trim();
        let match = trimmed.match(/^(\d{1,2})[\/\.\-](\d{1,2})[\/\.\-](\d{2,4})/);
        if(match){
            let part1 = parseInt(match[1], 10);
            let part2 = parseInt(match[2], 10);
            let year = parseInt(match[3], 10);
            if(year < 100){ year += (year <= 49 ? 2000 : 1900); }
            let month = part1 > 12 ? part2 : part1;
            let day = part1 > 12 ? part1 : part2;
            const d = new Date(year, month - 1, day);
            if(!isNaN(d) && d.getFullYear() > 2000){ return d; }
        }
        const fallback = new Date(trimmed);
        if(!isNaN(fallback)){
            const fixed = new Date(fallback.getFullYear(), fallback.getMonth(), fallback.getDate());
            if(fixed.getFullYear() < 100){ fixed.setFullYear(fixed.getFullYear() + (fixed.getFullYear() <= 49 ? 2000 : 1900)); }
            return fixed;
        }
        return fallback;
    }
    return new Date(NaN);
}

function parsePercentValue(pct){
    if(pct === null || pct === undefined || pct === '') return 0;
    let num = pct;
    if(typeof num === 'string') num = parseFloat(num.replace('%',''));
    if(isNaN(num)) return 0;
    if(num <= 1) num = num * 100;
    return num;
}

/* ---------- Civil Foundations ---------- */
async function loadCivilFoundations() {
  const tableBody = document.querySelector("#civilFoundationsTable tbody");
  if (!tableBody) return;

  try {
    const workbook = await loadWorkbook(DATA_FILE_PATH);
    const sheet = workbook.Sheets["Civil foundations"];
    if (!sheet) {
      tableBody.innerHTML = '<tr><td colspan="3">Civil foundations sheet not found.</td></tr>';
      return;
    }
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    const dataRows = rows.slice(1).filter(row =>
      row && row[0] !== undefined && row[0] !== null && String(row[0]).trim() !== ""
    );
    if (!dataRows.length) {
      tableBody.innerHTML = '<tr><td colspan="3">No civil foundations data found.</td></tr>';
      return;
    }
    tableBody.innerHTML = dataRows.map(row => `
      <tr>
        <td>${civilHtml(row[0])}</td>
        <td>${civilHtml(row[1] || "")}</td>
        <td>${civilHtml(row[2] || "")}</td>
      </tr>
    `).join("");
  } catch (error) {
    console.error("Civil foundations load failed:", error);
    tableBody.innerHTML = '<tr><td colspan="3">Could not load Data-file.xlsx.</td></tr>';
  }
}

function civilHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

/* ---------- ACC 3 Progress Chart ---------- */
function loadACC(workbook){
    const sheet = workbook.Sheets['ACC 3'];
    if(!sheet) return;

    const rows = XLSX.utils.sheet_to_json(sheet);
    const months = [], manpower = [], monthlyTonnage = [], cumulative = [];

    rows.forEach(r => {
        const keys = Object.keys(r);
        months.push(r[keys[0]]);
        manpower.push(r[keys[1]]);
        monthlyTonnage.push(r[keys[2]]);
        cumulative.push(r[keys[3]]);
    });

    if(chartObj) chartObj.destroy();

    const chartEl = document.getElementById('accChart');
    if(!chartEl) return;

    chartObj = new Chart(chartEl, {
        data:{
            labels: months,
            datasets:[
                {
                    type:'bar',
                    label:'Manpower Deployed',
                    data:manpower,
                    backgroundColor:'#0078D4',
                    yAxisID:'y'
                },
                {
                    type:'bar',
                    label:'Monthly Erection Tonnage (MT)',
                    data:monthlyTonnage,
                    backgroundColor:'#82c4f0',
                    yAxisID:'y1'
                },
                {
                    type:'line',
                    label:'Cumulative Erection (MT)',
                    data:cumulative,
                    borderColor:'#ff6600',
                    backgroundColor:'#ff6600',
                    borderWidth:4,
                    tension:0.3,
                    yAxisID:'y1'
                }
            ]
        },
        options:{
            responsive:true,
            interaction:{ mode:'index' },
            scales:{
                y:{
                    beginAtZero:true,
                    title:{ display:true, text:'Manpower' }
                },
                y1:{
                    position:'right',
                    beginAtZero:true,
                    grid:{ drawOnChartArea:false },
                    title:{ display:true, text:'Tonnage (MT)' }
                }
            }
        }
    });
}


/* ---------- Manpower ---------- */
function loadManpower(workbook){
    const sheet = workbook.Sheets['Manpower'];
    if(!sheet) return;

    const asOnEl = document.getElementById('manpowerAsOnDate');
    if(asOnEl){
        const yesterday = new Date();
        yesterday.setDate(yesterday.getDate() - 1);
        asOnEl.textContent = yesterday.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
    }

    const rows = XLSX.utils.sheet_to_json(sheet);
    const tbody = document.querySelector('#manpowerTable tbody');
    tbody.innerHTML='';

    rows.forEach(row => {
        const values = Object.values(row);
        const tr = document.createElement('tr');
        if(values[0] && values[0].toString().trim() === 'Total') tr.classList.add('total-row');
        tr.innerHTML=`
            <td>${values[0]}</td>
            <td>${values[1]}</td>
            <td>${values[2]}</td>
            <td class="shortfall">${values[3]}</td>
        `;
        tbody.appendChild(tr);
    });
}

function toggleDriveDropdown(contentId, btn) {
    const content = document.getElementById(contentId);
    const isOpen = content.classList.contains('open');
    if (isOpen) {
        content.classList.remove('open');
        btn.classList.remove('open');
    } else {
        content.classList.add('open');
        btn.classList.add('open');
    }
}

/* ---------- Electrical ---------- */
function loadElectrical(workbook) {
    const sheet = workbook.Sheets['Electrical'];
    if (!sheet) return;

    const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    const headerIndex = rawRows.findIndex(r => r && r.some(c => String(c).trim().toUpperCase() === 'MOTOR ID'));
    if (headerIndex === -1) return;

    const dataRows = rawRows.slice(headerIndex + 1);
    const tbody11 = document.querySelector('#table-11kv tbody');
    const tbody33 = document.querySelector('#table-33kv tbody');
    tbody11.innerHTML = '';
    tbody33.innerHTML = '';

    dataRows.forEach(row => {
        if (!row || row.every(cell => cell === undefined || cell === null || String(cell).trim() === '')) return;
        const sNo = row[0];
        const motorId = row[1] ? String(row[1]).trim() : '';
        const rating = row[2] || '-';
        const volt = row[3];
        const boardDate = row[4] || '-';
        const trialDate = row[5] || '-';

        if (motorId.includes('11 KV Drives')) { document.getElementById('badge-11kv').textContent = boardDate; return; }
        if (motorId.includes('3.3 KV Drives')) { document.getElementById('badge-33kv').textContent = boardDate; return; }

        const rowHtml = `
            <tr>
                <td>${sNo !== undefined && sNo !== null ? sNo : ''}</td>
                <td><strong>${motorId}</strong></td>
                <td>${rating}</td>
                <td>${boardDate}</td>
                <td>${trialDate}</td>
            </tr>
        `;
        if (volt === 11 || volt === '11') tbody11.innerHTML += rowHtml;
        else if (volt === 3.3 || volt === '3.3') tbody33.innerHTML += rowHtml;
    });
}

/* ---------- Milestones ---------- */
function loadMilestones(workbook){
    const sheet = workbook.Sheets['Milestone'];
    if(!sheet) return;

    const rows = XLSX.utils.sheet_to_json(sheet);
    const tbody = document.querySelector('#milestoneTable tbody');
    tbody.innerHTML='';

    rows.forEach(row => {
        const values = Object.values(row);
        const tr = document.createElement('tr');
        tr.innerHTML=`<td>${values[0]}</td><td>${values[1]}</td>`;
        tbody.appendChild(tr);
    });
}

/* ---------- Capex Progress ---------- */
function loadCapex(workbook){
    const sheet = workbook.Sheets['Capex'];
    if(!sheet) return;

    const rows = XLSX.utils.sheet_to_json(sheet);
    if(rows.length === 0) return;

    const labels = [];
    const values = [];
    const colors = ['#0078D4', '#28a745', '#ff6600', '#8e44ad'];

    rows.forEach(r => {
        const keys = Object.keys(r);
        labels.push(r[keys[0]]);
        values.push(parsePercentValue(r[keys[2]]));
    });

    if(capexChartObj) capexChartObj.destroy();

    capexChartObj = new Chart(document.getElementById('capexChart'), {
        type: 'bar',
        data:{
            labels: labels,
            datasets:[{ label: 'Progress (%)', data: values, backgroundColor: labels.map((l,i) => colors[i % colors.length]), borderRadius: 4 }]
        },
        options:{
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            plugins:{
                legend:{ display:false },
                tooltip:{ callbacks:{ label: (ctx) => ctx.parsed.x.toFixed(1) + '%' } }
            },
            scales:{
                x:{ beginAtZero:true, max:100, ticks:{ callback: (v) => v + '%' } },
                y:{ ticks:{ font:{ size:11 } } }
            }
        }
    });
}

/* ---------- Master Dashboard Load ---------- */
async function loadDashboard(){
    try {
        const workbook = await loadWorkbook(DATA_FILE_PATH);
        loadACC(workbook);
        loadManpower(workbook);
        loadMilestones(workbook);
        loadCapex(workbook);
        loadElectrical(workbook);

        document.getElementById('fileStatus').textContent = 'Data last refreshed at ' + new Date().toLocaleTimeString();
    } catch(err){
        console.error(err);
        document.getElementById('fileStatus').textContent = 'Could not load data file. Please check that "' + DATA_FILE_PATH + '" exists.';
    }
}

/* ---------- PDF Communications ---------- */
let currentCommPage = 1;
const COMM_PAGE_SIZE = 5;
let currentCommFiles = [];

function switchCommTab(folderName, tabEl){
    currentCommFolder = folderName;
    currentCommPage = 1;
    document.querySelectorAll('#commTabs .comm-tab').forEach(t => t.classList.remove('active'));
    tabEl.classList.add('active');
    loadRecentPdfs();
}

async function loadRecentPdfs(){
    const container = document.getElementById('pdfListContainer');
    const pagination = document.getElementById('commPagination');
    const folderPath = PDF_ROOT + '/' + currentCommFolder;
    const manifestPath = folderPath + '/manifest.json';

    container.innerHTML = '<div class="empty-msg">Loading recent communications...</div>';
    if(pagination) pagination.style.display = 'none';

    try {
        const response = await fetch(manifestPath + '?t=' + Date.now());
        if(!response.ok) throw new Error('Manifest not found');

        const files = await response.json();
        if(!files || files.length === 0){
            container.innerHTML = '<div class="empty-msg">No communications uploaded yet in ' + currentCommFolder + '.</div>';
            return;
        }

        currentCommFiles = files.slice().sort((a,b) => new Date(b.date) - new Date(a.date));
        const totalPages = Math.ceil(currentCommFiles.length / COMM_PAGE_SIZE) || 1;
        if (currentCommPage > totalPages) currentCommPage = totalPages;
        if (currentCommPage < 1) currentCommPage = 1;

        renderCommPage();
    } catch(err){
        console.error(err);
        container.innerHTML = '<div class="empty-msg">Could not load communications. Ensure "' + manifestPath + '" exists.</div>';
    }
}

function renderCommPage(){
    const container = document.getElementById('pdfListContainer');
    const pagination = document.getElementById('commPagination');
    const prevBtn = document.getElementById('commPrevBtn');
    const nextBtn = document.getElementById('commNextBtn');
    const pageInfo = document.getElementById('commPageInfo');
    const folderPath = PDF_ROOT + '/' + currentCommFolder;

    const totalFiles = currentCommFiles.length;
    const totalPages = Math.ceil(totalFiles / COMM_PAGE_SIZE) || 1;
    const startIndex = (currentCommPage - 1) * COMM_PAGE_SIZE;
    const pageFiles = currentCommFiles.slice(startIndex, startIndex + COMM_PAGE_SIZE);

    const ul = document.createElement('ul');
    ul.className = 'pdf-list';

    pageFiles.forEach(item => {
        const li = document.createElement('li');
        li.className = 'pdf-item';
        const dateStr = item.date ? new Date(item.date).toLocaleDateString() : '';
        li.innerHTML = `
            <a class="pdf-name" href="${folderPath}/${item.file}" target="_blank">&#128196; ${item.file}</a>
            <span class="pdf-date">${dateStr}</span>
        `;
        ul.appendChild(li);
    });

    container.innerHTML = '';
    container.appendChild(ul);
    pagination.style.display = totalFiles > COMM_PAGE_SIZE ? 'flex' : 'none';
    pageInfo.textContent = `Page ${currentCommPage} of ${totalPages} (${totalFiles} total)`;
    prevBtn.disabled = (currentCommPage === 1);
    nextBtn.disabled = (currentCommPage === totalPages);
}

function changeCommPage(delta){
    currentCommPage += delta;
    renderCommPage();
}

/* ---------- Tasks & Calendar ---------- */
let taskMap = {};
let calYear = new Date().getFullYear();
let calMonth = new Date().getMonth();

function classifyStatus(statusRaw){
    const status = (statusRaw || '').toString().trim().toLowerCase();
    if(status.startsWith('completed')) return 'completed';
    if(status.startsWith('delayed')) return 'delayed';
    return 'pending';
}

async function loadTasks(){
    try {
        const workbook = await loadWorkbook(TASK_FILE_PATH);
        const sheet = workbook.Sheets['Tasks'];
        if(!sheet) return;

        const allRows = XLSX.utils.sheet_to_json(sheet, {header:1, raw:true});
        if(allRows.length < 2) return;

        taskMap = {};
        rawTaskRows = [];

        allRows.slice(1).forEach(row => {
            if(!row || row.length === 0) return;
            const rawDate = row[0];
            const activity = row[1];
            const statusRaw = (row[2] !== undefined && row[2] !== null) ? row[2] : '';
            const department = (row[3] !== undefined && row[3] !== null) ? row[3].toString().trim() : '';

            if(rawDate === undefined || rawDate === null || rawDate === '') return;
            const dateObj = parseAnyDate(rawDate);
            if(isNaN(dateObj)) return;

            const statusClass = classifyStatus(statusRaw);

            rawTaskRows.push({
                date: dateObj,
                activity: activity,
                status: statusRaw || '',
                statusClass: statusClass,
                department: department || 'Unassigned'
            });

            const key = dateObj.getFullYear() + '-' + (dateObj.getMonth()+1) + '-' + dateObj.getDate();
            if(!taskMap[key]) taskMap[key] = [];
            taskMap[key].push({ activity: activity, status: statusRaw || '', statusClass: statusClass });
        });

	// <-- Initializes the slider range across all tasks & opens This Week

	tasksLoaded = true;
        renderCalendar();
        setupDeptWiseCommitments();
        initTimelineBounds(); // <-- Initializes slider range and displays current week

        if(deptChartOpen) renderDeptChart();
    } catch(err){
        console.error(err);
        document.getElementById('calendarGrid').innerHTML = '<div class="empty-msg">Could not load "' + TASK_FILE_PATH + '"</div>';
    }
}

function changeMonth(delta){
    calMonth += delta;
    if(calMonth < 0){ calMonth = 11; calYear -= 1; }
    else if(calMonth > 11){ calMonth = 0; calYear += 1; }
    renderCalendar();
    if(deptChartOpen) renderDeptChart();
}

function toggleDeptChart(){
    deptChartOpen = !deptChartOpen;
    const wrap = document.getElementById('deptChartWrap');
    const btn = document.getElementById('deptToggleBtn');

    if(deptChartOpen){
        wrap.classList.add('open');
        btn.classList.add('active');
        setTimeout(() => renderDeptChart(), 100);
    } else {
        wrap.classList.remove('open');
        btn.classList.remove('active');
    }
}

function renderDeptChart(){
    const canvasEl = document.getElementById('deptChart');
    const captionEl = document.getElementById('deptChartCaption');
    const debugEl = document.getElementById('deptDebug');
    if(!canvasEl) return;
    debugEl.textContent = '';

    if(!tasksLoaded){ captionEl.textContent = 'Loading task data...'; return; }

    const monthNames = ['January','February','March','April','May','June','July','August','September','October','November','December'];
    const monthTasks = rawTaskRows.filter(t => t.date.getFullYear() === calYear && t.date.getMonth() === calMonth);
    const deptStats = {};

    monthTasks.forEach(t => {
        if(!deptStats[t.department]) deptStats[t.department] = { total: 0, completed: 0 };
        deptStats[t.department].total += 1;
        if(t.statusClass === 'completed') deptStats[t.department].completed += 1;
    });

    const departments = Object.keys(deptStats);
    captionEl.textContent = departments.length > 0
        ? 'Targets vs Completed by Department - ' + monthNames[calMonth] + ' ' + calYear
        : 'No tasks with department data found for ' + monthNames[calMonth] + ' ' + calYear;

    if(deptChartObj){ deptChartObj.destroy(); deptChartObj = null; }
    if(departments.length === 0){
        debugEl.textContent = 'Total task rows loaded: ' + rawTaskRows.length + '. None matched ' + monthNames[calMonth] + ' ' + calYear + '.';
        return;
    }

    const ctx = canvasEl.getContext('2d');
    deptChartObj = new Chart(ctx, {
        type: 'bar',
        data:{
            labels: departments,
            datasets:[
                { label: 'Targets', data: departments.map(d => deptStats[d].total), backgroundColor: '#0078D4', borderRadius: 4 },
                { label: 'Completed', data: departments.map(d => deptStats[d].completed), backgroundColor: '#28a745', borderRadius: 4 }
            ]
        },
        options:{
            responsive:true,
            maintainAspectRatio:false,
            plugins:{ legend:{ display:true, position:'top' } },
            scales:{ y:{ beginAtZero:true, ticks:{ stepSize:1 } } }
        }
    });
}

function getDayOverallStatus(tasks, cellDate, today){
    if(tasks.some(t => t.statusClass === 'delayed')) return 'delayed';
    if(tasks.every(t => t.statusClass === 'completed')) return 'completed';
    if(cellDate > today) return 'upcoming';
    return 'pending';
}

function renderCalendar(){
    const grid = document.getElementById('calendarGrid');
    const monthNames = ['January','February','March','April','May','June','July','August','September','October','November','December'];
    document.getElementById('calMonthLabel').textContent = monthNames[calMonth] + ' ' + calYear;
    grid.innerHTML = '';

    ['Su','Mo','Tu','We','Th','Fr','Sa'].forEach(d => {
        const el = document.createElement('div');
        el.className = 'cal-day-name';
        el.textContent = d;
        grid.appendChild(el);
    });

    const firstDay = new Date(calYear, calMonth, 1).getDay();
    const daysInMonth = new Date(calYear, calMonth + 1, 0).getDate();
    const today = new Date();
    const todayMidnight = new Date(today.getFullYear(), today.getMonth(), today.getDate());

    for(let i = 0; i < firstDay; i++){
        const empty = document.createElement('div');
        empty.className = 'cal-cell empty';
        grid.appendChild(empty);
    }

    for(let d = 1; d <= daysInMonth; d++){
        const key = calYear + '-' + (calMonth+1) + '-' + d;
        const cell = document.createElement('div');
        const tasks = taskMap[key];
        const cellDate = new Date(calYear, calMonth, d);

        if(tasks && tasks.length > 0){
            const overallStatus = getDayOverallStatus(tasks, cellDate, todayMidnight);
            cell.className = 'cal-cell ' + overallStatus;
            cell.addEventListener('mouseenter', (e) => showCalTooltip(e, tasks, overallStatus));
            cell.addEventListener('mousemove', (e) => moveCalTooltip(e));
            cell.addEventListener('mouseleave', () => hideCalTooltip());
        } else {
            cell.className = 'cal-cell normal';
        }
        cell.textContent = d;
        grid.appendChild(cell);
    }
}

function showCalTooltip(e, tasks, overallStatus){
    const tooltip = document.getElementById('calTooltip');
    const itemsHtml = tasks.map(t => {
        let tagClass = t.statusClass;
        let tagText = t.status;
        if(t.statusClass === 'pending' && overallStatus === 'upcoming'){ tagClass = 'upcoming'; tagText = t.status || 'Upcoming'; }
        else if(!tagText){ tagText = 'Pending'; }
        return '<div style="margin-bottom:6px">&#8226; ' + t.activity + '<br><span class="status-tag ' + tagClass + '">' + tagText + '</span></div>';
    }).join('');
    tooltip.innerHTML = itemsHtml;
    tooltip.style.display = 'block';
    moveCalTooltip(e);
}

function moveCalTooltip(e){
    const tooltip = document.getElementById('calTooltip');
    tooltip.style.left = (e.clientX + 15) + 'px';
    tooltip.style.top = (e.clientY + 15) + 'px';
}

function hideCalTooltip(){
    document.getElementById('calTooltip').style.display = 'none';
}

/* ---------- Department & Month Task Filter ---------- */
function setupDeptWiseCommitments() {
    if (!rawTaskRows || rawTaskRows.length === 0) return;

    const depts = [...new Set(rawTaskRows.map(t => t.department).filter(d => d && d !== 'Unassigned'))];
    const preferredOrder = ['ME', 'CCD', 'EED', 'C&T', 'P&S'];
    depts.sort((a, b) => {
        let ia = preferredOrder.indexOf(a), ib = preferredOrder.indexOf(b);
        if (ia !== -1 && ib !== -1) return ia - ib;
        if (ia !== -1) return -1;
        if (ib !== -1) return 1;
        return a.localeCompare(b);
    });

    if (!selectedDept || !depts.includes(selectedDept)) selectedDept = depts[0] || '';

    renderDeptTabs(depts);
    renderMonthTabs();
    renderFilteredDeptTasks();
}

function renderDeptTabs(depts) {
    const container = document.getElementById('deptTabsContainer');
    if (!container) return;
    container.innerHTML = depts.map(dept => `
        <button class="dept-main-tab ${dept === selectedDept ? 'active' : ''}" onclick="selectDept('${dept}')">
            ${dept}
        </button>
    `).join('');
}

function selectDept(dept) {
    selectedDept = dept;
    selectedMonth = '';
    const depts = [...new Set(rawTaskRows.map(t => t.department).filter(d => d && d !== 'Unassigned'))];
    renderDeptTabs(depts);
    renderMonthTabs();
    renderFilteredDeptTasks();
}

function renderMonthTabs() {
    const container = document.getElementById('monthTabsContainer');
    if (!container) return;

    const deptTasks = rawTaskRows.filter(t => t.department === selectedDept);
    const monthMap = new Map();
    deptTasks.forEach(t => {
        if (!isNaN(t.date)) {
            const key = `${t.date.getFullYear()}-${String(t.date.getMonth() + 1).padStart(2, '0')}`;
            const label = t.date.toLocaleDateString('en-GB', { month: 'short', year: 'numeric' });
            monthMap.set(key, label);
        }
    });

    const sortedKeys = Array.from(monthMap.keys()).sort();
    if (sortedKeys.length === 0) {
        container.innerHTML = '<span style="color:#888; font-size:12px;">No tasks available for this department.</span>';
        selectedMonth = '';
        return;
    }

    if (!selectedMonth || !monthMap.has(selectedMonth)) selectedMonth = sortedKeys[0];

    container.innerHTML = sortedKeys.map(key => `
        <button class="month-tab ${key === selectedMonth ? 'active' : ''}" onclick="selectMonth('${key}')">
            ${monthMap.get(key)}
        </button>
    `).join('');
}

function selectMonth(monthKey) {
    selectedMonth = monthKey;
    renderMonthTabs();
    renderFilteredDeptTasks();
}

function renderFilteredDeptTasks() {
    const tbody = document.querySelector('#commitmentsTable tbody');
    if (!tbody) return;

    const todayMidnight = new Date();
    todayMidnight.setHours(0, 0, 0, 0);

    const filtered = rawTaskRows.filter(t => {
        if (t.department !== selectedDept) return false;
        if (!selectedMonth) return true;
        const key = `${t.date.getFullYear()}-${String(t.date.getMonth() + 1).padStart(2, '0')}`;
        return key === selectedMonth;
    });

    filtered.sort((a, b) => a.date - b.date);

    if (filtered.length === 0) {
        tbody.innerHTML = `<tr><td colspan="3" class="empty-msg">No tasks found for ${selectedDept} in this period.</td></tr>`;
        return;
    }

    tbody.innerHTML = filtered.map(t => {
        const dateStr = t.date.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
        let tagClass = t.statusClass;
        let tagText = t.status || 'Pending';
        if (tagClass === 'pending' && t.date > todayMidnight) {
            tagClass = 'upcoming';
            tagText = t.status || 'Upcoming';
        }
        return `
            <tr>
                <td><strong>${dateStr}</strong></td>
                <td>${t.activity}</td>
                <td><span class="status-tag ${tagClass}">${tagText}</span></td>
            </tr>
        `;
    }).join('');
}


/* ---------- Daily Major Updates ---------- */
async function loadDailyUpdates(){
    try {
        const workbook = await loadWorkbook(UPDATES_FILE_PATH);
        const sheet = workbook.Sheets['Updates'];
        if(!sheet) return;

        const rows = XLSX.utils.sheet_to_json(sheet);
        const latest = { 'ACC': null, 'Boiler': null, 'PCP Joint': null };

        rows.forEach(r => {
            const keys = Object.keys(r);
            let rawDate = r[keys[0]];
            let category = (r[keys[1]] || '').toString().trim();
            let update = r[keys[2]];
            const dateObj = parseAnyDate(rawDate);
            if(isNaN(dateObj) || !(category in latest)) return;
            if(!latest[category] || dateObj > latest[category].dateObj){
                latest[category] = { dateObj: dateObj, text: update };
            }
        });

        const mapping = { 'ACC':'acc', 'Boiler':'boiler', 'PCP Joint':'pcp' };
        Object.keys(mapping).forEach(cat => {
            const id = mapping[cat];
            const entry = latest[cat];
            if(entry){
                document.getElementById('date-' + id).textContent = entry.dateObj.toLocaleDateString();
                document.getElementById('text-' + id).textContent = entry.text;
            } else {
                document.getElementById('date-' + id).textContent = '';
                document.getElementById('text-' + id).textContent = 'No update yet.';
            }
        });
    } catch(err){
        console.error(err);
        ['acc','boiler','pcp'].forEach(id => {
            document.getElementById('text-' + id).textContent = 'Could not load "' + UPDATES_FILE_PATH + '"';
        });
    }
}

/* ---------- Critical Areas ---------- */
async function loadCriticalAreas() {
  try {
    const workbook = await loadWorkbook(UPDATES_FILE_PATH);
    const sheet = workbook.Sheets["Critical areas"];
    if (!sheet) return;

    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    if (rows.length < 2) return;

    const colNames = rows[0].slice(1);
    let targets = [null, null, null];
    const dailyRows = [];

    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row || row[0] === undefined || row[0] === null) continue;
      const firstCell = row[0].toString().trim().toLowerCase();

      if (firstCell.includes("target")) {
        targets = [row[1], row[2], row[3]];
        continue;
      }
      const dateObj = parseAnyDate(row[0]);
      if (!isNaN(dateObj)) dailyRows.push({ date: dateObj, values: [row[1], row[2], row[3]] });
    }

    dailyRows.sort((a, b) => a.date - b.date);
    const last7 = dailyRows.slice(-7);
    const labels = last7.map((r) => r.date.getDate() + "/" + (r.date.getMonth() + 1));

    renderCriticalChart("criticalChartAcc", labels, last7.map((r) => r.values[0] || 0), targets[0], "#0078D4", "target-badge-acc", colNames[0] || "ACC");
    renderCriticalChart("criticalChartBoiler", labels, last7.map((r) => r.values[1] || 0), targets[1], "#ff6600", "target-badge-boiler", colNames[1] || "Boiler");
    renderCriticalChart("criticalChartPcp", labels, last7.map((r) => r.values[2] || 0), targets[2], "#28a745", "target-badge-pcp", colNames[2] || "PCP");
  } catch (err) {
    console.error(err);
  }
}

const barValueLabelPlugin = {
  id: "barValueLabelPlugin",
  afterDatasetsDraw(chart) {
    const { ctx } = chart;
    chart.data.datasets.forEach((dataset, datasetIndex) => {
      if (dataset.type !== "bar") return;
      const meta = chart.getDatasetMeta(datasetIndex);
      meta.data.forEach((bar, index) => {
        const value = dataset.data[index];
        if (value === null || value === undefined) return;
        ctx.save();
        ctx.fillStyle = "#333";
        ctx.font = "bold 11px Segoe UI, Arial, sans-serif";
        ctx.textAlign = "center";
        ctx.textBaseline = "bottom";
        ctx.fillText(value, bar.x, bar.y - 4);
        ctx.restore();
      });
    });
  },
};
Chart.register(barValueLabelPlugin);

function renderCriticalChart(canvasId, labels, values, target, barColor, badgeId, seriesLabel) {
  const badge = document.getElementById(badgeId);
  if (badge) badge.textContent = "Target: " + (target !== null ? target : "N/A");

  if (canvasId === "criticalChartAcc" && criticalChartAcc) criticalChartAcc.destroy();
  else if (canvasId === "criticalChartBoiler" && criticalChartBoiler) criticalChartBoiler.destroy();
  else if (canvasId === "criticalChartPcp" && criticalChartPcp) criticalChartPcp.destroy();

  const targetLineData = target !== null ? labels.map(() => target) : [];

  const newChart = new Chart(document.getElementById(canvasId), {
    data: {
      labels: labels,
      datasets: [
        { type: "bar", label: seriesLabel + " - Daily", data: values, backgroundColor: barColor, borderRadius: 4, order: 2 },
        { type: "line", label: "Target", data: targetLineData, borderColor: "#d93025", borderWidth: 2, borderDash: [6, 4], pointRadius: 0, fill: false, order: 1 }
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      layout: { padding: { top: 20 } },
      plugins: { legend: { display: true, position: "top", labels: { font: { size: 10 }, boxWidth: 12 } } },
      scales: { y: { beginAtZero: true, ticks: { font: { size: 10 } } }, x: { ticks: { font: { size: 10 } } } }
    },
  });

  if (canvasId === "criticalChartAcc") criticalChartAcc = newChart;
  else if (canvasId === "criticalChartBoiler") criticalChartBoiler = newChart;
  else if (canvasId === "criticalChartPcp") criticalChartPcp = newChart;
}

function toggleCollapse(headerEl, bodyId){
    const body = document.getElementById(bodyId);
    const isOpen = body.classList.contains('open');

    if(isOpen){
        body.classList.remove('open');
        headerEl.classList.remove('open');
    } else {
        body.classList.add('open');
        headerEl.classList.add('open');
        setTimeout(() => {
            if(bodyId === 'critical-acc' && criticalChartAcc) criticalChartAcc.resize();
            if(bodyId === 'critical-boiler' && criticalChartBoiler) criticalChartBoiler.resize();
            if(bodyId === 'critical-pcp' && criticalChartPcp) criticalChartPcp.resize();
        }, 310);
    }
}

/* ---------- KPI Details ---------- */
let tonnageData = { 'Unit 1':[], 'Unit 2':[], 'Unit 3':[] };
let targetsData = [];

function formatPercent(pct){
    if(pct === null || pct === undefined || pct === '') return '';
    let num = pct;
    if(typeof num === 'string') num = parseFloat(num.replace('%',''));
    if(isNaN(num)) return pct;
    if(num <= 1) num = num * 100;
    return num.toFixed(1) + '%';
}

async function loadKpiDetails(){
    try {
        const workbook = await loadWorkbook(KPI_FILE_PATH);
        const tonnageSheet = workbook.Sheets['Tonnage'];
        if(tonnageSheet){
            const rows = XLSX.utils.sheet_to_json(tonnageSheet);
            tonnageData = { 'Unit 1':[], 'Unit 2':[], 'Unit 3':[] };
            rows.forEach(r => {
                const keys = Object.keys(r);
                const unit = (r[keys[0]] || '').toString().trim();
                if(tonnageData[unit]){
                    tonnageData[unit].push({ component: r[keys[1]], completed: r[keys[2]], total: r[keys[3]], pct: formatPercent(r[keys[4]]) });
                }
            });
        }

        const targetSheet = workbook.Sheets['SepTargets'];
        if(targetSheet){
            const rows = XLSX.utils.sheet_to_json(targetSheet);
            targetsData = rows.map(r => r[Object.keys(r)[1]]);
            renderTargets();
        }
    } catch(err){
        console.error(err);
        document.getElementById('targetList').innerHTML = '<li class="target-item">Could not load "' + KPI_FILE_PATH + '"</li>';
    }
}

function renderTargets(){
    const list = document.getElementById('targetList');
    list.innerHTML = '';
    targetsData.forEach((target, idx) => {
        const li = document.createElement('li');
        li.className = 'target-item';
        li.innerHTML = `<span class="target-num">${idx+1}</span><span>${target}</span>`;
        list.appendChild(li);
    });
}

function setupKpiTooltip(kpiId, unitName){
    const card = document.getElementById(kpiId);
    card.addEventListener('mouseenter', (e) => showKpiTooltip(e, unitName));
    card.addEventListener('mousemove', (e) => moveKpiTooltip(e));
    card.addEventListener('mouseleave', () => hideKpiTooltip());
}

function showKpiTooltip(e, unitName){
    const tooltip = document.getElementById('kpiTooltip');
    const details = tonnageData[unitName] || [];
    if(details.length === 0){
        tooltip.innerHTML = 'No tonnage details available for ' + unitName;
    } else {
        let rowsHtml = details.map(d => `<tr><td>${d.component}</td><td>${d.completed}</td><td>${d.total}</td><td>${d.pct}</td></tr>`).join('');
        tooltip.innerHTML = `<strong>${unitName} - Tonnage Details</strong><table><tr><th>Component</th><th>Done</th><th>Total</th><th>%</th></tr>${rowsHtml}</table>`;
    }
    tooltip.style.display = 'block';
    moveKpiTooltip(e);
}

function moveKpiTooltip(e){
    const tooltip = document.getElementById('kpiTooltip');
    tooltip.style.left = (e.clientX + 15) + 'px';
    tooltip.style.top = (e.clientY + 15) + 'px';
}

function hideKpiTooltip(){
    document.getElementById('kpiTooltip').style.display = 'none';
}

setupKpiTooltip('kpi-unit1', 'Unit 1');
setupKpiTooltip('kpi-unit2', 'Unit 2');
setupKpiTooltip('kpi-unit3', 'Unit 3');

/* ---------- Slides Carousel ---------- */
let pdfDoc = null;
let currentSlide = 1;
let totalSlides = 0;
let autoplayTimer = null;
const AUTOPLAY_INTERVAL = 6000;

if(window['pdfjsLib']){
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
}

async function loadSlides(){
    const caption = document.getElementById('carouselCaption');
    try {
        const loadingTask = pdfjsLib.getDocument(encodeURI(SLIDES_PDF_PATH) + '?t=' + Date.now());
        pdfDoc = await loadingTask.promise;
        totalSlides = pdfDoc.numPages;
        currentSlide = 1;

        buildDots();
        await renderSlide(currentSlide);
        startAutoplay();
    } catch(err){
        console.error(err);
        caption.textContent = 'Could not load slides. Ensure "' + SLIDES_PDF_PATH + '" exists.';
    }
}

function buildDots(){
    const dotsContainer = document.getElementById('carouselDots');
    dotsContainer.innerHTML = '';
    for(let i = 1; i <= totalSlides; i++){
        const dot = document.createElement('span');
        dot.className = 'carousel-dot' + (i === currentSlide ? ' active' : '');
        dot.onclick = () => { currentSlide = i; renderSlide(currentSlide); resetAutoplay(); };
        dotsContainer.appendChild(dot);
    }
}

function updateDots(){
    document.querySelectorAll('.carousel-dot').forEach((dot, idx) => dot.classList.toggle('active', (idx + 1) === currentSlide));
}

async function renderSlide(pageNum){
    if(!pdfDoc) return;
    const page = await pdfDoc.getPage(pageNum);
    const canvas = document.getElementById('slideCanvas');
    const context = canvas.getContext('2d');
    const viewport = page.getViewport({ scale: 1.4 });
    canvas.width = viewport.width;
    canvas.height = viewport.height;
    await page.render({ canvasContext: context, viewport: viewport }).promise;
    document.getElementById('carouselCaption').textContent = 'Slide ' + pageNum + ' of ' + totalSlides;
    updateDots();
}

function nextSlide(){ currentSlide = (currentSlide % totalSlides) + 1; renderSlide(currentSlide); resetAutoplay(); }
function prevSlide(){ currentSlide = (currentSlide - 2 + totalSlides) % totalSlides + 1; renderSlide(currentSlide); resetAutoplay(); }
function startAutoplay(){
    if(autoplayTimer) clearInterval(autoplayTimer);
    autoplayTimer = setInterval(() => { currentSlide = (currentSlide % totalSlides) + 1; renderSlide(currentSlide); }, AUTOPLAY_INTERVAL);
}
function resetAutoplay(){ startAutoplay(); }

document.getElementById('slidesSection').addEventListener('mouseenter', () => { if(autoplayTimer) clearInterval(autoplayTimer); });
document.getElementById('slidesSection').addEventListener('mouseleave', () => startAutoplay());

/* ---------- Graphic Timeline with Slider & Navigation ---------- */

let timelineMinDate = null;
let timelineMaxDate = null;
let timelineCurrentStartDate = null;

function initTimelineBounds() {
    if (!rawTaskRows || rawTaskRows.length === 0) return;

    // Filter valid dates from all tasks
    const validDates = rawTaskRows
        .map(t => t.date)
        .filter(d => d instanceof Date && !isNaN(d.getTime()))
        .sort((a, b) => a - b);

    if (validDates.length === 0) return;

    // Start from the earliest task date
    timelineMinDate = new Date(validDates[0].getFullYear(), validDates[0].getMonth(), validDates[0].getDate());
    
    // Cap the upper bound strictly to 31 March 2027
    timelineMaxDate = new Date(2027, 2, 31); // Month index 2 = March

    // Update Slider UI attributes
    const slider = document.getElementById('timelineDateSlider');
    const minLabel = document.getElementById('sliderMinLabel');
    const maxLabel = document.getElementById('sliderMaxLabel');

    const totalDays = Math.max(7, Math.round((timelineMaxDate - timelineMinDate) / (1000 * 60 * 60 * 24)));

    if (slider) {
        slider.min = 0;
        slider.max = totalDays;
    }
    if (minLabel) minLabel.textContent = timelineMinDate.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: '2-digit' });
    if (maxLabel) maxLabel.textContent = "Mar '27"; // Labels the end of the slider cleanly

    // Set Default: This Week (starting Monday)
    resetTimelineToThisWeek();
}

function resetTimelineToThisWeek() {
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    // Find Monday of the current week
    const dayOfWeek = today.getDay(); // 0 is Sunday, 1 is Monday
    const diff = today.getDate() - dayOfWeek + (dayOfWeek === 0 ? -6 : 1);
    const monday = new Date(today);
    monday.setDate(diff);

    setTimelineStartDate(monday);
}

function setTimelineStartDate(startDate) {
    if (!timelineMinDate || !timelineMaxDate) return;

    // Cap the maximum start date so the 7-day view finishes by 31 Mar 2027
    const maxStartDate = new Date(timelineMaxDate);
    maxStartDate.setDate(maxStartDate.getDate() - 6);

    if (startDate < timelineMinDate) startDate = new Date(timelineMinDate);
    if (startDate > maxStartDate) startDate = new Date(maxStartDate);

    timelineCurrentStartDate = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());

    // Sync Slider knob position
    const slider = document.getElementById('timelineDateSlider');
    if (slider) {
        const offsetDays = Math.round((timelineCurrentStartDate - timelineMinDate) / (1000 * 60 * 60 * 24));
        slider.value = Math.max(0, Math.min(parseInt(slider.max, 10), offsetDays));
    }

    renderTimelineView();
}

function onTimelineSliderChange(dayOffset) {
    if (!timelineMinDate) return;
    const newDate = new Date(timelineMinDate);
    newDate.setDate(timelineMinDate.getDate() + parseInt(dayOffset, 10));
    timelineCurrentStartDate = newDate;
    renderTimelineView();
}

function shiftTimelineDays(deltaDays) {
    if (!timelineCurrentStartDate) {
        resetTimelineToThisWeek();
    }
    const newDate = new Date(timelineCurrentStartDate);
    newDate.setDate(newDate.getDate() + deltaDays);
    setTimelineStartDate(newDate);
}

function renderTimelineView() {
    const track = document.getElementById('timelineTrack');
    const rangeLabel = document.getElementById('timelineWeekRange');
    if (!track || !timelineCurrentStartDate) return;

    // Build the 7 days starting from current date
    const days = [];
    for (let i = 0; i < 7; i++) {
        const d = new Date(timelineCurrentStartDate);
        d.setDate(timelineCurrentStartDate.getDate() + i);
        days.push(d);
    }

    if (rangeLabel) {
        const dStart = days[0].toLocaleDateString('en-GB', { day: '2-digit', month: 'short' });
        const dEnd = days[6].toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
        rangeLabel.textContent = `${dStart} – ${dEnd}`;
    }

    const dayFormatter = new Intl.DateTimeFormat('en-GB', { weekday: 'short' });
    const dateFormatter = new Intl.DateTimeFormat('en-GB', { day: '2-digit', month: 'short' });

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    let html = '';

    days.forEach(dateObj => {
        const key = `${dateObj.getFullYear()}-${dateObj.getMonth() + 1}-${dateObj.getDate()}`;
        const dayTasks = (taskMap && taskMap[key]) ? taskMap[key] : [];
        const hasTargets = dayTasks.length > 0;
        const isToday = (dateObj.getTime() === today.getTime());

        let targetsHtml = '';
        if (hasTargets) {
            targetsHtml = dayTasks.map(t => {
                const statusClass = t.statusClass || 'pending';
                const matched = rawTaskRows.find(r => r.date.getTime() === dateObj.getTime() && r.activity === t.activity);
                const dept = matched ? matched.department : '';

                return `
                    <div class="timeline-target-item ${statusClass}">
                        ${dept ? `<span class="timeline-dept-tag">${dept}</span>` : ''}
                        <div style="font-weight:600; color:#1e293b; margin-bottom:3px;">${t.activity}</div>
                        <span class="status-tag ${statusClass}">${t.status || 'Pending'}</span>
                    </div>
                `;
            }).join('');
        } else {
            targetsHtml = '<div class="timeline-empty">— No targets —</div>';
        }

        html += `
            <div class="timeline-col ${hasTargets ? 'has-targets' : ''} ${isToday ? 'today' : ''}">
                <div class="timeline-node">${dateObj.getDate()}</div>
                <div class="timeline-date-label">${dateFormatter.format(dateObj)}</div>
                <div class="timeline-day-sub">${dayFormatter.format(dateObj)}${isToday ? ' (Today)' : ''}</div>
                <div class="timeline-stem"></div>
                <div class="timeline-targets-wrap">
                    ${targetsHtml}
                </div>
            </div>
        `;
    });

    track.innerHTML = html;
}


/* ---------- Initialize Dashboard ---------- */
loadDashboard();
loadRecentPdfs();
loadTasks();
loadDailyUpdates();
loadKpiDetails();
loadSlides();
loadCriticalAreas();
loadCivilFoundations();

setInterval(loadDashboard, 60000);
setInterval(loadRecentPdfs, 60000);
setInterval(loadTasks, 60000);
setInterval(loadDailyUpdates, 60000);
setInterval(loadKpiDetails, 60000);
setInterval(loadCriticalAreas, 60000);