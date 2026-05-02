const STORAGE_KEYS = {
  companyListings: "derby-dib-company-listings-v3",
  followUps: "derby-dib-followups-v3",
  session: "derby-dib-session-v1",
};

const ACTIVE_STAGE_EXCLUSIONS = new Set(["BOOKED", "BOOKED -", "CANCEL"]);
const STAGE_SEQUENCE = [
  "Refer",
  "DI Refer",
  "Detail Data Entry",
  "CIF Account Linkage",
  "Detail Data Entry Quality Check",
  "Post Sanc Doc",
];

const MIS_COLUMN_NAMES = {
  agent: "Employee Name",
  stage: "Current Stage",
  amount: "Finance Amount",
  date: "Date",
};

const ACTIVE_AGENTS = [
  "OMAR ALI",
  "MOHAMED YOUSSEF",
  "MOHAB HESHAM",
  "EL SAYED",
  "ABDALLA HASSAN",
  "LAMIAA YOUSSEF",
  "MONA SOLIMAN",
  "SARRA MOUSSALI",
  "HANI WALID",
  "DONIA AHMED",
  "AHMED HAMDY",
  "AHMED MOHAMED",
];

const ACTIVE_AGENT_DISPLAY_NAMES = {
  "OMAR ALI": "OMAR ALI",
  "MOHAMED YOUSSEF": "MOHAMED YOUSSEF FAROUQ",
  "MOHAB HESHAM": "MOHAB HESHAM",
  "EL SAYED": "EL SAYED",
  "ABDALLA HASSAN": "ABDALLA HASSAN",
  "LAMIAA YOUSSEF": "LAMIAA YOUSSEF",
  "MONA SOLIMAN": "MONA SOLIMAN",
  "SARRA MOUSSALI": "SARRA MOUSSALI",
  "HANI WALID": "HANI WALID",
  "DONIA AHMED": "DONIA AHMED",
  "AHMED HAMDY": "AHMED HAMDY",
  "AHMED MOHAMED": "AHMED MOHAMED",
};

const DEFAULT_MONTHLY_TARGET = 500000;
const SALES_ALLOWED_PANELS = new Set([
  "dashboard",
  "dbr-dsr",
  "emi",
  "proposal",
  "company-listing",
  "follow-up",
]);

const state = {
  misRows: buildSampleMisRows(),
  activeStaff: buildSampleStaffRows(),
  sourceName: "starter sample data",
  importedAt: "",
  selectedMonth: "",
  search: "",
  detectedAgentColumn: MIS_COLUMN_NAMES.agent,
  session: readStorage(STORAGE_KEYS.session, null),
  managerReportSnapshot: null,
  companyListings: readStorage(STORAGE_KEYS.companyListings, [
    { id: createId(), company: "LISTED A", status: "Listed", notes: "Imported from active MIS trends" },
    { id: createId(), company: "NTML", status: "Under Review", notes: "Track listing and exceptions" },
  ]),
  followUps: readStorage(STORAGE_KEYS.followUps, [
    { id: createId(), appId: "PF-1005", customer: "Omar Yousif", stage: "CIF Account Linkage", notes: "Follow up for CIF completion" },
    { id: createId(), appId: "PF-1007", customer: "Layla Kareem", stage: "Post Sanc Doc", notes: "Collect sanction documents" },
  ]),
};

const el = {};
let lastGeneratedPdfDoc = null;

document.addEventListener("DOMContentLoaded", () => {
  captureElements();
  bindEvents();
  state.selectedMonth = state.selectedMonth || "ALL";
  hydrateSession();
  if (el.emiStartDate) {
    el.emiStartDate.value = toDateInputValue(new Date());
  }
  render();
});

function captureElements() {
  const ids = [
    "authScreen",
    "appShell",
    "loginForm",
    "loginError",
    "logoutButton",
    "exportPdfButton",
    "exportExcelButton",
    "dashboardExportStatus",
    "userAvatarText",
    "userDisplayName",
    "userDisplayRole",
    "sourceMeta",
    "activeStaffBadge",
    "activeCasesBadge",
    "managerDashboardContent",
    "salesDashboardContent",
    "salesAgentBadge",
    "salesCasesBadge",
    "portalSidebar",
    "sidebarOverlay",
    "sidebarToggle",
    "sidebarClose",
    "monthFilter",
    "searchInput",
    "uploadDropzone",
    "misUpload",
    "uploadStatus",
    "mappingMeta",
    "kpiGrid",
    "agentTableBody",
    "leaderboardBody",
    "pipelineKpiGrid",
    "pipelineTableBody",
    "stageList",
    "salesKpiGrid",
    "salesTargetGrid",
    "salesUrgentBody",
    "salesPipelineBody",
    "salesLeaderboardBody",
    "previewBody",
    "submissionMiniStats",
    "submissionTableBody",
    "dsrForm",
    "dsrResult",
    "emiForm",
    "emiStartDate",
    "emiSummary",
    "emiTableBody",
    "proposalForm",
    "proposalPreview",
    "generatePdfButton",
    "downloadPdfButton",
    "proposalStatus",
    "companyForm",
    "companyList",
    "followUpForm",
    "followUpList",
  ];

  ids.forEach((id) => {
    el[id] = document.getElementById(id);
  });

  el.navButtons = [...document.querySelectorAll(".nav-button")];
  el.panels = [...document.querySelectorAll(".panel")];
}

function bindEvents() {
  el.navButtons.forEach((button) => {
    button.addEventListener("click", () => {
      activatePanel(button.dataset.panelTarget);
      closeSidebar();
    });
  });

  el.sidebarToggle?.addEventListener("click", openSidebar);
  el.sidebarClose?.addEventListener("click", closeSidebar);
  el.sidebarOverlay?.addEventListener("click", closeSidebar);
  window.addEventListener("resize", handleViewportChange);
  window.addEventListener("keydown", (event) => {
    if (event.key === "Escape") {
      closeSidebar();
    }
  });

  el.monthFilter.addEventListener("change", (event) => {
    state.selectedMonth = event.target.value;
    render();
  });

  el.searchInput.addEventListener("input", (event) => {
    state.search = event.target.value.trim();
    render();
  });

  el.loginForm?.addEventListener("submit", handleLoginSubmit);
  el.logoutButton?.addEventListener("click", handleLogout);
  el.exportPdfButton?.addEventListener("click", exportManagerPdfReport);
  el.exportExcelButton?.addEventListener("click", exportManagerExcelReport);
  el.misUpload.addEventListener("change", handleMisUpload);
  bindUploadDropzone();
  el.dsrForm.addEventListener("input", renderDsrCalculator);
  el.emiForm.addEventListener("input", renderEmiCalculator);
  el.proposalForm.addEventListener("input", renderProposalPreview);
  el.generatePdfButton.addEventListener("click", () => {
    try {
      lastGeneratedPdfDoc = buildProposalPdf();
      el.proposalStatus.textContent = "PDF prepared successfully.";
    } catch (error) {
      console.error(error);
      el.proposalStatus.textContent = "PDF library unavailable. Please try again when the PDF tools are loaded.";
    }
  });
  el.downloadPdfButton.addEventListener("click", () => {
    try {
      const pdfDoc = lastGeneratedPdfDoc || buildProposalPdf();
      const customerName = getProposalData().customerName || "proposal";
      pdfDoc.save(`${customerName.replace(/\s+/g, "_")}_Derby_DIB_Proposal.pdf`);
      el.proposalStatus.textContent = "PDF downloaded.";
    } catch (error) {
      console.error(error);
      el.proposalStatus.textContent = "PDF download could not be completed. Please generate the PDF again.";
    }
  });

  el.companyForm.addEventListener("submit", handleCompanySubmit);
  el.followUpForm.addEventListener("submit", handleFollowUpSubmit);
}

function activatePanel(panelName) {
  if (!canAccessPanel(panelName)) {
    panelName = "dashboard";
  }

  el.navButtons.forEach((button) => {
    button.classList.toggle("active", button.dataset.panelTarget === panelName);
  });
  el.panels.forEach((panel) => {
    panel.classList.toggle("active", panel.dataset.panel === panelName);
  });
}

function openSidebar() {
  el.portalSidebar?.classList.add("open");
  el.sidebarOverlay?.classList.add("visible");
  document.body.style.overflow = "hidden";
}

function closeSidebar() {
  el.portalSidebar?.classList.remove("open");
  el.sidebarOverlay?.classList.remove("visible");
  document.body.style.overflow = "";
}

function handleViewportChange() {
  if (window.innerWidth > 1180) {
    closeSidebar();
  }
}

function hydrateSession() {
  if (!isValidSession(state.session)) {
    clearSession();
  }
}

function handleLoginSubmit(event) {
  event.preventDefault();
  const formData = new FormData(el.loginForm);
  const username = cleanAgentName(formData.get("username"));
  const password = textValue(formData.get("password"));
  const session = authenticateUser({ username, password });

  if (!session) {
    el.loginError.textContent = "Invalid login. Use the provided demo credentials.";
    return;
  }

  state.session = session;
  persistStorage(STORAGE_KEYS.session, session);
  state.search = "";
  state.selectedMonth = "ALL";
  el.loginError.textContent = "";
  activatePanel("dashboard");
  render();
}

function handleLogout() {
  clearSession();
  state.search = "";
  state.selectedMonth = "ALL";
  activatePanel("dashboard");
  closeSidebar();
  render();
}

function authenticateUser({ username, password }) {
  if (normalizeText(username) === "manager" && password === "admin123") {
    return {
      role: "manager",
      username: "manager",
      displayName: "Derby Manager",
    };
  }

  const agentKey = normalizeAgentKey(username);
  if (!ACTIVE_AGENTS.includes(agentKey) || password !== "1234") {
    return null;
  }

  return {
    role: "sales",
    username: agentKey,
    agentKey,
    displayName: ACTIVE_AGENT_DISPLAY_NAMES[agentKey] || agentKey,
  };
}

function isValidSession(session) {
  if (!session || !session.role) return false;
  if (session.role === "manager") return true;
  return session.role === "sales" && ACTIVE_AGENTS.includes(normalizeAgentKey(session.agentKey));
}

function clearSession() {
  state.session = null;
  state.managerReportSnapshot = null;
  localStorage.removeItem(STORAGE_KEYS.session);
}

async function handleMisUpload(event) {
  const [file] = event.target.files || [];
  if (!file) return;

  await importMisFile(file);
  event.target.value = "";
}

function bindUploadDropzone() {
  if (!el.uploadDropzone) return;

  ["dragenter", "dragover"].forEach((eventName) => {
    el.uploadDropzone.addEventListener(eventName, (event) => {
      event.preventDefault();
      el.uploadDropzone.classList.add("drag-active");
    });
  });

  ["dragleave", "dragend", "drop"].forEach((eventName) => {
    el.uploadDropzone.addEventListener(eventName, (event) => {
      event.preventDefault();
      if (eventName === "dragleave" && el.uploadDropzone.contains(event.relatedTarget)) {
        return;
      }
      el.uploadDropzone.classList.remove("drag-active");
    });
  });

  el.uploadDropzone.addEventListener("drop", async (event) => {
    const [file] = event.dataTransfer?.files || [];
    if (!file) return;
    await importMisFile(file);
  });
}

async function importMisFile(file) {
  if (!file) return;

  if (typeof XLSX === "undefined") {
    el.uploadStatus.textContent = "Excel library is unavailable.";
    return;
  }

  try {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    const misSheet = workbook.Sheets.MIS || workbook.Sheets[workbook.SheetNames[0]];
    const staffSheet = workbook.Sheets["Staff Tracker"];

    if (!misSheet || !staffSheet) {
      throw new Error("Required MIS or Staff Tracker sheet missing.");
    }

    const rawMisRows = XLSX.utils.sheet_to_json(misSheet, { defval: "" });
    const rawStaffRows = XLSX.utils.sheet_to_json(staffSheet, { defval: "" });
    const agentColumn = detectColumn(rawMisRows, MIS_COLUMN_NAMES.agent);
    const stageColumn = detectColumn(rawMisRows, MIS_COLUMN_NAMES.stage);
    const amountColumn = detectColumn(rawMisRows, MIS_COLUMN_NAMES.amount);
    const dateColumn = detectColumn(rawMisRows, MIS_COLUMN_NAMES.date);
    if (!agentColumn) {
      el.uploadStatus.textContent = "Employee Name column not found in MIS";
      el.mappingMeta.textContent = "";
      throw new Error("Employee Name column not found in MIS");
    }
    state.detectedAgentColumn = agentColumn;

    state.misRows = rawMisRows
      .map((row) =>
        normalizeMisRow(row, {
          agentColumn,
          stageColumn,
          amountColumn,
          dateColumn,
        }),
      )
      .filter((row) => row.appId || row.agent || row.customer || row.stage);

    state.activeStaff = rawStaffRows
      .map(normalizeStaffRow)
      .filter((row) => normalizeText(row.visaStatus) === "issued" && row.employeeName);

    state.sourceName = file.name;
    state.importedAt = new Date().toISOString();
    state.selectedMonth = "ALL";
    const metrics = computeMetrics(state.misRows);
    const { agentMap, performanceRows } = buildAgentPerformance(state.misRows);

    el.uploadStatus.textContent = `${file.name} imported successfully.`;
    el.mappingMeta.textContent = `Agent column: ${state.detectedAgentColumn || "Not detected"} | Active staff source: Employee Name`;
    console.log(
      "LAMIAA rows",
      rawMisRows.filter((r) => normalizeAgentKey(readField(r, agentColumn)) === "LAMIAA YOUSSEF"),
    );
    console.log("Agents detected:", Object.keys(agentMap));
    console.log("Performance rows:", performanceRows);
    if (!performanceRows.length) {
      console.error("No agent performance rows generated", {
        agentColumn,
        totalRows: state.misRows.length,
        sampleRow: rawMisRows[0],
      });
      console.log("Full MIS data sample:", state.misRows.slice(0, 10));
    }
    console.log({
      agentColumn,
      stageColumn,
      amountColumn,
      dateColumn,
      agentRows: state.misRows.filter((row) => row.agent).length,
      groupedAgents: Object.keys(agentMap),
      performanceRows,
      totalRows: state.misRows.length,
      submissionCount: metrics.activeSubmissionCount,
      bookedEndCount: metrics.bookedCasesCount,
      bookedEndValue: metrics.totalBookedValue,
      bookedMonthlyCount: metrics.currentMonthBookedCount,
      selectedMonth: state.selectedMonth,
    });
    render();
  } catch (error) {
    console.error(error);
    el.uploadStatus.textContent = "Unable to import workbook. Please verify the file structure.";
    el.mappingMeta.textContent = "";
  }
}

function render() {
  renderAuthState();
  if (!state.session) {
    return;
  }
  populateMonthFilter();
  renderDashboard();
  renderSubmissions();
  renderDsrCalculator();
  renderEmiCalculator();
  renderProposalPreview();
  renderCompanyListings();
  renderFollowUps();
}

function renderAuthState() {
  const loggedIn = Boolean(state.session);
  const isManager = state.session?.role === "manager";
  const displayName = state.session?.displayName || "Derby Team";
  const displayRole = isManager ? "Manager View" : "Sales Agent View";

  if (el.authScreen) {
    el.authScreen.hidden = loggedIn;
  }

  if (el.appShell) {
    el.appShell.hidden = !loggedIn;
  }

  if (el.userDisplayName) {
    el.userDisplayName.textContent = displayName;
  }

  if (el.userDisplayRole) {
    el.userDisplayRole.textContent = loggedIn ? displayRole : "Sales Management";
  }

  if (el.userAvatarText) {
    el.userAvatarText.textContent = getAvatarLabel(state.session);
  }

  el.navButtons.forEach((button) => {
    const isAllowed = canAccessPanel(button.dataset.panelTarget);
    button.hidden = !loggedIn || !isAllowed;
  });

  el.panels.forEach((panel) => {
    const isAllowed = canAccessPanel(panel.dataset.panel);
    panel.hidden = !loggedIn || !isAllowed;
    if (panel.hidden) {
      panel.classList.remove("active");
    }
  });

  if (el.managerDashboardContent) {
    el.managerDashboardContent.hidden = !isManager;
  }

  if (el.salesDashboardContent) {
    el.salesDashboardContent.hidden = !loggedIn || isManager;
  }

  if (loggedIn && !isManager) {
    const activePanel = el.panels.find((panel) => panel.classList.contains("active"))?.dataset.panel;
    if (!activePanel || !canAccessPanel(activePanel)) {
      activatePanel("dashboard");
    }
  }

  if (!loggedIn && el.loginError) {
    el.loginError.textContent = "";
  }

  if (!isManager) {
    state.managerReportSnapshot = null;
  }

  if (el.dashboardExportStatus && !loggedIn) {
    el.dashboardExportStatus.textContent = "";
  }

  syncProposalAgentContact();
}

function getAvatarLabel(session) {
  if (!session) return "DA";
  if (session.role === "manager") return "MG";

  return (session.agentKey || "")
    .split(/\s+/)
    .filter(Boolean)
    .slice(0, 2)
    .map((part) => part[0] || "")
    .join("");
}

function canAccessPanel(panelName) {
  if (!state.session) return false;
  if (state.session.role === "manager") return true;
  return SALES_ALLOWED_PANELS.has(panelName);
}

function syncProposalAgentContact() {
  if (!el.proposalForm || state.session?.role !== "sales") {
    return;
  }

  const agentContactInput = el.proposalForm.elements.namedItem("agentContact");
  if (!agentContactInput) {
    return;
  }

  const suggestedValue = `${state.session.displayName || state.session.agentKey} | Sales Agent`;
  if (!textValue(agentContactInput.value) || agentContactInput.value.includes("Omar Ali | 050 000 0000")) {
    agentContactInput.value = suggestedValue;
  }
}

function populateMonthFilter() {
  const options = getMonthOptions(state.misRows);
  if (!options.some((option) => option.value === state.selectedMonth)) {
    state.selectedMonth = "ALL";
  }

  el.monthFilter.innerHTML = options
    .map((option) => `<option value="${escapeHtml(option.value)}">${escapeHtml(option.label)}</option>`)
    .join("");
  el.monthFilter.value = state.selectedMonth;
  el.searchInput.value = state.search;
}

function renderDashboard() {
  const baseRows = getBaseRows();
  const visibleRows = getScopedRows();
  const metrics = computeMetrics(visibleRows);
  const dashboardEndRows = visibleRows.filter((row) => row.stageKey === "END");
  const dashboardBookedRows = visibleRows.filter((row) => row.stageKey === "BOOKED");
  const activeRows = visibleRows.filter(isActiveSubmission);
  const { agentMap, unmatchedRows, performanceRows } = buildAgentPerformance(visibleRows);
  const agentSummary = performanceRows;
  const teamPerformanceRows =
    state.session?.role === "sales" ? buildAgentPerformance(baseRows).performanceRows : performanceRows;
  const leaderboardRows = buildLeaderboardRows(teamPerformanceRows);
  const pipelineRows = buildPipelineRows(visibleRows);
  const pipelineKpis = computePipelineKpis(pipelineRows);
  const isManager = state.session?.role === "manager";

  el.sourceMeta.textContent = state.importedAt
    ? `${state.sourceName} imported on ${formatDateTime(state.importedAt)}${isManager ? "" : ` | Filtered for ${state.session.displayName}`}`
    : `Using ${state.sourceName}${isManager ? "." : ` | Filtered for ${state.session.displayName}.`}`;
  el.activeStaffBadge.textContent = `${state.activeStaff.length} active staff`;
  el.activeCasesBadge.textContent = `${metrics.activeSubmissionCount} active submissions`;
  console.log({
    selectedMonth: state.selectedMonth,
    totalRows: state.misRows.length,
    filteredMonthRows: filterRowsByMonth(state.misRows, state.selectedMonth).length,
    agentMatchedRows: Object.values(agentMap).flatMap((bucket) => bucket.rows).length,
    unmatchedRows,
    leaderboardRows,
    pipelineRows,
    pipelineKpis,
    performanceRows,
  });

  state.managerReportSnapshot = isManager
    ? buildManagerReportSnapshot({
        metrics,
        dashboardEndRows,
        dashboardBookedRows,
        agentSummary,
        leaderboardRows,
        pipelineKpis,
        pipelineRows,
      })
    : null;

  el.kpiGrid.innerHTML = [
    {
      label: "Total Submission",
      value: formatCurrency(metrics.totalSubmissionValue),
      foot: `${metrics.activeSubmissionCount} active submissions`,
    },
    {
      label: "Total Booked (END)",
      value: formatCurrency(dashboardEndRows.reduce((sum, row) => sum + row.financeAmount, 0)),
      foot: `${dashboardEndRows.length} END cases`,
    },
    {
      label: "Current Month Booked (BOOKED)",
      value: formatCurrency(dashboardBookedRows.reduce((sum, row) => sum + row.financeAmount, 0)),
      foot: `${dashboardBookedRows.length} BOOKED cases in ${monthLabelFromKey(state.selectedMonth)}`,
    },
    {
      label: "Conversion %",
      value: formatPercent(metrics.conversion),
      foot: `${metrics.conversionCaseCount} END / ${metrics.activeSubmissionCount} submissions`,
    },
  ]
    .map(
      (item) => `
        <article class="kpi-card">
          <p>${escapeHtml(item.label)}</p>
          <h3>${escapeHtml(item.value)}</h3>
          <div class="kpi-foot">${escapeHtml(item.foot)}</div>
        </article>
      `,
    )
    .join("");

  el.agentTableBody.innerHTML = agentSummary.length
    ? agentSummary
        .map(
          (row) => `
            <tr>
              <td>${escapeHtml(row.agentName)}</td>
              <td class="number">${formatCurrency(row.totalSubmissionValue)}</td>
              <td class="number">${formatCurrency(row.totalBookedValue)}</td>
              <td class="number">${formatCurrency(row.currentMonthBookedValue)}</td>
              <td class="number">${row.bookedCasesCount}</td>
              <td class="number">${formatPercent(row.conversion)}</td>
            </tr>
          `,
        )
        .join("")
    : `<tr><td colspan="6" class="empty-state">No agent data found — check MIS file format</td></tr>`;

  el.leaderboardBody.innerHTML = leaderboardRows.length
    ? leaderboardRows
        .map(
          (row) => `
            <tr>
              <td>${renderRankBadge(row.rank)}</td>
              <td>${escapeHtml(row.agentName.toUpperCase())}</td>
              <td class="number">${formatCurrency(row.totalSubmissionValue)}</td>
              <td class="number">${formatCurrency(row.totalBookedValue)}</td>
              <td class="number">${formatCurrency(row.currentMonthBookedValue)}</td>
              <td class="number">${row.bookedCasesCount}</td>
              <td class="number">${formatPercent(row.conversion)}</td>
              <td>${renderLeaderboardStatusBadge(row)}</td>
            </tr>
          `,
        )
        .join("")
    : `<tr><td colspan="8" class="empty-state">No leaderboard data available.</td></tr>`;

  el.pipelineKpiGrid.innerHTML = [
    ["Total Pipeline Cases", String(pipelineKpis.totalPipelineCases), "Active pending pipeline"],
    ["Total Pipeline Value", formatCurrency(pipelineKpis.totalPipelineValue), "Visible pipeline AED"],
    ["Refer Cases", String(pipelineKpis.referCases), "Stages containing REFER"],
    ["CIF / Post Sanc Cases", String(pipelineKpis.cifPostSancCases), "CIF and Post Sanc workload"],
    ["Urgent Follow-ups", String(pipelineKpis.urgentFollowUps), "Ageing 8 days or more"],
  ]
    .map(
      ([title, value, note]) => `
        <div class="result-card">
          <strong>${escapeHtml(title)}</strong>
          <div class="value">${escapeHtml(value)}</div>
          <span class="note">${escapeHtml(note)}</span>
        </div>
      `,
    )
    .join("");

  el.pipelineTableBody.innerHTML = pipelineRows.length
    ? pipelineRows
        .map(
          (row) => `
            <tr>
              <td>${formatDate(row.date)}</td>
              <td>${escapeHtml(row.agentName.toUpperCase())}</td>
              <td>${escapeHtml(row.customer)}</td>
              <td>${escapeHtml(row.company)}</td>
              <td class="number">${formatCurrency(row.financeAmount)}</td>
              <td>${renderStageChip(row.stage)}</td>
              <td class="number">${row.ageingDaysLabel}</td>
              <td>${renderFollowUpStatusBadge(row.followUpStatus)}</td>
            </tr>
          `,
        )
        .join("")
    : `<tr><td colspan="8" class="empty-state">No pipeline cases found.</td></tr>`;

  if (!isManager) {
    renderSalesDashboard({
      visibleRows,
      metrics,
      dashboardEndRows,
      dashboardBookedRows,
      pipelineRows,
      leaderboardRows,
    });
  }

  el.stageList.innerHTML = renderStageList(activeRows);
  el.previewBody.innerHTML = renderPreviewRows(visibleRows);
}

function renderSubmissions() {
  const visibleRows = getScopedRows();
  const metrics = computeMetrics(visibleRows);

  el.submissionMiniStats.innerHTML = [
    { label: "Active Submission Value", value: formatCurrency(metrics.totalSubmissionValue) },
    { label: "Total Booked Value (BOOKED -)", value: formatCurrency(metrics.totalBookedValue) },
    { label: "Current Month Booked (BOOKED)", value: formatCurrency(metrics.currentMonthBookedValue) },
    { label: "Conversion", value: formatPercent(metrics.conversion) },
  ]
    .map(
      (item) => `
        <div class="result-card">
          <strong>${escapeHtml(item.label)}</strong>
          <div class="value">${escapeHtml(item.value)}</div>
        </div>
      `,
    )
    .join("");

  el.submissionTableBody.innerHTML = visibleRows.length
    ? visibleRows
        .map(
          (row) => `
            <tr>
              <td>${formatDate(row.date)}</td>
              <td>${escapeHtml(row.appId)}</td>
              <td>${escapeHtml(row.agent)}</td>
              <td>${escapeHtml(row.customer)}</td>
              <td>${escapeHtml(row.company)}</td>
              <td>${renderStageChip(row.stage)}</td>
              <td class="number">${formatCurrency(row.financeAmount)}</td>
            </tr>
          `,
        )
        .join("")
    : `<tr><td colspan="7" class="empty-state">No submission rows available.</td></tr>`;
}

function renderSalesDashboard({
  visibleRows,
  metrics,
  dashboardEndRows,
  dashboardBookedRows,
  pipelineRows,
  leaderboardRows,
}) {
  const agentKey = state.session?.agentKey || "";
  const displayName = state.session?.displayName || ACTIVE_AGENT_DISPLAY_NAMES[agentKey] || agentKey;
  const monthlyTarget = getMonthlyTarget(agentKey);
  const achieved = dashboardBookedRows.reduce((sum, row) => sum + row.financeAmount, 0);
  const remaining = Math.max(0, monthlyTarget - achieved);
  const achievementRatio = monthlyTarget > 0 ? achieved / monthlyTarget : 0;
  const urgentRows = pipelineRows.filter((row) => row.followUpStatus === "Urgent");

  el.salesAgentBadge.textContent = displayName;
  el.salesCasesBadge.textContent = `${visibleRows.filter((row) => row.stageKey !== "CANCEL").length} visible cases`;

  el.salesKpiGrid.innerHTML = [
    ["Current Submission Value", formatCurrency(metrics.totalSubmissionValue), `${metrics.activeSubmissionCount} active submissions`],
    ["Total Booked Value", formatCurrency(dashboardEndRows.reduce((sum, row) => sum + row.financeAmount, 0)), `${dashboardEndRows.length} END cases`],
    ["Current Month Booked Value", formatCurrency(dashboardBookedRows.reduce((sum, row) => sum + row.financeAmount, 0)), `${dashboardBookedRows.length} BOOKED cases`],
    ["Conversion %", formatPercent(metrics.conversion), `${metrics.conversionCaseCount} END / ${metrics.activeSubmissionCount} submissions`],
  ]
    .map(
      ([label, value, note]) => `
        <article class="kpi-card">
          <p>${escapeHtml(label)}</p>
          <h3>${escapeHtml(value)}</h3>
          <div class="kpi-foot">${escapeHtml(note)}</div>
        </article>
      `,
    )
    .join("");

  el.salesTargetGrid.innerHTML = [
    ["Monthly Target", formatCurrency(monthlyTarget), "Configured sales target"],
    ["Achieved", formatCurrency(achieved), `BOOKED value in ${monthLabelFromKey(state.selectedMonth)}`],
    ["Remaining", formatCurrency(remaining), "Gap to monthly target"],
    ["Achievement %", formatPercent(achievementRatio), "Achieved against monthly target"],
  ]
    .map(
      ([title, value, note]) => `
        <div class="result-card">
          <strong>${escapeHtml(title)}</strong>
          <div class="value">${escapeHtml(value)}</div>
          <span class="note">${escapeHtml(note)}</span>
        </div>
      `,
    )
    .join("");

  el.salesUrgentBody.innerHTML = urgentRows.length
    ? urgentRows
        .map(
          (row) => `
            <tr>
              <td>${escapeHtml(row.customer)}</td>
              <td>${escapeHtml(row.company)}</td>
              <td class="number">${formatCurrency(row.financeAmount)}</td>
              <td>${renderStageChip(row.stage)}</td>
              <td class="number">${row.ageingDaysLabel}</td>
            </tr>
          `,
        )
        .join("")
    : `<tr><td colspan="5" class="empty-state">No urgent cases found.</td></tr>`;

  el.salesPipelineBody.innerHTML = pipelineRows.length
    ? pipelineRows
        .map(
          (row) => `
            <tr>
              <td>${escapeHtml(row.customer)}</td>
              <td>${escapeHtml(row.company)}</td>
              <td class="number">${formatCurrency(row.financeAmount)}</td>
              <td>${renderStageChip(row.stage)}</td>
              <td class="number">${row.ageingDaysLabel}</td>
              <td>${renderFollowUpStatusBadge(row.followUpStatus)}</td>
            </tr>
          `,
        )
        .join("")
    : `<tr><td colspan="6" class="empty-state">No pipeline cases found for this agent.</td></tr>`;

  el.salesLeaderboardBody.innerHTML = leaderboardRows.length
    ? leaderboardRows
        .map(
          (row) => `
            <tr>
              <td>${renderRankBadge(row.rank)}</td>
              <td>${escapeHtml(row.agentName.toUpperCase())}</td>
              <td class="number">${formatCurrency(row.totalBookedValue)}</td>
              <td class="number">${formatPercent(row.conversion)}</td>
            </tr>
          `,
        )
        .join("")
    : `<tr><td colspan="4" class="empty-state">No leaderboard data available.</td></tr>`;
}

function buildManagerReportSnapshot({
  metrics,
  dashboardEndRows,
  dashboardBookedRows,
  agentSummary,
  leaderboardRows,
  pipelineKpis,
  pipelineRows,
}) {
  return {
    reportDate: new Date().toISOString(),
    selectedMonth: state.selectedMonth,
    kpis: [
      {
        label: "Total Submission",
        value: metrics.totalSubmissionValue,
        foot: `${metrics.activeSubmissionCount} active submissions`,
      },
      {
        label: "Total Booked (END)",
        value: dashboardEndRows.reduce((sum, row) => sum + row.financeAmount, 0),
        foot: `${dashboardEndRows.length} END cases`,
      },
      {
        label: "Current Month Booked (BOOKED)",
        value: dashboardBookedRows.reduce((sum, row) => sum + row.financeAmount, 0),
        foot: `${dashboardBookedRows.length} BOOKED cases in ${monthLabelFromKey(state.selectedMonth)}`,
      },
      {
        label: "Conversion %",
        value: metrics.conversion,
        foot: `${metrics.conversionCaseCount} END / ${metrics.activeSubmissionCount} submissions`,
      },
    ],
    agentPerformance: agentSummary,
    leaderboard: leaderboardRows,
    pipelineKpis,
    pipelineRows,
  };
}

function exportManagerPdfReport() {
  if (state.session?.role !== "manager") {
    return;
  }

  if (!window.jspdf) {
    setDashboardExportStatus("PDF library unavailable. Please try again once PDF tools are loaded.");
    return;
  }

  const snapshot = state.managerReportSnapshot;
  if (!snapshot) {
    setDashboardExportStatus("No manager dashboard data available for export.");
    return;
  }

  try {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
    const monthLabel = monthLabelFromKey(snapshot.selectedMonth);

    doc.setFillColor(15, 61, 46);
    doc.rect(0, 0, 210, 24, "F");
    doc.setTextColor(255, 255, 255);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(18);
    doc.text("Derby - DIB Sales Portal", 14, 14);
    doc.setFontSize(10);
    doc.text("Management Report", 14, 20);

    doc.setTextColor(19, 40, 31);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(10);
    doc.text(`Report Date: ${formatDateTime(snapshot.reportDate)}`, 14, 32);
    doc.text(`Analysis Month: ${monthLabel}`, 14, 38);

    doc.setFont("helvetica", "bold");
    doc.setFontSize(12);
    doc.text("KPI Summary", 14, 50);
    doc.autoTable({
      startY: 54,
      head: [["Metric", "Value", "Notes"]],
      body: snapshot.kpis.map((item) => [
        item.label,
        item.label.includes("Conversion") ? formatPercent(item.value) : formatCurrency(item.value),
        item.foot,
      ]),
      theme: "grid",
      styles: { fontSize: 9, textColor: [19, 40, 31] },
      headStyles: { fillColor: [15, 61, 46], textColor: [255, 255, 255] },
    });

    doc.setFont("helvetica", "bold");
    doc.setFontSize(12);
    doc.text("Agent Performance", 14, doc.lastAutoTable.finalY + 10);
    doc.autoTable({
      startY: doc.lastAutoTable.finalY + 14,
      head: [[
        "Agent Name",
        "Current Submission Value",
        "Booked Value (END)",
        "Current Month Booked Value",
        "Booked Cases Count",
        "Conversion %",
      ]],
      body: snapshot.agentPerformance.map((row) => [
        row.agentName,
        formatCurrency(row.totalSubmissionValue),
        formatCurrency(row.totalBookedValue),
        formatCurrency(row.currentMonthBookedValue),
        String(row.bookedCasesCount),
        formatPercent(row.conversion),
      ]),
      theme: "striped",
      styles: { fontSize: 8, textColor: [19, 40, 31] },
      headStyles: { fillColor: [15, 61, 46], textColor: [255, 255, 255] },
    });

    doc.addPage();
    doc.setFillColor(15, 61, 46);
    doc.rect(0, 0, 210, 20, "F");
    doc.setTextColor(255, 255, 255);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(16);
    doc.text("Leaderboard and Pipeline Summary", 14, 13);

    doc.setTextColor(19, 40, 31);
    doc.setFontSize(12);
    doc.text("Leaderboard", 14, 30);
    doc.autoTable({
      startY: 34,
      head: [[
        "Rank",
        "Agent Name",
        "Current Submission Value",
        "Booked Value (END)",
        "Current Month Booked Value",
        "Booked Cases",
        "Conversion %",
      ]],
      body: snapshot.leaderboard.map((row) => [
        String(row.rank),
        row.agentName,
        formatCurrency(row.totalSubmissionValue),
        formatCurrency(row.totalBookedValue),
        formatCurrency(row.currentMonthBookedValue),
        String(row.bookedCasesCount),
        formatPercent(row.conversion),
      ]),
      theme: "grid",
      styles: { fontSize: 8, textColor: [19, 40, 31] },
      headStyles: { fillColor: [198, 161, 91], textColor: [19, 40, 31] },
    });

    doc.setFont("helvetica", "bold");
    doc.setFontSize(12);
    doc.text("Pipeline KPI Summary", 14, doc.lastAutoTable.finalY + 10);
    doc.autoTable({
      startY: doc.lastAutoTable.finalY + 14,
      head: [["Metric", "Value"]],
      body: [
        ["Total Pipeline Cases", String(snapshot.pipelineKpis.totalPipelineCases)],
        ["Total Pipeline Value", formatCurrency(snapshot.pipelineKpis.totalPipelineValue)],
        ["Refer Cases", String(snapshot.pipelineKpis.referCases)],
        ["CIF / Post Sanc Cases", String(snapshot.pipelineKpis.cifPostSancCases)],
        ["Urgent Follow-ups", String(snapshot.pipelineKpis.urgentFollowUps)],
      ],
      theme: "grid",
      styles: { fontSize: 9, textColor: [19, 40, 31] },
      headStyles: { fillColor: [15, 61, 46], textColor: [255, 255, 255] },
    });

    doc.save(`Derby_DIB_Report_${snapshot.selectedMonth || "ALL"}.pdf`);
    setDashboardExportStatus("PDF report downloaded.");
  } catch (error) {
    console.error(error);
    setDashboardExportStatus("PDF report could not be generated.");
  }
}

function exportManagerExcelReport() {
  if (state.session?.role !== "manager") {
    return;
  }

  if (typeof XLSX === "undefined") {
    setDashboardExportStatus("Excel library unavailable. Please try again once Excel tools are loaded.");
    return;
  }

  const snapshot = state.managerReportSnapshot;
  if (!snapshot) {
    setDashboardExportStatus("No manager dashboard data available for export.");
    return;
  }

  try {
    const workbook = XLSX.utils.book_new();
    const monthLabel = monthLabelFromKey(snapshot.selectedMonth);

    const kpiSheet = XLSX.utils.json_to_sheet([
      ...snapshot.kpis.map((item) => ({
        Metric: item.label,
        Value: item.label.includes("Conversion") ? formatPercent(item.value) : formatCurrency(item.value),
        Notes: item.foot,
        "Analysis Month": monthLabel,
        "Report Date": formatDateTime(snapshot.reportDate),
      })),
    ]);
    applyWorksheetLayout(kpiSheet);
    XLSX.utils.book_append_sheet(workbook, kpiSheet, "KPI Summary");

    const performanceSheet = XLSX.utils.json_to_sheet(
      snapshot.agentPerformance.map((row) => ({
        "Agent Name": row.agentName,
        "Current Submission Value": formatCurrency(row.totalSubmissionValue),
        "Booked Value (END)": formatCurrency(row.totalBookedValue),
        "Current Month Booked Value": formatCurrency(row.currentMonthBookedValue),
        "Booked Cases Count": row.bookedCasesCount,
        "Conversion %": formatPercent(row.conversion),
      })),
    );
    applyWorksheetLayout(performanceSheet);
    XLSX.utils.book_append_sheet(workbook, performanceSheet, "Agent Performance");

    const leaderboardSheet = XLSX.utils.json_to_sheet(
      snapshot.leaderboard.map((row) => ({
        Rank: row.rank,
        "Agent Name": row.agentName,
        "Current Submission Value": formatCurrency(row.totalSubmissionValue),
        "Booked Value (END)": formatCurrency(row.totalBookedValue),
        "Current Month Booked Value": formatCurrency(row.currentMonthBookedValue),
        "Booked Cases Count": row.bookedCasesCount,
        "Conversion %": formatPercent(row.conversion),
      })),
    );
    applyWorksheetLayout(leaderboardSheet);
    XLSX.utils.book_append_sheet(workbook, leaderboardSheet, "Leaderboard");

    const pipelineSheet = XLSX.utils.json_to_sheet(
      snapshot.pipelineRows.map((row) => ({
        Date: formatDate(row.date),
        "Agent Name": row.agentName,
        "Customer Name": row.customer,
        Company: row.company,
        "Finance Amount": formatCurrency(row.financeAmount),
        "Current Stage": row.stage,
        "Ageing Days": row.ageingDaysLabel,
        "Follow-up Status": row.followUpStatus,
      })),
    );
    applyWorksheetLayout(pipelineSheet);
    XLSX.utils.book_append_sheet(workbook, pipelineSheet, "Pipeline Cases");

    XLSX.writeFile(workbook, `Derby_DIB_Report_${snapshot.selectedMonth || "ALL"}.xlsx`);
    setDashboardExportStatus("Excel report downloaded.");
  } catch (error) {
    console.error(error);
    setDashboardExportStatus("Excel report could not be generated.");
  }
}

function applyWorksheetLayout(worksheet) {
  const ref = worksheet["!ref"];
  if (!ref) return;

  const range = XLSX.utils.decode_range(ref);
  const columnWidths = [];

  for (let col = range.s.c; col <= range.e.c; col += 1) {
    let maxLength = 12;
    for (let row = range.s.r; row <= range.e.r; row += 1) {
      const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
      const cellValue = worksheet[cellAddress]?.v;
      maxLength = Math.max(maxLength, String(cellValue ?? "").length + 2);
    }
    columnWidths.push({ wch: Math.min(maxLength, 32) });
  }

  worksheet["!cols"] = columnWidths;
  worksheet["!freeze"] = { xSplit: 0, ySplit: 1, topLeftCell: "A2", activePane: "bottomLeft", state: "frozen" };
}

function setDashboardExportStatus(message) {
  if (el.dashboardExportStatus) {
    el.dashboardExportStatus.textContent = message;
  }
}

function renderDsrCalculator() {
  const values = readNumberForm(el.dsrForm, {
    salary: 0,
    otherIncome: 0,
    existingEmi: 0,
    rent: 0,
    cardLimit: 0,
    otherObligations: 0,
    policyCap: 50,
  });

  const totalIncome = values.salary + values.otherIncome;
  const cardLiability = values.cardLimit * 0.05;
  const totalObligations =
    values.existingEmi +
    values.rent +
    cardLiability +
    values.otherObligations;
  const dsr = totalIncome ? totalObligations / totalIncome : 0;
  const availableInstallment = Math.max(0, totalIncome * (values.policyCap / 100) - totalObligations);

  el.dsrResult.innerHTML = [
    ["Total Income", formatCurrency(totalIncome), "Salary plus other income"],
    ["Card Liability", formatCurrency(cardLiability), "5% of total credit card limits"],
    ["Total Obligations", formatCurrency(totalObligations), "EMI, rent, card liability, and other obligations"],
    ["DSR", formatPercent(dsr), `Policy cap ${values.policyCap}%`],
    ["Available EMI Capacity", formatCurrency(availableInstallment), "Indicative new installment capacity"],
  ]
    .map(
      ([title, value, note]) => `
        <div class="result-card">
          <strong>${escapeHtml(title)}</strong>
          <div class="value">${escapeHtml(value)}</div>
          <span class="note">${escapeHtml(note)}</span>
        </div>
      `,
    )
    .join("");
}

function renderEmiCalculator() {
  const values = readNumberForm(el.emiForm, {
    principal: 0,
    annualRate: 0,
    tenureMonths: 1,
  });
  const startDate = el.emiStartDate.value || toDateInputValue(new Date());
  const schedule = buildEmiSchedule(values.principal, values.annualRate, values.tenureMonths, startDate);
  const totalPayable = schedule.reduce((sum, item) => sum + item.emi, 0);
  const totalProfit = schedule.reduce((sum, item) => sum + item.profit, 0);

  el.emiSummary.innerHTML = [
    ["Monthly EMI", formatCurrency(schedule[0]?.emi || 0), "Calculated installment"],
    ["Total Profit", formatCurrency(totalProfit), "Across full tenor"],
    ["Total Payable", formatCurrency(totalPayable), "Principal plus profit"],
    ["Maturity Date", schedule.at(-1)?.dateLabel || "-", "Final installment date"],
  ]
    .map(
      ([title, value, note]) => `
        <div class="result-card">
          <strong>${escapeHtml(title)}</strong>
          <div class="value">${escapeHtml(value)}</div>
          <span class="note">${escapeHtml(note)}</span>
        </div>
      `,
    )
    .join("");

  el.emiTableBody.innerHTML = schedule.length
    ? schedule
        .slice(0, 24)
        .map(
          (item) => `
            <tr>
              <td>${item.month}</td>
              <td>${escapeHtml(item.dateLabel)}</td>
              <td class="number">${formatCurrency(item.emi)}</td>
              <td class="number">${formatCurrency(item.principal)}</td>
              <td class="number">${formatCurrency(item.profit)}</td>
              <td class="number">${formatCurrency(item.balance)}</td>
            </tr>
          `,
        )
        .join("")
    : `<tr><td colspan="6" class="empty-state">No EMI schedule available.</td></tr>`;
}

function renderProposalPreview() {
  syncProposalAgentContact();
  const proposal = getProposalData();
  const emiSchedule = buildEmiSchedule(
    proposal.financeAmount,
    proposal.rate,
    proposal.tenure,
    el.emiStartDate.value || toDateInputValue(new Date()),
  );
  const monthlyEmi = emiSchedule[0]?.emi || 0;

  el.proposalPreview.innerHTML = [
    ["Customer Details", proposal.customerName || "-", `${proposal.appId || "-"} | ${proposal.company || "-"}`],
    ["Finance Details", formatCurrency(proposal.financeAmount), `${proposal.tenure} months at ${proposal.rate.toFixed(2)}%`],
    ["EMI Details", formatCurrency(monthlyEmi), "Indicative monthly installment"],
    ["Agent Contact", proposal.agentContact || "-", "Prepared through Derby - DIB Sales Portal"],
    ["Benefits", proposal.benefits || "-", "Key customer value proposition"],
    ["Notes", proposal.notes || "-", "Policy and approval reminder, including 5% credit card liability basis for DSR."],
  ]
    .map(
      ([title, value, note]) => `
        <div class="result-card">
          <strong>${escapeHtml(title)}</strong>
          <div class="value">${escapeHtml(value)}</div>
          <span class="note">${escapeHtml(note)}</span>
        </div>
      `,
    )
    .join("");
}

function buildProposalPdf() {
  if (!window.jspdf) {
    throw new Error("jsPDF library unavailable.");
  }

  const proposal = getProposalData();
  const emiSchedule = buildEmiSchedule(
    proposal.financeAmount,
    proposal.rate,
    proposal.tenure,
    el.emiStartDate.value || toDateInputValue(new Date()),
  );
  const monthlyEmi = emiSchedule[0]?.emi || 0;
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF();

  doc.setFillColor(15, 61, 46);
  doc.rect(0, 0, 210, 28, "F");
  doc.setTextColor(255, 255, 255);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(18);
  doc.text("Derby - DIB Sales Portal", 14, 16);
  doc.setFontSize(10);
  doc.text("Personal Finance Sales Management System", 14, 22);

  doc.setTextColor(19, 47, 37);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(12);
  doc.text("Customer Details", 14, 40);
  doc.setFont("helvetica", "normal");
  doc.text(`Customer Name: ${proposal.customerName || "-"}`, 14, 48);
  doc.text(`Customer ID / App ID: ${proposal.appId || "-"}`, 14, 55);
  doc.text(`Company: ${proposal.company || "-"}`, 14, 62);

  doc.setFont("helvetica", "bold");
  doc.text("Finance Details", 14, 78);
  doc.setFont("helvetica", "normal");
  doc.text(`Finance Amount: ${formatCurrency(proposal.financeAmount)}`, 14, 86);
  doc.text(`Tenure: ${proposal.tenure} months`, 14, 93);
  doc.text(`Rate: ${proposal.rate.toFixed(2)}%`, 14, 100);

  doc.setFont("helvetica", "bold");
  doc.text("EMI Details", 14, 116);
  doc.setFont("helvetica", "normal");
  doc.text(`Indicative EMI: ${formatCurrency(monthlyEmi)}`, 14, 124);
  doc.text(`Schedule Start: ${formatDate(el.emiStartDate.value || new Date())}`, 14, 131);

  doc.setFont("helvetica", "bold");
  doc.text("Benefits", 14, 147);
  doc.setFont("helvetica", "normal");
  doc.text(doc.splitTextToSize(proposal.benefits || "-", 180), 14, 155);

  doc.setFont("helvetica", "bold");
  doc.text("Notes", 14, 178);
  doc.setFont("helvetica", "normal");
  doc.text(
    doc.splitTextToSize(
      `${proposal.notes || "-"}\n\nPolicy basis: credit card liability is assessed at 5% of total card limits for affordability review.`,
      180,
    ),
    14,
    186,
  );

  doc.setFont("helvetica", "bold");
  doc.text("Agent Contact", 14, 220);
  doc.setFont("helvetica", "normal");
  doc.text(proposal.agentContact || "-", 14, 228);

  doc.autoTable({
    startY: 238,
    head: [["Month", "EMI", "Principal", "Profit"]],
    body: emiSchedule.slice(0, 6).map((item) => [
      String(item.month),
      formatCurrency(item.emi),
      formatCurrency(item.principal),
      formatCurrency(item.profit),
    ]),
    theme: "grid",
    styles: { fontSize: 9, textColor: [19, 47, 37] },
    headStyles: { fillColor: [195, 155, 78] },
  });

  return doc;
}

function handleCompanySubmit(event) {
  event.preventDefault();
  const formData = new FormData(event.currentTarget);
  const company = textValue(formData.get("company"));
  if (!company) return;

  state.companyListings.unshift({
    id: createId(),
    company,
    status: textValue(formData.get("status")) || "New",
    notes: textValue(formData.get("notes")),
  });
  persistStorage(STORAGE_KEYS.companyListings, state.companyListings);
  event.currentTarget.reset();
  renderCompanyListings();
}

function renderCompanyListings() {
  el.companyList.innerHTML = state.companyListings.length
    ? state.companyListings
        .map(
          (item) => `
            <div class="simple-item">
              <strong>${escapeHtml(item.company)}</strong>
              <div>${renderStageChip(item.status)}</div>
              <span>${escapeHtml(item.notes || "-")}</span>
            </div>
          `,
        )
        .join("")
    : `<div class="simple-item"><span>No companies tracked.</span></div>`;
}

function handleFollowUpSubmit(event) {
  event.preventDefault();
  const formData = new FormData(event.currentTarget);
  const appId = textValue(formData.get("appId"));
  if (!appId) return;

  state.followUps.unshift({
    id: createId(),
    appId,
    customer: textValue(formData.get("customer")),
    stage: textValue(formData.get("stage")),
    notes: textValue(formData.get("notes")),
    agentKey: state.session?.role === "sales" ? state.session.agentKey : "",
  });
  persistStorage(STORAGE_KEYS.followUps, state.followUps);
  event.currentTarget.reset();
  renderFollowUps();
}

function renderFollowUps() {
  const visibleFollowUps =
    state.session?.role === "sales"
      ? state.followUps.filter((item) => item.agentKey === state.session.agentKey)
      : state.followUps;

  el.followUpList.innerHTML = visibleFollowUps.length
    ? visibleFollowUps
        .map(
          (item) => `
            <div class="simple-item">
              <strong>${escapeHtml(item.appId)} | ${escapeHtml(item.customer || "-")}</strong>
              <div>${renderStageChip(item.stage || "-")}</div>
              <span>${escapeHtml(item.notes || "-")}</span>
            </div>
          `,
        )
        .join("")
    : `<div class="simple-item"><span>No follow-ups tracked.</span></div>`;
}

function computeMetrics(rows) {
  const activeSubmissionRows = rows.filter(isActiveSubmission);
  const endRows = rows.filter(isEndStage);
  const bookedMinusRows = rows.filter(isTotalBookedStage);
  const bookedRows = rows.filter((row) => row.stageKey === "BOOKED");

  return {
    activeSubmissionCount: activeSubmissionRows.length,
    totalSubmissionValue: activeSubmissionRows.reduce((sum, row) => sum + row.financeAmount, 0),
    totalBookedValue: bookedMinusRows.reduce((sum, row) => sum + row.financeAmount, 0),
    bookedCasesCount: bookedMinusRows.length,
    currentMonthBookedValue: bookedRows.reduce((sum, row) => sum + row.financeAmount, 0),
    currentMonthBookedCount: bookedRows.length,
    conversionCaseCount: endRows.length,
    conversion:
      activeSubmissionRows.reduce((sum, row) => sum + row.financeAmount, 0) > 0
        ? endRows.reduce((sum, row) => sum + row.financeAmount, 0) /
          activeSubmissionRows.reduce((sum, row) => sum + row.financeAmount, 0)
        : 0,
  };
}

function computeAgentSummary(rows) {
  return buildAgentPerformance(rows).performanceRows;
}

function buildAgentPerformance(rows) {
  const agentBuckets = {};
  ACTIVE_AGENTS.forEach((agent) => {
    agentBuckets[agent] = {
      agentName: ACTIVE_AGENT_DISPLAY_NAMES[agent] || agent,
      rows: [],
    };
  });
  const unmatchedRows = [];

  rows.forEach((row) => {
    const normalizedRow = {
      ...row,
      agentKey: row.agentKey || createAgentKey(row.agent),
      stageKey: row.stageKey || normalizeStage(row.stage),
    };

    if (agentBuckets[normalizedRow.agentKey]) {
      agentBuckets[normalizedRow.agentKey].rows.push(normalizedRow);
    } else {
      unmatchedRows.push(normalizedRow);
    }
  });

  const performanceRows = Object.values(agentBuckets)
    .map((agentBucket) => {
      const agentRows = agentBucket.rows;
      const currentSubmissionRows = agentRows.filter((row) => {
        const stage = row.stageKey;
        return !["BOOKED", "END", "BOOKED -", "CANCEL"].includes(stage);
      });
      const totalBookedValueRows = agentRows.filter((row) => {
        const stage = row.stageKey;
        return stage === "END";
      });
      const currentMonthBookedRows = agentRows.filter((row) => row.stageKey === "BOOKED");
      const bookedCaseRows = agentRows.filter((row) => {
        const stage = row.stageKey;
        return stage === "END" || stage === "BOOKED" || stage === "BOOKED -";
      });
      const totalCasesEntered = agentRows.length;

      return {
        agentName: agentBucket.agentName,
        totalSubmissionValue: currentSubmissionRows.reduce((sum, row) => sum + row.financeAmount, 0),
        totalBookedValue: totalBookedValueRows.reduce((sum, row) => sum + row.financeAmount, 0),
        currentMonthBookedValue: currentMonthBookedRows.reduce((sum, row) => sum + row.financeAmount, 0),
        bookedCasesCount: bookedCaseRows.length,
        submissionCount: totalCasesEntered,
        conversion:
          totalCasesEntered > 0
            ? bookedCaseRows.length / totalCasesEntered
            : 0,
      };
    })
    .sort((a, b) => b.totalSubmissionValue - a.totalSubmissionValue || b.totalBookedValue - a.totalBookedValue);

  console.log({
    agentBuckets,
    unmatchedRows,
    performanceRows,
  });

  return { agentMap: agentBuckets, unmatchedRows, performanceRows };
}

function buildLeaderboardRows(rows) {
  return [...rows]
    .sort(
      (a, b) =>
        b.totalBookedValue - a.totalBookedValue ||
        b.currentMonthBookedValue - a.currentMonthBookedValue ||
        b.conversion - a.conversion,
    )
    .map((row, index) => ({
      ...row,
      rank: index + 1,
    }));
}

function buildPipelineRows(rows) {
  return rows
    .filter((row) => !["BOOKED", "BOOKED -", "END", "CANCEL"].includes(row.stageKey))
    .map((row) => {
      const ageingDays = getAgeingDays(row.date);
      return {
        ...row,
        agentName: ACTIVE_AGENT_DISPLAY_NAMES[row.agentKey] || row.agent,
        ageingDays,
        ageingDaysLabel: ageingDays === null ? "-" : String(ageingDays),
        followUpStatus: getFollowUpStatus(ageingDays),
      };
    })
    .sort(
      (a, b) =>
        getFollowUpPriority(b.followUpStatus) - getFollowUpPriority(a.followUpStatus) ||
        (b.ageingDays ?? -1) - (a.ageingDays ?? -1),
    );
}

function computePipelineKpis(rows) {
  return {
    totalPipelineCases: rows.length,
    totalPipelineValue: rows.reduce((sum, row) => sum + row.financeAmount, 0),
    referCases: rows.filter((row) => row.stageKey.includes("REFER")).length,
    cifPostSancCases: rows.filter((row) => row.stageKey.includes("CIF") || row.stageKey.includes("POST SANC")).length,
    urgentFollowUps: rows.filter((row) => row.followUpStatus === "Urgent").length,
  };
}

function renderRankBadge(rank) {
  if (rank === 1) return `<span class="status-chip status-chip-end">#1</span>`;
  if (rank === 2) return `<span class="status-chip status-chip-average">#2</span>`;
  if (rank === 3) return `<span class="status-chip status-chip-default">#3</span>`;
  return `<span class="status-chip status-chip-default">#${rank}</span>`;
}

function renderLeaderboardStatusBadge(row) {
  if (row.rank === 1) return `<span class="status-chip status-chip-booked">Top Performer</span>`;
  if (row.conversion >= 0.4) return `<span class="status-chip status-chip-booked">Strong</span>`;
  if (row.conversion >= 0.2) return `<span class="status-chip status-chip-average">Average</span>`;
  return `<span class="status-chip status-chip-cancel">Needs Focus</span>`;
}

function renderFollowUpStatusBadge(status) {
  if (status === "Fresh") return `<span class="status-chip status-chip-booked">Fresh</span>`;
  if (status === "Follow Up") return `<span class="status-chip status-chip-average">Follow Up</span>`;
  if (status === "Urgent") return `<span class="status-chip status-chip-cancel">Urgent</span>`;
  return `<span class="status-chip status-chip-default">-</span>`;
}

function getAgeingDays(value) {
  const date = parseDate(value);
  if (!date) return null;
  const today = new Date();
  const startOfToday = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const startOfDate = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  return Math.max(0, Math.floor((startOfToday - startOfDate) / 86400000));
}

function getFollowUpStatus(ageingDays) {
  if (ageingDays === null) return "-";
  if (ageingDays <= 3) return "Fresh";
  if (ageingDays <= 7) return "Follow Up";
  return "Urgent";
}

function getFollowUpPriority(status) {
  if (status === "Urgent") return 3;
  if (status === "Follow Up") return 2;
  if (status === "Fresh") return 1;
  return 0;
}

function renderStageList(rows) {
  const stageSequence = getVisibleStageSequence(rows);
  const counts = stageSequence.map((stage) => ({
    stage,
    count: rows.filter((row) => normalizeStage(row.stage) === stage.toUpperCase()).length,
  }));
  const maxCount = Math.max(1, ...counts.map((item) => item.count));

  return counts
    .map(
      (item) => `
        <div class="stage-item">
          <strong>${renderStageChip(item.stage)}</strong>
          <span>${item.count} cases</span>
          <div class="progress-track"><i style="width:${((item.count / maxCount) * 100).toFixed(1)}%"></i></div>
        </div>
      `,
    )
    .join("");
}

function getVisibleStageSequence(rows) {
  const discoveredStages = [...new Set(
    rows
      .map((row) => textValue(row.stage))
      .filter(Boolean)
      .filter((stage) => !ACTIVE_STAGE_EXCLUSIONS.has(normalizeStage(stage))),
  )];

  return [
    ...STAGE_SEQUENCE.filter((stage) =>
      discoveredStages.some((candidate) => normalizeStage(candidate) === stage.toUpperCase()),
    ),
    ...discoveredStages.filter(
      (stage) => !STAGE_SEQUENCE.some((candidate) => candidate.toUpperCase() === normalizeStage(stage)),
    ),
  ];
}

function renderPreviewRows(rows) {
  if (!rows.length) {
    return `<tr><td colspan="7" class="empty-state">No MIS rows available.</td></tr>`;
  }

  return rows
    .slice(0, 15)
    .map(
      (row) => `
        <tr>
          <td>${formatDate(row.date)}</td>
          <td>${escapeHtml(row.appId)}</td>
          <td>${escapeHtml(row.agent)}</td>
          <td>${escapeHtml(row.customer)}</td>
          <td>${escapeHtml(row.company)}</td>
          <td>${renderStageChip(row.stage)}</td>
          <td class="number">${formatCurrency(row.financeAmount)}</td>
        </tr>
      `,
    )
    .join("");
}

function getProposalData() {
  const formData = new FormData(el.proposalForm);
  return {
    customerName: textValue(formData.get("customerName")),
    appId: textValue(formData.get("appId")),
    agentContact: textValue(formData.get("agentContact")),
    company: textValue(formData.get("company")),
    financeAmount: Number(formData.get("financeAmount") || 0),
    tenure: Number(formData.get("tenure") || 0),
    rate: Number(formData.get("rate") || 0),
    benefits: textValue(formData.get("benefits")),
    notes: textValue(formData.get("notes")),
  };
}

function readNumberForm(form, defaults) {
  const formData = new FormData(form);
  return Object.fromEntries(
    Object.entries(defaults).map(([key, fallback]) => [key, Number(formData.get(key) || fallback)]),
  );
}

function buildEmiSchedule(principal, annualRate, tenureMonths, startDate) {
  const amount = Number(principal) || 0;
  const months = Math.max(1, Math.round(Number(tenureMonths) || 1));
  const monthlyRate = (Number(annualRate) || 0) / 12 / 100;
  if (!amount) return [];

  const emi =
    monthlyRate === 0
      ? amount / months
      : (amount * monthlyRate * Math.pow(1 + monthlyRate, months)) /
        (Math.pow(1 + monthlyRate, months) - 1);

  let balance = amount;
  const baseDate = parseDate(startDate) || new Date();

  return Array.from({ length: months }, (_, index) => {
    const profit = monthlyRate === 0 ? 0 : balance * monthlyRate;
    const principalPart = Math.min(balance, emi - profit);
    balance = Math.max(0, balance - principalPart);
    const dueDate = new Date(baseDate);
    dueDate.setMonth(dueDate.getMonth() + index + 1);

    return {
      month: index + 1,
      dateLabel: formatDate(dueDate),
      emi,
      principal: principalPart,
      profit,
      balance,
    };
  });
}

function detectColumn(rows, columnName) {
  const headers = [...new Set(rows.flatMap((row) => Object.keys(row || {})))];
  return headers.find((header) => normalizeHeader(header) === normalizeHeader(columnName)) || "";
}

function normalizeMisRow(row, columns) {
  const agentName = cleanAgentName(readField(row, columns.agentColumn));
  const stageText = textValue(readField(row, columns.stageColumn));
  return {
    date: normalizeDateValue(readField(row, columns.dateColumn)),
    appId: textValue(row["APP ID"]),
    agent: agentName,
    agentKey: createAgentKey(agentName),
    customer: textValue(row["Customer Name"]),
    company: textValue(row["Company"]),
    stage: stageText,
    stageKey: normalizeStage(stageText),
    financeAmount: parseAmount(readField(row, columns.amountColumn)),
  };
}

function normalizeStaffRow(row) {
  return {
    employeeName: textValue(row["Employee Name"]),
    visaStatus: textValue(row["Visa Status"]),
  };
}

function readField(row, preferredHeader) {
  if (!preferredHeader) return "";
  const matchingKey = Object.keys(row || {}).find(
    (header) => normalizeHeader(header) === normalizeHeader(preferredHeader),
  );
  return matchingKey ? row[matchingKey] : "";
}

function isActiveSubmission(row) {
  return !ACTIVE_STAGE_EXCLUSIONS.has(normalizeStage(row.stage));
}

function isEndStage(row) {
  return normalizeStage(row.stage) === "END";
}

function isCurrentMonthBookedStage(row) {
  return normalizeStage(row.stage) === "BOOKED" && monthKey(row.date) === state.selectedMonth;
}

function isTotalBookedStage(row) {
  return normalizeStage(row.stage) === "BOOKED -";
}

function filterRowsBySearch(rows, search) {
  const query = normalizeText(search);
  const sortedRows = [...rows].sort((a, b) => compareDatesDesc(a.date, b.date));
  if (!query) return sortedRows;

  return sortedRows.filter((row) =>
    [row.appId, row.agent, row.customer, row.company, row.stage]
      .map(normalizeText)
      .some((value) => value.includes(query)),
  );
}

function filterRowsByMonth(rows, selectedMonth) {
  if (!selectedMonth || selectedMonth === "ALL") {
    return [...rows];
  }

  return rows.filter((row) => monthKey(row.date) === selectedMonth);
}

function getBaseRows() {
  return filterRowsBySearch(filterRowsByMonth(state.misRows, state.selectedMonth), state.search);
}

function getScopedRows() {
  const baseRows = getBaseRows();

  if (state.session?.role === "sales" && state.session.agentKey) {
    return baseRows.filter((row) => (row.agentKey || createAgentKey(row.agent)) === state.session.agentKey);
  }

  return baseRows;
}

function getMonthlyTarget(agentKey) {
  return ACTIVE_AGENTS.includes(agentKey) ? DEFAULT_MONTHLY_TARGET : 0;
}

function getMonthOptions(rows) {
  const monthKeys = [...new Set(rows.map((row) => monthKey(row.date)).filter(Boolean))].sort().reverse();
  return [{ value: "ALL", label: "All Months" }, ...monthKeys.map((value) => ({ value, label: monthLabelFromKey(value) }))];
}

function firstTwoNames(value) {
  const tokens = cleanAgentName(value)
    .toUpperCase()
    .replace(/[^A-Z0-9\s]/g, " ")
    .split(/\s+/)
    .filter(Boolean)
    .slice(0, 2)
    .map((token, index) => normalizeAgentToken(token, index));

  return tokens.join(" ");
}

function createAgentKey(value) {
  return firstTwoNames(value);
}

function normalizeAgentKey(value) {
  return createAgentKey(value);
}

function normalizeAgentToken(token, index) {
  if (index === 1 && token === "YOUSEF") {
    return "YOUSSEF";
  }
  return token;
}

function matchesAgentPrefix(agentName, staffPrefix) {
  const agentPrefix = firstTwoNames(agentName);
  return Boolean(agentPrefix) && agentPrefix === staffPrefix;
}

function cleanAgentName(value) {
  return textValue(`${value ?? ""}`.replace(/\u00A0/g, " ")).replace(/\s+/g, " ").trim();
}

function parseAmount(value) {
  if (typeof value === "number") return value;
  const cleaned = `${value || ""}`.replace(/[^0-9.-]/g, "");
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
}

function normalizeDateValue(value) {
  if (value instanceof Date) return value.toISOString();
  if (typeof value === "number" && typeof XLSX !== "undefined" && XLSX.SSF) {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) {
      return new Date(parsed.y, parsed.m - 1, parsed.d).toISOString();
    }
  }
  const date = parseDate(value);
  return date ? date.toISOString() : "";
}

function parseDate(value) {
  if (!value) return null;
  if (value instanceof Date) return value;
  const direct = new Date(value);
  if (!Number.isNaN(direct.getTime())) return direct;

  const tokens = `${value}`.split(/[-/]/).map((part) => part.trim());
  if (tokens.length === 3) {
    const guess = new Date(`${tokens[1]} ${tokens[0]}, ${tokens[2]}`);
    if (!Number.isNaN(guess.getTime())) return guess;
  }

  return null;
}

function compareDatesDesc(a, b) {
  return (parseDate(b)?.getTime() || 0) - (parseDate(a)?.getTime() || 0);
}

function monthKey(value) {
  const date = parseDate(value);
  if (!date) return "";
  return `${date.getFullYear()}-${`${date.getMonth() + 1}`.padStart(2, "0")}`;
}

function monthLabelFromKey(value) {
  if (value === "ALL") return "All Months";
  if (!value) return "-";
  const [year, month] = value.split("-");
  return new Intl.DateTimeFormat("en-GB", {
    month: "short",
    year: "numeric",
  }).format(new Date(Number(year), Number(month) - 1, 1));
}

function formatCurrency(value) {
  return new Intl.NumberFormat("en-AE", {
    style: "currency",
    currency: "AED",
    maximumFractionDigits: 0,
  }).format(Number(value) || 0);
}

function formatPercent(value) {
  return `${((Number(value) || 0) * 100).toFixed(1)}%`;
}

function formatDate(value) {
  const date = parseDate(value);
  if (!date) return "-";
  return new Intl.DateTimeFormat("en-GB", {
    day: "2-digit",
    month: "short",
    year: "numeric",
  }).format(date);
}

function formatDateTime(value) {
  const date = parseDate(value);
  if (!date) return "-";
  return new Intl.DateTimeFormat("en-GB", {
    day: "2-digit",
    month: "short",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  }).format(date);
}

function toDateInputValue(date) {
  return new Date(date.getTime() - date.getTimezoneOffset() * 60000).toISOString().slice(0, 10);
}

function normalizeStage(value) {
  return textValue(value)
    .replace(/\u00A0/g, " ")
    .replace(/\s*-\s*/g, " - ")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}

function normalizeText(value) {
  return textValue(value).toLowerCase();
}

function normalizeHeader(value) {
  return textValue(value).toLowerCase().replace(/[^a-z0-9]/g, "");
}

function textValue(value) {
  return `${value ?? ""}`.trim();
}

function renderStageChip(value) {
  const raw = textValue(value) || "-";
  const normalized = normalizeStage(raw);
  let tone = "default";

  if (normalized === "BOOKED") {
    tone = "booked";
  } else if (normalized === "BOOKED -") {
    tone = "booked-minus";
  } else if (normalized === "END") {
    tone = "end";
  } else if (normalized === "CANCEL") {
    tone = "cancel";
  } else if (normalized.includes("REFER") || normalized === "DETAIL DATA ENTRY") {
    tone = "refer";
  }

  return `<span class="status-chip status-chip-${tone}">${escapeHtml(raw)}</span>`;
}

function escapeHtml(value) {
  return `${value ?? ""}`
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function createId() {
  return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

function readStorage(key, fallback) {
  try {
    const value = localStorage.getItem(key);
    return value ? JSON.parse(value) : fallback;
  } catch {
    return fallback;
  }
}

function persistStorage(key, value) {
  localStorage.setItem(key, JSON.stringify(value));
}

function buildSampleMisRows() {
  return [
    {
      date: "2026-05-05",
      appId: "PF-1001",
      agent: "OMAR ALI",
      customer: "Ahmed Nasser",
      company: "Listed A",
      stage: "Refer",
      financeAmount: 210000,
    },
    {
      date: "2026-05-06",
      appId: "PF-1002",
      agent: "OMAR ALI",
      customer: "Layla Hasan",
      company: "MOD",
      stage: "Detail Data Entry",
      financeAmount: 160000,
    },
    {
      date: "2026-05-08",
      appId: "PF-1003",
      agent: "MOHAB HESHAM",
      customer: "Faris Khalid",
      company: "NTML",
      stage: "BOOKED",
      financeAmount: 395000,
    },
    {
      date: "2026-05-09",
      appId: "PF-1004",
      agent: "EL SAYED MANSOUR",
      customer: "Mariam Ali",
      company: "ADNOC",
      stage: "END",
      financeAmount: 325000,
    },
    {
      date: "2026-05-10",
      appId: "PF-1005",
      agent: "ABDALLA HASSAN",
      customer: "Omar Yousif",
      company: "Global Medical",
      stage: "CIF Account Linkage",
      financeAmount: 185000,
    },
    {
      date: "2026-05-11",
      appId: "PF-1006",
      agent: "MONA SOLIMAN",
      customer: "Sara Rahman",
      company: "Hospitality Group",
      stage: "BOOKED -",
      financeAmount: 270000,
    },
    {
      date: "2026-05-12",
      appId: "PF-1007",
      agent: "LAYLA KAREEM",
      customer: "Huda Saif",
      company: "Healthcare",
      stage: "Post Sanc Doc",
      financeAmount: 142000,
    },
  ];
}

function buildSampleStaffRows() {
  return [
    { employeeName: "OMAR ALI MOHAMED IBRAHIM", visaStatus: "Issued" },
    { employeeName: "Mohab Hesham Salah", visaStatus: "Issued" },
    { employeeName: "El Sayed Mansour El Sayed Soliman", visaStatus: "Issued" },
    { employeeName: "Abdalla Hassan Abdelfattah Mohamed", visaStatus: "Issued" },
    { employeeName: "Mona Soliman Ahmed Youssef", visaStatus: "Issued" },
    { employeeName: "Layla Kareem Hassan", visaStatus: "Issued" },
  ];
}
