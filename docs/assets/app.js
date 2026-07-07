const state = {
  config: null,
  questions: [],
  dashboardCards: [],
  descriptiveResults: [],
  multipleChoiceResults: [],
  crosstabResults: [],
  questionById: new Map(),
  resultByQuestionId: new Map(),
  crosstabById: new Map(),
  currentQuestionRows: [],
  currentCrosstabRows: [],
  mainChart: null,
  crosstabChart: null,
};

const els = {};

document.addEventListener("DOMContentLoaded", initApp);

async function initApp() {
  cacheElements();

  try {
    const config = await fetchJson("../data/app/app_config.json");
    const files = config.data_files || {};

    const [questions, dashboardCards, descriptiveResults, multipleChoiceResults, crosstabResults] = await Promise.all([
      fetchJson(files.questions || "../data/app/questions.json"),
      fetchJson(files.dashboard_cards || "../data/app/dashboard_cards.json"),
      fetchJson(files.descriptive_results || "../data/app/descriptive_results.json"),
      fetchJson(files.multiple_choice_results || "../data/app/multiple_choice_results.json"),
      fetchJson(files.crosstab_results || "../data/app/crosstab_results.json"),
    ]);

    state.config = config;
    state.questions = questions;
    state.dashboardCards = dashboardCards;
    state.descriptiveResults = descriptiveResults;
    state.multipleChoiceResults = multipleChoiceResults;
    state.crosstabResults = crosstabResults;
    state.questionById = new Map(questions.map((q) => [String(q.question_id), q]));
    state.resultByQuestionId = new Map([
      ...descriptiveResults.map((r) => [String(r.question_id), r]),
      ...multipleChoiceResults.map((r) => [String(r.question_id), r]),
    ]);
    state.crosstabById = new Map(crosstabResults.map((r) => [String(r.id), r]));

    renderShell();
    renderDashboard();
    renderTabs();
    renderMethodNotes();

    const firstSection = config.sections?.[0]?.id || "caracterizacao";
    selectSection(firstSection);

    showApp();
  } catch (error) {
    console.error(error);
    els.loadingPanel.innerHTML = `
      <h2>Erro ao carregar dados</h2>
      <p>Verifique se os arquivos JSON já foram gerados por <code>scripts/02_generate_app_outputs.py</code>.</p>
      <pre>${escapeHtml(String(error))}</pre>
    `;
  }
}

function cacheElements() {
  els.loadingPanel = document.getElementById("loading-panel");
  els.dashboardCards = document.getElementById("dashboard-cards");
  els.navigationPanel = document.getElementById("navigation-panel");
  els.analysisPanel = document.getElementById("analysis-panel");
  els.methodPanel = document.getElementById("method-panel");
  els.title = document.getElementById("app-title");
  els.subtitle = document.getElementById("app-subtitle");
  els.sectionTabs = document.getElementById("section-tabs");
  els.currentSectionLabel = document.getElementById("current-section-label");
  els.currentSectionTitle = document.getElementById("current-section-title");
  els.currentSectionObjective = document.getElementById("current-section-objective");
  els.questionArea = document.getElementById("question-area");
  els.crosstabArea = document.getElementById("crosstab-area");
  els.questionSelect = document.getElementById("question-select");
  els.crosstabSelect = document.getElementById("crosstab-select");
  els.questionIdLabel = document.getElementById("question-id-label");
  els.questionTitle = document.getElementById("question-title");
  els.questionMeta = document.getElementById("question-meta");
  els.questionSummary = document.getElementById("question-summary");
  els.frequencyTable = document.getElementById("frequency-table");
  els.downloadFrequencyBtn = document.getElementById("download-frequency-btn");
  els.crosstabTitle = document.getElementById("crosstab-title");
  els.crosstabMeta = document.getElementById("crosstab-meta");
  els.crosstabStats = document.getElementById("crosstab-stats");
  els.crosstabWarning = document.getElementById("crosstab-warning");
  els.crosstabCountsTable = document.getElementById("crosstab-counts-table");
  els.crosstabRowPctTable = document.getElementById("crosstab-rowpct-table");
  els.downloadCrosstabBtn = document.getElementById("download-crosstab-btn");
  els.methodNotes = document.getElementById("method-notes");

  els.questionSelect.addEventListener("change", () => renderQuestion(els.questionSelect.value));
  els.crosstabSelect.addEventListener("change", () => renderCrosstab(els.crosstabSelect.value));
  els.downloadFrequencyBtn.addEventListener("click", () => downloadRowsAsCsv(state.currentQuestionRows, "frequencia.csv"));
  els.downloadCrosstabBtn.addEventListener("click", () => downloadRowsAsCsv(state.currentCrosstabRows, "cruzamento.csv"));
}

async function fetchJson(url) {
  const response = await fetch(url);
  if (!response.ok) throw new Error(`${url}: ${response.status} ${response.statusText}`);
  return response.json();
}

function showApp() {
  els.loadingPanel.classList.add("hidden");
  els.dashboardCards.classList.remove("hidden");
  els.navigationPanel.classList.remove("hidden");
  els.analysisPanel.classList.remove("hidden");
  els.methodPanel.classList.remove("hidden");
}

function renderShell() {
  els.title.textContent = state.config.title || "Análise do questionário";
  els.subtitle.textContent = state.config.subtitle || "";
}

function renderDashboard() {
  els.dashboardCards.innerHTML = state.dashboardCards.map((card) => `
    <article class="metric-card">
      <p class="metric-card__label">${escapeHtml(card.title || card.id)}</p>
      <p class="metric-card__value">${escapeHtml(card.display_value ?? card.value ?? "—")}</p>
      <p class="metric-card__label">${escapeHtml(card.note || "")}</p>
    </article>
  `).join("");
}

function renderTabs() {
  els.sectionTabs.innerHTML = (state.config.sections || []).map((section) => `
    <button class="tab-button" type="button" data-section-id="${escapeHtml(section.id)}">
      ${escapeHtml(section.title)}
    </button>
  `).join("");

  els.sectionTabs.querySelectorAll("button").forEach((button) => {
    button.addEventListener("click", () => selectSection(button.dataset.sectionId));
  });
}

function renderMethodNotes() {
  els.methodNotes.innerHTML = (state.config.notes || []).map((note) => `<li>${escapeHtml(note)}</li>`).join("");
}

function selectSection(sectionId) {
  const section = (state.config.sections || []).find((s) => s.id === sectionId);
  if (!section) return;

  els.sectionTabs.querySelectorAll("button").forEach((button) => {
    button.classList.toggle("active", button.dataset.sectionId === sectionId);
  });

  els.currentSectionLabel.textContent = "Módulo de análise";
  els.currentSectionTitle.textContent = section.title;
  els.currentSectionObjective.textContent = section.objective || "";

  if (sectionId === "cruzamentos") {
    els.questionArea.classList.add("hidden");
    els.crosstabArea.classList.remove("hidden");
    populateCrosstabSelect();
    renderCrosstab(els.crosstabSelect.value);
  } else {
    els.crosstabArea.classList.add("hidden");
    els.questionArea.classList.remove("hidden");
    populateQuestionSelect(section);
    renderQuestion(els.questionSelect.value);
  }
}

function populateQuestionSelect(section) {
  const questions = (section.questions || [])
    .map((id) => state.questionById.get(String(id)))
    .filter((q) => q && state.resultByQuestionId.has(String(q.question_id)));

  els.questionSelect.innerHTML = questions.map((q) => `
    <option value="${escapeHtml(q.question_id)}">${escapeHtml(q.question_id)} — ${escapeHtml(stripQuestionPrefix(q.question_text))}</option>
  `).join("");
}

function populateCrosstabSelect() {
  els.crosstabSelect.innerHTML = state.crosstabResults.map((ct) => `
    <option value="${escapeHtml(ct.id)}">${escapeHtml(ct.title)}</option>
  `).join("");
}

function renderQuestion(questionId) {
  const result = state.resultByQuestionId.get(String(questionId));
  if (!result) {
    els.questionTitle.textContent = "Sem resultado pré-calculado";
    els.frequencyTable.innerHTML = "";
    destroyChart("mainChart");
    return;
  }

  els.questionIdLabel.textContent = result.question_id;
  els.questionTitle.textContent = result.question_text || result.question_label;
  els.questionMeta.textContent = `${result.app_section || ""} · ${labelVariableType(result.variable_type)}`;

  renderQuestionSummary(result);
  renderFrequencyTable(result);
  renderFrequencyChart(result);

  state.currentQuestionRows = result.rows || [];
}

function renderQuestionSummary(result) {
  const s = result.summary || {};
  if (result.variable_type === "multiple_choice") {
    els.questionSummary.innerHTML = `
      <span><strong>${formatNumber(s.total_respondents || 0)}</strong> respondentes válidas</span>
      <span><strong>${formatNumber(s.n_options || 0)}</strong> opções</span>
      <span><strong>${formatNumber(s.total_selected || 0)}</strong> seleções</span>
    `;
  } else {
    els.questionSummary.innerHTML = `
      <span><strong>${formatNumber(s.valid_n || 0)}</strong> respostas válidas</span>
      <span><strong>${formatNumber(s.missing_n || 0)}</strong> sem resposta</span>
      <span><strong>${formatNumber(s.n_categories || 0)}</strong> categorias</span>
    `;
  }
}

function renderFrequencyTable(result) {
  const rows = result.rows || [];
  const isMultiple = result.variable_type === "multiple_choice";

  if (isMultiple) {
    els.frequencyTable.innerHTML = makeTable(
      ["Opção", "Selecionaram (n)", "N válido", "% válido", "% total"],
      rows.map((r) => [r.option, r.selected_n, r.valid_n, formatPercent(r.percent_valid), formatPercent(r.percent_total_respondents)])
    );
  } else {
    els.frequencyTable.innerHTML = makeTable(
      ["Resposta", "n", "% válido", "% total"],
      rows.map((r) => [r.category, r.n, formatPercent(r.percent_valid), formatPercent(r.percent_total)])
    );
  }
}

function renderFrequencyChart(result) {
  const rows = result.rows || [];
  const isMultiple = result.variable_type === "multiple_choice";
  const labels = rows.map((r) => isMultiple ? r.option : r.category);
  const values = rows.map((r) => isMultiple ? r.selected_n : r.n);
  const isHorizontal = isMultiple || labels.length > 6;

  const ctx = document.getElementById("main-chart");
  destroyChart("mainChart");
  state.mainChart = new Chart(ctx, {
    type: "bar",
    data: {
      labels,
      datasets: [{
        label: isMultiple ? "Selecionaram (n)" : "Frequência (n)",
        data: values,
      }],
    },
    options: chartOptions({ horizontal: isHorizontal }),
  });
}

function renderCrosstab(crosstabId) {
  const result = state.crosstabById.get(String(crosstabId));
  if (!result) return;

  els.crosstabTitle.textContent = result.title;
  els.crosstabMeta.textContent = `${result.x_variable} × ${result.y_variable}`;

  const stats = result.statistics || {};
  els.crosstabStats.innerHTML = `
    <article><span>Qui-quadrado</span><strong>${formatStat(stats.chi_square)}</strong></article>
    <article><span>p-valor</span><strong>${formatStat(stats.p_value)}</strong></article>
    <article><span>Cramér's V</span><strong>${formatStat(stats.cramers_v)}</strong></article>
    <article><span>Efeito</span><strong>${escapeHtml(stats.effect_size_label || "—")}</strong></article>
  `;

  if (result.warning) {
    els.crosstabWarning.textContent = result.warning;
    els.crosstabWarning.classList.remove("hidden");
  } else {
    els.crosstabWarning.textContent = "";
    els.crosstabWarning.classList.add("hidden");
  }

  els.crosstabCountsTable.innerHTML = objectTable(result.counts_table || []);
  els.crosstabRowPctTable.innerHTML = objectTable(result.row_percent_table || [], true);
  renderCrosstabChart(result);

  state.currentCrosstabRows = result.plot_rows || [];
}

function renderCrosstabChart(result) {
  const plotRows = result.plot_rows || [];
  const rowCategories = [...new Set(plotRows.map((r) => r.row_category))];
  const colCategories = [...new Set(plotRows.map((r) => r.column_category))];

  const datasets = colCategories.map((col) => ({
    label: col,
    data: rowCategories.map((row) => {
      const found = plotRows.find((r) => r.row_category === row && r.column_category === col);
      return found ? found.row_percent : 0;
    }),
  }));

  const ctx = document.getElementById("crosstab-chart");
  destroyChart("crosstabChart");
  state.crosstabChart = new Chart(ctx, {
    type: "bar",
    data: { labels: rowCategories, datasets },
    options: chartOptions({ horizontal: rowCategories.length > 6, percentAxis: true }),
  });
}

function chartOptions({ horizontal = false, percentAxis = false } = {}) {
  return {
    indexAxis: horizontal ? "y" : "x",
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: { display: true },
      tooltip: {
        callbacks: {
          label: (context) => `${context.dataset.label}: ${percentAxis ? formatPercent(context.raw) : formatNumber(context.raw)}`,
        },
      },
    },
    scales: {
      x: { beginAtZero: true, ticks: { callback: (value) => percentAxis && !horizontal ? `${value}%` : value } },
      y: { beginAtZero: true, ticks: { callback: (value) => percentAxis && horizontal ? `${value}%` : value } },
    },
  };
}

function destroyChart(key) {
  if (state[key]) {
    state[key].destroy();
    state[key] = null;
  }
}

function makeTable(headers, rows) {
  return `
    <table>
      <thead><tr>${headers.map((h) => `<th>${escapeHtml(h)}</th>`).join("")}</tr></thead>
      <tbody>
        ${rows.map((row) => `<tr>${row.map((cell) => `<td>${escapeHtml(cell ?? "—")}</td>`).join("")}</tr>`).join("")}
      </tbody>
    </table>
  `;
}

function objectTable(rows, percent = false) {
  if (!rows.length) return "<p class='muted'>Sem dados para apresentar.</p>";
  const headers = Object.keys(rows[0]);
  return makeTable(headers, rows.map((row) => headers.map((h, index) => {
    const value = row[h];
    if (index > 0 && percent && typeof value === "number") return formatPercent(value);
    return value;
  })));
}

function downloadRowsAsCsv(rows, filename) {
  if (!rows || !rows.length) return;
  const headers = Object.keys(rows[0]);
  const csv = [headers.join(","), ...rows.map((row) => headers.map((h) => csvEscape(row[h])).join(","))].join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

function csvEscape(value) {
  const text = String(value ?? "");
  if (/[",\n]/.test(text)) return `"${text.replace(/"/g, '""')}"`;
  return text;
}

function stripQuestionPrefix(text) {
  if (!text) return "";
  const idx = text.indexOf(".");
  if (idx > 0 && idx < 8) return text.slice(idx + 1).trim();
  return text;
}

function labelVariableType(type) {
  return type === "multiple_choice" ? "Seleção múltipla" : "Resposta simples";
}

function formatNumber(value) {
  if (value === null || value === undefined || Number.isNaN(Number(value))) return "—";
  return new Intl.NumberFormat("pt-PT").format(Number(value));
}

function formatPercent(value) {
  if (value === null || value === undefined || Number.isNaN(Number(value))) return "—";
  return `${Number(value).toLocaleString("pt-PT", { minimumFractionDigits: 1, maximumFractionDigits: 1 })}%`;
}

function formatStat(value) {
  if (value === null || value === undefined || Number.isNaN(Number(value))) return "—";
  return Number(value).toLocaleString("pt-PT", { maximumFractionDigits: 4 });
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
