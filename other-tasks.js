const STORAGE_KEY = "project-dashboard-other-tasks";

const state = {
  tasks: [],
  searchQuery: "",
  showCompleted: false,
  activeTaskId: null,
  saveStatusTimer: null,
};

const elements = {
  searchInput: document.getElementById("task-search"),
  searchCount: document.getElementById("task-search-count"),
  showCompletedToggle: document.getElementById("show-completed-toggle"),
  addButton: document.getElementById("task-add"),
  form: document.getElementById("task-form"),
  formResetButton: document.getElementById("task-form-reset"),
  formDescription: document.getElementById("task-form-description"),
  formBuilding: document.getElementById("task-form-building"),
  formDateEntered: document.getElementById("task-form-date-entered"),
  formTargetDate: document.getElementById("task-form-target-date"),
  formNotes: document.getElementById("task-form-notes"),
  saveStatus: document.getElementById("task-save-status"),
  openCount: document.getElementById("task-open-count"),
  completedCount: document.getElementById("task-completed-count"),
  dueSoonCount: document.getElementById("task-due-soon-count"),
  tableBody: document.getElementById("task-table-body"),
  empty: document.getElementById("task-empty"),
  dashboardView: document.getElementById("task-dashboard-view"),
  detailView: document.getElementById("task-detail-view"),
  backButton: document.getElementById("task-back"),
  deleteButton: document.getElementById("task-delete"),
  detailTitle: document.getElementById("task-title"),
  detailSubtitle: document.getElementById("task-subtitle"),
  detailDescription: document.getElementById("task-detail-description"),
  detailBuilding: document.getElementById("task-detail-building"),
  detailDateEntered: document.getElementById("task-detail-date-entered"),
  detailTargetDate: document.getElementById("task-detail-target-date"),
  detailCompleted: document.getElementById("task-detail-completed"),
  detailNotes: document.getElementById("task-detail-notes"),
};

init();

function init() {
  state.tasks = loadTasks().map(normalizeTask);
  bindEvents();
  resetForm();
  renderDashboard();
}

function bindEvents() {
  elements.addButton.addEventListener("click", () => {
    switchView("dashboard");
    elements.formDescription.focus();
    elements.formDescription.select();
  });

  elements.form.addEventListener("submit", (event) => {
    event.preventDefault();

    const description = elements.formDescription.value.trim();
    if (!description) {
      elements.formDescription.focus();
      return;
    }

    const task = {
      id: createId(),
      description,
      building: elements.formBuilding.value.trim(),
      dateEntered: normalizeDateValue(elements.formDateEntered.value) || getToday(),
      targetDate: normalizeDateValue(elements.formTargetDate.value),
      notes: elements.formNotes.value.trim(),
      completed: false,
      completedAt: "",
    };

    state.tasks.push(task);
    persistTasks("Task saved locally");
    resetForm();
    renderDashboard();
  });

  elements.formResetButton.addEventListener("click", () => {
    resetForm();
    showSaveStatus("");
  });

  elements.searchInput.addEventListener("input", (event) => {
    state.searchQuery = event.target.value.trim().toLowerCase();
    renderDashboard();
  });

  elements.showCompletedToggle.addEventListener("change", (event) => {
    state.showCompleted = event.target.checked;
    renderDashboard();
  });

  elements.backButton.addEventListener("click", () => {
    cleanupEmptyActiveTask();
    state.activeTaskId = null;
    switchView("dashboard");
    renderDashboard();
  });

  elements.deleteButton.addEventListener("click", () => {
    if (!state.activeTaskId) {
      return;
    }
    state.tasks = state.tasks.filter((task) => task.id !== state.activeTaskId);
    persistTasks("Task deleted");
    state.activeTaskId = null;
    switchView("dashboard");
    renderDashboard();
  });

  elements.detailDescription.addEventListener("input", () => {
    updateActiveTask({ description: elements.detailDescription.value.trim() });
  });

  elements.detailBuilding.addEventListener("input", () => {
    updateActiveTask({ building: elements.detailBuilding.value.trim() });
  });

  elements.detailDateEntered.addEventListener("input", () => {
    updateActiveTask({ dateEntered: normalizeDateValue(elements.detailDateEntered.value) });
  });

  elements.detailTargetDate.addEventListener("input", () => {
    updateActiveTask({ targetDate: normalizeDateValue(elements.detailTargetDate.value) });
  });

  elements.detailCompleted.addEventListener("change", () => {
    const completed = elements.detailCompleted.checked;
    updateActiveTask({
      completed,
      completedAt: completed ? getToday() : "",
    });
  });

  elements.detailNotes.addEventListener("input", () => {
    updateActiveTask({ notes: elements.detailNotes.value });
  });

  document.addEventListener("keydown", (event) => {
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "k") {
      event.preventDefault();
      elements.searchInput.focus();
      elements.searchInput.select();
    }
  });
}

function renderDashboard() {
  const scopedTasks = state.tasks.filter((task) => task.completed === state.showCompleted);
  const query = state.searchQuery;
  const filtered = query
    ? scopedTasks.filter((task) => taskMatchesQuery(task, query))
    : scopedTasks.slice();

  filtered.sort(sortTasks);

  elements.tableBody.innerHTML = "";
  elements.empty.classList.toggle("is-hidden", filtered.length > 0);
  elements.searchCount.textContent = query
    ? `${filtered.length} result${filtered.length === 1 ? "" : "s"}`
    : "";
  renderSummary();

  filtered.forEach((task) => {
    const row = document.createElement("tr");
    const status = getTaskStatus(task);
    row.classList.toggle("task-row-completed", !!task.completed);
    row.classList.toggle("task-row-overdue", status === "Overdue");
    row.innerHTML = `
      <td><span class="task-status-pill task-status-${status.toLowerCase().replace(/\s+/g, "-")}">${escapeHtml(status)}</span></td>
      <td>${escapeHtml(task.description)}</td>
      <td>${escapeHtml(task.building || "")}</td>
      <td>${escapeHtml(task.dateEntered || "")}</td>
      <td>${escapeHtml(task.targetDate || "")}</td>
      <td>${escapeHtml(getNotesPreview(task.notes))}</td>
    `;
    row.addEventListener("click", () => openTask(task.id));
    elements.tableBody.appendChild(row);
  });

  switchView("dashboard");
}

function openTask(taskId) {
  const task = state.tasks.find((item) => item.id === taskId);
  if (!task) {
    return;
  }
  state.activeTaskId = taskId;
  elements.detailTitle.textContent = task.description || "New Task";
  elements.detailSubtitle.textContent = task.completedAt
    ? `Completed ${task.completedAt}`
    : `Entered ${task.dateEntered || ""}`;
  elements.detailDescription.value = task.description || "";
  elements.detailBuilding.value = task.building || "";
  elements.detailDateEntered.value = normalizeDateValue(task.dateEntered) || getToday();
  elements.detailTargetDate.value = normalizeDateValue(task.targetDate);
  elements.detailCompleted.checked = !!task.completed;
  elements.detailNotes.value = task.notes || "";
  switchView("detail");
}

function updateActiveTask(update) {
  if (!state.activeTaskId) {
    return;
  }
  const task = state.tasks.find((item) => item.id === state.activeTaskId);
  if (!task) {
    return;
  }
  Object.assign(task, update);
  if (!task.dateEntered) {
    task.dateEntered = getToday();
  }
  persistTasks("Task updated", { silent: true });
  elements.detailTitle.textContent = task.description || "New Task";
  elements.detailSubtitle.textContent = task.completedAt
    ? `Completed ${task.completedAt}`
    : `Entered ${task.dateEntered || ""}`;
  renderSummary();
}

function cleanupEmptyActiveTask() {
  if (!state.activeTaskId) {
    return;
  }
  const task = state.tasks.find((item) => item.id === state.activeTaskId);
  if (!task) {
    return;
  }
  const isEmpty =
    !task.description &&
    !task.building &&
    !task.targetDate &&
    !task.notes &&
    task.dateEntered === getToday() &&
    !task.completed;
  if (isEmpty) {
    state.tasks = state.tasks.filter((item) => item.id !== task.id);
    persistTasks("", { silent: true });
  }
}

function switchView(view) {
  const showDashboard = view === "dashboard";
  elements.dashboardView.classList.toggle("is-hidden", !showDashboard);
  elements.detailView.classList.toggle("is-hidden", showDashboard);
}

function createEmptyTask() {
  return {
    id: createId(),
    description: "",
    building: "",
    dateEntered: getToday(),
    targetDate: "",
    notes: "",
    completed: false,
    completedAt: "",
  };
}

function normalizeTask(task) {
  return {
    ...createEmptyTask(),
    ...task,
    dateEntered: normalizeDateValue(task?.dateEntered) || getToday(),
    targetDate: normalizeDateValue(task?.targetDate),
    completed: !!task?.completed,
    completedAt: normalizeDateValue(task?.completedAt),
  };
}

function resetForm() {
  elements.form.reset();
  elements.formDateEntered.value = getToday();
}

function renderSummary() {
  const openTasks = state.tasks.filter((task) => !task.completed);
  const completedTasks = state.tasks.filter((task) => task.completed);
  const dueSoonTasks = openTasks.filter((task) => isDueSoon(task.targetDate));

  elements.openCount.textContent = String(openTasks.length);
  elements.completedCount.textContent = String(completedTasks.length);
  elements.dueSoonCount.textContent = String(dueSoonTasks.length);
}

function getTaskStatus(task) {
  if (task.completed) {
    return "Completed";
  }

  if (task.targetDate && isOverdue(task.targetDate)) {
    return "Overdue";
  }

  if (task.targetDate && isDueSoon(task.targetDate)) {
    return "Due Soon";
  }

  return "Open";
}

function getNotesPreview(notes) {
  const normalized = String(notes || "").replace(/\s+/g, " ").trim();
  if (!normalized) {
    return "";
  }
  return normalized.length > 72 ? `${normalized.slice(0, 69)}...` : normalized;
}


function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll("\"", "&quot;")
    .replaceAll("'", "&#39;");
}

function taskMatchesQuery(task, query) {
  const parts = [
    task.description,
    task.building,
    task.notes,
    task.dateEntered,
    task.targetDate,
    getTaskStatus(task),
  ]
    .filter(Boolean)
    .join(" ")
    .toLowerCase();
  return parts.includes(query);
}

function sortTasks(a, b) {
  const aTarget = parseDateValue(a.targetDate);
  const bTarget = parseDateValue(b.targetDate);
  if (aTarget && bTarget && aTarget !== bTarget) {
    return aTarget - bTarget;
  }

  if (aTarget && !bTarget) {
    return -1;
  }

  if (!aTarget && bTarget) {
    return 1;
  }

  const aDate = parseDateValue(a.dateEntered);
  const bDate = parseDateValue(b.dateEntered);
  if (aDate === bDate) {
    return 0;
  }
  return bDate - aDate;
}

function parseDateValue(value) {
  if (!value) {
    return 0;
  }
  const [year, month, day] = value.split("-").map((part) => Number(part));
  if (!year || !month || !day) {
    return 0;
  }
  return new Date(year, month - 1, day).getTime();
}

function normalizeDateValue(value) {
  if (!value) {
    return "";
  }

  if (/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    return value;
  }

  const slashMatch = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (slashMatch) {
    const month = slashMatch[1].padStart(2, "0");
    const day = slashMatch[2].padStart(2, "0");
    const yearRaw = slashMatch[3];
    const year = yearRaw.length === 2 ? "20" + yearRaw : yearRaw;
    return `${year}-${month}-${day}`;
  }

  return "";
}

function getToday() {
  const today = new Date();
  const year = today.getFullYear();
  const month = String(today.getMonth() + 1).padStart(2, "0");
  const day = String(today.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function createId() {
  return `task-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function isOverdue(value) {
  const target = parseDateValue(value);
  const today = parseDateValue(getToday());
  return !!target && target < today;
}

function isDueSoon(value) {
  const target = parseDateValue(value);
  const today = parseDateValue(getToday());
  if (!target || target < today) {
    return false;
  }
  const difference = target - today;
  const sevenDays = 7 * 24 * 60 * 60 * 1000;
  return difference <= sevenDays;
}

function loadTasks() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      return [];
    }
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch (error) {
    return [];
  }
}

function saveTasks(tasks) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(tasks));
  } catch (error) {
    // Ignore storage failures and keep runtime state only.
  }
}

function persistTasks(message, options = {}) {
  saveTasks(state.tasks);
  if (!options.silent) {
    showSaveStatus(message);
  }
}

function showSaveStatus(message) {
  if (state.saveStatusTimer) {
    window.clearTimeout(state.saveStatusTimer);
    state.saveStatusTimer = null;
  }

  elements.saveStatus.textContent = message;
  if (!message) {
    return;
  }

  state.saveStatusTimer = window.setTimeout(() => {
    elements.saveStatus.textContent = "";
    state.saveStatusTimer = null;
  }, 2500);
}
