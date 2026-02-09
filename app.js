import { Editor, Node as TiptapNode } from "https://esm.sh/@tiptap/core";
import StarterKit from "https://esm.sh/@tiptap/starter-kit";
import { TaskList } from "https://esm.sh/@tiptap/extension-task-list";
import { TaskItem } from "https://esm.sh/@tiptap/extension-task-item";
import { Table } from "https://esm.sh/@tiptap/extension-table";
import { TableRow } from "https://esm.sh/@tiptap/extension-table-row";
import { TableHeader } from "https://esm.sh/@tiptap/extension-table-header";
import { TableCell } from "https://esm.sh/@tiptap/extension-table-cell";
import { Link } from "https://esm.sh/@tiptap/extension-link";
import { Placeholder } from "https://esm.sh/@tiptap/extension-placeholder";
import { Image } from "https://esm.sh/@tiptap/extension-image";

const ROOT_FOLDERS = ["EI", "SCP", "SER", "WO"];
const ROOT_BASE_PATH = "C:\\Users\\u144243\\OneDrive - Eastman Chemical Company\\Documents\\..Projects";
const COMPLETED_FOLDERS = new Set(["complete", "completed"]);
const FIELD_ORDER = [
  { key: "title", label: "Title" },
  { key: "description", label: "Description" },
  { key: "type", label: "Type" },
  { key: "building", label: "Building" },
  { key: "percentComplete", label: "% Complete", type: "number" },
  { key: "status", label: "Status" },
  { key: "assignedDate", label: "Assigned Date", type: "date" },
  { key: "ecDate", label: "EC Date", type: "date" },
  { key: "actualEcDate", label: "Actual EC Date", type: "date" },
  { key: "priority", label: "Priority", type: "number" },
  { key: "divRep", label: "Div Rep" },
  { key: "moc", label: "MOC" },
];

const TYPE_CLASSES = {
  EI: "type-ei",
  SCP: "type-scp",
  WO: "type-wo",
  OTHER: "type-other",
  BLANK: "type-blank",
};

const CALENDAR_FILTERS = new Set(["calendar-current", "calendar-completed"]);

const Summary = TiptapNode.create({
  name: "summary",
  group: "block",
  content: "inline*",
  defining: true,
  addAttributes() {
    return {
      level: {
        default: 0,
        parseHTML: (element) => {
          if (element.classList.contains("toggle-heading-1")) {
            return 1;
          }
          if (element.classList.contains("toggle-heading-2")) {
            return 2;
          }
          if (element.classList.contains("toggle-heading-3")) {
            return 3;
          }
          return 0;
        },
      },
    };
  },
  parseHTML() {
    return [{ tag: "summary" }];
  },
  renderHTML({ node }) {
    const classes = ["toggle-heading"];
    if (node.attrs.level) {
      classes.push("toggle-heading-" + node.attrs.level);
    }
    return ["summary", { class: classes.join(" ") }, 0];
  },
});

const Details = TiptapNode.create({
  name: "details",
  group: "block",
  content: "summary block*",
  defining: true,
  isolating: true,
  addAttributes() {
    return {
      open: {
        default: false,
        parseHTML: (element) => element.hasAttribute("open"),
        renderHTML: (attributes) => (attributes.open ? { open: "true" } : {}),
      },
    };
  },
  parseHTML() {
    return [{ tag: "details" }];
  },
  renderHTML({ HTMLAttributes }) {
    return ["details", HTMLAttributes, 0];
  },
});

const CustomTaskItem = TaskItem.extend({
  addKeyboardShortcuts() {
    return {
      Enter: () => {
        const info = getTaskItemInfo(this.editor);
        if (!info) {
          return false;
        }

        if (isTaskItemEmpty(info)) {
          return this.editor.commands.liftListItem("taskItem");
        }

        return this.editor.commands.splitListItem("taskItem");
      },
      Backspace: () => {
        const info = getTaskItemInfo(this.editor);
        if (!info) {
          return false;
        }

        if (isCursorAtTaskItemStart(this.editor, info)) {
          return this.editor.commands.liftListItem("taskItem");
        }

        return false;
      },
    };
  },
});

const state = {
  rootHandle: null,
  projects: [],
  filtered: [],
  selectedIndex: 0,
  filter: "current",
  searchQuery: "",
  searchResults: [],
  searchIndex: -1,
  activeProject: null,
  tiptapEditor: null,
  useTiptap: true,
  notesMenuActive: false,
  notesMenuQuery: "",
  notesMenuIndex: 0,
  notesMenuResults: [],
};

const elements = {
  rootStatus: document.getElementById("root-status"),
  selectRoot: document.getElementById("select-root"),
  refresh: document.getElementById("refresh"),
  filterButtons: Array.from(document.querySelectorAll(".filter-btn")),
  searchInput: document.getElementById("search-input"),
  searchCount: document.getElementById("search-count"),
  searchResults: document.getElementById("search-results"),
  tableBody: document.getElementById("project-table-body"),
  emptyState: document.getElementById("empty-state"),
  filterGroup: document.getElementById("filter-group"),
  dashboardView: document.getElementById("dashboard-view"),
  calendarView: document.getElementById("calendar-view"),
  calendarScroll: document.getElementById("calendar-scroll"),
  projectView: document.getElementById("project-view"),
  backButton: document.getElementById("back-button"),
  projectTitle: document.getElementById("project-title"),
  projectPath: document.getElementById("project-path"),
  projectFields: document.getElementById("project-fields"),
  notesEditor: document.getElementById("notes-editor"),
  notesMenu: document.getElementById("notes-menu"),
  tableToolbar: document.getElementById("table-toolbar"),
  tableAddRow: document.getElementById("table-add-row"),
  tableDelRow: document.getElementById("table-del-row"),
  tableAddCol: document.getElementById("table-add-col"),
  tableDelCol: document.getElementById("table-del-col"),
  tableDel: document.getElementById("table-del"),
  openProjectFolder: document.getElementById("open-project-folder"),
  copyStatus: document.getElementById("copy-status"),
};

let notesSaveTimer = null;
let jsonSaveTimer = null;

init();

function init() {
  if (!window.showDirectoryPicker) {
    elements.rootStatus.textContent = "This browser does not support folder access.";
    elements.selectRoot.disabled = true;
    return;
  }

  bindEvents();
  restoreRootHandle();
}

function bindEvents() {
  elements.selectRoot.addEventListener("click", selectRootFolder);
  elements.refresh.addEventListener("click", () => loadProjects());
  elements.filterButtons.forEach((button) => {
    button.addEventListener("click", () => {
      state.filter = button.dataset.filter;
      elements.filterButtons.forEach((btn) => btn.classList.toggle("is-active", btn === button));
      state.selectedIndex = 0;
      renderDashboard();
    });
  });

  elements.searchInput.addEventListener("input", (event) => {
    state.searchQuery = event.target.value.trim().toLowerCase();
    updateSearchResults();
  });

  elements.searchInput.addEventListener("keydown", (event) => {
    if (event.key === "ArrowDown") {
      event.preventDefault();
      moveSearchSelection(1);
    }

    if (event.key === "ArrowUp") {
      event.preventDefault();
      moveSearchSelection(-1);
    }

    if (event.key === "Enter") {
      event.preventDefault();
      openSelectedSearchResult();
    }

    if (event.key === "Escape") {
      hideSearchResults();
    }
  });

  document.addEventListener("keydown", (event) => {
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "k") {
      event.preventDefault();
      elements.searchInput.focus();
      elements.searchInput.select();
    }
  });

  document.addEventListener("click", (event) => {
    if (!elements.searchResults.contains(event.target) && !elements.searchInput.contains(event.target)) {
      hideSearchResults();
    }
  });

  document.addEventListener("click", (event) => {
    if (!state.notesMenuActive) {
      return;
    }

    if (elements.notesMenu.contains(event.target) || elements.notesEditor.contains(event.target)) {
      return;
    }

    hideNotesMenu();
  });

  elements.backButton.addEventListener("click", () => {
    state.activeProject = null;
    switchView("dashboard");
    hideSearchResults();
  });

  elements.openProjectFolder.addEventListener("click", () => {
    copyProjectFolderPath();
  });

  elements.tableAddRow.addEventListener("click", () => {
    if (state.tiptapEditor) {
      state.tiptapEditor.chain().focus().addRowAfter().run();
    }
  });

  elements.tableDelRow.addEventListener("click", () => {
    if (state.tiptapEditor) {
      state.tiptapEditor.chain().focus().deleteRow().run();
    }
  });

  elements.tableAddCol.addEventListener("click", () => {
    if (state.tiptapEditor) {
      state.tiptapEditor.chain().focus().addColumnAfter().run();
    }
  });

  elements.tableDelCol.addEventListener("click", () => {
    if (state.tiptapEditor) {
      state.tiptapEditor.chain().focus().deleteColumn().run();
    }
  });

  elements.tableDel.addEventListener("click", () => {
    if (state.tiptapEditor) {
      state.tiptapEditor.chain().focus().deleteTable().run();
    }
  });
}

async function selectRootFolder() {
  try {
    const handle = await window.showDirectoryPicker({ mode: "readwrite" });
    const hasPermission = await verifyPermission(handle, true);
    if (!hasPermission) {
      elements.rootStatus.textContent = "Read/write permission denied.";
      return;
    }

    state.rootHandle = handle;
    await saveRootHandle(handle);
    await loadProjects();
  } catch (error) {
    if (error && error.name !== "AbortError") {
      elements.rootStatus.textContent = "Unable to access the selected folder.";
    }
  }
}

async function restoreRootHandle() {
  const handle = await loadRootHandle();
  if (!handle) {
    return;
  }

  const hasPermission = await verifyPermission(handle, false);
  if (!hasPermission) {
    return;
  }

  state.rootHandle = handle;
  loadProjects();
}

async function loadProjects() {
  if (!state.rootHandle) {
    elements.rootStatus.textContent = "No folder selected.";
    return;
  }

  elements.rootStatus.textContent = "Scanning project folders...";

  const projects = [];

  for (const rootName of ROOT_FOLDERS) {
    const rootDir = await tryGetDirectoryHandle(state.rootHandle, rootName);
    if (!rootDir) {
      continue;
    }

    for await (const entry of rootDir.values()) {
      if (entry.kind !== "directory") {
        continue;
      }

      const buildingName = entry.name;
      const buildingDir = entry;
      const completedDir = await findCompletedDirectory(buildingDir);

      const projectDirs = [];
      for await (const buildingEntry of buildingDir.values()) {
        if (buildingEntry.kind !== "directory") {
          continue;
        }

        const isCompletedFolder = COMPLETED_FOLDERS.has(buildingEntry.name.toLowerCase());
        if (isCompletedFolder) {
          continue;
        }

        projectDirs.push({ dir: buildingEntry, isCompleted: false, completedFolderName: null });
      }

      if (completedDir) {
        for await (const completedEntry of completedDir.values()) {
          if (completedEntry.kind !== "directory") {
            continue;
          }

          projectDirs.push({
            dir: completedEntry,
            isCompleted: true,
            completedFolderName: completedDir.name,
          });
        }
      }

      for (const projectEntry of projectDirs) {
        const project = await readProjectFolder({
          projectDir: projectEntry.dir,
          rootName,
          buildingName,
          isCompleted: projectEntry.isCompleted,
          completedFolderName: projectEntry.completedFolderName,
        });

        if (project) {
          projects.push(project);
        }
      }
    }
  }

  projects.sort(sortProjects);
  state.projects = projects;
  state.selectedIndex = 0;
  elements.rootStatus.textContent = "Loaded " + projects.length + " projects.";
  renderDashboard();
}

async function findCompletedDirectory(buildingDir) {
  for await (const entry of buildingDir.values()) {
    if (entry.kind !== "directory") {
      continue;
    }

    if (COMPLETED_FOLDERS.has(entry.name.toLowerCase())) {
      return entry;
    }
  }

  return null;
}

async function readProjectFolder({
  projectDir,
  rootName,
  buildingName,
  isCompleted,
  completedFolderName,
}) {
  const jsonHandle = await tryGetFileHandle(projectDir, "project.json");
  if (!jsonHandle) {
    return null;
  }

  const jsonText = await readFile(jsonHandle);

  let data = {};
  try {
    data = JSON.parse(jsonText);
  } catch (error) {
    data = {};
  }

  const project = {
    id: rootName + "/" + buildingName + "/" + projectDir.name,
    rootName,
    buildingName,
    folderName: projectDir.name,
    isCompleted,
    completedFolderName: completedFolderName || null,
    data,
    jsonHandle,
    projectDir,
  };

  project.searchText = (
    (data.title || "") +
    " " +
    (data.description || "") +
    " " +
    project.folderName
  )
    .toLowerCase()
    .trim();

  return project;
}

function sortProjects(a, b) {
  const byType = a.rootName.localeCompare(b.rootName);
  if (byType !== 0) {
    return byType;
  }

  const byBuilding = a.buildingName.localeCompare(b.buildingName);
  if (byBuilding !== 0) {
    return byBuilding;
  }

  return (a.data.title || a.folderName).localeCompare(b.data.title || b.folderName);
}

function renderDashboard() {
  if (CALENDAR_FILTERS.has(state.filter)) {
    elements.dashboardView.classList.add("is-hidden");
    elements.calendarView.classList.remove("is-hidden");
    renderCalendar();
    return;
  }

  elements.calendarView.classList.add("is-hidden");
  elements.dashboardView.classList.remove("is-hidden");
  state.filtered = state.projects.filter((project) => {
    const matchesStatus = state.filter === "current" ? !project.isCompleted : project.isCompleted;
    if (!matchesStatus) {
      return false;
    }
    return true;
  });

  elements.tableBody.innerHTML = "";
  elements.emptyState.classList.toggle("is-hidden", state.filtered.length > 0);

  state.filtered.forEach((project, index) => {
    const row = document.createElement("tr");
    const typeClass = TYPE_CLASSES[normalizeType(project.data.type || project.rootName)];
    if (typeClass) {
      row.classList.add(typeClass);
    }
    if (index === state.selectedIndex) {
      row.classList.add("is-selected");
    }

    row.innerHTML = `
      <td>${escapeHtml(project.data.title || "")}</td>
      <td>${escapeHtml(project.data.description || project.folderName)}</td>
      <td>${escapeHtml(project.data.type || project.rootName)}</td>
      <td>${escapeHtml(project.data.building || project.buildingName)}</td>
      <td>${formatPercent(project.data.percentComplete)}</td>
      <td>${escapeHtml(project.data.status || "")}</td>
      <td>${escapeHtml(project.data.assignedDate || "")}</td>
      <td>${escapeHtml(project.data.ecDate || "")}</td>
      <td>${escapeHtml(project.data.actualEcDate || "")}</td>
      <td>${escapeHtml(project.data.divRep || "")}</td>
      <td>${escapeHtml(project.data.moc || "")}</td>
    `;

    row.addEventListener("click", () => openProject(project));
    row.addEventListener("mouseenter", () => {
      state.selectedIndex = index;
      updateSelection();
    });

    elements.tableBody.appendChild(row);
  });

  updateSelection();
}

function normalizeType(value) {
  const raw = String(value || "").trim().toUpperCase();
  if (!raw) {
    return "BLANK";
  }

  if (raw === "EI") {
    return "EI";
  }

  if (raw === "SCP") {
    return "SCP";
  }

  if (raw === "WO") {
    return "WO";
  }

  if (raw === "OTHER") {
    return "OTHER";
  }

  return "OTHER";
}

function updateSearchResults() {
  const query = state.searchQuery;
  if (!query) {
    hideSearchResults();
    return;
  }

  const results = state.projects.filter((project) => project.searchText.includes(query));
  state.searchResults = results.slice(0, 8);
  state.searchIndex = state.searchResults.length ? 0 : -1;
  elements.searchCount.textContent = results.length ? results.length + " results" : "0 results";

  renderSearchResults();
}

function renderSearchResults() {
  elements.searchResults.innerHTML = "";

  if (!state.searchResults.length) {
    elements.searchResults.classList.remove("is-hidden");
    const empty = document.createElement("div");
    empty.className = "search-result";
    empty.textContent = "No matches found.";
    elements.searchResults.appendChild(empty);
    return;
  }

  state.searchResults.forEach((project, index) => {
    const item = document.createElement("div");
    item.className = "search-result" + (index === state.searchIndex ? " is-active" : "");
    item.setAttribute("role", "option");
    item.setAttribute("aria-selected", index === state.searchIndex ? "true" : "false");

    const title = document.createElement("div");
    title.className = "search-result-title";
    title.innerHTML = escapeHtml(project.data.description || project.folderName);

    const meta = document.createElement("div");
    meta.className = "search-result-meta";
    const metaText = [project.data.title, project.data.type || project.rootName, project.data.building || project.buildingName]
      .filter(Boolean)
      .join(" • ");
    meta.innerHTML = escapeHtml(metaText);

    item.append(title, meta);
    item.addEventListener("click", () => {
      openProject(project);
      hideSearchResults();
    });
    item.addEventListener("mouseenter", () => {
      state.searchIndex = index;
      updateSearchSelection();
    });

    elements.searchResults.appendChild(item);
  });

  elements.searchResults.classList.remove("is-hidden");
}

function updateSearchSelection() {
  const items = Array.from(elements.searchResults.querySelectorAll(".search-result"));
  items.forEach((item, index) => {
    item.classList.toggle("is-active", index === state.searchIndex);
    item.setAttribute("aria-selected", index === state.searchIndex ? "true" : "false");
  });
}

function moveSearchSelection(delta) {
  if (!state.searchResults.length) {
    return;
  }

  state.searchIndex = Math.max(0, Math.min(state.searchIndex + delta, state.searchResults.length - 1));
  updateSearchSelection();
}

function openSelectedSearchResult() {
  if (!state.searchResults.length || state.searchIndex < 0) {
    return;
  }

  const project = state.searchResults[state.searchIndex];
  if (project) {
    openProject(project);
    hideSearchResults();
  }
}

function hideSearchResults() {
  state.searchResults = [];
  state.searchIndex = -1;
  elements.searchResults.classList.add("is-hidden");
  elements.searchResults.innerHTML = "";
  if (!state.searchQuery) {
    elements.searchCount.textContent = "";
  }
}

function updateSelection() {
  const rows = Array.from(elements.tableBody.querySelectorAll("tr"));
  rows.forEach((row, index) => row.classList.toggle("is-selected", index === state.selectedIndex));
}

function moveSelection(delta) {
  if (!state.filtered.length) {
    return;
  }

  state.selectedIndex = Math.max(0, Math.min(state.selectedIndex + delta, state.filtered.length - 1));
  updateSelection();
}

function openSelectedProject() {
  if (!state.filtered.length) {
    return;
  }

  const project = state.filtered[state.selectedIndex];
  if (project) {
    openProject(project);
  }
}

function openProject(project) {
  state.activeProject = project;
  elements.projectTitle.textContent = project.data.description || project.folderName;
  elements.projectPath.textContent = project.id;
  renderProjectFields(project);
  renderNotes(project);
  warmProjectFolderPath(project);
  switchView("project");
}

async function copyProjectFolderPath() {
  if (!state.activeProject || !state.activeProject.projectDir) {
    return;
  }

  try {
    const path = await getProjectFolderPath(state.activeProject);
    await navigator.clipboard.writeText(path);
    showCopyStatus("Copied");
  } catch (error) {
    showCopyStatus("Copy failed");
  }
}

async function warmProjectFolderPath(project) {
  if (!project || project.cachedPath) {
    return;
  }

  try {
    await getProjectFolderPath(project);
  } catch (error) {
    // Ignore cache failures; copy will fallback to recompute.
  }
}

async function getProjectFolderPath(project) {
  if (project.cachedPath) {
    return project.cachedPath;
  }

  if (!project.projectDir) {
    throw new Error("Missing project directory handle.");
  }

  const rootBase = ROOT_BASE_PATH || (state.rootHandle ? state.rootHandle.name : "");
  const expectedParts = getExpectedProjectParts(project);

  if (state.rootHandle && typeof state.rootHandle.resolve === "function") {
    try {
      const resolved = await state.rootHandle.resolve(project.projectDir);
      if (resolved && resolved.length) {
        if (expectedParts) {
          const matchesExpected = expectedParts.every(
            (part, index) =>
              String(resolved[index] || "").toLowerCase() === String(part).toLowerCase(),
          );
          if (!matchesExpected) {
            throw new Error("Resolved path does not match expected project metadata.");
          }
        }

        project.cachedPath = rootBase ? rootBase + "\\" + resolved.join("\\") : resolved.join("\\");
        return project.cachedPath;
      }
    } catch (error) {
      // Fall back to metadata or parent traversal.
    }
  }

  if (expectedParts) {
    project.cachedPath = rootBase ? rootBase + "\\" + expectedParts.join("\\") : expectedParts.join("\\");
    return project.cachedPath;
  }

  const handle = project.projectDir;
  const parts = [handle.name];
  let parent = handle;
  while (parent && parent !== state.rootHandle) {
    parent = await getParentHandle(parent);
    if (parent && parent !== state.rootHandle) {
      parts.unshift(parent.name);
    }
  }

  project.cachedPath = rootBase ? rootBase + "\\" + parts.join("\\") : parts.join("\\");
  return project.cachedPath;
}

function getExpectedProjectParts(project) {
  if (!project || !project.rootName || !project.buildingName) {
    return null;
  }

  const parts = [project.rootName, project.buildingName];
  if (project.isCompleted) {
    parts.push(project.completedFolderName || "Completed");
  }
  parts.push(project.folderName);
  return parts;
}

function showCopyStatus(message) {
  if (!elements.copyStatus) {
    return;
  }

  elements.copyStatus.textContent = message;
  clearTimeout(state.copyStatusTimer);
  state.copyStatusTimer = setTimeout(() => {
    elements.copyStatus.textContent = "";
  }, 2000);
}

async function getParentHandle(childHandle) {
  if (!state.rootHandle) {
    return null;
  }

  for await (const entry of state.rootHandle.values()) {
    if (entry.kind === "directory") {
      if (entry.name === childHandle.name) {
        return state.rootHandle;
      }
      const found = await findParentHandle(entry, childHandle.name);
      if (found) {
        return found;
      }
    }
  }

  return null;
}

async function findParentHandle(dir, targetName) {
  for await (const entry of dir.values()) {
    if (entry.kind === "directory") {
      if (entry.name === targetName) {
        return dir;
      }
      const found = await findParentHandle(entry, targetName);
      if (found) {
        return found;
      }
    }
  }

  return null;
}

function renderProjectFields(project) {
  elements.projectFields.innerHTML = "";
  const copyKeys = new Set(["title", "description", "moc"]);

  FIELD_ORDER.forEach((field) => {
    const wrapper = document.createElement("div");
    wrapper.className = "field";

    const header = document.createElement("div");
    header.className = "field-header";

    const label = document.createElement("label");
    label.textContent = field.label;

    const status = document.createElement("span");
    status.className = "field-copy-status";

    header.append(label, status);

    const input = document.createElement("input");
    input.type = field.type || "text";
    if (field.type === "date") {
      input.value = normalizeDateValue(project.data[field.key]);
    } else {
      input.value = project.data[field.key] ?? "";
    }

    input.addEventListener("input", () => {
      if (input.type === "date") {
        project.data[field.key] = normalizeDateValue(input.value);
      } else {
        const value = input.type === "number" ? Number(input.value) : input.value;
        project.data[field.key] = input.value === "" && input.type === "number" ? "" : value;
      }

      clearTimeout(jsonSaveTimer);
      jsonSaveTimer = setTimeout(() => saveProjectJson(project), 500);
    });

    const inputWrap = document.createElement("div");
    inputWrap.className = "field-input-wrap";
    inputWrap.appendChild(input);

    if (copyKeys.has(field.key)) {
      const copyButton = document.createElement("button");
      copyButton.type = "button";
      copyButton.className = "field-copy";
      copyButton.title = "Copy " + field.label;
      copyButton.innerHTML =
        "<svg viewBox=\"0 0 24 24\" aria-hidden=\"true\" focusable=\"false\"><path d=\"M9 9h10v11H9z\"/><path d=\"M5 5h10v11H5z\"/></svg>";
      copyButton.addEventListener("click", () => {
        copyFieldValue(project.data[field.key] ?? "", status);
      });
      inputWrap.appendChild(copyButton);
    }

    wrapper.append(header, inputWrap);
    elements.projectFields.append(wrapper);
  });
}

function normalizeDateValue(value) {
  if (!value) {
    return "";
  }

  if (typeof value === "string") {
    const trimmed = value.trim();
    if (!trimmed) {
      return "";
    }

    if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) {
      return trimmed;
    }

    const slashMatch = trimmed.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
    if (slashMatch) {
      const month = slashMatch[1].padStart(2, "0");
      const day = slashMatch[2].padStart(2, "0");
      const yearRaw = slashMatch[3];
      const year = yearRaw.length === 2 ? "20" + yearRaw : yearRaw;
      return `${year}-${month}-${day}`;
    }
  }

  return "";
}

async function copyFieldValue(value, statusElement) {
  try {
    await navigator.clipboard.writeText(String(value || ""));
    showFieldCopyStatus(statusElement, "Copied");
  } catch (error) {
    showFieldCopyStatus(statusElement, "Copy failed");
  }
}

function showFieldCopyStatus(statusElement, message) {
  if (!statusElement) {
    return;
  }

  statusElement.textContent = message;
  clearTimeout(statusElement._copyTimer);
  statusElement._copyTimer = setTimeout(() => {
    statusElement.textContent = "";
  }, 1500);
}

function renderNotes(project) {
  const doc = normalizeNotesDoc(project.data.notesDoc);
  if (state.tiptapEditor) {
    state.tiptapEditor.destroy();
    state.tiptapEditor = null;
  }

  elements.notesEditor.innerHTML = "";
  state.tiptapEditor = new Editor({
    element: elements.notesEditor,
    extensions: [
      Details,
      Summary,
      StarterKit.configure({
        heading: { levels: [1, 2, 3] },
      }),
      TaskList,
      CustomTaskItem.configure({ nested: true }),
      Table.configure({ resizable: false }),
      TableRow,
      TableHeader,
      TableCell,
      Image.configure({ allowBase64: true }),
      Link.configure({ openOnClick: false, autolink: true, linkOnPaste: true }),
      Placeholder.configure({ placeholder: "Write project notes..." }),
    ],
    content: doc,
    editorProps: {
      attributes: { class: "notes-prose" },
      handlePaste: (_view, event) => {
        if (!event.clipboardData || !state.tiptapEditor) {
          return false;
        }

        const files = Array.from(event.clipboardData.files || []).filter((file) =>
          file.type.startsWith("image/"),
        );
        if (!files.length) {
          return false;
        }

        event.preventDefault();
        void insertPastedImages(state.tiptapEditor, files);
        return true;
      },
      handleKeyDown: (_view, event) => {
        if (event.key === "/") {
          event.preventDefault();
          openNotesMenu();
          return true;
        }

        if ((event.key === "Backspace" || event.key === "Delete") && state.tiptapEditor) {
          const handled = handleImageDeletion(
            state.tiptapEditor,
            event.key === "Backspace" ? "backward" : "forward",
          );
          if (handled) {
            event.preventDefault();
            return true;
          }
        }

        if (event.key === "Backspace" && state.tiptapEditor) {
          const handled = handleToggleBackspace(state.tiptapEditor);
          if (handled) {
            event.preventDefault();
            return true;
          }
        }

        if (!state.notesMenuActive) {
          return false;
        }

        if (event.key === "Escape") {
          event.preventDefault();
          hideNotesMenu();
          return true;
        }

        if (event.key === "ArrowDown") {
          event.preventDefault();
          moveNotesMenu(1);
          return true;
        }

        if (event.key === "ArrowUp") {
          event.preventDefault();
          moveNotesMenu(-1);
          return true;
        }

        if (event.key === "Enter") {
          event.preventDefault();
          applyNotesMenuSelection();
          return true;
        }

        if (event.key === "Backspace") {
          event.preventDefault();
          state.notesMenuQuery = state.notesMenuQuery.slice(0, -1);
          renderNotesMenu();
          return true;
        }

        if (event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey) {
          event.preventDefault();
          state.notesMenuQuery += event.key.toLowerCase();
          renderNotesMenu();
          return true;
        }

        return false;
      },
      handleDOMEvents: {
        click: (_view, event) => {
          const summary = event.target.closest?.("summary");
          if (summary) {
            const details = summary.closest("details");
            if (details) {
              event.preventDefault();
              toggleDetailsOpen(_view, details);
              return true;
            }
          }

          const link = event.target.closest?.("a");
          if (link && (event.ctrlKey || event.metaKey)) {
            event.preventDefault();
            window.open(link.href, "_blank", "noopener");
            return true;
          }
          return false;
        },
      },
    },
    onUpdate: ({ editor }) => {
      project.data.notesDoc = editor.getJSON();
      clearTimeout(notesSaveTimer);
      notesSaveTimer = setTimeout(() => {
        saveProjectJson(project);
      }, 500);
    },
    onSelectionUpdate: ({ editor }) => {
      updateTableToolbar(editor);
    },
  });
  updateTableToolbar(state.tiptapEditor);
  hideNotesMenu();
}

async function insertPastedImages(editor, files) {
  for (const file of files) {
    const dataUrl = await readFileAsDataUrl(file);
    if (!dataUrl) {
      continue;
    }

    editor.chain().focus().setImage({ src: dataUrl, alt: file.name || "Pasted image" }).run();
  }
}

function readFileAsDataUrl(file) {
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onload = () => resolve(typeof reader.result === "string" ? reader.result : "");
    reader.onerror = () => resolve("");
    reader.readAsDataURL(file);
  });
}

function normalizeNotesDoc(notesDoc) {
  if (!notesDoc) {
    return { type: "doc", content: [] };
  }

  if (typeof notesDoc === "string") {
    try {
      const parsed = JSON.parse(notesDoc);
      return parsed && parsed.type === "doc" ? parsed : { type: "doc", content: [] };
    } catch (error) {
      return { type: "doc", content: [] };
    }
  }

  if (typeof notesDoc === "object" && notesDoc.type === "doc") {
    return notesDoc;
  }

  return { type: "doc", content: [] };
}

async function saveProjectJson(project) {
  try {
    const content = JSON.stringify(project.data, null, "\t") + "\n";
    await writeFile(project.jsonHandle, content);
  } catch (error) {
    elements.rootStatus.textContent = "Unable to save project.json for " + project.folderName;
  }
}

async function saveProjectNotes(project) {
  try {
    const doc = project.htmlDocument || document.implementation.createHTMLDocument("Project Notes");
    doc.body.innerHTML = elements.notesEditor.innerHTML.replace(/\u200B/g, "");
    const content = "<!DOCTYPE html>\n" + doc.documentElement.outerHTML;
    project.htmlText = content;
    await writeFile(project.htmlHandle, content);
  } catch (error) {
    elements.rootStatus.textContent = "Unable to save project.html for " + project.folderName;
  }
}

function parseHtml(htmlText) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(htmlText || "", "text/html");
  return { doc, bodyHtml: doc.body ? doc.body.innerHTML : "" };
}

function switchView(view) {
  const showDashboard = view === "dashboard";
  elements.dashboardView.classList.toggle("is-hidden", !showDashboard);
  elements.projectView.classList.toggle("is-hidden", showDashboard);
  elements.filterGroup.classList.toggle("is-hidden", !showDashboard);
  elements.calendarView.classList.toggle(
    "is-hidden",
    !showDashboard || !CALENDAR_FILTERS.has(state.filter),
  );
  if (!showDashboard) {
    hideSearchResults();
  }
  hideNotesMenu();
}

function renderCalendar() {
  if (!elements.calendarScroll) {
    return;
  }

  const showCompleted = state.filter === "calendar-completed";

  const projects = state.projects
    .filter((project) => (showCompleted ? project.isCompleted : !project.isCompleted))
    .map((project) => {
      const dateValue = normalizeDateValue(project.data.ecDate);
      return {
        project,
        dateValue,
      };
    })
    .filter((item) => item.dateValue);

  elements.calendarScroll.innerHTML = "";

  if (!projects.length) {
    const empty = document.createElement("p");
    empty.className = "empty-state";
    empty.textContent = showCompleted
      ? "No completed projects have an EC Date yet."
      : "No current projects have an EC Date yet.";
    elements.calendarScroll.appendChild(empty);
    return;
  }

  const dateMap = new Map();
  let minDate = null;
  let maxDate = null;

  projects.forEach(({ project, dateValue }) => {
    const [year, month, day] = dateValue.split("-").map((part) => Number(part));
    const date = new Date(year, month - 1, day);
    if (!minDate || date < minDate) {
      minDate = date;
    }
    if (!maxDate || date > maxDate) {
      maxDate = date;
    }

    if (!dateMap.has(dateValue)) {
      dateMap.set(dateValue, []);
    }
    dateMap.get(dateValue).push(project);
  });

  if (!minDate || !maxDate) {
    return;
  }

  const start = new Date(minDate.getFullYear(), minDate.getMonth(), 1);
  const end = new Date(maxDate.getFullYear(), maxDate.getMonth(), 1);

  const monthNames = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
  ];

  const dayNames = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];

  for (let cursor = new Date(start); cursor <= end; cursor.setMonth(cursor.getMonth() + 1)) {
    const monthStart = new Date(cursor.getFullYear(), cursor.getMonth(), 1);
    const monthEnd = new Date(cursor.getFullYear(), cursor.getMonth() + 1, 0);
    const monthCard = document.createElement("div");
    monthCard.className = "calendar-month";

    const title = document.createElement("div");
    title.className = "calendar-month-title";
    title.textContent = `${monthNames[monthStart.getMonth()]} ${monthStart.getFullYear()}`;
    monthCard.appendChild(title);

    const grid = document.createElement("div");
    grid.className = "calendar-grid";

    dayNames.forEach((label) => {
      const header = document.createElement("div");
      header.className = "calendar-day-header";
      header.textContent = label;
      grid.appendChild(header);
    });

    const firstDay = monthStart.getDay();
    const totalCells = firstDay + monthEnd.getDate();
    const rows = Math.ceil(totalCells / 7) * 7;

    for (let index = 0; index < rows; index += 1) {
      const cell = document.createElement("div");
      cell.className = "calendar-day";
      const dayNumber = index - firstDay + 1;
      const inMonth = dayNumber >= 1 && dayNumber <= monthEnd.getDate();
      if (!inMonth) {
        cell.classList.add("is-muted");
        grid.appendChild(cell);
        continue;
      }

      const dateValue = formatDateValue(monthStart.getFullYear(), monthStart.getMonth() + 1, dayNumber);
      const number = document.createElement("div");
      number.className = "calendar-day-number";
      number.textContent = String(dayNumber);
      cell.appendChild(number);

      const items = dateMap.get(dateValue) || [];
      if (items.length) {
        const list = document.createElement("div");
        list.className = "calendar-items";
        items.forEach((project) => {
          const item = document.createElement("div");
          item.className = "calendar-item";
          item.textContent = project.data.description || project.folderName;
          list.appendChild(item);
        });
        cell.appendChild(list);
      }

      grid.appendChild(cell);
    }

    monthCard.appendChild(grid);
    elements.calendarScroll.appendChild(monthCard);
  }
}

function formatDateValue(year, month, day) {
  const mm = String(month).padStart(2, "0");
  const dd = String(day).padStart(2, "0");
  return `${year}-${mm}-${dd}`;
}
function getNotesCommands(editor) {
  if (!editor) {
    return [];
  }

  return [
    {
      key: "body",
      label: "Body",
      description: "Normal text",
      action: () => editor.chain().focus().setParagraph().run(),
    },
    {
      key: "heading 1",
      label: "Heading 1",
      description: "Large section heading",
      action: () => editor.chain().focus().toggleHeading({ level: 1 }).run(),
    },
    {
      key: "heading 2",
      label: "Heading 2",
      description: "Medium section heading",
      action: () => editor.chain().focus().toggleHeading({ level: 2 }).run(),
    },
    {
      key: "heading 3",
      label: "Heading 3",
      description: "Small section heading",
      action: () => editor.chain().focus().toggleHeading({ level: 3 }).run(),
    },
    {
      key: "bulleted list",
      label: "Bulleted list",
      description: "Add bullet points",
      action: () => editor.chain().focus().toggleBulletList().run(),
    },
    {
      key: "numbered list",
      label: "Numbered list",
      description: "Add a numbered list",
      action: () => editor.chain().focus().toggleOrderedList().run(),
    },
    {
      key: "toggle",
      label: "Toggle",
      description: "Expandable section",
      action: () => insertToggle(editor, 0),
    },
    {
      key: "toggle heading 1",
      label: "Toggle Heading 1",
      description: "Large toggle heading",
      action: () => insertToggle(editor, 1),
    },
    {
      key: "toggle heading 2",
      label: "Toggle Heading 2",
      description: "Medium toggle heading",
      action: () => insertToggle(editor, 2),
    },
    {
      key: "toggle heading 3",
      label: "Toggle Heading 3",
      description: "Small toggle heading",
      action: () => insertToggle(editor, 3),
    },
    {
      key: "table",
      label: "Table",
      description: "Insert a simple table",
      action: () => editor.chain().focus().insertTable({ rows: 3, cols: 2, withHeaderRow: true }).run(),
    },
    {
      key: "checkbox",
      label: "Checkbox list",
      description: "Checklist item",
      action: () => editor.chain().focus().toggleTaskList().run(),
    },
    {
      key: "link",
      label: "Link",
      description: "Insert a hyperlink",
      action: () => insertLink(editor),
    },
  ];
}

function insertToggle(editor, level) {
  if (!editor) {
    return;
  }

  editor.chain().focus().insertContent({
    type: "details",
    attrs: { open: true },
    content: [
      {
        type: "summary",
        attrs: { level },
        content: [{ type: "text", text: "Toggle" }],
      },
      { type: "paragraph" },
    ],
  }).run();
}

function insertLink(editor) {
  const selection = editor.state.selection;
  const hasSelection = selection && !selection.empty;
  const url = window.prompt("Enter a URL", "https://");
  if (!url) {
    return;
  }

  if (!hasSelection) {
    const text = window.prompt("Link text", "Link") || "Link";
    const start = editor.state.selection.from;
    editor.chain().focus().insertContent(text).run();
    editor.commands.setTextSelection({ from: start, to: start + text.length });
  }

  editor.chain().focus().setLink({ href: url }).run();
}

function openNotesMenu() {
  state.notesMenuActive = true;
  state.notesMenuQuery = "";
  state.notesMenuIndex = 0;
  renderNotesMenu();
}

function renderNotesMenu() {
  if (!state.notesMenuActive || !state.tiptapEditor) {
    elements.notesMenu.classList.add("is-hidden");
    return;
  }

  const query = state.notesMenuQuery.trim();
  const commands = getNotesCommands(state.tiptapEditor);
  const results = commands.filter((cmd) => cmd.key.includes(query));
  state.notesMenuResults = results;
  state.notesMenuIndex = Math.min(state.notesMenuIndex, Math.max(0, results.length - 1));

  elements.notesMenu.innerHTML = "";
  if (!results.length) {
    const empty = document.createElement("div");
    empty.className = "notes-menu-item";
    empty.textContent = "No commands.";
    elements.notesMenu.appendChild(empty);
    elements.notesMenu.classList.remove("is-hidden");
    return;
  }

  results.forEach((cmd, index) => {
    const item = document.createElement("div");
    item.className = "notes-menu-item" + (index === state.notesMenuIndex ? " is-active" : "");

    const title = document.createElement("div");
    title.className = "notes-menu-title";
    title.textContent = cmd.label;

    const desc = document.createElement("div");
    desc.className = "notes-menu-desc";
    desc.textContent = cmd.description;

    item.append(title, desc);
    item.addEventListener("click", () => {
      cmd.action();
      hideNotesMenu();
    });

    elements.notesMenu.appendChild(item);
  });

  elements.notesMenu.classList.remove("is-hidden");
}

function moveNotesMenu(delta) {
  const results = state.notesMenuResults || [];
  if (!results.length) {
    return;
  }

  state.notesMenuIndex = Math.max(0, Math.min(state.notesMenuIndex + delta, results.length - 1));
  renderNotesMenu();
}

function applyNotesMenuSelection() {
  const results = state.notesMenuResults || [];
  const selected = results[state.notesMenuIndex];
  if (selected) {
    selected.action();
  }
  hideNotesMenu();
}

function hideNotesMenu() {
  state.notesMenuActive = false;
  state.notesMenuQuery = "";
  state.notesMenuIndex = 0;
  state.notesMenuResults = [];
  elements.notesMenu.classList.add("is-hidden");
  elements.notesMenu.innerHTML = "";
}

function toggleDetailsOpen(view, details) {
  if (!view || !details) {
    return;
  }

  const target = details.querySelector("summary") || details;
  const domPos = view.posAtDOM(target, 0);
  if (domPos == null) {
    return;
  }

  const $pos = view.state.doc.resolve(domPos);
  const detailsInfo = findParentNode($pos, "details");
  if (!detailsInfo) {
    return;
  }

  const nextAttrs = { ...detailsInfo.node.attrs, open: !detailsInfo.node.attrs.open };
  view.dispatch(view.state.tr.setNodeMarkup(detailsInfo.pos, undefined, nextAttrs));
}

function findParentNode($pos, typeName) {
  for (let depth = $pos.depth; depth > 0; depth -= 1) {
    const node = $pos.node(depth);
    if (node.type.name === typeName) {
      return { node, pos: $pos.before(depth), depth };
    }
  }
  return null;
}

function getTaskItemInfo(editor) {
  if (!editor) {
    return null;
  }

  const { selection } = editor.state;
  if (!selection.empty) {
    return null;
  }

  return findParentNode(selection.$from, "taskItem");
}

function isTaskItemEmpty(info) {
  return !info || !info.node || !info.node.textContent.trim();
}

function isCursorAtTaskItemStart(editor, info) {
  if (!editor || !info || typeof info.depth !== "number") {
    return false;
  }

  const $from = editor.state.selection.$from;
  const start = $from.start(info.depth);
  return $from.pos === start;
}

function updateTableToolbar(editor) {
  if (!elements.tableToolbar) {
    return;
  }

  const isTableActive = !!editor && editor.isActive("table");
  elements.tableToolbar.classList.toggle("is-hidden", !isTableActive);
}

function handleImageDeletion(editor, direction) {
  if (!editor) {
    return false;
  }

  const { state: editorState, view } = editor;
  const { selection } = editorState;
  const imageType = editorState.schema.nodes.image;
  if (!imageType) {
    return false;
  }

  if (selection.node && selection.node.type === imageType) {
    view.dispatch(editorState.tr.deleteSelection());
    return true;
  }

  if (!selection.empty) {
    return false;
  }

  const { $from } = selection;
  if (direction === "backward") {
    const nodeBefore = $from.nodeBefore;
    if (nodeBefore && nodeBefore.type === imageType) {
      const from = $from.pos - nodeBefore.nodeSize;
      view.dispatch(editorState.tr.delete(from, $from.pos));
      return true;
    }
  } else {
    const nodeAfter = $from.nodeAfter;
    if (nodeAfter && nodeAfter.type === imageType) {
      const to = $from.pos + nodeAfter.nodeSize;
      view.dispatch(editorState.tr.delete($from.pos, to));
      return true;
    }
  }

  return false;
}

function handleToggleBackspace(editor) {
  const { state: editorState, view } = editor;
  const { selection } = editorState;
  if (!selection.empty) {
    return false;
  }

  const $from = selection.$from;
  const detailsInfo = findParentNode($from, "details");
  if (detailsInfo && $from.parent.type.name === "summary" && $from.parentOffset === 0) {
    const tr = editorState.tr.delete(detailsInfo.pos, detailsInfo.pos + detailsInfo.node.nodeSize);
    view.dispatch(tr);
    return true;
  }

  if ($from.parentOffset === 0) {
    const beforePos = $from.before($from.depth);
    if (beforePos > 0) {
      const prevNode = editorState.doc.nodeAt(beforePos - 1);
      if (prevNode && prevNode.type.name === "details") {
        const tr = editorState.tr.delete(beforePos - prevNode.nodeSize, beforePos);
        view.dispatch(tr);
        return true;
      }
    }
  }

  return false;
}

function getCurrentBlock() {
  const selection = window.getSelection();
  if (!selection || !selection.rangeCount) {
    return null;
  }

  let node = selection.anchorNode;
  if (!node) {
    return null;
  }

  if (node.nodeType === Node.TEXT_NODE) {
    node = node.parentElement;
  }

  if (!node) {
    return null;
  }

  return node.closest("p, div, li, h1, h2, h3") || node;
}

function replaceBlock(block, newElement) {
  if (!block || !block.parentNode) {
    return;
  }

  block.parentNode.replaceChild(newElement, block);
  placeCaretAtEnd(newElement);
}

function applyHeading(level) {
  const block = getCurrentBlock();
  if (!block) {
    return;
  }

  const heading = document.createElement("h" + level);
  heading.textContent = stripSlashCommand(block.textContent);
  replaceBlock(block, heading);
}

function applyParagraph() {
  const block = getCurrentBlock();
  if (!block) {
    return;
  }

  const paragraph = document.createElement("p");
  paragraph.textContent = stripSlashCommand(block.textContent);
  replaceBlock(block, paragraph);
}

function applyList(type) {
  const block = getCurrentBlock();
  if (!block) {
    return;
  }

  const list = document.createElement(type);
  const item = document.createElement("li");
  item.textContent = stripSlashCommand(block.textContent);
  list.appendChild(item);
  replaceBlock(block, list);
}

function applyChecklist() {
  const block = getCurrentBlock();
  if (!block) {
    return;
  }

  const cell = getClosest(block, "td, th");
  if (cell) {
    insertCheckboxInCell(cell);
    return;
  }

  const list = document.createElement("ul");
  list.className = "checklist";
  const item = document.createElement("li");
  const checkbox = document.createElement("input");
  checkbox.type = "checkbox";
  const text = document.createElement("span");
  text.textContent = stripSlashCommand(block.textContent);
  item.append(checkbox, text);
  list.appendChild(item);
  replaceBlock(block, list);
}

function applyTable() {
  const block = getCurrentBlock();
  if (!block) {
    return;
  }

  const tableBlock = createTableBlock();
  replaceBlock(block, tableBlock);
  const firstCell = tableBlock.querySelector("td");
  if (firstCell) {
    placeCaretAtStart(firstCell);
  }
}

function createTableBlock() {
  const wrapper = document.createElement("div");
  wrapper.className = "table-block";
  wrapper.contentEditable = "false";

  const table = document.createElement("table");
  table.className = "notes-table";
  table.contentEditable = "false";

  const thead = document.createElement("thead");
  const headRow = document.createElement("tr");
  for (let i = 0; i < 3; i += 1) {
    const th = document.createElement("th");
    th.contentEditable = "true";
    th.innerHTML = "Header";
    headRow.appendChild(th);
  }
  thead.appendChild(headRow);

  const tbody = document.createElement("tbody");
  for (let r = 0; r < 2; r += 1) {
    tbody.appendChild(createTableRow(3));
  }

  table.append(thead, tbody);

  const controls = document.createElement("div");
  controls.className = "table-controls";
  controls.contentEditable = "false";

  const addRow = document.createElement("button");
  addRow.type = "button";
  addRow.className = "table-add-row";
  addRow.textContent = "+ Row";

  const addCol = document.createElement("button");
  addCol.type = "button";
  addCol.className = "table-add-col";
  addCol.textContent = "+ Column";

  const delRow = document.createElement("button");
  delRow.type = "button";
  delRow.className = "table-del-row";
  delRow.textContent = "- Row";

  const delCol = document.createElement("button");
  delCol.type = "button";
  delCol.className = "table-del-col";
  delCol.textContent = "- Column";

  controls.append(addRow, addCol, delRow, delCol);
  wrapper.append(table, controls);

  return wrapper;
}

function createTableRow(columnCount) {
  const row = document.createElement("tr");
  for (let i = 0; i < columnCount; i += 1) {
    const td = document.createElement("td");
    td.contentEditable = "true";
    td.innerHTML = "<br>";
    row.appendChild(td);
  }
  return row;
}

function addTableRow(table) {
  if (!table) {
    return;
  }

  const body = table.querySelector("tbody");
  const firstRow = body ? body.querySelector("tr") : null;
  const columnCount = firstRow ? firstRow.children.length : 3;
  if (body) {
    body.appendChild(createTableRow(columnCount));
  }
}

function addTableColumn(table) {
  if (!table) {
    return;
  }

  const rows = table.querySelectorAll("tr");
  rows.forEach((row, index) => {
    const cell = document.createElement(index === 0 ? "th" : "td");
    cell.contentEditable = "true";
    cell.innerHTML = index === 0 ? "Header" : "<br>";
    row.appendChild(cell);
  });
}

function deleteTableRow() {
  const cell = getActiveTableCell();
  if (!cell) {
    return;
  }

  const row = cell.parentElement;
  const tbody = row.parentElement;
  if (!tbody || tbody.tagName.toLowerCase() !== "tbody") {
    return;
  }

  if (tbody.children.length <= 1) {
    row.querySelectorAll("td").forEach((td) => {
      td.innerHTML = "<br>";
    });
    placeCaretAtStart(row.querySelector("td"));
    return;
  }

  const next = row.nextElementSibling || row.previousElementSibling;
  tbody.removeChild(row);
  if (next) {
    placeCaretAtStart(next.querySelector("td"));
  }
}

function deleteTableColumn() {
  const cell = getActiveTableCell();
  if (!cell) {
    return;
  }

  const row = cell.parentElement;
  const table = row.closest("table");
  if (!table) {
    return;
  }

  const index = Array.from(row.children).indexOf(cell);
  if (index < 0) {
    return;
  }

  const rows = table.querySelectorAll("tr");
  if (!rows.length || rows[0].children.length <= 1) {
    rows.forEach((r) => {
      const target = r.children[0];
      if (target) {
        target.innerHTML = r.parentElement.tagName.toLowerCase() === "thead" ? "Header" : "<br>";
      }
    });
    return;
  }

  rows.forEach((r) => {
    if (r.children[index]) {
      r.removeChild(r.children[index]);
    }
  });
}

function getActiveTableCell() {
  const selection = window.getSelection();
  if (selection && selection.rangeCount) {
    const node = selection.anchorNode;
    const cell = getClosest(node, "td, th");
    if (cell) {
      state.activeTableCell = cell;
      return cell;
    }
  }

  return state.activeTableCell;
}

function insertCheckboxInCell(cell) {
  const checkbox = document.createElement("input");
  checkbox.type = "checkbox";
  const spacer = document.createTextNode(" ");
  insertNodeAtCaret(checkbox);
  insertNodeAtCaret(spacer);
}

function applyToggle(level) {
  const block = getCurrentBlock();
  if (!block) {
    return;
  }

  const details = document.createElement("details");
  const summary = document.createElement("summary");
  if (level) {
    summary.classList.add("toggle-heading", "toggle-heading-" + level);
  }
  summary.textContent = stripSlashCommand(block.textContent) || "Toggle";
  const body = document.createElement("p");
  body.textContent = "";
  details.append(summary, body);
  replaceBlock(block, details);
  ensureParagraphAfter(details);
  placeCaretAtEnd(summary);
}

function focusToggleBody(details) {
  const body = ensureToggleBody(details);
  details.open = true;
  if (body) {
    placeCaretAtStart(body);
  }
}

function ensureToggleBody(details) {
  if (!details) {
    return null;
  }

  let body = details.querySelector("p");
  if (!body) {
    body = document.createElement("p");
    body.innerHTML = "<br>";
    details.appendChild(body);
  } else if (!body.innerHTML.trim()) {
    body.innerHTML = "<br>";
  }

  return body;
}

function normalizeToggleBodies() {
  const toggles = elements.notesEditor.querySelectorAll("details");
  toggles.forEach((details) => {
    ensureToggleBody(details);
  });
}

function isToggleMarkerClick(event, summary) {
  if (!summary) {
    return false;
  }

  const rect = summary.getBoundingClientRect();
  const markerWidth = 22;
  return event.clientX <= rect.left + markerWidth;
}

function applyLink() {
  const selection = window.getSelection();
  const selectedText = selection && selection.toString ? selection.toString() : "";
  const url = window.prompt("Enter a URL", "https://");
  if (!url) {
    return;
  }

  const text = selectedText || window.prompt("Link text", "Link") || "Link";
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.textContent = text;
  anchor.target = "_blank";
  anchor.rel = "noopener";
  insertNodeAtCaret(anchor);
}

function getChecklistItemFromSelection() {
  const selection = window.getSelection();
  if (!selection || !selection.rangeCount) {
    return null;
  }

  const node = selection.anchorNode;
  return getClosest(node, "ul.checklist li");
}

function createChecklistItem(text) {
  const item = document.createElement("li");
  const checkbox = document.createElement("input");
  checkbox.type = "checkbox";
  const span = document.createElement("span");
  if (text) {
    span.textContent = text;
  } else {
    span.textContent = "\u200B";
    span.dataset.placeholder = "true";
  }
  item.append(checkbox, span);
  return item;
}

function exitChecklist(item) {
  if (!item || !item.parentNode) {
    return;
  }

  const list = item.parentNode;
  const parent = list.parentNode;
  const paragraph = document.createElement("p");
  paragraph.innerHTML = "<br>";
  parent.insertBefore(paragraph, list.nextSibling);
  list.removeChild(item);
  if (!list.children.length) {
    parent.removeChild(list);
  }
  placeCaretAtStart(paragraph);
}

function focusChecklistSpan(span) {
  if (!span) {
    return;
  }

  if (!span.textContent) {
    span.textContent = "\u200B";
    span.dataset.placeholder = "true";
  }

  const textNode = span.firstChild;
  placeCaretAtTextEnd(textNode || span);
}

function cleanupChecklistPlaceholders() {
  const placeholders = elements.notesEditor.querySelectorAll("span[data-placeholder='true']");
  placeholders.forEach((span) => {
    if (span.textContent !== "\u200B") {
      span.removeAttribute("data-placeholder");
    }
  });
}

function stripSlashCommand(text) {
  return (text || "").replace(/^\/[\w-]+\s*/i, "");
}

function applyBlockMarkdown() {
  const block = getCurrentBlock();
  if (!block) {
    return;
  }

  const text = block.textContent || "";
  if (text.startsWith("# ")) {
    const heading = document.createElement("h1");
    heading.textContent = text.replace(/^#\s+/, "");
    replaceBlock(block, heading);
    return;
  }

  if (text.startsWith("## ")) {
    const heading = document.createElement("h2");
    heading.textContent = text.replace(/^##\s+/, "");
    replaceBlock(block, heading);
    return;
  }

  if (text.startsWith("### ")) {
    const heading = document.createElement("h3");
    heading.textContent = text.replace(/^###\s+/, "");
    replaceBlock(block, heading);
    return;
  }

  if (/^(\[\]|\[\s\])\s+/.test(text)) {
    const content = text.replace(/^(\[\]|\[\s\])\s+/, "");
    const list = document.createElement("ul");
    list.className = "checklist";
    const item = document.createElement("li");
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    const span = document.createElement("span");
    span.textContent = content;
    item.append(checkbox, span);
    list.appendChild(item);
    replaceBlock(block, list);
    return;
  }

  if (text.startsWith("- [ ] ") || text.startsWith("* [ ] ")) {
    const content = text.replace(/^[-*]\s\[\s\]\s+/, "");
    const list = document.createElement("ul");
    list.className = "checklist";
    const item = document.createElement("li");
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    const span = document.createElement("span");
    span.textContent = content;
    item.append(checkbox, span);
    list.appendChild(item);
    replaceBlock(block, list);
    return;
  }

  if (text.startsWith("- ") || text.startsWith("* ")) {
    const content = text.replace(/^[-*]\s+/, "");
    const list = document.createElement("ul");
    const item = document.createElement("li");
    item.textContent = content;
    list.appendChild(item);
    replaceBlock(block, list);
    return;
  }

  if (/^\d+\.\s+/.test(text)) {
    const content = text.replace(/^\d+\.\s+/, "");
    const list = document.createElement("ol");
    const item = document.createElement("li");
    item.textContent = content;
    list.appendChild(item);
    replaceBlock(block, list);
  }
}

function applyInlineMarkdown() {
  const block = getCurrentBlock();
  if (!block) {
    return;
  }

  const selection = window.getSelection();
  if (!selection || !selection.rangeCount || !selection.isCollapsed) {
    return;
  }

  if (!isSelectionInside(selection, block)) {
    return;
  }

  const text = block.textContent || "";
  if (!text.includes("*")) {
    return;
  }

  if (!hasBalancedInlineMarkers(text)) {
    return;
  }

  if (!isCaretAtEndOfBlock(selection, block)) {
    return;
  }

  const html = block.innerHTML;
  if (!html.includes("*")) {
    return;
  }

  const replaced = html
    .replace(/\*\*([^*]+)\*\*/g, "<strong>$1</strong>")
    .replace(/(^|[^*])\*([^*]+)\*(?!\*)/g, "$1<em>$2</em>");

  if (replaced !== html) {
    block.innerHTML = replaced;
    const trailing = ensureTrailingText(block);
    placeCaretAtTextEnd(trailing);
  }
}

function ensureTrailingText(block) {
  const last = block.lastChild;
  if (last && last.nodeType === Node.TEXT_NODE) {
    if (!last.textContent.endsWith(" ")) {
      last.textContent += " ";
    }
    return last;
  }

  const textNode = document.createTextNode(" ");
  block.appendChild(textNode);
  return textNode;
}

function placeCaretAtTextEnd(textNode) {
  if (!textNode) {
    return;
  }

  const range = document.createRange();
  const length = textNode.textContent.length;
  range.setStart(textNode, length);
  range.setEnd(textNode, length);
  const selection = window.getSelection();
  selection.removeAllRanges();
  selection.addRange(range);
}

function ensureParagraphAfter(element) {
  if (!element || !element.parentNode) {
    return;
  }

  const next = element.nextElementSibling;
  if (!next) {
    const paragraph = document.createElement("p");
    paragraph.innerHTML = "<br>";
    element.parentNode.appendChild(paragraph);
  }
}

function removeDetails(details) {
  if (!details || !details.parentNode) {
    return;
  }

  const next = details.nextElementSibling;
  const parent = details.parentNode;
  parent.removeChild(details);
  if (next) {
    placeCaretAtStart(next);
  } else {
    const paragraph = document.createElement("p");
    paragraph.innerHTML = "<br>";
    parent.appendChild(paragraph);
    placeCaretAtStart(paragraph);
  }
}

function placeCaretAtStart(element) {
  const range = document.createRange();
  range.selectNodeContents(element);
  range.collapse(true);
  const selection = window.getSelection();
  selection.removeAllRanges();
  selection.addRange(range);
}

function getClosest(node, selector) {
  if (!node) {
    return null;
  }

  const element = node.nodeType === Node.TEXT_NODE ? node.parentElement : node;
  if (!element) {
    return null;
  }

  return element.closest(selector);
}

function isCaretAtStart(range, block) {
  if (!range || !block) {
    return false;
  }

  if (range.startContainer === block && range.startOffset === 0) {
    return true;
  }

  if (range.startContainer.nodeType === Node.TEXT_NODE) {
    return range.startContainer.parentElement === block && range.startOffset === 0;
  }

  return false;
}

function isCaretAtStartOfSpan(selection, span) {
  if (!selection || !selection.rangeCount || !span) {
    return false;
  }

  const range = selection.getRangeAt(0);
  if (range.startContainer === span && range.startOffset === 0) {
    return true;
  }

  if (range.startContainer.nodeType === Node.TEXT_NODE) {
    return range.startContainer.parentElement === span && range.startOffset === 0;
  }

  return false;
}

function isSelectionInside(selection, block) {
  const range = selection.getRangeAt(0);
  const container = range.startContainer;
  if (!container) {
    return false;
  }

  const element = container.nodeType === Node.TEXT_NODE ? container.parentElement : container;
  return element ? block.contains(element) : false;
}

function isCaretAtEndOfBlock(selection, block) {
  const range = selection.getRangeAt(0);
  const preRange = range.cloneRange();
  preRange.selectNodeContents(block);
  preRange.setEnd(range.endContainer, range.endOffset);
  const offset = preRange.toString().length;
  const total = (block.textContent || "").length;
  return offset === total;
}

function hasBalancedInlineMarkers(text) {
  const doubleMatches = text.match(/\*\*/g) || [];
  if (doubleMatches.length % 2 !== 0) {
    return false;
  }

  const withoutDouble = text.replace(/\*\*/g, "");
  const singleMatches = withoutDouble.match(/\*/g) || [];
  return singleMatches.length % 2 === 0;
}

function insertNodeAtCaret(node) {
  const selection = window.getSelection();
  if (!selection || !selection.rangeCount) {
    elements.notesEditor.appendChild(node);
    return;
  }

  const range = selection.getRangeAt(0);
  range.deleteContents();
  range.insertNode(node);
  range.setStartAfter(node);
  range.setEndAfter(node);
  selection.removeAllRanges();
  selection.addRange(range);
}

function placeCaretAtEnd(element) {
  const range = document.createRange();
  range.selectNodeContents(element);
  range.collapse(false);
  const selection = window.getSelection();
  selection.removeAllRanges();
  selection.addRange(range);
}

async function readFile(handle) {
  const file = await handle.getFile();
  return file.text();
}

async function writeFile(handle, content) {
  const writable = await handle.createWritable();
  await writable.write(content);
  await writable.close();
}

async function tryGetDirectoryHandle(parent, name) {
  try {
    return await parent.getDirectoryHandle(name);
  } catch (error) {
    return null;
  }
}

async function tryGetFileHandle(parent, name) {
  try {
    return await parent.getFileHandle(name);
  } catch (error) {
    return null;
  }
}

function formatPercent(value) {
  if (value === undefined || value === null || value === "") {
    return "";
  }

  const numberValue = Number(value);
  if (Number.isNaN(numberValue)) {
    return String(value);
  }

  return numberValue + "%";
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll("\"", "&quot;")
    .replaceAll("'", "&#39;");
}

async function verifyPermission(handle, readwrite) {
  const options = { mode: readwrite ? "readwrite" : "read" };
  if ((await handle.queryPermission(options)) === "granted") {
    return true;
  }

  if ((await handle.requestPermission(options)) === "granted") {
    return true;
  }

  return false;
}

async function openDb() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open("project-dashboard", 1);
    request.onupgradeneeded = () => request.result.createObjectStore("handles");
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function saveRootHandle(handle) {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction("handles", "readwrite");
    tx.objectStore("handles").put(handle, "root");
    tx.oncomplete = () => resolve(true);
    tx.onerror = () => reject(tx.error);
  });
}

async function loadRootHandle() {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction("handles", "readonly");
    const request = tx.objectStore("handles").get("root");
    request.onsuccess = () => resolve(request.result || null);
    request.onerror = () => reject(request.error);
  });
}
