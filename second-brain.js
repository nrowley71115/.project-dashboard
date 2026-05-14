const searchInput = document.getElementById("brain-search");
const clearButton = document.getElementById("brain-search-clear");
const searchStatus = document.getElementById("brain-search-status");
const emptyState = document.getElementById("wiki-empty-state");
const sections = Array.from(document.querySelectorAll(".wiki-section"));
const navLinks = Array.from(document.querySelectorAll(".wiki-nav a"));

initSecondBrainSearch();

function initSecondBrainSearch() {
  if (!searchInput || !clearButton || !searchStatus || !emptyState || !sections.length) {
    return;
  }

  searchInput.addEventListener("input", applySearch);
  clearButton.addEventListener("click", () => {
    searchInput.value = "";
    applySearch();
    searchInput.focus();
  });

  document.addEventListener("keydown", (event) => {
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "k") {
      event.preventDefault();
      searchInput.focus();
      searchInput.select();
    }
  });

  applySearch();
}

function applySearch() {
  const query = searchInput.value.trim().toLowerCase();
  let visibleSections = 0;
  let visibleEntries = 0;

  sections.forEach((section) => {
    const heading = section.querySelector(".wiki-section-heading");
    const entryGrid = section.querySelector(".wiki-entry-grid");
    const entries = Array.from(section.querySelectorAll(".wiki-entry"));
    const headingText = heading ? heading.textContent.toLowerCase() : "";
    const sectionMatches = !query || headingText.includes(query);

    let sectionHasVisibleEntry = false;

    entries.forEach((entry) => {
      const entryMatches = !query || sectionMatches || entry.textContent.toLowerCase().includes(query);
      entry.classList.toggle("is-hidden", !entryMatches);
      if (entryMatches) {
        sectionHasVisibleEntry = true;
        visibleEntries += 1;
      }
    });

    const showSection = sectionMatches || sectionHasVisibleEntry;
    section.classList.toggle("is-hidden", !showSection);
    if (entryGrid) {
      entryGrid.classList.toggle("wiki-entry-grid-single", showSection && visibleEntries === 1 && query.length > 0);
    }
    if (showSection) {
      visibleSections += 1;
    }
  });

  updateNavVisibility(query);
  emptyState.classList.toggle("is-hidden", visibleSections > 0);

  if (!query) {
    searchStatus.textContent = "Press Ctrl+K to jump here quickly.";
    return;
  }

  searchStatus.textContent = `${visibleEntries} note card${visibleEntries === 1 ? "" : "s"} across ${visibleSections} topic${visibleSections === 1 ? "" : "s"}`;
}

function updateNavVisibility(query) {
  navLinks.forEach((link) => {
    const targetId = link.getAttribute("href")?.replace("#", "");
    const section = targetId ? document.getElementById(targetId) : null;
    const visible = !query || (section && !section.classList.contains("is-hidden"));
    link.classList.toggle("is-hidden", !visible);
  });
}