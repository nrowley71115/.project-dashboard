# Project Dashboard

A lightweight, browser-based dashboard for browsing, editing, and tracking project folders stored under the main Projects directory. It uses the File System Access API to read and update `project.json` files directly from the filesystem.

## Folder Structure

Expected layout under the selected root folder:

```
<Projects Root>/
  EI/
    B25/
      Completed/
      <Project Folder>/
  SCP/
    B43/
      Completed/
      <Project Folder>/
  SER/
    B107/
      <Project Folder>/
  WO/
    B55/
      Completed/
      <Project Folder>/
```

Files inside each project folder:

```
<Project Folder>/
  project.json
  (optional files and subfolders)
```

Notes:
- The root type folders must match `ROOT_FOLDERS` in `app.js` (EI, SCP, SER, WO).
- Completed projects are detected by a `Completed/` or `Complete/` subfolder.

## Main Features

- **Project list + filters**: Current vs Completed projects, plus a calendar view based on EC Date.
- **Search**: Quick search across title, description, and folder name.
- **Project details editor**: Edits `project.json` fields in place with autosave.
- **Notes editor**: Tiptap-based rich text editor with headings, lists, tables, toggles, task lists, links, and pasted images.
- **Engineering Second Brain**: Standalone wiki-style knowledge page (`second-brain.html`) with topic navigation and search for standards, instrumentation, technology notes, power, and field references.
- **Other Tasks workspace**: Separate locally stored task dashboard for non-project work with quick entry, notes, target dates, and completion tracking.
- **Copy project folder path**: One-click copy of the project directory path, or use **Ctrl+M** from a project detail page. `Ctrl+N` remains available for opening a new browser window.
- **SharePoint shortcut**: Opens the derived SharePoint folder for supported job numbers and project types.

## How It Works

- On load, you select a root folder with the File System Access API.
- The app scans `ROOT_FOLDERS` for building folders, then reads `project.json` from each project directory.
- Data is cached in memory and rendered in the dashboard table or calendar view.
- Editing a field updates the in-memory data and writes back to `project.json` after a short debounce.
- The notes editor stores its content in `project.json` under `notesDoc`.
- The Other Tasks workspace stores its data in browser `localStorage`, separate from project folders.
- The Second Brain is a normal static HTML page that you edit directly when you want to add links or notes.
- The copy button builds the project path from the known root/building metadata and uses `resolve()` when available.
- The SharePoint button derives a folder URL from the project job number (`title` in current data, or `jobNo` if present) plus the project type.

## How To Run

This project is static HTML, CSS, and JS. You can run it directly in a browser that supports the File System Access API (Chromium-based browsers).

Option A: Open the file directly
1. Open `project-dashboard/index.html` in Chrome or Edge.
2. Click **Select Projects Folder** and choose the main Projects root.

Option B: Use a local static server (recommended)
1. Start any static server in the `project-dashboard` folder.
2. Open the served URL in Chrome or Edge.
3. Click **Select Projects Folder** and choose the main Projects root.
4. Open **Engineering Second Brain** from the header and edit `second-brain.html` directly as needed.

## Replicating This Setup

1. Copy the `project-dashboard` folder to your desired location.
2. Ensure your Projects root follows the expected folder structure.
3. Update `ROOT_BASE_PATH` in `app.js` if you want the copy-path feature to use a different base path.
4. Open `index.html` in a Chromium-based browser and select the Projects root folder.

## Key Files

- `index.html` - Main UI layout.
- `second-brain.html` - Engineering knowledge landing page.
- `other-tasks.html` - Quick-capture dashboard for non-project tasks.
- `styles.css` - Styling for dashboard, editor, and controls.
- `app.js` - Application logic, filesystem access, and editor setup.
- `other-tasks.js` - Local task storage, filtering, and detail editing logic.

## Browser Requirements

- Chrome or Edge with File System Access API support.
- Local file access permission granted when selecting the Projects root.
