/**
 * STUDIO SQL - MINIMALIST DATABASE & FILE MANAGEMENT ENGINE
 * Features: Database Prompt Setup, Basic CRUD (Tables/Rows/Folders/Files),
 * Drag & Drop Files to Any Folder with Glowing Hover Animations, Database Saving & Viewer.
 */

// GLOBAL APPLICATION STATE
const state = {
  dbName: 'MyDatabase.sqlite',
  activeView: 'table-contents', // 'table-contents' | 'file-preview' | 'sql-query'
  activeTable: 'Users',
  activeFileId: null,
  activeFolderId: null,
  
  // Open Tabs Management
  tabs: [
    { id: 'table:Users', type: 'table', name: 'Users', icon: 'fa-table' }
  ],
  activeTabId: 'table:Users',

  // Database Schemas & Records Store
  tables: {
    'Users': {
      columns: [
        { name: 'id', type: 'INTEGER', pk: true, nn: true },
        { name: 'name', type: 'TEXT', pk: false, nn: true },
        { name: 'email', type: 'TEXT', pk: false, nn: true },
        { name: 'role', type: 'TEXT', pk: false, nn: false },
        { name: 'created_at', type: 'TEXT', pk: false, nn: false }
      ],
      rows: [
        { id: 1, name: 'John Doe', email: 'john@example.com', role: 'Admin', created_at: '2026-01-10' },
        { id: 2, name: 'Maria Garcia', email: 'maria@tech.org', role: 'Developer', created_at: '2026-01-15' },
        { id: 3, name: 'Alex Chen', email: 'alex@data.io', role: 'Analyst', created_at: '2026-02-01' }
      ]
    },
    'Orders': {
      columns: [
        { name: 'id', type: 'INTEGER', pk: true, nn: true },
        { name: 'user_id', type: 'INTEGER', pk: false, nn: true },
        { name: 'amount', type: 'REAL', pk: false, nn: true },
        { name: 'status', type: 'TEXT', pk: false, nn: false }
      ],
      rows: [
        { id: 101, user_id: 1, amount: 149.99, status: 'Completed' },
        { id: 102, user_id: 2, amount: 299.50, status: 'Pending' }
      ]
    },
    'Products': {
      columns: [
        { name: 'id', type: 'INTEGER', pk: true, nn: true },
        { name: 'title', type: 'TEXT', pk: false, nn: true },
        { name: 'price', type: 'REAL', pk: false, nn: true },
        { name: 'stock', type: 'INTEGER', pk: false, nn: false }
      ],
      rows: [
        { id: 501, title: 'Studio SQL Pro', price: 49.00, stock: 100 },
        { id: 502, title: 'Data Engine Plugin', price: 19.99, stock: 250 }
      ]
    }
  },

  // Folders & Files Store
  folders: {
    'f_media': {
      id: 'f_media',
      name: 'Media',
      icon: 'fa-folder',
      files: []
    },
    'f_documents': {
      id: 'f_documents',
      name: 'Documents',
      icon: 'fa-folder',
      files: []
    },
    'f_imports': {
      id: 'f_imports',
      name: 'Uploads',
      icon: 'fa-folder',
      files: []
    }
  },

  // Staged Files for Initial Prompt Setup Modal
  stagedFiles: [],

  // Record Editing State
  editingRowIndex: null
};

// INITIALIZATION
document.addEventListener('DOMContentLoaded', () => {
  loadSavedDatabase();
  setupEventListeners();
  renderSidebar();
  renderTabs();
  renderActiveView();

  // Check if first time or prompt setup requested
  if (!localStorage.getItem('studio_sql_initialized')) {
    openInitialSetupModal();
  }
});

/** Load Saved Database from LocalStorage */
function loadSavedDatabase() {
  try {
    const saved = localStorage.getItem('studio_sql_db');
    if (saved) {
      const parsed = JSON.parse(saved);
      if (parsed.dbName) state.dbName = parsed.dbName;
      if (parsed.tables) state.tables = parsed.tables;
      if (parsed.folders) state.folders = parsed.folders;
    }
  } catch (e) {
    console.error('Failed to load database state from storage:', e);
  }
}

/** Get Clean Base DB Name without extension */
function getCleanDbName() {
  return (state.dbName || 'MyDatabase').replace(/\.sqlite$/i, '').replace(/\.db$/i, '');
}

/** Save Database State to LocalStorage (Internal quiet save without downloading file) */
function saveDatabaseToStorage() {
  try {
    const dbPayload = {
      dbName: state.dbName,
      tables: state.tables,
      folders: state.folders
    };
    localStorage.setItem('studio_sql_db', JSON.stringify(dbPayload));
    localStorage.setItem('studio_sql_initialized', 'true');
  } catch (e) {
    console.error('Error saving state to localStorage:', e);
  }
}

/** Explicit User Action: Clicked "Save DB" Button in Header */
function handleSaveDbButtonClick() {
  saveDatabaseToStorage();
  exportDatabaseSpreadsheet();
  showToast(`Saved DB & generated Excel spreadsheet "${getCleanDbName()}.xlsx"`, 'success');
}

/** Export Database as Excel Spreadsheet file (.xlsx) with matching DB name */
function exportDatabaseSpreadsheet() {
  const baseName = getCleanDbName();
  const filename = `${baseName}.xlsx`;

  if (typeof XLSX !== 'undefined') {
    const wb = XLSX.utils.book_new();
    const tableKeys = Object.keys(state.tables);

    if (tableKeys.length === 0) {
      const ws = XLSX.utils.json_to_sheet([{ Note: 'Empty Database' }]);
      XLSX.utils.book_append_sheet(wb, ws, 'Database');
    } else {
      tableKeys.forEach(tableName => {
        const table = state.tables[tableName];
        const rows = (table && table.rows && table.rows.length > 0) ? table.rows : [];
        const ws = XLSX.utils.json_to_sheet(rows.length > 0 ? rows : [{}]);
        // Excel worksheet name max length 31 chars
        const sheetName = tableName.substring(0, 30);
        XLSX.utils.book_append_sheet(wb, ws, sheetName);
      });
    }

    // Save/Download Excel Spreadsheet File
    XLSX.writeFile(wb, filename);
  } else {
    // CSV Fallback
    let csvData = "";
    Object.keys(state.tables).forEach(tableName => {
      csvData += `=== TABLE: ${tableName} ===\n`;
      const table = state.tables[tableName];
      if (table && table.columns) {
        csvData += table.columns.map(c => c.name).join(",") + "\n";
        (table.rows || []).forEach(r => {
          csvData += table.columns.map(c => JSON.stringify(r[c.name] !== undefined ? r[c.name] : "")).join(",") + "\n";
        });
      }
      csvData += "\n";
    });

    const blob = new Blob([csvData], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = `${baseName}.csv`;
    document.body.appendChild(link);
    link.click();
    link.remove();
  }
}

/** Open Rename Database Modal */
function openRenameDbModal() {
  const input = document.getElementById('renameDbInput');
  if (input) input.value = getCleanDbName();
  openModal('renameDbModal');
}

/** Handle Rename Database Form Submit */
function handleRenameDbSubmit(e) {
  e.preventDefault();
  const input = document.getElementById('renameDbInput');
  let newName = input.value.trim() || 'MyDatabase';
  if (!newName.endsWith('.sqlite')) newName += '.sqlite';

  state.dbName = newName;
  document.getElementById('headerDbName').innerText = newName;
  saveDatabaseToStorage();
  closeModal('renameDbModal');
  showToast(`Database renamed to "${newName}"`, 'success');
}

/** Export Database as JSON file */
function exportDatabaseFile() {
  const dbPayload = {
    dbName: state.dbName,
    exportedAt: new Date().toISOString(),
    tables: state.tables,
    folders: state.folders
  };
  const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(dbPayload, null, 2));
  const downloadAnchor = document.createElement('a');
  downloadAnchor.setAttribute("href", dataStr);
  downloadAnchor.setAttribute("download", `${state.dbName.replace('.sqlite', '')}_export.json`);
  document.body.appendChild(downloadAnchor);
  downloadAnchor.click();
  downloadAnchor.remove();
  showToast(`Exported ${state.dbName}`, 'success');
}

/** Setup Event Listeners & Theme Switcher */
function setupEventListeners() {
  // Theme Toggle Button
  const themeBtn = document.getElementById('themeToggleBtn');
  if (themeBtn) {
    themeBtn.addEventListener('click', toggleTheme);
  }

  // Setup Global File Drop listeners to support folder drop targets
  setupFolderDragAndDropHandlers();
}

/** Toggle Theme */
function toggleTheme() {
  const body = document.body;
  const icon = document.querySelector('#themeToggleBtn i');
  if (body.classList.contains('dark-theme')) {
    body.classList.remove('dark-theme');
    body.classList.add('light-theme');
    icon.className = 'fa-solid fa-sun';
    showToast('Switched to Light Theme', 'info');
  } else {
    body.classList.remove('light-theme');
    body.classList.add('dark-theme');
    icon.className = 'fa-solid fa-moon';
    showToast('Switched to Dark Theme', 'info');
  }
}

/** RENDER SIDEBAR (TABLES & FOLDERS) */
function renderSidebar() {
  // Header Info
  document.getElementById('headerDbName').innerText = state.dbName;
  const tableKeys = Object.keys(state.tables);
  document.getElementById('headerTableCount').innerText = `${tableKeys.length} Tables`;

  // Render Tables List
  const tablesContainer = document.getElementById('tablesList');
  tablesContainer.innerHTML = '';

  if (tableKeys.length === 0) {
    tablesContainer.innerHTML = `<div class="placeholder-text p-2">No tables found. Click + to create one.</div>`;
  } else {
    tableKeys.forEach(tableName => {
      const isSelected = state.activeTable === tableName && state.activeView === 'table-contents';
      const item = document.createElement('div');
      item.className = `tree-item ${isSelected ? 'active' : ''}`;
      item.onclick = () => openTableTab(tableName);
      item.innerHTML = `
        <div class="tree-item-name">
          <i class="fa-solid fa-table"></i>
          <span>${tableName}</span>
        </div>
        <div class="tree-item-actions">
          <button class="btn-icon-xs" title="Modify Table" onclick="event.stopPropagation(); openEditSchemaModal('${tableName}')"><i class="fa-solid fa-pen"></i></button>
          <button class="btn-icon-xs text-danger" title="Delete Table" onclick="event.stopPropagation(); deleteTableByName('${tableName}')"><i class="fa-solid fa-trash-can"></i></button>
        </div>
      `;
      tablesContainer.appendChild(item);
    });
  }

  // Render Folders & Files Tree (WITH DRAG AND DROP TARGET BINDINGS)
  const foldersContainer = document.getElementById('foldersTree');
  foldersContainer.innerHTML = '';

  const folderIds = Object.keys(state.folders);
  if (folderIds.length === 0) {
    foldersContainer.innerHTML = `<div class="placeholder-text p-2">No folders found. Click + to create.</div>`;
  } else {
    folderIds.forEach(fId => {
      const folder = state.folders[fId];
      const folderBox = document.createElement('div');
      folderBox.className = 'folder-box';
      folderBox.setAttribute('data-folder-id', fId);

      // Render Folder Header & Files List
      let filesHtml = '';
      if (folder.files && folder.files.length > 0) {
        folder.files.forEach(file => {
          const isFileActive = state.activeFileId === file.id && state.activeView === 'file-preview';
          let iconClass = 'fa-file';
          if (file.type.startsWith('image/')) iconClass = 'fa-file-image';
          else if (file.type.startsWith('video/')) iconClass = 'fa-file-video';
          else if (file.name.endsWith('.csv')) iconClass = 'fa-file-csv';
          else if (file.name.endsWith('.sql')) iconClass = 'fa-file-code';

          filesHtml += `
            <div class="file-item ${isFileActive ? 'active' : ''}" onclick="openFilePreview('${fId}', '${file.id}')">
              <div class="file-name">
                <i class="fa-solid ${iconClass}"></i>
                <span>${file.name}</span>
              </div>
              <button class="btn-icon-xs text-danger" title="Delete File" onclick="event.stopPropagation(); deleteFile('${fId}', '${file.id}')"><i class="fa-solid fa-xmark"></i></button>
            </div>
          `;
        });
      } else {
        filesHtml = `<div class="placeholder-text" style="font-size: 10px; padding: 4px;">Drag & drop files here</div>`;
      }

      folderBox.innerHTML = `
        <div class="folder-header" onclick="triggerFolderUpload('${fId}')">
          <div class="folder-title">
            <i class="fa-solid fa-folder text-warning"></i>
            <span>${folder.name}</span>
          </div>
          <span class="folder-count">${folder.files ? folder.files.length : 0}</span>
        </div>
        <div class="file-list">
          ${filesHtml}
        </div>
      `;

      // BIND DRAG & DROP EVENTS FOR ANIMATIONS ON THIS FOLDER
      attachFolderDropListeners(folderBox, fId);

      foldersContainer.appendChild(folderBox);
    });
  }
}

/** BIND DRAG & DROP HOVER ANIMATIONS TO FOLDER BOXES */
function attachFolderDropListeners(element, folderId) {
  // Drag Enter
  element.addEventListener('dragenter', (e) => {
    e.preventDefault();
    e.stopPropagation();
    element.classList.add('drag-over');
  });

  // Drag Over
  element.addEventListener('dragover', (e) => {
    e.preventDefault();
    e.stopPropagation();
    element.classList.add('drag-over');
  });

  // Drag Leave
  element.addEventListener('dragleave', (e) => {
    e.preventDefault();
    e.stopPropagation();
    element.classList.remove('drag-over');
  });

  // Drop
  element.addEventListener('drop', (e) => {
    e.preventDefault();
    e.stopPropagation();
    element.classList.remove('drag-over');

    const files = e.dataTransfer.files;
    if (files && files.length > 0) {
      processFilesForFolder(folderId, files);
    }
  });
}

/** Global dragover setup */
function setupFolderDragAndDropHandlers() {
  window.addEventListener('dragover', (e) => e.preventDefault());
  window.addEventListener('drop', (e) => e.preventDefault());
}

/** Trigger upload for a specific folder */
function triggerFolderUpload(folderId) {
  state.activeFolderId = folderId;
  const input = document.getElementById('globalFileInput');
  if (input) input.click();
}

/** Handle File Selection from Hidden File Input */
function handleGlobalFileSelect(e) {
  const files = e.target.files;
  const targetFolderId = state.activeFolderId || Object.keys(state.folders)[0];
  if (files && files.length > 0 && targetFolderId) {
    processFilesForFolder(targetFolderId, files);
  }
}

/** Process Uploaded/Dropped Files and save into Target Folder */
function processFilesForFolder(folderId, fileList) {
  const folder = state.folders[folderId];
  if (!folder) return;

  let addedCount = 0;
  Array.from(fileList).forEach(file => {
    const reader = new FileReader();
    reader.onload = (event) => {
      const fileRecord = {
        id: 'file_' + Date.now() + '_' + Math.random().toString(36).substr(2, 5),
        name: file.name,
        size: formatBytes(file.size),
        type: file.type || 'file',
        content: event.target.result // DataURL or text
      };
      if (!folder.files) folder.files = [];
      folder.files.push(fileRecord);
      addedCount++;

      renderSidebar();
      saveDatabaseToStorage();

      // Automatically open preview for the last added file
      openFilePreview(folderId, fileRecord.id);
    };

    // Read as DataURL for media images/videos or text for code
    if (file.type.startsWith('image/') || file.type.startsWith('video/')) {
      reader.readAsDataURL(file);
    } else {
      reader.readAsText(file);
    }
  });

  showToast(`Uploaded files into "${folder.name}"`, 'success');
}

/** Format Bytes Utility */
function formatBytes(bytes) {
  if (!bytes) return '0 B';
  const k = 1024;
  const sizes = ['B', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
}

/** TAB MANAGEMENT SYSTEM */
function openTableTab(tableName) {
  if (!state.tables[tableName]) return;
  state.activeTable = tableName;
  state.activeView = 'table-contents';

  const tabId = `table:${tableName}`;
  if (!state.tabs.find(t => t.id === tabId)) {
    state.tabs.push({ id: tabId, type: 'table', name: tableName, icon: 'fa-table' });
  }
  state.activeTabId = tabId;

  renderTabs();
  renderActiveView();
}

function openFilePreview(folderId, fileId) {
  const folder = state.folders[folderId];
  if (!folder || !folder.files) return;
  const file = folder.files.find(f => f.id === fileId);
  if (!file) return;

  state.activeFolderId = folderId;
  state.activeFileId = fileId;
  state.activeView = 'file-preview';

  const tabId = `file:${fileId}`;
  if (!state.tabs.find(t => t.id === tabId)) {
    let icon = 'fa-file';
    if (file.type.startsWith('image/')) icon = 'fa-file-image';
    else if (file.type.startsWith('video/')) icon = 'fa-file-video';
    state.tabs.push({ id: tabId, type: 'file', folderId, fileId, name: file.name, icon });
  }
  state.activeTabId = tabId;

  renderTabs();
  renderActiveView();
}

function openQueryTab() {
  state.activeView = 'sql-query';
  const tabId = 'sql-console';
  if (!state.tabs.find(t => t.id === tabId)) {
    state.tabs.push({ id: tabId, type: 'query', name: 'SQL Query Console', icon: 'fa-terminal' });
  }
  state.activeTabId = tabId;
  renderTabs();
  renderActiveView();
}

function closeTab(e, tabId) {
  e.stopPropagation();
  state.tabs = state.tabs.filter(t => t.id !== tabId);
  if (state.activeTabId === tabId) {
    if (state.tabs.length > 0) {
      const nextTab = state.tabs[state.tabs.length - 1];
      state.activeTabId = nextTab.id;
      if (nextTab.type === 'table') {
        state.activeTable = nextTab.name;
        state.activeView = 'table-contents';
      } else if (nextTab.type === 'file') {
        state.activeFolderId = nextTab.folderId;
        state.activeFileId = nextTab.fileId;
        state.activeView = 'file-preview';
      } else {
        state.activeView = 'sql-query';
      }
    } else {
      openQueryTab();
    }
  }
  renderTabs();
  renderActiveView();
}

/** Render Tab Bar */
function renderTabs() {
  const container = document.getElementById('tabItemsWrapper');
  container.innerHTML = '';

  state.tabs.forEach(tab => {
    const isActive = tab.id === state.activeTabId;
    const tabEl = document.createElement('div');
    tabEl.className = `tab-item ${isActive ? 'active' : ''}`;
    tabEl.onclick = () => {
      state.activeTabId = tab.id;
      if (tab.type === 'table') {
        state.activeTable = tab.name;
        state.activeView = 'table-contents';
      } else if (tab.type === 'file') {
        state.activeFolderId = tab.folderId;
        state.activeFileId = tab.fileId;
        state.activeView = 'file-preview';
      } else {
        state.activeView = 'sql-query';
      }
      renderTabs();
      renderActiveView();
    };

    tabEl.innerHTML = `
      <i class="fa-solid ${tab.icon}"></i>
      <span>${tab.name}</span>
      <button class="tab-close" onclick="closeTab(event, '${tab.id}')"><i class="fa-solid fa-xmark"></i></button>
    `;
    container.appendChild(tabEl);
  });
}

/** RENDER MAIN ACTIVE VIEW */
function renderActiveView() {
  // Hide all view pages
  document.querySelectorAll('.view-page').forEach(v => v.classList.remove('active-view'));

  if (state.activeView === 'table-contents') {
    const view = document.getElementById('view-table-contents');
    if (view) view.classList.add('active-view');
    renderTableContents();
  } else if (state.activeView === 'file-preview') {
    const view = document.getElementById('view-file-preview');
    if (view) view.classList.add('active-view');
    renderFilePreview();
  } else if (state.activeView === 'sql-query') {
    const view = document.getElementById('view-sql-query');
    if (view) view.classList.add('active-view');
  }
}

/** RENDER TABLE CONTENTS DATA GRID */
function renderTableContents() {
  const tableName = state.activeTable;
  const tableData = state.tables[tableName];

  document.getElementById('activeTableName').innerHTML = `<i class="fa-solid fa-table text-accent"></i> ${tableName}`;
  if (!tableData) {
    document.getElementById('mainDataTableHeader').innerHTML = '';
    document.getElementById('mainDataTableBody').innerHTML = `<tr><td colspan="10" class="placeholder-text p-4">Table "${tableName}" not found.</td></tr>`;
    return;
  }

  const columns = tableData.columns || [];
  const rows = tableData.rows || [];

  document.getElementById('activeTableRowCount').innerText = `${rows.length} rows`;
  document.getElementById('activeTableColCount').innerText = `${columns.length} columns`;
  document.getElementById('gridFooterSummary').innerText = `Showing ${rows.length} records in table "${tableName}"`;

  // Render Table Header
  const headerTr = document.getElementById('mainDataTableHeader');
  headerTr.innerHTML = '<th>#</th>';
  columns.forEach(col => {
    headerTr.innerHTML += `<th>${col.name} <small class="text-muted">(${col.type})</small> ${col.pk ? '<span class="text-warning">PK</span>' : ''}</th>`;
  });
  headerTr.innerHTML += `<th style="width: 80px;">Actions</th>`;

  // Render Table Body
  renderTableRows(rows, columns);
}

/** Filter Table Rows based on Search Input */
function filterTableDataRows(query) {
  const tableName = state.activeTable;
  const tableData = state.tables[tableName];
  if (!tableData) return;

  const q = query.toLowerCase().trim();
  const filtered = tableData.rows.filter(row => {
    return Object.values(row).some(val => String(val).toLowerCase().includes(q));
  });

  renderTableRows(filtered, tableData.columns);
}

/** Render Table Row Elements */
function renderTableRows(rows, columns) {
  const body = document.getElementById('mainDataTableBody');
  body.innerHTML = '';

  if (rows.length === 0) {
    body.innerHTML = `<tr><td colspan="${columns.length + 2}" class="placeholder-text text-center p-4">No records found. Click "+ Add Record" to insert data.</td></tr>`;
    return;
  }

  rows.forEach((row, index) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td><strong>${index + 1}</strong></td>`;

    columns.forEach(col => {
      const val = row[col.name] !== undefined ? row[col.name] : '';
      tr.innerHTML += `<td>${val}</td>`;
    });

    tr.innerHTML += `
      <td>
        <div class="cell-actions">
          <button class="btn-icon-xs" title="Edit Record" onclick="openEditRecordModal(${index})"><i class="fa-solid fa-pen"></i></button>
          <button class="btn-icon-xs text-danger" title="Delete Record" onclick="deleteRecordRow(${index})"><i class="fa-solid fa-trash-can"></i></button>
        </div>
      </td>
    `;
    body.appendChild(tr);
  });
}

/** RENDER FILE PREVIEWER (VIDEO, IMAGE, CSV, CODE) */
function renderFilePreview() {
  const folder = state.folders[state.activeFolderId];
  if (!folder || !folder.files) return;
  const file = folder.files.find(f => f.id === state.activeFileId);
  if (!file) return;

  document.getElementById('previewFileName').innerText = file.name;
  document.getElementById('previewFileSize').innerText = file.size;

  const container = document.getElementById('previewBodyContainer');
  container.innerHTML = '';

  if (file.type.startsWith('image/')) {
    container.innerHTML = `<img src="${file.content}" class="media-preview-img" alt="${file.name}">`;
  } else if (file.type.startsWith('video/')) {
    container.innerHTML = `
      <video controls class="media-preview-video">
        <source src="${file.content}" type="${file.type}">
        Your browser does not support the video tag.
      </video>
    `;
  } else if (file.name.endsWith('.csv')) {
    // Render CSV Content as Interactive Table
    const lines = (file.content || '').split('\n').filter(l => l.trim());
    if (lines.length > 0) {
      let tableHtml = '<table class="grid-table"><thead><tr>';
      const headers = lines[0].split(',');
      headers.forEach(h => tableHtml += `<th>${h.trim()}</th>`);
      tableHtml += '</tr></thead><tbody>';

      for (let i = 1; i < lines.length; i++) {
        tableHtml += '<tr>';
        const cols = lines[i].split(',');
        cols.forEach(c => tableHtml += `<td>${c.trim()}</td>`);
        tableHtml += '</tr>';
      }
      tableHtml += 'tbody></table>';
      container.innerHTML = `<div class="data-grid-container" style="width: 100%; max-height: 80vh;">${tableHtml}</div>`;
    } else {
      container.innerHTML = `<div class="placeholder-text">CSV file is empty.</div>`;
    }
  } else {
    // Plain Text or Code
    container.innerHTML = `<pre class="code-preview-box"><code>${escapeHtml(file.content)}</code></pre>`;
  }
}

/** Escape HTML Utility */
function escapeHtml(str) {
  if (!str) return '';
  return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/** Delete Active Preview File */
function deleteActivePreviewFile() {
  if (state.activeFolderId && state.activeFileId) {
    deleteFile(state.activeFolderId, state.activeFileId);
  }
}

function deleteFile(folderId, fileId) {
  const folder = state.folders[folderId];
  if (!folder || !folder.files) return;
  
  if (confirm('Are you sure you want to delete this file?')) {
    folder.files = folder.files.filter(f => f.id !== fileId);
    closeTab(new Event('click'), `file:${fileId}`);
    renderSidebar();
    saveDatabaseToStorage();
    showToast('File deleted', 'info');
  }
}

/** MODAL MANAGEMENT FUNCTIONS */
function openModal(modalId) {
  const modal = document.getElementById(modalId);
  if (modal) modal.classList.remove('hidden');
}

function closeModal(modalId) {
  const modal = document.getElementById(modalId);
  if (modal) modal.classList.add('hidden');
}

/** INITIAL PROMPT SETUP MODAL */
function openInitialSetupModal() {
  state.stagedFiles = [];
  renderStagedFilesList();
  openModal('initialSetupModal');
}

function handleModalFileSelect(e) {
  const files = e.target.files;
  if (files && files.length > 0) {
    Array.from(files).forEach(f => state.stagedFiles.push(f));
    renderStagedFilesList();
  }
}

function renderStagedFilesList() {
  const container = document.getElementById('modalStagedFilesList');
  container.innerHTML = '';
  state.stagedFiles.forEach((file, index) => {
    const div = document.createElement('div');
    div.className = 'staged-file-item';
    div.innerHTML = `
      <span><i class="fa-solid fa-file text-accent"></i> ${file.name} (${formatBytes(file.size)})</span>
      <button class="btn-icon-xs text-danger" onclick="state.stagedFiles.splice(${index}, 1); renderStagedFilesList();"><i class="fa-solid fa-xmark"></i></button>
    `;
    container.appendChild(div);
  });
}

function handleInitialSetupSubmit(e) {
  e.preventDefault();
  const nameInput = document.getElementById('promptDbName');
  let name = nameInput.value.trim() || 'MyDatabase';
  if (!name.endsWith('.sqlite')) name += '.sqlite';

  state.dbName = name;

  // Process staged files into Uploads folder
  if (state.stagedFiles.length > 0) {
    processFilesForFolder('f_imports', state.stagedFiles);
  }

  // Create default sample tables if requested
  const scaffold = document.getElementById('scaffoldSampleTables').checked;
  if (!scaffold) {
    state.tables = {};
    state.activeTable = null;
  } else {
    state.activeTable = 'Users';
  }

  saveDatabaseToStorage();
  closeModal('initialSetupModal');
  renderSidebar();
  renderTabs();
  renderActiveView();
  showToast(`Database "${state.dbName}" initialized!`, 'success');
}

/** CREATE NEW TABLE MODAL */
function openCreateTableModal() {
  const oldInput = document.getElementById('editingOldTableName');
  if (oldInput) oldInput.value = '';

  document.getElementById('newTableName').value = '';
  const tbody = document.getElementById('createTableColumnsTbody');
  tbody.innerHTML = '';
  
  // Default first column: id INTEGER PK
  addCreateTableColumnRow('id', 'INTEGER', true, true);
  addCreateTableColumnRow('name', 'TEXT', false, true);

  openModal('createTableModal');
}

function addCreateTableColumnRow(name = '', type = 'TEXT', pk = false, nn = false, oldName = '') {
  const tbody = document.getElementById('createTableColumnsTbody');
  const tr = document.createElement('tr');
  const origName = oldName || name;
  tr.innerHTML = `
    <td><input type="text" class="form-control col-name" data-old-name="${origName}" value="${name}" placeholder="attribute_name" required></td>
    <td>
      <select class="form-control col-type">
        <option value="INTEGER" ${type === 'INTEGER' ? 'selected' : ''}>INTEGER</option>
        <option value="TEXT" ${type === 'TEXT' ? 'selected' : ''}>TEXT</option>
        <option value="REAL" ${type === 'REAL' ? 'selected' : ''}>REAL</option>
        <option value="DATETIME" ${type === 'DATETIME' ? 'selected' : ''}>DATETIME</option>
      </select>
    </td>
    <td><input type="checkbox" class="col-pk" ${pk ? 'checked' : ''}></td>
    <td><input type="checkbox" class="col-nn" ${nn ? 'checked' : ''}></td>
    <td><button type="button" class="btn-icon-xs text-danger" onclick="this.closest('tr').remove()"><i class="fa-solid fa-trash-can"></i></button></td>
  `;
  tbody.appendChild(tr);
}

function handleCreateTableSubmit(e) {
  e.preventDefault();
  const oldTableName = (document.getElementById('editingOldTableName') ? document.getElementById('editingOldTableName').value : '').trim();
  const newTableName = document.getElementById('newTableName').value.trim();
  if (!newTableName) return;

  const trs = document.querySelectorAll('#createTableColumnsTbody tr');
  const newColumns = [];
  trs.forEach(tr => {
    const nameInput = tr.querySelector('.col-name');
    const colName = nameInput.value.trim();
    const oldColName = (nameInput.getAttribute('data-old-name') || '').trim();
    const colType = tr.querySelector('.col-type').value;
    const pk = tr.querySelector('.col-pk').checked;
    const nn = tr.querySelector('.col-nn').checked;

    if (colName) {
      newColumns.push({ name: colName, oldName: oldColName, type: colType, pk, nn });
    }
  });

  if (newColumns.length === 0) {
    showToast('Please define at least one column attribute', 'danger');
    return;
  }

  // Preserve existing row contents & migrate attributes when renamed
  let finalRows = [];
  if (oldTableName && state.tables[oldTableName]) {
    const oldRows = state.tables[oldTableName].rows || [];
    const oldColumns = state.tables[oldTableName].columns || [];

    finalRows = oldRows.map(row => {
      const migratedRow = {};
      newColumns.forEach((newCol, colIdx) => {
        // 1. Direct attribute match
        if (row[newCol.name] !== undefined) {
          migratedRow[newCol.name] = row[newCol.name];
        } 
        // 2. Renamed attribute match from oldName
        else if (newCol.oldName && row[newCol.oldName] !== undefined) {
          migratedRow[newCol.name] = row[newCol.oldName];
        } 
        // 3. Positional attribute match by column index
        else if (oldColumns[colIdx] && row[oldColumns[colIdx].name] !== undefined) {
          migratedRow[newCol.name] = row[oldColumns[colIdx].name];
        } 
        // Default empty string
        else {
          migratedRow[newCol.name] = '';
        }
      });
      return migratedRow;
    });

    if (oldTableName !== newTableName) {
      delete state.tables[oldTableName];
      const oldTab = state.tabs.find(t => t.id === `table:${oldTableName}`);
      if (oldTab) {
        oldTab.id = `table:${newTableName}`;
        oldTab.name = newTableName;
      }
      if (state.activeTabId === `table:${oldTableName}`) {
        state.activeTabId = `table:${newTableName}`;
      }
    }
  }

  // Clean column definitions for storage
  const cleanedColumns = newColumns.map(({ name, type, pk, nn }) => ({ name, type, pk, nn }));

  state.tables[newTableName] = {
    columns: cleanedColumns,
    rows: finalRows
  };
  state.activeTable = newTableName;

  closeModal('createTableModal');
  saveDatabaseToStorage();
  openTableTab(newTableName);
  renderSidebar();
  renderTabs();
  renderActiveView();
  showToast(`Table "${newTableName}" modified. All attribute contents preserved!`, 'success');
}

/** EDIT / MODIFY SCHEMA MODAL */
function openEditSchemaModal(targetTable = null) {
  const tableName = targetTable || state.activeTable;
  if (!tableName || !state.tables[tableName]) return;

  const oldInput = document.getElementById('editingOldTableName');
  if (oldInput) oldInput.value = tableName;

  document.getElementById('newTableName').value = tableName;
  const tbody = document.getElementById('createTableColumnsTbody');
  tbody.innerHTML = '';

  state.tables[tableName].columns.forEach(col => {
    addCreateTableColumnRow(col.name, col.type, col.pk, col.nn, col.name);
  });

  openModal('createTableModal');
}

/** DELETE TABLE */
function deleteActiveTable() {
  if (state.activeTable) {
    deleteTableByName(state.activeTable);
  }
}

function deleteTableByName(tableName) {
  if (confirm(`Are you sure you want to delete table "${tableName}"?`)) {
    delete state.tables[tableName];
    closeTab(new Event('click'), `table:${tableName}`);
    state.activeTable = Object.keys(state.tables)[0] || null;
    renderSidebar();
    renderTabs();
    renderActiveView();
    saveDatabaseToStorage();
    showToast(`Table "${tableName}" deleted`, 'info');
  }
}

/** ADD / EDIT RECORD ROW MODAL */
function openAddRecordModal() {
  const tableName = state.activeTable;
  const tableData = state.tables[tableName];
  if (!tableData) return;

  state.editingRowIndex = null;
  document.getElementById('recordModalTitle').innerText = `Add Record to "${tableName}"`;
  document.getElementById('recordModalSub').innerText = 'Enter values for each column field';

  renderRecordFormFields({});
  openModal('recordModal');
}

function openEditRecordModal(rowIndex) {
  const tableName = state.activeTable;
  const tableData = state.tables[tableName];
  if (!tableData || !tableData.rows[rowIndex]) return;

  state.editingRowIndex = rowIndex;
  document.getElementById('recordModalTitle').innerText = `Edit Record in "${tableName}"`;
  document.getElementById('recordModalSub').innerText = `Editing Row #${rowIndex + 1}`;

  renderRecordFormFields(tableData.rows[rowIndex]);
  openModal('recordModal');
}

function renderRecordFormFields(rowObj) {
  const tableName = state.activeTable;
  const columns = state.tables[tableName].columns;

  const container = document.getElementById('recordFieldsContainer');
  container.innerHTML = '';

  columns.forEach(col => {
    const val = rowObj[col.name] !== undefined ? rowObj[col.name] : '';
    const fieldGroup = document.createElement('div');
    fieldGroup.className = 'form-group';
    fieldGroup.innerHTML = `
      <label>${col.name} <small class="text-muted">(${col.type})</small></label>
      <input type="${col.type === 'INTEGER' || col.type === 'REAL' ? 'number' : 'text'}" 
             class="form-control record-input-field" 
             data-col-name="${col.name}" 
             value="${val}" 
             ${col.nn ? 'required' : ''}>
    `;
    container.appendChild(fieldGroup);
  });
}

function handleRecordFormSubmit(e) {
  e.preventDefault();
  const tableName = state.activeTable;
  const tableData = state.tables[tableName];
  if (!tableData) return;

  const inputs = document.querySelectorAll('.record-input-field');
  const recordRow = {};

  inputs.forEach(input => {
    const colName = input.getAttribute('data-col-name');
    let val = input.value;
    if (input.type === 'number') val = val !== '' ? Number(val) : '';
    recordRow[colName] = val;
  });

  if (state.editingRowIndex !== null) {
    const pkCol = tableData.columns.find(c => c.pk) || tableData.columns[0];
    const pkVal = tableData.rows[state.editingRowIndex] ? tableData.rows[state.editingRowIndex][pkCol.name] : undefined;
    tableData.rows[state.editingRowIndex] = recordRow;
    const sqlStr = generateUpdateSql(tableName, recordRow, pkCol ? pkCol.name : null, pkVal);
    showToast(`Executed: ${sqlStr}`, 'success');
  } else {
    tableData.rows.push(recordRow);
    const sqlStr = generateInsertSql(tableName, recordRow);
    showToast(`Executed: ${sqlStr}`, 'success');
  }

  closeModal('recordModal');
  saveDatabaseToStorage();
  renderTableContents();
}

function deleteRecordRow(index) {
  const tableName = state.activeTable;
  const tableData = state.tables[tableName];
  if (!tableData || !tableData.rows[index]) return;

  if (confirm('Delete this record?')) {
    const pkCol = tableData.columns.find(c => c.pk) || tableData.columns[0];
    const pkVal = tableData.rows[index][pkCol ? pkCol.name : ''];
    const sqlStr = generateDeleteSql(tableName, pkCol ? pkCol.name : null, pkVal);

    tableData.rows.splice(index, 1);
    saveDatabaseToStorage();
    renderTableContents();
    showToast(`Executed: ${sqlStr}`, 'info');
  }
}

/** CREATE FOLDER MODAL */
function openCreateFolderModal() {
  document.getElementById('newFolderName').value = '';
  openModal('createFolderModal');
}

function handleCreateFolderSubmit(e) {
  e.preventDefault();
  const name = document.getElementById('newFolderName').value.trim();
  if (!name) return;

  const fId = 'f_' + Date.now();
  state.folders[fId] = {
    id: fId,
    name: name,
    icon: 'fa-folder',
    files: []
  };

  closeModal('createFolderModal');
  saveDatabaseToStorage();
  renderSidebar();
  showToast(`Folder "${name}" created`, 'success');
}

/** =========================================================================
 * ACCURATE SQL GENERATOR & EXECUTION ENGINE
 * Supports: SELECT, INSERT INTO, UPDATE, DELETE FROM, CREATE TABLE, DROP TABLE
 * ========================================================================= */

/** Generate accurate CREATE TABLE DDL SQL statement string */
function generateCreateTableSql(tableName, columns) {
  const colDefs = columns.map(c => {
    let def = `  "${c.name}" ${c.type}`;
    if (c.pk) def += " PRIMARY KEY";
    if (c.nn) def += " NOT NULL";
    return def;
  }).join(",\n");
  return `CREATE TABLE "${tableName}" (\n${colDefs}\n);`;
}

/** Generate accurate INSERT DML SQL statement string */
function generateInsertSql(tableName, rowObj) {
  const cols = Object.keys(rowObj);
  const vals = cols.map(c => {
    const v = rowObj[c];
    if (v === null || v === undefined || v === '') return 'NULL';
    if (typeof v === 'number') return v;
    return `'${String(v).replace(/'/g, "''")}'`;
  });
  return `INSERT INTO "${tableName}" (${cols.map(c => `"${c}"`).join(", ")}) VALUES (${vals.join(", ")});`;
}

/** Generate accurate UPDATE DML SQL statement string */
function generateUpdateSql(tableName, rowObj, whereCol, whereVal) {
  const setClause = Object.keys(rowObj).map(c => {
    const v = rowObj[c];
    let valStr = 'NULL';
    if (v !== null && v !== undefined && v !== '') {
      valStr = typeof v === 'number' ? v : `'${String(v).replace(/'/g, "''")}'`;
    }
    return `"${c}" = ${valStr}`;
  }).join(", ");

  let whereClause = "";
  if (whereCol && whereVal !== undefined) {
    const wVal = typeof whereVal === 'number' ? whereVal : `'${String(whereVal).replace(/'/g, "''")}'`;
    whereClause = ` WHERE "${whereCol}" = ${wVal}`;
  }
  return `UPDATE "${tableName}" SET ${setClause}${whereClause};`;
}

/** Generate accurate DELETE DML SQL statement string */
function generateDeleteSql(tableName, whereCol, whereVal) {
  let whereClause = "";
  if (whereCol && whereVal !== undefined) {
    const wVal = typeof whereVal === 'number' ? whereVal : `'${String(whereVal).replace(/'/g, "''")}'`;
    whereClause = ` WHERE "${whereCol}" = ${wVal}`;
  }
  return `DELETE FROM "${tableName}"${whereClause};`;
}

/** Generate accurate DROP TABLE DDL SQL statement string */
function generateDropTableSql(tableName) {
  return `DROP TABLE "${tableName}";`;
}

/** 
 * ACCURATE SQL INTERPRETER & EXECUTION ENGINE 
 * Parses and executes single or multi-line SQL queries entered by the user.
 */
function parseAndExecuteSql(rawSql) {
  const cleanSql = rawSql.trim().replace(/;$/, '');
  if (!cleanSql) return { success: false, error: 'Empty query statement.' };

  const tokens = cleanSql.split(/\s+/);
  const command = tokens[0].toUpperCase();

  try {
    if (command === 'SELECT') {
      return executeSelectQuery(cleanSql);
    } else if (command === 'INSERT') {
      return executeInsertQuery(cleanSql);
    } else if (command === 'UPDATE') {
      return executeUpdateQuery(cleanSql);
    } else if (command === 'DELETE') {
      return executeDeleteQuery(cleanSql);
    } else if (command === 'CREATE') {
      return executeCreateTableQuery(cleanSql);
    } else if (command === 'DROP') {
      return executeDropTableQuery(cleanSql);
    } else {
      return { success: false, error: `Unsupported SQL command: "${command}". Supported: SELECT, INSERT, UPDATE, DELETE, CREATE TABLE, DROP TABLE.` };
    }
  } catch (err) {
    return { success: false, error: `SQL Syntax / Runtime Error: ${err.message}` };
  }
}

/** Execute SELECT statement */
function executeSelectQuery(sql) {
  const match = sql.match(/SELECT\s+(.+?)\s+FROM\s+(["`\w]+)(?:\s+WHERE\s+(.+?))?(?:\s+ORDER\s+BY\s+(\w+)(?:\s+(ASC|DESC))?)?(?:\s+LIMIT\s+(\d+))?$/i);
  let selectedColsStr, tableName, whereClause = null, orderCol = null, orderDir = 'ASC', limitVal = null;

  if (!match) {
    const simpleMatch = sql.match(/SELECT\s+(.+?)\s+FROM\s+(["`\w]+)/i);
    if (!simpleMatch) throw new Error('Invalid SELECT syntax. Usage: SELECT * FROM TableName WHERE col = val;');
    selectedColsStr = simpleMatch[1].trim();
    tableName = simpleMatch[2].replace(/["`]/g, '').trim();
  } else {
    selectedColsStr = match[1].trim();
    tableName = match[2].replace(/["`]/g, '').trim();
    whereClause = match[3] ? match[3].trim() : null;
    orderCol = match[4] ? match[4].trim() : null;
    orderDir = match[5] ? match[5].toUpperCase() : 'ASC';
    limitVal = match[6] ? parseInt(match[6], 10) : null;
  }

  if (!state.tables[tableName]) {
    throw new Error(`Table "${tableName}" does not exist in active database.`);
  }

  const table = state.tables[tableName];
  let rows = [...table.rows];

  if (whereClause) {
    rows = applyWhereCondition(rows, whereClause);
  }

  if (orderCol) {
    rows.sort((a, b) => {
      const valA = a[orderCol], valB = b[orderCol];
      if (valA < valB) return orderDir === 'ASC' ? -1 : 1;
      if (valA > valB) return orderDir === 'ASC' ? 1 : -1;
      return 0;
    });
  }

  if (limitVal !== null && limitVal > 0) {
    rows = rows.slice(0, limitVal);
  }

  let columns = table.columns;
  if (selectedColsStr !== '*') {
    const colNames = selectedColsStr.split(',').map(c => c.replace(/["`]/g, '').trim());
    columns = table.columns.filter(c => colNames.includes(c.name));
  }

  return {
    success: true,
    type: 'SELECT',
    tableName: tableName,
    columns: columns,
    rows: rows,
    rowCount: rows.length,
    message: `Query returned ${rows.length} record(s) from table "${tableName}".`
  };
}

/** Execute WHERE Clause Condition */
function applyWhereCondition(rows, whereClause) {
  const opMatch = whereClause.match(/(["`\w]+)\s*(=|!=|<>|>|<|>=|<=|LIKE)\s*(.+)/i);
  if (!opMatch) return rows;

  const field = opMatch[1].replace(/["`]/g, '').trim();
  const op = opMatch[2].toUpperCase();
  let rawVal = opMatch[3].trim().replace(/^['"]|['"]$/g, '');

  return rows.filter(row => {
    const cellVal = row[field];
    if (cellVal === undefined) return false;

    let targetVal = rawVal;
    if (typeof cellVal === 'number') {
      targetVal = Number(rawVal);
    }

    switch (op) {
      case '=': return cellVal == targetVal;
      case '!=':
      case '<>': return cellVal != targetVal;
      case '>': return cellVal > targetVal;
      case '<': return cellVal < targetVal;
      case '>=': return cellVal >= targetVal;
      case '<=': return cellVal <= targetVal;
      case 'LIKE': 
        const pattern = String(rawVal).replace(/%/g, '.*');
        return new RegExp('^' + pattern + '$', 'i').test(String(cellVal));
      default: return true;
    }
  });
}

/** Execute INSERT INTO statement */
function executeInsertQuery(sql) {
  const match = sql.match(/INSERT\s+INTO\s+(["`\w]+)\s*\((.+?)\)\s*VALUES\s*\((.+?)\)$/i);
  if (!match) {
    throw new Error("Invalid INSERT syntax. Usage: INSERT INTO TableName (col1, col2) VALUES ('val1', 'val2');");
  }

  const tableName = match[1].replace(/["`]/g, '').trim();
  if (!state.tables[tableName]) {
    throw new Error(`Table "${tableName}" does not exist.`);
  }

  const cols = match[2].split(',').map(c => c.replace(/["`]/g, '').trim());
  const vals = match[3].split(',').map(v => {
    let val = v.trim().replace(/^['"]|['"]$/g, '');
    if (!isNaN(val) && val !== '') val = Number(val);
    return val;
  });

  if (cols.length !== vals.length) {
    throw new Error("Number of columns does not match number of values.");
  }

  const newRow = {};
  cols.forEach((col, idx) => {
    newRow[col] = vals[idx];
  });

  state.tables[tableName].rows.push(newRow);
  saveDatabaseToStorage();

  if (state.activeTable === tableName) {
    renderTableContents();
  }

  return {
    success: true,
    type: 'DML',
    tableName: tableName,
    affectedRows: 1,
    message: `1 row inserted into "${tableName}".`
  };
}

/** Execute UPDATE statement */
function executeUpdateQuery(sql) {
  const match = sql.match(/UPDATE\s+(["`\w]+)\s+SET\s+(.+?)(?:\s+WHERE\s+(.+?))?$/i);
  if (!match) {
    throw new Error("Invalid UPDATE syntax. Usage: UPDATE TableName SET col1 = 'val1' WHERE id = 1;");
  }

  const tableName = match[1].replace(/["`]/g, '').trim();
  if (!state.tables[tableName]) {
    throw new Error(`Table "${tableName}" does not exist.`);
  }

  const setStr = match[2].trim();
  const whereClause = match[3] ? match[3].trim() : null;

  const setPairs = setStr.split(',').map(p => {
    const parts = p.split('=');
    const key = parts[0].replace(/["`]/g, '').trim();
    let val = parts[1].trim().replace(/^['"]|['"]$/g, '');
    if (!isNaN(val) && val !== '') val = Number(val);
    return { key, val };
  });

  const table = state.tables[tableName];
  let updatedCount = 0;

  table.rows.forEach(row => {
    let shouldUpdate = true;
    if (whereClause) {
      const filtered = applyWhereCondition([row], whereClause);
      if (filtered.length === 0) shouldUpdate = false;
    }

    if (shouldUpdate) {
      setPairs.forEach(pair => {
        row[pair.key] = pair.val;
      });
      updatedCount++;
    }
  });

  saveDatabaseToStorage();
  if (state.activeTable === tableName) {
    renderTableContents();
  }

  return {
    success: true,
    type: 'DML',
    tableName: tableName,
    affectedRows: updatedCount,
    message: `${updatedCount} row(s) updated in "${tableName}".`
  };
}

/** Execute DELETE statement */
function executeDeleteQuery(sql) {
  const match = sql.match(/DELETE\s+FROM\s+(["`\w]+)(?:\s+WHERE\s+(.+?))?$/i);
  if (!match) {
    throw new Error("Invalid DELETE syntax. Usage: DELETE FROM TableName WHERE id = 1;");
  }

  const tableName = match[1].replace(/["`]/g, '').trim();
  if (!state.tables[tableName]) {
    throw new Error(`Table "${tableName}" does not exist.`);
  }

  const whereClause = match[2] ? match[2].trim() : null;
  const table = state.tables[tableName];
  const initialLength = table.rows.length;

  if (whereClause) {
    table.rows = table.rows.filter(row => {
      const matchRows = applyWhereCondition([row], whereClause);
      return matchRows.length === 0;
    });
  } else {
    table.rows = [];
  }

  const deletedCount = initialLength - table.rows.length;

  saveDatabaseToStorage();
  if (state.activeTable === tableName) {
    renderTableContents();
  }

  return {
    success: true,
    type: 'DML',
    tableName: tableName,
    affectedRows: deletedCount,
    message: `${deletedCount} row(s) deleted from "${tableName}".`
  };
}

/** Execute CREATE TABLE statement */
function executeCreateTableQuery(sql) {
  const match = sql.match(/CREATE\s+TABLE\s+(["`\w]+)\s*\((.+?)\)$/i);
  if (!match) {
    throw new Error("Invalid CREATE TABLE syntax. Usage: CREATE TABLE TableName (id INTEGER PRIMARY KEY, name TEXT);");
  }

  const tableName = match[1].replace(/["`]/g, '').trim();
  const colDefsStr = match[2].trim();

  const columns = colDefsStr.split(',').map(def => {
    const parts = def.trim().split(/\s+/);
    const name = parts[0].replace(/["`]/g, '');
    const type = (parts[1] || 'TEXT').toUpperCase();
    const pk = def.toUpperCase().includes('PRIMARY KEY');
    const nn = def.toUpperCase().includes('NOT NULL');
    return { name, type, pk, nn };
  });

  state.tables[tableName] = {
    columns: columns,
    rows: []
  };

  saveDatabaseToStorage();
  openTableTab(tableName);
  renderSidebar();

  return {
    success: true,
    type: 'DDL',
    tableName: tableName,
    message: `Table "${tableName}" created successfully.`
  };
}

/** Execute DROP TABLE statement */
function executeDropTableQuery(sql) {
  const match = sql.match(/DROP\s+TABLE\s+(["`\w]+)$/i);
  if (!match) {
    throw new Error("Invalid DROP TABLE syntax. Usage: DROP TABLE TableName;");
  }

  const tableName = match[1].replace(/["`]/g, '').trim();
  if (!state.tables[tableName]) {
    throw new Error(`Table "${tableName}" does not exist.`);
  }

  delete state.tables[tableName];
  closeTab(new Event('click'), `table:${tableName}`);
  state.activeTable = Object.keys(state.tables)[0] || null;

  saveDatabaseToStorage();
  renderSidebar();
  renderTabs();
  renderActiveView();

  return {
    success: true,
    type: 'DDL',
    tableName: tableName,
    message: `Table "${tableName}" dropped successfully.`
  };
}

/** CONSOLE SQL RUNNER INTERFACE */
function runConsoleSqlQuery() {
  const input = document.getElementById('consoleQueryInput').value.trim();
  const resultsContainer = document.getElementById('consoleQueryResults');
  if (!input) return;

  const statements = input.split(';').map(s => s.trim()).filter(s => s.length > 0);
  resultsContainer.innerHTML = '';

  statements.forEach(stmt => {
    const time = new Date().toLocaleTimeString();
    const logHeader = `<div class="log-line info" style="margin-bottom: 4px; font-weight: 600;">[${time}] &gt; ${escapeHtml(stmt)};</div>`;

    const res = parseAndExecuteSql(stmt);
    if (!res.success) {
      resultsContainer.innerHTML += `${logHeader}<div class="text-danger mb-3" style="padding-left: 10px;"><i class="fa-solid fa-circle-xmark"></i> ${escapeHtml(res.error)}</div>`;
    } else {
      let content = `<div class="text-success mb-2" style="padding-left: 10px;"><i class="fa-solid fa-circle-check"></i> ${escapeHtml(res.message)}</div>`;

      if (res.type === 'SELECT' && res.columns && res.rows) {
        let tableHtml = '<div class="data-grid-container mb-3" style="max-height: 300px;"><table class="grid-table"><thead><tr>';
        res.columns.forEach(c => tableHtml += `<th>${c.name} <small class="text-muted">(${c.type})</small></th>`);
        tableHtml += '</tr></thead><tbody>';

        if (res.rows.length === 0) {
          tableHtml += `<tr><td colspan="${res.columns.length}" class="placeholder-text p-2">0 rows returned.</td></tr>`;
        } else {
          res.rows.forEach(r => {
            tableHtml += '<tr>';
            res.columns.forEach(c => tableHtml += `<td>${r[c.name] !== undefined ? r[c.name] : ''}</td>`);
            tableHtml += '</tr>';
          });
        }
        tableHtml += '</tbody></table></div>';
        content += tableHtml;
      }
      resultsContainer.innerHTML += `${logHeader}${content}`;
    }
  });
}

function clearConsoleQuery() {
  document.getElementById('consoleQueryInput').value = '';
  document.getElementById('consoleQueryResults').innerHTML = '<p class="placeholder-text">Run a query to see results.</p>';
}

/** TOAST NOTIFICATION SYSTEM */
function showToast(message, type = 'info') {
  const container = document.getElementById('toastContainer');
  if (!container) return;

  const toast = document.createElement('div');
  toast.className = `toast ${type}`;
  let icon = 'fa-circle-info';
  if (type === 'success') icon = 'fa-circle-check';
  if (type === 'danger') icon = 'fa-triangle-exclamation';

  toast.innerHTML = `<i class="fa-solid ${icon}"></i> <span>${message}</span>`;
  container.appendChild(toast);

  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateX(100%)';
    setTimeout(() => toast.remove(), 300);
  }, 3000);
}
