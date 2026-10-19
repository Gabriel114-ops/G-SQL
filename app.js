/**
 * IWM_SQLite - MICROSOFT SQL SERVER MANAGEMENT STUDIO (SSMS) ENGINE
 * Complete relational database studio:
 * - Authentic SSMS "Connect to Server" (Windows Auth vs SQL Auth / Laptop / Remote Server)
 * - SSMS Object Explorer with expandable tree (Databases, Tables, Columns, Keys, Views, Security)
 * - SSMS Right-Click Context Menu (Select Top 1000, Edit Top 200, Design, Script as, Delete)
 * - SSMS Table Designer (Column grid, PK toggle, Allow Nulls, Column Properties, guaranteed safe edit)
 * - SSMS Table Data Editor ("Edit Top 200 Rows" with inline editing, insert & delete)
 * - SSMS SQL Query Console (Execute / F5, Results Grid & Messages tabs, syntax highlighting)
 * - SSMS "Delete Objects" confirmation modal with clean foreign key integrity cleanup
 * - Relational Database Diagrams (interactive table cards and bezier connectors)
 * - Excel (.xlsx) Export & JSON Backup / Restore
 */

// DEFAULT SAMPLE DATABASE SCHEMA TEMPLATE
const SAMPLE_DATABASE_TABLES = {
  'Users': {
    columns: [
      { name: 'id', type: 'INTEGER', pk: true, nn: true, uq: true, def: '' },
      { name: 'name', type: 'TEXT', pk: false, nn: true, uq: false, def: '' },
      { name: 'email', type: 'TEXT', pk: false, nn: true, uq: true, def: '' },
      { name: 'role', type: 'TEXT', pk: false, nn: false, uq: false, def: 'User' },
      { name: 'created_at', type: 'TEXT', pk: false, nn: false, uq: false, def: '2026-01-01' }
    ],
    foreignKeys: [],
    rows: [
      { id: 1, name: 'John Doe', email: 'john@example.com', role: 'Admin', created_at: '2026-01-10' },
      { id: 2, name: 'Maria Garcia', email: 'maria@tech.org', role: 'Developer', created_at: '2026-01-15' },
      { id: 3, name: 'Alex Chen', email: 'alex@data.io', role: 'Analyst', created_at: '2026-02-01' }
    ]
  },
  'Orders': {
    columns: [
      { name: 'id', type: 'INTEGER', pk: true, nn: true, uq: true, def: '' },
      { name: 'user_id', type: 'INTEGER', pk: false, nn: true, uq: false, def: '' },
      { name: 'amount', type: 'REAL', pk: false, nn: true, uq: false, def: '0.00' },
      { name: 'status', type: 'TEXT', pk: false, nn: false, uq: false, def: 'Pending' }
    ],
    foreignKeys: [
      { column: 'user_id', refTable: 'Users', refColumn: 'id', onDelete: 'CASCADE' }
    ],
    rows: [
      { id: 101, user_id: 1, amount: 149.99, status: 'Completed' },
      { id: 102, user_id: 2, amount: 299.50, status: 'Pending' },
      { id: 103, user_id: 1, amount: 49.00, status: 'Completed' }
    ]
  },
  'Products': {
    columns: [
      { name: 'id', type: 'INTEGER', pk: true, nn: true, uq: true, def: '' },
      { name: 'title', type: 'TEXT', pk: false, nn: true, uq: true, def: '' },
      { name: 'price', type: 'REAL', pk: false, nn: true, uq: false, def: '0.00' },
      { name: 'stock', type: 'INTEGER', pk: false, nn: false, uq: false, def: '0' }
    ],
    foreignKeys: [],
    rows: [
      { id: 501, title: 'Studio SQL Pro', price: 49.00, stock: 100 },
      { id: 502, title: 'Data Engine Plugin', price: 19.99, stock: 250 },
      { id: 503, title: 'Cloud Sync Connector', price: 89.00, stock: 45 }
    ]
  }
};

// GLOBAL APPLICATION STATE (CLEAN BACKGROUND FIRST PRIOR TO CONNECTION)
const state = {
  dbName: 'MyDatabase.sqlite',
  activeView: 'empty-workspace', // Starts with clear empty background
  activeTable: null,

  // Connection State (SSMS Server & Authentication)
  connection: {
    connected: false, // Disconnected until user connects
    serverName: 'GABRIEL-PC\\SQLEXPRESS',
    serverType: 'Database Engine (IWM_SQLite)',
    authType: 'windows', // 'windows' | 'sql'
    userName: 'GABRIEL-PC\\Gabriel',
    networkProtocol: 'TCP/IP',
    rememberPassword: true
  },

  // Open Document Tabs (MDI Workspace - clear on launch)
  tabs: [],
  activeTabId: null,

  // Relational Tables Store
  tables: JSON.parse(JSON.stringify(SAMPLE_DATABASE_TABLES)),

  // Active Context Menu Target
  contextTargetTable: null,

  // Table Designer State
  designer: {
    editingTableName: null,
    isNew: false
  },

  // Record Editing State
  editingRowIndex: null,

  // Table Data Grid Sorting & Pagination
  sort: {
    column: null,
    direction: 'asc'
  },
  pagination: {
    pageSize: 200,
    currentPage: 1
  },

  // Database Diagrams State
  diagram: {
    zoom: 1,
    tablePositions: {},
    selectedTable: null,
    draggingTable: null,
    dragStart: { x: 0, y: 0 }
  },

  // Query Counter & History
  queryCounter: 1,
  queryHistory: []
};

// INITIALIZATION (PROMPT CONNECT PANEL FIRST ON LAUNCH OVER CLEAR BACKGROUND)
document.addEventListener('DOMContentLoaded', () => {
  loadSavedDatabase();
  loadQueryHistory();
  setupEventListeners();
  setupKeyboardShortcuts();
  updateConnectionUI();
  populateToolbarDbSelect();
  renderObjectExplorer();
  renderTabs();
  renderActiveView();
  setupLineNumbers();

  // Prompt Connect to Server dialog FIRST on launch
  openConnectServerModal();
});

/** =========================================================================
 * STORAGE & DATABASE PERSISTENCE
 * ========================================================================= */

function loadSavedDatabase() {
  try {
    const saved = localStorage.getItem('studio_sql_db');
    if (saved) {
      const parsed = JSON.parse(saved);
      if (parsed.dbName) state.dbName = parsed.dbName;
      if (parsed.connection) {
        state.connection.serverName = parsed.connection.serverName || state.connection.serverName;
        state.connection.authType = parsed.connection.authType || state.connection.authType;
        state.connection.userName = parsed.connection.userName || state.connection.userName;
      }
      if (parsed.tables && Object.keys(parsed.tables).length > 0) {
        state.tables = parsed.tables;
        Object.keys(state.tables).forEach(tName => {
          const t = state.tables[tName];
          if (!t.foreignKeys) t.foreignKeys = [];
          if (!t.rows) t.rows = [];
          if (t.columns) {
            t.columns.forEach(c => {
              if (c.uq === undefined) c.uq = false;
              if (c.def === undefined) c.def = '';
            });
          }
        });
      }
      if (parsed.diagramPositions) state.diagram.tablePositions = parsed.diagramPositions;
    }
  } catch (e) {
    console.error('Failed to load database state from storage:', e);
  }
}

function saveDatabaseToStorage() {
  try {
    const dbPayload = {
      dbName: state.dbName,
      connection: state.connection,
      tables: state.tables,
      diagramPositions: state.diagram.tablePositions
    };
    localStorage.setItem('studio_sql_db', JSON.stringify(dbPayload));
  } catch (e) {
    console.error('Error saving state to localStorage:', e);
  }
}

function getCleanDbName() {
  return (state.dbName || 'MyDatabase').replace(/\.sqlite$/i, '').replace(/\.db$/i, '');
}

function handleSaveDbButtonClick() {
  saveDatabaseToStorage();
  exportDatabaseSpreadsheet();
  showToast(`Database "${state.dbName}" saved & exported to Excel (.xlsx)`, 'success');
}

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
      tableKeys.forEach(tName => {
        const table = state.tables[tName];
        let ws;
        if (table.rows && table.rows.length > 0) {
          ws = XLSX.utils.json_to_sheet(table.rows);
        } else {
          const emptyRow = {};
          (table.columns || []).forEach(c => emptyRow[c.name] = '');
          ws = XLSX.utils.json_to_sheet([emptyRow]);
        }
        const sheetName = tName.substring(0, 31);
        XLSX.utils.book_append_sheet(wb, ws, sheetName);
      });
    }
    XLSX.writeFile(wb, filename);
  }
}

function exportDatabaseFile() {
  const dbPayload = {
    dbName: state.dbName,
    exportedAt: new Date().toISOString(),
    connection: state.connection,
    tables: state.tables,
    diagramPositions: state.diagram.tablePositions
  };
  const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(dbPayload, null, 2));
  const dl = document.createElement('a');
  dl.setAttribute("href", dataStr);
  dl.setAttribute("download", `${getCleanDbName()}_SSMS_Backup.json`);
  document.body.appendChild(dl);
  dl.click();
  dl.remove();
  showToast(`Exported "${state.dbName}" JSON backup`, 'success');
  closeAllMenus();
}

/** =========================================================================
 * SSMS CONNECT TO SERVER (WINDOWS AUTH VS SQL AUTH / LAPTOP / SERVER)
 * ========================================================================= */

function openConnectServerModal() {
  closeAllMenus();
  const serverInput = document.getElementById('connServerName');
  if (serverInput) serverInput.value = state.connection.serverName || 'GABRIEL-PC\\SQLEXPRESS';

  const currentAuth = state.connection.authType || 'windows';
  selectAuthMode(currentAuth);

  const winUser = document.getElementById('connWindowsUser');
  if (winUser) winUser.value = state.connection.userName || 'GABRIEL-PC\\Gabriel';

  const dbInput = document.getElementById('connDatabaseName');
  if (dbInput) dbInput.value = state.dbName || 'MyDatabase.sqlite';

  updateConnectionPreview();
  openModal('connectServerModal');
}

function selectAuthMode(mode) {
  const cardWin = document.getElementById('authCardWindows');
  const cardSql = document.getElementById('authCardSql');
  const authInput = document.getElementById('connAuthType');
  const infoText = document.getElementById('authInfoText');
  const winUserVal = (document.getElementById('connWindowsUser') ? document.getElementById('connWindowsUser').value : '') || 'GABRIEL-PC\\Gabriel';

  if (authInput) authInput.value = mode;

  if (mode === 'windows') {
    if (cardWin) {
      cardWin.classList.add('active');
      const radio = cardWin.querySelector('.auth-card-radio');
      if (radio) radio.innerHTML = '<i class="fa-solid fa-circle-dot"></i>';
    }
    if (cardSql) {
      cardSql.classList.remove('active');
      const radio = cardSql.querySelector('.auth-card-radio');
      if (radio) radio.innerHTML = '<i class="fa-regular fa-circle"></i>';
    }
    if (infoText) {
      infoText.innerHTML = `Windows Authentication connects directly using your workstation credentials (<strong>${escapeHtml(winUserVal)}</strong>). No password required. Ideal for local development on your PC or laptop.`;
    }
    handleAuthTypeChange('windows');
  } else {
    if (cardSql) {
      cardSql.classList.add('active');
      const radio = cardSql.querySelector('.auth-card-radio');
      if (radio) radio.innerHTML = '<i class="fa-solid fa-circle-dot"></i>';
    }
    if (cardWin) {
      cardWin.classList.remove('active');
      const radio = cardWin.querySelector('.auth-card-radio');
      if (radio) radio.innerHTML = '<i class="fa-regular fa-circle"></i>';
    }
    if (infoText) {
      infoText.innerHTML = `Server Authentication connects using database credentials. Enter the remote or local SQL instance host, Admin Login (e.g. <code>sa</code>), and password. Ideal for connecting to another laptop or remote server.`;
    }
    handleAuthTypeChange('sql');
  }
  updateConnectionPreview();
}

function openConnectServerModalWithAuth(mode) {
  openConnectServerModal();
  selectAuthMode(mode);
}

function applyServerPreset(val) {
  const input = document.getElementById('connServerName');
  if (input) input.value = val;
  const select = document.getElementById('connServerPreset');
  if (select) select.value = val;
  updateConnectionPreview();
}

function updateConnectionPreview() {
  const previewEl = document.getElementById('connLivePreviewText');
  if (!previewEl) return;

  const server = (document.getElementById('connServerName') ? document.getElementById('connServerName').value.trim() : '') || 'GABRIEL-PC\\SQLEXPRESS';
  const auth = (document.getElementById('connAuthType') ? document.getElementById('connAuthType').value : '') || 'windows';
  const db = (document.getElementById('connDatabaseName') ? document.getElementById('connDatabaseName').value.trim() : '') || 'MyDatabase.sqlite';
  const login = (document.getElementById('connSqlLogin') ? document.getElementById('connSqlLogin').value.trim() : '') || 'sa';

  const authLabel = auth === 'windows' ? 'Windows Auth' : `Server Auth (${login})`;
  previewEl.innerText = `Target: ${server} • ${authLabel} • ${db}`;
}

function togglePasswordVisibility() {
  const pwdInput = document.getElementById('connSqlPassword');
  const icon = document.getElementById('pwdToggleIcon');
  if (!pwdInput) return;

  if (pwdInput.type === 'password') {
    pwdInput.type = 'text';
    if (icon) icon.className = 'fa-solid fa-eye-slash';
  } else {
    pwdInput.type = 'password';
    if (icon) icon.className = 'fa-solid fa-eye';
  }
}

function handleAuthTypeChange(authType) {
  const rowWin = document.getElementById('rowWindowsUser');
  const rowLogin = document.getElementById('rowSqlLogin');
  const rowPwd = document.getElementById('rowSqlPassword');

  if (authType === 'windows') {
    if (rowWin) rowWin.classList.remove('hidden');
    if (rowLogin) rowLogin.classList.add('hidden');
    if (rowPwd) rowPwd.classList.add('hidden');
  } else {
    if (rowWin) rowWin.classList.add('hidden');
    if (rowLogin) rowLogin.classList.remove('hidden');
    if (rowPwd) rowPwd.classList.remove('hidden');
  }
}

function toggleConnectOptions() {
  const panel = document.getElementById('connOptionsPanel');
  const arrow = document.getElementById('connOptionsArrow');
  if (panel) {
    const isHidden = panel.classList.toggle('hidden');
    if (arrow) {
      arrow.className = isHidden ? 'fa-solid fa-chevron-right' : 'fa-solid fa-chevron-down';
    }
  }
}

function handleConnectServerSubmit(e) {
  e.preventDefault();
  const serverName = document.getElementById('connServerName').value.trim() || 'GABRIEL-PC\\SQLEXPRESS';
  const authType = document.getElementById('connAuthType').value;
  const dbNameInput = document.getElementById('connDatabaseName').value.trim();

  let userName = '';
  if (authType === 'windows') {
    userName = document.getElementById('connWindowsUser').value.trim() || 'GABRIEL-PC\\Gabriel';
  } else {
    userName = document.getElementById('connSqlLogin').value.trim() || 'sa';
  }

  // Schema seed choice
  const seedOption = document.querySelector('input[name="connSchemaSeed"]:checked');
  if (seedOption && seedOption.value === 'blank') {
    state.tables = {};
    state.activeTable = null;
    state.diagram.tablePositions = {};
  } else if (!state.tables || Object.keys(state.tables).length === 0) {
    state.tables = JSON.parse(JSON.stringify(SAMPLE_DATABASE_TABLES));
    state.activeTable = 'Users';
  }

  state.connection.connected = true;
  state.connection.serverName = serverName;
  state.connection.authType = authType;
  state.connection.userName = userName;

  if (dbNameInput) {
    let cleanDb = dbNameInput;
    if (!cleanDb.endsWith('.sqlite')) cleanDb += '.sqlite';
    state.dbName = cleanDb;
  }

  saveDatabaseToStorage();
  updateConnectionUI();
  populateToolbarDbSelect();
  renderObjectExplorer();
  closeModal('connectServerModal');

  // Open appropriate initial workspace tab
  if (state.tabs.length === 0) {
    if (state.tables && Object.keys(state.tables).length > 0) {
      const firstTable = Object.keys(state.tables)[0];
      openTableDataTab(firstTable);
    } else {
      openNewQueryTab();
    }
  } else {
    renderTabs();
    renderActiveView();
  }

  showToast(`Connected to server "${serverName}" via ${authType === 'windows' ? 'Windows' : 'Server'} Authentication`, 'success');
}

function disconnectServer() {
  closeAllMenus();
  state.connection.connected = false;
  state.tabs = [];
  state.activeTabId = null;
  state.activeView = 'empty-workspace';

  updateConnectionUI();
  renderObjectExplorer();
  renderTabs();
  renderActiveView();

  showToast(`Disconnected from "${state.connection.serverName}". Workspace cleared.`, 'info');
}

function updateConnectionUI() {
  const tag = document.getElementById('titlebarConnectionTag');
  const sbServer = document.getElementById('sbServerName');
  const sbAuth = document.getElementById('sbAuthUser');
  const sbDb = document.getElementById('sbDbName');

  const sName = state.connection.serverName || 'GABRIEL-PC\\SQLEXPRESS';
  const uName = state.connection.userName || 'GABRIEL-PC\\Gabriel';
  const isConn = state.connection.connected;

  if (tag) {
    tag.innerText = isConn ? `${sName} (Connected)` : `${sName} (Disconnected)`;
    tag.style.color = isConn ? 'var(--color-success)' : 'var(--color-danger)';
  }
  if (sbServer) sbServer.innerText = sName;
  if (sbAuth) sbAuth.innerText = `${uName} (${state.connection.authType === 'windows' ? 'Windows Auth' : 'SQL Auth'})`;
  if (sbDb) sbDb.innerText = state.dbName;
}

/** =========================================================================
 * SSMS OBJECT EXPLORER (DOCK PANEL TREE)
 * ========================================================================= */

function renderObjectExplorer() {
  const container = document.getElementById('objectExplorerTree');
  if (!container) return;
  container.innerHTML = '';

  const isConnected = state.connection.connected;
  const serverName = state.connection.serverName || 'GABRIEL-PC\\SQLEXPRESS';
  const userName = state.connection.userName || 'GABRIEL-PC\\Gabriel';

  // Server Root Node
  const rootNode = document.createElement('div');
  rootNode.className = 'tree-item-group';

  const rootHeader = document.createElement('div');
  rootHeader.className = `tree-node ${isConnected ? '' : 'text-muted'}`;
  rootHeader.innerHTML = `
    <span class="tree-expander expanded" onclick="toggleTreeNode(this)"><i class="fa-solid fa-caret-right"></i></span>
    <span class="tree-icon"><i class="fa-solid fa-server ${isConnected ? 'text-accent' : 'text-danger'}"></i></span>
    <span class="tree-label"><strong>${serverName}</strong> (SQL Server 16.0 - ${userName})</span>
  `;
  rootNode.appendChild(rootHeader);

  const rootChildren = document.createElement('div');
  rootChildren.className = 'tree-children open';

  if (!isConnected) {
    rootChildren.innerHTML = `
      <div class="tree-node text-muted" onclick="openConnectServerModal()">
        <span class="tree-icon"><i class="fa-solid fa-plug text-danger"></i></span>
        <span class="tree-label"><em>Disconnected. Click here to connect to server...</em></span>
      </div>
    `;
    rootNode.appendChild(rootChildren);
    container.appendChild(rootNode);
    return;
  }

  // DATABASES FOLDER
  const dbFolderNode = document.createElement('div');
  dbFolderNode.className = 'tree-item-group';
  dbFolderNode.innerHTML = `
    <div class="tree-node">
      <span class="tree-expander expanded" onclick="toggleTreeNode(this)"><i class="fa-solid fa-caret-right"></i></span>
      <span class="tree-icon"><i class="fa-solid fa-folder text-gold"></i></span>
      <span class="tree-label">Databases</span>
    </div>
  `;

  const dbFolderChildren = document.createElement('div');
  dbFolderChildren.className = 'tree-children open';

  // System Databases
  const sysDbNode = document.createElement('div');
  sysDbNode.className = 'tree-item-group';
  sysDbNode.innerHTML = `
    <div class="tree-node text-muted">
      <span class="tree-expander" onclick="toggleTreeNode(this)"><i class="fa-solid fa-caret-right"></i></span>
      <span class="tree-icon"><i class="fa-solid fa-folder"></i></span>
      <span class="tree-label">System Databases</span>
    </div>
    <div class="tree-children">
      <div class="tree-node text-muted"><span class="tree-icon"><i class="fa-solid fa-database"></i></span><span class="tree-label">master</span></div>
      <div class="tree-node text-muted"><span class="tree-icon"><i class="fa-solid fa-database"></i></span><span class="tree-label">model</span></div>
      <div class="tree-node text-muted"><span class="tree-icon"><i class="fa-solid fa-database"></i></span><span class="tree-label">msdb</span></div>
      <div class="tree-node text-muted"><span class="tree-icon"><i class="fa-solid fa-database"></i></span><span class="tree-label">tempdb</span></div>
    </div>
  `;
  dbFolderChildren.appendChild(sysDbNode);

  // Active User Database
  const activeDbNode = document.createElement('div');
  activeDbNode.className = 'tree-item-group';

  const activeDbHeader = document.createElement('div');
  activeDbHeader.className = 'tree-node active';
  activeDbHeader.innerHTML = `
    <span class="tree-expander expanded" onclick="toggleTreeNode(this)"><i class="fa-solid fa-caret-right"></i></span>
    <span class="tree-icon"><i class="fa-solid fa-database text-accent"></i></span>
    <span class="tree-label"><strong>${state.dbName}</strong></span>
  `;
  activeDbNode.appendChild(activeDbHeader);

  const activeDbChildren = document.createElement('div');
  activeDbChildren.className = 'tree-children open';

  // Database Diagrams Node
  const diagNode = document.createElement('div');
  diagNode.className = 'tree-node';
  diagNode.onclick = () => openErDiagramTab();
  diagNode.innerHTML = `
    <span class="tree-expander"><i class="fa-solid fa-caret-right" style="opacity:0;"></i></span>
    <span class="tree-icon"><i class="fa-solid fa-diagram-project text-info"></i></span>
    <span class="tree-label">Database Diagrams</span>
  `;
  activeDbChildren.appendChild(diagNode);

  // Tables Folder
  const tablesNode = document.createElement('div');
  tablesNode.className = 'tree-item-group';

  const tableKeys = Object.keys(state.tables);
  const tablesHeader = document.createElement('div');
  tablesHeader.className = 'tree-node';
  tablesHeader.innerHTML = `
    <span class="tree-expander expanded" onclick="toggleTreeNode(this)"><i class="fa-solid fa-caret-right"></i></span>
    <span class="tree-icon"><i class="fa-solid fa-folder text-gold"></i></span>
    <span class="tree-label">Tables (${tableKeys.length})</span>
    <div class="tree-node-actions">
      <button class="tree-action-btn" title="New Table Designer" onclick="event.stopPropagation(); openTableDesignerTab()"><i class="fa-solid fa-plus"></i></button>
    </div>
  `;
  tablesNode.appendChild(tablesHeader);

  const tablesChildren = document.createElement('div');
  tablesChildren.className = 'tree-children open';

  if (tableKeys.length === 0) {
    tablesChildren.innerHTML = `
      <div class="tree-node text-muted" onclick="openTableDesignerTab()">
        <span class="tree-label"><em>No tables. Click to create new table...</em></span>
      </div>
    `;
  } else {
    tableKeys.forEach(tName => {
      const tData = state.tables[tName];
      const isSelected = state.activeTable === tName && state.activeView === 'table-contents';

      const tNode = document.createElement('div');
      tNode.className = 'tree-item-group';

      const tHeader = document.createElement('div');
      tHeader.className = `tree-node ${isSelected ? 'active' : ''}`;
      tHeader.oncontextmenu = (e) => showSSMSContextMenu(e, tName);
      tHeader.onclick = () => openTableDataTab(tName);

      const rowCount = tData.rows ? tData.rows.length : 0;
      tHeader.innerHTML = `
        <span class="tree-expander" onclick="event.stopPropagation(); toggleTreeNode(this)"><i class="fa-solid fa-caret-right"></i></span>
        <span class="tree-icon"><i class="fa-solid fa-table text-accent"></i></span>
        <span class="tree-label">dbo.${tName}</span>
        <div class="tree-node-actions">
          <button class="tree-action-btn" title="Edit Top 200 Rows" onclick="event.stopPropagation(); openTableDataTab('${tName}')"><i class="fa-solid fa-pen-to-square"></i></button>
          <button class="tree-action-btn" title="Design Table" onclick="event.stopPropagation(); openTableDesignerForTable('${tName}')"><i class="fa-solid fa-table-columns"></i></button>
          <button class="tree-action-btn text-danger" title="Delete Table" onclick="event.stopPropagation(); promptDeleteTable('${tName}')"><i class="fa-solid fa-trash-can"></i></button>
        </div>
      `;
      tNode.appendChild(tHeader);

      // Columns & Keys under table
      const tChildren = document.createElement('div');
      tChildren.className = 'tree-children';

      // Columns Folder
      const colsFolder = document.createElement('div');
      colsFolder.className = 'tree-item-group';
      colsFolder.innerHTML = `
        <div class="tree-node">
          <span class="tree-expander" onclick="toggleTreeNode(this)"><i class="fa-solid fa-caret-right"></i></span>
          <span class="tree-icon"><i class="fa-solid fa-folder"></i></span>
          <span class="tree-label">Columns</span>
        </div>
      `;
      const colsChildren = document.createElement('div');
      colsChildren.className = 'tree-children';

      (tData.columns || []).forEach(col => {
        const isFk = (tData.foreignKeys || []).some(fk => fk.column === col.name);
        let keyIcon = '';
        if (col.pk) keyIcon = '<i class="fa-solid fa-key badge-pk"></i> ';
        else if (isFk) keyIcon = '<i class="fa-solid fa-link badge-fk"></i> ';

        const nullText = col.nn ? 'not null' : 'null';
        const colEl = document.createElement('div');
        colEl.className = 'tree-node';
        colEl.innerHTML = `
          <span class="tree-icon">${keyIcon || '<i class="fa-solid fa-table-cells text-muted"></i>'}</span>
          <span class="tree-label">${col.name} (${col.type.toLowerCase()}, ${nullText})</span>
        `;
        colsChildren.appendChild(colEl);
      });
      colsFolder.appendChild(colsChildren);
      tChildren.appendChild(colsFolder);

      // Keys Folder
      const keysFolder = document.createElement('div');
      keysFolder.className = 'tree-item-group';
      keysFolder.innerHTML = `
        <div class="tree-node">
          <span class="tree-expander" onclick="toggleTreeNode(this)"><i class="fa-solid fa-caret-right"></i></span>
          <span class="tree-icon"><i class="fa-solid fa-folder"></i></span>
          <span class="tree-label">Keys</span>
        </div>
      `;
      const keysChildren = document.createElement('div');
      keysChildren.className = 'tree-children';

      const pkCol = (tData.columns || []).find(c => c.pk);
      if (pkCol) {
        keysChildren.innerHTML += `
          <div class="tree-node"><span class="tree-icon"><i class="fa-solid fa-key badge-pk"></i></span><span class="tree-label">PK_${tName}</span></div>
        `;
      }
      (tData.foreignKeys || []).forEach(fk => {
        keysChildren.innerHTML += `
          <div class="tree-node"><span class="tree-icon"><i class="fa-solid fa-link badge-fk"></i></span><span class="tree-label">FK_${tName}_${fk.refTable}</span></div>
        `;
      });
      keysFolder.appendChild(keysChildren);
      tChildren.appendChild(keysFolder);

      tNode.appendChild(tChildren);
      tablesChildren.appendChild(tNode);
    });
  }

  tablesNode.appendChild(tablesChildren);
  activeDbChildren.appendChild(tablesNode);

  // Views Folder
  const viewsNode = document.createElement('div');
  viewsNode.className = 'tree-node text-muted';
  viewsNode.innerHTML = `
    <span class="tree-expander"><i class="fa-solid fa-caret-right" style="opacity:0;"></i></span>
    <span class="tree-icon"><i class="fa-solid fa-folder"></i></span>
    <span class="tree-label">Views</span>
  `;
  activeDbChildren.appendChild(viewsNode);

  // Programmability Folder
  const progNode = document.createElement('div');
  progNode.className = 'tree-node text-muted';
  progNode.innerHTML = `
    <span class="tree-expander"><i class="fa-solid fa-caret-right" style="opacity:0;"></i></span>
    <span class="tree-icon"><i class="fa-solid fa-folder"></i></span>
    <span class="tree-label">Programmability</span>
  `;
  activeDbChildren.appendChild(progNode);

  activeDbNode.appendChild(activeDbChildren);
  dbFolderChildren.appendChild(activeDbNode);
  dbFolderNode.appendChild(dbFolderChildren);
  rootChildren.appendChild(dbFolderNode);

  // SECURITY FOLDER
  const secNode = document.createElement('div');
  secNode.className = 'tree-item-group';
  secNode.innerHTML = `
    <div class="tree-node">
      <span class="tree-expander" onclick="toggleTreeNode(this)"><i class="fa-solid fa-caret-right"></i></span>
      <span class="tree-icon"><i class="fa-solid fa-shield-halved text-accent"></i></span>
      <span class="tree-label">Security</span>
    </div>
    <div class="tree-children">
      <div class="tree-node"><span class="tree-icon"><i class="fa-solid fa-user-lock"></i></span><span class="tree-label">Logins</span></div>
      <div class="tree-node"><span class="tree-icon"><i class="fa-solid fa-users"></i></span><span class="tree-label">Server Roles</span></div>
    </div>
  `;
  rootChildren.appendChild(secNode);

  rootNode.appendChild(rootChildren);
  container.appendChild(rootNode);
}

function toggleTreeNode(expanderEl) {
  expanderEl.classList.toggle('expanded');
  const parentNode = expanderEl.closest('.tree-item-group');
  if (parentNode) {
    const children = parentNode.querySelector('.tree-children');
    if (children) children.classList.toggle('open');
  }
}

function refreshObjectExplorer() {
  renderObjectExplorer();
  populateToolbarDbSelect();
  showToast('Object Explorer refreshed', 'info');
}

function toggleObjectExplorer() {
  const oe = document.getElementById('ssmsObjectExplorer');
  if (oe) {
    oe.classList.toggle('hidden');
  }
}

function filterObjectExplorerTree(query) {
  const q = query.trim().toLowerCase();
  const nodes = document.querySelectorAll('#objectExplorerTree .tree-node');
  nodes.forEach(n => {
    const text = n.innerText.toLowerCase();
    n.style.display = (!q || text.includes(q)) ? '' : 'none';
  });
}

function populateToolbarDbSelect() {
  const select = document.getElementById('toolbarDbSelect');
  if (!select) return;
  select.innerHTML = '';

  const opt1 = document.createElement('option');
  opt1.value = state.dbName;
  opt1.innerText = state.dbName;
  opt1.selected = true;
  select.appendChild(opt1);

  ['master', 'model', 'msdb', 'tempdb'].forEach(sysDb => {
    const opt = document.createElement('option');
    opt.value = sysDb;
    opt.innerText = sysDb;
    select.appendChild(opt);
  });
}

function handleToolbarDbChange(val) {
  if (val !== state.dbName && !['master', 'model', 'msdb', 'tempdb'].includes(val)) {
    state.dbName = val;
    renderObjectExplorer();
    updateConnectionUI();
  }
}

/** =========================================================================
 * SSMS CONTEXT MENU (RIGHT CLICK ACTIONS ON TABLE)
 * ========================================================================= */

function showSSMSContextMenu(e, tableName) {
  e.preventDefault();
  state.contextTargetTable = tableName;
  const menu = document.getElementById('ssmsContextMenu');
  if (!menu) return;

  menu.style.left = `${Math.min(e.clientX, window.innerWidth - 190)}px`;
  menu.style.top = `${Math.min(e.clientY, window.innerHeight - 220)}px`;
  menu.classList.add('show');
}

function hideSSMSContextMenu() {
  const menu = document.getElementById('ssmsContextMenu');
  if (menu) menu.classList.remove('show');
}

function handleCtxSelectTop1000() {
  hideSSMSContextMenu();
  const tName = state.contextTargetTable || state.activeTable;
  if (!tName) return;

  openNewQueryTab(`-- Select Top 1000 Rows from [dbo].[${tName}]\nSELECT TOP 1000 * FROM [dbo].[${tName}];`);
  runConsoleSqlQuery();
}

function handleCtxEditTop200() {
  hideSSMSContextMenu();
  const tName = state.contextTargetTable || state.activeTable;
  if (tName) openTableDataTab(tName);
}

function handleCtxDesignTable() {
  hideSSMSContextMenu();
  const tName = state.contextTargetTable || state.activeTable;
  if (tName) openTableDesignerForTable(tName);
}

function handleCtxScriptCreate() {
  hideSSMSContextMenu();
  const tName = state.contextTargetTable || state.activeTable;
  const tData = state.tables[tName];
  if (!tData) return;

  let script = `-- Create Table Script for [dbo].[${tName}]\nCREATE TABLE [dbo].[${tName}] (\n`;
  const colLines = (tData.columns || []).map(c => {
    let line = `  [${c.name}] ${c.type}`;
    if (c.pk) line += ' PRIMARY KEY';
    if (c.nn) line += ' NOT NULL';
    if (c.def) line += ` DEFAULT '${c.def}'`;
    return line;
  });
  script += colLines.join(',\n') + '\n);\nGO';

  openNewQueryTab(script);
}

function handleCtxRenameTable() {
  hideSSMSContextMenu();
  const tName = state.contextTargetTable || state.activeTable;
  if (!tName) return;

  const newName = prompt(`Enter new name for table "${tName}":`, tName);
  if (newName && newName.trim() && newName.trim() !== tName) {
    renameTable(tName, newName.trim());
  }
}

function handleCtxDeleteTable() {
  hideSSMSContextMenu();
  const tName = state.contextTargetTable || state.activeTable;
  if (tName) promptDeleteTable(tName);
}

/** =========================================================================
 * GUARANTEED TABLE DELETION & SAFE REFERENTIAL INTEGRITY CLEANUP
 * ========================================================================= */

function promptDeleteTable(tableName) {
  closeAllMenus();
  if (!tableName || !state.tables[tableName]) return;

  state.contextTargetTable = tableName;
  const label = document.getElementById('deleteTargetTableName');
  if (label) label.innerText = tableName;

  // Check if any other table references this table via Foreign Key
  const dependentFks = [];
  Object.keys(state.tables).forEach(otherTName => {
    if (otherTName !== tableName) {
      (state.tables[otherTName].foreignKeys || []).forEach(fk => {
        if (fk.refTable === tableName) {
          dependentFks.push(`${otherTName}.${fk.column} &rarr; ${tableName}.${fk.refColumn}`);
        }
      });
    }
  });

  const warningBox = document.getElementById('deleteDependentFkWarning');
  const warningText = document.getElementById('deleteDependentFkText');
  if (warningBox && warningText) {
    if (dependentFks.length > 0) {
      warningBox.classList.remove('hidden');
      warningText.innerHTML = `The following foreign key relationships reference this table and will be cleanly removed to maintain integrity:<br><strong>${dependentFks.join(', ')}</strong>`;
    } else {
      warningBox.classList.add('hidden');
    }
  }

  openModal('deleteTableModal');
}

function confirmExecuteTableDeletion() {
  const tableName = state.contextTargetTable;
  if (!tableName || !state.tables[tableName]) {
    closeModal('deleteTableModal');
    return;
  }

  deleteTableByName(tableName);
  closeModal('deleteTableModal');
}

function deleteActiveTable() {
  closeAllMenus();
  if (state.activeTable) promptDeleteTable(state.activeTable);
}

function deleteTableByName(tableName) {
  if (!state.tables[tableName]) return;

  // 1. Remove table schema and rows
  delete state.tables[tableName];
  delete state.diagram.tablePositions[tableName];

  // 2. Clean up any foreign key in OTHER tables that referenced this deleted table
  Object.keys(state.tables).forEach(otherName => {
    if (state.tables[otherName].foreignKeys) {
      state.tables[otherName].foreignKeys = state.tables[otherName].foreignKeys.filter(
        fk => fk.refTable !== tableName
      );
    }
  });

  // 3. Close open tabs for this table
  state.tabs = state.tabs.filter(t => t.id !== `table:${tableName}` && t.id !== `design:${tableName}`);

  // 4. Update active table pointer
  const remainingTables = Object.keys(state.tables);
  if (state.activeTable === tableName) {
    state.activeTable = remainingTables.length > 0 ? remainingTables[0] : null;
  }

  // 5. If closed active tab, switch to next tab or new query
  if (state.activeTabId === `table:${tableName}` || state.activeTabId === `design:${tableName}`) {
    if (state.tabs.length > 0) {
      const nextTab = state.tabs[state.tabs.length - 1];
      switchActiveTab(nextTab.id);
    } else if (state.activeTable) {
      openTableDataTab(state.activeTable);
    } else {
      openNewQueryTab();
    }
  }

  saveDatabaseToStorage();
  renderObjectExplorer();
  renderTabs();
  renderActiveView();
  showToast(`Table "dbo.${tableName}" deleted successfully.`, 'info');
}

function renameTable(oldName, newName) {
  if (!oldName || !newName || !state.tables[oldName]) return;
  if (state.tables[newName]) {
    showToast(`Table "${newName}" already exists!`, 'danger');
    return;
  }

  state.tables[newName] = state.tables[oldName];
  delete state.tables[oldName];

  if (state.diagram.tablePositions[oldName]) {
    state.diagram.tablePositions[newName] = state.diagram.tablePositions[oldName];
    delete state.diagram.tablePositions[oldName];
  }

  // Update foreign keys referencing old table
  Object.keys(state.tables).forEach(t => {
    (state.tables[t].foreignKeys || []).forEach(fk => {
      if (fk.refTable === oldName) fk.refTable = newName;
    });
  });

  // Update tabs
  state.tabs.forEach(tab => {
    if (tab.id === `table:${oldName}`) {
      tab.id = `table:${newName}`;
      tab.name = `dbo.${newName} [Edit]`;
      tab.tableName = newName;
    }
    if (tab.id === `design:${oldName}`) {
      tab.id = `design:${newName}`;
      tab.name = `dbo.${newName} [Design]`;
      tab.tableName = newName;
    }
  });
  if (state.activeTabId === `table:${oldName}`) state.activeTabId = `table:${newName}`;
  if (state.activeTabId === `design:${oldName}`) state.activeTabId = `design:${newName}`;
  if (state.activeTable === oldName) state.activeTable = newName;

  saveDatabaseToStorage();
  renderObjectExplorer();
  renderTabs();
  renderActiveView();
  showToast(`Table renamed to "dbo.${newName}"`, 'success');
}

/** =========================================================================
 * SSMS TABLE DESIGNER (SAFE SCHEMA EDITING FOR 1ST, 2ND & ALL TABLES)
 * ========================================================================= */

function openTableDesignerTab() {
  closeAllMenus();
  state.designer.isNew = true;
  state.designer.editingTableName = '';

  const designerTabId = 'design:NewTable';
  if (!state.tabs.find(t => t.id === designerTabId)) {
    state.tabs.push({ id: designerTabId, type: 'designer', name: 'dbo.NewTable [Design]', icon: 'fa-table-columns' });
  }
  state.activeTabId = designerTabId;
  state.activeView = 'table-designer';

  const titleEl = document.getElementById('designerTitle');
  if (titleEl) titleEl.innerHTML = `<i class="fa-solid fa-table-columns text-accent"></i> <span>dbo.NewTable [Design]</span>`;
  const nameInput = document.getElementById('designerTableNameInput');
  if (nameInput) nameInput.value = 'Table1';

  const tbody = document.getElementById('designerColumnsTbody');
  if (tbody) {
    tbody.innerHTML = '';
    addDesignerColumnRow('id', 'INTEGER', true, true, true, '');
    addDesignerColumnRow('name', 'TEXT', false, true, false, '');
  }

  renderTabs();
  renderActiveView();
}

function openTableDesignerForActiveTable() {
  if (state.activeTable) openTableDesignerForTable(state.activeTable);
}

function openTableDesignerForTable(tableName) {
  closeAllMenus();
  if (!tableName || !state.tables[tableName]) return;

  state.designer.isNew = false;
  state.designer.editingTableName = tableName;

  const designerTabId = `design:${tableName}`;
  if (!state.tabs.find(t => t.id === designerTabId)) {
    state.tabs.push({ id: designerTabId, type: 'designer', name: `dbo.${tableName} [Design]`, tableName: tableName, icon: 'fa-table-columns' });
  }
  state.activeTabId = designerTabId;
  state.activeView = 'table-designer';

  const titleEl = document.getElementById('designerTitle');
  if (titleEl) titleEl.innerHTML = `<i class="fa-solid fa-table-columns text-accent"></i> <span>dbo.${tableName} [Design]</span>`;
  const nameInput = document.getElementById('designerTableNameInput');
  if (nameInput) nameInput.value = tableName;

  const tbody = document.getElementById('designerColumnsTbody');
  if (tbody) {
    tbody.innerHTML = '';
    const tableData = state.tables[tableName];
    (tableData.columns || []).forEach(col => {
      addDesignerColumnRow(col.name, col.type, col.pk, col.nn, col.uq, col.def, col.name);
    });
  }

  renderTabs();
  renderActiveView();
}

function addDesignerColumnRow(name = '', type = 'TEXT', pk = false, nn = false, uq = false, def = '', oldName = '') {
  const tbody = document.getElementById('designerColumnsTbody');
  if (!tbody) return;

  const tr = document.createElement('tr');
  const origName = oldName || name;

  tr.innerHTML = `
    <td style="text-align: center;">
      <button type="button" class="btn-icon-xs col-pk-btn ${pk ? 'active-pk' : ''}" title="Toggle Primary Key" onclick="toggleDesignerRowPk(this)">
        <i class="fa-solid fa-key ${pk ? 'badge-pk' : 'text-muted'}"></i>
      </button>
      <input type="checkbox" class="col-pk-check hidden" ${pk ? 'checked' : ''}>
    </td>
    <td>
      <input type="text" class="designer-col-name" data-old-name="${origName}" value="${name}" placeholder="Column name" onfocus="updatePropertiesPane(this)">
    </td>
    <td>
      <select class="designer-col-type" onchange="updatePropertiesPane(this)">
        <option value="INTEGER" ${type.toUpperCase() === 'INTEGER' ? 'selected' : ''}>int (INTEGER)</option>
        <option value="TEXT" ${type.toUpperCase() === 'TEXT' ? 'selected' : ''}>varchar(50) (TEXT)</option>
        <option value="REAL" ${type.toUpperCase() === 'REAL' ? 'selected' : ''}>decimal(18,2) (REAL)</option>
        <option value="DATETIME" ${type.toUpperCase() === 'DATETIME' ? 'selected' : ''}>datetime (DATETIME)</option>
      </select>
    </td>
    <td style="text-align: center;">
      <input type="checkbox" class="designer-col-null" ${!nn ? 'checked' : ''} onchange="updatePropertiesPane(this)">
    </td>
    <td style="text-align: center;">
      <input type="checkbox" class="designer-col-uq" ${uq ? 'checked' : ''}>
    </td>
    <td>
      <input type="text" class="designer-col-def" value="${def || ''}" placeholder="((NULL))">
    </td>
    <td style="text-align: center;">
      <button type="button" class="btn-icon-xs text-danger" title="Delete Column" onclick="this.closest('tr').remove()"><i class="fa-solid fa-trash-can"></i></button>
    </td>
  `;

  tbody.appendChild(tr);
  tr.onclick = () => {
    tbody.querySelectorAll('tr').forEach(r => r.classList.remove('selected'));
    tr.classList.add('selected');
    updatePropertiesPane(tr.querySelector('.designer-col-name'));
  };
}

function toggleDesignerRowPk(btn) {
  const tr = btn.closest('tr');
  const check = tr.querySelector('.col-pk-check');
  const icon = btn.querySelector('i');
  const nullCheck = tr.querySelector('.designer-col-null');

  const isPk = !check.checked;
  check.checked = isPk;
  if (isPk) {
    icon.className = 'fa-solid fa-key badge-pk';
    btn.classList.add('active-pk');
    if (nullCheck) nullCheck.checked = false; // PK cannot be null
  } else {
    icon.className = 'fa-solid fa-key text-muted';
    btn.classList.remove('active-pk');
  }
}

function updatePropertiesPane(el) {
  const tr = el.closest('tr');
  if (!tr) return;

  const colName = tr.querySelector('.designer-col-name').value || '(unnamed)';
  const colType = tr.querySelector('.designer-col-type').value;
  const isNull = tr.querySelector('.designer-col-null').checked;
  const isPk = tr.querySelector('.col-pk-check').checked;

  const titleDisp = document.getElementById('propColNameDisplay');
  const nameVal = document.getElementById('propNameVal');
  const typeVal = document.getElementById('propTypeVal');
  const nullVal = document.getElementById('propNullVal');
  const idVal = document.getElementById('propIdentityVal');

  if (titleDisp) titleDisp.innerText = colName;
  if (nameVal) nameVal.innerText = colName;
  if (typeVal) typeVal.innerText = colType;
  if (nullVal) nullVal.innerText = isNull ? 'Yes' : 'No';
  if (idVal) idVal.innerText = isPk ? 'Yes (Primary Key)' : 'No';
}

function handleSaveTableDesigner() {
  const oldTableName = state.designer.editingTableName;
  const newTableName = (document.getElementById('designerTableNameInput').value || '').trim();

  if (!newTableName) {
    showToast('Please specify a table name', 'danger');
    return;
  }

  const trs = document.querySelectorAll('#designerColumnsTbody tr');
  const newColumns = [];

  trs.forEach(tr => {
    const nameInput = tr.querySelector('.designer-col-name');
    const colName = nameInput.value.trim();
    const oldColName = (nameInput.getAttribute('data-old-name') || '').trim();
    const colType = tr.querySelector('.designer-col-type').value;
    const pk = tr.querySelector('.col-pk-check').checked;
    const nn = !tr.querySelector('.designer-col-null').checked; // Allow null unchecked means NOT NULL
    const uq = tr.querySelector('.designer-col-uq').checked;
    const def = tr.querySelector('.designer-col-def').value.trim();

    if (colName) {
      newColumns.push({ name: colName, oldName: oldColName, type: colType, pk, nn, uq, def });
    }
  });

  if (newColumns.length === 0) {
    showToast('Please define at least one column for the table', 'danger');
    return;
  }

  // Preserve existing rows and migrate data safely
  let finalRows = [];
  let existingFks = [];

  if (oldTableName && state.tables[oldTableName]) {
    const oldRows = state.tables[oldTableName].rows || [];
    existingFks = state.tables[oldTableName].foreignKeys || [];

    finalRows = oldRows.map(row => {
      const migratedRow = {};
      newColumns.forEach(newCol => {
        if (row[newCol.name] !== undefined) {
          migratedRow[newCol.name] = row[newCol.name];
        } else if (newCol.oldName && row[newCol.oldName] !== undefined) {
          migratedRow[newCol.name] = row[newCol.oldName];
        } else {
          migratedRow[newCol.name] = newCol.def || '';
        }
      });
      return migratedRow;
    });

    if (oldTableName !== newTableName) {
      delete state.tables[oldTableName];
      if (state.diagram.tablePositions[oldTableName]) {
        state.diagram.tablePositions[newTableName] = state.diagram.tablePositions[oldTableName];
        delete state.diagram.tablePositions[oldTableName];
      }
      // Update foreign keys pointing to old table
      Object.keys(state.tables).forEach(t => {
        (state.tables[t].foreignKeys || []).forEach(fk => {
          if (fk.refTable === oldTableName) fk.refTable = newTableName;
        });
      });
    }
  }

  const cleanedColumns = newColumns.map(({ name, type, pk, nn, uq, def }) => ({ name, type, pk, nn, uq, def }));

  state.tables[newTableName] = {
    columns: cleanedColumns,
    foreignKeys: existingFks,
    rows: finalRows
  };
  state.activeTable = newTableName;
  state.designer.editingTableName = newTableName;
  state.designer.isNew = false;

  // Update tabs
  const curTab = state.tabs.find(t => t.id === state.activeTabId);
  if (curTab) {
    curTab.id = `design:${newTableName}`;
    curTab.name = `dbo.${newTableName} [Design]`;
    curTab.tableName = newTableName;
    state.activeTabId = curTab.id;
  }

  saveDatabaseToStorage();
  renderObjectExplorer();
  renderTabs();
  showToast(`Table "dbo.${newTableName}" schema saved successfully!`, 'success');
}

/** =========================================================================
 * SSMS TABLE DATA EDITOR ("Edit Top 200 Rows")
 * ========================================================================= */

function openTableDataTab(tableName) {
  closeAllMenus();
  if (!state.tables[tableName]) return;

  state.activeTable = tableName;
  state.activeView = 'table-contents';

  const tabId = `table:${tableName}`;
  if (!state.tabs.find(t => t.id === tabId)) {
    state.tabs.push({ id: tabId, type: 'table', name: `dbo.${tableName} [Edit]`, tableName: tableName, icon: 'fa-table' });
  }
  state.activeTabId = tabId;

  renderObjectExplorer();
  renderTabs();
  renderActiveView();
}

function renderTableContents() {
  const tableName = state.activeTable;
  const tableData = state.tables[tableName];

  const titleEl = document.getElementById('activeTableName');
  if (titleEl) titleEl.innerHTML = `<i class="fa-solid fa-table text-accent"></i> dbo.${tableName || 'No Table'}`;

  const rowCountEl = document.getElementById('activeTableRowCount');
  const colCountEl = document.getElementById('activeTableColCount');

  if (!tableData) {
    document.getElementById('mainDataTableHeader').innerHTML = '';
    document.getElementById('mainDataTableBody').innerHTML = `<tr><td colspan="10" class="p-4 text-muted">Table "${tableName}" not found.</td></tr>`;
    if (rowCountEl) rowCountEl.innerText = '0 rows';
    if (colCountEl) colCountEl.innerText = '0 cols';
    return;
  }

  const columns = tableData.columns || [];
  let rows = [...(tableData.rows || [])];

  if (rowCountEl) rowCountEl.innerText = `${rows.length} rows`;
  if (colCountEl) colCountEl.innerText = `${columns.length} cols`;

  // Render Header
  const headerTr = document.getElementById('mainDataTableHeader');
  headerTr.innerHTML = '<th style="width: 44px; text-align: center;">#</th>';

  columns.forEach(col => {
    let sortIndicator = '';
    if (state.sort.column === col.name) {
      sortIndicator = `<span class="text-accent ml-1">${state.sort.direction === 'asc' ? '▲' : '▼'}</span>`;
    }

    let badges = '';
    if (col.pk) badges += '<span class="badge-pk ml-1" title="Primary Key"><i class="fa-solid fa-key"></i></span>';
    if (col.uq) badges += '<span class="text-info ml-1" style="font-size: 9px; font-weight: 700;">UQ</span>';
    if (col.nn) badges += '<span class="text-danger ml-1" style="font-size: 9px; font-weight: 700;">*</span>';

    const isFk = (tableData.foreignKeys || []).some(fk => fk.column === col.name);
    if (isFk) badges += '<span class="badge-fk ml-1" title="Foreign Key"><i class="fa-solid fa-link"></i></span>';

    headerTr.innerHTML += `
      <th class="sortable-th" onclick="sortTableByColumn('${col.name}')" title="Sort by ${col.name}">
        <span>${col.name}</span>
        <small class="text-muted" style="font-size: 10px; font-weight: normal;">(${col.type})</small>
        ${badges}
        ${sortIndicator}
      </th>
    `;
  });
  headerTr.innerHTML += `<th style="width: 70px; text-align: center;">Actions</th>`;

  // Sorting
  if (state.sort.column) {
    const colName = state.sort.column;
    const dir = state.sort.direction === 'asc' ? 1 : -1;
    rows.sort((a, b) => {
      let valA = a[colName] !== undefined ? a[colName] : '';
      let valB = b[colName] !== undefined ? b[colName] : '';
      if (typeof valA === 'number' && typeof valB === 'number') return (valA - valB) * dir;
      return String(valA).localeCompare(String(valB)) * dir;
    });
  }

  // Pagination
  const totalRows = rows.length;
  const pageSize = state.pagination.pageSize === 'all' ? (totalRows || 1) : parseInt(state.pagination.pageSize, 10);
  const totalPages = Math.max(1, Math.ceil(totalRows / pageSize));

  if (state.pagination.currentPage > totalPages) state.pagination.currentPage = totalPages;
  if (state.pagination.currentPage < 1) state.pagination.currentPage = 1;

  const startIndex = (state.pagination.currentPage - 1) * pageSize;
  const pageRows = rows.slice(startIndex, startIndex + pageSize);

  renderTableRows(pageRows, columns, startIndex);

  const summary = document.getElementById('gridFooterSummary');
  if (summary) summary.innerText = `Showing ${totalRows === 0 ? 0 : startIndex + 1}-${Math.min(startIndex + pageSize, totalRows)} of ${totalRows} records`;

  const pageText = document.getElementById('pageIndicatorText');
  if (pageText) pageText.innerText = `Page ${state.pagination.currentPage} of ${totalPages}`;

  const sbRows = document.getElementById('sbRowCount');
  if (sbRows) sbRows.innerText = `${totalRows} rows`;
}

function renderTableRows(rows, columns, offset = 0) {
  const body = document.getElementById('mainDataTableBody');
  body.innerHTML = '';

  const tableName = state.activeTable;
  const fks = (state.tables[tableName] && state.tables[tableName].foreignKeys) ? state.tables[tableName].foreignKeys : [];

  if (rows.length === 0) {
    body.innerHTML = `<tr><td colspan="${columns.length + 2}" class="p-4 text-muted text-center">No rows. Click "+ Add Row" to insert data into dbo.${tableName}.</td></tr>`;
    return;
  }

  rows.forEach((row, index) => {
    const originalRowIndex = offset + index;
    const tr = document.createElement('tr');
    tr.innerHTML = `<td style="text-align: center; color: var(--text-muted);">${originalRowIndex + 1}</td>`;

    columns.forEach(col => {
      const val = row[col.name] !== undefined ? row[col.name] : '';
      const fkMatch = fks.find(fk => fk.column === col.name);

      let cellContent = escapeHtml(val);
      if (fkMatch && val !== '' && val !== null && val !== undefined) {
        cellContent = `
          <span class="fk-pill" title="Foreign Key referencing ${fkMatch.refTable}.${fkMatch.refColumn}" onclick="event.stopPropagation(); navigateToRelatedRow('${fkMatch.refTable}', '${fkMatch.refColumn}', '${escapeHtml(val)}')">
            <i class="fa-solid fa-link"></i> ${escapeHtml(val)} &rarr; ${fkMatch.refTable}
          </span>
        `;
      }

      tr.innerHTML += `
        <td class="editable-cell" 
            data-row-idx="${originalRowIndex}" 
            data-col-name="${col.name}" 
            ondblclick="startInlineEdit(this, ${originalRowIndex}, '${col.name}')" 
            title="Double-click to edit cell">
          ${cellContent}
        </td>
      `;
    });

    tr.innerHTML += `
      <td style="text-align: center;">
        <div class="flex-center" style="justify-content: center; gap: 4px;">
          <button class="btn-icon-xs" title="Edit Record" onclick="openEditRecordModal(${originalRowIndex})"><i class="fa-solid fa-pen"></i></button>
          <button class="btn-icon-xs text-danger" title="Delete Record" onclick="deleteRecordRow(${originalRowIndex})"><i class="fa-solid fa-trash-can"></i></button>
        </div>
      </td>
    `;
    body.appendChild(tr);
  });
}

function startInlineEdit(td, rowIndex, colName) {
  if (td.querySelector('input')) return;

  const tableName = state.activeTable;
  const tableData = state.tables[tableName];
  if (!tableData || !tableData.rows[rowIndex]) return;

  const currentVal = tableData.rows[rowIndex][colName] !== undefined ? tableData.rows[rowIndex][colName] : '';
  const colDef = tableData.columns.find(c => c.name === colName);

  const input = document.createElement('input');
  input.type = (colDef && (colDef.type === 'INTEGER' || colDef.type === 'REAL')) ? 'number' : 'text';
  input.className = 'inline-cell-input';
  input.value = currentVal;

  td.innerHTML = '';
  td.appendChild(input);
  input.focus();
  input.select();

  const commitEdit = () => {
    let newVal = input.value.trim();
    if (input.type === 'number') newVal = newVal !== '' ? Number(newVal) : '';

    if (colDef && colDef.nn && (newVal === '' || newVal === null)) {
      showToast(`Column "${colName}" cannot be null`, 'danger');
      renderTableContents();
      return;
    }

    if (colDef && colDef.uq && newVal !== '') {
      const exists = tableData.rows.some((r, idx) => idx !== rowIndex && String(r[colName]) === String(newVal));
      if (exists) {
        showToast(`UNIQUE violation on "${colName}"`, 'danger');
        renderTableContents();
        return;
      }
    }

    tableData.rows[rowIndex][colName] = newVal;
    saveDatabaseToStorage();
    renderTableContents();
  };

  input.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') commitEdit();
    else if (e.key === 'Escape') renderTableContents();
  });
  input.addEventListener('blur', commitEdit);
}

function deleteRecordRow(index) {
  const tableName = state.activeTable;
  const tableData = state.tables[tableName];
  if (!tableData || !tableData.rows[index]) return;

  if (confirm(`Delete row #${index + 1} from dbo.${tableName}?`)) {
    tableData.rows.splice(index, 1);
    saveDatabaseToStorage();
    renderTableContents();
    renderObjectExplorer();
    showToast(`Deleted row #${index + 1}`, 'info');
  }
}

function openAddRecordModal() {
  const tableName = state.activeTable;
  const tableData = state.tables[tableName];
  if (!tableData) return;

  state.editingRowIndex = null;
  document.getElementById('recordModalTitle').innerText = `Add Record to dbo.${tableName}`;
  document.getElementById('recordModalSub').innerText = 'Enter values for row fields';

  const defaultObj = {};
  tableData.columns.forEach(col => {
    if (col.def) defaultObj[col.name] = col.def;
  });

  renderRecordFormFields(defaultObj);
  openModal('recordModal');
}

function openEditRecordModal(rowIndex) {
  const tableName = state.activeTable;
  const tableData = state.tables[tableName];
  if (!tableData || !tableData.rows[rowIndex]) return;

  state.editingRowIndex = rowIndex;
  document.getElementById('recordModalTitle').innerText = `Edit Record in dbo.${tableName}`;
  document.getElementById('recordModalSub').innerText = `Row #${rowIndex + 1}`;

  renderRecordFormFields(tableData.rows[rowIndex]);
  openModal('recordModal');
}

function renderRecordFormFields(rowObj) {
  const tableName = state.activeTable;
  const columns = state.tables[tableName].columns;
  const fks = state.tables[tableName].foreignKeys || [];

  const container = document.getElementById('recordFieldsContainer');
  container.innerHTML = '';

  columns.forEach(col => {
    const val = rowObj[col.name] !== undefined ? rowObj[col.name] : '';
    const fkMatch = fks.find(fk => fk.column === col.name);

    const fieldGroup = document.createElement('div');
    fieldGroup.className = 'form-group';

    let badges = '';
    if (col.pk) badges += ' <span class="badge-pk">(PK)</span>';
    if (col.uq) badges += ' <span class="text-info">(UQ)</span>';
    if (col.nn) badges += ' <span class="text-danger">*</span>';

    let fkHint = '';
    if (fkMatch) {
      fkHint = `<small class="form-help text-accent"><i class="fa-solid fa-link"></i> References ${fkMatch.refTable}.${fkMatch.refColumn}</small>`;
    }

    fieldGroup.innerHTML = `
      <label>${col.name} <small class="text-muted">(${col.type})</small>${badges}</label>
      <input type="${col.type === 'INTEGER' || col.type === 'REAL' ? 'number' : 'text'}" 
             class="form-control record-input-field" 
             data-col-name="${col.name}" 
             value="${val}" 
             ${col.nn && !col.pk ? 'required' : ''}>
      ${fkHint}
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

  for (let col of tableData.columns) {
    if (col.uq && recordRow[col.name] !== undefined && recordRow[col.name] !== '') {
      const exists = tableData.rows.some((r, idx) => {
        if (state.editingRowIndex !== null && idx === state.editingRowIndex) return false;
        return String(r[col.name]) === String(recordRow[col.name]);
      });
      if (exists) {
        showToast(`UNIQUE violation on "${col.name}" with value "${recordRow[col.name]}"`, 'danger');
        return;
      }
    }
  }

  if (state.editingRowIndex !== null) {
    tableData.rows[state.editingRowIndex] = recordRow;
    showToast(`Updated row in dbo.${tableName}`, 'success');
  } else {
    tableData.rows.push(recordRow);
    showToast(`Inserted row into dbo.${tableName}`, 'success');
  }

  closeModal('recordModal');
  saveDatabaseToStorage();
  renderTableContents();
  renderObjectExplorer();
}

function navigateToRelatedRow(targetTable, targetCol, value) {
  if (!state.tables[targetTable]) return;
  openTableDataTab(targetTable);
  const searchInput = document.getElementById('tableDataSearchInput');
  if (searchInput) {
    searchInput.value = String(value);
    filterTableDataRows(String(value));
  }
  showToast(`Jumped to dbo.${targetTable} where ${targetCol} = "${value}"`, 'info');
}

function filterTableDataRows(query) {
  const q = query.trim().toLowerCase();
  const rows = document.querySelectorAll('#mainDataTableBody tr');
  rows.forEach(r => {
    const text = r.innerText.toLowerCase();
    r.style.display = text.includes(q) ? '' : 'none';
  });
}

function sortTableByColumn(colName) {
  if (state.sort.column === colName) {
    if (state.sort.direction === 'asc') state.sort.direction = 'desc';
    else { state.sort.column = null; state.sort.direction = 'asc'; }
  } else {
    state.sort.column = colName;
    state.sort.direction = 'asc';
  }
  renderTableContents();
}

function changePageSize(val) {
  state.pagination.pageSize = val;
  state.pagination.currentPage = 1;
  renderTableContents();
}

function goToPage(action) {
  const tableName = state.activeTable;
  const totalRows = (state.tables[tableName] && state.tables[tableName].rows) ? state.tables[tableName].rows.length : 0;
  const pageSize = state.pagination.pageSize === 'all' ? (totalRows || 1) : parseInt(state.pagination.pageSize, 10);
  const totalPages = Math.max(1, Math.ceil(totalRows / pageSize));

  if (action === 'first') state.pagination.currentPage = 1;
  else if (action === 'prev') state.pagination.currentPage = Math.max(1, state.pagination.currentPage - 1);
  else if (action === 'next') state.pagination.currentPage = Math.min(totalPages, state.pagination.currentPage + 1);
  else if (action === 'last') state.pagination.currentPage = totalPages;

  renderTableContents();
}

/** =========================================================================
 * SSMS DOCUMENT TABS ARCHITECTURE (MDI)
 * ========================================================================= */

function openNewQueryTab(initialSql = '') {
  closeAllMenus();
  state.queryCounter++;
  const tabId = `query:${state.queryCounter}`;
  const tabName = `SQLQuery${state.queryCounter}.sql`;

  state.tabs.push({ id: tabId, type: 'query', name: tabName, sql: initialSql, icon: 'fa-file-code' });
  state.activeTabId = tabId;
  state.activeView = 'sql-query';

  renderTabs();
  renderActiveView();

  const input = document.getElementById('consoleQueryInput');
  if (input) {
    input.value = initialSql || `-- SQL Query (${tabName})\nSELECT * FROM Users;`;
    input.focus();
    setupLineNumbers();
  }
}

function openErDiagramTab() {
  closeAllMenus();
  state.activeView = 'er-diagram';
  const tabId = 'er-diagram';
  if (!state.tabs.find(t => t.id === tabId)) {
    state.tabs.push({ id: tabId, type: 'diagram', name: 'Database Diagrams', icon: 'fa-diagram-project' });
  }
  state.activeTabId = tabId;

  renderTabs();
  renderActiveView();
}

function switchActiveTab(tabId) {
  state.activeTabId = tabId;
  const tab = state.tabs.find(t => t.id === tabId);
  if (!tab) return;

  if (tab.type === 'query') {
    state.activeView = 'sql-query';
  } else if (tab.type === 'table') {
    state.activeTable = tab.tableName;
    state.activeView = 'table-contents';
  } else if (tab.type === 'designer') {
    state.activeView = 'table-designer';
    state.designer.editingTableName = tab.tableName;
  } else if (tab.type === 'diagram') {
    state.activeView = 'er-diagram';
  }

  renderTabs();
  renderActiveView();
  renderObjectExplorer();
}

function closeTab(e, tabId) {
  if (e && e.stopPropagation) e.stopPropagation();
  state.tabs = state.tabs.filter(t => t.id !== tabId);

  if (state.activeTabId === tabId) {
    if (state.tabs.length > 0) {
      const next = state.tabs[state.tabs.length - 1];
      switchActiveTab(next.id);
    } else {
      state.activeTabId = null;
      state.activeView = 'empty-workspace';
      renderTabs();
      renderActiveView();
    }
  } else {
    renderTabs();
  }
}

function renderTabs() {
  const container = document.getElementById('tabItemsWrapper');
  if (!container) return;
  container.innerHTML = '';

  state.tabs.forEach(tab => {
    const isActive = tab.id === state.activeTabId;
    const tabEl = document.createElement('div');
    tabEl.className = `ssms-tab-item ${isActive ? 'active' : ''}`;
    tabEl.onclick = () => switchActiveTab(tab.id);

    tabEl.innerHTML = `
      <i class="fa-solid ${tab.icon}"></i>
      <span>${tab.name}</span>
      <button class="ssms-tab-close" onclick="closeTab(event, '${tab.id}')" title="Close"><i class="fa-solid fa-xmark"></i></button>
    `;
    container.appendChild(tabEl);
  });
}

function renderActiveView() {
  document.querySelectorAll('.ssms-view-page').forEach(v => v.classList.remove('active-view'));

  if (state.activeView === 'empty-workspace') {
    const view = document.getElementById('view-empty-workspace');
    if (view) view.classList.add('active-view');
  } else if (state.activeView === 'sql-query') {
    const view = document.getElementById('view-sql-query');
    if (view) view.classList.add('active-view');
  } else if (state.activeView === 'table-contents') {
    const view = document.getElementById('view-table-contents');
    if (view) view.classList.add('active-view');
    renderTableContents();
  } else if (state.activeView === 'table-designer') {
    const view = document.getElementById('view-table-designer');
    if (view) view.classList.add('active-view');
  } else if (state.activeView === 'er-diagram') {
    const view = document.getElementById('view-er-diagram');
    if (view) view.classList.add('active-view');
    renderErDiagram();
  }
}

/** =========================================================================
 * SSMS SQL QUERY CONSOLE & INTERPRETER
 * ========================================================================= */

function setupLineNumbers() {
  const textarea = document.getElementById('consoleQueryInput');
  const lineNumbers = document.getElementById('queryLineNumbers');
  if (!textarea || !lineNumbers) return;

  const update = () => {
    const lines = (textarea.value || '').split('\n').length;
    let numbersHtml = '';
    for (let i = 1; i <= Math.max(lines, 1); i++) {
      numbersHtml += `${i}<br>`;
    }
    lineNumbers.innerHTML = numbersHtml;
  };

  textarea.addEventListener('input', update);
  textarea.addEventListener('scroll', () => {
    lineNumbers.scrollTop = textarea.scrollTop;
  });
  update();
}

function runConsoleSqlQuery() {
  const input = document.getElementById('consoleQueryInput');
  const rawSql = input ? input.value.trim() : '';

  if (!rawSql) {
    showToast('Enter a SQL query to execute', 'info');
    return;
  }

  const startTime = performance.now();
  const timer = document.getElementById('queryExecTimer');
  if (timer) timer.innerText = 'Executing query...';

  try {
    const result = interpretSqlQuery(rawSql);
    const duration = ((performance.now() - startTime) / 1000).toFixed(2);

    if (timer) timer.innerText = `Query executed in ${duration}s`;
    recordQueryHistory(rawSql, true);

    const msgBox = document.getElementById('consoleQueryMessages');
    if (msgBox) {
      msgBox.innerText = `(${result.rowCount || 0} row(s) affected)\nQuery executed successfully.\nExecution time: 00:00:${duration}`;
    }

    if (result.type === 'SELECT') {
      renderQueryResultsGrid(result.columns, result.rows);
      switchOutputTab('results');
    } else {
      switchOutputTab('messages');
      renderObjectExplorer();
      populateToolbarDbSelect();
      if (state.activeTable && state.tables[state.activeTable]) {
        renderTableContents();
      }
    }
    showToast('Query executed successfully.', 'success');
  } catch (err) {
    const duration = ((performance.now() - startTime) / 1000).toFixed(2);
    if (timer) timer.innerText = `Failed (${duration}s)`;
    recordQueryHistory(rawSql, false);

    const msgBox = document.getElementById('consoleQueryMessages');
    if (msgBox) {
      msgBox.innerText = `Msg 102, Level 15, State 1:\n${err.message}`;
    }
    switchOutputTab('messages');
    showToast(`SQL Error: ${err.message}`, 'danger');
  }
}

function cancelExecutingQuery() {
  const timer = document.getElementById('queryExecTimer');
  if (timer) timer.innerText = 'Canceled by user.';
  showToast('Query execution canceled.', 'info');
}

function clearConsoleQuery() {
  const input = document.getElementById('consoleQueryInput');
  if (input) {
    input.value = '';
    setupLineNumbers();
  }
}

function switchOutputTab(tabType) {
  const btnRes = document.getElementById('tabOutputResults');
  const btnMsg = document.getElementById('tabOutputMessages');
  const paneRes = document.getElementById('outputPaneResults');
  const paneMsg = document.getElementById('outputPaneMessages');

  if (tabType === 'results') {
    if (btnRes) btnRes.classList.add('active');
    if (btnMsg) btnMsg.classList.remove('active');
    if (paneRes) paneRes.classList.add('active');
    if (paneMsg) paneMsg.classList.remove('active');
  } else {
    if (btnRes) btnRes.classList.remove('active');
    if (btnMsg) btnMsg.classList.add('active');
    if (paneRes) paneRes.classList.remove('active');
    if (paneMsg) paneMsg.classList.add('active');
  }
}

function renderQueryResultsGrid(columns, rows) {
  const container = document.getElementById('consoleQueryResults');
  if (!container) return;

  if (!columns || columns.length === 0) {
    container.innerHTML = '<div class="p-3 text-muted">Query executed successfully. (0 columns returned)</div>';
    return;
  }

  let tableHtml = `<table class="grid-table"><thead><tr><th style="width:40px; text-align:center;">#</th>`;
  columns.forEach(c => tableHtml += `<th>${escapeHtml(c)}</th>`);
  tableHtml += `</tr></thead><tbody>`;

  if (rows.length === 0) {
    tableHtml += `<tr><td colspan="${columns.length + 1}" class="p-3 text-muted text-center">0 rows returned.</td></tr>`;
  } else {
    rows.forEach((r, idx) => {
      tableHtml += `<tr><td style="text-align:center; color:var(--text-muted);">${idx + 1}</td>`;
      columns.forEach(c => {
        tableHtml += `<td>${escapeHtml(r[c] !== undefined ? r[c] : 'NULL')}</td>`;
      });
      tableHtml += `</tr>`;
    });
  }
  tableHtml += `</tbody></table>`;
  container.innerHTML = tableHtml;
}

function interpretSqlQuery(sql) {
  const clean = sql.replace(/--.*$/gm, '').trim();
  const firstWord = clean.split(/\s+/)[0].toUpperCase();

  if (firstWord === 'SELECT') return executeSelectQuery(clean);
  if (firstWord === 'INSERT') return executeInsertQuery(clean);
  if (firstWord === 'UPDATE') return executeUpdateQuery(clean);
  if (firstWord === 'DELETE') return executeDeleteQuery(clean);
  if (firstWord === 'CREATE') return executeCreateQuery(clean);
  if (firstWord === 'DROP') return executeDropQuery(clean);
  if (firstWord === 'ALTER') return executeAlterQuery(clean);

  throw new Error(`Unsupported SQL command "${firstWord}". Supported: SELECT, INSERT, UPDATE, DELETE, CREATE, DROP, ALTER.`);
}

function executeSelectQuery(sql) {
  const match = sql.match(/SELECT\s+(?:TOP\s+(\d+)\s+)?([\s\S]+?)\s+FROM\s+(?:\[dbo\]\.)?\[?([a-zA-Z0-9_]+)\]?(?:\s+WHERE\s+([\s\S]+?))?(?:\s+ORDER\s+BY\s+([\s\S]+?))?$/i);
  if (!match) throw new Error("Invalid SELECT syntax. Example: SELECT * FROM Users; or SELECT TOP 10 * FROM Users;");

  const topLimit = match[1] ? parseInt(match[1], 10) : null;
  const colClause = match[2].trim();
  const tableName = match[3].replace(/[\[\]]/g, '').trim();
  const whereClause = match[4] ? match[4].trim() : null;

  if (!state.tables[tableName]) throw new Error(`Table "dbo.${tableName}" does not exist in ${state.dbName}.`);

  const tableData = state.tables[tableName];
  let resultRows = [...tableData.rows];

  if (whereClause) {
    resultRows = resultRows.filter(row => evaluateWhereCondition(row, whereClause));
  }

  if (topLimit !== null) {
    resultRows = resultRows.slice(0, topLimit);
  }

  let finalColumns = [];
  if (colClause === '*') {
    finalColumns = tableData.columns.map(c => c.name);
  } else {
    finalColumns = colClause.split(',').map(c => c.trim().replace(/[\[\]]/g, ''));
  }

  return { type: 'SELECT', columns: finalColumns, rows: resultRows, rowCount: resultRows.length };
}

function evaluateWhereCondition(row, condition) {
  try {
    const eqMatch = condition.match(/\[?(\w+)\]?\s*=\s*(?:'([^']*)'|(\d+))/);
    if (eqMatch) {
      const col = eqMatch[1];
      const val = eqMatch[2] !== undefined ? eqMatch[2] : Number(eqMatch[3]);
      return String(row[col]) === String(val);
    }
    return true;
  } catch (e) {
    return true;
  }
}

function executeInsertQuery(sql) {
  const match = sql.match(/INSERT\s+INTO\s+(?:\[dbo\]\.)?\[?(\w+)\]?\s*\((.+?)\)\s*VALUES\s*\((.+?)\)/i);
  if (!match) throw new Error("Invalid INSERT syntax. Example: INSERT INTO Users (name, email) VALUES ('Alex', 'alex@mail.com');");

  const tableName = match[1];
  if (!state.tables[tableName]) throw new Error(`Table "dbo.${tableName}" does not exist.`);

  const cols = match[2].split(',').map(c => c.trim().replace(/[\[\]]/g, ''));
  const vals = match[3].split(',').map(v => v.trim().replace(/^'|'$/g, ''));

  const newRow = {};
  state.tables[tableName].columns.forEach(c => newRow[c.name] = c.def || '');
  cols.forEach((col, idx) => {
    newRow[col] = vals[idx] !== undefined ? vals[idx] : '';
  });

  state.tables[tableName].rows.push(newRow);
  saveDatabaseToStorage();
  return { type: 'INSERT', rowCount: 1 };
}

function executeUpdateQuery(sql) {
  const match = sql.match(/UPDATE\s+(?:\[dbo\]\.)?\[?(\w+)\]?\s+SET\s+(.+?)(?:\s+WHERE\s+(.+?))?$/i);
  if (!match) throw new Error("Invalid UPDATE syntax. Example: UPDATE Users SET role = 'Admin' WHERE id = 1;");

  const tableName = match[1];
  if (!state.tables[tableName]) throw new Error(`Table "dbo.${tableName}" does not exist.`);

  const setStr = match[2];
  const whereClause = match[3] ? match[3].trim() : null;

  const sets = {};
  setStr.split(',').forEach(s => {
    const p = s.split('=');
    if (p.length === 2) sets[p[0].trim().replace(/[\[\]]/g, '')] = p[1].trim().replace(/^'|'$/g, '');
  });

  let count = 0;
  state.tables[tableName].rows.forEach(row => {
    if (!whereClause || evaluateWhereCondition(row, whereClause)) {
      Object.assign(row, sets);
      count++;
    }
  });

  saveDatabaseToStorage();
  return { type: 'UPDATE', rowCount: count };
}

function executeDeleteQuery(sql) {
  const match = sql.match(/DELETE\s+FROM\s+(?:\[dbo\]\.)?\[?(\w+)\]?(?:\s+WHERE\s+(.+?))?$/i);
  if (!match) throw new Error("Invalid DELETE syntax. Example: DELETE FROM Users WHERE id = 2;");

  const tableName = match[1];
  if (!state.tables[tableName]) throw new Error(`Table "dbo.${tableName}" does not exist.`);

  const whereClause = match[2] ? match[2].trim() : null;
  const initial = state.tables[tableName].rows.length;

  if (whereClause) {
    state.tables[tableName].rows = state.tables[tableName].rows.filter(r => !evaluateWhereCondition(r, whereClause));
  } else {
    state.tables[tableName].rows = [];
  }

  const affected = initial - state.tables[tableName].rows.length;
  saveDatabaseToStorage();
  return { type: 'DELETE', rowCount: affected };
}

function executeCreateQuery(sql) {
  const match = sql.match(/CREATE\s+TABLE\s+(?:\[dbo\]\.)?\[?(\w+)\]?\s*\((.+?)\)/i);
  if (!match) throw new Error("Invalid CREATE TABLE syntax. Usage: CREATE TABLE TableName (id int primary key, name varchar(50) not null);");

  const tableName = match[1];
  const colDefsStr = match[2];

  const cols = colDefsStr.split(',').map(def => {
    const parts = def.trim().split(/\s+/);
    const name = parts[0].replace(/[\[\]]/g, '');
    const type = (parts[1] || 'TEXT').toUpperCase();
    const pk = def.toUpperCase().includes('PRIMARY KEY');
    const nn = def.toUpperCase().includes('NOT NULL');
    return { name, type, pk, nn, uq: pk, def: '' };
  });

  state.tables[tableName] = { columns: cols, foreignKeys: [], rows: [] };
  saveDatabaseToStorage();
  return { type: 'CREATE', rowCount: 0 };
}

function executeDropQuery(sql) {
  const match = sql.match(/DROP\s+TABLE\s+(?:\[dbo\]\.)?\[?(\w+)\]?/i);
  if (!match) throw new Error("Invalid DROP TABLE syntax. Usage: DROP TABLE TableName;");

  const tableName = match[1];
  if (!state.tables[tableName]) throw new Error(`Table "${tableName}" does not exist.`);

  deleteTableByName(tableName);
  return { type: 'DROP', rowCount: 0 };
}

function executeAlterQuery(sql) {
  const match = sql.match(/ALTER\s+TABLE\s+(?:\[dbo\]\.)?\[?(\w+)\]?\s+ADD\s+(?:COLUMN\s+)?(\w+)\s+(\w+)/i);
  if (!match) throw new Error("Invalid ALTER TABLE syntax. Usage: ALTER TABLE TableName ADD COLUMN colName int;");

  const tableName = match[1];
  const colName = match[2];
  const colType = match[3].toUpperCase();

  if (!state.tables[tableName]) throw new Error(`Table "${tableName}" does not exist.`);
  state.tables[tableName].columns.push({ name: colName, type: colType, pk: false, nn: false, uq: false, def: '' });
  saveDatabaseToStorage();
  return { type: 'ALTER', rowCount: 0 };
}

function recordQueryHistory(sql, success) {
  state.queryHistory.unshift({
    sql: sql,
    time: new Date().toLocaleTimeString(),
    success: success
  });
  if (state.queryHistory.length > 50) state.queryHistory.pop();
  try {
    localStorage.setItem('studio_sql_history', JSON.stringify(state.queryHistory));
  } catch (e) {}
  renderQueryHistory();
}

function loadQueryHistory() {
  try {
    const saved = localStorage.getItem('studio_sql_history');
    if (saved) state.queryHistory = JSON.parse(saved);
  } catch (e) {}
}

function renderQueryHistory() {
  const container = document.getElementById('queryHistoryList');
  if (!container) return;
  container.innerHTML = '';

  state.queryHistory.forEach(item => {
    const el = document.createElement('div');
    el.className = 'history-item';
    el.onclick = () => {
      const input = document.getElementById('consoleQueryInput');
      if (input) {
        input.value = item.sql;
        setupLineNumbers();
      }
    };
    el.innerHTML = `
      <div class="history-query">${escapeHtml(item.sql)}</div>
      <div class="history-time"><i class="fa-solid ${item.success ? 'fa-check text-success' : 'fa-xmark text-danger'}"></i> ${item.time}</div>
    `;
    container.appendChild(el);
  });
}

function toggleQueryHistoryPanel() {
  const panel = document.getElementById('queryHistorySidebar');
  if (panel) panel.classList.toggle('hidden');
}

function clearQueryHistory() {
  state.queryHistory = [];
  try { localStorage.removeItem('studio_sql_history'); } catch (e) {}
  renderQueryHistory();
}

function exportQueryResultsCsv() {
  const container = document.getElementById('consoleQueryResults');
  const table = container ? container.querySelector('table') : null;
  if (!table) {
    showToast('No query results to export.', 'info');
    return;
  }
  let csv = '';
  table.querySelectorAll('tr').forEach(row => {
    const rowData = [];
    row.querySelectorAll('th, td').forEach(cell => rowData.push(`"${cell.innerText.replace(/"/g, '""')}"`));
    csv += rowData.join(',') + '\n';
  });

  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `QueryResults_${Date.now()}.csv`;
  document.body.appendChild(a);
  a.click();
  a.remove();
}

/** =========================================================================
 * SSMS RELATIONAL DATABASE DIAGRAMS
 * ========================================================================= */

function renderErDiagram() {
  const tablesContainer = document.getElementById('diagramTablesContainer');
  const statsLabel = document.getElementById('diagramStats');
  if (!tablesContainer) return;

  tablesContainer.innerHTML = '';
  const tableKeys = Object.keys(state.tables);

  let totalRelationships = 0;
  tableKeys.forEach(tName => {
    totalRelationships += (state.tables[tName].foreignKeys || []).length;
  });

  if (statsLabel) {
    statsLabel.innerText = `${tableKeys.length} Tables, ${totalRelationships} Relationship${totalRelationships === 1 ? '' : 's'}`;
  }

  let defaultX = 50, defaultY = 50;
  tableKeys.forEach((tName, idx) => {
    if (!state.diagram.tablePositions[tName]) {
      state.diagram.tablePositions[tName] = {
        x: defaultX + (idx % 3) * 280,
        y: defaultY + Math.floor(idx / 3) * 260
      };
    }
  });

  tableKeys.forEach(tName => {
    const tableData = state.tables[tName];
    const pos = state.diagram.tablePositions[tName];

    const card = document.createElement('div');
    card.id = `diagram-card-${tName}`;
    card.className = 'diagram-table-card';
    card.style.left = `${pos.x}px`;
    card.style.top = `${pos.y}px`;

    let columnsHtml = '';
    (tableData.columns || []).forEach(col => {
      const isFk = (tableData.foreignKeys || []).some(fk => fk.column === col.name);
      columnsHtml += `
        <div class="card-col-row" id="col-row-${tName}-${col.name}">
          <div class="flex-center gap-1">
            ${col.pk ? '<i class="fa-solid fa-key badge-pk"></i>' : ''}
            ${isFk ? '<i class="fa-solid fa-link badge-fk"></i>' : ''}
            <span>${col.name}</span>
          </div>
          <span class="card-col-type">${col.type}</span>
        </div>
      `;
    });

    card.innerHTML = `
      <div class="card-header">
        <div class="flex-center gap-1">
          <i class="fa-solid fa-table text-accent"></i>
          <span>dbo.${tName}</span>
        </div>
        <div class="flex-center gap-1">
          <button class="btn-icon-xs" title="Edit Top 200 Rows" onclick="event.stopPropagation(); openTableDataTab('${tName}')"><i class="fa-solid fa-table"></i></button>
          <button class="btn-icon-xs" title="Design Table" onclick="event.stopPropagation(); openTableDesignerForTable('${tName}')"><i class="fa-solid fa-table-columns"></i></button>
          <button class="btn-icon-xs text-danger" title="Delete Table" onclick="event.stopPropagation(); promptDeleteTable('${tName}')"><i class="fa-solid fa-trash-can"></i></button>
        </div>
      </div>
      <div class="card-columns">
        ${columnsHtml}
      </div>
    `;

    card.addEventListener('mousedown', (e) => {
      if (e.target.closest('button')) return;
      state.diagram.draggingTable = tName;
      state.diagram.dragStart.x = e.clientX;
      state.diagram.dragStart.y = e.clientY;
    });

    tablesContainer.appendChild(card);
  });

  setTimeout(() => redrawDiagramConnectors(), 20);
}

function redrawDiagramConnectors() {
  const svg = document.getElementById('diagramSvgLayer');
  if (!svg) return;
  svg.innerHTML = '';

  const tableKeys = Object.keys(state.tables);
  tableKeys.forEach(sourceTableName => {
    const sourceTable = state.tables[sourceTableName];
    (sourceTable.foreignKeys || []).forEach(fk => {
      const targetTableName = fk.refTable;
      const targetTable = state.tables[targetTableName];
      if (!targetTable) return;

      const sourceCard = document.getElementById(`diagram-card-${sourceTableName}`);
      const targetCard = document.getElementById(`diagram-card-${targetTableName}`);
      if (!sourceCard || !targetCard) return;

      const sX = parseInt(sourceCard.style.left, 10) + sourceCard.offsetWidth;
      const sY = parseInt(sourceCard.style.top, 10) + 40;
      const tX = parseInt(targetCard.style.left, 10);
      const tY = parseInt(targetCard.style.top, 10) + 40;

      const dx = Math.abs(tX - sX) * 0.5;
      const pathData = `M ${sX} ${sY} C ${sX + dx} ${sY}, ${tX - dx} ${tY}, ${tX} ${tY}`;

      const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
      path.setAttribute('d', pathData);
      path.setAttribute('stroke', '#38bdf8');
      path.setAttribute('stroke-width', '2');
      path.setAttribute('fill', 'none');
      path.setAttribute('stroke-dasharray', '4 2');
      svg.appendChild(path);
    });
  });
}

function autoArrangeDiagramTables() {
  const tableKeys = Object.keys(state.tables);
  tableKeys.forEach((tName, idx) => {
    state.diagram.tablePositions[tName] = {
      x: 60 + (idx % 3) * 280,
      y: 60 + Math.floor(idx / 3) * 260
    };
  });
  renderErDiagram();
}

function zoomDiagram(delta) {
  state.diagram.zoom = Math.max(0.4, Math.min(2.0, state.diagram.zoom + delta));
  const canvas = document.getElementById('diagramCanvas');
  if (canvas) canvas.style.transform = `scale(${state.diagram.zoom})`;
}

function resetDiagramZoom() {
  state.diagram.zoom = 1;
  const canvas = document.getElementById('diagramCanvas');
  if (canvas) canvas.style.transform = 'scale(1)';
}

function exportDiagramSvg() {
  const svg = document.getElementById('diagramSvgLayer');
  if (!svg) return;
  const data = new XMLSerializer().serializeToString(svg);
  const blob = new Blob([data], { type: 'image/svg+xml;charset=utf-8' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `Database_Diagram_${getCleanDbName()}.svg`;
  document.body.appendChild(a);
  a.click();
  a.remove();
}

/** =========================================================================
 * FOREIGN KEY RELATIONSHIP WIZARD
 * ========================================================================= */

function openAddRelationshipModal(presetSourceTable = null) {
  closeAllMenus();
  const sSelect = document.getElementById('fkSourceTable');
  const tSelect = document.getElementById('fkTargetTable');
  if (!sSelect || !tSelect) return;

  sSelect.innerHTML = '';
  tSelect.innerHTML = '';

  const tableKeys = Object.keys(state.tables);
  const sourceTable = presetSourceTable || state.activeTable || tableKeys[0];

  tableKeys.forEach(t => {
    sSelect.innerHTML += `<option value="${t}" ${t === sourceTable ? 'selected' : ''}>${t}</option>`;
    tSelect.innerHTML += `<option value="${t}" ${t !== sourceTable ? 'selected' : ''}>${t}</option>`;
  });

  updateFkSourceColumns();
  updateFkTargetColumns();
  openModal('addRelationshipModal');
}

function updateFkSourceColumns() {
  const sTable = document.getElementById('fkSourceTable').value;
  const colSelect = document.getElementById('fkSourceCol');
  if (!colSelect) return;
  colSelect.innerHTML = '';

  if (state.tables[sTable]) {
    state.tables[sTable].columns.forEach(col => {
      colSelect.innerHTML += `<option value="${col.name}">${col.name} (${col.type})</option>`;
    });
  }
}

function updateFkTargetColumns() {
  const tTable = document.getElementById('fkTargetTable').value;
  const colSelect = document.getElementById('fkTargetCol');
  if (!colSelect) return;
  colSelect.innerHTML = '';

  if (state.tables[tTable]) {
    state.tables[tTable].columns.forEach(col => {
      colSelect.innerHTML += `<option value="${col.name}" ${col.pk ? 'selected' : ''}>${col.name} (${col.type}) ${col.pk ? '★ PK' : ''}</option>`;
    });
  }
}

function handleAddRelationshipSubmit(e) {
  e.preventDefault();
  const sTable = document.getElementById('fkSourceTable').value;
  const sCol = document.getElementById('fkSourceCol').value;
  const tTable = document.getElementById('fkTargetTable').value;
  const tCol = document.getElementById('fkTargetCol').value;
  const onDelete = document.getElementById('fkOnDelete').value;

  if (sTable === tTable && sCol === tCol) {
    showToast('Cannot connect a column to itself.', 'danger');
    return;
  }

  if (!state.tables[sTable].foreignKeys) state.tables[sTable].foreignKeys = [];
  state.tables[sTable].foreignKeys = state.tables[sTable].foreignKeys.filter(fk => fk.column !== sCol);
  state.tables[sTable].foreignKeys.push({ column: sCol, refTable: tTable, refColumn: tCol, onDelete: onDelete });

  saveDatabaseToStorage();
  closeModal('addRelationshipModal');
  renderObjectExplorer();
  renderErDiagram();
  if (state.activeView === 'table-contents') renderTableContents();
  showToast(`Enforced Foreign Key: ${sTable}.${sCol} &rarr; ${tTable}.${tCol}`, 'success');
}

/** =========================================================================
 * NEW DATABASE & RESTORE / PROPERTIES
 * ========================================================================= */

function openInitialSetupModal() {
  closeAllMenus();
  openModal('initialSetupModal');
}

function handleInitialSetupSubmit(e) {
  e.preventDefault();
  const nameInput = document.getElementById('promptDbName');
  let name = nameInput.value.trim() || 'MyDatabase';
  if (!name.endsWith('.sqlite')) name += '.sqlite';

  state.dbName = name;
  const scaffold = document.getElementById('scaffoldSampleTables').checked;

  if (!scaffold) {
    state.tables = {};
    state.activeTable = null;
    state.diagram.tablePositions = {};
  } else {
    state.tables = {
      'Users': {
        columns: [
          { name: 'id', type: 'INTEGER', pk: true, nn: true, uq: true, def: '' },
          { name: 'name', type: 'TEXT', pk: false, nn: true, uq: false, def: '' },
          { name: 'email', type: 'TEXT', pk: false, nn: true, uq: true, def: '' },
          { name: 'role', type: 'TEXT', pk: false, nn: false, uq: false, def: 'User' },
          { name: 'created_at', type: 'TEXT', pk: false, nn: false, uq: false, def: '2026-01-01' }
        ],
        foreignKeys: [],
        rows: [
          { id: 1, name: 'John Doe', email: 'john@example.com', role: 'Admin', created_at: '2026-01-10' },
          { id: 2, name: 'Maria Garcia', email: 'maria@tech.org', role: 'Developer', created_at: '2026-01-15' }
        ]
      },
      'Orders': {
        columns: [
          { name: 'id', type: 'INTEGER', pk: true, nn: true, uq: true, def: '' },
          { name: 'user_id', type: 'INTEGER', pk: false, nn: true, uq: false, def: '' },
          { name: 'amount', type: 'REAL', pk: false, nn: true, uq: false, def: '0.00' },
          { name: 'status', type: 'TEXT', pk: false, nn: false, uq: false, def: 'Pending' }
        ],
        foreignKeys: [
          { column: 'user_id', refTable: 'Users', refColumn: 'id', onDelete: 'CASCADE' }
        ],
        rows: [
          { id: 101, user_id: 1, amount: 149.99, status: 'Completed' }
        ]
      }
    };
    state.activeTable = 'Users';
    state.diagram.tablePositions = {};
  }

  saveDatabaseToStorage();
  closeModal('initialSetupModal');
  updateConnectionUI();
  populateToolbarDbSelect();
  renderObjectExplorer();
  openNewQueryTab();
  showToast(`Database "${state.dbName}" initialized!`, 'success');
}

function openRenameDbModal() {
  closeAllMenus();
  const input = document.getElementById('renameDbInput');
  if (input) input.value = getCleanDbName();
  openModal('renameDbModal');
}

function handleRenameDbSubmit(e) {
  e.preventDefault();
  const input = document.getElementById('renameDbInput');
  let newName = input.value.trim() || 'MyDatabase';
  if (!newName.endsWith('.sqlite')) newName += '.sqlite';

  state.dbName = newName;
  saveDatabaseToStorage();
  closeModal('renameDbModal');
  updateConnectionUI();
  populateToolbarDbSelect();
  renderObjectExplorer();
  showToast(`Database renamed to "${newName}"`, 'success');
}

function openDbPropertiesModal() {
  closeAllMenus();
  const dbNameEl = document.getElementById('propDbNameVal');
  const serverEl = document.getElementById('propServerHostVal');
  const tableCountEl = document.getElementById('propTableCountVal');
  const rowCountEl = document.getElementById('propRowCountVal');
  const fkCountEl = document.getElementById('propFkCountVal');

  const tableKeys = Object.keys(state.tables);
  let totalRows = 0;
  let totalFks = 0;
  tableKeys.forEach(t => {
    totalRows += (state.tables[t].rows || []).length;
    totalFks += (state.tables[t].foreignKeys || []).length;
  });

  if (dbNameEl) dbNameEl.innerText = state.dbName;
  if (serverEl) serverEl.innerText = state.connection.serverName;
  if (tableCountEl) tableCountEl.innerText = tableKeys.length;
  if (rowCountEl) rowCountEl.innerText = totalRows;
  if (fkCountEl) fkCountEl.innerText = totalFks;

  openModal('dbPropertiesModal');
}

function openImportDbModal() {
  closeAllMenus();
  openModal('importDbModal');
}

function handleImportFileSelect(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (event) => {
    try {
      const data = JSON.parse(event.target.result);
      if (data.tables) {
        state.tables = data.tables;
        if (data.dbName) state.dbName = data.dbName;
        saveDatabaseToStorage();
        closeModal('importDbModal');
        updateConnectionUI();
        populateToolbarDbSelect();
        renderObjectExplorer();
        renderTabs();
        renderActiveView();
        showToast('Database backup successfully restored!', 'success');
      } else {
        showToast('Invalid database JSON backup file.', 'danger');
      }
    } catch (err) {
      showToast('Error reading database file.', 'danger');
    }
  };
  reader.readAsText(file);
}

function exportActiveTableCsv() {
  closeAllMenus();
  const tableName = state.activeTable;
  const tableData = state.tables[tableName];
  if (!tableData) return;

  const cols = tableData.columns.map(c => c.name);
  let csv = cols.join(',') + '\n';
  (tableData.rows || []).forEach(row => {
    const line = cols.map(c => `"${String(row[c] !== undefined ? row[c] : '').replace(/"/g, '""')}"`);
    csv += line.join(',') + '\n';
  });

  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = `dbo_${tableName}.csv`;
  document.body.appendChild(link);
  link.click();
  link.remove();
}

function printActiveTable() {
  window.print();
}

/** =========================================================================
 * UI INTERACTION & MENU HELPERS
 * ========================================================================= */

function setupEventListeners() {
  const themeBtn = document.getElementById('themeToggleBtn');
  if (themeBtn) themeBtn.addEventListener('click', toggleTheme);

  // Diagram card drag listener
  window.addEventListener('mousemove', (e) => {
    if (state.diagram.draggingTable) {
      const dx = (e.clientX - state.diagram.dragStart.x) / state.diagram.zoom;
      const dy = (e.clientY - state.diagram.dragStart.y) / state.diagram.zoom;
      const tName = state.diagram.draggingTable;
      const card = document.getElementById(`diagram-card-${tName}`);
      if (card && state.diagram.tablePositions[tName]) {
        state.diagram.tablePositions[tName].x += dx;
        state.diagram.tablePositions[tName].y += dy;
        card.style.left = `${state.diagram.tablePositions[tName].x}px`;
        card.style.top = `${state.diagram.tablePositions[tName].y}px`;
        state.diagram.dragStart.x = e.clientX;
        state.diagram.dragStart.y = e.clientY;
        redrawDiagramConnectors();
      }
    }
  });

  window.addEventListener('mouseup', () => {
    if (state.diagram.draggingTable) {
      state.diagram.draggingTable = null;
      saveDatabaseToStorage();
    }
  });

  // Global click closes dropdowns and context menus
  window.addEventListener('click', (e) => {
    hideSSMSContextMenu();
    if (!e.target.closest('.menu-item')) {
      closeAllMenus();
    }
  });
}

function setupKeyboardShortcuts() {
  window.addEventListener('keydown', (e) => {
    if (e.key === 'F5' || (e.ctrlKey && e.key === 'Enter')) {
      e.preventDefault();
      runConsoleSqlQuery();
    }
    if (e.ctrlKey && (e.key === 'n' || e.key === 'N')) {
      e.preventDefault();
      openNewQueryTab();
    }
    if (e.ctrlKey && (e.key === 's' || e.key === 'S')) {
      e.preventDefault();
      if (state.activeView === 'table-designer') {
        handleSaveTableDesigner();
      } else {
        handleSaveDbButtonClick();
      }
    }
    if (e.key === 'Escape') {
      closeAllMenus();
      hideSSMSContextMenu();
      document.querySelectorAll('.modal-backdrop:not(.hidden)').forEach(m => m.classList.add('hidden'));
    }
  });
}

function toggleTheme() {
  const body = document.body;
  const icon = document.querySelector('#themeToggleBtn i');
  if (body.classList.contains('dark-theme')) {
    body.classList.remove('dark-theme');
    body.classList.add('light-theme');
    if (icon) icon.className = 'fa-solid fa-sun';
    showToast('Switched to Classic SSMS Light Theme', 'info');
  } else {
    body.classList.remove('light-theme');
    body.classList.add('dark-theme');
    if (icon) icon.className = 'fa-solid fa-moon';
    showToast('Switched to Visual Studio Dark Theme', 'info');
  }
}

function toggleMenuDropdown(e, menuId) {
  if (e && e.stopPropagation) e.stopPropagation();
  const menu = document.getElementById(menuId);
  const isAlreadyOpen = menu ? menu.classList.contains('show') : false;
  closeAllMenus();
  if (menu && !isAlreadyOpen) menu.classList.add('show');
}

function closeAllMenus() {
  document.querySelectorAll('.ssms-dropdown').forEach(d => d.classList.remove('show'));
}

function openModal(modalId) {
  const modal = document.getElementById(modalId);
  if (modal) modal.classList.remove('hidden');
}

function closeModal(modalId) {
  const modal = document.getElementById(modalId);
  if (modal) modal.classList.add('hidden');
}

function openShortcutsModal() {
  closeAllMenus();
  openModal('shortcutsModal');
}

function escapeHtml(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

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
    toast.style.transform = 'translateX(20px)';
    setTimeout(() => toast.remove(), 250);
  }, 3200);
}
