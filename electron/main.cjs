'use strict';

const { app, BrowserWindow, ipcMain, shell, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');

// ─── Paths ──────────────────────────────────────────────────────────────────
const DB_FILE = path.join(app.getPath('userData'), 'db.json');
const FILES_DIR = path.join(os.homedir(), 'Documents', 'SIPA_FTK_Files');

// ─── DB Helpers ─────────────────────────────────────────────────────────────

function ensureFilesDir() {
  if (!fs.existsSync(FILES_DIR)) {
    fs.mkdirSync(FILES_DIR, { recursive: true });
  }
}

function readDb() {
  if (!fs.existsSync(DB_FILE)) {
    const initial = { surat_masuk: [], surat_keluar: [] };
    fs.writeFileSync(DB_FILE, JSON.stringify(initial, null, 2), 'utf8');
    return initial;
  }
  try {
    return JSON.parse(fs.readFileSync(DB_FILE, 'utf8'));
  } catch {
    return { surat_masuk: [], surat_keluar: [] };
  }
}

function writeDb(data) {
  fs.writeFileSync(DB_FILE, JSON.stringify(data, null, 2), 'utf8');
}

function generateId() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 9);
}

function generateRandom16CharFilename(originalName) {
  const chars = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  let result = '';
  for (let i = 0; i < 16; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  const dotIndex = originalName.lastIndexOf('.');
  const ext = dotIndex !== -1 ? originalName.substring(dotIndex) : '';
  return result + ext;
}

// ─── IPC Handlers ───────────────────────────────────────────────────────────

// Read all letters
ipcMain.handle('db:getLetters', () => {
  return readDb();
});

// Add a letter entry to a table ('surat_masuk' or 'surat_keluar')
ipcMain.handle('db:addLetter', (_event, table, entry) => {
  const db = readDb();
  const newEntry = {
    id: generateId(),
    created_at: new Date().toISOString(),
    ...entry,
  };
  db[table] = [newEntry, ...(db[table] || [])];
  writeDb(db);
  return newEntry;
});

// Delete entries by id array from a table
ipcMain.handle('db:deleteLetters', (_event, table, ids) => {
  const db = readDb();
  const idSet = new Set(ids);
  db[table] = (db[table] || []).filter(item => !idSet.has(item.id));
  writeDb(db);
  return { deleted: ids.length };
});

// Copy an uploaded file to local files directory
ipcMain.handle('file:upload', (_event, srcPath, originalName) => {
  ensureFilesDir();
  const randomName = generateRandom16CharFilename(originalName);
  const destPath = path.join(FILES_DIR, randomName);
  fs.copyFileSync(srcPath, destPath);
  return { fileName: randomName, filePath: destPath };
});

// Open a file in the system's default viewer
ipcMain.handle('file:open', (_event, filePath) => {
  if (!filePath || !fs.existsSync(filePath)) {
    return { error: 'File not found: ' + filePath };
  }
  shell.openPath(filePath);
  return { ok: true };
});

// Return the db path and files directory for settings display
ipcMain.handle('app:getPaths', () => {
  return { dbPath: DB_FILE, filesDir: FILES_DIR };
});

// Open a native file-open dialog (so user picks a file to upload)
ipcMain.handle('dialog:openFile', async () => {
  const result = await dialog.showOpenDialog({
    properties: ['openFile'],
    filters: [
      { name: 'Documents & Images', extensions: ['pdf', 'png', 'jpg', 'jpeg', 'doc', 'docx', 'xls', 'xlsx'] },
      { name: 'All Files', extensions: ['*'] }
    ]
  });
  if (result.canceled || result.filePaths.length === 0) return null;
  return result.filePaths[0];
});

// ─── Window ─────────────────────────────────────────────────────────────────

let win;

function createWindow() {
  win = new BrowserWindow({
    width: 1280,
    height: 800,
    minWidth: 900,
    minHeight: 600,
    title: 'SIPA FTK UnHar',
    show: false,
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  win.once('ready-to-show', () => win.show());

  const isDev = process.env.NODE_ENV === 'development' || process.argv.includes('--dev');

  if (isDev) {
    // Poll until Vite dev server is ready, then load it
    const DEV_URL = 'http://localhost:5173';
    const tryLoad = () => {
      const http = require('http');
      const req = http.get(DEV_URL, () => {
        win.loadURL(DEV_URL);
        win.webContents.openDevTools();
      });
      req.on('error', () => setTimeout(tryLoad, 500));
      req.end();
    };
    tryLoad();
  } else {
    // In production, load built assets
    win.loadFile(path.join(__dirname, '..', 'dist', 'index.html'));
  }
}

app.whenReady().then(() => {
  // remove global menu bar
  const { Menu } = require('electron');
  Menu.setApplicationMenu(null);

  // Ensure db and files dir exist at startup
  readDb();
  ensureFilesDir();
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});
