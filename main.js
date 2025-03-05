const { app, BrowserWindow, dialog, ipcMain } = require('electron');
const fs = require('fs');
const path = require('path');
const iconv = require('iconv-lite'); // Importing iconv-lite for encoding support

let win;

const createWindow = () => {
  win = new BrowserWindow({
    width: 600,
    height: 400,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
      enableRemoteModule: true,
    },
  });

  win.loadFile('index.html');
};

app.whenReady().then(() => {
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

ipcMain.on('open-file-dialog', (event) => {
  dialog.showOpenDialog(win, {
    properties: ['openFile'],
    // filters: [{ name: 'Text Files', extensions: ['txt'] }],
    filters: [{ name: 'Text & Log Files', extensions: ['txt', 'log'] }], // Supports both .txt and .log
  }).then((file) => {
    if (!file.canceled && file.filePaths.length > 0) {
      const inputFilePath = file.filePaths[0];
      const inputDir = path.dirname(inputFilePath);
      // const baseName = path.basename(inputFilePath, '.txt');
      const baseName = path.basename(inputFilePath, path.extname(inputFilePath)); // Supports both .txt and .log

      // const content = fs.readFileSync(inputFilePath, 'utf-8');
      // Read the file content with Shift_JIS encoding
      const content = iconv.decode(fs.readFileSync(inputFilePath), 'Shift_JIS');
      event.reply('file-path', inputFilePath);
      event.reply('file-content', content);

      // Automatically create only the Excel file
      const fileToCreate = { suffix: '_hyojun.xlsx', eventName: 'file-created-4' };
      const newFilePath = path.join(inputDir, `${baseName}${fileToCreate.suffix}`);
      fs.writeFileSync(newFilePath, '', 'utf-8'); // Creates an empty Excel file
      event.reply(fileToCreate.eventName, newFilePath);
    }
  });
});

ipcMain.on('open-save-dialog4', (event) => {
  dialog.showSaveDialog(win, {
    defaultPath: 'new-file.xlsx',
    filters: [{ name: 'Excel Files', extensions: ['xlsx'] }],
  }).then((file) => {
    if (!file.canceled) {
      fs.writeFileSync(file.filePath, '', 'utf-8');
      event.reply('file-created-4', file.filePath);
    }
  });
});
