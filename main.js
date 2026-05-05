const { app, BrowserWindow, Menu } = require('electron');
const path = require('path');

function createWindow() {
  // 创建浏览器窗口
  const mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    minWidth: 800,
    minHeight: 600,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
      enableRemoteModule: true
    },
    icon: path.join(__dirname, 'icon.ico'),
    title: '防火墙配置翻译工具'
  });

  // 加载 index.html
  mainWindow.loadFile('index.html');

  // 隐藏菜单栏
  Menu.setApplicationMenu(null);

  // 打开开发者工具（开发模式）
  // mainWindow.webContents.openDevTools();
}

// Electron 准备就绪后创建窗口
app.whenReady().then(() => {
  createWindow();

  app.on('activate', function () {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

// 关闭所有窗口时退出应用（macOS除外）
app.on('window-all-closed', function () {
  if (process.platform !== 'darwin') app.quit();
});
