const { app, BrowserWindow } = require('electron');
const path = require('path');
const { spawn } = require('child_process');

let mainWindow;
let backendProcess;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1400,
    height: 900,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false
    }
  });

  mainWindow.loadFile('frontend/index.html');
  mainWindow.setMenuBarVisibility(false); // Üstteki File/Edit menüsünü gizler, daha şık durur

  // Uygulamanın paketlenmiş (.exe) olup olmadığını kontrol et
  if (app.isPackaged) {
    // Paketlenmiş sürüm: PyInstaller ile oluşturduğumuz app.exe'yi çalıştır
    const backendPath = path.join(process.resourcesPath, 'app.exe');
    backendProcess = spawn(backendPath);
  } else {
    // Geliştirme sürümü: Normal Python scriptini çalıştır
    const pythonCmd = process.platform === 'win32' ? 'python' : 'python3';
    backendProcess = spawn(pythonCmd, [path.join(__dirname, 'backend', 'app.py')]);
  }

  backendProcess.stdout.on('data', (data) => console.log(`Backend: ${data}`));
  backendProcess.stderr.on('data', (data) => console.error(`Backend Error: ${data}`));
}

app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
  if (backendProcess) backendProcess.kill();
  app.quit();
});