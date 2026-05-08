const { app, BrowserWindow } = require("electron");

function createWindow() {

  const win = new BrowserWindow({
    width: 1600,
    height: 1000,
  });

  win.loadURL("http://localhost:5173");
}

app.whenReady().then(createWindow);