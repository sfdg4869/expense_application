import { app, BrowserWindow, ipcMain, dialog } from "electron";
import path from "path";
import fs from "fs";
import { spawn } from "child_process";
import { parseCardStatement } from "./parser";

const isDev = !app.isPackaged;

function getScriptPath(): string {
  if (isDev) {
    return path.join(app.getAppPath(), "scripts", "fill-expense-form.ps1");
  }
  return path.join(process.resourcesPath, "scripts", "fill-expense-form.ps1");
}

function getDefaultsPath(): string {
  return path.join(app.getPath("userData"), "defaults.json");
}

function createWindow() {
  const win = new BrowserWindow({
    width: 1100,
    height: 800,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  if (isDev) {
    win.loadURL("http://localhost:5175");
    win.webContents.openDevTools({ mode: "detach" });
  } else {
    win.loadFile(path.join(__dirname, "../dist/index.html"));
  }
}

app.whenReady().then(() => {
  createWindow();
  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});

ipcMain.handle("pick-file", async (_e, filters) => {
  const r = await dialog.showOpenDialog({
    properties: ["openFile"],
    filters: filters ?? [{ name: "Excel", extensions: ["xls", "xlsx"] }],
  });
  return r.canceled || !r.filePaths[0] ? null : r.filePaths[0];
});

ipcMain.handle("pick-folder", async () => {
  const r = await dialog.showOpenDialog({ properties: ["openDirectory"] });
  return r.canceled || !r.filePaths[0] ? null : r.filePaths[0];
});

ipcMain.handle("pick-images", async () => {
  const r = await dialog.showOpenDialog({
    properties: ["openFile", "multiSelections"],
    filters: [{ name: "Images", extensions: ["jpg", "jpeg", "png", "webp", "heic"] }],
  });
  return r.canceled ? [] : r.filePaths;
});

ipcMain.handle("parse-statement", async (_e, filePath: string) => {
  return parseCardStatement(filePath);
});

ipcMain.handle("load-defaults", async () => {
  const p = getDefaultsPath();
  if (!fs.existsSync(p)) {
    return {
      account: "복리후생비",
      client: "엑쥐",
      userName: "정경수",
      detail: "야근식대",
    };
  }
  return JSON.parse(fs.readFileSync(p, "utf-8"));
});

ipcMain.handle("save-defaults", async (_e, defaults) => {
  fs.writeFileSync(getDefaultsPath(), JSON.stringify(defaults, null, 2), "utf-8");
});



ipcMain.handle("load-images-from-folder", async (_e, folderPath: string) => {
  const exts = new Set([".jpg", ".jpeg", ".png", ".webp", ".heic"]);
  const entries = fs.readdirSync(folderPath, { withFileTypes: true });
  return entries
    .filter((e) => e.isFile() && exts.has(path.extname(e.name).toLowerCase()))
    .map((e) => path.join(folderPath, e.name))
    .sort((a, b) => a.localeCompare(b, "ko"));
});
ipcMain.handle("pick-save", async (_e, defaultName: string) => {
  const r = await dialog.showSaveDialog({
    defaultPath: defaultName,
    filters: [{ name: "Excel 97-2003", extensions: ["xls"] }],
  });
  return r.canceled || !r.filePath ? null : r.filePath;
});
ipcMain.handle("generate", async (_e, payload) => {
  const scriptPath = getScriptPath();
  if (!fs.existsSync(scriptPath)) {
    return { success: false, error: `?ㅽ겕由쏀듃瑜?李얠쓣 ???놁뒿?덈떎: ${scriptPath}` };
  }

  const tmpDir = path.join(app.getPath("temp"), "expense-app");
  fs.mkdirSync(tmpDir, { recursive: true });
  const configPath = path.join(tmpDir, `job-${Date.now()}.json`);
  fs.writeFileSync(configPath, JSON.stringify(payload, null, 2), "utf-8");

  return new Promise((resolve) => {
    const ps = spawn(
      "powershell.exe",
      [
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        scriptPath,
        "-ConfigPath",
        configPath,
      ],
      { windowsHide: true, env: { ...process.env, PYTHONUTF8: "1" } }
    );

    let stdout = "";
    let stderr = "";
    ps.stdout.on("data", (d) => (stdout += d.toString("utf8")));
    ps.stderr.on("data", (d) => (stderr += d.toString("utf8")));

    ps.on("close", (code) => {
      try {
        fs.unlinkSync(configPath);
      } catch {
        /* ignore */
      }
      if (code === 0) {
        resolve({ success: true, outputPath: payload.outputPath });
      } else {
        resolve({
          success: false,
          error: stderr || stdout || `PowerShell 醫낅즺 肄붾뱶 ${code}`,
        });
      }
    });
  });
});

