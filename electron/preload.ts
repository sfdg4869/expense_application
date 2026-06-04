import { contextBridge, ipcRenderer } from "electron";

contextBridge.exposeInMainWorld("expenseApi", {
  pickFile: (filters: { name: string; extensions: string[] }[]) =>
    ipcRenderer.invoke("pick-file", filters),
  pickSave: (defaultName: string) => ipcRenderer.invoke("pick-save", defaultName),
  pickFolder: () => ipcRenderer.invoke("pick-folder"),
  loadImagesFromFolder: (folderPath: string) =>
    ipcRenderer.invoke("load-images-from-folder", folderPath),
  pickImages: () => ipcRenderer.invoke("pick-images"),
  parseStatement: (filePath: string) => ipcRenderer.invoke("parse-statement", filePath),
  loadDefaults: () => ipcRenderer.invoke("load-defaults"),
  saveDefaults: (defaults: unknown) => ipcRenderer.invoke("save-defaults", defaults),
  generate: (payload: unknown) => ipcRenderer.invoke("generate", payload),
});