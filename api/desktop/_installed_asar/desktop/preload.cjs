const { contextBridge } = require("electron");

contextBridge.exposeInMainWorld("ironlogDesktop", {
  isDesktop: true,
});

