import { getDb } from "./firebase.js";

window.claude = {
  async use(name) {
    if (name === "db") {
      const db = getDb();
      if (!db) {
        throw new Error("Firebase is not configured yet. Falling back to localStore.");
      }
      return db;
    }
    if (name === "user") {
      if (window._userShim) return window._userShim;
      throw new Error("user capability not configured yet");
    }
    if (name === "assets") {
      if (window._assetsShim) return window._assetsShim;
      throw new Error("assets capability not configured yet");
    }
    if (name === "downloads") {
      if (window._downloadsShim) return window._downloadsShim;
      throw new Error("downloads capability not configured yet");
    }
    if (name === "sample") {
      if (window._sampleShim) return window._sampleShim;
      throw new Error("sample capability not configured yet");
    }
    throw new Error("unknown capability: " + name);
  }
};
