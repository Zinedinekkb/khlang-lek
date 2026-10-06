import { getDb, getAuth, getStorage } from "./firebase.js";

// In-memory or localStorage cache for local assets fallback
const localAssetCache = new Map();

export function assetUrl(id) {
  if (!id) return "";
  if (id.startsWith("http://") || id.startsWith("https://") || id.startsWith("data:") || id.startsWith("blob:")) {
    return id;
  }
  if (localAssetCache.has(id)) {
    return localAssetCache.get(id);
  }
  return id;
}

window.assetUrl = assetUrl;

const downloadsShim = {
  async save({ filename, data }) {
    try {
      let blob;
      if (typeof data === "string") {
        blob = new Blob([data], { type: "text/html;charset=utf-8" });
      } else if (data instanceof ArrayBuffer || data instanceof Uint8Array) {
        blob = new Blob([data], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      } else {
        blob = new Blob([data]);
      }
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = filename || "download";
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 10000);
      return true;
    } catch (err) {
      console.error("Download failed:", err);
      throw { code: "declined", message: "Download failed" };
    }
  }
};

const userShim = {
  isOwner: () => true, // สิทธิ์เจ้าของสำหรับทั้งสองคน
  id: () => {
    const auth = getAuth();
    return auth?.currentUser?.uid || "owner-1";
  },
  profiles: async (ids) => {
    const auth = getAuth();
    const currentName = auth?.currentUser?.displayName || "เจ้าของร้าน";
    const result = {};
    if (Array.isArray(ids)) {
      ids.forEach(id => {
        result[id] = { name: currentName };
      });
    }
    return result;
  }
};

const assetsShim = {
  async upload(blob, { type = "image/jpeg" } = {}) {
    const storage = getStorage();
    const randomId = Date.now().toString(36) + Math.random().toString(36).slice(2, 7);

    if (storage) {
      try {
        const fileRef = storage.ref(`receipts/${randomId}.jpg`);
        const snapshot = await fileRef.put(blob, { contentType: type });
        const downloadUrl = await snapshot.ref.getDownloadURL();
        localAssetCache.set(randomId, downloadUrl);
        return { id: downloadUrl };
      } catch (err) {
        console.warn("Storage upload failed, using local blob fallback:", err);
      }
    }

    // Local fallback: convert blob to data URL
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onloadend = () => {
        const dataUrl = reader.result;
        localAssetCache.set(randomId, dataUrl);
        resolve({ id: dataUrl });
      };
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    });
  }
};

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
      return userShim;
    }
    if (name === "downloads") {
      return downloadsShim;
    }
    if (name === "assets") {
      return assetsShim;
    }
    if (name === "sample") {
      if (window._sampleShim) return window._sampleShim;
      throw new Error("sample capability not configured yet");
    }
    throw new Error("unknown capability: " + name);
  }
};
