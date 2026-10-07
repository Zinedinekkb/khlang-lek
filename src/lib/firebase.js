import firebase from "firebase/compat/app";
import "firebase/compat/firestore";
import "firebase/compat/auth";
import "firebase/compat/storage";

const firebaseConfig = {
  apiKey: import.meta.env.VITE_FIREBASE_API_KEY || "AIzaSyBblsal6kMJfzi6oNQMeBTwJ4FfvZpRNac",
  authDomain: import.meta.env.VITE_FIREBASE_AUTH_DOMAIN || "jin-replit1.firebaseapp.com",
  projectId: import.meta.env.VITE_FIREBASE_PROJECT_ID || "jin-replit1",
  storageBucket: import.meta.env.VITE_FIREBASE_STORAGE_BUCKET || "jin-replit1.firebasestorage.app",
  messagingSenderId: import.meta.env.VITE_FIREBASE_MESSAGING_SENDER_ID || "790901511056",
  appId: import.meta.env.VITE_FIREBASE_APP_ID || "1:790901511056:web:4fc780107bba2e90cc1915"
};

export const isFirebaseConfigured = Boolean(
  firebaseConfig.apiKey &&
  firebaseConfig.projectId &&
  !firebaseConfig.apiKey.includes("your_api_key")
);

let app = null;
let db = null;
let auth = null;
let storage = null;

if (isFirebaseConfigured) {
  try {
    app = firebase.apps.length ? firebase.app() : firebase.initializeApp(firebaseConfig);
    db = firebase.firestore();
    auth = firebase.auth();
    storage = firebase.storage();

    // Enable offline persistence
    db.enablePersistence({ synchronizeTabs: true }).catch((err) => {
      if (err.code === "failed-precondition") {
        console.warn("Firestore persistence failed: Multiple tabs open");
      } else if (err.code === "unimplemented") {
        console.warn("Firestore persistence is not supported in this browser");
      } else {
        console.warn("Firestore persistence error:", err);
      }
    });
  } catch (err) {
    console.error("Firebase initialization failed:", err);
    app = null;
    db = null;
    auth = null;
    storage = null;
  }
}

export function getDb() {
  return db;
}

export function getAuth() {
  return auth;
}

export function getStorage() {
  return storage;
}

export function getFirebaseApp() {
  return app;
}

export { firebase };
