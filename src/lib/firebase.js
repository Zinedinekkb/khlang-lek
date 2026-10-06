import firebase from "firebase/compat/app";
import "firebase/compat/firestore";

const firebaseConfig = {
  apiKey: import.meta.env.VITE_FIREBASE_API_KEY,
  authDomain: import.meta.env.VITE_FIREBASE_AUTH_DOMAIN,
  projectId: import.meta.env.VITE_FIREBASE_PROJECT_ID,
  storageBucket: import.meta.env.VITE_FIREBASE_STORAGE_BUCKET,
  messagingSenderId: import.meta.env.VITE_FIREBASE_MESSAGING_SENDER_ID,
  appId: import.meta.env.VITE_FIREBASE_APP_ID
};

export const isFirebaseConfigured = Boolean(
  firebaseConfig.apiKey &&
  firebaseConfig.projectId &&
  !firebaseConfig.apiKey.includes("your_api_key")
);

let app = null;
let db = null;

if (isFirebaseConfigured) {
  try {
    app = firebase.apps.length ? firebase.app() : firebase.initializeApp(firebaseConfig);
    db = firebase.firestore();

    // Enable offline persistence if supported
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
  }
}

export function getDb() {
  return db;
}

export function getFirebaseApp() {
  return app;
}

export { firebase };
