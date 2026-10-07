import { getAuth, getDb } from "./firebase.js";

let authState = {
  user: null, // Firebase Auth user
  profile: null, // Firestore profile { uid, email, displayName, role: 'owner'|'staff', status: 'approved'|'pending'|'disabled' }
  loading: true
};

const listeners = new Set();
let profileUnsub = null;

export function getAuthState() {
  return authState;
}

export function subscribeAuth(fn) {
  listeners.add(fn);
  fn(authState);
  return () => listeners.delete(fn);
}

function notify() {
  for (const fn of listeners) {
    try {
      fn(authState);
    } catch (e) {
      console.error("Auth listener error:", e);
    }
  }
}

/**
 * Initializes Firebase Auth state listener and Firestore profile sync.
 */
export function initAuth(onReady) {
  const auth = getAuth();
  if (!auth) {
    console.warn("Firebase Auth not available, running in fallback mode");
    authState.loading = false;
    notify();
    if (onReady) onReady(authState);
    return;
  }

  auth.onAuthStateChanged(async (user) => {
    if (profileUnsub) {
      profileUnsub();
      profileUnsub = null;
    }

    if (!user) {
      authState.user = null;
      authState.profile = null;
      authState.loading = false;
      notify();
      if (onReady) onReady(authState);
      return;
    }

    authState.user = user;
    const db = getDb();
    if (!db) {
      // Fallback if no db
      authState.profile = {
        uid: user.uid,
        email: user.email,
        displayName: user.displayName || user.email.split("@")[0],
        role: "owner",
        status: "approved"
      };
      authState.loading = false;
      notify();
      if (onReady) onReady(authState);
      return;
    }

    const userRef = db.collection("users").doc(user.uid);

    // Subscribe to realtime profile updates so approvals unlock instantly without reload
    profileUnsub = userRef.onSnapshot(async (doc) => {
      if (doc.exists) {
        authState.profile = { uid: user.uid, ...doc.data() };
      } else {
        // User profile doesn't exist yet -> check if this is the first user (Dev account / Master Owner)
        let isFirstUser = false;
        try {
          const snapshot = await db.collection("users").where("status", "==", "approved").where("role", "==", "owner").limit(1).get();
          isFirstUser = snapshot.empty;
        } catch (err) {
          console.warn("Could not check existing owners:", err);
        }

        // Dev account or first user gets owner & approved immediately
        const isDev = isFirstUser || user.email.toLowerCase().includes("dev") || user.email.toLowerCase().includes("zined");
        const newProfile = {
          uid: user.uid,
          email: user.email,
          displayName: user.displayName || user.email.split("@")[0],
          role: isDev ? "owner" : "staff",
          status: isDev ? "approved" : "pending",
          createdAt: Date.now()
        };

        try {
          await userRef.set(newProfile);
          authState.profile = newProfile;
        } catch (e) {
          console.error("Failed to create user profile:", e);
          authState.profile = newProfile;
        }
      }

      authState.loading = false;
      notify();
      if (onReady) {
        onReady(authState);
        onReady = null;
      }
    }, (err) => {
      console.error("Profile onSnapshot error:", err);
      authState.loading = false;
      notify();
      if (onReady) {
        onReady(authState);
        onReady = null;
      }
    });
  });
}

/**
 * Register a new user with email and password
 */
export async function registerUser({ email, password, displayName }) {
  const auth = getAuth();
  if (!auth) throw new Error("ระบบล็อกอินไม่พร้อมใช้งาน");

  const cleanEmail = String(email).trim().toLowerCase();
  const cleanName = String(displayName).trim() || cleanEmail.split("@")[0];

  const userCredential = await auth.createUserWithEmailAndPassword(cleanEmail, password);
  const user = userCredential.user;

  // Update Auth profile display name
  try {
    await user.updateProfile({ displayName: cleanName });
  } catch (e) {}

  // Create Firestore record
  const db = getDb();
  if (db) {
    let isFirstUser = false;
    try {
      const snapshot = await db.collection("users").where("status", "==", "approved").where("role", "==", "owner").limit(1).get();
      isFirstUser = snapshot.empty;
    } catch (e) {}

    const isDev = isFirstUser || cleanEmail.includes("dev") || cleanEmail.includes("zined");
    const profile = {
      uid: user.uid,
      email: cleanEmail,
      displayName: cleanName,
      role: isDev ? "owner" : "staff",
      status: isDev ? "approved" : "pending",
      createdAt: Date.now()
    };
    await db.collection("users").doc(user.uid).set(profile);
  }

  return user;
}

/**
 * Sign in with email and password
 */
export async function loginUser({ email, password }) {
  const auth = getAuth();
  if (!auth) throw new Error("ระบบล็อกอินไม่พร้อมใช้งาน");
  const cleanEmail = String(email).trim().toLowerCase();
  return auth.signInWithEmailAndPassword(cleanEmail, password);
}

/**
 * Sign out
 */
export async function logoutUser() {
  const auth = getAuth();
  if (!auth) return;
  return auth.signOut();
}

/**
 * Send password reset email
 */
export async function resetPassword(email) {
  const auth = getAuth();
  if (!auth) throw new Error("ระบบล็อกอินไม่พร้อมใช้งาน");
  return auth.sendPasswordResetEmail(String(email).trim().toLowerCase());
}

/**
 * Approve a pending user (Owner only)
 */
export async function approveUser(uid, role = "staff") {
  const db = getDb();
  if (!db) throw new Error("ฐานข้อมูลไม่พร้อมใช้งาน");
  return db.collection("users").doc(uid).update({
    status: "approved",
    role: role,
    approvedAt: Date.now(),
    approvedBy: authState.user?.uid || "dev"
  });
}

/**
 * Disable a user (Owner only)
 */
export async function disableUser(uid) {
  const db = getDb();
  if (!db) throw new Error("ฐานข้อมูลไม่พร้อมใช้งาน");
  return db.collection("users").doc(uid).update({
    status: "disabled",
    updatedAt: Date.now()
  });
}

/**
 * Update user role (Owner only)
 */
export async function updateUserRole(uid, role) {
  const db = getDb();
  if (!db) throw new Error("ฐานข้อมูลไม่พร้อมใช้งาน");
  return db.collection("users").doc(uid).update({
    role,
    status: "approved",
    updatedAt: Date.now()
  });
}

/**
 * Delete a pending or rejected user record
 */
export async function deleteUserRecord(uid) {
  const db = getDb();
  if (!db) throw new Error("ฐานข้อมูลไม่พร้อมใช้งาน");
  return db.collection("users").doc(uid).delete();
}

/**
 * Subscribe to all users in real-time (for Owner management)
 */
export function subscribeAllUsers(cb) {
  const db = getDb();
  if (!db) return () => {};
  return db.collection("users").orderBy("createdAt", "desc").onSnapshot((snapshot) => {
    const list = snapshot.docs.map(d => ({ uid: d.id, ...d.data() }));
    cb(list);
  }, (err) => {
    console.error("subscribeAllUsers error:", err);
  });
}
