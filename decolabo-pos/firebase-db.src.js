// Firebase adapter for Decolabo POS. Build: npx esbuild firebase-db.src.js --bundle --format=esm --minify --target=es2020 --outfile=firebase-db.bundle.js (firebase@13.0.0).
// Exposes the same small db interface the page uses with
// claude.ai's shared store: db.doc(path).onSnapshot/set and db.collection(name).orderBy().limit().onSnapshot/doc(id).
import { initializeApp } from "firebase/app";
import { getAuth, onAuthStateChanged, signInAnonymously } from "firebase/auth";
import {
  initializeFirestore, persistentLocalCache, persistentMultipleTabManager,
  doc, collection, query, orderBy, limit, onSnapshot, setDoc, deleteDoc,
} from "firebase/firestore";

export async function createDb(cfg, onWriteError) {
  const app = initializeApp(cfg);
  const auth = getAuth(app);
  // Anonymous sign-in: no login screen of Firebase's own; the POS keeps its ID/password screen.
  await new Promise((resolve, reject) => {
    const off = onAuthStateChanged(auth, u => { if (u) { off(); resolve(); } }, reject);
    signInAnonymously(auth).catch(reject);
  });
  // Cache on the device so the register keeps working offline and sends writes when back online.
  const db = initializeFirestore(app, {
    ignoreUndefinedProperties: true,
    localCache: persistentLocalCache({ tabManager: persistentMultipleTabManager() }),
  });
  const code = e => Object.assign(new Error(e && e.message), { code: mapCode(e && e.code) });
  const wrapDoc = ref => ({
    onSnapshot: (ok, ng) => onSnapshot(ref, s => ok({ exists: s.exists(), data: () => s.data() }), e => ng && ng(code(e))),
    // Writes land in the local cache at once; the server copy follows when online. Waiting for the
    // server would freeze checkout while offline, so failures are reported through onWriteError instead.
    set: data => { setDoc(ref, data).catch(e => onWriteError && onWriteError(code(e))); return Promise.resolve(); },
    delete: () => { deleteDoc(ref).catch(e => onWriteError && onWriteError(code(e))); return Promise.resolve(); },
  });
  const wrapQuery = (col, cons) => ({
    orderBy: (f, dir) => wrapQuery(col, cons.concat(orderBy(f, dir))),
    limit: n => wrapQuery(col, cons.concat(limit(n))),
    onSnapshot: (ok, ng) => onSnapshot(query(col, ...cons),
      s => ok({ docs: s.docs.map(d => ({ id: d.id, data: () => d.data() })) }), e => ng && ng(code(e))),
    doc: id => wrapDoc(doc(col, id)),
  });
  return {
    doc: path => wrapDoc(doc(db, path)),
    collection: name => wrapQuery(collection(db, name), []),
  };
}
function mapCode(c) {
  if (c === "permission-denied" || c === "unauthenticated") return "invalid_argument";
  if (c === "resource-exhausted") return "quota_exceeded";
  return "unavailable";
}
