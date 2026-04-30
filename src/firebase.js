// src/firebase.js
import { initializeApp } from 'firebase/app';
import {
  getFirestore, doc, setDoc, getDoc, deleteDoc,
  collection, getDocs
} from 'firebase/firestore';

const firebaseConfig = {
  apiKey: "AIzaSyB8HDXTgwSf_kNkSmjzIYiqfV2_9ZxUwoM",
  authDomain: "taller-doe.firebaseapp.com",
  projectId: "taller-doe",
  storageBucket: "taller-doe.firebasestorage.app",
  messagingSenderId: "482880338231",
  appId: "1:482880338231:web:137d73f59db39721d342b2",
  measurementId: "G-G4NHGR9WTY"
};

const app = initializeApp(firebaseConfig);
const db = getFirestore(app);

// Adaptador con la misma interfaz que window.storage
// para que el resto del codigo no cambie
export const storage = {
  async set(key, value, shared = true) {
    const cedula = key.replace(/^doe_taller_v5:/, '');
    await setDoc(doc(db, 'submissions', cedula), {
      payload: value,
      updatedAt: new Date().toISOString()
    });
    return { key, value, shared };
  },

  async get(key, shared = true) {
    const cedula = key.replace(/^doe_taller_v5:/, '');
    const snap = await getDoc(doc(db, 'submissions', cedula));
    if (!snap.exists()) throw new Error('Not found');
    return { key, value: snap.data().payload, shared };
  },

  async delete(key, shared = true) {
    const cedula = key.replace(/^doe_taller_v5:/, '');
    await deleteDoc(doc(db, 'submissions', cedula));
    return { key, deleted: true, shared };
  },

  async list(prefix = '', shared = true) {
    const snap = await getDocs(collection(db, 'submissions'));
    const keys = snap.docs.map(d => `doe_taller_v5:${d.id}`);
    return { keys, prefix, shared };
  }
};

// Lo conectamos a window.storage para que el codigo existente funcione
if (typeof window !== 'undefined') {
  window.storage = storage;
}
