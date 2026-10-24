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

// Las entregas del semestre 2026-2 se guardan en la misma coleccion con el prefijo
// TERM en el id del documento, para separarlas de las de 2026-1 sin cambiar las reglas.
const TERM = '2026-2_';

const app = initializeApp(firebaseConfig);
const db = getFirestore(app);

// Adaptador con la misma interfaz que window.storage
// para que el resto del codigo no cambie
export const storage = {
  async set(key, value, shared = true) {
    const cedula = key.replace(/^doe_taller_v6:/, '');
    await setDoc(doc(db, 'submissions', TERM + cedula), {
      payload: value,
      updatedAt: new Date().toISOString()
    });
    return { key, value, shared };
  },

  async get(key, shared = true) {
    const cedula = key.replace(/^doe_taller_v6:/, '');
    const snap = await getDoc(doc(db, 'submissions', TERM + cedula));
    if (!snap.exists()) throw new Error('Not found');
    return { key, value: snap.data().payload, shared };
  },

  async delete(key, shared = true) {
    const cedula = key.replace(/^doe_taller_v6:/, '');
    await deleteDoc(doc(db, 'submissions', TERM + cedula));
    return { key, deleted: true, shared };
  },

  async list(prefix = '', shared = true) {
    const snap = await getDocs(collection(db, 'submissions'));
    const keys = snap.docs
      .filter(d => d.id.startsWith(TERM))
      .map(d => `doe_taller_v6:${d.id.slice(TERM.length)}`);
    return { keys, prefix, shared };
  }
};

// Lo conectamos a window.storage para que el codigo existente funcione
if (typeof window !== 'undefined') {
  window.storage = storage;
}
