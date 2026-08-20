import { initializeApp } from 'firebase/app';

const firebaseConfig = {
  apiKey: "AIzaSyAPrXGgWqee7s5H2V8jY5-MMpKI0p2hXrU",
  authDomain: "gridpath.firebaseapp.com",
  projectId: "gridpath",
  storageBucket: "gridpath.firebasestorage.app",
  messagingSenderId: "441581008186",
  appId: "1:441581008186:web:811e8808f72141ff7d8c03",
  measurementId: "G-W68WY2W0ZC"
};

export const firebaseApp = initializeApp(firebaseConfig);
