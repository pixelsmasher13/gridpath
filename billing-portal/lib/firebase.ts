import admin from 'firebase-admin';

let firebaseApp: admin.app.App | null = null;

export function initializeFirebase() {
  if (firebaseApp) {
    return firebaseApp;
  }

  try {
    const privateKey = process.env.FIREBASE_PRIVATE_KEY?.replace(/\\n/g, '\n');

    if (!privateKey || !process.env.FIREBASE_PROJECT_ID || !process.env.FIREBASE_CLIENT_EMAIL) {
      console.warn('Firebase environment variables not configured - Firebase auth will not work');
      return null;
    }

    firebaseApp = admin.initializeApp({
      credential: admin.credential.cert({
        projectId: process.env.FIREBASE_PROJECT_ID,
        privateKey: privateKey,
        clientEmail: process.env.FIREBASE_CLIENT_EMAIL,
      }),
      projectId: process.env.FIREBASE_PROJECT_ID,
    });

    console.log('Firebase Admin initialized successfully for billing portal');
    return firebaseApp;
  } catch (error) {
    console.error('Error initializing Firebase Admin:', error);
    return null;
  }
}

export async function verifyFirebaseToken(idToken: string): Promise<{
  valid: boolean;
  uid?: string;
  email?: string;
  name?: string;
}> {
  try {
    const app = initializeFirebase();
    if (!app) {
      return { valid: false };
    }

    const decodedToken = await admin.auth().verifyIdToken(idToken);

    return {
      valid: true,
      uid: decodedToken.uid,
      email: decodedToken.email,
      name: decodedToken.name,
    };
  } catch (error) {
    console.error('Firebase token verification failed:', error);
    return { valid: false };
  }
}
