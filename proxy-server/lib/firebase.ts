import admin from 'firebase-admin';

let firebaseApp: admin.app.App;

export function initializeFirebase() {
  if (firebaseApp) {
    return firebaseApp;
  }

  try {
    const privateKey = process.env.FIREBASE_PRIVATE_KEY?.replace(/\\n/g, '\n');
    
    if (!privateKey || !process.env.FIREBASE_PROJECT_ID || !process.env.FIREBASE_CLIENT_EMAIL) {
      throw new Error('Missing Firebase environment variables');
    }

    firebaseApp = admin.initializeApp({
      credential: admin.credential.cert({
        projectId: process.env.FIREBASE_PROJECT_ID,
        privateKey: privateKey,
        clientEmail: process.env.FIREBASE_CLIENT_EMAIL,
      }),
      projectId: process.env.FIREBASE_PROJECT_ID,
    });

    console.log('Firebase Admin initialized successfully');
    return firebaseApp;
  } catch (error) {
    console.error('Error initializing Firebase Admin:', error);
    throw error;
  }
}

export async function verifyFirebaseToken(idToken: string): Promise<{
  valid: boolean;
  uid?: string;
  email?: string;
  name?: string;
}> {
  try {
    initializeFirebase();
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