import { open } from '@tauri-apps/plugin-shell';
import {
  getAuth,
  GoogleAuthProvider,
  signInWithCredential,
  signInWithEmailAndPassword,
  createUserWithEmailAndPassword,
  sendPasswordResetEmail,
  sendEmailVerification,
  onAuthStateChanged,
  User
} from 'firebase/auth';
import { firebaseApp } from './app';

const auth = getAuth(firebaseApp);

const openBrowserToConsent = (port: string) => {
  return open('https://accounts.google.com/o/oauth2/auth?' +
    'response_type=token&' +
    'client_id=441581008186-9klu9129oap07d7l8a63idt7l5ab0r24.apps.googleusercontent.com&' +
    `redirect_uri=http%3A//localhost:${port}&` +
    'scope=email%20profile%20openid&' +
    'prompt=consent'
  );
};

export const openGoogleSignIn = (port: string) => {
  return new Promise((resolve, reject) => {
    openBrowserToConsent(port).then(resolve).catch(reject);
  });
};

export const googleSignIn = async (payload: string) => {
  const url = new URL(payload);
  // Get `access_token` from redirect_uri param
  const access_token = new URLSearchParams(url.hash.substring(1)).get('access_token');

  if (!access_token) {
    throw new Error('No access token found in redirect URL');
  }

  const credential = GoogleAuthProvider.credential(null, access_token);

  try {
    const result = await signInWithCredential(auth, credential);
    return result.user;
  } catch (error: any) {
    console.error('Sign in error:', error.code, error.message);
    throw error;
  }
};

export const signOut = () => {
  return auth.signOut();
};

export const getCurrentUser = (): User | null => {
  return auth.currentUser;
};

export const onAuthChange = (callback: (user: User | null) => void) => {
  return onAuthStateChanged(auth, callback);
};

// Email/Password authentication methods
export const signInWithEmail = async (email: string, password: string) => {
  try {
    const result = await signInWithEmailAndPassword(auth, email, password);
    return result.user;
  } catch (error: any) {
    console.error('Email sign in error:', error.code, error.message);
    throw error;
  }
};

export const signUpWithEmail = async (email: string, password: string) => {
  try {
    const result = await createUserWithEmailAndPassword(auth, email, password);

    // Send verification email immediately after signup
    if (result.user && !result.user.emailVerified) {
      await sendEmailVerification(result.user);
      console.log('Verification email sent to:', result.user.email);
    }

    return result.user;
  } catch (error: any) {
    console.error('Email sign up error:', error.code, error.message);
    throw error;
  }
};

export const resetPassword = async (email: string) => {
  try {
    await sendPasswordResetEmail(auth, email);
  } catch (error: any) {
    console.error('Password reset error:', error.code, error.message);
    throw error;
  }
};

// Send verification email to current user
export const sendVerificationEmail = async () => {
  const user = auth.currentUser;
  if (user && !user.emailVerified) {
    try {
      await sendEmailVerification(user);
      return true;
    } catch (error: any) {
      console.error('Error sending verification email:', error.code, error.message);
      throw error;
    }
  }
  return false;
};

// Check if current user's email is verified
export const isEmailVerified = (): boolean => {
  const user = auth.currentUser;
  return user ? user.emailVerified : false;
};

// Reload user to get latest verification status
export const reloadUser = async () => {
  const user = auth.currentUser;
  if (user) {
    await user.reload();
    return user;
  }
  return null;
}; 