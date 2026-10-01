import { invoke } from '@tauri-apps/api/core';
import { listen } from '@tauri-apps/api/event';
import { firebaseAuth } from '../firebase';
import callbackTemplate from './callback.template';

const {
  openGoogleSignIn,
  googleSignIn,
  signOut,
  signInWithEmail: firebaseSignInWithEmail,
  signUpWithEmail: firebaseSignUpWithEmail,
  resetPassword: firebaseResetPassword,
  sendVerificationEmail: firebaseSendVerificationEmail,
  isEmailVerified: firebaseIsEmailVerified,
  reloadUser: firebaseReloadUser
} = firebaseAuth;

export const loginWithGoogle = async () => {
  // Set up listener for OAuth callback
  const unlisten = await listen('oauth://url', async (data) => {
    try {
      await googleSignIn(data.payload as string);
      unlisten(); // Clean up listener after successful login
    } catch (error) {
      console.error('Google sign in error:', error);
      unlisten();
      throw error;
    }
  });
  
  try {
    // Start Tauri OAuth server
    const port = await invoke('plugin:oauth|start', {
      config: {
        // Use callback template for friendly response page
        response: callbackTemplate,
      },
    });
    
    // Open browser for Google sign-in
    await openGoogleSignIn(port as string);
  } catch (error) {
    unlisten(); // Clean up listener if error occurs
    throw error;
  }
};

// Keep login as alias for backward compatibility
export const login = loginWithGoogle;

// Email/Password authentication
export const signInWithEmail = async (email: string, password: string) => {
  return firebaseSignInWithEmail(email, password);
};

export const signUpWithEmail = async (email: string, password: string) => {
  return firebaseSignUpWithEmail(email, password);
};

export const resetPassword = async (email: string) => {
  return firebaseResetPassword(email);
};

export const sendVerificationEmail = async () => {
  return firebaseSendVerificationEmail();
};

export const isEmailVerified = () => {
  return firebaseIsEmailVerified();
};

export const reloadUser = async () => {
  return firebaseReloadUser();
};

export const logout = () => {
  return signOut();
};

export { getCurrentUser, onAuthChange } from '../firebase/auth'; 