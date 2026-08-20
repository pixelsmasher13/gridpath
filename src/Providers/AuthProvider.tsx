import React, { createContext, useContext, useState, useEffect, type FC, type ReactNode } from 'react';
import { invoke } from '@tauri-apps/api/core';
import { useToast } from '@chakra-ui/react';
import { User as FirebaseUser } from 'firebase/auth';
import {
  login as firebaseLogin,
  loginWithGoogle,
  signInWithEmail,
  signUpWithEmail,
  resetPassword,
  logout as firebaseLogout,
  onAuthChange
} from '../services/auth';
import { usageTracking, initializeUsageTracking } from '../services/usageTracking';

interface User {
  userId: string;
  email: string;
  authToken: string;
  expiresAt: string;
  firebaseUid?: string;
}

interface AuthContextType {
  user: User | null;
  firebaseUser: FirebaseUser | null;
  isLoading: boolean;
  login: () => Promise<void>;
  loginWithEmail: (email: string, password: string) => Promise<void>;
  signUpWithEmail: (email: string, password: string) => Promise<void>;
  resetPassword: (email: string) => Promise<void>;
  logout: () => Promise<void>;
  checkAuth: () => Promise<void>;
}

const AuthContext = createContext<AuthContextType | undefined>(undefined);

export const useAuth = () => {
  const context = useContext(AuthContext);
  if (!context) {
    throw new Error('useAuth must be used within AuthProvider');
  }
  return context;
};

interface AuthProviderProps {
  children: ReactNode;
}

export const AuthProvider: FC<AuthProviderProps> = ({ children }) => {
  const [user, setUser] = useState<User | null>(null);
  const [firebaseUser, setFirebaseUser] = useState<FirebaseUser | null>(null);
  const [isLoading, setIsLoading] = useState(true);
  const toast = useToast();

  const syncWithBackend = async (fbUser: FirebaseUser) => {
    try {
      const idToken = await fbUser.getIdToken();
      
      // Exchange Firebase token for proxy JWT token
      console.log('🔐 Exchanging Firebase token for proxy JWT...');
      console.log('🔑 Firebase ID token (first 20 chars):', idToken.substring(0, 20) + '...');
      
      const proxyResponse = await fetch('https://gridpath-theta.vercel.app/api/auth/google', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ idToken }),
      });
      
      console.log('📡 Proxy response status:', proxyResponse.status);
      
      if (!proxyResponse.ok) {
        const errorData = await proxyResponse.json().catch(() => ({}));
        console.error('❌ Proxy authentication failed:', errorData);
        throw new Error(errorData.error || 'Failed to authenticate with proxy server');
      }
      
      const proxyData = await proxyResponse.json();
      console.log('📋 Proxy response data:', { success: proxyData.success, hasToken: !!proxyData.token });
      
      if (!proxyData.success || !proxyData.token) {
        throw new Error('Invalid response from proxy server');
      }
      
      console.log('✅ Successfully got JWT token from proxy (first 20 chars):', proxyData.token.substring(0, 20) + '...');
      
      // Set the auth token for usage tracking and localStorage
      usageTracking.setAuthToken(proxyData.token);
      localStorage.setItem('authToken', proxyData.token);
      
      // Now store the proxy JWT token in the backend
      const authResult = await invoke<User>('authenticate_user', { 
        email: fbUser.email || '', 
        password: proxyData.token // Using proxy JWT token
      });
      
      console.log('💾 Stored auth in backend:', { userId: authResult.userId, email: authResult.email });
      
      setUser({
        ...authResult,
        firebaseUid: fbUser.uid
      });
      
      toast({
        title: 'Logged in successfully',
        description: `Welcome back, ${fbUser.displayName || fbUser.email}!`,
        status: 'success',
        duration: 3000,
        isClosable: true,
      });
      
      return authResult;
    } catch (error) {
      console.error('Failed to sync with backend:', error);
      
      // Clear user state to force login screen
      setUser(null);
      setFirebaseUser(null);
      setIsLoading(false);
      
      // Handle specific error cases
      let errorMessage = 'Failed to complete authentication';
      if (error instanceof Error) {
        // If it's a token error, give a friendly message
        if (error.message.includes('token') || error.message.includes('401')) {
          errorMessage = 'Your session has expired. Please sign in again.';
        } else {
          errorMessage = error.message;
        }
      }

      toast({
        title: 'Authentication Error',
        description: errorMessage,
        status: 'error',
        duration: 5000,
        isClosable: true,
      });
      
      // We don't rethrow to prevent unhandled promise rejections
      // The cleared state will naturally return user to login
      return null;
    }
  };

  const checkAuth = async () => {
    if (firebaseUser) {
      await syncWithBackend(firebaseUser);
    }
  };

  const login = async () => {
    try {
      setIsLoading(true);
      await firebaseLogin();
      // Firebase auth state change will handle the rest
      toast({
        title: 'Signing in with Google...',
        description: 'Please complete the sign-in process in your browser',
        status: 'info',
        duration: 5000,
        isClosable: true,
      });
    } catch (error) {
      toast({
        title: 'Login failed',
        description: error as string,
        status: 'error',
        duration: 5000,
        isClosable: true,
      });
      setIsLoading(false);
      throw error;
    }
  };

  const loginWithEmailHandler = async (email: string, password: string) => {
    try {
      setIsLoading(true);
      await signInWithEmail(email, password);
      // Firebase auth state change will handle the rest
    } catch (error) {
      setIsLoading(false);
      throw error;
    }
  };

  const signUpWithEmailHandler = async (email: string, password: string) => {
    try {
      setIsLoading(true);
      await signUpWithEmail(email, password);
      // Firebase auth state change will handle the rest
    } catch (error) {
      setIsLoading(false);
      throw error;
    }
  };

  const resetPasswordHandler = async (email: string) => {
    try {
      await resetPassword(email);
    } catch (error) {
      throw error;
    }
  };

  const logout = async () => {
    try {
      await firebaseLogout();
      await invoke('logout_user').catch(console.error);
      setUser(null);
      setFirebaseUser(null);
      localStorage.removeItem('authToken');
      toast({
        title: 'Logged out',
        description: 'You have been successfully logged out',
        status: 'info',
        duration: 3000,
        isClosable: true,
      });
    } catch (error) {
      console.error('Logout error:', error);
      // Clear user even if backend fails
      setUser(null);
      setFirebaseUser(null);
    }
  };

  useEffect(() => {
    // Initialize usage tracking
    initializeUsageTracking();
    
    // Listen to Firebase auth state changes
    const unsubscribe = onAuthChange(async (fbUser) => {
      setFirebaseUser(fbUser);
      
      if (fbUser) {
        await syncWithBackend(fbUser);
      } else {
        setUser(null);
      }
      
      setIsLoading(false);
    });

    return () => unsubscribe();
  }, []);

  return (
    <AuthContext.Provider value={{
      user,
      firebaseUser,
      isLoading,
      login,
      loginWithEmail: loginWithEmailHandler,
      signUpWithEmail: signUpWithEmailHandler,
      resetPassword: resetPasswordHandler,
      logout,
      checkAuth
    }}>
      {children}
    </AuthContext.Provider>
  );
}; 