import React, { useState, useEffect } from "react";
import { AuthProvider, useAuth } from './Providers/AuthProvider';
import { LoginScreen } from './screens/LoginScreen';
import { Routes, Route, Navigate } from 'react-router-dom';
import { SpreadsheetScreen } from './screens/SpreadsheetScreen';
import { SpreadsheetOnboarding } from './screens/SpreadsheetScreen/components/SpreadsheetOnboarding';
import { EmailVerificationScreen } from './screens/EmailVerificationScreen';
import { ColorModeManager } from './components/ColorModeManager';

export const AppContent: React.FC = () => {
  const [hasCompletedOnboarding, setHasCompletedOnboarding] = useState(false);
  const { user, firebaseUser, isLoading } = useAuth();

  useEffect(() => {
    const onboardingCompleted = localStorage.getItem("onboardingCompleted");
    if (onboardingCompleted) {
      setHasCompletedOnboarding(true);
    }
  }, []);

  const completeOnboarding = () => {
    setHasCompletedOnboarding(true);
    localStorage.setItem("onboardingCompleted", "true");
  };

  if (isLoading) {
    return (
      <div style={{
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        height: '100vh',
        background: '#1e1e1e',
        color: '#d4d4d4',
      }}>
        Loading…
      </div>
    );
  }

  if (!user) {
    return <LoginScreen />;
  }

  // Email/password users have to verify before continuing; Google users are auto-verified.
  if (firebaseUser && firebaseUser.email && !firebaseUser.emailVerified) {
    const isEmailPasswordUser = !firebaseUser.providerData.some(
      provider => provider.providerId === 'google.com'
    );
    if (isEmailPasswordUser) {
      return <EmailVerificationScreen />;
    }
  }

  if (!hasCompletedOnboarding) {
    return <SpreadsheetOnboarding onComplete={completeOnboarding} />;
  }

  return (
    <Routes>
      <Route path="/" element={<SpreadsheetScreen />} />
      <Route path="*" element={<Navigate to="/" replace />} />
    </Routes>
  );
};

export const App: React.FC = () => {
  return (
    <AuthProvider>
      <ColorModeManager />
      <AppContent />
    </AuthProvider>
  );
};
