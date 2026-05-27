import React, { useState, useEffect } from "react";
import { Routes, Route, Navigate } from 'react-router-dom';
import { SpreadsheetScreen } from './screens/SpreadsheetScreen';
import { SpreadsheetOnboarding } from './screens/SpreadsheetScreen/components/SpreadsheetOnboarding';
import { ColorModeManager } from './components/ColorModeManager';

export const AppContent: React.FC = () => {
  const [hasCompletedOnboarding, setHasCompletedOnboarding] = useState(false);

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
    <>
      <ColorModeManager />
      <AppContent />
    </>
  );
};
