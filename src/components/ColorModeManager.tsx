import { useEffect } from 'react';
import { useColorMode } from '@chakra-ui/react';

export const ColorModeManager = () => {
  const { colorMode } = useColorMode();

  useEffect(() => {
    const root = document.documentElement;
    root.setAttribute('data-theme', colorMode);
    // Also set on body for good measure if needed by some libs
    document.body.setAttribute('data-theme', colorMode);
  }, [colorMode]);

  return null;
};

