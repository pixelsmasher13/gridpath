import { extendTheme, type ThemeConfig } from '@chakra-ui/react';

// Color mode config
const config: ThemeConfig = {
  initialColorMode: 'light',
  useSystemColorMode: false,
}

export const theme = extendTheme({
  config,
  fonts: {
    heading: `-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif`,
    body: `-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif`,
  },
  colors: {
    brand: {
      main: "#3363AD",
      100: "#3363AD1F",
      300: "#4070b8",
      400: "#3363AD",
      600: "#2a5489",
    },
  },
  components: {
    Button: {
      variants: {
        primary: {
          bg: "brand.400",
          color: "white", // Text color
          _hover: {
            bg: "brand.600", // Color on hover
          },
        },
      },
    },
  },
});
