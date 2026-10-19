import '@/styles/globals.css';
import type { AppProps } from 'next/app';
import { useEffect } from 'react';
import { useRouter } from 'next/router';

export default function App({ Component, pageProps }: AppProps) {
  const router = useRouter();

  useEffect(() => {
    // Check for auth token on mount
    const token = localStorage.getItem('auth-token');
    const isAuthPage = router.pathname === '/auth' || router.pathname === '/auth/callback';
    
    if (!token && !isAuthPage) {
      // Redirect to auth if no token
      router.push('/auth');
    }
  }, [router]);

  return <Component {...pageProps} />;
}