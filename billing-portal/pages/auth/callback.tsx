import { useEffect } from 'react';
import { useRouter } from 'next/router';
import Head from 'next/head';

export default function AuthCallback() {
  const router = useRouter();

  useEffect(() => {
    const { token, state, error } = router.query;

    if (error) {
      // Handle auth error
      router.push(`/auth?error=${error}`);
      return;
    }

    if (token && typeof token === 'string') {
      // Verify state matches
      const savedState = sessionStorage.getItem('auth-state');
      if (state !== savedState && savedState) {
        router.push('/auth?error=Invalid state');
        return;
      }

      // Store token
      localStorage.setItem('auth-token', token);
      document.cookie = `auth-token=${token}; path=/; max-age=${7 * 24 * 60 * 60}`;
      
      // Clear state
      sessionStorage.removeItem('auth-state');
      
      // Redirect to dashboard
      router.push('/dashboard');
    }
  }, [router]);

  return (
    <>
      <Head>
        <title>Authenticating... - Linefox Billing</title>
      </Head>
      
      <div className="min-h-screen bg-gradient-to-br from-primary-50 to-primary-100 flex items-center justify-center">
        <div className="text-center">
          <svg className="animate-spin h-12 w-12 mx-auto mb-4 text-primary-600" viewBox="0 0 24 24">
            <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" fill="none" />
            <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z" />
          </svg>
          <h2 className="text-xl font-semibold text-gray-900">Authenticating...</h2>
          <p className="mt-2 text-gray-600">Please wait while we sign you in</p>
        </div>
      </div>
    </>
  );
}