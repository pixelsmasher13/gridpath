import { useEffect, useState } from 'react';
import { useRouter } from 'next/router';
import Head from 'next/head';
import Image from 'next/image';

export default function AuthPage() {
  const router = useRouter();
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    // Check if we have a token from the desktop app
    const { token, return_to } = router.query;
    
    if (token && typeof token === 'string') {
      // Validate the JWT token first
      validateAndStoreToken(token, return_to as string);
    }
  }, [router]);

  const validateAndStoreToken = async (token: string, returnTo?: string) => {
    setLoading(true);
    setError(null);

    console.log('🔐 Validating token:', {
      tokenLength: token?.length,
      tokenPreview: token?.substring(0, 50) + '...',
      apiUrl: process.env.NEXT_PUBLIC_API_URL,
      returnTo
    });

    try {
      // Validate the token by making a test API call
      const response = await fetch(`${process.env.NEXT_PUBLIC_API_URL}/api/billing/credits-status`, {
        headers: {
          'Authorization': `Bearer ${token}`,
        },
      });

      console.log('📡 Validation response:', {
        status: response.status,
        ok: response.ok,
        statusText: response.statusText
      });

      if (!response.ok) {
        const errorText = await response.text();
        console.error('❌ Token validation failed:', {
          status: response.status,
          error: errorText
        });
        throw new Error(`Token validation failed: ${response.status} - ${errorText}`);
      }

      const data = await response.json();
      console.log('✅ Token validated successfully:', data);

      // Token is valid, store it
      localStorage.setItem('auth-token', token);
      document.cookie = `auth-token=${token}; path=/; max-age=${7 * 24 * 60 * 60}; SameSite=Strict`;
      
      // Redirect to dashboard or return URL
      const redirectUrl = returnTo || '/dashboard';
      console.log('🚀 Redirecting to:', redirectUrl);
      router.push(redirectUrl);
    } catch (err: any) {
      console.error('Token validation failed:', err);
      setError(`Authentication failed: ${err?.message || 'Unknown error'}. Please try again.`);
      setLoading(false);
    }
  };

  const handleDesktopAuth = () => {
    setLoading(true);
    setError(null);

    // Generate a unique state for this auth session
    const state = Math.random().toString(36).substring(7);
    sessionStorage.setItem('auth-state', state);

    // Redirect back to desktop app for authentication
    const returnUrl = `${window.location.origin}/auth/callback`;
    const desktopAuthUrl = `heelix://auth?return_url=${encodeURIComponent(returnUrl)}&state=${state}`;
    
    // Try to open desktop app
    window.location.href = desktopAuthUrl;

    // Fallback message after 3 seconds if app doesn't open
    setTimeout(() => {
      if (document.visibilityState === 'visible') {
        setLoading(false);
        setError('Could not open Linefox app. Please make sure it is installed and running.');
      }
    }, 3000);
  };

  return (
    <>
      <Head>
        <title>Sign In - Linefox Billing</title>
      </Head>
      
      <div className="min-h-screen bg-gradient-to-br from-primary-50 to-primary-100 flex items-center justify-center px-4" style={{ fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif" }}>
        <div className="max-w-md w-full">
          <div className="card">
            <div className="text-center mb-8">
              <div className="flex items-center justify-center space-x-2 mb-4">
                <Image
                  src="/logo-black.png"
                  alt="Linefox"
                  width={40}
                  height={40}
                />
                <span className="text-2xl font-semibold text-gray-900">Linefox</span>
              </div>
              <h1 className="text-xl font-medium text-gray-700 mb-2">Billing Portal</h1>
              <p className="text-gray-500">Sign in to manage your billing and usage</p>
            </div>

            {error && (
              <div className="mb-6 p-4 bg-red-50 border border-red-200 rounded-lg">
                <p className="text-red-700 text-sm">{error}</p>
              </div>
            )}

            <div className="space-y-4">
              <button
                onClick={handleDesktopAuth}
                disabled={loading}
                className="w-full btn-primary py-3 flex items-center justify-center space-x-2"
              >
                {loading ? (
                  <>
                    <svg className="animate-spin h-5 w-5" viewBox="0 0 24 24">
                      <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" fill="none" />
                      <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z" />
                    </svg>
                    <span>Connecting to Linefox...</span>
                  </>
                ) : (
                  <>
                    <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12l2 2 4-4m5.618-4.016A11.955 11.955 0 0112 2.944a11.955 11.955 0 01-8.618 3.04A12.02 12.02 0 003 9c0 5.591 3.824 10.29 9 11.622 5.176-1.332 9-6.03 9-11.622 0-1.042-.133-2.052-.382-3.016z" />
                    </svg>
                    <span>Sign in with Linefox Desktop</span>
                  </>
                )}
              </button>

              <div className="relative">
                <div className="absolute inset-0 flex items-center">
                  <div className="w-full border-t border-gray-300"></div>
                </div>
                <div className="relative flex justify-center text-sm">
                  <span className="px-2 bg-white text-gray-500">Or</span>
                </div>
              </div>

              <div className="text-center">
                <p className="text-sm text-gray-600">
                  Don't have Linefox installed?{' '}
                  <a href="https://heelix.com/download" className="text-primary-600 hover:text-primary-700 font-medium">
                    Download now
                  </a>
                </p>
              </div>
            </div>
          </div>

          <div className="mt-8 text-center">
            <p className="text-sm text-gray-600">
              Need help?{' '}
              <a href="https://heelix.com/support" className="text-primary-600 hover:text-primary-700">
                Contact support
              </a>
            </p>
          </div>
        </div>
      </div>
    </>
  );
}