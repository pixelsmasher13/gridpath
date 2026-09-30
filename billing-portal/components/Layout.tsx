import { useRouter } from 'next/router';
import Link from 'next/link';
import Image from 'next/image';
import { useEffect, useState } from 'react';
import jwt from 'jsonwebtoken';

interface UserInfo {
  email: string;
  name?: string;
  uid: string;
}

export function Layout({ children }: { children: React.ReactNode }) {
  const router = useRouter();
  const [user, setUser] = useState<UserInfo | null>(null);
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false);

  useEffect(() => {
    // Decode JWT to get user info
    const token = localStorage.getItem('auth-token');
    if (token) {
      try {
        // Decode without verification (since we already validated with the API)
        const decoded = jwt.decode(token) as any;
        if (decoded) {
          setUser({
            email: decoded.email,
            name: decoded.name,
            uid: decoded.uid,
          });
        }
      } catch (error) {
        console.error('Error decoding token:', error);
      }
    }
  }, []);

  const handleSignOut = () => {
    localStorage.removeItem('auth-token');
    document.cookie = 'auth-token=; path=/; expires=Thu, 01 Jan 1970 00:00:01 GMT';
    router.push('/auth');
  };

  return (
    <div className="min-h-screen bg-gray-50" style={{ fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif" }}>
      {/* Navigation */}
      <nav className="bg-white shadow-sm border-b border-gray-200">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="flex justify-between items-center h-14">
            {/* Left side - Logo and nav links */}
            <div className="flex items-center">
              <Link href="/dashboard" className="flex items-center space-x-2">
                <Image
                  src="/logo-black.png"
                  alt="Linefox"
                  width={28}
                  height={28}
                  className="w-7 h-7"
                />
                <span className="text-lg font-semibold text-gray-900">Linefox</span>
              </Link>

              {/* Desktop nav links */}
              <div className="hidden md:flex ml-8 items-center space-x-1">
                <Link
                  href="/dashboard"
                  className={`px-3 py-1.5 text-sm font-medium rounded-md transition-colors ${
                    router.pathname === '/dashboard'
                      ? 'text-primary-600 bg-primary-50'
                      : 'text-gray-600 hover:text-gray-900 hover:bg-gray-50'
                  }`}
                >
                  Dashboard
                </Link>
                <Link
                  href="/usage"
                  className={`px-3 py-1.5 text-sm font-medium rounded-md transition-colors ${
                    router.pathname === '/usage'
                      ? 'text-primary-600 bg-primary-50'
                      : 'text-gray-600 hover:text-gray-900 hover:bg-gray-50'
                  }`}
                >
                  Usage
                </Link>
              </div>
            </div>

            {/* Right side - User info (desktop) and hamburger (mobile) */}
            <div className="flex items-center space-x-3">
              {/* Desktop user info */}
              <div className="hidden md:flex items-center space-x-3">
                {user && (
                  <div className="flex items-center space-x-2 px-2 py-1 bg-gray-50 rounded-lg">
                    <div className="flex items-center justify-center w-7 h-7 bg-primary-600 text-white rounded-full text-xs font-medium">
                      {user.name ? user.name[0].toUpperCase() : user.email[0].toUpperCase()}
                    </div>
                    <span className="text-sm text-gray-700 max-w-[120px] truncate">
                      {user.email}
                    </span>
                  </div>
                )}
                <button
                  onClick={handleSignOut}
                  className="text-gray-600 hover:text-gray-900 text-sm font-medium px-3 py-1.5 rounded-md hover:bg-gray-100 transition-colors"
                >
                  Sign Out
                </button>
              </div>

              {/* Mobile hamburger button */}
              <button
                onClick={() => setMobileMenuOpen(!mobileMenuOpen)}
                className="md:hidden p-2 rounded-md text-gray-600 hover:text-gray-900 hover:bg-gray-100"
              >
                {mobileMenuOpen ? (
                  <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
                  </svg>
                ) : (
                  <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 6h16M4 12h16M4 18h16" />
                  </svg>
                )}
              </button>
            </div>
          </div>
        </div>

        {/* Mobile menu */}
        {mobileMenuOpen && (
          <div className="md:hidden border-t border-gray-200 bg-white">
            <div className="px-4 py-3 space-y-1">
              <Link
                href="/dashboard"
                onClick={() => setMobileMenuOpen(false)}
                className={`block px-3 py-2 text-base font-medium rounded-md ${
                  router.pathname === '/dashboard'
                    ? 'text-primary-600 bg-primary-50'
                    : 'text-gray-700 hover:bg-gray-50'
                }`}
              >
                Dashboard
              </Link>
              <Link
                href="/usage"
                onClick={() => setMobileMenuOpen(false)}
                className={`block px-3 py-2 text-base font-medium rounded-md ${
                  router.pathname === '/usage'
                    ? 'text-primary-600 bg-primary-50'
                    : 'text-gray-700 hover:bg-gray-50'
                }`}
              >
                Usage History
              </Link>
            </div>

            {/* Mobile user section */}
            <div className="border-t border-gray-200 px-4 py-3">
              {user && (
                <div className="flex items-center space-x-3 mb-3">
                  <div className="flex items-center justify-center w-8 h-8 bg-primary-600 text-white rounded-full text-sm font-medium">
                    {user.name ? user.name[0].toUpperCase() : user.email[0].toUpperCase()}
                  </div>
                  <div className="flex-1 min-w-0">
                    <p className="text-sm font-medium text-gray-900 truncate">
                      {user.name || user.email.split('@')[0]}
                    </p>
                    <p className="text-xs text-gray-500 truncate">{user.email}</p>
                  </div>
                </div>
              )}
              <button
                onClick={handleSignOut}
                className="w-full text-left px-3 py-2 text-base font-medium text-gray-700 hover:bg-gray-50 rounded-md"
              >
                Sign Out
              </button>
            </div>
          </div>
        )}
      </nav>

      {/* Main Content */}
      <main>{children}</main>
    </div>
  );
}
