import { useEffect } from 'react';
import { useRouter } from 'next/router';
import Head from 'next/head';
import Link from 'next/link';

export default function PaymentSuccess() {
  const router = useRouter();
  const { session_id, amount } = router.query;

  useEffect(() => {
    // Auto-redirect to dashboard after 5 seconds
    const timer = setTimeout(() => {
      router.push('/dashboard');
    }, 5000);

    return () => clearTimeout(timer);
  }, [router]);

  return (
    <>
      <Head>
        <title>Payment Successful - Linefox Billing</title>
      </Head>

      <div className="min-h-screen bg-gradient-to-br from-green-50 to-green-100 flex items-center justify-center px-4">
        <div className="max-w-md w-full">
          <div className="card text-center">
            <div className="w-16 h-16 bg-green-100 rounded-full flex items-center justify-center mx-auto mb-4">
              <svg className="w-8 h-8 text-green-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 13l4 4L19 7" />
              </svg>
            </div>

            <h1 className="text-2xl font-bold text-gray-900 mb-2">Payment Successful!</h1>
            
            <p className="text-gray-600 mb-6">
              Your credits have been added to your account.
              {amount && ` You purchased $${amount} worth of credits.`}
            </p>

            <div className="space-y-3">
              <Link href="/dashboard" className="block w-full btn-primary py-3">
                Go to Dashboard
              </Link>
              
              <Link href="/invoices" className="block w-full btn-secondary py-3">
                View Invoice
              </Link>
            </div>

            <p className="text-sm text-gray-500 mt-6">
              You will be redirected to the dashboard in a few seconds...
            </p>
          </div>

          <div className="mt-6 text-center">
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