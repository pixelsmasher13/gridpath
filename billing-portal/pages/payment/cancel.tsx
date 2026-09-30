import Head from 'next/head';
import Link from 'next/link';

export default function PaymentCancel() {
  return (
    <>
      <Head>
        <title>Payment Cancelled - Linefox Billing</title>
      </Head>

      <div className="min-h-screen bg-gradient-to-br from-gray-50 to-gray-100 flex items-center justify-center px-4">
        <div className="max-w-md w-full">
          <div className="card text-center">
            <div className="w-16 h-16 bg-amber-100 rounded-full flex items-center justify-center mx-auto mb-4">
              <svg className="w-8 h-8 text-amber-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z" />
              </svg>
            </div>

            <h1 className="text-2xl font-bold text-gray-900 mb-2">Payment Cancelled</h1>
            
            <p className="text-gray-600 mb-6">
              Your payment was cancelled and no charges were made to your account.
            </p>

            <div className="space-y-3">
              <Link href="/dashboard" className="block w-full btn-primary py-3">
                Return to Dashboard
              </Link>
              
              <button
                onClick={() => window.history.back()}
                className="block w-full btn-secondary py-3"
              >
                Try Again
              </button>
            </div>

            <div className="mt-6 pt-6 border-t border-gray-200">
              <p className="text-sm text-gray-600 mb-3">
                Having trouble with payment?
              </p>
              <ul className="text-xs text-gray-500 space-y-1 text-left max-w-xs mx-auto">
                <li>• Check your card details are correct</li>
                <li>• Ensure your card has sufficient funds</li>
                <li>• Try a different payment method</li>
                <li>• Contact your bank if issues persist</li>
              </ul>
            </div>
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