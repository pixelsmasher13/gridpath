import { useEffect, useState } from 'react';
import { useRouter } from 'next/router';
import Head from 'next/head';
import useSWR from 'swr';
import { billingAPI, BillingStatus } from '@/lib/api';
import { CreditBalanceCard } from '@/components/CreditBalanceCard';
import { PurchaseCreditsCard } from '@/components/PurchaseCreditsCard';
import { Layout } from '@/components/Layout';

const fetcher = async () => {
  const token = localStorage.getItem('auth-token');
  if (!token) throw new Error('No auth token');
  billingAPI.setToken(token);
  return billingAPI.getStatus();
};

export default function Dashboard() {
  const router = useRouter();
  const [userEmail, setUserEmail] = useState<string>('');
  const { data: status, error, mutate } = useSWR<BillingStatus>('billing-status', fetcher, {
    refreshInterval: 30000, // Refresh every 30 seconds
  });

  useEffect(() => {
    const token = localStorage.getItem('auth-token');
    if (!token) {
      router.push('/auth');
    } else {
      // Decode token to get user email
      try {
        const decoded = JSON.parse(atob(token.split('.')[1]));
        setUserEmail(decoded.email || '');
      } catch (e) {
        console.error('Error decoding token:', e);
      }
    }
  }, [router]);

  if (error) {
    return (
      <Layout>
        <div className="max-w-7xl mx-auto px-4 py-4 sm:py-6">
          <div className="bg-red-50 border border-red-200 rounded-lg p-4 sm:p-6">
            <h2 className="text-base sm:text-lg font-semibold text-red-900 mb-2">Error loading billing data</h2>
            <p className="text-sm text-red-700">{error.message}</p>
            <button
              onClick={() => mutate()}
              className="mt-4 btn-primary"
            >
              Try Again
            </button>
          </div>
        </div>
      </Layout>
    );
  }

  if (!status) {
    return (
      <Layout>
        <div className="max-w-7xl mx-auto px-4 py-4 sm:py-6">
          <div className="flex items-center justify-center py-12">
            <svg className="animate-spin h-6 w-6 sm:h-8 sm:w-8 text-primary-600" viewBox="0 0 24 24">
              <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" fill="none" />
              <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z" />
            </svg>
            <span className="ml-3 text-sm sm:text-base text-gray-600">Loading...</span>
          </div>
        </div>
      </Layout>
    );
  }

  return (
    <Layout>
      <Head>
        <title>Billing - Linefox</title>
      </Head>

      <div className="max-w-7xl mx-auto px-4 py-4 sm:py-6 lg:px-8">
        {/* Header - Compact on mobile */}
        <div className="mb-4 sm:mb-6">
          <h1 className="text-xl sm:text-2xl font-bold text-gray-900">Billing</h1>
          <p className="mt-1 text-sm text-gray-500">Manage credits and payments</p>
        </div>

        {/* Alerts - Compact on mobile */}
        {status.needsCredits && !status.hasPaymentMethod && (
          <div className="mb-4 bg-amber-50 border border-amber-200 rounded-lg p-3 sm:p-4">
            <div className="flex items-start">
              <svg className="h-4 w-4 sm:h-5 sm:w-5 text-amber-400 mt-0.5 flex-shrink-0" viewBox="0 0 20 20" fill="currentColor">
                <path fillRule="evenodd" d="M8.257 3.099c.765-1.36 2.722-1.36 3.486 0l5.58 9.92c.75 1.334-.213 2.98-1.742 2.98H4.42c-1.53 0-2.493-1.646-1.743-2.98l5.58-9.92zM11 13a1 1 0 11-2 0 1 1 0 012 0zm-1-8a1 1 0 00-1 1v3a1 1 0 002 0V6a1 1 0 00-1-1z" clipRule="evenodd" />
              </svg>
              <div className="ml-2 sm:ml-3">
                <p className="text-sm font-medium text-amber-800">Purchase credits to continue</p>
              </div>
            </div>
          </div>
        )}

        {status.needsCredits && status.hasPaymentMethod && !status.autoRecharge.enabled && (
          <div className="mb-4 bg-blue-50 border border-blue-200 rounded-lg p-3 sm:p-4">
            <div className="flex items-start">
              <svg className="h-4 w-4 sm:h-5 sm:w-5 text-blue-400 mt-0.5 flex-shrink-0" viewBox="0 0 20 20" fill="currentColor">
                <path fillRule="evenodd" d="M18 10a8 8 0 11-16 0 8 8 0 0116 0zm-7-4a1 1 0 11-2 0 1 1 0 012 0zM9 9a1 1 0 000 2v3a1 1 0 001 1h1a1 1 0 100-2v-3a1 1 0 00-1-1H9z" clipRule="evenodd" />
              </svg>
              <div className="ml-2 sm:ml-3">
                <p className="text-sm font-medium text-blue-800">Running low on credits</p>
              </div>
            </div>
          </div>
        )}

        {/* Main Content - Stack on mobile, grid on desktop */}
        <div className="space-y-4 sm:space-y-6 lg:grid lg:grid-cols-2 lg:gap-6 lg:space-y-0">
          {/* Balance & Purchase - Primary focus */}
          <div className="space-y-4 sm:space-y-6">
            <CreditBalanceCard status={status} />
            <PurchaseCreditsCard status={status} onPurchaseSuccess={() => mutate()} />
          </div>

          {/* Settings & History - Secondary */}
          <div className="space-y-4 sm:space-y-6">
            {/* Billing Settings */}
            <div className="card">
              <h3 className="text-base sm:text-lg font-semibold text-gray-900 mb-3 sm:mb-4">Settings</h3>
              <div className="space-y-2 sm:space-y-3">
                <button
                  onClick={async () => {
                    try {
                      const token = localStorage.getItem('auth-token');
                      const response = await fetch(`${process.env.NEXT_PUBLIC_API_URL}/api/billing/create-portal-session`, {
                        method: 'POST',
                        headers: {
                          'Authorization': `Bearer ${token}`,
                          'Content-Type': 'application/json',
                        },
                      });
                      const data = await response.json();
                      if (data.url) {
                        window.location.href = data.url;
                      }
                    } catch (error) {
                      console.error('Error opening Stripe portal:', error);
                    }
                  }}
                  className="w-full text-left px-3 sm:px-4 py-2.5 sm:py-3 border border-gray-200 rounded-lg hover:bg-gray-50 transition-colors flex justify-between items-center group"
                >
                  <div>
                    <p className="text-sm font-medium text-gray-900">Payment Methods</p>
                    <p className="text-xs text-gray-500 hidden sm:block">Manage cards and billing info</p>
                  </div>
                  <svg className="w-4 h-4 sm:w-5 sm:h-5 text-gray-400 group-hover:text-gray-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 5l7 7-7 7" />
                  </svg>
                </button>
                <button
                  onClick={async () => {
                    try {
                      const token = localStorage.getItem('auth-token');
                      const response = await fetch(`${process.env.NEXT_PUBLIC_API_URL}/api/billing/create-portal-session`, {
                        method: 'POST',
                        headers: {
                          'Authorization': `Bearer ${token}`,
                          'Content-Type': 'application/json',
                        },
                      });
                      const data = await response.json();
                      if (data.url) {
                        window.location.href = data.url;
                      }
                    } catch (error) {
                      console.error('Error opening Stripe portal:', error);
                    }
                  }}
                  className="w-full text-left px-3 sm:px-4 py-2.5 sm:py-3 border border-gray-200 rounded-lg hover:bg-gray-50 transition-colors flex justify-between items-center group"
                >
                  <div>
                    <p className="text-sm font-medium text-gray-900">Invoices</p>
                    <p className="text-xs text-gray-500 hidden sm:block">View receipts and history</p>
                  </div>
                  <svg className="w-4 h-4 sm:w-5 sm:h-5 text-gray-400 group-hover:text-gray-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 5l7 7-7 7" />
                  </svg>
                </button>
              </div>
            </div>

            {/* Recent Purchases - Hide on mobile if empty */}
            {status.recentPurchases && status.recentPurchases.length > 0 && (
              <div className="card hidden sm:block">
                <h3 className="text-base sm:text-lg font-semibold text-gray-900 mb-3 sm:mb-4">Recent Purchases</h3>
                <div className="space-y-2">
                  {status.recentPurchases.slice(0, 5).map((purchase: any, index: number) => {
                    const amount = purchase.amountDollars || purchase.amount_dollars || (purchase.amount_cents ? purchase.amount_cents / 100 : 0);
                    const date = purchase.createdAt || purchase.created_at;
                    const validDate = date ? new Date(date) : null;

                    return (
                      <div key={index} className="flex justify-between items-center py-2 border-b border-gray-100 last:border-0">
                        <div>
                          <p className="text-sm font-medium text-gray-900">
                            ${typeof amount === 'number' ? amount.toFixed(2) : '0.00'}
                          </p>
                          <p className="text-xs text-gray-500">
                            {validDate && !isNaN(validDate.getTime()) ? validDate.toLocaleDateString() : 'Recent'}
                          </p>
                        </div>
                        <span className="text-xs px-2 py-0.5 bg-green-100 text-green-800 rounded-full">
                          {purchase.status || 'completed'}
                        </span>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}
          </div>
        </div>
      </div>
    </Layout>
  );
}
