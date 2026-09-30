import { useState } from 'react';
import { loadStripe } from '@stripe/stripe-js';
import { BillingStatus, billingAPI } from '@/lib/api';

const stripePromise = loadStripe(process.env.NEXT_PUBLIC_STRIPE_PUBLISHABLE_KEY!);

interface Props {
  status: BillingStatus;
  onPurchaseSuccess: () => void;
}

export function PurchaseCreditsCard({ status, onPurchaseSuccess }: Props) {
  const [loading, setLoading] = useState(false);
  const [amount, setAmount] = useState(5);
  const [error, setError] = useState<string | null>(null);

  const handlePurchase = async (useQuickPurchase = false) => {
    setLoading(true);
    setError(null);

    try {
      const token = localStorage.getItem('auth-token');
      if (!token) throw new Error('Not authenticated');

      billingAPI.setToken(token);
      const response = await billingAPI.purchaseCredits(amount, useQuickPurchase);

      if (response.checkoutUrl && !useQuickPurchase) {
        // Redirect to Stripe Checkout
        window.location.href = response.checkoutUrl;
      } else if (response.success) {
        // Quick purchase succeeded
        onPurchaseSuccess();
      }
    } catch (err: any) {
      setError(err.message || 'Failed to initiate purchase');
      setLoading(false);
    }
  };

  const creditOptions = [5, 10, 20, 50];
  const minutesPerDollar = 12; // $5/hour = 60 min/hour = 12 min/dollar

  return (
    <div className="card">
      <h2 className="text-base sm:text-lg font-semibold text-gray-900 mb-3 sm:mb-4">Buy Credits</h2>

      {error && (
        <div className="mb-3 p-2.5 sm:p-3 bg-red-50 border border-red-200 rounded-lg">
          <p className="text-sm text-red-700">{error}</p>
        </div>
      )}

      <div className="space-y-3 sm:space-y-4">
        {/* Amount Selection - 4 columns on mobile */}
        <div className="grid grid-cols-4 gap-1.5 sm:gap-2">
          {creditOptions.map((option) => (
            <button
              key={option}
              onClick={() => setAmount(option)}
              className={`p-2 sm:p-3 rounded-lg border-2 transition-all ${
                amount === option
                  ? 'border-primary-500 bg-primary-50'
                  : 'border-gray-200 hover:border-gray-300'
              }`}
            >
              <div className="text-base sm:text-lg font-semibold text-gray-900">${option}</div>
              <div className="text-xs text-gray-500">
                {option * minutesPerDollar}m
              </div>
            </button>
          ))}
        </div>

        {/* Custom Amount - Compact on mobile */}
        <div>
          <label className="block text-xs sm:text-sm text-gray-600 mb-1.5">
            Custom amount
          </label>
          <div className="relative">
            <span className="absolute left-3 top-1/2 transform -translate-y-1/2 text-gray-400 text-sm">
              $
            </span>
            <input
              type="number"
              min="5"
              step="5"
              value={amount}
              onChange={(e) => setAmount(Math.max(5, parseInt(e.target.value) || 5))}
              className="input pl-7 text-base"
              placeholder="5"
            />
          </div>
        </div>

        {/* Summary - More compact */}
        <div className="bg-gray-50 rounded-lg p-3 sm:p-4">
          <div className="flex justify-between items-center">
            <span className="text-sm text-gray-600">Total</span>
            <div className="text-right">
              <span className="text-lg sm:text-xl font-bold text-gray-900">${amount}</span>
              <span className="text-xs sm:text-sm text-gray-500 ml-2">= {amount * minutesPerDollar} min</span>
            </div>
          </div>
        </div>

        {/* Purchase Button - Prominent */}
        <div className="space-y-2">
          {status.hasPaymentMethod ? (
            <>
              <button
                onClick={() => handlePurchase(true)}
                disabled={loading}
                className="w-full btn-primary flex items-center justify-center py-2.5 sm:py-3 text-sm sm:text-base"
              >
                {loading ? (
                  <svg className="animate-spin h-5 w-5" viewBox="0 0 24 24">
                    <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" fill="none" />
                    <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z" />
                  </svg>
                ) : (
                  <>
                    <svg className="w-4 h-4 sm:w-5 sm:h-5 mr-1.5 sm:mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M13 10V3L4 14h7v7l9-11h-7z" />
                    </svg>
                    Quick Buy
                  </>
                )}
              </button>
              <button
                onClick={() => handlePurchase(false)}
                disabled={loading}
                className="w-full btn-secondary py-2 text-sm"
              >
                Different payment
              </button>
            </>
          ) : (
            <button
              onClick={() => handlePurchase(false)}
              disabled={loading}
              className="w-full btn-primary flex items-center justify-center py-2.5 sm:py-3 text-sm sm:text-base"
            >
              {loading ? (
                <svg className="animate-spin h-5 w-5" viewBox="0 0 24 24">
                  <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" fill="none" />
                  <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z" />
                </svg>
              ) : (
                <>
                  <svg className="w-4 h-4 sm:w-5 sm:h-5 mr-1.5 sm:mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M3 10h18M7 15h1m4 0h1m-7 4h12a3 3 0 003-3V8a3 3 0 00-3-3H6a3 3 0 00-3 3v8a3 3 0 003 3z" />
                  </svg>
                  Buy Credits
                </>
              )}
            </button>
          )}
        </div>

        {/* Info */}
        <p className="text-xs text-gray-400 text-center">
          Secure payment via Stripe
        </p>
      </div>
    </div>
  );
}
