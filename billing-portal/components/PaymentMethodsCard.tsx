import { useState } from 'react';
import { PaymentMethod, billingAPI } from '@/lib/api';

interface Props {
  paymentMethods: PaymentMethod[];
  onUpdate: () => void;
}

export function PaymentMethodsCard({ paymentMethods, onUpdate }: Props) {
  const [loading, setLoading] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);

  const handleSetDefault = async (methodId: string) => {
    setLoading(methodId);
    setError(null);

    try {
      const token = localStorage.getItem('auth-token');
      if (!token) throw new Error('Not authenticated');
      
      billingAPI.setToken(token);
      await billingAPI.updatePaymentMethod(methodId);
      onUpdate();
    } catch (err: any) {
      setError(err.message || 'Failed to update payment method');
    } finally {
      setLoading(null);
    }
  };

  const handleRemove = async (methodId: string) => {
    if (!confirm('Are you sure you want to remove this payment method?')) return;
    
    setLoading(methodId);
    setError(null);

    try {
      const token = localStorage.getItem('auth-token');
      if (!token) throw new Error('Not authenticated');
      
      billingAPI.setToken(token);
      await billingAPI.removePaymentMethod(methodId);
      onUpdate();
    } catch (err: any) {
      setError(err.message || 'Failed to remove payment method');
    } finally {
      setLoading(null);
    }
  };

  const cardBrandIcons: Record<string, string> = {
    visa: '💳',
    mastercard: '💳',
    amex: '💳',
    discover: '💳',
  };

  return (
    <div className="card">
      <h2 className="text-lg font-semibold text-gray-900 mb-4">Payment Methods</h2>

      {error && (
        <div className="mb-4 p-3 bg-red-50 border border-red-200 rounded-lg">
          <p className="text-sm text-red-700">{error}</p>
        </div>
      )}

      {paymentMethods.length === 0 ? (
        <div className="text-center py-8">
          <svg className="w-12 h-12 mx-auto text-gray-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M3 10h18M7 15h1m4 0h1m-7 4h12a3 3 0 003-3V8a3 3 0 00-3-3H6a3 3 0 00-3 3v8a3 3 0 003 3z" />
          </svg>
          <p className="mt-2 text-gray-600">No payment methods on file</p>
          <p className="mt-1 text-sm text-gray-500">
            Add a payment method when you make your first purchase
          </p>
        </div>
      ) : (
        <div className="space-y-3">
          {paymentMethods.map((method) => (
            <div
              key={method.id}
              className={`border rounded-lg p-4 ${
                method.isDefault ? 'border-primary-500 bg-primary-50' : 'border-gray-200'
              }`}
            >
              <div className="flex items-start justify-between">
                <div className="flex items-center space-x-3">
                  <span className="text-2xl">{cardBrandIcons[method.brand.toLowerCase()] || '💳'}</span>
                  <div>
                    <p className="font-medium text-gray-900">
                      {method.brand} •••• {method.last4}
                    </p>
                    <p className="text-sm text-gray-600">
                      Expires {method.expiryMonth}/{method.expiryYear}
                    </p>
                    {method.isDefault && (
                      <span className="inline-flex items-center px-2 py-0.5 rounded text-xs font-medium bg-primary-100 text-primary-800 mt-1">
                        Default
                      </span>
                    )}
                  </div>
                </div>

                <div className="flex items-center space-x-2">
                  {!method.isDefault && (
                    <button
                      onClick={() => handleSetDefault(method.id)}
                      disabled={loading === method.id}
                      className="text-sm text-primary-600 hover:text-primary-700 font-medium"
                    >
                      {loading === method.id ? 'Setting...' : 'Set Default'}
                    </button>
                  )}
                  <button
                    onClick={() => handleRemove(method.id)}
                    disabled={loading === method.id || (method.isDefault && paymentMethods.length === 1)}
                    className="text-sm text-red-600 hover:text-red-700 font-medium disabled:opacity-50 disabled:cursor-not-allowed"
                  >
                    {loading === method.id ? 'Removing...' : 'Remove'}
                  </button>
                </div>
              </div>
            </div>
          ))}
        </div>
      )}

      <div className="mt-4 pt-4 border-t border-gray-200">
        <p className="text-xs text-gray-500">
          Payment methods are saved securely via Stripe Link when you make a purchase.
          You can manage additional settings in your{' '}
          <a href="https://link.stripe.com/" target="_blank" rel="noopener noreferrer" className="text-primary-600 hover:text-primary-700">
            Stripe Link account
          </a>.
        </p>
      </div>
    </div>
  );
}