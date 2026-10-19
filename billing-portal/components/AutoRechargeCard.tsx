import { useState } from 'react';
import { BillingStatus, billingAPI } from '@/lib/api';

interface Props {
  status: BillingStatus;
  onToggle: () => void;
}

export function AutoRechargeCard({ status, onToggle }: Props) {
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const handleToggle = async () => {
    setLoading(true);
    setError(null);

    try {
      const token = localStorage.getItem('auth-token');
      if (!token) throw new Error('Not authenticated');
      
      billingAPI.setToken(token);
      await billingAPI.toggleAutoRecharge(!status.autoRecharge.enabled);
      onToggle();
    } catch (err: any) {
      setError(err.message || 'Failed to update auto-recharge');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="card">
      <div className="flex items-start justify-between">
        <div className="flex-1">
          <h2 className="text-lg font-semibold text-gray-900">Auto-Recharge</h2>
          <p className="mt-1 text-sm text-gray-600">
            Automatically add credits when your balance runs low
          </p>
        </div>
        
        <button
          onClick={handleToggle}
          disabled={loading}
          className={`relative inline-flex h-6 w-11 items-center rounded-full transition-colors ${
            status.autoRecharge.enabled ? 'bg-primary-600' : 'bg-gray-200'
          } ${loading ? 'opacity-50 cursor-not-allowed' : ''}`}
        >
          <span
            className={`inline-block h-4 w-4 transform rounded-full bg-white transition-transform ${
              status.autoRecharge.enabled ? 'translate-x-6' : 'translate-x-1'
            }`}
          />
        </button>
      </div>

      {error && (
        <div className="mt-4 p-3 bg-red-50 border border-red-200 rounded-lg">
          <p className="text-sm text-red-700">{error}</p>
        </div>
      )}

      <div className="mt-4 space-y-3">
        <div className="bg-gray-50 rounded-lg p-4">
          <div className="space-y-2">
            <div className="flex justify-between text-sm">
              <span className="text-gray-600">Status</span>
              <span className={`font-medium ${status.autoRecharge.enabled ? 'text-green-600' : 'text-gray-900'}`}>
                {status.autoRecharge.enabled ? 'Enabled' : 'Disabled'}
              </span>
            </div>
            <div className="flex justify-between text-sm">
              <span className="text-gray-600">Trigger threshold</span>
              <span className="font-medium">${status.autoRecharge.thresholdDollars}</span>
            </div>
            <div className="flex justify-between text-sm">
              <span className="text-gray-600">Recharge amount</span>
              <span className="font-medium">${status.autoRecharge.amountDollars}</span>
            </div>
          </div>
        </div>

        {status.autoRecharge.enabled && (
          <div className="flex items-start space-x-2">
            <svg className="w-5 h-5 text-primary-600 mt-0.5 flex-shrink-0" fill="currentColor" viewBox="0 0 20 20">
              <path fillRule="evenodd" d="M18 10a8 8 0 11-16 0 8 8 0 0116 0zm-7-4a1 1 0 11-2 0 1 1 0 012 0zM9 9a1 1 0 000 2v3a1 1 0 001 1h1a1 1 0 100-2v-3a1 1 0 00-1-1H9z" clipRule="evenodd" />
            </svg>
            <p className="text-sm text-gray-600">
              When your balance drops below ${status.autoRecharge.thresholdDollars}, 
              we'll automatically charge ${status.autoRecharge.amountDollars} to your default payment method.
            </p>
          </div>
        )}

        {!status.autoRecharge.enabled && (
          <div className="flex items-start space-x-2">
            <svg className="w-5 h-5 text-amber-500 mt-0.5 flex-shrink-0" fill="currentColor" viewBox="0 0 20 20">
              <path fillRule="evenodd" d="M8.257 3.099c.765-1.36 2.722-1.36 3.486 0l5.58 9.92c.75 1.334-.213 2.98-1.742 2.98H4.42c-1.53 0-2.493-1.646-1.743-2.98l5.58-9.92zM11 13a1 1 0 11-2 0 1 1 0 012 0zm-1-8a1 1 0 00-1 1v3a1 1 0 002 0V6a1 1 0 00-1-1z" clipRule="evenodd" />
            </svg>
            <p className="text-sm text-gray-600">
              Auto-recharge is disabled. You'll need to manually purchase credits when your balance runs out.
            </p>
          </div>
        )}
      </div>
    </div>
  );
}