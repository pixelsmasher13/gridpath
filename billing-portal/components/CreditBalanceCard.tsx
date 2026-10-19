import { BillingStatus } from '@/lib/api';

interface Props {
  status: BillingStatus;
}

export function CreditBalanceCard({ status }: Props) {
  const totalMinutes = status.estimatedMinutesRemaining;
  const totalHours = (totalMinutes / 60).toFixed(1);

  return (
    <div className="card">
      {/* Header with status */}
      <div className="flex items-center justify-between mb-3 sm:mb-4">
        <h2 className="text-base sm:text-lg font-semibold text-gray-900">Balance</h2>
        {status.canContinue ? (
          <span className="flex items-center space-x-1 text-green-600 text-xs sm:text-sm">
            <svg className="w-4 h-4" fill="currentColor" viewBox="0 0 20 20">
              <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd" />
            </svg>
            <span className="font-medium">Active</span>
          </span>
        ) : (
          <span className="flex items-center space-x-1 text-red-600 text-xs sm:text-sm">
            <svg className="w-4 h-4" fill="currentColor" viewBox="0 0 20 20">
              <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zM8.707 7.293a1 1 0 00-1.414 1.414L8.586 10l-1.293 1.293a1 1 0 101.414 1.414L10 11.414l1.293 1.293a1 1 0 001.414-1.414L11.414 10l1.293-1.293a1 1 0 00-1.414-1.414L10 8.586 8.707 7.293z" clipRule="evenodd" />
            </svg>
            <span className="font-medium">Add credits</span>
          </span>
        )}
      </div>

      {/* Main Balance */}
      <div className="bg-primary-50 rounded-lg p-3 sm:p-4 mb-4">
        <div className="flex items-baseline justify-center space-x-2">
          <p className="text-xl sm:text-2xl font-semibold text-primary-900">
            {totalMinutes}
          </p>
          <p className="text-sm text-primary-700">minutes</p>
          <p className="text-xs text-primary-600">({totalHours}h)</p>
        </div>
      </div>

      {/* Breakdown */}
      <div className="space-y-2 sm:space-y-3">
        {/* Free Minutes - only show if remaining > 0 */}
        {status.freeMinutesRemaining > 0 && (
          <div className="flex justify-between items-center py-2 border-b border-gray-100">
            <p className="text-sm text-gray-700">Free minutes</p>
            <p className="text-base sm:text-lg font-semibold text-gray-900">
              {status.freeMinutesRemaining}
            </p>
          </div>
        )}

        {/* Purchased Credits */}
        <div className="flex justify-between items-center py-2">
          <div>
            <p className="text-sm text-gray-700">Purchased</p>
            <p className="text-xs text-gray-500">
              ${status.creditBalance.dollars.toFixed(2)} credit
            </p>
          </div>
          <p className="text-base sm:text-lg font-semibold text-gray-900">
            {status.creditBalance.minutes}
          </p>
        </div>
      </div>
    </div>
  );
}
