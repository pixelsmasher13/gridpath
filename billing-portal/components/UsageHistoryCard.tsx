import { CreditPurchase } from '@/lib/api';

interface Props {
  purchases: CreditPurchase[];
}

export function UsageHistoryCard({ purchases }: Props) {
  const formatDate = (dateString: string) => {
    const date = new Date(dateString);
    return date.toLocaleDateString('en-US', {
      month: 'short',
      day: 'numeric',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
    });
  };

  const getStatusColor = (status: string) => {
    switch (status) {
      case 'succeeded':
        return 'text-green-600 bg-green-50';
      case 'pending':
        return 'text-amber-600 bg-amber-50';
      case 'failed':
        return 'text-red-600 bg-red-50';
      default:
        return 'text-gray-600 bg-gray-50';
    }
  };

  return (
    <div className="card">
      <div className="flex items-center justify-between mb-4">
        <h2 className="text-lg font-semibold text-gray-900">Recent Purchases</h2>
        <a href="/invoices" className="text-sm text-primary-600 hover:text-primary-700 font-medium">
          View All
        </a>
      </div>

      {purchases.length === 0 ? (
        <div className="text-center py-8">
          <svg className="w-12 h-12 mx-auto text-gray-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2" />
          </svg>
          <p className="mt-2 text-gray-600">No purchases yet</p>
          <p className="mt-1 text-sm text-gray-500">
            Your credit purchases will appear here
          </p>
        </div>
      ) : (
        <div className="space-y-3">
          {purchases.slice(0, 5).map((purchase) => (
            <div
              key={purchase.purchaseId}
              className="border border-gray-200 rounded-lg p-4 hover:bg-gray-50 transition-colors"
            >
              <div className="flex items-start justify-between">
                <div className="flex-1">
                  <div className="flex items-center space-x-2">
                    <p className="font-medium text-gray-900">
                      ${purchase.amountDollars}.00
                    </p>
                    <span className={`inline-flex items-center px-2 py-0.5 rounded text-xs font-medium ${getStatusColor(purchase.status)}`}>
                      {purchase.status}
                    </span>
                  </div>
                  <p className="text-sm text-gray-600 mt-1">
                    {purchase.creditsAdded} minutes added
                  </p>
                  <p className="text-xs text-gray-500 mt-1">
                    {formatDate(purchase.createdAt)}
                  </p>
                </div>

                <div className="text-right">
                  <p className="text-xs text-gray-500">
                    {purchase.paymentMethod}
                  </p>
                  {purchase.status === 'succeeded' && (
                    <a
                      href={`/invoices/${purchase.purchaseId}`}
                      className="text-xs text-primary-600 hover:text-primary-700 mt-1 inline-block"
                    >
                      View Invoice
                    </a>
                  )}
                </div>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}