import { useEffect, useState } from 'react';
import { useRouter } from 'next/router';
import Head from 'next/head';
import useSWR from 'swr';
import { Layout } from '@/components/Layout';

interface UsageSession {
  sessionId: string;
  startedAt: string;
  endedAt: string | null;
  durationMinutes: number;
  billedMinutes: number;
  freeMinutesUsed: number;
  creditMinutesUsed: number;
  creditsDeducted: number;
  status: 'active' | 'completed';
  taskName?: string;
}

interface UsageStats {
  todayMinutes: number;
  weekMinutes: number;
  monthMinutes: number;
  totalMinutes: number;
  sessions: UsageSession[];
}

const API_URL = process.env.NEXT_PUBLIC_API_URL || 'https://bitworks-topaz.vercel.app';

const fetcher = async (): Promise<UsageStats> => {
  const token = localStorage.getItem('auth-token');
  if (!token) throw new Error('No auth token');
  
  const response = await fetch(`${API_URL}/api/usage/history`, {
    headers: {
      'Authorization': `Bearer ${token}`,
    },
  });

  if (!response.ok) {
    throw new Error('Failed to fetch usage history');
  }

  return response.json();
};

export default function UsagePage() {
  const router = useRouter();
  const { data: usage, error, mutate } = useSWR<UsageStats>('usage-history', fetcher, {
    refreshInterval: 60000, // Refresh every minute
  });

  useEffect(() => {
    const token = localStorage.getItem('auth-token');
    if (!token) {
      router.push('/auth');
    }
  }, [router]);

  const formatDate = (dateString: string) => {
    const date = new Date(dateString);
    return date.toLocaleDateString('en-US', {
      month: 'short',
      day: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
    });
  };

  const formatDuration = (minutes: number) => {
    if (minutes < 60) {
      return `${minutes}m`;
    }
    const hours = Math.floor(minutes / 60);
    const mins = minutes % 60;
    return `${hours}h ${mins}m`;
  };

  const getSessionColor = (session: UsageSession) => {
    if (session.status === 'active') return 'text-green-600 bg-green-50';
    if (session.creditMinutesUsed > 0) return 'text-blue-600 bg-blue-50';
    return 'text-gray-600 bg-gray-50';
  };

  if (error) {
    return (
      <Layout>
        <div className="max-w-7xl mx-auto px-4 py-4 sm:py-6">
          <div className="bg-red-50 border border-red-200 rounded-lg p-4 sm:p-6">
            <h2 className="text-base sm:text-lg font-semibold text-red-900 mb-2">Error loading usage data</h2>
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

  if (!usage) {
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
        <title>Usage History - Heelix</title>
      </Head>

      <div className="max-w-7xl mx-auto px-4 py-4 sm:py-6 lg:px-8">
        {/* Header */}
        <div className="mb-4 sm:mb-6">
          <h1 className="text-xl sm:text-2xl font-bold text-gray-900">Usage History</h1>
          <p className="mt-1 text-sm text-gray-500">Track your agent usage</p>
        </div>

        {/* Stats Cards - 2 columns on mobile, 4 on desktop */}
        <div className="grid grid-cols-2 md:grid-cols-4 gap-2 sm:gap-4 mb-4 sm:mb-6">
          <div className="card p-3 sm:p-4">
            <div className="text-xs sm:text-sm text-gray-600">Today</div>
            <div className="mt-0.5 sm:mt-1 text-lg sm:text-2xl font-semibold text-gray-900">
              {formatDuration(usage.todayMinutes)}
            </div>
          </div>
          <div className="card p-3 sm:p-4">
            <div className="text-xs sm:text-sm text-gray-600">This Week</div>
            <div className="mt-0.5 sm:mt-1 text-lg sm:text-2xl font-semibold text-gray-900">
              {formatDuration(usage.weekMinutes)}
            </div>
          </div>
          <div className="card p-3 sm:p-4">
            <div className="text-xs sm:text-sm text-gray-600">This Month</div>
            <div className="mt-0.5 sm:mt-1 text-lg sm:text-2xl font-semibold text-gray-900">
              {formatDuration(usage.monthMinutes)}
            </div>
          </div>
          <div className="card p-3 sm:p-4">
            <div className="text-xs sm:text-sm text-gray-600">All Time</div>
            <div className="mt-0.5 sm:mt-1 text-lg sm:text-2xl font-semibold text-gray-900">
              {formatDuration(usage.totalMinutes)}
            </div>
          </div>
        </div>

        {/* Sessions */}
        <div className="card">
          <div className="flex items-center justify-between mb-3 sm:mb-4">
            <h2 className="text-base sm:text-lg font-semibold text-gray-900">Recent Sessions</h2>
            <button
              onClick={() => mutate()}
              className="text-xs sm:text-sm text-primary-600 hover:text-primary-700 font-medium"
            >
              Refresh
            </button>
          </div>

          {usage.sessions.length === 0 ? (
            <div className="text-center py-6 sm:py-8">
              <svg className="w-10 h-10 sm:w-12 sm:h-12 mx-auto text-gray-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 8v4l3 3m6-3a9 9 0 11-18 0 9 9 0 0118 0z" />
              </svg>
              <p className="mt-2 text-sm sm:text-base text-gray-600">No usage sessions yet</p>
              <p className="mt-1 text-xs sm:text-sm text-gray-500">
                Your task sessions will appear here
              </p>
            </div>
          ) : (
            <>
              {/* Mobile: Card-based list */}
              <div className="sm:hidden space-y-3">
                {usage.sessions.map((session) => (
                  <div key={session.sessionId} className="bg-gray-50 rounded-lg p-3">
                    <div className="flex justify-between items-start mb-2">
                      <div>
                        <div className="text-sm font-medium text-gray-900">
                          {session.taskName || 'Task'}
                        </div>
                        <div className="text-xs text-gray-500">
                          {formatDate(session.startedAt)}
                        </div>
                      </div>
                      <span className={`inline-flex items-center px-2 py-0.5 rounded text-xs font-medium ${getSessionColor(session)}`}>
                        {session.status === 'active' ? 'Active' : 'Done'}
                      </span>
                    </div>
                    <div className="grid grid-cols-3 gap-2 text-center">
                      <div>
                        <div className="text-xs text-gray-500">Duration</div>
                        <div className="text-sm font-medium text-gray-900">{formatDuration(session.durationMinutes)}</div>
                      </div>
                      <div>
                        <div className="text-xs text-gray-500">Billed</div>
                        <div className="text-sm font-medium text-gray-900">{session.billedMinutes}m</div>
                      </div>
                      <div>
                        <div className="text-xs text-gray-500">Cost</div>
                        <div className="text-sm font-medium text-gray-900">${session.creditsDeducted.toFixed(2)}</div>
                      </div>
                    </div>
                  </div>
                ))}
              </div>

              {/* Desktop: Table */}
              <div className="hidden sm:block overflow-x-auto">
                <table className="min-w-full">
                  <thead className="bg-gray-50">
                    <tr>
                      <th className="px-4 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                        Session
                      </th>
                      <th className="px-4 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                        Started
                      </th>
                      <th className="px-4 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                        Duration
                      </th>
                      <th className="px-4 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                        Billed
                      </th>
                      <th className="px-4 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                        Credits Used
                      </th>
                      <th className="px-4 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                        Status
                      </th>
                    </tr>
                  </thead>
                  <tbody className="bg-white divide-y divide-gray-200">
                    {usage.sessions.map((session) => (
                      <tr key={session.sessionId} className="hover:bg-gray-50">
                        <td className="px-4 py-4 whitespace-nowrap">
                          <div>
                            <div className="text-sm font-medium text-gray-900">
                              {session.taskName || 'Task'}
                            </div>
                            <div className="text-xs text-gray-500">
                              {session.sessionId.slice(0, 8)}...
                            </div>
                          </div>
                        </td>
                        <td className="px-4 py-4 whitespace-nowrap text-sm text-gray-900">
                          {formatDate(session.startedAt)}
                        </td>
                        <td className="px-4 py-4 whitespace-nowrap">
                          <div className="text-sm text-gray-900">
                            {formatDuration(session.durationMinutes)}
                          </div>
                        </td>
                        <td className="px-4 py-4 whitespace-nowrap">
                          <div className="text-sm text-gray-900">
                            {session.billedMinutes}m
                          </div>
                          <div className="text-xs text-gray-500">
                            {session.freeMinutesUsed > 0 && `${session.freeMinutesUsed}m free`}
                            {session.freeMinutesUsed > 0 && session.creditMinutesUsed > 0 && ', '}
                            {session.creditMinutesUsed > 0 && `${session.creditMinutesUsed}m paid`}
                          </div>
                        </td>
                        <td className="px-4 py-4 whitespace-nowrap text-sm text-gray-900">
                          ${session.creditsDeducted.toFixed(2)}
                        </td>
                        <td className="px-4 py-4 whitespace-nowrap">
                          <span className={`inline-flex items-center px-2 py-0.5 rounded text-xs font-medium ${getSessionColor(session)}`}>
                            {session.status === 'active' ? 'Active' : 'Completed'}
                          </span>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </>
          )}
        </div>
      </div>
    </Layout>
  );
}