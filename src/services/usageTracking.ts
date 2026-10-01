import { invoke } from '@tauri-apps/api/core';
import { listen } from '@tauri-apps/api/event';
// billingPortal lived in the legacy automation flow — removed in the
// spreadsheet-workspace cleanup. usageTracking still ships for the
// authToken/session plumbing AuthProvider relies on, but its billing
// methods are stubbed (the spreadsheet agent doesn't gate on credits).

interface UsageSession {
  sessionId: string;
  automationId: number;
  startedAt: string;
  endedAt?: string;
  totalSeconds: number;
  billedMinutes: number;
  status: 'active' | 'completed' | 'failed';
}

interface UsageUpdate {
  sessionId: string;
  elapsedSeconds: number;
  billedMinutes: number;
}

interface QueuedUsageData {
  endpoint: string;
  data: any;
  timestamp: number;
  retries: number;
}

class UsageTrackingService {
  private baseUrl: string;
  private authToken: string | null = null;
  private currentSession: UsageSession | null = null;
  private updateInterval: NodeJS.Timeout | null = null;
  private offlineQueue: QueuedUsageData[] = [];
  private isOnline: boolean = navigator.onLine;

  constructor() {
    // Use the same proxy URL as LLM requests
    this.baseUrl = import.meta.env.VITE_PROXY_URL || 'https://gridpath-theta.vercel.app';
    
    // Listen for online/offline events
    window.addEventListener('online', () => {
      this.isOnline = true;
      this.processOfflineQueue();
    });
    
    window.addEventListener('offline', () => {
      this.isOnline = false;
    });

    // Load offline queue from localStorage
    this.loadOfflineQueue();
  }

  setAuthToken(token: string) {
    this.authToken = token;
  }

  private async makeRequest(endpoint: string, data: any): Promise<any> {
    if (!this.authToken) {
      throw new Error('No auth token available');
    }

    if (!this.isOnline) {
      this.addToOfflineQueue(endpoint, data);
      return { queued: true };
    }

    try {
      const response = await fetch(`${this.baseUrl}/api/usage${endpoint}`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${this.authToken}`
        },
        body: JSON.stringify(data)
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      return await response.json();
    } catch (error) {
      console.error('Usage tracking request failed:', error);
      // Add to offline queue if request fails
      this.addToOfflineQueue(endpoint, data);
      throw error;
    }
  }

  async checkCreditsBeforeStart(): Promise<{ canContinue: boolean; needsPayment: boolean }> {
    if (!this.authToken) {
      console.warn('No auth token set, cannot check credits');
      return { canContinue: true, needsPayment: false }; // Allow offline usage
    }

    // Spreadsheet agent doesn't gate on credits. Always allow.
    return { canContinue: true, needsPayment: false };
  }

  async startSession(automationId: number, automationName: string): Promise<string> {
    console.log('📊 Usage Tracking: Starting session for', automationName, 'with auth token:', this.authToken ? 'set' : 'not set');
    
    // Check credits before starting
    const { canContinue, needsPayment } = await this.checkCreditsBeforeStart();
    if (!canContinue) {
      console.warn('Cannot start automation: No credits available');
      throw new Error('Credits required. Please purchase credits to continue.');
    }
    
    const sessionId = `session_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
    const startedAt = new Date().toISOString();

    this.currentSession = {
      sessionId,
      automationId,
      startedAt,
      totalSeconds: 0,
      billedMinutes: 0,
      status: 'active'
    };

    // Send to server
    try {
      await this.makeRequest('/start', {
        sessionId,
        automationId,
        automationName,
        startedAt
      });
    } catch (error) {
      console.error('Failed to start usage session on server:', error);
    }

    // Start periodic updates
    this.startPeriodicUpdates();

    // Store session locally for recovery
    this.storeSessionLocally();

    return sessionId;
  }

  async updateSession(elapsedSeconds: number, billedMinutes: number) {
    if (!this.currentSession) return;

    this.currentSession.totalSeconds = elapsedSeconds;
    this.currentSession.billedMinutes = billedMinutes;

    try {
      await this.makeRequest('/update', {
        sessionId: this.currentSession.sessionId,
        elapsedSeconds,
        billedMinutes,
        timestamp: new Date().toISOString()
      });
    } catch (error) {
      console.error('Failed to update usage session:', error);
    }

    // Update local storage
    this.storeSessionLocally();
  }

  async endSession(status: 'completed' | 'failed' = 'completed') {
    console.log('📊 Usage Tracking: Ending session with status:', status, 'Current session:', this.currentSession?.sessionId);
    if (!this.currentSession) return;

    const endedAt = new Date().toISOString();
    this.currentSession.endedAt = endedAt;
    this.currentSession.status = status;

    // Stop periodic updates
    if (this.updateInterval) {
      clearInterval(this.updateInterval);
      this.updateInterval = null;
    }

    let billingResult = null;
    try {
      // Use the credit-based endpoint instead of the old /end endpoint
      const response = await this.makeRequest('/end-credits', {
        sessionId: this.currentSession.sessionId,
        endedAt,
        totalSeconds: this.currentSession.totalSeconds,
        billedMinutes: this.currentSession.billedMinutes,
        status
      });
      
      // The credit endpoint returns billing information
      if (response && response.billing) {
        billingResult = response.billing;
        console.log('💰 Billing update:', {
          freeMinutesUsed: billingResult.freeMinutesUsed,
          creditMinutesUsed: billingResult.creditMinutesUsed,
          creditsDeducted: billingResult.creditsDeducted,
          remainingCredits: billingResult.remainingCredits,
          canContinue: billingResult.canContinue
        });
        
        // Check if user needs to add credits
        if (!billingResult.canContinue) {
          const event = new CustomEvent('billing:credits-exhausted', {
            detail: { 
              message: 'Your credits have been exhausted. Please purchase more credits to continue.',
              remainingCredits: billingResult.remainingCredits,
              needsPayment: true 
            }
          });
          window.dispatchEvent(event);
        }
      }
    } catch (error) {
      console.error('Failed to end usage session:', error);
    }

    // Clear local storage
    this.clearSessionLocally();
    this.currentSession = null;
    
    return billingResult;
  }

  private startPeriodicUpdates() {
    // Send updates every 30 seconds
    this.updateInterval = setInterval(() => {
      if (this.currentSession && this.currentSession.status === 'active') {
        this.updateSession(
          this.currentSession.totalSeconds,
          this.currentSession.billedMinutes
        );
      }
    }, 30000);
  }

  private storeSessionLocally() {
    if (this.currentSession) {
      localStorage.setItem('heelix_current_session', JSON.stringify(this.currentSession));
    }
  }

  private clearSessionLocally() {
    localStorage.removeItem('heelix_current_session');
  }

  async recoverSession() {
    const storedSession = localStorage.getItem('heelix_current_session');
    if (storedSession) {
      try {
        const session = JSON.parse(storedSession) as UsageSession;
        // Check if session is stale (older than 1 hour)
        const startTime = new Date(session.startedAt).getTime();
        const now = Date.now();
        const hourInMs = 60 * 60 * 1000;

        if (now - startTime > hourInMs) {
          // Session is stale, end it
          this.currentSession = session;
          await this.endSession('failed');
        } else {
          // Session might still be active, recover it
          this.currentSession = session;
          this.startPeriodicUpdates();
        }
      } catch (error) {
        console.error('Failed to recover session:', error);
        this.clearSessionLocally();
      }
    }
  }

  private addToOfflineQueue(endpoint: string, data: any) {
    const queueItem: QueuedUsageData = {
      endpoint,
      data,
      timestamp: Date.now(),
      retries: 0
    };

    this.offlineQueue.push(queueItem);
    this.saveOfflineQueue();
  }

  private loadOfflineQueue() {
    const stored = localStorage.getItem('heelix_usage_queue');
    if (stored) {
      try {
        this.offlineQueue = JSON.parse(stored);
      } catch (error) {
        console.error('Failed to load offline queue:', error);
        this.offlineQueue = [];
      }
    }
  }

  private saveOfflineQueue() {
    localStorage.setItem('heelix_usage_queue', JSON.stringify(this.offlineQueue));
  }

  async processOfflineQueue() {
    if (!this.isOnline || this.offlineQueue.length === 0) return;

    const queue = [...this.offlineQueue];
    this.offlineQueue = [];

    for (const item of queue) {
      try {
        await this.makeRequest(item.endpoint, item.data);
      } catch (error) {
        item.retries++;
        if (item.retries < 3) {
          this.offlineQueue.push(item);
        } else {
          console.error('Failed to process queued item after 3 retries:', item);
        }
      }
    }

    this.saveOfflineQueue();
  }

  // Get usage summary from server
  async getUsageSummary(period: 'day' | 'week' | 'month' = 'month'): Promise<any> {
    if (!this.authToken) {
      throw new Error('No auth token available');
    }

    try {
      const response = await fetch(`${this.baseUrl}/api/usage/summary?period=${period}`, {
        headers: {
          'Authorization': `Bearer ${this.authToken}`
        }
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      return await response.json();
    } catch (error) {
      console.error('Failed to get usage summary:', error);
      throw error;
    }
  }
}

// Create singleton instance
export const usageTracking = new UsageTrackingService();

// Listen for Tauri usage events and forward to server
export function initializeUsageTracking() {
  // Listen for automation started
  listen('automation_started', async (event: any) => {
    console.log('📊 Usage Tracking: automation_started event received', event.payload);
    const { automationId, automationName } = event.payload;
    await usageTracking.startSession(automationId, automationName);
  });

  // Listen for usage updates
  listen<UsageUpdate>('usage_update', async (event) => {
    const { elapsedSeconds, billedMinutes } = event.payload;
    await usageTracking.updateSession(elapsedSeconds, billedMinutes);
  });

  // Listen for automation stopped
  listen('automation_stopped', async () => {
    console.log('📊 Usage Tracking: automation_stopped event received');
    await usageTracking.endSession('completed');
  });

  // Listen for automation failed
  listen('automation_failed', async () => {
    console.log('📊 Usage Tracking: automation_failed event received');
    await usageTracking.endSession('failed');
  });

  // Recover any existing session on startup
  usageTracking.recoverSession();

  // Process offline queue on startup
  usageTracking.processOfflineQueue();
}