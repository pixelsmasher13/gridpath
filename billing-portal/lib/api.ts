const API_URL = process.env.NEXT_PUBLIC_API_URL || 'https://bitworks-topaz.vercel.app';

export interface CreditBalance {
  dollars: number;
  minutes: number;
}

export interface AutoRechargeSettings {
  enabled: boolean;
  thresholdDollars: number;
  amountDollars: number;
}

export interface PaymentMethod {
  id: string;
  brand: string;
  last4: string;
  expiryMonth: number;
  expiryYear: number;
  isDefault: boolean;
}

export interface CreditPurchase {
  purchaseId: string;
  amountDollars: number;
  creditsAdded: number;
  paymentMethod: string;
  status: 'pending' | 'succeeded' | 'failed';
  createdAt: string;
}

export interface BillingStatus {
  freeMinutesRemaining: number;
  freeMinutesResetDate: string;
  creditBalance: CreditBalance;
  autoRecharge: AutoRechargeSettings;
  hasPaymentMethod: boolean;
  needsCredits: boolean;
  canContinue: boolean;
  paymentMethods: PaymentMethod[];
  recentPurchases: CreditPurchase[];
  estimatedMinutesRemaining: number;
}

class BillingAPI {
  private token: string | null = null;

  setToken(token: string) {
    this.token = token;
  }

  private getHeaders(): HeadersInit {
    if (!this.token) {
      throw new Error('No auth token set');
    }
    return {
      'Authorization': `Bearer ${this.token}`,
      'Content-Type': 'application/json',
    };
  }

  async getStatus(): Promise<BillingStatus> {
    const response = await fetch(`${API_URL}/api/billing/credits-status`, {
      headers: this.getHeaders(),
    });

    if (!response.ok) {
      throw new Error('Failed to fetch billing status');
    }

    return response.json();
  }

  async purchaseCredits(amount: number = 5, useQuickPurchase: boolean = false) {
    const response = await fetch(`${API_URL}/api/billing/buy-credits`, {
      method: 'POST',
      headers: this.getHeaders(),
      body: JSON.stringify({
        amountDollars: amount,
        useQuickPurchase,
        origin: typeof window !== 'undefined' ? window.location.origin : undefined,
      }),
    });

    if (!response.ok) {
      const error = await response.json().catch(() => ({ error: 'Failed to purchase credits' }));
      throw new Error(error.error || 'Failed to purchase credits');
    }

    return response.json();
  }

  async toggleAutoRecharge(enabled: boolean) {
    const response = await fetch(`${API_URL}/api/billing/toggle-auto-recharge`, {
      method: 'POST',
      headers: this.getHeaders(),
      body: JSON.stringify({ enabled }),
    });

    if (!response.ok) {
      throw new Error('Failed to toggle auto-recharge');
    }

    return response.json();
  }

  async updatePaymentMethod(paymentMethodId: string) {
    const response = await fetch(`${API_URL}/api/billing/update-payment-method`, {
      method: 'POST',
      headers: this.getHeaders(),
      body: JSON.stringify({ paymentMethodId }),
    });

    if (!response.ok) {
      throw new Error('Failed to update payment method');
    }

    return response.json();
  }

  async removePaymentMethod(paymentMethodId: string) {
    const response = await fetch(`${API_URL}/api/billing/remove-payment-method`, {
      method: 'POST',
      headers: this.getHeaders(),
      body: JSON.stringify({ paymentMethodId }),
    });

    if (!response.ok) {
      throw new Error('Failed to remove payment method');
    }

    return response.json();
  }
}

export const billingAPI = new BillingAPI();