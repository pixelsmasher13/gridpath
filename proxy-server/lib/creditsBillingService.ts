import { query } from './db';
import Stripe from 'stripe';
import { sendWelcomeEmail } from './emailService';

const stripe = new Stripe(process.env.STRIPE_SECRET_KEY!, {
  apiVersion: '2025-07-30.basil',
});

// Constants
const HOURLY_RATE_DOLLARS = 5; // $5/hour
const CENTS_PER_MINUTE = 9; // ~$0.083/minute (rounded from 8.33)
const FREE_MINUTES_PER_MONTH = 60;
const STANDARD_CREDIT_PURCHASE = 5; // $5
const AUTO_RECHARGE_THRESHOLD = 1; // $1
const AUTO_RECHARGE_AMOUNT = 5; // $5

export interface UserBillingCredits {
  userId: string;
  stripeCustomerId?: string;
  hasPaymentMethod: boolean;
  defaultPaymentMethodId?: string;
  freeMinutesRemaining: number;
  freeMinutesResetDate: Date;
  creditBalanceDollars: number;
  autoRechargeEnabled: boolean;
  autoRechargeThresholdDollars: number;
  autoRechargeAmountDollars: number;
  totalMinutesUsed: number;
  billingStatus: 'active' | 'paused' | 'payment_failed';
  canContinue: boolean;
}

export interface UsageSessionCredits {
  sessionId: string;
  minutesUsed: number;
  freeMinutesUsed: number;
  creditMinutesUsed: number;
  creditsDeductedDollars: number;
  remainingBalanceDollars: number;
  needsRecharge: boolean;
  canContinue: boolean;
}

export interface CreditPurchase {
  purchaseId: number;
  amountDollars: number;
  creditsAddedDollars: number;
  purchaseType: 'manual' | 'auto_recharge';
  status: string;
  createdAt: Date;
}

export interface PaymentMethod {
  id: string;
  brand: string;
  last4: string;
  expiryMonth: number;
  expiryYear: number;
  isDefault: boolean;
}

export class CreditsBillingService {
  // Get or create user billing record
  // Pass userEmail to send welcome email on first creation
  static async getUserBilling(userId: string, userEmail?: string): Promise<UserBillingCredits> {
    // Try to insert new user - RETURNING tells us if a row was actually created
    const ensureText = `
      INSERT INTO user_billing_credits (user_id)
      VALUES ($1)
      ON CONFLICT (user_id) DO NOTHING
      RETURNING user_id
    `;
    const insertResult = await query(ensureText, [userId]);
    const isNewUser = insertResult.rows.length > 0;

    // Send welcome email for new users (fire and forget - don't block on it)
    if (isNewUser && userEmail) {
      sendWelcomeEmail(userEmail).catch(err =>
        console.error('Failed to send welcome email:', err)
      );
    }

    // Note: Free minutes are a one-time allowance (60 total), not monthly reset

    // Get current state
    const text = `
      SELECT * FROM user_billing_credits WHERE user_id = $1
    `;
    const result = await query(text, [userId]);
    const row = result.rows[0];

    return {
      userId: row.user_id,
      stripeCustomerId: row.stripe_customer_id,
      hasPaymentMethod: row.has_payment_method,
      defaultPaymentMethodId: row.default_payment_method_id,
      freeMinutesRemaining: row.free_minutes_remaining,
      freeMinutesResetDate: row.free_minutes_reset_date,
      creditBalanceDollars: (row.credit_balance_cents || 0) / 100,
      autoRechargeEnabled: row.auto_recharge_enabled || false,
      autoRechargeThresholdDollars: (row.auto_recharge_threshold_cents || 100) / 100,
      autoRechargeAmountDollars: (row.auto_recharge_amount_cents || 500) / 100,
      totalMinutesUsed: row.total_minutes_used || 0,
      billingStatus: row.billing_status,
      canContinue: row.billing_status === 'active' &&
                   (row.free_minutes_remaining > 0 || row.credit_balance_cents > 0),
    };
  }

  // Create or get Stripe customer
  static async ensureStripeCustomer(
    userId: string,
    email: string,
    name?: string
  ): Promise<string> {
    const billing = await this.getUserBilling(userId);
    
    if (billing.stripeCustomerId) {
      return billing.stripeCustomerId;
    }

    const customer = await stripe.customers.create({
      email,
      name,
      metadata: { userId },
    });

    const updateText = `
      UPDATE user_billing_credits 
      SET stripe_customer_id = $2, updated_at = CURRENT_TIMESTAMP
      WHERE user_id = $1
    `;
    
    await query(updateText, [userId, customer.id]);
    
    return customer.id;
  }

  // Create checkout session for credit purchase with Stripe Link
  static async createCreditPurchaseSession(
    userId: string,
    amountDollars: number = STANDARD_CREDIT_PURCHASE,
    origin?: string
  ): Promise<string> {
    const billing = await this.getUserBilling(userId);

    if (!billing.stripeCustomerId) {
      throw new Error('Stripe customer not found');
    }

    // Use provided origin or fall back to production domain
    const baseUrl = origin || 'https://billing.linefox.ai';

    // Create checkout session with Stripe Link enabled
    const session = await stripe.checkout.sessions.create({
      customer: billing.stripeCustomerId,
      payment_method_types: ['card', 'link'],
      mode: 'payment',
      line_items: [{
        price_data: {
          currency: 'usd',
          product_data: {
            name: 'Linefox Credits',
            description: `${amountDollars * 12} minutes of usage`,
          },
          unit_amount: amountDollars * 100,
        },
        quantity: 1,
      }],
      success_url: `${baseUrl}/payment/success?amount=${amountDollars}`,
      cancel_url: `${baseUrl}/payment/cancel`,
      metadata: {
        userId,
        type: 'credit_purchase',
        amountDollars: amountDollars.toString(),
      },
      payment_intent_data: {
        metadata: {
          userId,
          type: 'credit_purchase',
        },
      },
      // Enable saving payment method for future use
      customer_update: {
        address: 'auto',
      },
    });

    return session.url!;
  }

  // Quick purchase using saved payment method
  static async quickPurchaseCredits(
    userId: string,
    amountDollars: number = STANDARD_CREDIT_PURCHASE
  ): Promise<void> {
    const billing = await this.getUserBilling(userId);
    
    if (!billing.stripeCustomerId || !billing.defaultPaymentMethodId) {
      throw new Error('No payment method on file');
    }

    // Create payment intent
    const paymentIntent = await stripe.paymentIntents.create({
      amount: amountDollars * 100,
      currency: 'usd',
      customer: billing.stripeCustomerId,
      payment_method: billing.defaultPaymentMethodId,
      off_session: true,
      confirm: true,
      description: `Linefox Credits - ${amountDollars * 12} minutes`,
      metadata: {
        userId,
        type: 'credit_purchase',
      },
    });

    // Add credits if successful
    if (paymentIntent.status === 'succeeded') {
      await query('SELECT add_credits($1, $2, $3, $4)', [
        userId,
        amountDollars * 100,
        paymentIntent.id,
        'manual',
      ]);
    } else {
      throw new Error('Payment failed');
    }
  }

  // Process auto-recharge
  static async processAutoRecharge(userId: string): Promise<void> {
    const billing = await this.getUserBilling(userId);
    
    if (!billing.autoRechargeEnabled || !billing.hasPaymentMethod) {
      return;
    }

    if (billing.creditBalanceDollars >= billing.autoRechargeThresholdDollars) {
      return; // No recharge needed
    }

    try {
      await this.quickPurchaseCredits(userId, billing.autoRechargeAmountDollars);
      
      // Record as auto-recharge
      await query(
        `UPDATE credit_purchases 
         SET purchase_type = 'auto_recharge' 
         WHERE user_id = $1 
         ORDER BY created_at DESC 
         LIMIT 1`,
        [userId]
      );
    } catch (error) {
      console.error('Auto-recharge failed:', error);
      // Update billing status if payment failed
      await query(
        'UPDATE user_billing_credits SET billing_status = $2 WHERE user_id = $1',
        [userId, 'payment_failed']
      );
    }
  }

  // Toggle auto-recharge
  static async toggleAutoRecharge(
    userId: string,
    enabled: boolean
  ): Promise<void> {
    await query('SELECT toggle_auto_recharge($1, $2, $3)', [
      userId,
      enabled,
      'user',
    ]);
  }

  // Process usage for a session
  static async processUsageSession(
    userId: string,
    sessionId: string,
    minutesUsed: number
  ): Promise<UsageSessionCredits> {
    const text = `
      SELECT * FROM process_usage_session_credits($1, $2, $3)
    `;
    
    const result = await query(text, [userId, sessionId, minutesUsed]);
    const data = result.rows[0];
    
    // Check if auto-recharge needed
    if (data.needs_recharge) {
      await this.processAutoRecharge(userId);
    }
    
    return {
      sessionId,
      minutesUsed,
      freeMinutesUsed: data.free_minutes_used,
      creditMinutesUsed: data.credit_minutes_used,
      creditsDeductedDollars: data.credits_deducted_cents / 100,
      remainingBalanceDollars: data.remaining_balance_cents / 100,
      needsRecharge: data.needs_recharge,
      canContinue: data.can_continue,
    };
  }

  // Get user's payment methods
  static async getPaymentMethods(userId: string): Promise<PaymentMethod[]> {
    const billing = await this.getUserBilling(userId);
    
    if (!billing.stripeCustomerId) {
      return [];
    }

    const paymentMethods = await stripe.paymentMethods.list({
      customer: billing.stripeCustomerId,
      type: 'card',
    });

    return paymentMethods.data.map(pm => ({
      id: pm.id,
      brand: pm.card!.brand,
      last4: pm.card!.last4,
      expiryMonth: pm.card!.exp_month,
      expiryYear: pm.card!.exp_year,
      isDefault: pm.id === billing.defaultPaymentMethodId,
    }));
  }

  // Save payment method after setup
  static async savePaymentMethod(
    userId: string,
    paymentMethodId: string,
    setAsDefault: boolean = true
  ): Promise<void> {
    const billing = await this.getUserBilling(userId);
    
    if (!billing.stripeCustomerId) {
      throw new Error('Stripe customer not found');
    }

    // Attach payment method to customer
    await stripe.paymentMethods.attach(paymentMethodId, {
      customer: billing.stripeCustomerId,
    });

    if (setAsDefault) {
      // Set as default payment method
      await stripe.customers.update(billing.stripeCustomerId, {
        invoice_settings: {
          default_payment_method: paymentMethodId,
        },
      });

      // Update database
      const updateText = `
        UPDATE user_billing_credits 
        SET 
          has_payment_method = true,
          default_payment_method_id = $2,
          updated_at = CURRENT_TIMESTAMP
        WHERE user_id = $1
      `;
      
      await query(updateText, [userId, paymentMethodId]);
    }
  }

  // Get billing status for UI
  // Pass userEmail to send welcome email on first user creation
  static async getBillingStatus(userId: string, userEmail?: string): Promise<{
    freeMinutesRemaining: number;
    freeMinutesResetDate: Date;
    creditBalance: {
      dollars: number;
      minutes: number;
    };
    autoRecharge: {
      enabled: boolean;
      thresholdDollars: number;
      amountDollars: number;
    };
    totalMinutesUsed: number;
    hasPaymentMethod: boolean;
    needsCredits: boolean;
    billingStatus: string;
    canContinue: boolean;
    estimatedMinutesRemaining: number;
  }> {
    const billing = await this.getUserBilling(userId, userEmail);
    
    const creditMinutes = Math.floor((billing.creditBalanceDollars * 60) / HOURLY_RATE_DOLLARS);
    const totalMinutesAvailable = billing.freeMinutesRemaining + creditMinutes;
    const needsCredits = billing.freeMinutesRemaining === 0 && billing.creditBalanceDollars < 1;
    
    return {
      freeMinutesRemaining: billing.freeMinutesRemaining,
      freeMinutesResetDate: billing.freeMinutesResetDate,
      creditBalance: {
        dollars: billing.creditBalanceDollars,
        minutes: creditMinutes,
      },
      autoRecharge: {
        enabled: billing.autoRechargeEnabled,
        thresholdDollars: billing.autoRechargeThresholdDollars,
        amountDollars: billing.autoRechargeAmountDollars,
      },
      totalMinutesUsed: billing.totalMinutesUsed,
      hasPaymentMethod: billing.hasPaymentMethod,
      needsCredits,
      billingStatus: billing.billingStatus,
      canContinue: billing.canContinue,
      estimatedMinutesRemaining: totalMinutesAvailable,
    };
  }

  // Check if user can start new session
  static async canStartSession(userId: string): Promise<{
    canStart: boolean;
    reason?: string;
  }> {
    const billing = await this.getUserBilling(userId);
    
    if (billing.billingStatus === 'payment_failed') {
      return { canStart: false, reason: 'Payment method failed. Please update your card.' };
    }
    
    if (billing.billingStatus === 'paused') {
      return { canStart: false, reason: 'Account paused. Please contact support.' };
    }
    
    if (billing.freeMinutesRemaining === 0 && billing.creditBalanceDollars === 0) {
      return { canStart: false, reason: 'No credits available. Please purchase credits to continue.' };
    }
    
    return { canStart: true };
  }

  // Get recent purchases
  static async getRecentPurchases(userId: string, limit = 10): Promise<CreditPurchase[]> {
    const text = `
      SELECT * FROM credit_purchases
      WHERE user_id = $1
      ORDER BY created_at DESC
      LIMIT $2
    `;
    
    const result = await query(text, [userId, limit]);
    
    return result.rows.map(row => ({
      purchaseId: row.purchase_id,
      amountDollars: row.amount_cents / 100,
      creditsAddedDollars: row.credits_added_cents / 100,
      purchaseType: row.purchase_type,
      status: row.status,
      createdAt: row.created_at,
    }));
  }

  // Handle webhook for successful payment
  static async handleSuccessfulPayment(
    paymentIntentId: string,
    userId: string,
    amountCents: number
  ): Promise<void> {
    await query('SELECT add_credits($1, $2, $3, $4)', [
      userId,
      amountCents,
      paymentIntentId,
      'manual',
    ]);
  }
}