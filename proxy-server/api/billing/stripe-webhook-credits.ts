import { VercelRequest, VercelResponse } from '@vercel/node';
import Stripe from 'stripe';
import { CreditsBillingService } from '../../lib/creditsBillingService';
import { buffer } from 'micro';

const stripe = new Stripe(process.env.STRIPE_SECRET_KEY!, {
  apiVersion: '2025-07-30.basil',
});

// Disable bodyParser to get raw body for webhook signature verification
export const config = {
  api: {
    bodyParser: false,
  },
};

export default async function handler(
  req: VercelRequest,
  res: VercelResponse
) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const buf = await buffer(req);
  const sig = req.headers['stripe-signature'] as string;

  if (!sig) {
    return res.status(400).json({ error: 'Missing stripe signature' });
  }

  let event: Stripe.Event;

  try {
    event = stripe.webhooks.constructEvent(
      buf,
      sig,
      process.env.STRIPE_WEBHOOK_SECRET!
    );
  } catch (err: any) {
    console.error('Webhook signature verification failed:', err.message);
    return res.status(400).json({ error: `Webhook Error: ${err.message}` });
  }

  try {
    switch (event.type) {
      case 'payment_intent.succeeded': {
        const paymentIntent = event.data.object as Stripe.PaymentIntent;
        
        // Check if this is a credit purchase
        if (paymentIntent.metadata.type === 'credit_purchase') {
          const userId = paymentIntent.metadata.userId;
          
          if (!userId) {
            console.error('No userId in payment intent metadata');
            return res.status(400).json({ error: 'Invalid payment metadata' });
          }

          // Add credits to user account
          await CreditsBillingService.handleSuccessfulPayment(
            paymentIntent.id,
            userId,
            paymentIntent.amount
          );
          
          console.log(`Credits added for user ${userId}: $${paymentIntent.amount / 100}`);
        }
        break;
      }

      case 'payment_intent.payment_failed': {
        const paymentIntent = event.data.object as Stripe.PaymentIntent;
        
        if (paymentIntent.metadata.type === 'credit_purchase') {
          const userId = paymentIntent.metadata.userId;
          
          console.error(`Payment failed for user ${userId}:`, paymentIntent.last_payment_error);
          
          // Could update user billing status if needed
          // await query(
          //   'UPDATE user_billing_credits SET billing_status = $2 WHERE user_id = $1',
          //   [userId, 'payment_failed']
          // );
        }
        break;
      }

      case 'checkout.session.completed': {
        const session = event.data.object as Stripe.Checkout.Session;
        
        // Handle successful checkout session
        if (session.metadata?.type === 'credit_purchase') {
          const userId = session.metadata.userId;
          const amountDollars = parseInt(session.metadata.amountDollars || '5');
          
          if (!userId) {
            console.error('No userId in checkout session metadata');
            return res.status(400).json({ error: 'Invalid session metadata' });
          }

          // Payment intent will be processed in payment_intent.succeeded
          console.log(`Checkout completed for user ${userId}: $${amountDollars}`);
        }
        break;
      }

      case 'setup_intent.succeeded': {
        const setupIntent = event.data.object as Stripe.SetupIntent;
        
        // Handle successful card setup
        if (setupIntent.metadata?.userId && setupIntent.payment_method) {
          await CreditsBillingService.savePaymentMethod(
            setupIntent.metadata.userId,
            setupIntent.payment_method as string
          );
          
          console.log(`Payment method saved for user ${setupIntent.metadata.userId}`);
        }
        break;
      }

      default:
        console.log(`Unhandled event type: ${event.type}`);
    }

    return res.status(200).json({ received: true });
  } catch (error) {
    console.error('Error processing webhook:', error);
    return res.status(500).json({ error: 'Webhook processing failed' });
  }
}