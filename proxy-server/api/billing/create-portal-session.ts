import { VercelRequest, VercelResponse } from '@vercel/node';
import Stripe from 'stripe';
import { corsMiddleware, authenticateUser } from '../../lib/middleware';
import { CreditsBillingService } from '../../lib/creditsBillingService';

const stripe = new Stripe(process.env.STRIPE_SECRET_KEY!, {
  apiVersion: '2025-07-30.basil',
});

export default async function handler(
  req: VercelRequest,
  res: VercelResponse
) {
  // Apply CORS
  if (!corsMiddleware(req, res)) {
    return;
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    // Authenticate user (supports both Firebase tokens and JWT)
    const user = await authenticateUser(req, res);
    if (!user) {
      return; // authenticateUser already sent error response
    }

    // Get or create Stripe customer
    const customerId = await CreditsBillingService.ensureStripeCustomer(user.uid, user.email);

    // Create a portal session for the customer
    const session = await stripe.billingPortal.sessions.create({
      customer: customerId,
      return_url: `${req.headers.origin || 'https://billing.heelix.com'}/dashboard`,
    });

    return res.status(200).json({
      url: session.url,
    });
  } catch (error) {
    console.error('Error creating portal session:', error);
    return res.status(500).json({ 
      error: 'Failed to create portal session',
      message: process.env.NODE_ENV === 'development' ? error : undefined
    });
  }
}