import { VercelRequest, VercelResponse } from '@vercel/node';
import { CreditsBillingService } from '../../lib/creditsBillingService';
import { corsMiddleware, authenticateUser } from '../../lib/middleware';

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

    const {
      amountDollars = 5, // Default $5
      useQuickPurchase = false,
      origin // Return URL origin (e.g. https://billing.linefox.ai)
    } = req.body;

    // Validate amount (must be in $5 increments)
    if (amountDollars % 5 !== 0 || amountDollars < 5 || amountDollars > 100) {
      return res.status(400).json({
        error: 'Invalid amount. Must be $5, $10, $15, etc. up to $100'
      });
    }

    // Ensure Stripe customer exists
    await CreditsBillingService.ensureStripeCustomer(
      user.uid,
      user.email || `${user.uid}@heelix.app`,
      user.name
    );

    if (useQuickPurchase) {
      // Quick purchase with saved payment method
      try {
        await CreditsBillingService.quickPurchaseCredits(user.uid, amountDollars);
        
        const status = await CreditsBillingService.getBillingStatus(user.uid);
        
        return res.status(200).json({
          success: true,
          message: `Successfully purchased $${amountDollars} in credits`,
          billingStatus: status,
        });
      } catch (error: any) {
        if (error.message.includes('No payment method')) {
          // Fall back to checkout session
          const checkoutUrl = await CreditsBillingService.createCreditPurchaseSession(
            user.uid,
            amountDollars,
            origin
          );
          
          return res.status(200).json({
            success: false,
            requiresCheckout: true,
            checkoutUrl,
          });
        }
        throw error;
      }
    } else {
      // Create Stripe checkout session with Link
      const checkoutUrl = await CreditsBillingService.createCreditPurchaseSession(
        user.uid,
        amountDollars,
        origin
      );

      return res.status(200).json({
        success: true,
        checkoutUrl,
      });
    }
  } catch (error) {
    console.error('Error purchasing credits:', error);
    return res.status(500).json({ 
      error: 'Failed to process credit purchase',
      message: process.env.NODE_ENV === 'development' ? error : undefined
    });
  }
}