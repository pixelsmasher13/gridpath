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

  if (req.method !== 'GET') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    // Authenticate user (supports both Firebase tokens and JWT)
    const user = await authenticateUser(req, res);
    if (!user) {
      return; // authenticateUser already sent error response
    }

    // Get billing status (pass email to send welcome email for new users)
    const status = await CreditsBillingService.getBillingStatus(user.uid, user.email);
    
    // Get payment methods
    const paymentMethods = await CreditsBillingService.getPaymentMethods(user.uid);

    // Get recent purchases
    const recentPurchases = await CreditsBillingService.getRecentPurchases(user.uid, 10);

    // Check if can start session
    const canStartSession = await CreditsBillingService.canStartSession(user.uid);

    return res.status(200).json({
      ...status,
      paymentMethods,
      recentPurchases,
      canStartNewSession: canStartSession.canStart,
      cantStartReason: canStartSession.reason,
    });
  } catch (error) {
    console.error('Error fetching billing status:', error);
    return res.status(500).json({ 
      error: 'Failed to fetch billing status',
      message: process.env.NODE_ENV === 'development' ? error : undefined
    });
  }
}