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

    const { enabled } = req.body;
    
    if (typeof enabled !== 'boolean') {
      return res.status(400).json({ 
        error: 'Invalid request. "enabled" must be a boolean' 
      });
    }

    // Check if user has payment method before enabling auto-recharge
    if (enabled) {
      const billing = await CreditsBillingService.getUserBilling(user.uid);
      if (!billing.hasPaymentMethod) {
        return res.status(400).json({ 
          error: 'Cannot enable auto-recharge without a payment method' 
        });
      }
    }

    // Toggle auto-recharge
    await CreditsBillingService.toggleAutoRecharge(user.uid, enabled);
    
    // Get updated billing status
    const status = await CreditsBillingService.getBillingStatus(user.uid);

    return res.status(200).json({
      success: true,
      message: `Auto-recharge ${enabled ? 'enabled' : 'disabled'}`,
      billingStatus: status,
    });
  } catch (error) {
    console.error('Error toggling auto-recharge:', error);
    return res.status(500).json({ 
      error: 'Failed to toggle auto-recharge',
      message: process.env.NODE_ENV === 'development' ? error : undefined
    });
  }
}