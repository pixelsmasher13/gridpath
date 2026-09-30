import type { VercelRequest, VercelResponse } from '@vercel/node';

export default async function handler(
  req: VercelRequest,
  res: VercelResponse
) {
  res.status(200).json({ 
    message: 'Billing endpoints are deployed!',
    endpoints: [
      '/api/billing/credits-status',
      '/api/billing/buy-credits',
      '/api/usage/end-credits',
      '/api/usage/history'
    ],
    environment: {
      hasJwtSecret: !!process.env.JWT_SECRET,
      hasDatabaseUrl: !!process.env.DATABASE_URL,
      hasStripeKey: !!process.env.STRIPE_SECRET_KEY
    }
  });
}