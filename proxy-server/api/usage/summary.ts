import { VercelRequest, VercelResponse } from '@vercel/node';
import { UsageService } from '../../lib/usageService';
import { verifyJWT } from '../../lib/jwt';
import { corsMiddleware } from '../../lib/middleware';

export default async function handler(
  req: VercelRequest,
  res: VercelResponse
) {
  // Apply CORS
  if (!corsMiddleware(req, res)) {
    return;
  }

  // Only allow GET
  if (req.method !== 'GET') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    // Verify JWT token using your existing auth
    const token = req.headers.authorization?.replace('Bearer ', '');
    if (!token) {
      return res.status(401).json({ error: 'No token provided' });
    }

    const user = verifyJWT(token);
    if (!user || !user.uid) {
      return res.status(401).json({ error: 'Invalid token' });
    }

    // Get period from query params
    const { period = 'month' } = req.query;

    // Validate period
    if (period !== 'day' && period !== 'week' && period !== 'month') {
      return res.status(400).json({ 
        error: 'Invalid period. Must be one of: day, week, month' 
      });
    }

    // Get usage summary using uid as userId
    const summary = await UsageService.getUsageSummary(
      user.uid,  // Using uid from your JWT
      period as 'day' | 'week' | 'month'
    );

    // Set cache headers for performance (cache for 5 minutes)
    res.setHeader('Cache-Control', 's-maxage=300, stale-while-revalidate');

    return res.status(200).json(summary);

  } catch (error) {
    console.error('Error getting usage summary:', error);
    
    const errorMessage = error instanceof Error ? error.message : 'Unknown error';
    
    if (errorMessage.includes('JWT_SECRET')) {
      return res.status(500).json({ error: 'Server configuration error' });
    }
    
    return res.status(500).json({ 
      error: 'Internal server error',
      message: process.env.NODE_ENV === 'development' ? errorMessage : undefined
    });
  }
}