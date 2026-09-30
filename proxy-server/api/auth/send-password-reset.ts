import { VercelRequest, VercelResponse } from '@vercel/node';
import { sendPasswordResetEmail } from '../../lib/emailService';
import { corsMiddleware } from '../../lib/middleware';

export default async function handler(req: VercelRequest, res: VercelResponse) {
  // Apply CORS
  if (!corsMiddleware(req, res)) {
    return;
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const { email } = req.body;

    if (!email) {
      return res.status(400).json({ error: 'Email is required' });
    }

    // Send password reset email via Resend
    const result = await sendPasswordResetEmail(email);

    if (!result.success) {
      return res.status(500).json({
        error: result.error || 'Failed to send password reset email'
      });
    }

    res.status(200).json({
      success: true,
      message: 'Password reset email sent'
    });
  } catch (error: any) {
    console.error('Send password reset error:', error);
    res.status(500).json({
      error: error.message || 'Failed to send password reset email'
    });
  }
}
