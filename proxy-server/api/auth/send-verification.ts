import { VercelRequest, VercelResponse } from '@vercel/node';
import { verifyFirebaseToken } from '../../lib/firebase';
import { sendVerificationEmail } from '../../lib/emailService';
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
    const { idToken } = req.body;

    if (!idToken) {
      return res.status(400).json({ error: 'Firebase ID token is required' });
    }

    // Verify Firebase token to get user email
    const firebaseResult = await verifyFirebaseToken(idToken);

    if (!firebaseResult.valid || !firebaseResult.email) {
      return res.status(401).json({ error: 'Invalid Firebase token' });
    }

    // Send verification email via Resend
    const result = await sendVerificationEmail(firebaseResult.email);

    if (!result.success) {
      return res.status(500).json({
        error: result.error || 'Failed to send verification email'
      });
    }

    res.status(200).json({
      success: true,
      message: 'Verification email sent'
    });
  } catch (error: any) {
    console.error('Send verification error:', error);
    res.status(500).json({
      error: error.message || 'Failed to send verification email'
    });
  }
}
