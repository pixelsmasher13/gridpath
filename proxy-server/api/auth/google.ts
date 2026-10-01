import { VercelRequest, VercelResponse } from '@vercel/node';
import { verifyFirebaseToken } from '../../lib/firebase';
import { generateJWT } from '../../lib/jwt';
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

    // Verify Firebase token
    const firebaseResult = await verifyFirebaseToken(idToken);
    
    if (!firebaseResult.valid || !firebaseResult.uid || !firebaseResult.email) {
      return res.status(401).json({ error: 'Invalid Firebase token' });
    }

    // Generate JWT for our system
    const jwtToken = generateJWT({
      uid: firebaseResult.uid,
      email: firebaseResult.email,
      name: firebaseResult.name,
    });

    // Return user info and JWT
    res.status(200).json({
      success: true,
      token: jwtToken,
      user: {
        uid: firebaseResult.uid,
        email: firebaseResult.email,
        name: firebaseResult.name,
      },
    });
  } catch (error) {
    console.error('Google auth error:', error);
    res.status(500).json({ error: 'Authentication failed' });
  }
} 