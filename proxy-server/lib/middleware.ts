import { VercelRequest, VercelResponse } from '@vercel/node';
import { verifyJWT, UserPayload } from './jwt';
import { verifyFirebaseToken } from './firebase';

export interface AuthenticatedRequest extends VercelRequest {
  user: UserPayload;
}

export function corsMiddleware(req: VercelRequest, res: VercelResponse): boolean {
  // Set CORS headers
  const origin = req.headers.origin;
  const allowedOrigins = [
    'http://localhost:3000',
    'http://localhost:5173',
    'http://localhost:5174',
    'http://localhost:5175',
    'https://app.linefox.ai',
    'https://linefox.ai',
    'https://www.linefox.ai',
    'https://billing.linefox.ai',
  ];

  if (allowedOrigins.includes(origin as string)) {
    res.setHeader('Access-Control-Allow-Origin', origin as string);
  }

  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
  res.setHeader('Access-Control-Allow-Credentials', 'true');

  // Handle preflight requests
  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return false; // Don't continue processing
  }

  return true; // Continue processing
}

export async function authenticateUser(req: VercelRequest, res: VercelResponse): Promise<UserPayload | null> {
  // Internal service key — bypasses Firebase verification for server-to-server calls
  const internalKey = req.headers['x-internal-key'] as string;
  const internalUserId = req.headers['x-internal-user-id'] as string;
  const expectedInternalKey = process.env.INTERNAL_SERVICE_KEY;

  if (expectedInternalKey && internalKey === expectedInternalKey && internalUserId) {
    return {
      uid: internalUserId,
      email: 'internal@gridpath.ai',
      name: 'Internal Service'
    };
  }

  const authHeader = req.headers.authorization;

  if (!authHeader || !authHeader.startsWith('Bearer ')) {
    res.status(401).json({ error: 'Missing or invalid authorization header' });
    return null;
  }

  const token = authHeader.substring(7); // Remove 'Bearer ' prefix

  // Try Firebase token first (for web app)
  const firebaseResult = await verifyFirebaseToken(token);
  if (firebaseResult.valid && firebaseResult.uid) {
    return {
      uid: firebaseResult.uid,
      email: firebaseResult.email || '',
      name: firebaseResult.name
    };
  }

  // Fall back to custom JWT (for desktop app)
  const jwtUser = verifyJWT(token);
  if (jwtUser) {
    return jwtUser;
  }

  res.status(401).json({ error: 'Invalid or expired token' });
  return null;
}

export function withAuth(
  handler: (req: AuthenticatedRequest, res: VercelResponse) => Promise<void>
) {
  return async (req: VercelRequest, res: VercelResponse) => {
    // Apply CORS
    if (!corsMiddleware(req, res)) {
      return; // Early return for OPTIONS requests
    }

    // Authenticate user (now async)
    const user = await authenticateUser(req, res);
    if (!user) {
      return; // Early return if authentication failed
    }

    // Add user to request object
    (req as AuthenticatedRequest).user = user;

    // Call the actual handler
    try {
      await handler(req as AuthenticatedRequest, res);
    } catch (error) {
      console.error('Handler error:', error);
      res.status(500).json({ error: 'Internal server error' });
    }
  };
} 