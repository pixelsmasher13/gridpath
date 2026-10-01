import jwt from 'jsonwebtoken';
import { NextRequest } from 'next/server';
import { verifyFirebaseToken } from './firebase';

// This should be the same secret used by the proxy server to sign JWT tokens
const JWT_SECRET = process.env.JWT_SECRET || 'your-jwt-secret';

export interface UserPayload {
  userId: string;
  email: string;
  name?: string;
  // Additional fields from proxy JWT
  iat?: number;
  exp?: number;
}

export function generateToken(payload: UserPayload): string {
  return jwt.sign(payload, JWT_SECRET, {
    expiresIn: '7d',
  });
}

export async function verifyToken(token: string): Promise<UserPayload> {
  // Try Firebase token first (for web app users)
  try {
    const firebaseResult = await verifyFirebaseToken(token);
    if (firebaseResult.valid && firebaseResult.uid) {
      console.log('✅ Verified Firebase token for user:', firebaseResult.email);
      return {
        userId: firebaseResult.uid,
        email: firebaseResult.email || '',
        name: firebaseResult.name,
      };
    }
  } catch (error) {
    console.log('Not a Firebase token, trying JWT...');
  }

  // Fall back to JWT token (for desktop app users)
  try {
    const decoded = jwt.verify(token, JWT_SECRET) as any;

    console.log('✅ Verified JWT token for user:', decoded.email);
    // The proxy server JWT contains these fields
    return {
      userId: decoded.userId || decoded.sub || decoded.id || decoded.uid,
      email: decoded.email,
      name: decoded.name || decoded.displayName,
      iat: decoded.iat,
      exp: decoded.exp,
    };
  } catch (error) {
    console.error('Token verification failed (both Firebase and JWT):', error);
    throw new Error('Invalid token');
  }
}

export function getTokenFromRequest(req: NextRequest): string | null {
  const authHeader = req.headers.get('authorization');
  if (authHeader?.startsWith('Bearer ')) {
    return authHeader.substring(7);
  }
  
  const cookie = req.cookies.get('auth-token');
  return cookie?.value || null;
}

export async function validateRequest(req: NextRequest): Promise<UserPayload | null> {
  const token = getTokenFromRequest(req);
  if (!token) return null;

  try {
    return await verifyToken(token);
  } catch {
    return null;
  }
}