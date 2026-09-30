import jwt from 'jsonwebtoken';

export interface UserPayload {
  uid: string;
  email: string;
  name?: string;
}

export function generateJWT(user: UserPayload): string {
  const secret = process.env.JWT_SECRET;
  if (!secret) {
    throw new Error('JWT_SECRET environment variable is required');
  }

  return jwt.sign(
    {
      uid: user.uid,
      email: user.email,
      name: user.name,
    },
    secret,
    {
      expiresIn: '7d', // Token expires in 7 days
      issuer: 'heelix-proxy',
      audience: 'heelix-app',
    }
  );
}

export function verifyJWT(token: string): UserPayload | null {
  try {
    const secret = process.env.JWT_SECRET;
    if (!secret) {
      throw new Error('JWT_SECRET environment variable is required');
    }

    const decoded = jwt.verify(token, secret, {
      issuer: 'heelix-proxy',
      audience: 'heelix-app',
    }) as any;

    return {
      uid: decoded.uid,
      email: decoded.email,
      name: decoded.name,
    };
  } catch (error) {
    console.error('JWT verification failed:', error);
    return null;
  }
} 