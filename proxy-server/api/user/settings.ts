import { VercelRequest, VercelResponse } from '@vercel/node';
import { verifyFirebaseToken } from '../../lib/firebase';
import { query } from '../../lib/db';
import { corsMiddleware } from '../../lib/middleware';

export default async function handler(req: VercelRequest, res: VercelResponse) {
  // Apply CORS
  if (!corsMiddleware(req, res)) {
    return; // Early return for OPTIONS requests
  }

  try {
    // Authenticate user
    const authHeader = req.headers.authorization;
    if (!authHeader?.startsWith('Bearer ')) {
      return res.status(401).json({ error: 'Unauthorized' });
    }

    const firebaseToken = authHeader.substring(7);
    const decoded = await verifyFirebaseToken(firebaseToken);

    if (!decoded.valid || !decoded.uid) {
      return res.status(401).json({ error: 'Invalid token' });
    }

    const userId = decoded.uid;

    if (req.method === 'GET') {
      // Get user settings
      const result = await query(
        'SELECT * FROM user_settings WHERE user_id = $1',
        [userId]
      );

      if (result.rows.length === 0) {
        // No settings yet - return defaults
        return res.status(200).json({
          user_id: userId,
          save_browser_profile: null,  // null means user hasn't been asked yet
          use_pro_model: false,
          use_ultra_model: false,
          created_at: null,
          updated_at: null
        });
      }

      return res.status(200).json(result.rows[0]);
    }

    if (req.method === 'POST' || req.method === 'PUT') {
      // Update user settings - only update fields that are provided.
      // Pro and Ultra are mutually exclusive: when the client sets one to true we
      // force the other to false so the picker can't end up in an inconsistent
      // both-on state regardless of the order of upserts.
      const { save_browser_profile, use_pro_model, use_ultra_model } = req.body;

      let resolvedPro: boolean | null = use_pro_model ?? null;
      let resolvedUltra: boolean | null = use_ultra_model ?? null;
      if (use_ultra_model === true) resolvedPro = false;
      if (use_pro_model === true) resolvedUltra = false;

      // Upsert user settings with COALESCE to preserve existing values for fields
      // the caller didn't set
      const result = await query(
        `INSERT INTO user_settings (user_id, save_browser_profile, use_pro_model, use_ultra_model, created_at, updated_at)
         VALUES ($1, $2, $3, $4, NOW(), NOW())
         ON CONFLICT (user_id)
         DO UPDATE SET
           save_browser_profile = COALESCE($2, user_settings.save_browser_profile),
           use_pro_model = COALESCE($3, user_settings.use_pro_model),
           use_ultra_model = COALESCE($4, user_settings.use_ultra_model),
           updated_at = NOW()
         RETURNING *`,
        [userId, save_browser_profile ?? null, resolvedPro, resolvedUltra]
      );

      return res.status(200).json(result.rows[0]);
    }

    return res.status(405).json({ error: 'Method not allowed' });

  } catch (error: any) {
    console.error('User settings error:', error);
    return res.status(500).json({
      error: 'Failed to process user settings',
      details: error.message
    });
  }
}
