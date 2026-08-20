import { query, getClient } from './db';

export interface UsageSession {
  id: string;
  userId: string;
  automationId: number;
  automationName: string;
  startedAt: Date;
  endedAt?: Date;
  totalSeconds: number;
  billedMinutes: number;
  status: 'active' | 'completed' | 'failed';
}

export interface UsageSummary {
  totalMinutes: number;
  totalSessions: number;
  averageSessionMinutes: number;
  periodComparison: {
    percentageChange: number;
    direction: 'increase' | 'decrease' | 'none';
  };
  topAutomations: Array<{
    name: string;
    minutes: number;
    count: number;
  }>;
  dailyUsage: Array<{
    date: string;
    minutes: number;
  }>;
}

export class UsageService {
  // Start a new usage session
  static async startSession(
    sessionId: string,
    userId: string,
    automationId: number,
    automationName: string,
    startedAt: Date
  ): Promise<void> {
    const text = `
      INSERT INTO usage_sessions (id, user_id, automation_id, automation_name, started_at, status)
      VALUES ($1, $2, $3, $4, $5, 'active')
      ON CONFLICT (id) DO NOTHING
    `;
    
    await query(text, [sessionId, userId, automationId, automationName, startedAt]);
  }

  // Update an active session
  static async updateSession(
    sessionId: string,
    elapsedSeconds: number,
    billedMinutes: number
  ): Promise<void> {
    const text = `
      UPDATE usage_sessions 
      SET total_seconds = $2, 
          billed_minutes = $3,
          updated_at = CURRENT_TIMESTAMP
      WHERE id = $1 AND status = 'active'
    `;
    
    await query(text, [sessionId, elapsedSeconds, billedMinutes]);
  }

  // End a session
  static async endSession(
    sessionId: string,
    endedAt: Date,
    totalSeconds: number,
    billedMinutes: number,
    status: 'completed' | 'failed'
  ): Promise<void> {
    const text = `
      UPDATE usage_sessions 
      SET ended_at = $2,
          total_seconds = $3,
          billed_minutes = $4,
          status = $5,
          updated_at = CURRENT_TIMESTAMP
      WHERE id = $1
    `;
    
    await query(text, [sessionId, endedAt, totalSeconds, billedMinutes, status]);
  }

  // Get active sessions for a user (for recovery)
  static async getActiveSessions(userId: string): Promise<UsageSession[]> {
    const text = `
      SELECT * FROM usage_sessions 
      WHERE user_id = $1 AND status = 'active'
      ORDER BY started_at DESC
    `;
    
    const result = await query(text, [userId]);
    return result.rows;
  }

  // Get usage summary for a user
  static async getUsageSummary(userId: string, period: 'day' | 'week' | 'month'): Promise<UsageSummary> {
    // Determine date range
    const now = new Date();
    let startDate: Date;
    let previousStartDate: Date;
    
    switch (period) {
      case 'day':
        startDate = new Date(now.getFullYear(), now.getMonth(), now.getDate());
        previousStartDate = new Date(startDate);
        previousStartDate.setDate(previousStartDate.getDate() - 1);
        break;
      case 'week':
        const dayOfWeek = now.getDay();
        startDate = new Date(now);
        startDate.setDate(now.getDate() - dayOfWeek);
        startDate.setHours(0, 0, 0, 0);
        previousStartDate = new Date(startDate);
        previousStartDate.setDate(previousStartDate.getDate() - 7);
        break;
      case 'month':
        startDate = new Date(now.getFullYear(), now.getMonth(), 1);
        previousStartDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
        break;
    }

    // Get current period stats
    const currentStatsQuery = `
      SELECT 
        COALESCE(SUM(billed_minutes), 0) as total_minutes,
        COUNT(*) as total_sessions,
        CASE 
          WHEN COUNT(*) > 0 THEN COALESCE(SUM(billed_minutes), 0) / COUNT(*)
          ELSE 0
        END as avg_minutes
      FROM usage_sessions
      WHERE user_id = $1 
        AND started_at >= $2
        AND status IN ('completed', 'failed')
    `;
    
    const currentStats = await query(currentStatsQuery, [userId, startDate]);
    const current = currentStats.rows[0];

    // Get previous period stats for comparison
    const previousStats = await query(currentStatsQuery, [userId, previousStartDate]);
    const previous = previousStats.rows[0];

    // Calculate percentage change
    const percentageChange = previous.total_minutes > 0
      ? Math.round(((current.total_minutes - previous.total_minutes) / previous.total_minutes) * 100)
      : 100;

    const direction = percentageChange > 0 ? 'increase' : percentageChange < 0 ? 'decrease' : 'none';

    // Get top automations
    const topAutomationsQuery = `
      SELECT 
        automation_name as name,
        SUM(billed_minutes) as minutes,
        COUNT(*) as count
      FROM usage_sessions
      WHERE user_id = $1 
        AND started_at >= $2
        AND status IN ('completed', 'failed')
      GROUP BY automation_name
      ORDER BY minutes DESC
      LIMIT 5
    `;
    
    const topAutomations = await query(topAutomationsQuery, [userId, startDate]);

    // Get daily usage
    const dailyUsageQuery = `
      SELECT 
        DATE(started_at) as date,
        SUM(billed_minutes) as minutes
      FROM usage_sessions
      WHERE user_id = $1 
        AND started_at >= $2
        AND status IN ('completed', 'failed')
      GROUP BY DATE(started_at)
      ORDER BY date ASC
    `;
    
    const dailyUsage = await query(dailyUsageQuery, [userId, startDate]);

    return {
      totalMinutes: parseInt(current.total_minutes) || 0,
      totalSessions: parseInt(current.total_sessions) || 0,
      averageSessionMinutes: Math.round(parseFloat(current.avg_minutes) || 0),
      periodComparison: {
        percentageChange: Math.abs(percentageChange),
        direction: direction as 'increase' | 'decrease' | 'none'
      },
      topAutomations: topAutomations.rows.map(row => ({
        name: row.name,
        minutes: parseInt(row.minutes) || 0,
        count: parseInt(row.count) || 0
      })),
      dailyUsage: dailyUsage.rows.map(row => ({
        date: row.date.toISOString().split('T')[0],
        minutes: parseInt(row.minutes) || 0
      }))
    };
  }

  // Clean up stale sessions (can be called periodically)
  static async cleanupStaleSessions(hoursOld: number = 24): Promise<number> {
    const text = `
      UPDATE usage_sessions
      SET status = 'failed',
          ended_at = CURRENT_TIMESTAMP,
          updated_at = CURRENT_TIMESTAMP
      WHERE status = 'active'
        AND started_at < NOW() - INTERVAL '${hoursOld} hours'
    `;
    
    const result = await query(text);
    return result.rowCount || 0;
  }

  // Run aggregation for a specific date
  static async runAggregation(date?: Date): Promise<void> {
    const targetDate = date || new Date();
    const text = 'SELECT aggregate_usage_data($1::date)';
    await query(text, [targetDate]);
  }
}