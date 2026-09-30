import { Pool } from 'pg';

// Ensure DATABASE_URL is available
if (!process.env.DATABASE_URL) {
  console.error('DATABASE_URL environment variable is not set');
  throw new Error('DATABASE_URL is required');
}

// Log connection info (without password) for debugging
const dbUrl = process.env.DATABASE_URL;
const urlParts = dbUrl.split('@');
if (urlParts.length > 1) {
  console.log('Connecting to database:', urlParts[1]);
}

// Create a connection pool optimized for serverless
// You'll need to set these environment variables in your Vercel project:
// DATABASE_URL or individual: PGHOST, PGUSER, PGPASSWORD, PGDATABASE
const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: process.env.NODE_ENV === 'production' ? { rejectUnauthorized: false } : false,
  // Serverless-optimized settings
  max: 1, // Use only 1 connection per function instance
  idleTimeoutMillis: 0, // Disable idle timeout
  connectionTimeoutMillis: 5000, // Increase connection timeout
  // Supabase pooler requires these settings
  keepAlive: false,
  allowExitOnIdle: true,
});

// Helper function to execute queries
export async function query(text: string, params?: any[]) {
  const start = Date.now();
  let client;
  try {
    client = await pool.connect();
    const res = await client.query(text, params);
    const duration = Date.now() - start;
    console.log('Executed query', { text, duration, rows: res.rowCount });
    return res;
  } catch (error) {
    console.error('Database query error:', error);
    throw error;
  } finally {
    // Always release the client back to the pool
    if (client) {
      client.release();
    }
  }
}

// Helper function to get a client for transactions
export async function getClient() {
  const client = await pool.connect();
  const query = client.query.bind(client);
  const release = () => {
    client.release();
  };

  // Set a timeout of 5 seconds, after which we will log this client's last query
  const timeout = setTimeout(() => {
    console.error('A client has been checked out for more than 5 seconds!');
  }, 5000);

  client.release = () => {
    clearTimeout(timeout);
    client.release();
  };

  return { query, release, client };
}

// Export pool for advanced use cases
export { pool };