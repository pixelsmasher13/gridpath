// Supabase client for the desktop app.
// Shares the same Supabase project as the web app.
// Used for remote access relay (desktop_agents + remote_commands tables).

import { createClient, SupabaseClient } from '@supabase/supabase-js'

const supabaseUrl = import.meta.env.VITE_SUPABASE_URL
const supabaseAnonKey = import.meta.env.VITE_SUPABASE_ANON_KEY

let _supabase: SupabaseClient | null = null

export function getSupabase(): SupabaseClient | null {
  if (_supabase) return _supabase
  if (!supabaseUrl || !supabaseAnonKey) {
    console.warn('[Supabase] Missing VITE_SUPABASE_URL or VITE_SUPABASE_ANON_KEY — remote access disabled')
    return null
  }
  _supabase = createClient(supabaseUrl, supabaseAnonKey)
  return _supabase
}

export function isSupabaseConfigured(): boolean {
  return !!(supabaseUrl && supabaseAnonKey)
}
