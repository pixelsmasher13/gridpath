-- Credit-based billing schema
-- $10/hour with 1 hour free monthly, purchase credits in $10 increments

-- User billing with credit system
CREATE TABLE IF NOT EXISTS user_billing_credits (
  user_id VARCHAR(255) PRIMARY KEY,
  stripe_customer_id VARCHAR(255) UNIQUE,
  
  -- Free tier (resets monthly)
  free_minutes_remaining INTEGER DEFAULT 60,
  free_minutes_reset_date DATE DEFAULT (DATE_TRUNC('month', CURRENT_DATE) + INTERVAL '1 month'),
  
  -- Credit balance
  credit_balance_cents INTEGER DEFAULT 0,
  
  -- Auto-recharge settings
  auto_recharge_enabled BOOLEAN DEFAULT false,
  auto_recharge_threshold_cents INTEGER DEFAULT 100, -- Recharge when balance < $1
  auto_recharge_amount_cents INTEGER DEFAULT 500, -- Always recharge $5
  
  -- Payment method
  has_payment_method BOOLEAN DEFAULT false,
  default_payment_method_id VARCHAR(255),
  
  -- Usage tracking (current period)
  current_period_start DATE DEFAULT DATE_TRUNC('month', CURRENT_DATE),
  total_minutes_used INTEGER DEFAULT 0,
  
  -- Status
  billing_status VARCHAR(50) DEFAULT 'active' CHECK (billing_status IN ('active', 'paused', 'payment_failed')),
  
  -- Metadata
  created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
);

-- Usage sessions tracking
CREATE TABLE IF NOT EXISTS usage_sessions_credits (
  session_id VARCHAR(255) PRIMARY KEY,
  user_id VARCHAR(255) NOT NULL REFERENCES user_billing_credits(user_id),
  
  started_at TIMESTAMP WITH TIME ZONE NOT NULL,
  ended_at TIMESTAMP WITH TIME ZONE,
  
  total_minutes INTEGER DEFAULT 0,
  free_minutes_used INTEGER DEFAULT 0,
  credit_minutes_used INTEGER DEFAULT 0,
  credits_deducted_cents INTEGER DEFAULT 0,
  
  status VARCHAR(50) DEFAULT 'active' CHECK (status IN ('active', 'completed', 'failed')),
  created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
);

-- Credit purchases and auto-recharges
CREATE TABLE IF NOT EXISTS credit_purchases (
  purchase_id SERIAL PRIMARY KEY,
  user_id VARCHAR(255) NOT NULL REFERENCES user_billing_credits(user_id),
  
  amount_cents INTEGER NOT NULL,
  credits_added_cents INTEGER NOT NULL,
  
  purchase_type VARCHAR(50) NOT NULL CHECK (purchase_type IN ('manual', 'auto_recharge')),
  
  stripe_payment_intent_id VARCHAR(255) UNIQUE,
  stripe_invoice_id VARCHAR(255),
  
  status VARCHAR(50) DEFAULT 'pending' CHECK (status IN ('pending', 'succeeded', 'failed', 'refunded')),
  failure_reason TEXT,
  
  created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
);

-- Auto-recharge history
CREATE TABLE IF NOT EXISTS auto_recharge_history (
  id SERIAL PRIMARY KEY,
  user_id VARCHAR(255) NOT NULL REFERENCES user_billing_credits(user_id),
  enabled BOOLEAN NOT NULL,
  changed_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
  changed_by VARCHAR(50) -- 'user' or 'system'
);

-- Indexes
CREATE INDEX idx_credits_billing_stripe ON user_billing_credits(stripe_customer_id);
CREATE INDEX idx_credits_billing_status ON user_billing_credits(billing_status);
CREATE INDEX idx_credits_sessions_user_date ON usage_sessions_credits(user_id, started_at DESC);
CREATE INDEX idx_credit_purchases_user ON credit_purchases(user_id, created_at DESC);
CREATE INDEX idx_credit_purchases_stripe ON credit_purchases(stripe_payment_intent_id);

-- Pricing configuration
CREATE TABLE IF NOT EXISTS billing_config_credits (
  key VARCHAR(100) PRIMARY KEY,
  value_cents INTEGER,
  description TEXT
);

INSERT INTO billing_config_credits (key, value_cents, description) VALUES
  ('hourly_rate_cents', 500, '$5 per hour'),
  ('credit_purchase_amount', 500, 'Standard credit purchase $5'),
  ('auto_recharge_threshold', 100, 'Auto-recharge when balance < $1'),
  ('auto_recharge_amount', 500, 'Auto-recharge $5')
ON CONFLICT (key) DO NOTHING;

-- Function to process usage and deduct credits
CREATE OR REPLACE FUNCTION process_usage_session_credits(
  p_user_id VARCHAR(255),
  p_session_id VARCHAR(255),
  p_minutes_used INTEGER
) RETURNS TABLE(
  free_minutes_used INTEGER,
  credit_minutes_used INTEGER,
  credits_deducted_cents INTEGER,
  remaining_balance_cents INTEGER,
  needs_recharge BOOLEAN,
  can_continue BOOLEAN
) AS $$
DECLARE
  v_free_remaining INTEGER;
  v_credit_balance INTEGER;
  v_auto_recharge_enabled BOOLEAN;
  v_auto_recharge_threshold INTEGER;
  v_has_payment BOOLEAN;
  v_billing_status VARCHAR(50);
  v_free_used INTEGER := 0;
  v_credit_used INTEGER := 0;
  v_credits_deducted INTEGER := 0;
  v_rate_per_minute INTEGER := 9; -- $5/hour = ~$0.083/minute (rounded)
  v_needs_recharge BOOLEAN := false;
  v_can_continue BOOLEAN := true;
  v_existing_session_id VARCHAR(255);
BEGIN
  -- Check if session already processed to prevent double billing
  SELECT session_id INTO v_existing_session_id
  FROM usage_sessions_credits
  WHERE session_id = p_session_id;

  IF v_existing_session_id IS NOT NULL THEN
    -- Session already processed, return existing values without charging again
    RETURN QUERY
    SELECT
      usage_sessions_credits.free_minutes_used,
      usage_sessions_credits.credit_minutes_used,
      0, -- No new credits deducted
      (SELECT credit_balance_cents FROM user_billing_credits WHERE user_id = p_user_id),
      false, -- No recharge needed
      true -- Can continue
    FROM usage_sessions_credits
    WHERE session_id = p_session_id
    LIMIT 1;
    RETURN;
  END IF;

  -- Get or create user billing record
  INSERT INTO user_billing_credits (user_id)
  VALUES (p_user_id)
  ON CONFLICT (user_id) DO NOTHING;

  -- Note: Free minutes are a one-time allowance (60 total), not monthly reset

  -- Get current state
  SELECT 
    free_minutes_remaining,
    credit_balance_cents,
    auto_recharge_enabled,
    auto_recharge_threshold_cents,
    has_payment_method,
    billing_status
  INTO v_free_remaining, v_credit_balance, v_auto_recharge_enabled,
       v_auto_recharge_threshold, v_has_payment, v_billing_status
  FROM user_billing_credits
  WHERE user_id = p_user_id
  FOR UPDATE;
  
  -- Calculate free vs credit minutes
  IF v_free_remaining > 0 THEN
    v_free_used := LEAST(v_free_remaining, p_minutes_used);
    v_credit_used := p_minutes_used - v_free_used;
  ELSE
    v_credit_used := p_minutes_used;
  END IF;
  
  -- Calculate credits to deduct
  v_credits_deducted := v_credit_used * v_rate_per_minute;
  
  -- Don't let balance go negative
  IF v_credits_deducted > v_credit_balance THEN
    v_credits_deducted := v_credit_balance;
    v_credit_used := v_credit_balance / v_rate_per_minute;
  END IF;
  
  -- Update user billing
  UPDATE user_billing_credits
  SET 
    free_minutes_remaining = GREATEST(0, free_minutes_remaining - v_free_used),
    credit_balance_cents = GREATEST(0, credit_balance_cents - v_credits_deducted),
    total_minutes_used = total_minutes_used + p_minutes_used,
    updated_at = CURRENT_TIMESTAMP
  WHERE user_id = p_user_id;
  
  -- Record session (already checked for duplicates above)
  INSERT INTO usage_sessions_credits (
    session_id, user_id, started_at, ended_at,
    total_minutes, free_minutes_used, credit_minutes_used, credits_deducted_cents
  ) VALUES (
    p_session_id, p_user_id, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP,
    p_minutes_used, v_free_used, v_credit_used, v_credits_deducted
  );
  
  -- Check if needs recharge
  v_needs_recharge := v_auto_recharge_enabled AND 
                      (v_credit_balance - v_credits_deducted) < v_auto_recharge_threshold AND
                      v_has_payment;
  
  -- Check if user can continue
  IF v_billing_status != 'active' THEN
    v_can_continue := false;
  ELSIF v_free_remaining - v_free_used <= 0 AND v_credit_balance - v_credits_deducted <= 0 THEN
    v_can_continue := false;
  END IF;
  
  RETURN QUERY 
  SELECT 
    v_free_used,
    v_credit_used,
    v_credits_deducted,
    v_credit_balance - v_credits_deducted,
    v_needs_recharge,
    v_can_continue;
END;
$$ LANGUAGE plpgsql;

-- Function to add credits after purchase
CREATE OR REPLACE FUNCTION add_credits(
  p_user_id VARCHAR(255),
  p_credits_cents INTEGER,
  p_payment_intent_id VARCHAR(255),
  p_purchase_type VARCHAR(50)
) RETURNS void AS $$
BEGIN
  -- Add credits to balance
  UPDATE user_billing_credits
  SET 
    credit_balance_cents = credit_balance_cents + p_credits_cents,
    billing_status = 'active',
    updated_at = CURRENT_TIMESTAMP
  WHERE user_id = p_user_id;
  
  -- Record the purchase
  INSERT INTO credit_purchases (
    user_id, amount_cents, credits_added_cents,
    purchase_type, stripe_payment_intent_id, status
  ) VALUES (
    p_user_id, p_credits_cents, p_credits_cents,
    p_purchase_type, p_payment_intent_id, 'succeeded'
  );
END;
$$ LANGUAGE plpgsql;

-- Function to toggle auto-recharge
CREATE OR REPLACE FUNCTION toggle_auto_recharge(
  p_user_id VARCHAR(255),
  p_enabled BOOLEAN,
  p_changed_by VARCHAR(50)
) RETURNS void AS $$
BEGIN
  -- Update auto-recharge setting
  UPDATE user_billing_credits
  SET 
    auto_recharge_enabled = p_enabled,
    updated_at = CURRENT_TIMESTAMP
  WHERE user_id = p_user_id;
  
  -- Record history
  INSERT INTO auto_recharge_history (
    user_id, enabled, changed_by
  ) VALUES (
    p_user_id, p_enabled, p_changed_by
  );
END;
$$ LANGUAGE plpgsql;