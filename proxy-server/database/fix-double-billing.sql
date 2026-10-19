-- Fix for double billing issue in process_usage_session_credits
-- This script prevents sessions from being billed multiple times

DROP FUNCTION IF EXISTS process_usage_session_credits CASCADE;

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
      free_minutes_used,
      credit_minutes_used,
      0::INTEGER, -- No new credits deducted
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

  -- Check and reset free minutes if needed
  UPDATE user_billing_credits
  SET
    free_minutes_remaining = 60,
    free_minutes_reset_date = DATE_TRUNC('month', CURRENT_DATE) + INTERVAL '1 month',
    total_minutes_used = 0,
    current_period_start = DATE_TRUNC('month', CURRENT_DATE)
  WHERE user_id = p_user_id
    AND current_period_start < DATE_TRUNC('month', CURRENT_DATE);

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
$$ LANGUAGE plpgsql SECURITY DEFINER;

-- Grant execute permission
GRANT EXECUTE ON FUNCTION process_usage_session_credits TO authenticated;
GRANT EXECUTE ON FUNCTION process_usage_session_credits TO service_role;