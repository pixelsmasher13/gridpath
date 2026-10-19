# Complete Billing Implementation Guide

## System Overview

This guide covers the **credit-based billing system** - the recommended approach for Heelix.

### Billing Model
- **$5/hour** usage rate (better than Cursor's $15/hour)
- Purchase credits in **$5 increments**
- **60 free minutes** per month (resets monthly)
- **Optional auto-recharge** - Users can enable automatic $5 recharge when balance < $1
- Uses **Stripe Link** for one-click purchases

### Auto-Recharge Option (YES, it's available!)
Users can choose to enable auto-recharge for a seamless experience:
- **OFF by default** - Users must explicitly opt-in
- **Triggers at $1 balance** - Predictable threshold
- **Always $5** - Fixed amount, no surprises
- **Can toggle on/off anytime** - Full user control
- **Requires saved payment method** - From first manual purchase

### Why Credits with Optional Auto-recharge?
- **Best of both worlds** - Manual control OR automation
- **Lower barrier to entry** - Start with just $5, no auto-charge fear
- **User choice** - Power users can enable auto-recharge
- **Trust building** - Start manual, upgrade to auto when comfortable
- **Predictable** - Always $5, not variable amounts like old system

## File Structure

### Credit-Based System (RECOMMENDED - Use These)
```
proxy-server/
├── database/
│   └── credits-billing-schema.sql          # Database schema
├── lib/
│   └── creditsBillingService.ts           # Core billing logic
├── api/
│   ├── billing/
│   │   ├── buy-credits.ts                 # Purchase credits endpoint
│   │   ├── credits-status.ts              # Get billing status
│   │   ├── toggle-auto-recharge.ts        # Enable/disable auto-recharge
│   │   └── stripe-webhook-credits.ts      # Stripe webhook handler
│   └── usage/
│       └── end-credits.ts                 # End session & deduct credits
```

### Legacy Auto-charge System (DO NOT USE)
```
# These files are from the old system - ignore them:
- database/usage-billing-schema.sql
- lib/usageBillingService.ts  
- api/billing/status.ts
- api/billing/add-payment-method.ts
- api/billing/update-spending-limit.ts
- api/billing/stripe-webhook.ts
- api/usage/end-usage.ts
```

## Step-by-Step Implementation

### Step 1: Database Setup

Run the credit-based schema:
```bash
psql $DATABASE_URL < proxy-server/database/credits-billing-schema.sql
```

This creates:
- `user_billing_credits` - Main user billing table
- `usage_sessions_credits` - Usage tracking
- `credit_purchases` - Purchase history
- `auto_recharge_history` - Auto-recharge audit log

### Step 2: Stripe Configuration

#### 2.1 Get API Keys
1. Go to [Stripe Dashboard](https://dashboard.stripe.com/apikeys)
2. Copy:
   - **Secret key**: `sk_test_...` or `sk_live_...`
   - **Publishable key**: `pk_test_...` or `pk_live_...`

#### 2.2 Enable Stripe Link
1. Go to [Payment Methods Settings](https://dashboard.stripe.com/settings/payment_methods)
2. Enable **Link** (for one-click checkout)
3. This is crucial for user experience!

#### 2.3 Create Webhook
1. Go to [Webhooks](https://dashboard.stripe.com/webhooks)
2. Click **"Add endpoint"**
3. URL: `https://YOUR-DOMAIN.vercel.app/api/billing/stripe-webhook-credits`
4. Select events:
   - `payment_intent.succeeded`
   - `payment_intent.payment_failed`
   - `checkout.session.completed`
   - `setup_intent.succeeded`
5. Copy the **Signing secret** (`whsec_...`)

### Step 3: Environment Variables

Add to Vercel/your deployment:

```env
# Stripe
STRIPE_SECRET_KEY=sk_test_...
STRIPE_PUBLISHABLE_KEY=pk_test_...
STRIPE_WEBHOOK_SECRET=whsec_...

# Database
DATABASE_URL=postgresql://...

# Auth
JWT_SECRET=your-random-secret-string

# App
APP_URL=https://your-app.vercel.app
```

### Step 4: API Endpoints Reference

#### Purchase Credits
```typescript
POST /api/billing/buy-credits
Authorization: Bearer <jwt-token>

// Standard purchase (redirects to Stripe)
{
  "amountDollars": 5  // Must be in $5 increments
}

// Quick purchase (uses saved payment)
{
  "amountDollars": 10,
  "useQuickPurchase": true
}

Response:
{
  "success": true,
  "checkoutUrl": "https://checkout.stripe.com/..."  // Redirect here
}
```

#### Get Billing Status
```typescript
GET /api/billing/credits-status
Authorization: Bearer <jwt-token>

Response:
{
  "freeMinutesRemaining": 45,
  "freeMinutesResetDate": "2025-10-01",
  "creditBalance": {
    "dollars": 5.50,
    "minutes": 33
  },
  "autoRecharge": {
    "enabled": false,
    "thresholdDollars": 1,
    "amountDollars": 10
  },
  "hasPaymentMethod": true,
  "needsCredits": false,
  "canContinue": true,
  "paymentMethods": [...],
  "recentPurchases": [...]
}
```

#### Toggle Auto-Recharge
```typescript
POST /api/billing/toggle-auto-recharge
Authorization: Bearer <jwt-token>

{
  "enabled": true  // or false
}
```

#### End Usage Session
```typescript
POST /api/usage/end-credits
Authorization: Bearer <jwt-token>

{
  "sessionId": "session-123",
  "endedAt": "2025-09-08T10:30:00Z",
  "totalSeconds": 300,
  "billedMinutes": 5
}

Response:
{
  "billing": {
    "freeMinutesUsed": 5,
    "creditMinutesUsed": 0,
    "creditsDeducted": 0,
    "remainingCredits": 10.00,
    "canContinue": true
  }
}
```

### Step 5: Frontend Integration

#### Install Dependencies
```bash
npm install @stripe/stripe-js @stripe/react-stripe-js
```

#### Purchase Flow Example
```typescript
// services/billing.ts
export async function purchaseCredits(amount: number = 10) {
  const response = await fetch('/api/billing/buy-credits', {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${getToken()}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({ amountDollars: amount })
  });
  
  const data = await response.json();
  
  if (data.checkoutUrl) {
    // Redirect to Stripe
    window.location.href = data.checkoutUrl;
  }
  
  return data;
}

// Quick purchase for returning users
export async function quickPurchase(amount: number = 10) {
  return fetch('/api/billing/buy-credits', {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${getToken()}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({ 
      amountDollars: amount,
      useQuickPurchase: true 
    })
  });
}
```

#### UI Component Example
```tsx
// components/CreditBalance.tsx
export function CreditBalance() {
  const [status, setStatus] = useState(null);
  
  useEffect(() => {
    fetch('/api/billing/credits-status', {
      headers: { 'Authorization': `Bearer ${token}` }
    })
    .then(res => res.json())
    .then(setStatus);
  }, []);
  
  if (!status) return <div>Loading...</div>;
  
  return (
    <div>
      <h3>Balance</h3>
      <p>Free: {status.freeMinutesRemaining} min</p>
      <p>Credits: ${status.creditBalance.dollars}</p>
      <p>Total: {status.estimatedMinutesRemaining} min</p>
      
      {status.needsCredits && (
        <button onClick={() => purchaseCredits(5)}>
          Buy $5 Credits
        </button>
      )}
      
      <label>
        <input 
          type="checkbox" 
          checked={status.autoRecharge.enabled}
          onChange={(e) => toggleAutoRecharge(e.target.checked)}
        />
        Auto-recharge when balance < $1
      </label>
    </div>
  );
}
```

### Step 6: Testing

#### Test Cards
- **Success**: `4242 4242 4242 4242`
- **Decline**: `4000 0000 0000 0002`
- Use any future expiry date and any CVC

#### Test Flow - Manual Purchase
1. New user signs up → 60 free minutes
2. Use automation → Free minutes consumed
3. Free minutes exhausted → Prompt to buy credits
4. Purchase $5 → Stripe Link checkout
5. Success → 60 minutes added to account
6. Continue using until credits run out
7. Purchase more credits when needed

#### Test Flow - With Auto-Recharge
1. Complete initial $5 purchase (saves payment via Stripe Link)
2. Enable auto-recharge in settings
3. Use automation normally
4. When balance drops below $1 → Automatically charges $5
5. Continues seamlessly without interruption
6. User can disable auto-recharge anytime

### Step 7: Deployment Checklist

- [ ] Run database schema (`credits-billing-schema.sql`)
- [ ] Set all environment variables
- [ ] Configure Stripe webhook endpoint
- [ ] Enable Stripe Link in dashboard
- [ ] Deploy API endpoints
- [ ] Update frontend to use credit endpoints
- [ ] Test purchase flow with test card
- [ ] Test auto-recharge (if applicable)
- [ ] Switch to live Stripe keys for production

## Common Issues & Solutions

### Issue: "No payment method on file"
**Solution**: User needs to complete first purchase through Stripe checkout to save payment method via Link.

### Issue: Webhook not receiving events
**Solution**: 
1. Check webhook URL is correct
2. Verify signing secret matches
3. Ensure endpoint returns 200 status

### Issue: Credits not added after purchase
**Solution**: Check webhook logs in Stripe dashboard. Ensure `payment_intent.succeeded` event is handled.

### Issue: Auto-recharge not working
**Solution**: 
1. Verify user has payment method saved
2. Check auto-recharge is enabled
3. Confirm balance is below threshold

## Migration from Old System

If you have the old auto-charge system deployed:

1. **Keep old tables temporarily** (for reference)
2. **Run new schema** alongside
3. **Migrate user data**:
   ```sql
   -- Convert unbilled usage to credits (optional)
   UPDATE user_billing_credits 
   SET credit_balance_cents = (
     SELECT current_usage_cents 
     FROM user_billing_usage 
     WHERE user_id = user_billing_credits.user_id
   );
   ```
4. **Update webhook URL** to new endpoint
5. **Deploy new API endpoints**
6. **Update frontend** to use new endpoints
7. **Test thoroughly**
8. **Remove old tables** after confirmation

## Support & Debugging

### Check User's Billing Status
```sql
SELECT * FROM user_billing_credits WHERE user_id = 'USER_ID';
SELECT * FROM credit_purchases WHERE user_id = 'USER_ID' ORDER BY created_at DESC;
```

### Manual Credit Addition (Emergency)
```sql
SELECT add_credits('USER_ID', 500, 'manual-admin', 'manual');
-- Adds $5 (500 cents) to user account
```

### View Stripe Logs
1. Go to [Stripe Dashboard](https://dashboard.stripe.com)
2. Check **Developers → Logs** for API calls
3. Check **Webhooks → [Your endpoint] → Attempts** for webhook issues

## Architecture Decision Record

**Why Credits over Auto-charge:**
- Reduces signup friction (no card required initially)
- Users control spending explicitly
- $5 entry point vs unlimited liability fear
- Stripe Link enables fast repeat purchases
- Optional auto-recharge satisfies power users
- Better for initial user acquisition

**Why $5/hour vs $15/hour:**
- Significantly more competitive with alternatives
- Psychological pricing (very low barrier)
- Easier mental math for users
- Scales with volume

This is the complete, production-ready billing system. Use this guide for all implementation and troubleshooting.