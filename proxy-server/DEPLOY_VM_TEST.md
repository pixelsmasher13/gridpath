# VM Test Deployment Instructions

## 1. Set Environment Variables in Vercel Dashboard

Go to your Vercel project settings and add:

```env
# AWS Credentials
AWS_ACCESS_KEY_ID=your_key_here
AWS_SECRET_ACCESS_KEY=your_secret_here
AWS_REGION=us-east-1

# AWS Resources (from your setup)
SUBNET_ID=subnet-0bb41e42389024b54
SECURITY_GROUP_ID=sg-05f2123070cc30f88
SQS_QUEUE_URL=https://queue.amazonaws.com/510055035486/heelix-job-queue

# VM Configuration
AMI_ID=ami-0a3366e611aa218d9
INSTANCE_TYPE=t3.medium

# Security
CRON_SECRET=any_random_string_here

# Database URL (already set in Vercel)
# DATABASE_URL=postgresql://...
```

## 2. Deploy to Vercel

```bash
cd proxy-server
vercel --prod
```

## 3. Test Endpoints

### Check Pool Status
```bash
curl https://your-proxy-server.vercel.app/api/cloud/vm/status
```

### Launch First VM
```bash
curl -X POST https://your-proxy-server.vercel.app/api/cloud/vm/launch \
  -H "Content-Type: application/json"
```

### Wait for Cron (or trigger manually)
The cron job runs every 5 minutes to maintain 1 VM.

To trigger manually:
```bash
curl -X POST https://your-proxy-server.vercel.app/api/cron/vm-pool \
  -H "Authorization: Bearer YOUR_CRON_SECRET"
```

## 4. Monitor VM Launch

Keep checking status until you see:
- state: "running"
- health: "healthy"
- ip_address: "x.x.x.x"

The VM takes 2-3 minutes to fully initialize.

## 5. Test VM Access

Once running, you should be able to:
- RDP to the public IP (if you update security group)
- Check NVDA API: http://VM_IP:8765/api/status

## Current Structure

```
api/cloud/vm/
  ├── launch.ts     ✅ Deploy
  ├── status.ts     ✅ Deploy
  ├── command.ts    ✅ Deploy
  └── poll.ts       ✅ Deploy

api/cloud/
  ├── sessions.disabled/   ❌ Not deployed
  ├── recording.disabled/  ❌ Not deployed
  ├── events.disabled/     ❌ Not deployed
  ├── automations.ts.disabled ❌ Not deployed
  └── run.ts.disabled      ❌ Not deployed

api/cron/
  └── vm-pool.ts    ✅ Deploy
```