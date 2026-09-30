#!/bin/bash

VM_IP="44.198.185.12"

echo "Testing NVDA on VM at $VM_IP:8765"
echo "======================================="
echo ""

echo "1. Testing NVDA Status endpoint..."
curl -X GET "http://$VM_IP:8765/api/status" \
  --connect-timeout 5 \
  --max-time 10 \
  -w "\nHTTP Status: %{http_code}\n" \
  2>/dev/null || echo "Connection failed - NVDA may not be accessible"

echo ""
echo "======================================="
echo "Note: If connection failed, the security group likely blocks inbound traffic on port 8765"
echo "Security Group: sg-05f2123070cc30f88"
echo ""
echo "To fix: Add inbound rule in AWS Console for port 8765 from your IP or 0.0.0.0/0 (for testing)"