#!/bin/bash

VM_IP="44.198.185.12"

echo "Testing various ports on VM at $VM_IP"
echo "======================================="
echo ""

# Test common ports
echo "Testing RDP (3389)..."
nc -zv -w 2 $VM_IP 3389 2>&1 | grep -o "succeeded\|refused\|timed out" || echo "No response"

echo ""
echo "Testing VNC (5900)..."
nc -zv -w 2 $VM_IP 5900 2>&1 | grep -o "succeeded\|refused\|timed out" || echo "No response"

echo ""
echo "Testing NVDA (8765)..."
nc -zv -w 2 $VM_IP 8765 2>&1 | grep -o "succeeded\|refused\|timed out" || echo "No response"

echo ""
echo "Testing HTTP (80)..."
nc -zv -w 2 $VM_IP 80 2>&1 | grep -o "succeeded\|refused\|timed out" || echo "No response"

echo ""
echo "Testing Alt NVDA (8080)..."
nc -zv -w 2 $VM_IP 8080 2>&1 | grep -o "succeeded\|refused\|timed out" || echo "No response"

echo ""
echo "======================================="
echo "Results:"
echo "- 'succeeded' = Port is open and accepting connections"
echo "- 'refused' = Port is reachable but service not running"
echo "- 'timed out' = Port is blocked by firewall/security group"