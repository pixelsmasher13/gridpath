// Quick script to check VM status in database
import { query } from '../lib/db'

async function checkVMStatus() {
  try {
    console.log('Checking VM pool status in database...\n')

    const result = await query('SELECT * FROM vm_pool ORDER BY created_at DESC')

    if (result.rows.length === 0) {
      console.log('No VMs found in database')
    } else {
      console.log(`Found ${result.rows.length} VM(s):\n`)

      for (const vm of result.rows) {
        console.log(`Instance ID: ${vm.instance_id}`)
        console.log(`State: ${vm.state}`)
        console.log(`Health: ${vm.health_status}`)
        console.log(`Pool Status: ${vm.pool_status}`)
        console.log(`Public IP: ${vm.ip_address || 'Not assigned'}`)
        console.log(`Private IP: ${vm.private_ip || 'Not assigned'}`)
        console.log(`Created: ${vm.created_at}`)
        console.log(`Last Health Check: ${vm.last_health_check || 'Never'}`)
        console.log('---')
      }
    }

    process.exit(0)
  } catch (error) {
    console.error('Error checking VM status:', error)
    process.exit(1)
  }
}

checkVMStatus()