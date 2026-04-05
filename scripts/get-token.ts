#!/usr/bin/env node
import { getBotToken } from '../dist/token.js'

async function main () {
  const clientId = process.env.CLIENT_ID
  const clientSecret = process.env.CLIENT_SECRET
  const tenantId = process.env.TENANT_ID

  if (!clientId || !clientSecret) {
    console.error('Error: CLIENT_ID and CLIENT_SECRET environment variables are required')
    console.error('')
    console.error('Usage:')
    console.error('  export CLIENT_ID="your-app-id"')
    console.error('  export CLIENT_SECRET="your-client-secret"')
    console.error('  export TENANT_ID="your-tenant-id" (optional)')
    console.error('  npm run get-token')
    console.error('')
    console.error('Or run with:')
    console.error('  CLIENT_ID="xxx" CLIENT_SECRET="yyy" TENANT_ID="zzz" npm run get-token')
    process.exit(1)
  }

  try {
    console.log('Acquiring bot token...')
    const tokenInfo = await getBotToken({
      clientId,
      clientSecret,
      tenantId,
    })

    console.log('')
    console.log('=== Bot Token Details ===')
    console.log('')
    console.log(`App ID:      ${tokenInfo.appId}`)
    console.log(`Tenant ID:   ${tokenInfo.tenantId ?? 'N/A'}`)
    console.log(`Service URL: ${tokenInfo.serviceUrl}`)
    console.log(`Expires:     ${new Date(tokenInfo.expiration).toISOString()}`)
    console.log('')
    console.log('=== Environment Variables ===')
    console.log('')
    console.log('# Add these to your shell or .env file:')
    console.log(`export TEAMS_SERVICE_URL="${tokenInfo.serviceUrl}"`)
    console.log(`export TEAMS_BOT_TOKEN="${tokenInfo.token}"`)
    console.log(`export TEST_TENANT_ID="${tokenInfo.tenantId ?? ''}"`)
    console.log(`export TEST_APP_ID="${tokenInfo.appId}"`)
    console.log('')
  } catch (error) {
    console.error('Failed to acquire token:', error instanceof Error ? error.message : error)
    process.exit(1)
  }
}

main()
