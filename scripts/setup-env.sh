#!/bin/bash
# Setup script for Microsoft Teams Bot credentials
# Usage: source setup-env.sh

echo "=== Microsoft Teams Bot Environment Setup ==="
echo ""

# Check if logged in to Azure
echo "Checking Azure CLI login status..."
if ! az account show &>/dev/null; then
    echo "Not logged in to Azure. Please run: az login"
    return 1
fi

echo "✓ Logged in to Azure"
echo ""

# Show available subscriptions
echo "Available subscriptions:"
az account list --query '[].{name:name, id:id, tenant:tenantId}' -o table
echo ""

# Get subscription ID
read -p "Enter subscription ID: " SUBSCRIPTION_ID
export SUBSCRIPTION_ID

# Get tenant ID
TENANT_ID=$(az account show --query tenantId -o tsv)
echo "Tenant ID: $TENANT_ID"
export TENANT_ID

# List Bot Services
echo ""
echo "Available Bot Services:"
az resource list --resource-type Microsoft.BotService/botServices --subscription "$SUBSCRIPTION_ID" \
    --query '[].{name:name, id:id, location:location}' -o table 2>/dev/null || echo "No Bot Services found"

# Ask for bot name
read -p "Enter Bot name: " BOT_NAME
echo ""

# Get Bot details
echo "Fetching Bot details..."
BOT_DETAILS=$(az bot show --name "$BOT_NAME" --resource-group "$(az bot show --name "$BOT_NAME" --query resourceGroup -o tsv)" 2>/dev/null)

if [ $? -eq 0 ]; then
    APP_ID=$(echo "$BOT_DETAILS" | jq -r '.properties.appId')
    DISPLAY_NAME=$(echo "$BOT_DETAILS" | jq -r '.properties.displayName')
    ENDPOINT=$(echo "$BOT_DETAILS" | jq -r '.properties.endpoint')
    
    echo ""
    echo "=== Bot Configuration ==="
    echo "Display Name: $DISPLAY_NAME"
    echo "App ID: $APP_ID"
    echo "Endpoint: $ENDPOINT"
    
    # Get app password (client secret)
    echo ""
    echo "To get the bot token (client secret):"
    echo "1. Go to Azure Portal"
    echo "2. Navigate to: Azure Active Directory > App registrations > $DISPLAY_NAME"
    echo "3. Go to 'Certificates & secrets'"
    echo "4. Create a new client secret"
    echo "5. Copy the secret value (this is your TEAMS_BOT_TOKEN)"
    
    echo ""
    echo "=== Environment Variables ==="
    echo ""
    echo "export TEAMS_SERVICE_URL=\"$ENDPOINT\""
    echo "export TEST_TENANT_ID=\"$TENANT_ID\""
    echo "export TEST_APP_ID=\"$APP_ID\""
    echo "# export TEAMS_BOT_TOKEN=\"<your-client-secret>\""
    echo ""
    echo "Add these to your .env file or shell profile."
else
    echo "Bot '$BOT_NAME' not found. Please check the name and try again."
fi
