# Agent Guidelines for mcp-apx-ts

## Quick Commands

```bash
# Install dependencies
npm install

# Build the project
npm run build

# Run in development mode
npm run dev

# Run all tests
npm test
```

## Running Tests

| Command | Description |
|---------|-------------|
| `npm test` | Run all tests |
| `npm run test:unit` | Run unit tests only |
| `npm run test:integration` | Run integration tests |
| `npm run test:coverage` | Run tests with coverage |

## Linting

```bash
npm run lint        # Check for issues
npm run lint:fix    # Auto-fix issues
npm run lint:tests  # Lint test files
```

## Building

```bash
npm run build       # Compile TypeScript
npm run clean       # Remove dist folder
```

## Required Environment Variables

- `CLIENT_ID` - Azure app client ID
- `CLIENT_SECRET` - Azure app client secret
- `TENANT_ID` - Azure tenant ID

## Verification Steps

After making changes:
1. Run `npm run build` to compile
2. Run `npm run lint` to check code style
3. Run `npm test` to run tests