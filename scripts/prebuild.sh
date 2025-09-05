
#!/bin/bash

# Prebuild script to handle Prisma setup before Next.js build
set -e

echo "🔧 Starting prebuild process..."

# Check if DATABASE_URL is set
if [ -z "$DATABASE_URL" ]; then
    echo "⚠️  DATABASE_URL not set, using default for build"
    export DATABASE_URL="postgresql://user:password@localhost:5432/taxdb"
fi

# Set build-time environment variables
export NODE_ENV="${NODE_ENV:-production}"
export PRISMA_CLI_BINARY_TARGETS="native,linux-musl-arm64-openssl-3.0.x"

echo "📦 Installing dependencies if needed..."
if [ ! -d "node_modules" ]; then
    npm ci --only=production
fi

echo "🗄️  Generating Prisma client..."
npx prisma generate --schema=./prisma/schema.prisma

# Only run migrations in production/deployment
if [ "$NODE_ENV" = "production" ] && [ -n "$DATABASE_URL" ] && [[ "$DATABASE_URL" != *"localhost"* ]]; then
    echo "🚀 Running database migrations..."
    npx prisma migrate deploy --schema=./prisma/schema.prisma || echo "⚠️  Migration failed or not needed, continuing..."
else
    echo "⏭️  Skipping migrations (development/local build)"
fi

echo "✅ Prebuild completed successfully!"
