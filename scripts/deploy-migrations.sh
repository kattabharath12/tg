#!/bin/bash

# Deploy migrations script for Railway
# This script can be run manually to deploy migrations to the production database

echo "🚀 Starting database migration deployment..."

# Check if DATABASE_URL is set
if [ -z "$DATABASE_URL" ]; then
    echo "❌ ERROR: DATABASE_URL environment variable is not set"
    echo "Please set your Railway database URL:"
    echo "export DATABASE_URL='your_railway_database_url_here'"
    exit 1
fi

echo "✅ DATABASE_URL is configured"

# Generate Prisma client
echo "📦 Generating Prisma client..."
npx prisma generate

# Deploy migrations
echo "🔄 Deploying migrations..."
npx prisma migrate deploy

# Check migration status
echo "📊 Checking migration status..."
npx prisma migrate status

echo "✅ Migration deployment completed successfully!"
echo "Your database schema is now up to date."
