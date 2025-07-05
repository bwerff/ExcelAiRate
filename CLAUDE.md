# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Repository Overview

This is an AI-powered Excel Assistant built with a simplified architecture. The project provides natural language AI analysis for Excel data using Supabase as the backend and OpenAI for AI capabilities.

## Tech Stack

- **Package Manager**: pnpm v9.0.0 with workspaces
- **Language**: TypeScript v5.8.2
- **Frontend**: 
  - Web Dashboard: Next.js 15.3.5 with React 19, TailwindCSS v4
  - Excel Add-in: Vite with TypeScript and Office.js
- **Backend**: Supabase (PostgreSQL, Auth, Edge Functions)
- **AI**: OpenAI GPT-4 Turbo
- **Payments**: Stripe Checkout
- **Node Version**: >=18

## Project Structure (Simplified)

```
excel-addin/       # Microsoft Excel Add-in (Vite + Office.js)
web/              # Next.js web dashboard
lib/              # Shared libraries (supabase.ts, pricing.ts)
supabase/         # Backend
├── migrations/   # Database schema (4 tables)
└── functions/    # Edge Functions
    ├── ai/              # Unified AI endpoint
    ├── stripe-checkout/ # Create payment sessions
    ├── stripe-portal/   # Manage subscriptions
    └── stripe-webhook/  # Handle Stripe events
```

## Common Commands

### Development
```bash
# Install dependencies
pnpm install

# Run all apps in development mode
pnpm dev

# Run specific app
pnpm --filter @excelairate/web dev
pnpm --filter @excelairate/excel-addin dev
```

### Build & Production
```bash
# Build all packages
pnpm build

# Start production server (web only)
pnpm --filter @excelairate/web start
```

### Code Quality
```bash
# Run linting
pnpm lint

# Format code with Prettier
pnpm format

# TypeScript type checking
pnpm check-types
```

## Architecture Overview

### Simplified Design
- **4 Database Tables**: profiles, usage_logs, ai_cache, templates
- **Single AI Endpoint**: /ai handles analyze, generate, and explain
- **Magic Link Auth**: No passwords, just email OTP
- **3-Tier Pricing**: Free (10), Pro ($29/500), Team ($99/5000)
- **Stripe Managed**: All subscription logic handled by Stripe

### Key Features
- Natural language Excel data analysis
- AI-powered content generation
- Response caching for 60% cost reduction
- Usage tracking and limits
- Template library with full-text search

## Development Guidelines

### Database
- Uses Supabase PostgreSQL with Row Level Security
- Simple schema focused on MVP functionality
- No complex types or vector search for now

### Authentication
- Magic link only (signInWithOtp)
- No password management
- Automatic profile creation on signup

### AI Integration
- Single Edge Function handles all AI operations
- Automatic response caching
- Usage tracking per user

### Important Notes

- No testing framework is currently configured
- The project uses cutting-edge versions (React 19, Next.js 15, TailwindCSS 4)
- NEVER BY ANY MEANS USE 'any' TYPE IN TYPESCRIPT
- Keep the architecture simple - avoid adding complexity unless absolutely necessary