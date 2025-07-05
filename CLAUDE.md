# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Repository Overview

This is an AI-powered Excel Assistant monorepo built with Turborepo. The project aims to create the world's first native AI-powered Excel Add-in that transforms spreadsheet users into data scientists through natural language processing.

## Tech Stack

- **Package Manager**: pnpm v9.0.0 with workspaces
- **Language**: TypeScript v5.8.2
- **Frontend**: 
  - Web Dashboard: Next.js 15.3.5 with React 19, TailwindCSS v4
  - Excel Add-in: Vite with TypeScript and Office.js
- **Node Version**: >=18

## Project Structure

```
apps/
├── excel-addin/    # Microsoft Excel Add-in built with Vite and Office.js
└── web/           # Next.js web application for dashboard/admin

packages/
├── ui/            # Shared React UI components (@repo/ui)
├── eslint-config/ # Shared ESLint configurations
└── typescript-config/ # Shared TypeScript configurations
```

## Common Commands

### Development
```bash
# Install dependencies
pnpm install

# Run all apps in development mode
pnpm dev

# Run specific app
pnpm --filter excel-addin dev
pnpm --filter web dev
```

### Build & Production
```bash
# Build all packages
pnpm build

# Build specific app
pnpm --filter excel-addin build
pnpm --filter web build

# Start production server (web app only)
pnpm --filter web start
```

### Code Quality
```bash
# Run linting across all packages
pnpm lint

# Format code with Prettier
pnpm format

# TypeScript type checking
pnpm check-types
```

## Architecture Overview

### Core Product Features
1. **AI Data Analysis**: Natural language analysis of Excel data with structured insights
2. **Content Generation**: AI-powered content creation for Excel worksheets
3. **User Authentication & Subscription**: Secure accounts with subscription-based access

### Service Architecture
The application follows a microservices architecture with:
- Event-driven architecture for scalability
- Multi-layer caching strategy for performance
- Zero-trust security model

### Data Architecture
- Primary Database: Supabase PostgreSQL with Row Level Security (is this needed?)
- Cache Layer: Redis
- AI Integration: OpenAI GPT-4 Turbo
- Payment Processing: Stripe

## Development Guidelines

### Monorepo Workflow
- Dependencies between packages are managed automatically
- Use `pnpm --filter <package-name>` to run commands in specific packages

### Excel Add-in Development
- Uses Office.js for Excel integration
- Vite for fast development and HMR
- TypeScript for type safety
- Compatible with Excel 2016+, Excel Online, Excel Mobile

### Web Dashboard Development
- Next.js 15 with App Router
- React 19 with Server Components
- TailwindCSS v4 for styling
- TypeScript for type safety

## Important Notes

- No testing framework is currently configured
- The project uses cutting-edge versions (React 19, Next.js 15, TailwindCSS 4)
- Build outputs are configured for `.next/**` in turbo.json
- All packages use TypeScript with shared configurations
- NEVER BY ANY MEANS USE 'any' TYPE IN TYPESCRIPT