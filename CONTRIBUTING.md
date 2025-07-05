# Contributing to ExcelAIRate

Thank you for your interest in contributing to ExcelAIRate! This document provides guidelines and instructions for contributing to the project.

## Development Setup

### Prerequisites

- Node.js >= 18 (use the version specified in `.nvmrc`)
- pnpm 9.0.0 (will be automatically installed via corepack)
- Git

### Getting Started

1. Clone the repository:
```bash
git clone https://github.com/yourusername/excelairate.git
cd excelairate
```

2. Install dependencies:
```bash
pnpm install
```

3. Set up environment variables:
```bash
cp .env.example .env
```
Fill in the required environment variables in `.env`.

4. Start development servers:
```bash
pnpm dev
```

This will start both the web dashboard and Excel add-in in development mode.

## Project Structure

```
excel-addin/      # Microsoft Excel Add-in
web/              # Next.js web dashboard  
lib/              # Shared libraries
shared/           # Shared React components
supabase/         # Backend configuration
├── migrations/   # Database schema
└── functions/    # Edge Functions
```

## Development Guidelines

### Code Style

- We use Prettier for code formatting. Run `pnpm format` before committing.
- TypeScript is used throughout the project. **Never use `any` type**.
- Follow the existing code patterns and conventions.

### Git Workflow

1. Create a feature branch from `main`:
```bash
git checkout -b feature/your-feature-name
```

2. Make your changes and commit with clear messages:
```bash
git commit -m "feat: add new feature"
```

3. Push your branch and create a pull request.

### Commit Message Format

We follow conventional commits:
- `feat:` New features
- `fix:` Bug fixes
- `docs:` Documentation changes
- `style:` Code style changes (formatting, etc.)
- `refactor:` Code refactoring
- `test:` Test additions or changes
- `chore:` Build process or auxiliary tool changes

### Testing

Currently, no testing framework is configured. When adding new features, ensure:
- Manual testing in both web and Excel add-in environments
- No TypeScript errors (`pnpm check-types`)
- No linting errors (`pnpm lint`)

### Common Commands

```bash
# Development
pnpm dev                   # Start all apps in dev mode
pnpm --filter @excelairate/web dev       # Start only web app
pnpm --filter @excelairate/excel-addin dev  # Start only Excel add-in

# Build
pnpm build                 # Build all packages
pnpm --filter @excelairate/web build     # Build only web app

# Code Quality
pnpm lint                  # Run ESLint
pnpm format               # Format code with Prettier
pnpm check-types          # TypeScript type checking
```

### Working with Supabase

1. Database migrations go in `supabase/migrations/`
2. Edge Functions go in `supabase/functions/`
3. Use Row Level Security (RLS) for all tables
4. Follow the simplified architecture - avoid unnecessary complexity

### Excel Add-in Development

1. The add-in uses Vite for development
2. Test in Excel Desktop using:
```bash
pnpm --filter @excelairate/excel-addin office:start
```
3. Validate manifest before deployment:
```bash
pnpm --filter @excelairate/excel-addin office:validate
```

## Architecture Principles

- Keep it simple - avoid adding complexity unless absolutely necessary
- Use modern, cutting-edge versions (React 19, Next.js 15, TailwindCSS 4)
- Single AI endpoint handles all operations
- Magic link authentication only
- Let Stripe handle subscription complexity

## Getting Help

- Check existing issues before creating new ones
- Provide clear reproduction steps for bugs
- Include relevant error messages and logs

## License

By contributing, you agree that your contributions will be licensed under the same license as the project.