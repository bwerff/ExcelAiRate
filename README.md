# ExcelAIRate - AI-Powered Excel Assistant

Transform your Excel data with AI-powered analysis and content generation. Built with Supabase, Next.js, and Office.js.

## 🚀 Quick Start

### Prerequisites

- Node.js 18+ 
- pnpm 9.0.0+
- Supabase account
- OpenAI API key
- Stripe account (for payments)
- Microsoft 365 account (for testing Excel add-in)

### 1. Clone and Install

```bash
git clone https://github.com/yourusername/excelairate.git
cd excelairate
pnpm install
```

### 2. Set up Supabase

1. Create a new Supabase project at [supabase.com](https://supabase.com)
2. Run the simplified database migration:
   ```bash
   npx supabase db push --file supabase/migrations/20250105_simplified_schema.sql
   ```
3. Deploy Edge Functions:
   ```bash
   npx supabase functions deploy ai
   npx supabase functions deploy stripe-checkout
   npx supabase functions deploy stripe-portal
   npx supabase functions deploy stripe-webhook
   ```

### 3. Configure Environment Variables

```bash
cp .env.example .env.local
```

Update `.env.local` with your credentials:
- Supabase URL and keys
- OpenAI API key
- Stripe keys and price IDs

### 4. Start Development

```bash
# Start all services
pnpm dev

# Or start individually:
pnpm --filter @excelairate/web dev        # Web dashboard on http://localhost:3000
pnpm --filter @excelairate/excel-addin dev # Excel add-in on https://localhost:5173
```

### 5. Load Excel Add-in

1. Open Excel (Desktop or Online)
2. Go to Insert → My Add-ins → Upload My Add-in
3. Browse to `excel-addin/manifest.xml`
4. Click Upload

## 📁 Project Structure

```
excelairate/
├── excel-addin/        # Excel Add-in (Vite + Office.js)
├── web/               # Web Dashboard (Next.js)
├── ui/                # Shared UI Components
├── config/            # Configuration packages
│   ├── eslint/        # ESLint config
│   └── typescript/    # TypeScript config
├── supabase/          # Database & Edge Functions
│   ├── migrations/    # SQL migrations
│   └── functions/     # Edge Functions
└── docs/              # Documentation
```

## 🏗️ Architecture (Ultra-Simplified)

### Tech Stack
- **Backend**: Supabase (Database + Auth + API + Edge Functions)
- **AI**: OpenAI GPT-4 Turbo
- **Payments**: Stripe Checkout + Customer Portal
- **Frontend**: Next.js + Excel Add-in

### Key Simplifications
- **4 database tables** instead of 8+
- **Single AI endpoint** for all operations
- **Magic link auth** (no passwords)
- **3-tier pricing** (Free, Pro $29, Team $99)
- **Stripe handles** all subscription logic
- **No vector search** (using full-text search)

### Features Maintained
- Natural language Excel analysis
- AI content generation
- 60% cost reduction via caching
- Usage tracking and limits
- Template library
- Real-time updates

## 🔧 Development

### Database Commands

```bash
# Run migrations
npx supabase db push

# Reset database
npx supabase db reset

# Generate types
npx supabase gen types typescript --local > types/supabase.ts
```

### Testing

```bash
# Run tests (when implemented)
pnpm test

# Type checking
pnpm check-types

# Linting
pnpm lint
```

### Deployment

```bash
# Build all packages
pnpm build

# Deploy Edge Functions
npx supabase functions deploy

# Deploy web dashboard (Vercel)
vercel --prod
```

## 📊 Pricing (Simplified 3-Tier)

- **Free**: 10 queries/month
- **Pro**: $29/month - 500 queries (or $290/year)
- **Team**: $99/month - 5000 queries (or $990/year)

## 🛠️ Technology Stack

- **Frontend**: Next.js 15, React 19, TailwindCSS 4
- **Excel Add-in**: Vite, TypeScript, Office.js
- **Backend**: Supabase (PostgreSQL, Edge Functions, Auth)
- **AI**: OpenAI GPT-4 Turbo
- **Payments**: Stripe
- **Infrastructure**: Vercel (web), Supabase (backend)

## 📝 License

MIT License - see LICENSE file for details

## 🤝 Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📧 Support

- Documentation: [docs.excelairate.com](https://docs.excelairate.com)
- Email: support@excelairate.com
- Issues: [GitHub Issues](https://github.com/yourusername/excelairate/issues)