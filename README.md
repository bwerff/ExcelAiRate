🎯 1. Executive Summary {#executive-summary}
Product Vision
Create the world's first native AI-powered Excel Add-in that transforms spreadsheet users into data scientists through natural language processing, enabling instant analysis, content generation, and insights without any coding knowledge.
Business Opportunity

Market Size: 1.3 billion Excel users worldwide
Addressable Market: 50M+ business users needing advanced analytics
Revenue Potential: $100K+ MRR within 12 months
Competitive Moat: First-mover advantage in Excel-native AI integration

Key Value Propositions

Zero Learning Curve: Natural language interface for complex analysis
Native Integration: Works directly within Excel, no app switching
Cost Effective: 90% cheaper than enterprise BI tools
Instant Results: Analysis completed in seconds, not hours
Scalable Pricing: From $12/month to enterprise solutions

Financial Summary

Target Profit Margin: 95%+ (optimized architecture)
Customer Acquisition Cost: $25-50 (organic + paid)
Customer Lifetime Value: $400-800 (20-month avg retention)
Break-even Point: Month 4 with 200 paid users


🏢 2. Product Overview {#product-overview}
What We're Building
Primary Product: Microsoft Excel Add-in with AI-powered data analysis
Secondary Product: Web dashboard for subscription management
Core Technology: Supabase + OpenAI GPT-4 Turbo + Office.js
Core Functionality
AI Data Analysis

Select Excel data → Ask questions in plain English → Get instant insights
Support for: trends, patterns, anomalies, correlations, forecasts
Response formats: Text insights, structured analysis, recommendations

Content Generation

Generate sample data, formulas, reports, charts recommendations
Fill data patterns automatically using AI predictions
Create professional summaries and presentations

Smart Features

Response caching for 60% cost reduction
Streaming responses for better UX
Usage tracking and limits by subscription tier
Template library for common analysis types

User Workflow

Install: Download from Microsoft AppSource
Sign Up: 14-day free trial, no credit card required
Analyze: Select data, type question, get AI insights
Subscribe: Choose plan based on usage needs
Scale: Use advanced features as business grows


👥 3. Market Analysis {#market-analysis}
Target Market Segmentation
Primary Market: Business Analysts (60% of customers)

Size: 15M+ globally
Pain Points: Manual analysis, complex formulas, time-consuming reports
Current Solutions: Excel VBA, Power BI ($10-20/user/month)
Willingness to Pay: $20-50/month for significant time savings
Use Cases: Financial modeling, sales analysis, performance reports

Secondary Market: Small Business Owners (30% of customers)

Size: 30M+ globally
Pain Points: Limited analytics skills, expensive BI tools
Current Solutions: Basic Excel, Google Sheets
Willingness to Pay: $10-25/month for simple AI insights
Use Cases: Sales tracking, expense analysis, customer insights

Tertiary Market: Enterprise Teams (10% of customers)

Size: 5M+ teams globally
Pain Points: Inconsistent analysis, team productivity bottlenecks
Current Solutions: Tableau ($70/user), Power BI Premium ($20/user)
Willingness to Pay: $50-200/month for team efficiency
Use Cases: Department reporting, standardized analysis, team collaboration

Competitive Landscape
Direct Competitors
CompetitorPriceStrengthsWeaknessesMicrosoft Copilot$30/monthNative integrationEnterprise-only, expensivePower BI$10-20/monthAdvanced visualizationSeparate app, complex setupTableau$70/monthProfessional featuresHigh cost, steep learning curve
Competitive Advantages

Price Point: 50-80% cheaper than alternatives
Ease of Use: Natural language vs complex interfaces
Excel Native: No app switching or data exports
Quick Setup: Working in minutes vs days/weeks
Flexible Pricing: Plans for individuals to enterprises


🏗️ 4. Technical Architecture {#technical-architecture}
System Overview
Excel Add-in ↔ Supabase (Auth/DB/Edge Functions) ↔ OpenAI GPT-4 Turbo
              ↕                                     ↕
          Stripe Webhooks                     Response Cache
Technology Stack
Frontend

Excel Plugin: Office.js Add-in (HTML/CSS/JavaScript)
Web Dashboard: Next.js 14 with App Router
Styling: TailwindCSS for responsive design
State Management: React Context + Zustand
Hosting: Vercel with global CDN

Backend (Serverless)

Database: Supabase PostgreSQL with Row Level Security
Authentication: Supabase Auth (JWT-based)
API: Supabase Edge Functions (Deno runtime)
AI Integration: OpenAI SDK v4.24+ (GPT-4 Turbo)
Payments: Stripe with webhooks
File Storage: Supabase Storage (for templates/exports)

Infrastructure

Primary Hosting: Supabase (database + edge functions)
Frontend CDN: Vercel Edge Network
Domain Management: Cloudflare (DNS + security)
Monitoring: Supabase Analytics + Sentry
CI/CD: GitHub Actions with automated deployments

Cost-Optimized Architecture Benefits

95% Profit Margins: Optimized AI usage and serverless architecture
Global Performance: Edge functions for <100ms response times
Auto-Scaling: Serverless handles 0-10K users seamlessly
High Availability: 99.9% uptime with Supabase infrastructure
Security: SOC 2 compliant with enterprise-grade encryption


⚙️ 5. Feature Specifications {#feature-specifications}
MVP Features (Version 1.0)
5.1 AI Data Analysis
Description: Natural language analysis of Excel data with structured insights
User Stories:

As a business analyst, I want to select sales data and ask "What are the trends?" to get instant insights
As a manager, I want to understand anomalies in my performance data quickly
As a consultant, I want to identify correlations between different metrics

Functional Requirements:

Support Excel ranges up to 50,000 cells
Handle mixed data types (numbers, text, dates)
Generate structured responses with summary, insights, trends, recommendations
Provide confidence scores for analysis reliability
Export results to new worksheet or clipboard

Technical Requirements:

OpenAI GPT-4 Turbo integration with function calling
Response caching for 60% cost reduction
Streaming responses for better UX
Error handling for malformed data
Rate limiting by subscription tier

Acceptance Criteria:

✅ User can analyze any Excel range in <30 seconds
✅ AI provides relevant, actionable insights 90%+ of the time
✅ Results formatted for business users (structured JSON)
✅ Handles edge cases (empty cells, mixed data types, large datasets)
✅ Usage tracked and enforced by subscription limits

5.2 Content Generation
Description: AI-powered content creation for Excel worksheets
User Stories:

As a financial analyst, I want to generate sample financial data for modeling
As a project manager, I want to create task lists and timelines
As a sales manager, I want to generate customer persona templates

Functional Requirements:

Generate structured data (tables, lists, formulas)
Create Excel formulas from natural language descriptions
Produce reports and summaries from existing data
Fill data patterns automatically using AI
Support multiple output formats (CSV, JSON, Excel-compatible)

Technical Requirements:

Template-based generation system
Excel formula generation engine
Data validation and type checking
Bulk data insertion with proper formatting
Undo/redo functionality for generated content

Acceptance Criteria:

✅ Generates Excel-compatible data in proper formats
✅ Maintains data relationships and constraints
✅ Provides preview before insertion
✅ Supports rollback of generated content
✅ Works with existing Excel formatting and formulas

5.3 User Authentication & Subscription
Description: Secure user accounts with subscription-based feature access
User Stories:

As a new user, I want to try the product for free before purchasing
As a subscriber, I want to manage my plan and usage
As an admin, I want to track team usage and costs

Functional Requirements:

Email/password registration and authentication
14-day free trial with no credit card required
Subscription plan management (upgrade/downgrade/cancel)
Real-time usage tracking and limit enforcement
Self-service account dashboard

Technical Requirements:

Supabase Auth with Row Level Security
JWT token-based authentication
Stripe subscription management
Usage monitoring with daily aggregation
Automated billing and invoice generation

Acceptance Criteria:

✅ Secure authentication flow with proper session management
✅ Seamless payment processing with Stripe
✅ Real-time usage tracking visible to users
✅ Subscription changes reflected immediately
✅ Proper error handling for payment failures

Enhanced Features (Version 1.5)
5.4 Advanced Analytics

Statistical analysis (correlation, regression, significance testing)
Predictive modeling and forecasting
Data visualization recommendations
Custom analysis templates
Batch processing for large datasets

5.5 Team Collaboration

Shared analysis templates across team members
Team usage dashboards for managers
Collaborative workspaces for projects
Usage analytics and cost allocation
Role-based access controls

5.6 Enterprise Features

Single Sign-On (SSO) integration (SAML, OAuth)
Custom branding and white-label options
Advanced security controls and audit logs
Dedicated support and onboarding
Custom AI model fine-tuning

Future Features (Version 2.0+)
5.7 Multi-AI Integration

Support for Claude (Anthropic) and Gemini (Google)
Model comparison and selection
Custom model routing based on query type
Cost optimization across providers

5.8 Advanced Integrations

Power BI connector for enhanced visualization
SharePoint integration for enterprise collaboration
Microsoft Teams notifications and bot
Outlook automation for scheduled reports
API for third-party integrations

5.9 AI Vision Capabilities

Chart and graph analysis from screenshots
Image-to-data conversion
Document parsing and data extraction
Visual dashboard creation


🗄️ 6. Database Design {#database-design}
Core Schema (Optimized for Performance & Cost)
6.1 User Profiles
sqlCREATE TABLE profiles (
  id uuid references auth.users(id) on delete cascade primary key,
  
  -- Subscription Information
  plan text default 'free' check (plan in ('free', 'starter', 'professional', 'business', 'enterprise')),
  status text default 'active' check (status in ('active', 'trialing', 'canceled', 'past_due', 'unpaid')),
  
  -- Usage Tracking (reset monthly)
  queries_used integer default 0,
  queries_limit integer default 10,
  
  -- Billing Integration
  stripe_customer_id text unique,
  stripe_subscription_id text unique,
  current_period_start timestamptz,
  current_period_end timestamptz,
  
  -- User Preferences
  preferred_model text default 'gpt-4-turbo',
  timezone text default 'UTC',
  
  -- Metadata
  created_at timestamptz default now(),
  updated_at timestamptz default now(),
  last_active_at timestamptz default now(),
  
  -- Indexes for performance
  CONSTRAINT idx_profiles_stripe_customer ON stripe_customer_id,
  CONSTRAINT idx_profiles_status ON status,
  CONSTRAINT idx_profiles_updated_at ON updated_at
);
6.2 Usage Analytics (Aggregated)
sqlCREATE TABLE usage_daily (
  user_id uuid references auth.users(id) on delete cascade,
  date date default current_date,
  
  -- Usage Counters
  queries_total integer default 0,
  queries_analysis integer default 0,
  queries_generation integer default 0,
  queries_explanation integer default 0,
  
  -- Cost Tracking
  tokens_total integer default 0,
  cost_total decimal(8,4) default 0,
  
  -- Performance Metrics
  avg_response_time integer, -- milliseconds
  cache_hit_rate decimal(3,2) default 0,
  success_rate decimal(3,2) default 1.00,
  
  -- AI Model Usage
  gpt4_queries integer default 0,
  gpt35_queries integer default 0,
  
  primary key (user_id, date)
);

-- Partition by month for performance
ALTER TABLE usage_daily PARTITION BY RANGE (date);
6.3 Response Cache (Cost Optimization)
sqlCREATE TABLE response_cache (
  prompt_hash text primary key,
  response_data jsonb not null,
  model_used text not null,
  
  -- Cache Management
  created_at timestamptz default now(),
  last_accessed_at timestamptz default now(),
  hit_count integer default 1,
  
  -- Response Metadata
  tokens_saved integer default 0,
  cost_saved decimal(8,4) default 0,
  
  -- TTL constraint (auto-cleanup)
  CONSTRAINT cache_ttl CHECK (created_at > now() - interval '7 days')
);

-- Indexes for fast lookups
CREATE INDEX idx_response_cache_last_accessed ON response_cache(last_accessed_at);
CREATE INDEX idx_response_cache_model ON response_cache(model_used);
6.4 AI Templates (User-Generated)
sqlCREATE TABLE ai_templates (
  id uuid default gen_random_uuid() primary key,
  user_id uuid references auth.users(id) on delete cascade,
  
  -- Template Content
  name text not null,
  description text,
  prompt_template text not null,
  category text default 'general',
  
  -- Sharing & Discovery
  is_public boolean default false,
  is_featured boolean default false,
  
  -- Usage Statistics
  usage_count integer default 0,
  avg_rating decimal(2,1) default 0,
  
  -- Metadata
  created_at timestamptz default now(),
  updated_at timestamptz default now(),
  
  -- Indexes
  CONSTRAINT idx_ai_templates_public ON (is_public, category),
  CONSTRAINT idx_ai_templates_user ON user_id,
  CONSTRAINT idx_ai_templates_usage ON usage_count
);
6.5 System Metrics (Monitoring)
sqlCREATE TABLE system_metrics (
  id uuid default gen_random_uuid() primary key,
  metric_name text not null,
  metric_value jsonb not null,
  timestamp timestamptz default now(),
  
  -- Automatic partitioning
  PARTITION BY RANGE (timestamp)
);

-- Common metrics stored:
-- - daily_costs: { ai_cost, infra_cost, total_cost }
-- - daily_revenue: { mrr, arr, churn_rate }
-- - performance: { avg_response_time, uptime, error_rate }
-- - usage_patterns: { peak_hours, popular_features }
6.6 Database Functions (Performance Optimized)
Fast Usage Checking
sqlCREATE OR REPLACE FUNCTION check_usage_limit(user_uuid uuid)
RETURNS jsonb
LANGUAGE sql
SECURITY DEFINER
STABLE
AS $$
  SELECT jsonb_build_object(
    'can_query', 
    CASE 
      WHEN p.status != 'active' THEN false
      WHEN p.queries_used >= p.queries_limit THEN false
      WHEN p.current_period_end < now() AND p.plan != 'free' THEN false
      ELSE true
    END,
    'usage', jsonb_build_object(
      'current', p.queries_used,
      'limit', p.queries_limit,
      'plan', p.plan,
      'reset_date', p.current_period_end
    )
  )
  FROM profiles p
  WHERE p.id = user_uuid;
$$;
Atomic Usage Increment
sqlCREATE OR REPLACE FUNCTION increment_usage(
  user_uuid uuid,
  query_type text,
  tokens_used integer,
  cost decimal(8,4),
  response_time integer,
  from_cache boolean default false
)
RETURNS void
LANGUAGE plpgsql
SECURITY DEFINER
AS $$
BEGIN
  -- Update profile counter
  UPDATE profiles 
  SET 
    queries_used = queries_used + 1,
    last_active_at = now(),
    updated_at = now()
  WHERE id = user_uuid;
  
  -- Update daily aggregates
  INSERT INTO usage_daily (
    user_id, 
    queries_total, 
    tokens_total, 
    cost_total,
    avg_response_time,
    cache_hit_rate
  )
  VALUES (
    user_uuid, 
    1, 
    tokens_used, 
    cost,
    response_time,
    CASE WHEN from_cache THEN 1.0 ELSE 0.0 END
  )
  ON CONFLICT (user_id, date)
  DO UPDATE SET
    queries_total = usage_daily.queries_total + 1,
    tokens_total = usage_daily.tokens_total + tokens_used,
    cost_total = usage_daily.cost_total + cost,
    avg_response_time = (usage_daily.avg_response_time + response_time) / 2,
    cache_hit_rate = (usage_daily.cache_hit_rate * usage_daily.queries_total + 
                     CASE WHEN from_cache THEN 1.0 ELSE 0.0 END) / 
                     (usage_daily.queries_total + 1);
                     
  -- Update query type specific counter
  CASE query_type
    WHEN 'analysis' THEN
      UPDATE usage_daily SET queries_analysis = queries_analysis + 1
      WHERE user_id = user_uuid AND date = current_date;
    WHEN 'generation' THEN
      UPDATE usage_daily SET queries_generation = queries_generation + 1
      WHERE user_id = user_uuid AND date = current_date;
    WHEN 'explanation' THEN
      UPDATE usage_daily SET queries_explanation = queries_explanation + 1
      WHERE user_id = user_uuid AND date = current_date;
  END CASE;
END;
$$;
Monthly Usage Reset (Billing Cycle)
sqlCREATE OR REPLACE FUNCTION reset_monthly_usage()
RETURNS void
LANGUAGE plpgsql
SECURITY DEFINER
AS $$
BEGIN
  UPDATE profiles 
  SET 
    queries_used = 0,
    updated_at = now()
  WHERE current_period_end <= now()
    AND status = 'active';
END;
$$;

-- Schedule to run daily
SELECT cron.schedule('reset-monthly-usage', '0 1 * * *', 'SELECT reset_monthly_usage();');
6.7 Row Level Security (RLS) Policies
sql-- Enable RLS on all tables
ALTER TABLE profiles ENABLE ROW LEVEL SECURITY;
ALTER TABLE usage_daily ENABLE ROW LEVEL SECURITY;
ALTER TABLE ai_templates ENABLE ROW LEVEL SECURITY;
ALTER TABLE response_cache ENABLE ROW LEVEL SECURITY;

-- Profiles: Users can only access their own data
CREATE POLICY "users_own_profile" ON profiles
  FOR ALL USING (auth.uid() = id);

-- Usage: Users can only see their own usage
CREATE POLICY "users_own_usage" ON usage_daily
  FOR SELECT USING (auth.uid() = user_id);

-- Templates: Users can manage their own, view public ones
CREATE POLICY "users_own_templates" ON ai_templates
  FOR ALL USING (auth.uid() = user_id);

CREATE POLICY "public_templates_readable" ON ai_templates
  FOR SELECT USING (is_public = true);

-- Cache: Service role only (system managed)
CREATE POLICY "service_role_cache" ON response_cache
  USING (auth.role() = 'service_role');

🔌 7. API Specifications {#api-specifications}
7.1 Authentication Endpoints
POST /auth/register
Register new user with email and password
Request Body:
json{
  "email": "user@example.com",
  "password": "securePassword123",
  "name": "John Doe"
}
Response (201):
json{
  "user": {
    "id": "uuid",
    "email": "user@example.com",
    "name": "John Doe",
    "created_at": "2024-01-01T00:00:00Z"
  },
  "session": {
    "access_token": "jwt_token",
    "refresh_token": "refresh_token",
    "expires_in": 3600
  },
  "profile": {
    "plan": "free",
    "queries_used": 0,
    "queries_limit": 10,
    "trial_ends_at": "2024-01-15T00:00:00Z"
  }
}
POST /auth/login
Authenticate existing user
Request Body:
json{
  "email": "user@example.com",
  "password": "securePassword123"
}
Response (200):
json{
  "user": { /* user object */ },
  "session": { /* session object */ },
  "profile": { /* profile object */ }
}
7.2 AI Analysis Endpoints
POST /functions/v1/ai-analyze
Analyze Excel data using AI
Headers:
Authorization: Bearer {jwt_token}
Content-Type: application/json
Request Body:
json{
  "prompt": "What trends do you see in this sales data?",
  "context": "Month,Sales,Region\nJan,1000,North\nFeb,1200,North\n...",
  "options": {
    "stream": false,
    "model": "gpt-4-turbo",
    "max_tokens": 2000,
    "include_confidence": true
  }
}
Response (200):
json{
  "success": true,
  "response": {
    "summary": "Sales data shows consistent growth with seasonal patterns",
    "insights": [
      "Sales increased 20% from Jan to Feb",
      "North region outperforms South by 15%",
      "Q1 trending 25% above last year"
    ],
    "trends": [
      "Upward trajectory in all regions",
      "Seasonal spike in March"
    ],
    "recommendations": [
      "Increase inventory for March peak",
      "Focus marketing efforts on South region"
    ],
    "confidence_score": 0.89
  },
  "usage": {
    "queries_used": 15,
    "queries_limit": 500,
    "tokens_consumed": 1250,
    "model_used": "gpt-4-turbo",
    "cost": 0.0125
  },
  "metadata": {
    "processing_time_ms": 2400,
    "cached": false,
    "timestamp": "2024-01-01T12:00:00Z"
  }
}
Error Response (429 - Rate Limited):
json{
  "error": "Usage limit exceeded",
  "code": "USAGE_LIMIT_EXCEEDED",
  "details": {
    "current_usage": 500,
    "usage_limit": 500,
    "reset_date": "2024-02-01T00:00:00Z",
    "upgrade_url": "https://app.excelaix.com/upgrade"
  }
}
POST /functions/v1/ai-generate
Generate content for Excel
Request Body:
json{
  "prompt": "Generate 10 rows of sample customer data with name, email, phone, and purchase amount",
  "output_format": "excel",
  "options": {
    "structure": "table",
    "include_headers": true,
    "data_types": ["string", "email", "phone", "currency"]
  }
}
Response (200):
json{
  "success": true,
  "response": {
    "content": "Name\tEmail\tPhone\tPurchase Amount\nJohn Smith\tjohn@email.com\t555-0123\t$299.99\n...",
    "format": "tsv",
    "rows": 10,
    "columns": 4
  },
  "usage": { /* usage object */ },
  "metadata": { /* metadata object */ }
}
POST /functions/v1/ai-explain
Quick explanation of Excel data
Request Body:
json{
  "data": "Q1,100\nQ2,150\nQ3,120\nQ4,180",
  "context": "Quarterly sales figures"
}
Response (200):
json{
  "success": true,
  "response": {
    "explanation": "This appears to be quarterly sales data showing growth from Q1 to Q4, with a slight dip in Q3. Overall trend is positive with 80% growth year-over-year.",
    "key_points": [
      "Strong Q4 performance",
      "Q3 seasonal dip typical",
      "Consistent growth pattern"
    ]
  },
  "usage": { /* usage object */ }
}
7.3 Subscription Management
POST /functions/v1/stripe/create-checkout
Create Stripe checkout session
Request Body:
json{
  "plan": "professional",
  "billing_cycle": "monthly",
  "success_url": "https://app.excelaix.com/success",
  "cancel_url": "https://app.excelaix.com/pricing"
}
Response (200):
json{
  "success": true,
  "checkout_url": "https://checkout.stripe.com/pay/cs_...",
  "session_id": "cs_test_..."
}
POST /functions/v1/stripe/portal
Create customer portal session
Response (200):
json{
  "success": true,
  "portal_url": "https://billing.stripe.com/session/..."
}
GET /functions/v1/user/usage
Get current user usage and subscription info
Response (200):
json{
  "success": true,
  "subscription": {
    "plan": "professional",
    "status": "active",
    "current_period_start": "2024-01-01T00:00:00Z",
    "current_period_end": "2024-02-01T00:00:00Z",
    "cancel_at_period_end": false
  },
  "usage": {
    "current_period": {
      "queries_used": 127,
      "queries_limit": 500,
      "percentage_used": 25.4
    },
    "this_month": {
      "total_queries": 127,
      "analysis_queries": 89,
      "generation_queries": 28,
      "explanation_queries": 10
    },
    "daily_average": 4.2,
    "most_active_day": "Tuesday"
  },
  "costs": {
    "this_month": {
      "total_cost": 1.47,
      "ai_cost": 1.35,
      "infrastructure_cost": 0.12
    },
    "cost_per_query": 0.012
  }
}
7.4 Templates & Settings
GET /rest/v1/ai_templates
Get AI templates (public + user's private)
Query Parameters:

category: Filter by category
public: Boolean, include public templates
limit: Number of results (default: 20)
offset: Pagination offset

Response (200):
json{
  "success": true,
  "templates": [
    {
      "id": "uuid",
      "name": "Sales Trend Analysis",
      "description": "Analyze sales data for trends and patterns",
      "category": "sales",
      "prompt_template": "Analyze this sales data for trends: {data}",
      "is_public": true,
      "usage_count": 245,
      "avg_rating": 4.7,
      "created_by": "System"
    }
  ],
  "pagination": {
    "total": 156,
    "limit": 20,
    "offset": 0,
    "has_more": true
  }
}
POST /rest/v1/ai_templates
Create new AI template
Request Body:
json{
  "name": "Customer Segmentation Analysis",
  "description": "Segment customers based on purchase behavior",
  "category": "marketing",
  "prompt_template": "Segment these customers based on {criteria}: {data}",
  "is_public": false
}

💻 8. Frontend Implementation {#frontend-implementation}
8.1 Excel Plugin Architecture
Project Structure
excel-plugin/
├── manifest.xml              # Office Add-in manifest
├── src/
│   ├── index.html            # Main plugin interface
│   ├── taskpane/
│   │   ├── taskpane.html     # Task pane UI
│   │   ├── taskpane.js       # Main plugin logic
│   │   └── taskpane.css      # Plugin styling
│   ├── services/
│   │   ├── auth.service.js   # Supabase authentication
│   │   ├── ai.service.js     # AI API integration
│   │   └── excel.service.js  # Excel operations
│   ├── components/
│   │   ├── auth-panel.js     # Login/register UI
│   │   ├── analysis-panel.js # AI analysis interface
│   │   ├── generate-panel.js # Content generation
│   │   └── settings-panel.js # User settings
│   └── utils/
│       ├── constants.js      # App constants
│       ├── helpers.js        # Utility functions
│       └── cache.js          # Local caching
└── assets/
    ├── icons/               # Plugin icons (16px, 32px, 80px)
    └── images/             # UI images
Manifest Configuration
File: manifest.xml
xml<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
           xsi:type="TaskPaneApp">
    
    <Id>excel-ai-assistant-pro</Id>
    <Version>1.0.0.0</Version>
    <ProviderName>ExcelAI Solutions</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    
    <DisplayName DefaultValue="AI Assistant for Excel" />
    <Description DefaultValue="Transform your Excel data with AI-powered analysis and content generation" />
    <IconUrl DefaultValue="https://app.excelaix.com/assets/icons/icon-32.png" />
    <HighResolutionIconUrl DefaultValue="https://app.excelaix.com/assets/icons/icon-64.png" />
    <SupportUrl DefaultValue="https://excelaix.com/support" />
    
    <Hosts>
        <Host Name="Workbook" />
    </Hosts>
    
    <DefaultSettings>
        <SourceLocation DefaultValue="https://app.excelaix.com/excel-plugin/taskpane.html" />
    </DefaultSettings>
    
    <Permissions>ReadWriteDocument</Permissions>
    
    <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
        <Requirements>
            <Sets DefaultMinVersion="1.1">
                <Set Name="ExcelApi" MinVersion="1.9"/>
            </Sets>
        </Requirements>
        
        <Hosts>
            <Host xsi:type="Workbook">
                <DesktopFormFactor>
                    <GetStarted>
                        <Title resid="GetStarted.Title"/>
                        <Description resid="GetStarted.Description"/>
                        <LearnMoreUrl resid="GetStarted.LearnMoreUrl"/>
                    </GetStarted>
                    
                    <ExtensionPoint xsi:type="PrimaryCommandSurface">
                        <OfficeTab id="TabHome">
                            <Group id="AIAssistantGroup">
                                <Label resid="AIAssistantGroup.Label" />
                                <Icon>
                                    <bt:Image size="16" resid="Icon.16x16"/>
                                    <bt:Image size="32" resid="Icon.32x32"/>
                                    <bt:Image size="80" resid="Icon.80x80"/>
                                </Icon>
                                
                                <Control xsi:type="Button" id="OpenTaskPane">
                                    <Label resid="TaskpaneButton.Label" />
                                    <Supertip>
                                        <Title resid="TaskpaneButton.Label" />
                                        <Description resid="TaskpaneButton.Tooltip" />
                                    </Supertip>
                                    <Icon>
                                        <bt:Image size="16" resid="Icon.16x16"/>
                                        <bt:Image size="32" resid="Icon.32x32"/>
                                        <bt:Image size="80" resid="Icon.80x80"/>
                                    </Icon>
                                    <Action xsi:type="ShowTaskpane">
                                        <TaskpaneId>AIAssistantTaskPane</TaskpaneId>
                                        <SourceLocation resid="Taskpane.Url"/>
                                    </Action>
                                </Control>
                                
                                <Control xsi:type="Button" id="QuickAnalyze">
                                    <Label resid="QuickAnalyze.Label" />
                                    <Supertip>
                                        <Title resid="QuickAnalyze.Label" />
                                        <Description resid="QuickAnalyze.Tooltip" />
                                    </Supertip>
                                    <Icon>
                                        <bt:Image size="16" resid="Icon.16x16"/>
                                        <bt:Image size="32" resid="Icon.32x32"/>
                                        <bt:Image size="80" resid="Icon.80x80"/>
                                    </Icon>
                                    <Action xsi:type="ExecuteFunction">
                                        <FunctionName>quickAnalyze</FunctionName>
                                    </Action>
                                </Control>
                            </Group>
                        </OfficeTab>
                    </ExtensionPoint>
                </DesktopFormFactor>
            </Host>
        </Hosts>
        
        <Resources>
            <bt:Images>
                <bt:Image id="Icon.16x16" DefaultValue="https://app.excelaix.com/assets/icons/icon-16.png"/>
                <bt:Image id="Icon.32x32" DefaultValue="https://app.excelaix.com/assets/icons/icon-32.png"/>
                <bt:Image id="Icon.80x80" DefaultValue="https://app.excelaix.com/assets/icons/icon-80.png"/>
            </bt:Images>
            <bt:Urls>
                <bt:Url id="GetStarted.LearnMoreUrl" DefaultValue="https://excelaix.com/getting-started"/>
                <bt:Url id="Taskpane.Url" DefaultValue="https://app.excelaix.com/excel-plugin/taskpane.html"/>
            </bt:Urls>
            <bt:ShortStrings>
                <bt:String id="GetStarted.Title" DefaultValue="Get started with AI Assistant"/>
                <bt:String id="AIAssistantGroup.Label" DefaultValue="AI Assistant"/>
                <bt:String id="TaskpaneButton.Label" DefaultValue="AI Assistant"/>
                <bt:String id="QuickAnalyze.Label" DefaultValue="Quick Analyze"/>
            </bt:ShortStrings>
            <bt:LongStrings>
                <bt:String id="GetStarted.Description" DefaultValue="Use AI to analyze your Excel data instantly"/>
                <bt:String id="TaskpaneButton.Tooltip" DefaultValue="Open AI Assistant panel for data analysis"/>
                <bt:String id="QuickAnalyze.Tooltip" DefaultValue="Quickly analyze selected data with AI"/>
            </bt:LongStrings>
        </Resources>
    </VersionOverrides>
</OfficeApp>
Main Plugin Interface
File: src/taskpane/taskpane.html
html<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>AI Assistant for Excel</title>
    
    <!-- Office.js -->
    <script src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>
    
    <!-- Supabase -->
    <script src="https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2"></script>
    
    <!-- Styles -->
    <link rel="stylesheet" href="taskpane.css">
</head>

<body class="ms-welcome">
    <div id="app-container">
        <!-- Loading Screen -->
        <div id="loading-screen" class="loading-screen">
            <div class="loading-spinner"></div>
            <p>Initializing AI Assistant...</p>
        </div>
        
        <!-- Authentication Panel -->
        <div id="auth-panel" class="panel" style="display: none;">
            <div class="auth-header">
                <h1>🤖 AI Assistant</h1>
                <p>Sign in to unlock AI-powered Excel analysis</p>
            </div>
            
            <div class="auth-tabs">
                <button id="login-tab" class="tab-button active">Sign In</button>
                <button id="register-tab" class="tab-button">Create Account</button>
            </div>
            
            <!-- Login Form -->
            <form id="login-form" class="auth-form">
                <div class="form-group">
                    <label for="login-email">Email</label>
                    <input type="email" id="login-email" required>
                </div>
                <div class="form-group">
                    <label for="login-password">Password</label>
                    <input type="password" id="login-password" required>
                </div>
                <button type="submit" class="btn btn-primary">Sign In</button>
                <a href="#" id="forgot-password">Forgot Password?</a>
            </form>
            
            <!-- Register Form -->
            <form id="register-form" class="auth-form" style="display: none;">
                <div class="form-group">
                    <label for="register-name">Full Name</label>
                    <input type="text" id="register-name" required>
                </div>
                <div class="form-group">
                    <label for="register-email">Email</label>
                    <input type="email" id="register-email" required>
                </div>
                <div class="form-group">
                    <label for="register-password">Password</label>
                    <input type="password" id="register-password" required minlength="6">
                </div>
                <button type="submit" class="btn btn-primary">Create Account</button>
                <p class="trial-info">✨ 14-day free trial, no credit card required</p>
            </form>
            
            <div class="auth-footer">
                <p><a href="https://excelaix.com/privacy">Privacy Policy</a> | 
                   <a href="https://excelaix.com/terms">Terms of Service</a></p>
            </div>
        </div>
        
        <!-- Main Application -->
        <div id="main-app" class="panel" style="display: none;">
            <!-- Header -->
            <div class="app-header">
                <div class="user-info">
                    <span id="user-name">Welcome!</span>
                    <button id="user-menu-btn" class="btn-icon">⚙️</button>
                </div>
                <div class="usage-indicator">
                    <span id="usage-text">0/10 queries</span>
                    <div class="usage-bar">
                        <div id="usage-fill" class="usage-fill"></div>
                    </div>
                </div>
            </div>
            
            <!-- Mode Selector -->
            <div class="mode-selector">
                <button id="analyze-mode-btn" class="mode-btn active" data-mode="analyze">
                    📊 Analyze Data
                </button>
                <button id="generate-mode-btn" class="mode-btn" data-mode="generate">
                    ✨ Generate Content
                </button>
                <button id="templates-mode-btn" class="mode-btn" data-mode="templates">
                    📋 Templates
                </button>
            </div>
            
            <!-- Analysis Mode -->
            <div id="analyze-mode" class="mode-panel">
                <div class="quick-templates">
                    <label>Quick Start:</label>
                    <select id="analysis-templates">
                        <option value="">Choose a template...</option>
                        <option value="trends">📈 Find trends and patterns</option>
                        <option value="summary">📋 Summarize my data</option>
                        <option value="insights">💡 Extract key insights</option>
                        <option value="anomalies">🔍 Detect anomalies</option>
                        <option value="compare">⚖️ Compare segments</option>
                        <option value="forecast">🔮 Generate forecasts</option>
                    </select>
                </div>
                
                <div class="prompt-section">
                    <label for="analysis-prompt">Ask AI about your data:</label>
                    <textarea 
                        id="analysis-prompt" 
                        placeholder="e.g., 'What trends do you see in my sales data?' or 'Are there any unusual patterns?'"
                        rows="3"></textarea>
                    <div class="prompt-tips">
                        <small>💡 Select your data first, then ask specific questions for best results</small>
                    </div>
                </div>
                
                <div class="action-buttons">
                    <button id="analyze-btn" class="btn btn-primary">
                        🔍 Analyze Selection
                    </button>
                    <button id="explain-btn" class="btn btn-secondary">
                        💬 Quick Explain
                    </button>
                </div>
                
                <div class="streaming-indicator" id="streaming-indicator" style="display: none;">
                    <div class="streaming-dots">
                        <span></span><span></span><span></span>
                    </div>
                    <span>AI is analyzing your data...</span>
                </div>
            </div>
            
            <!-- Generation Mode -->
            <div id="generate-mode" class="mode-panel" style="display: none;">
                <div class="quick-templates">
                    <label>Generate:</label>
                    <select id="generation-templates">
                        <option value="">Choose what to generate...</option>
                        <option value="sample-data">📊 Sample data</option>
                        <option value="formulas">⚡ Excel formulas</option>
                        <option value="reports">📄 Data reports</option>
                        <option value="charts">📈 Chart suggestions</option>
                        <option value="fill-pattern">🔄 Fill data pattern</option>
                    </select>
                </div>
                
                <div class="prompt-section">
                    <label for="generation-prompt">Describe what you want to generate:</label>
                    <textarea 
                        id="generation-prompt" 
                        placeholder="e.g., 'Generate 50 rows of customer data with names, emails, and purchase amounts'"
                        rows="3"></textarea>
                </div>
                
                <div class="generation-options">
                    <div class="option-group">
                        <label>Output Format:</label>
                        <select id="output-format">
                            <option value="table">Table (rows & columns)</option>
                            <option value="list">List (single column)</option>
                            <option value="formulas">Excel formulas</option>
                        </select>
                    </div>
                </div>
                
                <div class="action-buttons">
                    <button id="generate-btn" class="btn btn-primary">
                        ✨ Generate Content
                    </button>
                    <button id="smart-fill-btn" class="btn btn-secondary">
                        🎯 Smart Fill Pattern
                    </button>
                </div>
            </div>
            
            <!-- Templates Mode -->
            <div id="templates-mode" class="mode-panel" style="display: none;">
                <div class="templates-header">
                    <h3>AI Templates</h3>
                    <button id="create-template-btn" class="btn btn-secondary">+ Create New</button>
                </div>
                
                <div class="template-categories">
                    <button class="category-btn active" data-category="all">All</button>
                    <button class="category-btn" data-category="analysis">Analysis</button>
                    <button class="category-btn" data-category="generation">Generation</button>
                    <button class="category-btn" data-category="financial">Financial</button>
                    <button class="category-btn" data-category="marketing">Marketing</button>
                </div>
                
                <div id="templates-list" class="templates-list">
                    <!-- Templates loaded dynamically -->
                </div>
            </div>
            
            <!-- Results Section -->
            <div id="results-section" class="results-section" style="display: none;">
                <div class="results-header">
                    <h3>📊 Results</h3>
                    <div class="results-actions">
                        <button id="copy-results-btn" class="btn btn-secondary">📋 Copy</button>
                        <button id="insert-results-btn" class="btn btn-primary">📥 Insert to Excel</button>
                        <button id="export-results-btn" class="btn btn-secondary">💾 Export</button>
                    </div>
                </div>
                
                <div id="results-content" class="results-content">
                    <!-- Results content loaded here -->
                </div>
                
                <div class="results-metadata">
                    <span id="processing-time">Processing time: --</span>
                    <span id="tokens-used">Tokens: --</span>
                    <span id="model-used">Model: --</span>
                </div>
            </div>
        </div>
        
        <!-- Upgrade Modal -->
        <div id="upgrade-modal" class="modal" style="display: none;">
            <div class="modal-content">
                <div class="modal-header">
                    <h3>🚀 Upgrade Your Plan</h3>
                    <button id="close-modal" class="btn-close">&times;</button>
                </div>
                <div class="modal-body">
                    <div class="upgrade-reason">
                        <p id="upgrade-message">You've reached your usage limit for this month.</p>
                    </div>
                    
                    <div class="pricing-cards">
                        <div class="pricing-card">
                            <h4>Starter</h4>
                            <div class="price">$12<span>/month</span></div>
                            <ul>
                                <li>150 AI queries/month</li>
                                <li>GPT-4 access (25 queries)</li>
                                <li>Email support</li>
                            </ul>
                            <button class="btn btn-primary" onclick="upgradeToplan('starter')">
                                Choose Starter
                            </button>
                        </div>
                        
                        <div class="pricing-card featured">
                            <div class="badge">Most Popular</div>
                            <h4>Professional</h4>
                            <div class="price">$29<span>/month</span></div>
                            <ul>
                                <li>500 AI queries/month</li>
                                <li>GPT-4 access (150 queries)</li>
                                <li>Priority support</li>
                                <li>Custom templates</li>
                            </ul>
                            <button class="btn btn-primary" onclick="upgradeToPlan('professional')">
                                Choose Professional
                            </button>
                        </div>
                        
                        <div class="pricing-card">
                            <h4>Business</h4>
                            <div class="price">$79<span>/month</span></div>
                            <ul>
                                <li>2000 AI queries/month</li>
                                <li>GPT-4 access (800 queries)</li>
                                <li>Team features</li>
                                <li>Phone support</li>
                            </ul>
                            <button class="btn btn-primary" onclick="upgradeToPlan('business')">
                                Choose Business
                            </button>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        
        <!-- Settings Panel -->
        <div id="settings-panel" class="side-panel" style="display: none;">
            <div class="panel-header">
                <h3>⚙️ Settings</h3>
                <button id="close-settings" class="btn-close">&times;</button>
            </div>
            
            <div class="panel-content">
                <div class="setting-group">
                    <h4>Account</h4>
                    <div class="setting-item">
                        <label>Name:</label>
                        <input type="text" id="user-name-input" />
                    </div>
                    <div class="setting-item">
                        <label>Email:</label>
                        <input type="email" id="user-email-input" />
                    </div>
                    <button id="save-profile" class="btn btn-secondary">Save Changes</button>
                </div>
                
                <div class="setting-group">
                    <h4>Preferences</h4>
                    <div class="setting-item">
                        <label>Preferred AI Model:</label>
                        <select id="preferred-model">
                            <option value="gpt-4-turbo">GPT-4 Turbo (Recommended)</option>
                            <option value="gpt-3.5-turbo">GPT-3.5 Turbo (Faster)</option>
                        </select>
                    </div>
                    <div class="setting-item">
                        <label>Response Style:</label>
                        <select id="response-style">
                            <option value="detailed">Detailed Analysis</option>
                            <option value="concise">Concise Summary</option>
                            <option value="bullet">Bullet Points</option>
                        </select>
                    </div>
                </div>
                
                <div class="setting-group">
                    <h4>Subscription</h4>
                    <div class="subscription-info">
                        <p><strong>Current Plan:</strong> <span id="current-plan">Loading...</span></p>
                        <p><strong>Usage:</strong> <span id="usage-details">Loading...</span></p>
                        <p><strong>Billing:</strong> <span id="billing-info">Loading...</span></p>
                    </div>
                    <div class="subscription-actions">
                        <button id="manage-billing" class="btn btn-secondary">Manage Billing</button>
                        <button id="upgrade-plan" class="btn btn-primary">Upgrade Plan</button>
                    </div>
                </div>
                
                <div class="setting-group">
                    <h4>Support</h4>
                    <div class="support-links">
                        <a href="https://excelaix.com/docs" target="_blank">📚 Documentation</a>
                        <a href="https://excelaix.com/support" target="_blank">💬 Contact Support</a>
                        <a href="https://excelaix.com/feedback" target="_blank">💡 Send Feedback</a>
                    </div>
                </div>
                
                <div class="setting-group">
                    <button id="sign-out" class="btn btn-danger">Sign Out</button>
                </div>
            </div>
        </div>
        
        <!-- Toast Notifications -->
        <div id="toast-container" class="toast-container"></div>
        
        <!-- Error Messages -->
        <div id="error-banner" class="error-banner" style="display: none;">
            <span id="error-message"></span>
            <button id="close-error" class="btn-close">&times;</button>
        </div>
    </div>
    
    <!-- Scripts -->
    <script src="services/auth.service.js"></script>
    <script src="services/ai.service.js"></script>
    <script src="services/excel.service.js"></script>
    <script src="components/auth-panel.js"></script>
    <script src="components/analysis-panel.js"></script>
    <script src="components/generate-panel.js"></script>
    <script src="components/templates-panel.js"></script>
    <script src="utils/helpers.js"></script>
    <script src="taskpane.js"></script>
</body>
</html>
Core Plugin Logic
File: src/taskpane/taskpane.js
javascript// Main application controller
class ExcelAIApp {
    constructor() {
        this.currentUser = null;
        this.currentMode = 'analyze';
        this.isStreaming = false;
        this.usage = { used: 0, limit: 10 };
        
        // Initialize services
        this.authService = new AuthService();
        this.aiService = new AIService();
        this.excelService = new ExcelService();
        
        this.init();
    }
    
    async init() {
        try {
            // Wait for Office.js to load
            await new Promise((resolve) => {
                Office.onReady(() => resolve());
            });
            
            // Initialize UI
            this.setupEventListeners();
            this.hideLoadingScreen();
            
            // Check authentication
            const user = await this.authService.getCurrentUser();
            if (user) {
                await this.handleUserAuthenticated(user);
            } else {
                this.showPanel('auth-panel');
            }
            
        } catch (error) {
            console.error('App initialization error:', error);
            this.showError('Failed to initialize app. Please refresh and try again.');
        }
    }
    
    setupEventListeners() {
        // Authentication
        document.getElementById('login-form').addEventListener('submit', (e) => this.handleLogin(e));
        document.getElementById('register-form').addEventListener('submit', (e) => this.handleRegister(e));
        document.getElementById('login-tab').addEventListener('click', () => this.switchAuthTab('login'));
        document.getElementById('register-tab').addEventListener('click', () => this.switchAuthTab('register'));
        
        // Mode switching
        document.querySelectorAll('.mode-btn').forEach(btn => {
            btn.addEventListener('click', (e) => this.switchMode(e.target.dataset.mode));
        });
        
        // Analysis mode
        document.getElementById('analyze-btn').addEventListener('click', () => this.analyzeData());
        document.getElementById('explain-btn').addEventListener('click', () => this.quickExplain());
        document.getElementById('analysis-templates').addEventListener('change', (e) => this.loadAnalysisTemplate(e.target.value));
        
        // Generation mode
        document.getElementById('generate-btn').addEventListener('click', () => this.generateContent());
        document.getElementById('smart-fill-btn').addEventListener('click', () => this.smartFill());
        document.getElementById('generation-templates').addEventListener('change', (e) => this.loadGenerationTemplate(e.target.value));
        
        // Results actions
        document.getElementById('copy-results-btn').addEventListener('click', () => this.copyResults());
        document.getElementById('insert-results-btn').addEventListener('click', () => this.insertResults());
        document.getElementById('export-results-btn').addEventListener('click', () => this.exportResults());
        
        // Settings and account
        document.getElementById('user-menu-btn').addEventListener('click', () => this.toggleSettings());
        document.getElementById('close-settings').addEventListener('click', () => this.hideSettings());
        document.getElementById('sign-out').addEventListener('click', () => this.signOut());
        document.getElementById('manage-billing').addEventListener('click', () => this.manageBilling());
        
        // Modal controls
        document.getElementById('close-modal').addEventListener('click', () => this.hideModal());
        document.getElementById('close-error').addEventListener('click', () => this.hideError());
        
        // Templates
        document.getElementById('create-template-btn').addEventListener('click', () => this.createTemplate());
        
        // Auto-save prompt content
        ['analysis-prompt', 'generation-prompt'].forEach(id => {
            document.getElementById(id).addEventListener('input', (e) => {
                localStorage.setItem(`excelaix_${id}`, e.target.value);
            });
        });
        
        // Load saved prompts
        ['analysis-prompt', 'generation-prompt'].forEach(id => {
            const saved = localStorage.getItem(`excelaix_${id}`);
            if (saved) {
                document.getElementById(id).value = saved;
            }
        });
    }
    
    async handleLogin(e) {
        e.preventDefault();
        
        try {
            this.showLoading('Signing in...');
            
            const email = document.getElementById('login-email').value;
            const password = document.getElementById('login-password').value;
            
            const user = await this.authService.signIn(email, password);
            await this.handleUserAuthenticated(user);
            
        } catch (error) {
            this.showError(error.message);
        } finally {
            this.hideLoading();
        }
    }
    
    async handleRegister(e) {
        e.preventDefault();
        
        try {
            this.showLoading('Creating account...');
            
            const name = document.getElementById('register-name').value;
            const email = document.getElementById('register-email').value;
            const password = document.getElementById('register-password').value;
            
            const user = await this.authService.signUp(email, password, name);
            await this.handleUserAuthenticated(user);
            
            this.showToast('Account created successfully! You have 14 days free trial.', 'success');
            
        } catch (error) {
            this.showError(error.message);
        } finally {
            this.hideLoading();
        }
    }
    
    async handleUserAuthenticated(user) {
        this.currentUser = user;
        
        // Load user profile and usage
        const profile = await this.authService.getProfile();
        this.usage = {
            used: profile.queries_used,
            limit: profile.queries_limit,
            plan: profile.plan
        };
        
        // Update UI
        document.getElementById('user-name').textContent = `Welcome, ${user.name}!`;
        this.updateUsageDisplay();
        
        // Show main app
        this.showPanel('main-app');
        
        // Load templates
        await this.loadTemplates();
    }
    
    async analyzeData() {
        if (!await this.checkUsageLimit()) return;
        
        try {
            const prompt = document.getElementById('analysis-prompt').value.trim();
            if (!prompt) {
                this.showError('Please enter a prompt for analysis.');
                return;
            }
            
            // Get selected Excel data
            const selectedData = await this.excelService.getSelectedData();
            if (!selectedData || selectedData.length === 0) {
                this.showError('Please select some data in Excel first.');
                return;
            }
            
            this.showStreamingIndicator(true);
            this.showResultsSection();
            
            // Call AI service with streaming
            const startTime = Date.now();
            const result = await this.aiService.analyzeData(prompt, selectedData, {
                stream: true,
                onProgress: (partialResponse) => {
                    this.updateResults(partialResponse, true);
                }
            });
            
            // Update final results
            this.updateResults(result.response, false);
            this.updateUsageDisplay(result.usage);
            this.updateResultsMetadata({
                processingTime: Date.now() - startTime,
                tokensUsed: result.usage.tokens_consumed,
                model: result.usage.model_used
            });
            
            // Save to local history
            this.saveToHistory('analysis', prompt, result.response);
            
        } catch (error) {
            this.handleAIError(error);
        } finally {
            this.showStreamingIndicator(false);
        }
    }
    
    async generateContent() {
        if (!await this.checkUsageLimit()) return;
        
        try {
            const prompt = document.getElementById('generation-prompt').value.trim();
            const outputFormat = document.getElementById('output-format').value;
            
            if (!prompt) {
                this.showError('Please describe what you want to generate.');
                return;
            }
            
            this.showStreamingIndicator(true);
            
            const result = await this.aiService.generateContent(prompt, {
                outputFormat,
                stream: true,
                onProgress: (partialResponse) => {
                    this.updateResults(partialResponse, true);
                }
            });
            
            // If it's structured data, offer to insert directly
            if (result.format === 'table' || outputFormat === 'table') {
                this.showInsertOption(result.response);
            }
            
            this.updateResults(result.response, false);
            this.updateUsageDisplay(result.usage);
            this.saveToHistory('generation', prompt, result.response);
            
        } catch (error) {
            this.handleAIError(error);
        } finally {
            this.showStreamingIndicator(false);
        }
    }
    
    async quickExplain() {
        if (!await this.checkUsageLimit()) return;
        
        try {
            const selectedData = await this.excelService.getSelectedData();
            if (!selectedData || selectedData.length === 0) {
                this.showError('Please select some data in Excel first.');
                return;
            }
            
            this.showStreamingIndicator(true);
            
            const result = await this.aiService.explainData(selectedData);
            
            this.updateResults(result.response, false);
            this.updateUsageDisplay(result.usage);
            this.showResultsSection();
            
        } catch (error) {
            this.handleAIError(error);
        } finally {
            this.showStreamingIndicator(false);
        }
    }
    
    async checkUsageLimit() {
        const usage = await this.authService.checkUsage();
        
        if (!usage.canUse) {
            this.showUpgradeModal(usage);
            return false;
        }
        
        this.usage = usage;
        this.updateUsageDisplay();
        return true;
    }
    
    updateResults(content, isPartial = false) {
        const resultsContent = document.getElementById('results-content');
        
        try {
            // Try to parse as structured response
            const parsed = typeof content === 'string' ? JSON.parse(content) : content;
            
            if (parsed.summary || parsed.insights) {
                // Structured analysis response
                resultsContent.innerHTML = `
                    <div class="analysis-result ${isPartial ? 'streaming' : ''}">
                        ${parsed.summary ? `
                            <div class="result-section">
                                <h4>📊 Summary</h4>
                                <p>${parsed.summary}</p>
                            </div>
                        ` : ''}
                        
                        ${parsed.insights && parsed.insights.length > 0 ? `
                            <div class="result-section">
                                <h4>💡 Key Insights</h4>
                                <ul>
                                    ${parsed.insights.map(insight => `<li>${insight}</li>`).join('')}
                                </ul>
                            </div>
                        ` : ''}
                        
                        ${parsed.trends && parsed.trends.length > 0 ? `
                            <div class="result-section">
                                <h4>📈 Trends</h4>
                                <ul>
                                    ${parsed.trends.map(trend => `<li>${trend}</li>`).join('')}
                                </ul>
                            </div>
                        ` : ''}
                        
                        ${parsed.recommendations && parsed.recommendations.length > 0 ? `
                            <div class="result-section">
                                <h4>🎯 Recommendations</h4>
                                <ul>
                                    ${parsed.recommendations.map(rec => `<li>${rec}</li>`).join('')}
                                </ul>
                            </div>
                        ` : ''}
                        
                        ${parsed.confidence_score ? `
                            <div class="confidence-section">
                                <div class="confidence-bar">
                                    <div class="confidence-fill" style="width: ${parsed.confidence_score * 100}%"></div>
                                </div>
                                <span class="confidence-text">Confidence: ${Math.round(parsed.confidence_score * 100)}%</span>
                            </div>
                        ` : ''}
                    </div>
                `;
            } else {
                // Plain text response
                resultsContent.innerHTML = `
                    <div class="text-result ${isPartial ? 'streaming' : ''}">
                        <pre>${content}</pre>
                    </div>
                `;
            }
        } catch (e) {
            // Fallback to plain text
            resultsContent.innerHTML = `
                <div class="text-result ${isPartial ? 'streaming' : ''}">
                    <pre>${content}</pre>
                </div>
            `;
        }
        
        if (!isPartial) {
            resultsContent.classList.remove('streaming');
        }
    }
    
    updateUsageDisplay(usage = this.usage) {
        const usageText = document.getElementById('usage-text');
        const usageFill = document.getElementById('usage-fill');
        
        usageText.textContent = `${usage.used}/${usage.limit} queries`;
        
        const percentage = (usage.used / usage.limit) * 100;
        usageFill.style.width = `${Math.min(percentage, 100)}%`;
        
        // Color coding
        if (percentage >= 90) {
            usageFill.className = 'usage-fill danger';
        } else if (percentage >= 70) {
            usageFill.className = 'usage-fill warning';
        } else {
            usageFill.className = 'usage-fill normal';
        }
        
        this.usage = usage;
    }
    
    showUpgradeModal(usage) {
        const modal = document.getElementById('upgrade-modal');
        const message = document.getElementById('upgrade-message');
        
        if (usage.plan === 'free') {
            message.textContent = `You've used all ${usage.limit} free queries. Upgrade to continue using AI features.`;
        } else {
            message.textContent = `You've reached your monthly limit of ${usage.limit} queries. Upgrade for more capacity.`;
        }
        
        modal.style.display = 'flex';
    }
    
    async upgradeToPlan(plan) {
        try {
            const checkoutUrl = await this.authService.createCheckoutSession(plan);
            window.open(checkoutUrl, '_blank');
            this.hideModal();
        } catch (error) {
            this.showError('Failed to start upgrade process. Please try again.');
        }
    }
    
    async manageBilling() {
        try {
            const portalUrl = await this.authService.createPortalSession();
            window.open(portalUrl, '_blank');
        } catch (error) {
            this.showError('Failed to open billing portal. Please try again.');
        }
    }
    
    // UI Helper Methods
    showPanel(panelId) {
        document.querySelectorAll('.panel').forEach(panel => {
            panel.style.display = 'none';
        });
        document.getElementById(panelId).style.display = 'block';
    }
    
    showResultsSection() {
        document.getElementById('results-section').style.display = 'block';
    }
    
    showStreamingIndicator(show) {
        document.getElementById('streaming-indicator').style.display = show ? 'flex' : 'none';
        this.isStreaming = show;
    }
    
    showError(message) {
        const errorBanner = document.getElementById('error-banner');
        const errorMessage = document.getElementById('error-message');
        
        errorMessage.textContent = message;
        errorBanner.style.display = 'flex';
        
        // Auto-hide after 5 seconds
        setTimeout(() => this.hideError(), 5000);
    }
    
    hideError() {
        document.getElementById('error-banner').style.display = 'none';
    }
    
    showToast(message, type = 'info') {
        const container = document.getElementById('toast-container');
        const toast = document.createElement('div');
        toast.className = `toast toast-${type}`;
        toast.textContent = message;
        
        container.appendChild(toast);
        
        // Animate in
        setTimeout(() => toast.classList.add('show'), 100);
        
        // Remove after 4 seconds
        setTimeout(() => {
            toast.classList.remove('show');
            setTimeout(() => container.removeChild(toast), 300);
        }, 4000);
    }
    
    hideLoadingScreen() {
        document.getElementById('loading-screen').style.display = 'none';
    }
    
    hideModal() {
        document.getElementById('upgrade-modal').style.display = 'none';
    }
    
    // Additional methods for templates, settings, etc...
    async loadTemplates() {
        // Implementation for loading AI templates
    }
    
    saveToHistory(type, prompt, response) {
        // Save to local storage for quick access
        const history = JSON.parse(localStorage.getItem('excelaix_history') || '[]');
        history.unshift({
            type,
            prompt,
            response: response.substring(0, 200) + '...',
            timestamp: new Date().toISOString()
        });
        
        // Keep only last 20 items
        if (history.length > 20) {
            history = history.slice(0, 20);
        }
        
        localStorage.setItem('excelaix_history', JSON.stringify(history));
    }
}

// Initialize app when DOM is loaded
document.addEventListener('DOMContentLoaded', () => {
    window.excelAIApp = new ExcelAIApp();
});
8.2 Service Layer Implementation
Authentication Service
File: src/services/auth.service.js
javascriptclass AuthService {
    constructor() {
        this.supabase = window.supabase.createClient(
            'https://your-project.supabase.co',
            'your-anon-key'
        );
        this.currentUser = null;
    }
    
    async signUp(email, password, name) {
        try {
            const { data, error } = await this.supabase.auth.signUp({
                email,
                password,
                options: {
                    data: { name }
                }
            });
            
            if (error) throw error;
            
            // Create profile
            if (data.user) {
                await this.supabase
                    .from('profiles')
                    .insert({
                        id: data.user.id,
                        plan: 'free',
                        queries_limit: 10
                    });
            }
            
            return data.user;
        } catch (error) {
            throw new Error(error.message || 'Registration failed');
        }
    }
    
    async signIn(email, password) {
        try {
            const { data, error } = await this.supabase.auth.signInWithPassword({
                email,
                password
            });
            
            if (error) throw error;
            
            this.currentUser = data.user;
            return data.user;
        } catch (error) {
            throw new Error(error.message || 'Sign in failed');
        }
    }
    
    async signOut() {
        try {
            const { error } = await this.supabase.auth.signOut();
            if (error) throw error;
            
            this.currentUser = null;
            
            // Clear local storage
            localStorage.removeItem('excelaix_history');
            localStorage.removeItem('excelaix_analysis-prompt');
            localStorage.removeItem('excelaix_generation-prompt');
            
        } catch (error) {
            throw new Error('Sign out failed');
        }
    }
    
    async getCurrentUser() {
        try {
            const { data: { user } } = await this.supabase.auth.getUser();
            this.currentUser = user;
            return user;
        } catch (error) {
            return null;
        }
    }
    
    async getProfile() {
        if (!this.currentUser) return null;
        
        try {
            const { data, error } = await this.supabase
                .from('profiles')
                .select('*')
                .eq('id', this.currentUser.id)
                .single();
            
            if (error) throw error;
            return data;
        } catch (error) {
            console.error('Failed to load profile:', error);
            return null;
        }
    }
    
    async checkUsage() {
        if (!this.currentUser) return { canUse: false, used: 0, limit: 0 };
        
        try {
            const { data } = await this.supabase
                .rpc('check_usage_limit', { user_uuid: this.currentUser.id });
            
            return data;
        } catch (error) {
            console.error('Failed to check usage:', error);
            return { canUse: false, used: 0, limit: 0 };
        }
    }
    
    async createCheckoutSession(plan) {
        try {
            const { data: { session } } = await this.supabase.auth.getSession();
            
            const response = await fetch('https://your-project.supabase.co/functions/v1/stripe-checkout', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': `Bearer ${session.access_token}`
                },
                body: JSON.stringify({ plan })
            });
            
            if (!response.ok) {
                throw new Error('Failed to create checkout session');
            }
            
            const { checkout_url } = await response.json();
            return checkout_url;
        } catch (error) {
            throw new Error('Failed to start upgrade process');
        }
    }
    
    async createPortalSession() {
        try {
            const { data: { session } } = await this.supabase.auth.getSession();
            
            const response = await fetch('https://your-project.supabase.co/functions/v1/stripe-portal', {
                method: 'POST',
                headers: {
                    'Authorization': `Bearer ${session.access_token}`
                }
            });
            
            if (!response.ok) {
                throw new Error('Failed to create portal session');
            }
            
            const { portal_url } = await response.json();
            return portal_url;
        } catch (error) {
            throw new Error('Failed to open billing portal');
        }
    }
}
AI Service
File: src/services/ai.service.js
javascriptclass AIService {
    constructor() {
        this.supabase = window.supabase.createClient(
            'https://your-project.supabase.co',
            'your-anon-key'
        );
        this.baseUrl = 'https://your-project.supabase.co/functions/v1';
    }
    
    async analyzeData(prompt, context, options = {}) {
        const { stream = false, onProgress } = options;
        
        try {
            const { data: { session } } = await this.supabase.auth.getSession();
            
            if (!session) {
                throw new Error('Please sign in to use AI features');
            }
            
            const response = await fetch(`${this.baseUrl}/ai-analyze`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': `Bearer ${session.access_token}`
                },
                body: JSON.stringify({
                    prompt,
                    context: this.formatContextData(context),
                    stream
                })
            });
            
            if (!response.ok) {
                const error = await response.json();
                throw new Error(error.error || 'Analysis failed');
            }
            
            if (stream) {
                return this.handleStreamingResponse(response, onProgress);
            } else {
                return await response.json();
            }
        } catch (error) {
            console.error('AI Analysis Error:', error);
            throw error;
        }
    }
    
    async generateContent(prompt, options = {}) {
        const { outputFormat = 'text', stream = false, onProgress } = options;
        
        try {
            const { data: { session } } = await this.supabase.auth.getSession();
            
            const response = await fetch(`${this.baseUrl}/ai-generate`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': `Bearer ${session.access_token}`
                },
                body: JSON.stringify({
                    prompt,
                    output_format: outputFormat,
                    stream
                })
            });
            
            if (!response.ok) {
                const error = await response.json();
                throw new Error(error.error || 'Generation failed');
            }
            
            if (stream) {
                return this.handleStreamingResponse(response, onProgress);
            } else {
                return await response.json();
            }
        } catch (error) {
            console.error('AI Generation Error:', error);
            throw error;
        }
    }
    
    async explainData(data) {
        try {
            const { data: { session } } = await this.supabase.auth.getSession();
            
            const response = await fetch(`${this.baseUrl}/ai-explain`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': `Bearer ${session.access_token}`
                },
                body: JSON.stringify({
                    data: this.formatContextData(data)
                })
            });
            
            if (!response.ok) {
                const error = await response.json();
                throw new Error(error.error || 'Explanation failed');
            }
            
            return await response.json();
        } catch (error) {
            console.error('AI Explain Error:', error);
            throw error;
        }
    }
    
    async handleStreamingResponse(response, onProgress) {
        const reader = response.body?.getReader();
        const decoder = new TextDecoder();
        let fullResponse = '';
        
        if (!reader) {
            throw new Error('Streaming not supported');
        }
        
        try {
            while (true) {
                const { done, value } = await reader.read();
                
                if (done) break;
                
                const chunk = decoder.decode(value);
                const lines = chunk.split('\n');
                
                for (const line of lines) {
                    if (line.startsWith('data: ')) {
                        const data = line.slice(6);
                        
                        if (data === '[DONE]') {
                            return { response: fullResponse, stream: true };
                        }
                        
                        try {
                            const parsed = JSON.parse(data);
                            if (parsed.content) {
                                fullResponse += parsed.content;
                                onProgress?.(fullResponse);
                            }
                        } catch (e) {
                            // Ignore invalid JSON chunks
                        }
                    }
                }
            }
        } finally {
            reader.releaseLock();
        }
        
        return { response: fullResponse, stream: true };
    }
    
    formatContextData(data) {
        if (typeof data === 'string') {
            return data;
        }
        
        if (Array.isArray(data)) {
            // Convert 2D array to CSV-like format
            return data.map(row => 
                Array.isArray(row) ? row.join('\t') : row
            ).join('\n');
        }
        
        return JSON.stringify(data);
    }
    
    // Template methods
    async getTemplates(category = 'all') {
        try {
            let query = this.supabase
                .from('ai_templates')
                .select('*')
                .or('is_public.eq.true,user_id.eq.' + (await this.getCurrentUserId()));
            
            if (category !== 'all') {
                query = query.eq('category', category);
            }
            
            const { data, error } = await query.order('usage_count', { ascending: false });
            
            if (error) throw error;
            return data;
        } catch (error) {
            console.error('Failed to load templates:', error);
            return [];
        }
    }
    
    async saveTemplate(template) {
        try {
            const userId = await this.getCurrentUserId();
            
            const { data, error } = await this.supabase
                .from('ai_templates')
                .insert({
                    ...template,
                    user_id: userId
                });
            
            if (error) throw error;
            return data;
        } catch (error) {
            console.error('Failed to save template:', error);
            throw error;
        }
    }
    
    async getCurrentUserId() {
        const { data: { user } } = await this.supabase.auth.getUser();
        return user?.id;
    }
}
Excel Service
File: src/services/excel.service.js
javascriptclass ExcelService {
    constructor() {
        this.maxCells = 50000; // Performance limit
    }
    
    async getSelectedData() {
        return new Promise((resolve, reject) => {
            Excel.run(async (context) => {
                try {
                    const range = context.workbook.getSelectedRange();
                    range.load(['values', 'address', 'rowCount', 'columnCount']);
                    await context.sync();
                    
                    // Check size limits
                    const totalCells = range.rowCount * range.columnCount;
                    if (totalCells > this.maxCells) {
                        reject(new Error(`Selection too large (${totalCells} cells). Please select fewer than ${this.maxCells} cells.`));
                        return;
                    }
                    
                    if (range.values.length === 0) {
                        reject(new Error('No data selected. Please select a range with data.'));
                        return;
                    }
                    
                    resolve({
                        values: range.values,
                        address: range.address,
                        rowCount: range.rowCount,
                        columnCount: range.columnCount
                    });
                } catch (error) {
                    reject(error);
                }
            });
        });
    }
    
    async insertData(data, startCell = null) {
        return new Promise((resolve, reject) => {
            Excel.run(async (context) => {
                try {
                    let range;
                    
                    if (startCell) {
                        // Insert at specific cell
                        range = context.workbook.worksheets.getActiveWorksheet().getRange(startCell);
                    } else {
                        // Insert at current selection
                        range = context.workbook.getSelectedRange();
                    }
                    
                    // Parse data into 2D array
                    const dataArray = this.parseDataForExcel(data);
                    
                    // Resize range to fit data
                    const resizedRange = range.getResizedRange(dataArray.length - 1, dataArray[0].length - 1);
                    resizedRange.values = dataArray;
                    
                    await context.sync();
                    resolve(resizedRange.address);
                } catch (error) {
                    reject(error);
                }
            });
        });
    }
    
    async createNewWorksheet(name = 'AI Analysis') {
        return new Promise((resolve, reject) => {
            Excel.run(async (context) => {
                try {
                    const worksheet = context.workbook.worksheets.add(name);
                    worksheet.activate();
                    await context.sync();
                    resolve(worksheet);
                } catch (error) {
                    reject(error);
                }
            });
        });
    }
    
    async insertAnalysisResults(results, createNewSheet = false) {
        return new Promise((resolve, reject) => {
            Excel.run(async (context) => {
                try {
                    let worksheet;
                    
                    if (createNewSheet) {
                        worksheet = context.workbook.worksheets.add('AI Analysis ' + new Date().toLocaleDateString());
                    } else {
                        worksheet = context.workbook.worksheets.getActiveWorksheet();
                    }
                    
                    // Insert results as formatted text
                    const range = worksheet.getRange('A1');
                    range.values = [['AI Analysis Results']];
                    range.format.font.bold = true;
                    range.format.font.size = 14;
                    
                    // Insert actual results
                    const resultsRange = worksheet.getRange('A3');
                    resultsRange.values = [[results]];
                    resultsRange.format.wrapText = true;
                    
                    // Auto-fit columns
                    worksheet.getUsedRange().format.autofitColumns();
                    
                    if (createNewSheet) {
                        worksheet.activate();
                    }
                    
                    await context.sync();
                    resolve(worksheet.name);
                } catch (error) {
                    reject(error);
                }
            });
        });
    }
    
    parseDataForExcel(data) {
        if (typeof data === 'string') {
            // Try to parse as TSV/CSV
            const lines = data.split('\n').filter(line => line.trim());
            
            // Check if it's tab-separated
            if (lines[0].includes('\t')) {
                return lines.map(line => line.split('\t'));
            }
            
            // Check if it's comma-separated
            if (lines[0].includes(',')) {
                return lines.map(line => line.split(','));
            }
            
            // Single column data
            return lines.map(line => [line]);
        }
        
        if (Array.isArray(data)) {
            return data;
        }
        
        // Fallback: convert to string
        return [[data.toString()]];
    }
    
    async exportToFile(data, filename = 'ai-analysis.txt') {
        try {
            const blob = new Blob([data], { type: 'text/plain' });
            const url = URL.createObjectURL(blob);
            
            const a = document.createElement('a');
            a.href = url;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            
            URL.revokeObjectURL(url);
        } catch (error) {
            console.error('Export failed:', error);
            throw new Error('Failed to export file');
        }
    }
    
    async copyToClipboard(data) {
        try {
            await navigator.clipboard.writeText(data);
            return true;
        } catch (error) {
            console.error('Copy failed:', error);
            
            // Fallback for older browsers
            const textArea = document.createElement('textarea');
            textArea.value = data;
            document.body.appendChild(textArea);
            textArea.select();
            const success = document.execCommand('copy');
            document.body.removeChild(textArea);
            
            if (!success) {
                throw new Error('Failed to copy to clipboard');
            }
            
            return true;
        }
    }
    
    // Utility methods for data formatting
    formatForExcel(data) {
        // Ensure data is in proper format for Excel
        if (typeof data === 'object' && data !== null) {
            return JSON.stringify(data, null, 2);
        }
        return data.toString();
    }
    
    detectDataType(values) {
        // Analyze data to suggest appropriate charts/analysis
        const flatValues = values.flat().filter(v => v !== null && v !== '');
        
        const hasNumbers = flatValues.some(v => !isNaN(v) && v !== '');
        const hasDates = flatValues.some(v => !isNaN(Date.parse(v)));
        const hasText = flatValues.some(v => isNaN(v) && isNaN(Date.parse(v)));
        
        return {
            hasNumbers,
            hasDates,
            hasText,
            rowCount: values.length,
            columnCount: values[0]?.length || 0,
            totalCells: values.length * (values[0]?.length || 0)
        };
    }
}

