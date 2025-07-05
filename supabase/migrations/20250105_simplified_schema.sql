-- Simplified Schema for ExcelAIRate
-- Only 4 core tables for MVP

-- Enable required extensions
CREATE EXTENSION IF NOT EXISTS "uuid-ossp";

-- 1. PROFILES TABLE (includes subscription data)
CREATE TABLE public.profiles (
  -- User identity
  id UUID REFERENCES auth.users(id) ON DELETE CASCADE PRIMARY KEY,
  email TEXT UNIQUE NOT NULL,
  full_name TEXT,
  
  -- Subscription (simplified)
  plan TEXT DEFAULT 'free' CHECK (plan IN ('free', 'pro', 'team')),
  stripe_customer_id TEXT UNIQUE,
  stripe_subscription_id TEXT UNIQUE,
  subscription_status TEXT DEFAULT 'active',
  current_period_end TIMESTAMPTZ,
  
  -- Usage (simplified)
  queries_used INTEGER DEFAULT 0 NOT NULL,
  queries_limit INTEGER DEFAULT 10 NOT NULL,
  
  -- Metadata
  created_at TIMESTAMPTZ DEFAULT NOW() NOT NULL,
  updated_at TIMESTAMPTZ DEFAULT NOW() NOT NULL
);

-- Simple indexes
CREATE INDEX idx_profiles_email ON profiles(email);
CREATE INDEX idx_profiles_stripe_customer ON profiles(stripe_customer_id);

-- 2. USAGE_LOGS TABLE (combined tracking)
CREATE TABLE public.usage_logs (
  id UUID DEFAULT uuid_generate_v4() PRIMARY KEY,
  user_id UUID REFERENCES auth.users(id) ON DELETE CASCADE NOT NULL,
  
  -- What happened
  action_type TEXT NOT NULL CHECK (action_type IN ('analyze', 'generate', 'explain')),
  prompt TEXT NOT NULL,
  response JSONB NOT NULL,
  
  -- Performance & cost
  model_used TEXT NOT NULL,
  tokens_used INTEGER DEFAULT 0,
  response_time_ms INTEGER,
  from_cache BOOLEAN DEFAULT false,
  
  -- When
  created_at TIMESTAMPTZ DEFAULT NOW() NOT NULL
);

-- Indexes for common queries
CREATE INDEX idx_usage_logs_user_id ON usage_logs(user_id);
CREATE INDEX idx_usage_logs_created_at ON usage_logs(created_at);
CREATE INDEX idx_usage_logs_user_date ON usage_logs(user_id, created_at);

-- 3. AI_CACHE TABLE (simplified caching)
CREATE TABLE public.ai_cache (
  prompt_hash TEXT PRIMARY KEY,
  prompt TEXT NOT NULL,
  response JSONB NOT NULL,
  model_used TEXT NOT NULL,
  hit_count INTEGER DEFAULT 0,
  created_at TIMESTAMPTZ DEFAULT NOW() NOT NULL,
  expires_at TIMESTAMPTZ DEFAULT (NOW() + INTERVAL '7 days') NOT NULL
);

-- Cache indexes
CREATE INDEX idx_ai_cache_expires ON ai_cache(expires_at);

-- 4. TEMPLATES TABLE (simplified, no vectors)
CREATE TABLE public.templates (
  id UUID DEFAULT uuid_generate_v4() PRIMARY KEY,
  user_id UUID REFERENCES auth.users(id) ON DELETE CASCADE,
  
  -- Template data
  name TEXT NOT NULL,
  description TEXT,
  prompt_template TEXT NOT NULL,
  category TEXT DEFAULT 'general',
  
  -- Sharing
  is_public BOOLEAN DEFAULT false,
  usage_count INTEGER DEFAULT 0,
  
  -- Metadata
  created_at TIMESTAMPTZ DEFAULT NOW() NOT NULL,
  updated_at TIMESTAMPTZ DEFAULT NOW() NOT NULL
);

-- Template indexes
CREATE INDEX idx_templates_user_id ON templates(user_id);
CREATE INDEX idx_templates_public ON templates(is_public);
-- Full text search index
CREATE INDEX idx_templates_search ON templates 
  USING gin(to_tsvector('english', name || ' ' || coalesce(description, '')));

-- ROW LEVEL SECURITY
ALTER TABLE profiles ENABLE ROW LEVEL SECURITY;
ALTER TABLE usage_logs ENABLE ROW LEVEL SECURITY;
ALTER TABLE templates ENABLE ROW LEVEL SECURITY;
ALTER TABLE ai_cache ENABLE ROW LEVEL SECURITY;

-- Profiles: Users can only see/edit their own
CREATE POLICY "Users can view own profile" ON profiles
  FOR SELECT USING (auth.uid() = id);
CREATE POLICY "Users can update own profile" ON profiles
  FOR UPDATE USING (auth.uid() = id);

-- Usage logs: Users can only see their own
CREATE POLICY "Users can view own usage" ON usage_logs
  FOR SELECT USING (auth.uid() = user_id);

-- Templates: Users can manage their own, view public ones
CREATE POLICY "Users can manage own templates" ON templates
  FOR ALL USING (auth.uid() = user_id);
CREATE POLICY "Anyone can view public templates" ON templates
  FOR SELECT USING (is_public = true);

-- Cache: Service role only
CREATE POLICY "Service role only" ON ai_cache
  FOR ALL USING (auth.role() = 'service_role');

-- SIMPLIFIED FUNCTIONS

-- Check if user can make a query
CREATE OR REPLACE FUNCTION can_user_query(user_id UUID)
RETURNS BOOLEAN
LANGUAGE plpgsql
SECURITY DEFINER
AS $$
DECLARE
  profile_record RECORD;
BEGIN
  SELECT * INTO profile_record
  FROM profiles
  WHERE id = user_id;
  
  -- Free users: check monthly limit
  IF profile_record.plan = 'free' THEN
    RETURN profile_record.queries_used < profile_record.queries_limit;
  END IF;
  
  -- Paid users: check if subscription is active
  RETURN profile_record.subscription_status = 'active' 
    AND (profile_record.current_period_end IS NULL 
         OR profile_record.current_period_end > NOW());
END;
$$;

-- Increment usage (simplified)
CREATE OR REPLACE FUNCTION increment_usage(user_id UUID)
RETURNS VOID
LANGUAGE plpgsql
SECURITY DEFINER
AS $$
BEGIN
  UPDATE profiles
  SET 
    queries_used = queries_used + 1,
    updated_at = NOW()
  WHERE id = user_id;
END;
$$;

-- Get usage stats for current period
CREATE OR REPLACE FUNCTION get_usage_stats(user_id UUID)
RETURNS JSON
LANGUAGE plpgsql
SECURITY DEFINER
AS $$
DECLARE
  profile_record RECORD;
  period_start TIMESTAMPTZ;
  usage_count INTEGER;
BEGIN
  SELECT * INTO profile_record
  FROM profiles
  WHERE id = user_id;
  
  -- Determine period start
  IF profile_record.plan = 'free' THEN
    period_start := date_trunc('month', NOW());
  ELSE
    period_start := profile_record.current_period_end - INTERVAL '30 days';
  END IF;
  
  -- Count usage in period
  SELECT COUNT(*) INTO usage_count
  FROM usage_logs
  WHERE user_id = user_id
    AND created_at >= period_start;
  
  RETURN json_build_object(
    'used', usage_count,
    'limit', profile_record.queries_limit,
    'plan', profile_record.plan,
    'period_start', period_start,
    'period_end', CASE 
      WHEN profile_record.plan = 'free' 
      THEN date_trunc('month', NOW()) + INTERVAL '1 month'
      ELSE profile_record.current_period_end
    END
  );
END;
$$;

-- Reset free tier usage monthly (called by cron)
CREATE OR REPLACE FUNCTION reset_free_tier_usage()
RETURNS VOID
LANGUAGE plpgsql
SECURITY DEFINER
AS $$
BEGIN
  UPDATE profiles
  SET 
    queries_used = 0,
    updated_at = NOW()
  WHERE 
    plan = 'free'
    AND date_trunc('month', NOW()) = date_trunc('day', NOW());
END;
$$;

-- Auto-create profile on signup
CREATE OR REPLACE FUNCTION handle_new_user()
RETURNS TRIGGER
LANGUAGE plpgsql
SECURITY DEFINER
AS $$
BEGIN
  INSERT INTO public.profiles (id, email, full_name)
  VALUES (
    NEW.id,
    NEW.email,
    NEW.raw_user_meta_data->>'full_name'
  );
  RETURN NEW;
END;
$$;

CREATE TRIGGER on_auth_user_created
  AFTER INSERT ON auth.users
  FOR EACH ROW
  EXECUTE FUNCTION handle_new_user();

-- Update timestamp trigger
CREATE OR REPLACE FUNCTION update_updated_at()
RETURNS TRIGGER
LANGUAGE plpgsql
AS $$
BEGIN
  NEW.updated_at = NOW();
  RETURN NEW;
END;
$$;

CREATE TRIGGER update_profiles_updated_at
  BEFORE UPDATE ON profiles
  FOR EACH ROW
  EXECUTE FUNCTION update_updated_at();

CREATE TRIGGER update_templates_updated_at
  BEFORE UPDATE ON templates
  FOR EACH ROW
  EXECUTE FUNCTION update_updated_at();

-- Set up cron job for free tier reset (if pg_cron is available)
-- SELECT cron.schedule('reset-free-usage', '0 0 1 * *', 'SELECT reset_free_tier_usage();');

-- Clean expired cache daily
-- SELECT cron.schedule('clean-cache', '0 2 * * *', 'DELETE FROM ai_cache WHERE expires_at < NOW();');