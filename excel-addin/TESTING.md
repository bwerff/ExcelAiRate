# Testing ExcelAIRate Add-in

## Prerequisites

1. Set up environment variables in `.env.local`:
   ```
   VITE_SUPABASE_URL=your-actual-supabase-url
   VITE_SUPABASE_ANON_KEY=your-actual-anon-key
   ```

2. Install dependencies:
   ```bash
   pnpm install
   ```

## Build and Run

1. Build the add-in:
   ```bash
   pnpm build:dev
   ```

2. Start the development server:
   ```bash
   pnpm dev-server
   ```

3. In another terminal, sideload the add-in:
   ```bash
   pnpm start
   ```

## Testing Custom Functions

### Method 1: In Excel

1. Open Excel with the add-in loaded
2. Click "Open ExcelAiRate" in the Home tab
3. Sign in with your email (magic link)
4. In any cell, type `=AI` to see all available functions:
   - `=AIANALYZE(data_range, analysis_type)`
   - `=AIGENERATE(prompt, context)`
   - `=AIEXPLAIN(formula)`
   - `=AISUMMARIZE(data_range, max_length)`
   - `=AIUSAGE()`

### Method 2: Test Script

Run the test script to verify function registration:
```bash
pnpm test:functions
```

## Verifying Functions Work

### Test 1: Basic Analysis
```excel
=AIANALYZE(A1:C10, "summary")
```

### Test 2: Generate Content
```excel
=AIGENERATE("Create a product description for Widget A")
```

### Test 3: Explain Formula
```excel
=AIEXPLAIN("=VLOOKUP(A2,B:C,2,FALSE)")
```

### Test 4: Check Usage
```excel
=AIUSAGE()
```

## Troubleshooting

### Functions Not Appearing
1. Check if `functions.json` was generated in the dist folder
2. Verify the manifest.xml has CustomFunctions extension point
3. Check browser console for errors

### Authentication Issues
1. Ensure you're signed in via the task pane first
2. Check if session is persisted in Office.context.document.settings
3. Verify Supabase credentials are correct

### API Errors
1. Check network tab for failed requests
2. Verify the AI edge function is deployed
3. Check usage limits haven't been exceeded

## Session Persistence Test

1. Sign in via task pane
2. Close and reopen Excel
3. Functions should still work without re-authentication
4. Session auto-refreshes before expiry

## Common Issues Fixed

✅ Webpack entry points corrected to use actual files
✅ API parameters standardized to use 'type' instead of 'operation'
✅ Session persistence implemented with auto-refresh
✅ Custom functions properly registered with Office.js

## Next Steps

1. Deploy the Supabase edge functions
2. Set up production URLs in manifest.xml
3. Test with real data and multiple users
4. Monitor usage and performance