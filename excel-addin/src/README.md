# Excel Add-in Source Structure

This directory contains the source code for the ExcelAiRate Excel Add-in.

## Directory Structure

- **taskpane/** - Task pane UI components
  - `taskpane.html` - HTML structure for the task pane
  - `taskpane.css` - Styles for the task pane
  - `taskpane.ts` - TypeScript logic for authentication and UI interactions

- **commands/** - Ribbon command handlers
  - `commands.ts` - Handles ribbon button clicks and other commands

- **functions/** - Custom Excel functions
  - `functions.ts` - AI-powered custom functions:
    - `EXCELAIRATE.AIANALYZE()` - Analyze data with AI
    - `EXCELAIRATE.AIGENERATE()` - Generate content with AI
    - `EXCELAIRATE.AIEXPLAIN()` - Explain formulas or concepts
    - `EXCELAIRATE.AISUMMARIZE()` - Summarize data
    - `EXCELAIRATE.AIUSAGE()` - Check usage limits

- **types/** - TypeScript type definitions
  - `globals.d.ts` - Global type declarations

## Key Features

1. **Authentication**: Magic link authentication via Supabase
2. **AI Operations**: Analyze, generate, explain, and summarize using AI
3. **Usage Tracking**: Monitor API usage and limits
4. **Custom Functions**: Excel formulas that call AI services
5. **Task Pane**: Interactive UI for more complex operations

## Environment Variables

The add-in requires these environment variables (set in `.env.local`):
- `VITE_SUPABASE_URL` - Your Supabase project URL
- `VITE_SUPABASE_ANON_KEY` - Your Supabase anonymous key