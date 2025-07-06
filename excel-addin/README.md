# ExcelAiRate - Advanced Excel AI Assistant

A comprehensive Excel Add-in that brings the power of AI to your spreadsheets with 19 custom functions, an intelligent task pane, and advanced business intelligence capabilities.

## Features

### 🎯 19 AI-Powered Custom Functions

#### Data Analysis
- `=AI.ANALYZE(data, [type])` - Analyze data trends and patterns
- `=AI.INSIGHTS(data, [focus])` - Extract business insights
- `=AI.ANALYZE_ML(data, type, [target])` - Advanced statistical analysis
- `=AI.FORECAST(data, [type], [periods])` - Predictive analytics

#### Data Operations
- `=AI.CLEAN(data, [instructions])` - Clean and format data
- `=AI.CATEGORIZE(data, [criteria])` - Categorize data intelligently
- `=AI.VALIDATE(data, [rules])` - Data quality assessment
- `=AI.TRANSLATE(text, language)` - Translate text
- `=AI.GENERATE(type, [rows], [params])` - Generate sample data

#### Visualization & BI
- `=AI.DASHBOARD(range, [type], [prefs])` - Create interactive dashboards
- `=AI.VISUALIZE(data, type, [style])` - Advanced visualizations

#### Excel Enhancement
- `=AI.FORMULA(description)` - Generate Excel formulas
- `=AI.OPTIMIZE(target, [type])` - Performance optimization
- `=AI.AUTOMATE(task, [script-type])` - Generate automation code

#### Data Integration
- `=AI.POWERQUERY(transformation, [source])` - Generate PowerQuery M code
- `=AI.SQL(description, [schema], [dialect])` - Generate SQL queries
- `=AI.CONNECT(source, type, [config])` - API & data connections
- `=AI.DAX(description, [context])` - Create DAX formulas

#### Financial Modeling
- `=AI.FINMODEL(type, [data], [scenarios])` - Financial models

### 🎨 Intelligent Task Pane

- **Interactive Chat Interface**: Natural language interaction with AI
- **Quick Actions**: One-click access to common tasks
- **Advanced BI Tools**: Dashboard creation, PowerQuery wizard, DAX helper
- **ML & Analytics**: Forecasting, ML analysis, advanced visualizations
- **Automation Tools**: VBA generator, SQL queries, API connections
- **Financial Modeling**: DCF models, scenario analysis, Monte Carlo simulations

### 🚀 Key Capabilities

1. **Shared Runtime Architecture**: Seamless communication between custom functions and task pane
2. **Context-Aware AI**: Understands your workbook context for better recommendations
3. **Response Caching**: Improved performance and reduced API costs
4. **Usage Tracking**: Monitor AI usage with built-in statistics
5. **Enterprise Security**: Secure authentication via Supabase

## Installation

### Prerequisites

- Node.js >= 18
- pnpm >= 9.0.0
- Microsoft Excel (Desktop or Online)
- ExcelAiRate account for authentication

### Setup

1. Clone the repository:
```bash
git clone https://github.com/excelairate/excelairate.git
cd excelairate/excel-addin
```

2. Install dependencies:
```bash
pnpm install
```

3. Create `.env.local` file:
```env
VITE_SUPABASE_URL=your_supabase_url
VITE_SUPABASE_ANON_KEY=your_supabase_anon_key
VITE_OPENAI_API_KEY=your_openai_api_key
```

4. Start the development server:
```bash
pnpm run dev-server
```

5. Sideload the add-in:
```bash
pnpm start
```

## Usage

### Using Custom Functions

1. Sign in through the task pane
2. Use any AI function in a cell:
   ```excel
   =AI.ANALYZE(A1:D100, "trend")
   =AI.FORMULA("calculate compound annual growth rate")
   =AI.FORECAST(B2:B50, "seasonal", 12)
   ```

### Using the Task Pane

1. Click "Open ExcelAiRate" in the Home tab
2. Sign in with your email
3. Use natural language to:
   - Analyze selected data
   - Generate formulas
   - Create dashboards
   - Build financial models

## Architecture

### Technology Stack

- **Frontend**: TypeScript, Office.js, Webpack
- **Backend**: Supabase Edge Functions
- **AI**: OpenAI GPT-4o-mini with fallback to GPT-3.5-turbo
- **Auth**: Supabase Auth with magic links

### Project Structure

```
excel-addin/
├── src/
│   ├── services/
│   │   └── ai-service.ts          # Core AI service layer
│   ├── functions/
│   │   └── custom-functions.ts    # All 19 custom functions
│   ├── taskpane/
│   │   ├── taskpane-enhanced.ts   # Enhanced task pane logic
│   │   ├── taskpane-enhanced.html # Task pane UI
│   │   └── taskpane-enhanced.css  # Styling
│   ├── utils/
│   │   ├── excel-helpers.ts       # Excel utility functions
│   │   ├── dashboard-builder.ts   # Dashboard creation
│   │   └── dialog-manager.ts      # Dialog UI management
│   └── commands/
│       └── commands.ts            # Ribbon commands
├── manifest.xml                   # Add-in manifest
├── webpack.config.js              # Build configuration
└── package.json                   # Dependencies
```

## Advanced Features

### Dashboard Creation
The add-in can automatically create comprehensive dashboards with:
- KPI cards with trends
- Optimal chart selection based on data
- Professional themes
- Interactive elements

### PowerQuery Integration
Generate M code for complex data transformations:
- Data cleaning and shaping
- Table merges and appends
- Custom column creation
- Performance optimization

### Financial Modeling
Build sophisticated financial models:
- DCF valuations with sensitivity analysis
- Monte Carlo simulations
- Scenario planning
- Budget variance analysis

## Development

### Building for Production

```bash
pnpm run build
```

### Running Tests

```bash
pnpm test
```

### Code Quality

```bash
pnpm run lint
pnpm run format
```

## Deployment

1. Update the production URL in `webpack.config.js`
2. Build the production bundle
3. Deploy to your hosting service
4. Update the manifest with production URLs

## Best Practices

1. **Performance**: Use shared runtime for optimal performance
2. **Caching**: Leverage built-in response caching
3. **Error Handling**: All functions include comprehensive error handling
4. **User Experience**: Provide clear feedback and loading states

## Troubleshooting

### Common Issues

1. **Functions not appearing**: Ensure you're signed in via the task pane
2. **Slow performance**: Check your internet connection and Excel calculation settings
3. **Authentication errors**: Clear browser cache and re-authenticate

### Debug Mode

Enable debug logging:
```javascript
// In console
localStorage.setItem('DEBUG', 'true');
```

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Submit a pull request

## License

MIT License - see LICENSE file for details

## Support

- Documentation: https://docs.excelairate.com
- Issues: https://github.com/excelairate/excelairate/issues
- Email: support@excelairate.com