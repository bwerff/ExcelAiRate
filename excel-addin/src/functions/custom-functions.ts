/**
 * Excel AI Assistant - Custom Functions
 * All 19 AI-powered functions for Excel
 */

import { aiService, sharedState } from '../services/ai-service';

/* global CustomFunctions */

/**
 * Analyze data using AI
 * @customfunction
 * @param {any[][]} data Data range to analyze
 * @param {string} [analysisType] Analysis type (trend, summary, insights, correlation)
 * @returns {Promise<string>} Analysis result
 */
export async function ANALYZE(data: any[][], analysisType: string = 'summary'): Promise<string> {
  try {
    if (!data || data.length === 0) {
      return 'No data provided for analysis';
    }

    return await aiService.analyzeData(data, analysisType);
  } catch (error: any) {
    console.error('ANALYZE function error:', error);
    return `Analysis error: ${error.message}`;
  }
}

/**
 * Generate formulas using AI
 * @customfunction
 * @param {string} description Describe what you want the formula to do
 * @param {string} [dataRange] Reference data range (optional)
 * @returns {Promise<string>} Excel formula
 */
export async function FORMULA(description: string, dataRange?: string): Promise<string> {
  try {
    return await aiService.generateFormula(description, dataRange);
  } catch (error: any) {
    console.error('FORMULA function error:', error);
    return `Formula error: ${error.message}`;
  }
}

/**
 * Clean and format data using AI
 * @customfunction
 * @param {any[][]} data Data to clean
 * @param {string} [instructions] Cleaning instructions
 * @returns {Promise<string>} Cleaned data
 */
export async function CLEAN(data: any[][], instructions: string = 'standardize and clean'): Promise<string> {
  try {
    if (!data) return '';
    
    const context = aiService.buildContext(data, 'cleaning', instructions);
    
    const systemPrompt = `You are a data cleaning expert. Clean and standardize data according to instructions.
    Return ONLY the cleaned data in the same format as input. No explanations.`;
    
    const prompt = `Clean this data according to instructions: ${instructions}\n${context}`;

    const response = await aiService.callAI(prompt, {
      systemPrompt,
      maxTokens: 200,
    });
    
    return response.content.trim();
  } catch (error: any) {
    console.error('CLEAN function error:', error);
    return Array.isArray(data) ? data[0][0] : data; // Return original data if cleaning fails
  }
}

/**
 * Translate text using AI
 * @customfunction
 * @param {string} text Text to translate
 * @param {string} targetLanguage Target language
 * @returns {Promise<string>} Translated text
 */
export async function TRANSLATE(text: string, targetLanguage: string): Promise<string> {
  try {
    if (!text || !targetLanguage) {
      return 'Please provide text and target language';
    }

    const systemPrompt = `You are a professional translator. Translate text accurately while preserving meaning and context.
    Return ONLY the translated text, no explanations.`;
    
    const prompt = `Translate this text to ${targetLanguage}: "${text}"`;

    const response = await aiService.callAI(prompt, {
      systemPrompt,
      maxTokens: 150,
    });
    
    return response.content.trim();
  } catch (error: any) {
    console.error('TRANSLATE function error:', error);
    return `Translation error: ${error.message}`;
  }
}

/**
 * Extract insights from data
 * @customfunction
 * @param {any[][]} data Data range
 * @param {string} [focus] Focus area (trends, anomalies, patterns, predictions)
 * @returns {Promise<string>} Insights
 */
export async function INSIGHTS(data: any[][], focus: string = 'general'): Promise<string> {
  try {
    if (!data || data.length === 0) {
      return 'No data provided for insights';
    }

    const context = aiService.buildContext(data, 'insights extraction', `Focus: ${focus}`);
    
    const systemPrompt = `You are a business intelligence expert. Extract actionable insights and recommendations 
    from data. Focus on business value, opportunities, and strategic implications.`;
    
    let prompt = `Extract key insights from this data with focus on ${focus}:\n${context}`;
    
    switch (focus.toLowerCase()) {
      case 'trends':
        prompt += '\n\nFocus on trend analysis and future predictions.';
        break;
      case 'anomalies':
        prompt += '\n\nFocus on identifying outliers and unusual patterns.';
        break;
      case 'patterns':
        prompt += '\n\nFocus on recurring patterns and cyclical behavior.';
        break;
      case 'predictions':
        prompt += '\n\nFocus on predictive insights and forecasting.';
        break;
    }

    const response = await aiService.callAI(prompt, {
      systemPrompt,
      maxTokens: 400,
      operation: 'analyze',
    });
    
    return response.content;
  } catch (error: any) {
    console.error('INSIGHTS function error:', error);
    return `Insights error: ${error.message}`;
  }
}

/**
 * Categorize data using AI
 * @customfunction
 * @param {any} data Data to categorize
 * @param {string} [criteria] Category criteria or examples
 * @returns {Promise<string>} Category
 */
export async function CATEGORIZE(data: any, criteria: string = 'best fit categories'): Promise<string> {
  try {
    if (!data) return '';
    
    const context = aiService.buildContext(data, 'categorization', criteria);
    
    const systemPrompt = `You are a data categorization expert. Categorize data based on the provided criteria.
    Return ONLY the category name for the data, no explanations.`;
    
    const prompt = `Categorize this data based on: ${criteria}\n${context}`;

    const response = await aiService.callAI(prompt, {
      systemPrompt,
      maxTokens: 100,
    });
    
    return response.content.trim();
  } catch (error: any) {
    console.error('CATEGORIZE function error:', error);
    return 'Categorization error';
  }
}

/**
 * Generate sample data
 * @customfunction
 * @param {string} dataType Type of data to generate
 * @param {number} [rows] Number of rows
 * @param {string} [parameters] Additional parameters
 * @returns {Promise<string>} Generated data
 */
export async function GENERATE(dataType: string, rows: number = 10, parameters?: string): Promise<string> {
  try {
    const systemPrompt = `You are a data generation expert. Generate realistic sample data based on specifications.
    Return data in a format suitable for Excel (comma-separated or tab-separated values).`;
    
    let prompt = `Generate ${rows} rows of ${dataType} data`;
    if (parameters) {
      prompt += ` with these parameters: ${parameters}`;
    }
    prompt += '\n\nReturn only the data, formatted for Excel import.';

    const response = await aiService.callAI(prompt, {
      systemPrompt,
      maxTokens: 500,
    });
    
    return response.content.trim();
  } catch (error: any) {
    console.error('GENERATE function error:', error);
    return `Generation error: ${error.message}`;
  }
}

/**
 * Create dashboard visualizations
 * @customfunction
 * @param {string} dataRange Data range for dashboard
 * @param {string} [dashboardType] Dashboard type (executive, sales, financial, operational)
 * @param {string} [chartPrefs] Chart preferences
 * @returns {Promise<string>} Dashboard specifications
 */
export async function DASHBOARD(dataRange: string, dashboardType: string = 'executive', chartPrefs?: string): Promise<string> {
  try {
    if (!dataRange) {
      return 'Please specify a data range for the dashboard';
    }

    const systemPrompt = `You are a business intelligence dashboard expert. Design comprehensive dashboards 
    that provide actionable insights. Focus on the most important KPIs and visualizations for the specified type.`;
    
    let prompt = `Design a ${dashboardType} dashboard for data in range ${dataRange}.`;
    if (chartPrefs) {
      prompt += ` Chart preferences: ${chartPrefs}.`;
    }
    prompt += `\n\nProvide:
    1. Key metrics to track
    2. Recommended chart types
    3. Layout suggestions
    4. Color scheme recommendations
    5. Interactive elements to include`;

    const response = await aiService.callAI(prompt, {
      systemPrompt,
      maxTokens: 600,
    });
    
    return response.content;
  } catch (error: any) {
    console.error('DASHBOARD function error:', error);
    return `Dashboard error: ${error.message}`;
  }
}

/**
 * Generate PowerQuery M code
 * @customfunction
 * @param {string} transformation Data transformation description
 * @param {string} [sourceInfo] Source data info
 * @returns {Promise<string>} PowerQuery M code
 */
export async function POWERQUERY(transformation: string, sourceInfo?: string): Promise<string> {
  try {
    const systemPrompt = `You are a PowerQuery expert. Generate efficient M language code for data transformations.
    Create clean, optimized code that follows PowerQuery best practices. Include proper error handling and documentation.`;
    
    let prompt = `Generate PowerQuery M code for: ${transformation}`;
    if (sourceInfo) {
      prompt += `\nSource data info: ${sourceInfo}`;
    }
    prompt += `\n\nProvide:
    1. Complete M code
    2. Step-by-step explanation
    3. Performance optimization tips
    
    Return the M code in a code block, followed by explanations.`;

    const response = await aiService.callAI(prompt, {
      systemPrompt,
      maxTokens: 700,
    });
    
    return response.content;
  } catch (error: any) {
    console.error('POWERQUERY function error:', error);
    return `PowerQuery error: ${error.message}`;
  }
}

/**
 * Generate DAX formulas for PowerPivot
 * @customfunction
 * @param {string} measureDescription Measure description
 * @param {string} [tableContext] Table/column context
 * @returns {Promise<string>} DAX formula
 */
export async function DAX(measureDescription: string, tableContext?: string): Promise<string> {
  try {
    const systemPrompt = `You are a DAX expert. Create efficient, accurate DAX formulas for PowerPivot data models.
    Follow DAX best practices, use proper context transition, and optimize for performance.`;
    
    let prompt = `Generate a DAX formula for: ${measureDescription}`;
    if (tableContext) {
      prompt += `\nTable/column context: ${tableContext}`;
    }
    prompt += `\n\nProvide:
    1. The DAX formula
    2. Explanation of the logic
    3. Performance considerations
    4. Alternative approaches if applicable
    
    Start with the formula, then provide explanations.`;

    const response = await aiService.callAI(prompt, {
      systemPrompt,
      maxTokens: 500,
    });
    
    return response.content;
  } catch (error: any) {
    console.error('DAX function error:', error);
    return `DAX error: ${error.message}`;
  }
}

/**
 * Advanced predictive analytics and forecasting
 * @customfunction
 * @param {any[][]} historicalData Historical data for prediction
 * @param {string} [forecastType] Forecast type (linear, exponential, seasonal, arima)
 * @param {number} [periods] Number of periods to forecast
 * @returns {Promise<string>} Forecast results
 */
export async function FORECAST(historicalData: any[][], forecastType: string = 'linear', periods: number = 12): Promise<string> {
  try {
    if (!historicalData || historicalData.length === 0) {
      return 'No historical data provided for forecasting';
    }

    const context = aiService.buildContext(historicalData, 'forecasting', `Type: ${forecastType}, Periods: ${periods}`);
    
    const systemPrompt = `You are a predictive analytics expert. Generate accurate forecasts using statistical methods.
    Consider seasonality, trends, and data patterns. Provide confidence intervals and methodology explanations.`;
    
    let prompt = `Create a ${forecastType} forecast for ${periods} future periods using this historical data:\n${context}`;
    
    prompt += `\n\nProvide:
    1. Forecasted values for each period
    2. Confidence intervals (upper/lower bounds)
    3. Key assumptions and methodology
    4. Recommended Excel formulas for implementation
    5. Data quality assessment`;

    const response = await aiService.callAI(prompt, {
      systemPrompt,
      maxTokens: 600,
    });
    
    return response.content;
  } catch (error: any) {
    console.error('FORECAST function error:', error);
    return `Forecast error: ${error.message}`;
  }
}

/**
 * Statistical analysis and machine learning
 * @customfunction
 * @param {any[][]} dataset Dataset for analysis
 * @param {string} analysisType Analysis type (regression, correlation, clustering, anomaly)
 * @param {string} [targetVariable] Target variable or parameters
 * @returns {Promise<string>} Analysis results
 */
export async function ANALYZE_ML(dataset: any[][], analysisType: string, targetVariable?: string): Promise<string> {
  try {
    if (!dataset || dataset.length === 0) {
      return 'No dataset provided for ML analysis';
    }

    const context = aiService.buildContext(dataset, 'machine learning analysis', `Type: ${analysisType}, Target: ${targetVariable || 'N/A'}`);
    
    const systemPrompt = `You are a data scientist expert. Perform advanced statistical analysis and machine learning tasks.
    Provide actionable insights, statistical significance, and practical recommendations.`;
    
    let prompt = `Perform ${analysisType} analysis on this dataset:\n${context}`;
    if (targetVariable) {
      prompt += `\nTarget variable: ${targetVariable}`;
    }
    
    switch (analysisType.toLowerCase()) {
      case 'regression':
        prompt += '\n\nProvide: R-squared, coefficients, significance tests, prediction equation, residual analysis';
        break;
      case 'correlation':
        prompt += '\n\nProvide: Correlation matrix, strongest relationships, statistical significance, recommendations';
        break;
      case 'clustering':
        prompt += '\n\nProvide: Optimal cluster count, cluster characteristics, assignment rules, business interpretation';
        break;
      case 'anomaly':
        prompt += '\n\nProvide: Anomaly detection method, outliers identified, severity scores, investigation priorities';
        break;
      default:
        prompt += '\n\nProvide: Key findings, statistical measures, business insights, next steps';
    }

    const response = await aiService.callAI(prompt, {
      systemPrompt,
      maxTokens: 700,
    });
    
    return response.content;
  } catch (error: any) {
    console.error('ANALYZE_ML function error:', error);
    return `ML Analysis error: ${error.message}`;
  }
}

/**
 * Generate VBA/Office Scripts automation
 * @customfunction
 * @param {string} taskDescription Task to automate
 * @param {string} [scriptType] Script type (vba, office-script, power-automate)
 * @returns {Promise<string>} Automation script
 */
export async function AUTOMATE(taskDescription: string, scriptType: string = 'office-script'): Promise<string> {
  try {
    const systemPrompt = `You are an Excel automation expert. Generate efficient, well-documented automation scripts.
    Follow best practices for error handling, performance, and maintainability.`;
    
    let prompt = `Generate a ${scriptType} script to: ${taskDescription}`;
    
    switch (scriptType.toLowerCase()) {
      case 'vba':
        prompt += '\n\nProvide VBA code with proper error handling, comments, and user-friendly features.';
        break;
      case 'office-script':
        prompt += '\n\nProvide TypeScript Office Script code with modern syntax and async/await patterns.';
        break;
      case 'power-automate':
        prompt += '\n\nProvide Power Automate flow steps and configuration details.';
        break;
      default:
        prompt += '\n\nProvide Office Script code as the default automation solution.';
    }
    
    prompt += '\n\nInclude: Complete code, setup instructions, usage guidelines, troubleshooting tips.';

    const response = await aiService.callAI(prompt, {
      systemPrompt,
      maxTokens: 800,
    });
    
    return response.content;
  } catch (error: any) {
    console.error('AUTOMATE function error:', error);
    return `Automation error: ${error.message}`;
  }
}

/**
 * Generate SQL queries for data integration
 * @customfunction
 * @param {string} queryDescription Query description
 * @param {string} [schemaInfo] Database schema information
 * @param {string} [dialect] SQL dialect (tsql, mysql, postgres, oracle)
 * @returns {Promise<string>} SQL query
 */
export async function SQL(queryDescription: string, schemaInfo?: string, dialect: string = 'tsql'): Promise<string> {
  try {
    const systemPrompt = `You are a SQL expert. Generate efficient, optimized SQL queries for data integration.
    Follow best practices for performance, readability, and maintainability.`;
    
    let prompt = `Generate a ${dialect.toUpperCase()} query for: ${queryDescription}`;
    if (schemaInfo) {
      prompt += `\nSchema information: ${schemaInfo}`;
    }
    
    prompt += `\n\nProvide:
    1. Complete SQL query with proper formatting
    2. Explanation of the query logic
    3. Performance optimization tips
    4. Integration steps for Excel
    5. Alternative approaches if applicable`;

    const response = await aiService.callAI(prompt, {
      systemPrompt,
      maxTokens: 600,
    });
    
    return response.content;
  } catch (error: any) {
    console.error('SQL function error:', error);
    return `SQL error: ${error.message}`;
  }
}

/**
 * Financial modeling and scenario analysis
 * @customfunction
 * @param {string} modelType Model type (dcf, budget, forecast, sensitivity)
 * @param {any[][]} [financialData] Financial data
 * @param {string} [scenarios] Scenario parameters
 * @returns {Promise<string>} Financial model
 */
export async function FINMODEL(modelType: string, financialData?: any[][], scenarios?: string): Promise<string> {
  try {
    const systemPrompt = `You are a financial modeling expert. Create comprehensive, professional financial models
    with proper assumptions, sensitivity analysis, and scenario planning.`;
    
    let prompt = `Create a ${modelType} financial model`;
    if (financialData) {
      prompt += ` using this data: ${aiService.preprocessData(financialData)}`;
    }
    if (scenarios) {
      prompt += `\nScenario parameters: ${scenarios}`;
    }
    
    switch (modelType.toLowerCase()) {
      case 'dcf':
        prompt += '\n\nInclude: Cash flow projections, discount rate analysis, terminal value, sensitivity analysis';
        break;
      case 'budget':
        prompt += '\n\nInclude: Revenue forecasts, expense categories, variance analysis, budget vs actual tracking';
        break;
      case 'forecast':
        prompt += '\n\nInclude: Historical trend analysis, growth assumptions, seasonality factors, confidence intervals';
        break;
      case 'sensitivity':
        prompt += '\n\nInclude: Key variable identification, tornado charts, scenario comparison, risk assessment';
        break;
      default:
        prompt += '\n\nInclude: Model structure, key assumptions, calculations, scenario analysis';
    }
    
    prompt += '\n\nProvide: Excel formulas, model structure, assumptions documentation, validation checks';

    const response = await aiService.callAI(prompt, {
      systemPrompt,
      maxTokens: 800,
    });
    
    return response.content;
  } catch (error: any) {
    console.error('FINMODEL function error:', error);
    return `Financial modeling error: ${error.message}`;
  }
}

/**
 * Data quality assessment and validation
 * @customfunction
 * @param {any[][]} data Data to validate
 * @param {string} [validationRules] Validation rules or criteria
 * @returns {Promise<string>} Validation report
 */
export async function VALIDATE(data: any[][], validationRules?: string): Promise<string> {
  try {
    if (!data || data.length === 0) {
      return 'No data provided for validation';
    }

    const context = aiService.buildContext(data, 'data validation', validationRules || '');
    
    const systemPrompt = `You are a data quality expert. Assess data quality, identify issues, and provide remediation strategies.
    Focus on completeness, accuracy, consistency, and business rule compliance.`;
    
    let prompt = `Validate this data for quality issues:\n${context}`;
    if (validationRules) {
      prompt += `\nValidation rules: ${validationRules}`;
    }
    
    prompt += `\n\nProvide:
    1. Data quality score (0-100)
    2. Issues identified (missing values, duplicates, outliers, format errors)
    3. Impact assessment for each issue
    4. Remediation recommendations
    5. Excel formulas for ongoing validation`;

    const response = await aiService.callAI(prompt, {
      systemPrompt,
      maxTokens: 600,
    });
    
    return response.content;
  } catch (error: any) {
    console.error('VALIDATE function error:', error);
    return `Validation error: ${error.message}`;
  }
}

/**
 * Advanced visualization and custom charts
 * @customfunction
 * @param {any[][]} data Data for visualization
 * @param {string} vizType Visualization type (heatmap, treemap, waterfall, etc.)
 * @param {string} [styleOptions] Style preferences
 * @returns {Promise<string>} Visualization specifications
 */
export async function VISUALIZE(data: any[][], vizType: string, styleOptions?: string): Promise<string> {
  try {
    if (!data || data.length === 0) {
      return 'No data provided for visualization';
    }

    const context = aiService.buildContext(data, 'visualization', `Type: ${vizType}, Style: ${styleOptions || 'default'}`);
    
    const systemPrompt = `You are a data visualization expert. Create compelling, informative visualizations
    that tell clear stories and drive business decisions.`;
    
    let prompt = `Create a ${vizType} visualization for this data:\n${context}`;
    if (styleOptions) {
      prompt += `\nStyle preferences: ${styleOptions}`;
    }
    
    prompt += `\n\nProvide:
    1. Visualization design recommendations
    2. Color scheme and formatting suggestions
    3. Excel implementation steps
    4. Alternative chart types to consider
    5. Interactive elements to include`;

    const response = await aiService.callAI(prompt, {
      systemPrompt,
      maxTokens: 600,
    });
    
    return response.content;
  } catch (error: any) {
    console.error('VISUALIZE function error:', error);
    return `Visualization error: ${error.message}`;
  }
}

/**
 * Performance optimization and formula enhancement
 * @customfunction
 * @param {string} target Formula or workbook area to optimize
 * @param {string} [optimizationType] Optimization type (speed, memory, accuracy, size)
 * @returns {Promise<string>} Optimization recommendations
 */
export async function OPTIMIZE(target: string, optimizationType: string = 'speed'): Promise<string> {
  try {
    const systemPrompt = `You are an Excel performance optimization expert. Identify bottlenecks and provide
    specific, actionable recommendations for improving speed, memory usage, and accuracy.`;
    
    let prompt = `Optimize this Excel element for ${optimizationType}: ${target}`;
    
    switch (optimizationType.toLowerCase()) {
      case 'speed':
        prompt += '\n\nFocus on: Formula efficiency, calculation mode, volatile functions, array formulas';
        break;
      case 'memory':
        prompt += '\n\nFocus on: File size reduction, unused objects, formatting optimization, data compression';
        break;
      case 'accuracy':
        prompt += '\n\nFocus on: Numerical precision, error handling, validation rules, data integrity';
        break;
      case 'size':
        prompt += '\n\nFocus on: File compression, unused elements, formatting consolidation, data archiving';
        break;
    }
    
    prompt += `\n\nProvide:
    1. Specific optimization recommendations
    2. Before/after performance estimates
    3. Implementation steps
    4. Potential risks and mitigation
    5. Monitoring and maintenance tips`;

    const response = await aiService.callAI(prompt, {
      systemPrompt,
      maxTokens: 600,
    });
    
    return response.content;
  } catch (error: any) {
    console.error('OPTIMIZE function error:', error);
    return `Optimization error: ${error.message}`;
  }
}

/**
 * API integration and web data extraction
 * @customfunction
 * @param {string} dataSource Data source description
 * @param {string} connectionType Connection type (api, web-scrape, database, file)
 * @param {string} [config] Authentication or configuration details
 * @returns {Promise<string>} Connection instructions
 */
export async function CONNECT(dataSource: string, connectionType: string, config?: string): Promise<string> {
  try {
    const systemPrompt = `You are a data integration expert. Create robust, secure connections to external data sources
    with proper error handling and data transformation capabilities.`;
    
    let prompt = `Create a ${connectionType} connection to: ${dataSource}`;
    if (config) {
      prompt += `\nConfiguration: ${config}`;
    }
    
    switch (connectionType.toLowerCase()) {
      case 'api':
        prompt += '\n\nProvide: API endpoint setup, authentication, data parsing, error handling, refresh automation';
        break;
      case 'web-scrape':
        prompt += '\n\nProvide: Web scraping strategy, data extraction rules, anti-bot considerations, legal compliance';
        break;
      case 'database':
        prompt += '\n\nProvide: Connection string, query optimization, security considerations, data mapping';
        break;
      case 'file':
        prompt += '\n\nProvide: File format handling, import automation, data validation, update procedures';
        break;
    }
    
    prompt += `\n\nInclude: Step-by-step setup, PowerQuery code, troubleshooting guide, maintenance recommendations`;

    const response = await aiService.callAI(prompt, {
      systemPrompt,
      maxTokens: 700,
    });
    
    return response.content;
  } catch (error: any) {
    console.error('CONNECT function error:', error);
    return `Connection error: ${error.message}`;
  }
}

// Register all custom functions
CustomFunctions.associate('ANALYZE', ANALYZE);
CustomFunctions.associate('FORMULA', FORMULA);
CustomFunctions.associate('CLEAN', CLEAN);
CustomFunctions.associate('TRANSLATE', TRANSLATE);
CustomFunctions.associate('INSIGHTS', INSIGHTS);
CustomFunctions.associate('CATEGORIZE', CATEGORIZE);
CustomFunctions.associate('GENERATE', GENERATE);
CustomFunctions.associate('DASHBOARD', DASHBOARD);
CustomFunctions.associate('POWERQUERY', POWERQUERY);
CustomFunctions.associate('DAX', DAX);
CustomFunctions.associate('FORECAST', FORECAST);
CustomFunctions.associate('ANALYZE_ML', ANALYZE_ML);
CustomFunctions.associate('AUTOMATE', AUTOMATE);
CustomFunctions.associate('SQL', SQL);
CustomFunctions.associate('FINMODEL', FINMODEL);
CustomFunctions.associate('VALIDATE', VALIDATE);
CustomFunctions.associate('VISUALIZE', VISUALIZE);
CustomFunctions.associate('OPTIMIZE', OPTIMIZE);
CustomFunctions.associate('CONNECT', CONNECT);