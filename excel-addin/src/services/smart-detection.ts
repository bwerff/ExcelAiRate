/**
 * Smart Range Detection & Context Awareness Service
 * Automatically detects data types, headers, and relationships in Excel data
 */

import { ExcelHelpers } from '../utils/excel-helpers';

/* global Excel */

// Data type definitions
export type DataType = 'numeric' | 'date' | 'currency' | 'percentage' | 'text' | 'boolean' | 'mixed' | 'empty';

export interface DetectedDataInfo {
  range: string;
  dataType: DataType;
  columnTypes?: DataType[];
  headers: string[];
  hasHeaders: boolean;
  namedRanges: string[];
  relatedRanges?: string[];
  statistics?: DataStatistics;
  format?: string;
  nullCount: number;
  uniqueValues?: number;
  patterns?: DataPattern[];
  suggestions?: string[];
}

export interface DataStatistics {
  count: number;
  sum?: number;
  average?: number;
  min?: number;
  max?: number;
  median?: number;
  stdDev?: number;
  quartiles?: number[];
  mode?: any;
}

export interface DataPattern {
  type: 'email' | 'phone' | 'url' | 'id' | 'postal' | 'custom';
  confidence: number;
  examples: string[];
  regex?: string;
}

export interface SmartContext {
  primaryData: DetectedDataInfo;
  relatedData: DetectedDataInfo[];
  worksheetContext: WorksheetContext;
  suggestions: ContextSuggestion[];
}

export interface WorksheetContext {
  name: string;
  totalRows: number;
  totalColumns: number;
  namedRanges: NamedRangeInfo[];
  tables: TableInfo[];
  charts: ChartInfo[];
  pivotTables: PivotInfo[];
}

export interface NamedRangeInfo {
  name: string;
  address: string;
  scope: 'workbook' | 'worksheet';
}

export interface TableInfo {
  name: string;
  range: string;
  headers: string[];
  rowCount: number;
}

export interface ChartInfo {
  name: string;
  type: string;
  dataRange: string;
}

export interface PivotInfo {
  name: string;
  sourceData: string;
  location: string;
}

export interface ContextSuggestion {
  type: 'analysis' | 'formatting' | 'validation' | 'relationship';
  description: string;
  confidence: number;
  action?: () => Promise<void>;
}

export class SmartDetectionService {
  private excelHelpers: ExcelHelpers;
  private cache: Map<string, DetectedDataInfo> = new Map();

  constructor() {
    this.excelHelpers = new ExcelHelpers();
  }

  /**
   * Analyze selected range and detect all relevant information
   */
  async analyzeSelection(): Promise<SmartContext> {
    return await Excel.run(async (context) => {
      const selection = context.workbook.getSelectedRange();
      selection.load(['address', 'values', 'numberFormat', 'rowCount', 'columnCount']);
      
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      worksheet.load(['name', 'tables', 'charts']);
      
      await context.sync();

      // Analyze primary selection
      const primaryData = await this.detectDataInfo(selection, context);
      
      // Get worksheet context
      const worksheetContext = await this.getWorksheetContext(worksheet, context);
      
      // Find related data
      const relatedData = await this.findRelatedData(primaryData, worksheetContext, context);
      
      // Generate suggestions
      const suggestions = this.generateSuggestions(primaryData, relatedData, worksheetContext);

      return {
        primaryData,
        relatedData,
        worksheetContext,
        suggestions
      };
    });
  }

  /**
   * Detect comprehensive information about a data range
   */
  private async detectDataInfo(range: Excel.Range, context: Excel.RequestContext): Promise<DetectedDataInfo> {
    const cacheKey = range.address;
    if (this.cache.has(cacheKey)) {
      return this.cache.get(cacheKey)!;
    }

    const values = range.values;
    const numberFormats = range.numberFormat;
    
    // Detect headers
    const headerDetection = this.detectHeaders(values);
    
    // Detect data types
    const typeDetection = this.detectDataTypes(values, numberFormats, headerDetection.hasHeaders);
    
    // Calculate statistics if numeric
    const statistics = this.calculateStatistics(values, typeDetection.columnTypes, headerDetection.hasHeaders);
    
    // Detect patterns
    const patterns = this.detectPatterns(values, headerDetection.hasHeaders);
    
    // Count nulls and unique values
    const { nullCount, uniqueValues } = this.analyzeDataQuality(values, headerDetection.hasHeaders);
    
    // Find named ranges that include this range
    const namedRanges = await this.findNamedRangesForRange(range.address, context);

    const info: DetectedDataInfo = {
      range: range.address,
      dataType: typeDetection.overallType,
      columnTypes: typeDetection.columnTypes,
      headers: headerDetection.headers,
      hasHeaders: headerDetection.hasHeaders,
      namedRanges,
      statistics,
      format: this.detectFormat(numberFormats),
      nullCount,
      uniqueValues,
      patterns,
      suggestions: this.generateDataSuggestions(typeDetection, patterns, statistics)
    };

    this.cache.set(cacheKey, info);
    return info;
  }

  /**
   * Detect if the first row contains headers
   */
  private detectHeaders(values: any[][]): { hasHeaders: boolean; headers: string[] } {
    if (!values || values.length < 2) {
      return { hasHeaders: false, headers: [] };
    }

    const firstRow = values[0];
    const secondRow = values[1];

    // Check if first row is all text and different type from other rows
    const firstRowTypes = firstRow.map(cell => this.detectCellType(cell));
    const secondRowTypes = secondRow.map(cell => this.detectCellType(cell));

    let headerScore = 0;

    // All text in first row
    if (firstRowTypes.every(type => type === 'text')) {
      headerScore += 3;
    }

    // Different types between first and second row
    const typeDifferences = firstRowTypes.filter((type, i) => type !== secondRowTypes[i]).length;
    if (typeDifferences > firstRowTypes.length / 2) {
      headerScore += 2;
    }

    // Check for common header patterns
    const headerPatterns = ['id', 'name', 'date', 'amount', 'total', 'count', 'description', 'category', 'type', 'status'];
    const firstRowLower = firstRow.map(cell => String(cell).toLowerCase());
    const patternMatches = firstRowLower.filter(cell => 
      headerPatterns.some(pattern => cell.includes(pattern))
    ).length;
    
    if (patternMatches > 0) {
      headerScore += patternMatches;
    }

    // Statistical check: headers usually have more unique values in columns
    const columnUniqueness = this.calculateColumnUniqueness(values);
    if (columnUniqueness.withHeaders > columnUniqueness.withoutHeaders * 1.2) {
      headerScore += 2;
    }

    const hasHeaders = headerScore >= 3;
    const headers = hasHeaders ? firstRow.map(h => String(h)) : [];

    return { hasHeaders, headers };
  }

  /**
   * Detect data types for the range
   */
  private detectDataTypes(values: any[][], numberFormats: any[][], hasHeaders: boolean): {
    overallType: DataType;
    columnTypes: DataType[];
  } {
    const startRow = hasHeaders ? 1 : 0;
    const columnTypes: DataType[] = [];
    const typeCountMap = new Map<DataType, number>();

    for (let col = 0; col < values[0].length; col++) {
      const columnData = values.slice(startRow).map(row => row[col]);
      const columnFormats = numberFormats.slice(startRow).map(row => row[col]);
      const columnType = this.detectColumnType(columnData, columnFormats);
      
      columnTypes.push(columnType);
      typeCountMap.set(columnType, (typeCountMap.get(columnType) || 0) + 1);
    }

    // Determine overall type
    let overallType: DataType = 'mixed';
    const uniqueTypes = Array.from(typeCountMap.keys()).filter(t => t !== 'empty');
    
    if (uniqueTypes.length === 1) {
      overallType = uniqueTypes[0];
    } else if (uniqueTypes.length === 0) {
      overallType = 'empty';
    }

    return { overallType, columnTypes };
  }

  /**
   * Detect type of a single column
   */
  private detectColumnType(columnData: any[], formats?: string[]): DataType {
    const nonEmptyData = columnData.filter(val => val != null && val !== '');
    
    if (nonEmptyData.length === 0) {
      return 'empty';
    }

    const types = nonEmptyData.map((val, i) => {
      const format = formats?.[i];
      return this.detectCellType(val, format);
    });

    // Check for consistency
    const typeCount = new Map<DataType, number>();
    types.forEach(type => {
      typeCount.set(type, (typeCount.get(type) || 0) + 1);
    });

    // If 90% or more are the same type, use that type
    const threshold = nonEmptyData.length * 0.9;
    for (const [type, count] of typeCount.entries()) {
      if (count >= threshold) {
        return type;
      }
    }

    return 'mixed';
  }

  /**
   * Detect type of a single cell
   */
  private detectCellType(value: any, format?: string): DataType {
    if (value == null || value === '') {
      return 'empty';
    }

    // Check format first
    if (format) {
      if (format.includes('$') || format.includes('¤')) {
        return 'currency';
      }
      if (format.includes('%')) {
        return 'percentage';
      }
      if (format.includes('d') || format.includes('m') || format.includes('y')) {
        return 'date';
      }
    }

    // Check value type
    if (typeof value === 'boolean') {
      return 'boolean';
    }

    if (typeof value === 'number' || !isNaN(Number(value))) {
      // Excel dates are numbers
      if (value > 25569 && value < 60000) { // Rough date serial number range
        return 'date';
      }
      return 'numeric';
    }

    // Check for date strings
    const dateVal = Date.parse(String(value));
    if (!isNaN(dateVal)) {
      const year = new Date(dateVal).getFullYear();
      if (year > 1900 && year < 2100) {
        return 'date';
      }
    }

    // Check for boolean strings
    const strVal = String(value).toLowerCase();
    if (['true', 'false', 'yes', 'no', 'y', 'n'].includes(strVal)) {
      return 'boolean';
    }

    return 'text';
  }

  /**
   * Calculate statistics for numeric data
   */
  private calculateStatistics(values: any[][], columnTypes: DataType[], hasHeaders: boolean): DataStatistics | undefined {
    const numericColumns = columnTypes
      .map((type, index) => ({ type, index }))
      .filter(({ type }) => type === 'numeric' || type === 'currency' || type === 'percentage')
      .map(({ index }) => index);

    if (numericColumns.length === 0) {
      return undefined;
    }

    const startRow = hasHeaders ? 1 : 0;
    const numericData: number[] = [];

    for (let row = startRow; row < values.length; row++) {
      for (const col of numericColumns) {
        const val = values[row][col];
        if (val != null && !isNaN(Number(val))) {
          numericData.push(Number(val));
        }
      }
    }

    if (numericData.length === 0) {
      return undefined;
    }

    return this.excelHelpers.calculateStatistics(numericData);
  }

  /**
   * Detect patterns in text data
   */
  private detectPatterns(values: any[][], hasHeaders: boolean): DataPattern[] {
    const patterns: DataPattern[] = [];
    const startRow = hasHeaders ? 1 : 0;
    
    // Email pattern
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    const emailExamples: string[] = [];
    
    // Phone pattern
    const phoneRegex = /^[\+]?[(]?[0-9]{3}[)]?[-\s\.]?[(]?[0-9]{3}[)]?[-\s\.]?[0-9]{4,6}$/;
    const phoneExamples: string[] = [];
    
    // URL pattern
    const urlRegex = /^(https?:\/\/)?([\da-z\.-]+)\.([a-z\.]{2,6})([\/\w \.-]*)*\/?$/;
    const urlExamples: string[] = [];

    // Scan data for patterns
    for (let row = startRow; row < values.length && row < startRow + 100; row++) { // Sample first 100 rows
      for (let col = 0; col < values[row].length; col++) {
        const val = String(values[row][col]);
        
        if (emailRegex.test(val) && emailExamples.length < 3) {
          emailExamples.push(val);
        }
        if (phoneRegex.test(val) && phoneExamples.length < 3) {
          phoneExamples.push(val);
        }
        if (urlRegex.test(val) && urlExamples.length < 3) {
          urlExamples.push(val);
        }
      }
    }

    // Add detected patterns
    if (emailExamples.length > 0) {
      patterns.push({
        type: 'email',
        confidence: Math.min(emailExamples.length / 3, 1),
        examples: emailExamples,
        regex: emailRegex.source
      });
    }
    
    if (phoneExamples.length > 0) {
      patterns.push({
        type: 'phone',
        confidence: Math.min(phoneExamples.length / 3, 1),
        examples: phoneExamples,
        regex: phoneRegex.source
      });
    }
    
    if (urlExamples.length > 0) {
      patterns.push({
        type: 'url',
        confidence: Math.min(urlExamples.length / 3, 1),
        examples: urlExamples,
        regex: urlRegex.source
      });
    }

    return patterns;
  }

  /**
   * Analyze data quality metrics
   */
  private analyzeDataQuality(values: any[][], hasHeaders: boolean): {
    nullCount: number;
    uniqueValues: number;
  } {
    const startRow = hasHeaders ? 1 : 0;
    let nullCount = 0;
    const uniqueSet = new Set<string>();

    for (let row = startRow; row < values.length; row++) {
      for (let col = 0; col < values[row].length; col++) {
        const val = values[row][col];
        if (val == null || val === '') {
          nullCount++;
        } else {
          uniqueSet.add(String(val));
        }
      }
    }

    return {
      nullCount,
      uniqueValues: uniqueSet.size
    };
  }

  /**
   * Find named ranges that include the given address
   */
  private async findNamedRangesForRange(address: string, context: Excel.RequestContext): Promise<string[]> {
    const names: string[] = [];
    
    try {
      const namedItems = context.workbook.names;
      namedItems.load('items');
      await context.sync();

      for (const namedItem of namedItems.items) {
        try {
          const range = namedItem.getRange();
          range.load('address');
          await context.sync();
          
          if (this.rangesOverlap(address, range.address)) {
            names.push(namedItem.name);
          }
        } catch (e) {
          // Named item might not be a range
          continue;
        }
      }
    } catch (error) {
      console.error('Error finding named ranges:', error);
    }

    return names;
  }

  /**
   * Check if two range addresses overlap
   */
  private rangesOverlap(range1: string, range2: string): boolean {
    // Simple check - in practice, you'd parse the addresses and check overlap
    return range1 === range2 || range2.includes(range1) || range1.includes(range2);
  }

  /**
   * Get comprehensive worksheet context
   */
  private async getWorksheetContext(worksheet: Excel.Worksheet, context: Excel.RequestContext): Promise<WorksheetContext> {
    // Load worksheet properties
    const usedRange = worksheet.getUsedRange();
    usedRange.load(['rowCount', 'columnCount']);
    
    // Load tables
    worksheet.tables.load(['items', 'count']);
    
    // Load charts
    worksheet.charts.load(['items', 'count']);
    
    await context.sync();

    // Get named ranges
    const namedRanges = await this.getNamedRanges(context);
    
    // Get table info
    const tables: TableInfo[] = [];
    for (const table of worksheet.tables.items) {
      table.load(['name', 'id']);
      const range = table.getRange();
      range.load('address');
      const headerRange = table.getHeaderRowRange();
      headerRange.load('values');
      await context.sync();
      
      tables.push({
        name: table.name,
        range: range.address,
        headers: headerRange.values[0].map(h => String(h)),
        rowCount: table.getDataBodyRange().getRowCount()
      });
    }

    // Get chart info
    const charts: ChartInfo[] = [];
    for (const chart of worksheet.charts.items) {
      chart.load(['name', 'chartType']);
      await context.sync();
      
      charts.push({
        name: chart.name || 'Unnamed Chart',
        type: chart.chartType,
        dataRange: '' // Would need more complex logic to get data range
      });
    }

    return {
      name: worksheet.name,
      totalRows: usedRange.rowCount || 0,
      totalColumns: usedRange.columnCount || 0,
      namedRanges,
      tables,
      charts,
      pivotTables: [] // Pivot tables require more complex detection
    };
  }

  /**
   * Get all named ranges in the workbook
   */
  private async getNamedRanges(context: Excel.RequestContext): Promise<NamedRangeInfo[]> {
    const namedRanges: NamedRangeInfo[] = [];
    
    try {
      const names = context.workbook.names;
      names.load('items');
      await context.sync();

      for (const namedItem of names.items) {
        namedItem.load(['name', 'scope', 'type']);
        try {
          const range = namedItem.getRange();
          range.load('address');
          await context.sync();
          
          namedRanges.push({
            name: namedItem.name,
            address: range.address,
            scope: namedItem.scope === 'Workbook' ? 'workbook' : 'worksheet'
          });
        } catch (e) {
          // Not a range reference
          continue;
        }
      }
    } catch (error) {
      console.error('Error loading named ranges:', error);
    }

    return namedRanges;
  }

  /**
   * Find related data based on various criteria
   */
  private async findRelatedData(
    primaryData: DetectedDataInfo,
    worksheetContext: WorksheetContext,
    context: Excel.RequestContext
  ): Promise<DetectedDataInfo[]> {
    const relatedData: DetectedDataInfo[] = [];

    // Find data in the same table
    for (const table of worksheetContext.tables) {
      if (this.rangesOverlap(primaryData.range, table.range)) {
        // This is the same table, look for other columns
        const tableRange = context.workbook.worksheets.getActiveWorksheet().tables.getItem(table.name).getRange();
        tableRange.load(['address', 'values', 'numberFormat']);
        await context.sync();
        
        const tableInfo = await this.detectDataInfo(tableRange, context);
        if (tableInfo.range !== primaryData.range) {
          relatedData.push(tableInfo);
        }
      }
    }

    // Find data with similar headers
    if (primaryData.hasHeaders && primaryData.headers.length > 0) {
      // This would require scanning other ranges, which could be expensive
      // For now, we'll check named ranges
      for (const namedRange of worksheetContext.namedRanges) {
        if (!primaryData.namedRanges.includes(namedRange.name)) {
          try {
            const range = context.workbook.names.getItem(namedRange.name).getRange();
            range.load(['address', 'values', 'numberFormat']);
            await context.sync();
            
            const rangeInfo = await this.detectDataInfo(range, context);
            if (rangeInfo.hasHeaders && this.headersMatch(primaryData.headers, rangeInfo.headers)) {
              relatedData.push(rangeInfo);
            }
          } catch (e) {
            continue;
          }
        }
      }
    }

    return relatedData;
  }

  /**
   * Check if headers match (allowing for some differences)
   */
  private headersMatch(headers1: string[], headers2: string[]): boolean {
    const set1 = new Set(headers1.map(h => h.toLowerCase()));
    const set2 = new Set(headers2.map(h => h.toLowerCase()));
    
    let matches = 0;
    for (const header of set1) {
      if (set2.has(header)) {
        matches++;
      }
    }

    // Consider related if at least 30% of headers match
    return matches >= Math.min(set1.size, set2.size) * 0.3;
  }

  /**
   * Calculate column uniqueness for header detection
   */
  private calculateColumnUniqueness(values: any[][]): {
    withHeaders: number;
    withoutHeaders: number;
  } {
    if (values.length < 2) {
      return { withHeaders: 0, withoutHeaders: 0 };
    }

    let uniquenessWithHeaders = 0;
    let uniquenessWithoutHeaders = 0;

    for (let col = 0; col < values[0].length; col++) {
      // With headers
      const columnWithHeaders = values.slice(1).map(row => row[col]);
      const uniqueWithHeaders = new Set(columnWithHeaders).size;
      uniquenessWithHeaders += uniqueWithHeaders / columnWithHeaders.length;

      // Without headers
      const columnWithoutHeaders = values.map(row => row[col]);
      const uniqueWithoutHeaders = new Set(columnWithoutHeaders).size;
      uniquenessWithoutHeaders += uniqueWithoutHeaders / columnWithoutHeaders.length;
    }

    return {
      withHeaders: uniquenessWithHeaders / values[0].length,
      withoutHeaders: uniquenessWithoutHeaders / values[0].length
    };
  }

  /**
   * Detect the format pattern from number formats
   */
  private detectFormat(numberFormats: any[][]): string {
    const formats = numberFormats.flat().filter(f => f && f !== 'General');
    if (formats.length === 0) {
      return 'General';
    }

    // Find most common format
    const formatCount = new Map<string, number>();
    formats.forEach(format => {
      formatCount.set(format, (formatCount.get(format) || 0) + 1);
    });

    let maxCount = 0;
    let mostCommonFormat = 'General';
    
    for (const [format, count] of formatCount.entries()) {
      if (count > maxCount) {
        maxCount = count;
        mostCommonFormat = format;
      }
    }

    return mostCommonFormat;
  }

  /**
   * Generate suggestions based on detected data
   */
  private generateDataSuggestions(
    typeDetection: { overallType: DataType; columnTypes: DataType[] },
    patterns: DataPattern[],
    statistics?: DataStatistics
  ): string[] {
    const suggestions: string[] = [];

    // Type-based suggestions
    if (typeDetection.overallType === 'numeric' || typeDetection.columnTypes.includes('numeric')) {
      suggestions.push('Consider creating a chart to visualize numeric trends');
      suggestions.push('Use ANALYZE function for statistical insights');
    }

    if (typeDetection.columnTypes.includes('date')) {
      suggestions.push('Create a timeline visualization for date data');
      suggestions.push('Use FORECAST function for time-series predictions');
    }

    if (typeDetection.columnTypes.includes('currency')) {
      suggestions.push('Apply currency formatting for better readability');
      suggestions.push('Use FINMODEL function for financial analysis');
    }

    // Pattern-based suggestions
    patterns.forEach(pattern => {
      switch (pattern.type) {
        case 'email':
          suggestions.push('Validate email addresses for data quality');
          break;
        case 'phone':
          suggestions.push('Standardize phone number format');
          break;
        case 'url':
          suggestions.push('Check URL validity and accessibility');
          break;
      }
    });

    // Statistics-based suggestions
    if (statistics) {
      if (statistics.stdDev && statistics.average && statistics.stdDev > statistics.average * 0.5) {
        suggestions.push('High variance detected - check for outliers');
      }
    }

    return suggestions;
  }

  /**
   * Generate context-aware suggestions
   */
  private generateSuggestions(
    primaryData: DetectedDataInfo,
    relatedData: DetectedDataInfo[],
    worksheetContext: WorksheetContext
  ): ContextSuggestion[] {
    const suggestions: ContextSuggestion[] = [];

    // Analysis suggestions based on data type
    if (primaryData.dataType === 'numeric' || primaryData.dataType === 'currency') {
      suggestions.push({
        type: 'analysis',
        description: 'Perform statistical analysis to find trends and patterns',
        confidence: 0.9,
        action: async () => {
          // Trigger AI analysis
          console.log('Triggering statistical analysis');
        }
      });
    }

    // Formatting suggestions
    if (primaryData.dataType === 'currency' && primaryData.format === 'General') {
      suggestions.push({
        type: 'formatting',
        description: 'Apply currency formatting for better readability',
        confidence: 0.95,
        action: async () => {
          await Excel.run(async (context) => {
            const range = context.workbook.worksheets.getActiveWorksheet().getRange(primaryData.range);
            range.numberFormat = [['$#,##0.00']];
            await context.sync();
          });
        }
      });
    }

    // Validation suggestions
    if (primaryData.patterns.some(p => p.type === 'email')) {
      suggestions.push({
        type: 'validation',
        description: 'Add email validation to ensure data quality',
        confidence: 0.85,
        action: async () => {
          // Add data validation
          console.log('Adding email validation');
        }
      });
    }

    // Relationship suggestions
    if (relatedData.length > 0) {
      suggestions.push({
        type: 'relationship',
        description: `Found ${relatedData.length} related data ranges. Consider creating a dashboard.`,
        confidence: 0.8,
        action: async () => {
          // Create dashboard
          console.log('Creating dashboard with related data');
        }
      });
    }

    // Table suggestions
    if (!worksheetContext.tables.some(t => this.rangesOverlap(primaryData.range, t.range))) {
      suggestions.push({
        type: 'formatting',
        description: 'Convert data to Excel Table for better functionality',
        confidence: 0.7,
        action: async () => {
          await Excel.run(async (context) => {
            const worksheet = context.workbook.worksheets.getActiveWorksheet();
            const range = worksheet.getRange(primaryData.range);
            const table = worksheet.tables.add(range, primaryData.hasHeaders);
            table.name = `Table_${Date.now()}`;
            table.style = 'TableStyleMedium2';
            await context.sync();
          });
        }
      });
    }

    return suggestions.sort((a, b) => b.confidence - a.confidence);
  }

  /**
   * Clear the detection cache
   */
  clearCache(): void {
    this.cache.clear();
  }

  /**
   * Get cached detection info for a range
   */
  getCachedInfo(rangeAddress: string): DetectedDataInfo | undefined {
    return this.cache.get(rangeAddress);
  }
}

// Export singleton instance
export const smartDetection = new SmartDetectionService();