/**
 * Excel Helper Functions
 * Utilities for working with Excel data and objects
 */

/* global Excel */

export class ExcelHelpers {
  /**
   * Get currently selected data from Excel
   */
  async getSelectedData(): Promise<any[][] | null> {
    try {
      return await Excel.run(async (context) => {
        const selection = context.workbook.getSelectedRange();
        selection.load('values');
        await context.sync();
        return selection.values;
      });
    } catch (error) {
      console.error('Error getting selection:', error);
      return null;
    }
  }

  /**
   * Get selected range address
   */
  async getSelectedRange(): Promise<string | null> {
    try {
      return await Excel.run(async (context) => {
        const selection = context.workbook.getSelectedRange();
        selection.load('address');
        await context.sync();
        return selection.address;
      });
    } catch (error) {
      console.error('Error getting selection:', error);
      return null;
    }
  }

  /**
   * Insert data at current selection
   */
  async insertData(data: any[][]): Promise<void> {
    try {
      await Excel.run(async (context) => {
        const selection = context.workbook.getSelectedRange();
        selection.values = data;
        await context.sync();
      });
    } catch (error) {
      console.error('Error inserting data:', error);
      throw error;
    }
  }

  /**
   * Create a new worksheet
   */
  async createWorksheet(name: string): Promise<Excel.Worksheet> {
    return await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      
      // Check if sheet exists and delete if it does
      try {
        const existingSheet = sheets.getItem(name);
        existingSheet.delete();
        await context.sync();
      } catch (e) {
        // Sheet doesn't exist, which is fine
      }
      
      const newSheet = sheets.add(name);
      newSheet.activate();
      await context.sync();
      
      return newSheet;
    });
  }

  /**
   * Get all worksheet names
   */
  async getWorksheetNames(): Promise<string[]> {
    return await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load('items/name');
      await context.sync();
      
      return sheets.items.map(sheet => sheet.name);
    });
  }

  /**
   * Create a chart
   */
  async createChart(
    worksheet: Excel.Worksheet,
    dataRange: Excel.Range,
    chartType: Excel.ChartType,
    position: { top: number; left: number; height: number; width: number }
  ): Promise<Excel.Chart> {
    const chart = worksheet.charts.add(chartType, dataRange, Excel.ChartSeriesBy.auto);
    chart.top = position.top;
    chart.left = position.left;
    chart.height = position.height;
    chart.width = position.width;
    
    return chart;
  }

  /**
   * Apply conditional formatting
   */
  async applyConditionalFormatting(
    range: Excel.Range,
    type: 'dataBar' | 'colorScale' | 'iconSet',
    options?: any
  ): Promise<void> {
    await Excel.run(async (context) => {
      let format: any;
      
      switch (type) {
        case 'dataBar':
          format = range.conditionalFormats.add(Excel.ConditionalFormatType.dataBar);
          if (options?.color) {
            format.dataBar.barDirection = Excel.ConditionalDataBarDirection.context;
            format.dataBar.showDataBarOnly = false;
          }
          break;
          
        case 'colorScale':
          format = range.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
          format.colorScale.criteria = {
            minimum: { 
              color: options?.minColor || '#F8696B', 
              type: Excel.ConditionalFormatColorCriterionType.lowestValue 
            },
            midpoint: { 
              color: options?.midColor || '#FFEB84', 
              type: Excel.ConditionalFormatColorCriterionType.percentile, 
              value: 50 
            },
            maximum: { 
              color: options?.maxColor || '#63BE7B', 
              type: Excel.ConditionalFormatColorCriterionType.highestValue 
            }
          };
          break;
          
        case 'iconSet':
          format = range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);
          format.iconSet.style = options?.style || Excel.IconSet.threeTrafficLights1;
          break;
      }
      
      await context.sync();
    });
  }

  /**
   * Format range as table
   */
  async formatAsTable(
    worksheet: Excel.Worksheet,
    range: string,
    tableName: string,
    style: string = 'TableStyleMedium2'
  ): Promise<Excel.Table> {
    return await Excel.run(async (context) => {
      const dataRange = worksheet.getRange(range);
      const table = worksheet.tables.add(dataRange, true);
      table.name = tableName;
      table.style = style;
      
      await context.sync();
      return table;
    });
  }

  /**
   * Get data from a named range
   */
  async getNamedRangeData(name: string): Promise<any[][] | null> {
    try {
      return await Excel.run(async (context) => {
        const namedRange = context.workbook.names.getItem(name);
        const range = namedRange.getRange();
        range.load('values');
        await context.sync();
        return range.values;
      });
    } catch (error) {
      console.error('Error getting named range:', error);
      return null;
    }
  }

  /**
   * Create a named range
   */
  async createNamedRange(name: string, rangeAddress: string): Promise<void> {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(rangeAddress);
      
      context.workbook.names.add(name, range);
      await context.sync();
    });
  }

  /**
   * Auto-fit columns
   */
  async autoFitColumns(worksheet: Excel.Worksheet): Promise<void> {
    await Excel.run(async (context) => {
      const usedRange = worksheet.getUsedRange();
      if (usedRange) {
        usedRange.format.autofitColumns();
        await context.sync();
      }
    });
  }

  /**
   * Apply number format
   */
  async applyNumberFormat(range: Excel.Range, format: string): Promise<void> {
    await Excel.run(async (context) => {
      range.numberFormat = [[format]];
      await context.sync();
    });
  }

  /**
   * Get unique values from a column
   */
  async getUniqueValues(data: any[][]): Promise<any[]> {
    const flatData = data.flat();
    return [...new Set(flatData)];
  }

  /**
   * Calculate statistics for numerical data
   */
  calculateStatistics(data: number[]): {
    count: number;
    sum: number;
    average: number;
    min: number;
    max: number;
    median: number;
    stdDev: number;
  } {
    const count = data.length;
    const sum = data.reduce((a, b) => a + b, 0);
    const average = sum / count;
    const min = Math.min(...data);
    const max = Math.max(...data);
    
    // Calculate median
    const sorted = [...data].sort((a, b) => a - b);
    const median = count % 2 === 0
      ? (sorted[count / 2 - 1] + sorted[count / 2]) / 2
      : sorted[Math.floor(count / 2)];
    
    // Calculate standard deviation
    const squaredDiffs = data.map(x => Math.pow(x - average, 2));
    const avgSquaredDiff = squaredDiffs.reduce((a, b) => a + b, 0) / count;
    const stdDev = Math.sqrt(avgSquaredDiff);
    
    return { count, sum, average, min, max, median, stdDev };
  }

  /**
   * Detect data types in columns
   */
  detectColumnTypes(data: any[][]): string[] {
    if (data.length === 0) return [];
    
    const headers = data[0];
    const columnTypes: string[] = [];
    
    for (let col = 0; col < headers.length; col++) {
      const columnData = data.slice(1).map(row => row[col]).filter(val => val != null);
      
      if (columnData.length === 0) {
        columnTypes.push('empty');
        continue;
      }
      
      // Check if all values are numbers
      if (columnData.every(val => !isNaN(Number(val)))) {
        columnTypes.push('number');
      }
      // Check if all values are dates
      else if (columnData.every(val => !isNaN(Date.parse(String(val))))) {
        columnTypes.push('date');
      }
      // Check if boolean
      else if (columnData.every(val => 
        typeof val === 'boolean' || 
        ['true', 'false', 'yes', 'no'].includes(String(val).toLowerCase())
      )) {
        columnTypes.push('boolean');
      }
      // Default to text
      else {
        columnTypes.push('text');
      }
    }
    
    return columnTypes;
  }

  /**
   * Create a pivot table
   */
  async createPivotTable(
    sourceData: string,
    targetCell: string,
    config: {
      rows: string[];
      columns: string[];
      values: Array<{ field: string; operation: string }>;
    }
  ): Promise<void> {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      
      // This is a placeholder as Office.js doesn't directly support pivot table creation
      // In practice, you would use VBA or Office Scripts for this
      const message = `Pivot Table Configuration:
      Source: ${sourceData}
      Target: ${targetCell}
      Rows: ${config.rows.join(', ')}
      Columns: ${config.columns.join(', ')}
      Values: ${config.values.map(v => `${v.field} (${v.operation})`).join(', ')}`;
      
      sheet.getRange(targetCell).values = [[message]];
      await context.sync();
    });
  }
}