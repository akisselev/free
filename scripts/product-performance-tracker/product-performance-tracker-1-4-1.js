/**
 * ====================================
 * PRODUCT PERFORMANCE TRACKER v1.4.1
 * ====================================
 * 
 * A Google Ads script that analyzes month-over-month product impression activity changes
 * by automatically pulling Shopping product data from Google Ads.
 * 
 * 
 * Features:
 * - Automatically pulls product data from Google Ads (no manual uploads needed)
 * - Captures account currency for proper formatting
 * - Identifies newly active, inactive, and continuing products
 * - Analyzes performance impact of impression activity changes
 * - Provides category-level insights by product type
 * - Generates comprehensive comparison reports with proper formatting
 * - AI-generated business insights in plain language
 * - Professional formatting with conditional highlighting
 * 
 * Data Source:
 * Automatically pulls Shopping product data directly from Google Ads for:
 * - Last calendar month (e.g., if run in March, analyzes February 1-28/29)
 * - Two calendar months ago (e.g., if run in March, analyzes January 1-31)
 * 
 * IMPORTANT: Uses COMPLETE CALENDAR MONTHS, not rolling 30-day periods.
 * This ensures consistent month-over-month comparisons with complete data.
 * 
 * Collected metrics:
 * - Product ID, Title, Types (1st, 2nd, 3rd level)
 * - Impressions, Clicks, Conversions (formatted to 2 decimals)
 * - Conversion Value, Cost (formatted with currency symbol and 2 decimals)
 * - Conv. value / cost (calculated and formatted)
 * - Account Currency
 * 
 * Installation Steps:
 * 1. (Optional) Set SHEET_URL to your Google Sheet URL, or leave empty to auto-create
 * 2. (Optional) Set OPENAI_API_KEY to your OpenAI API key for AI insights
 * 3. Run this script in Google Ads Scripts
 * 4. Script will automatically pull data and generate reports
 * 5. If new spreadsheet created, save the URL from logs for future use
 * 
 * Author: Andrey Kisselev
 * Version: 1.4.1
 * ====================================
 */

// CONFIGURATION - UPDATE THESE VALUES
const SHEET_URL = ''; // Leave empty to auto-create a new spreadsheet, or enter your Google Sheet URL here
const OPENAI_API_KEY = ''; // Enter your OpenAI API key here

// Date ranges are calculated dynamically in calculateDateRanges() function
// to ensure we get complete calendar months (not rolling 30-day periods)

// Output tab names (will be created by script)
const OUTPUT_TABS = {
  ANALYSIS: 'Impression Activity Analysis',
  NEW_PRODUCTS: 'Newly Active Products',
  DISCONTINUED: 'Inactive Products', 
  PERFORMANCE_CHANGES: 'Performance Changes',
  EXECUTIVE_INSIGHTS: 'Executive Insights',
  RAW_DATA_LAST_MONTH: 'Raw Data - Last Month',
  RAW_DATA_TWO_MONTHS_AGO: 'Raw Data - Two Months Ago'
};

// Global variable to store account currency
let ACCOUNT_CURRENCY = 'USD';

function main() {
  try {
   
    Logger.log('Thanks for using Product Performance Tracker by Andrey Kisselev (c) 2025 version 1.4.1');
    Logger.log('🚀 STARTING PRODUCT PERFORMANCE TRACKER v1.4.1');
   
    
    // Get account currency first
    Logger.log('💰 Getting account currency...');
    ACCOUNT_CURRENCY = getAccountCurrency();
    Logger.log(`✅ Account currency: ${ACCOUNT_CURRENCY}`);
    
    // Get or create spreadsheet
    Logger.log('📊 Getting spreadsheet...');
    const spreadsheet = getOrCreateSpreadsheet();
    Logger.log('✅ Spreadsheet ready');
    
    // Calculate date ranges
    const dateRanges = calculateDateRanges();
    Logger.log(`📅 Date ranges: Last calendar month (${dateRanges.lastMonth.start} to ${dateRanges.lastMonth.end}), Two calendar months ago (${dateRanges.twoMonthsAgo.start} to ${dateRanges.twoMonthsAgo.end})`);
    Logger.log(`📅 IMPORTANT: Using COMPLETE CALENDAR MONTHS, not rolling 30-day periods`);
    
    // Pull data automatically from Google Ads
    Logger.log('📖 Pulling last month data from Google Ads...');
    const lastMonthData = pullShoppingProductData(dateRanges.lastMonth, 'Last Month');
    
    Logger.log('📖 Pulling two months ago data from Google Ads...');
    const twoMonthsAgoData = pullShoppingProductData(dateRanges.twoMonthsAgo, 'Two Months Ago');
    
    if (!lastMonthData || !twoMonthsAgoData) {
      Logger.log('❌ ERROR: Failed to pull data from Google Ads');
      return;
    }
    
    // Save raw data to spreadsheet for reference
    Logger.log('💾 Saving raw data to spreadsheet...');
    saveRawDataToSheet(spreadsheet, lastMonthData, OUTPUT_TABS.RAW_DATA_LAST_MONTH);
    saveRawDataToSheet(spreadsheet, twoMonthsAgoData, OUTPUT_TABS.RAW_DATA_TWO_MONTHS_AGO);
    
    // Perform product performance comparison analysis
    Logger.log('🔍 Analyzing product performance changes...');
    const analysis = performInventoryAnalysis(lastMonthData, twoMonthsAgoData);
    
    // Create comprehensive reports
    Logger.log('📊 Creating analysis reports...');
    createAnalysisReports(spreadsheet, analysis);
    
    // Generate AI-powered business insights
    Logger.log('🤖 Generating AI-powered business insights...');
    generateExecutiveInsights(spreadsheet, analysis);
    
    Logger.log('='.repeat(60));
    Logger.log('🎉 SUCCESS! AUTOMATED PRODUCT PERFORMANCE ANALYSIS COMPLETED');
    Logger.log('='.repeat(60));
    Logger.log(`✅ Analyzed ${analysis.summary.totalProductsLastMonth} products from last month`);
    Logger.log(`✅ Analyzed ${analysis.summary.totalProductsTwoMonthsAgo} products from two months ago`);
    Logger.log(`📈 Found ${analysis.summary.newProducts} newly active products`);
    Logger.log(`📉 Found ${analysis.summary.discontinuedProducts} inactive products`);
    Logger.log(`🔄 Found ${analysis.summary.continuingProducts} continuing products`);
    Logger.log(`💰 Currency: ${ACCOUNT_CURRENCY}`);
    Logger.log(`📊 Spreadsheet: ${spreadsheet.getUrl()}`);
    Logger.log('='.repeat(60));
    
  } catch (error) {
    Logger.log('='.repeat(60));
    Logger.log('❌ FATAL ERROR');
    Logger.log('='.repeat(60));
    Logger.log('Error: ' + error.toString());
    Logger.log('Stack trace: ' + error.stack);
    Logger.log('='.repeat(60));
  }
}

/**
 * Gets the account currency code
 */
function getAccountCurrency() {
  try {
    const account = AdsApp.currentAccount();
    return account.getCurrencyCode() || 'USD';
  } catch (error) {
    Logger.log(`⚠️ Warning: Could not get account currency, defaulting to USD: ${error.message}`);
    return 'USD';
  }
}

/**
 * Gets existing spreadsheet or creates a new one if SHEET_URL is empty
 */
function getOrCreateSpreadsheet() {
  try {
    // If SHEET_URL is provided and not empty, try to open it
    if (SHEET_URL && SHEET_URL.trim() !== '') {
      Logger.log('📋 Opening existing spreadsheet from provided URL...');
      try {
        const spreadsheet = SpreadsheetApp.openByUrl(SHEET_URL);
        Logger.log(`✅ Successfully opened existing spreadsheet: ${spreadsheet.getName()}`);
        return spreadsheet;
      } catch (error) {
        Logger.log(`⚠️ Warning: Could not open spreadsheet from provided URL: ${error.message}`);
        Logger.log('📋 Creating new spreadsheet instead...');
      }
    }
    
    // Create a new spreadsheet
    Logger.log('📋 Creating new spreadsheet...');
    const timestamp = Utilities.formatDate(new Date(), AdsApp.currentAccount().getTimeZone(), 'yyyy-MM-dd_HH-mm');
    const spreadsheetName = `Product Performance Analysis - ${timestamp}`;
    
    const spreadsheet = SpreadsheetApp.create(spreadsheetName);
    
    // Create our first analysis sheet, then remove the default sheet
    const analysisSheet = spreadsheet.insertSheet(OUTPUT_TABS.ANALYSIS);
    
    // Now safely remove the default sheet
    const defaultSheet = spreadsheet.getSheets().find(sheet => sheet.getName() === 'Sheet1');
    if (defaultSheet) {
      spreadsheet.deleteSheet(defaultSheet);
    }
    
    Logger.log(`✅ Created new spreadsheet: ${spreadsheetName}`);
    Logger.log(`📊 Spreadsheet URL: ${spreadsheet.getUrl()}`);
    Logger.log('💡 TIP: Save this URL to use the same spreadsheet next time!');
    
    return spreadsheet;
    
  } catch (error) {
    Logger.log(`❌ Error getting/creating spreadsheet: ${error.message}`);
    throw error;
  }
}

/**
 * Calculates date ranges for last calendar month and two calendar months ago
 * 
 * IMPORTANT: This calculates COMPLETE CALENDAR MONTHS, not rolling 30-day periods.
 * For example, if run in March 2024:
 * - Last month = February 1-29, 2024 (complete February)
 * - Two months ago = January 1-31, 2024 (complete January)
 * 
 * This ensures consistent month-over-month comparisons with complete data.
 */
function calculateDateRanges() {
  const today = new Date();
  const timeZone = AdsApp.currentAccount().getTimeZone();
  
  // Last calendar month (e.g., if today is March 15, this gets February 1-28/29)
  const lastMonthStart = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  const lastMonthEnd = new Date(today.getFullYear(), today.getMonth(), 0); // Last day of previous month
  
  // Two calendar months ago (e.g., if today is March 15, this gets January 1-31)
  const twoMonthsAgoStart = new Date(today.getFullYear(), today.getMonth() - 2, 1);
  const twoMonthsAgoEnd = new Date(today.getFullYear(), today.getMonth() - 1, 0); // Last day of month before last
  
  return {
    lastMonth: {
      start: Utilities.formatDate(lastMonthStart, timeZone, 'yyyy-MM-dd'),
      end: Utilities.formatDate(lastMonthEnd, timeZone, 'yyyy-MM-dd')
    },
    twoMonthsAgo: {
      start: Utilities.formatDate(twoMonthsAgoStart, timeZone, 'yyyy-MM-dd'),
      end: Utilities.formatDate(twoMonthsAgoEnd, timeZone, 'yyyy-MM-dd')
    }
  };
}

/**
 * Pulls Shopping product data from Google Ads for a specific date range using Reports API
 */
function pullShoppingProductData(dateRange, periodName) {
  try {
    Logger.log(`📊 Pulling ${periodName} data from ${dateRange.start} to ${dateRange.end}...`);
    
    const productData = new Map();
    
    // Use Google Ads Reporting API to get Shopping product data
    const reportQuery = `
      SELECT 
        segments.product_item_id,
        segments.product_title,
        segments.product_type_l1,
        segments.product_type_l2,
        segments.product_type_l3,
        metrics.impressions,
        metrics.clicks,
        metrics.conversions,
        metrics.conversions_value,
        metrics.cost_micros
      FROM shopping_performance_view 
      WHERE segments.date BETWEEN '${dateRange.start}' AND '${dateRange.end}'
    `;
    
    Logger.log(`📊 Running Shopping product report for ${periodName}...`);
    
    const report = AdsApp.report(reportQuery);
    const rows = report.rows();
    let processedCount = 0;
    
    while (rows.hasNext()) {
      const row = rows.next();
      
      try {
        // Get product identifiers
        const productId = row['segments.product_item_id'] || `Unknown_${processedCount}`;
        const productTitle = row['segments.product_title'] || '';
        
        // Get product categories
        const productTypeL1 = row['segments.product_type_l1'] || '';
        const productTypeL2 = row['segments.product_type_l2'] || '';
        const productTypeL3 = row['segments.product_type_l3'] || '';
        
        // Get performance metrics
        const impressions = parseInt(row['metrics.impressions']) || 0;
        const clicks = parseInt(row['metrics.clicks']) || 0;
        const conversions = parseFloat(row['metrics.conversions']) || 0;
        const conversionValue = parseFloat(row['metrics.conversions_value']) || 0;
        const costMicros = parseInt(row['metrics.cost_micros']) || 0;
        const cost = costMicros / 1000000; // Convert micros to actual currency
        
        // Calculate conv. value / cost
        const convValuePerCost = cost > 0 ? (conversionValue / cost) : 0;
        
        const productRecord = {
          productId: productId.toString(),
          productTitle: productTitle,
          impressions: impressions,
          clicks: clicks,
          productTypeL1: productTypeL1,
          productTypeL2: productTypeL2,
          productTypeL3: productTypeL3,
          conversions: conversions,
          conversionValue: conversionValue,
          cost: cost,
          convValuePerCost: convValuePerCost,
          currency: ACCOUNT_CURRENCY
        };
        
        // If we already have this product, aggregate the metrics
        if (productData.has(productId.toString())) {
          const existing = productData.get(productId.toString());
          existing.impressions += impressions;
          existing.clicks += clicks;
          existing.conversions += conversions;
          existing.conversionValue += conversionValue;
          existing.cost += cost;
          existing.convValuePerCost = existing.cost > 0 ? (existing.conversionValue / existing.cost) : 0;
        } else {
          productData.set(productId.toString(), productRecord);
        }
        
        processedCount++;
        
        // Log progress every 1000 rows
        if (processedCount % 1000 === 0) {
          Logger.log(`📊 Processed ${processedCount} rows for ${periodName}...`);
        }
        
      } catch (error) {
        Logger.log(`⚠️ Error processing row in ${periodName}: ${error}`);
      }
    }
    
    Logger.log(`✅ ${periodName}: Processed ${processedCount} rows, found ${productData.size} unique products`);
    
    return {
      data: productData,
      sheetName: periodName,
      dateRange: dateRange
    };
    
  } catch (error) {
    Logger.log(`❌ Error pulling ${periodName} data from Google Ads: ${error}`);
    Logger.log(`❌ Error details: ${error.message}`);
    return null;
  }
}

/**
 * Saves raw data to a sheet for reference with proper formatting
 */
function saveRawDataToSheet(spreadsheet, productData, sheetName) {
  try {
    let sheet = spreadsheet.getSheetByName(sheetName);
    if (sheet) {
      sheet.clear();
    } else {
      sheet = spreadsheet.insertSheet(sheetName);
    }
    
    // Create headers
    const headers = [
      'Product ID',
      'Product Title',
      'Impressions',
      'Clicks',
      'Product Type L1',
      'Product Type L2',
      'Product Type L3',
      'Conversions',
      'Conversion Value',
      'Cost',
      'Conv. Value / Cost',
      'Currency'
    ];
    
    // Create data rows
    const dataRows = [headers];
    for (const product of productData.data.values()) {
      dataRows.push([
        product.productId,
        product.productTitle,
        product.impressions,
        product.clicks,
        product.productTypeL1,
        product.productTypeL2,
        product.productTypeL3,
        product.conversions,
        product.conversionValue,
        product.cost,
        product.convValuePerCost,
        product.currency
      ]);
    }
    
    // Write data
    if (dataRows.length > 1) {
      const range = sheet.getRange(1, 1, dataRows.length, headers.length);
      range.setValues(dataRows);
      
      // Format headers
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#e8f0fe');
      headerRange.setFontColor('#1a73e8');
      
      // Format specific columns with proper formatting
      if (dataRows.length > 1) {
        // Conversions - 2 decimal places
        sheet.getRange(2, 8, dataRows.length - 1, 1).setNumberFormat('#,##0.00');
        
        // Conversion Value - currency with 2 decimal places
        const currencyFormat = getCurrencyFormat(ACCOUNT_CURRENCY);
        sheet.getRange(2, 9, dataRows.length - 1, 1).setNumberFormat(currencyFormat);
        
        // Cost - currency with 2 decimal places
        sheet.getRange(2, 10, dataRows.length - 1, 1).setNumberFormat(currencyFormat);
        
        // Conv. Value / Cost - 2 decimal places
        sheet.getRange(2, 11, dataRows.length - 1, 1).setNumberFormat('#,##0.00');
      }
      
      // Auto-resize columns
      sheet.autoResizeColumns(1, headers.length);
      
      // Freeze header row
      sheet.setFrozenRows(1);
    }
    
    Logger.log(`✅ Raw data saved to ${sheetName} (${productData.data.size} products) with ${ACCOUNT_CURRENCY} formatting`);
    
  } catch (error) {
    Logger.log(`❌ Error saving raw data to ${sheetName}: ${error}`);
  }
}

/**
 * Returns the appropriate currency format for the given currency code
 */
function getCurrencyFormat(currencyCode) {
  const currencyFormats = {
    'USD': '$#,##0.00',
    'EUR': '€#,##0.00',
    'GBP': '£#,##0.00',
    'JPY': '¥#,##0',
    'CAD': 'C$#,##0.00',
    'AUD': 'A$#,##0.00',
    'CHF': 'CHF#,##0.00',
    'SEK': '#,##0.00 kr',
    'NOK': '#,##0.00 kr',
    'DKK': '#,##0.00 kr'
  };
  
  return currencyFormats[currencyCode] || `${currencyCode} #,##0.00`;
}

/**
 * Performs comprehensive product performance analysis
 */
function performInventoryAnalysis(lastMonthData, twoMonthsAgoData) {
  Logger.log('📊 Starting comprehensive product performance analysis...');
  
  const lastMonthProducts = lastMonthData.data;
  const twoMonthsAgoProducts = twoMonthsAgoData.data;
  
  // Get all unique product IDs
  const allLastMonthIds = new Set(lastMonthProducts.keys());
  const allTwoMonthsAgoIds = new Set(twoMonthsAgoProducts.keys());
  const allProductIds = new Set([...allLastMonthIds, ...allTwoMonthsAgoIds]);
  
  // Categorize products
  const newProducts = new Map();
  const discontinuedProducts = new Map();
  const continuingProducts = new Map();
  
  // Analyze each product
  for (const productId of allProductIds) {
    const lastMonthProduct = lastMonthProducts.get(productId);
    const twoMonthsAgoProduct = twoMonthsAgoProducts.get(productId);
    
    if (lastMonthProduct && !twoMonthsAgoProduct) {
      // Product that started getting impressions
      newProducts.set(productId, lastMonthProduct);
    } else if (!lastMonthProduct && twoMonthsAgoProduct) {
      // Product that stopped getting impressions
      discontinuedProducts.set(productId, twoMonthsAgoProduct);
    } else if (lastMonthProduct && twoMonthsAgoProduct) {
      // Continuing product - calculate performance changes
      const changes = {
        productId: productId,
        lastMonth: lastMonthProduct,
        twoMonthsAgo: twoMonthsAgoProduct,
        changes: {
          impressionsDelta: lastMonthProduct.impressions - twoMonthsAgoProduct.impressions,
          clicksDelta: lastMonthProduct.clicks - twoMonthsAgoProduct.clicks,
          conversionsDelta: lastMonthProduct.conversions - twoMonthsAgoProduct.conversions,
          conversionValueDelta: lastMonthProduct.conversionValue - twoMonthsAgoProduct.conversionValue,
          convValuePerCostDelta: lastMonthProduct.convValuePerCost - twoMonthsAgoProduct.convValuePerCost,
          costDelta: lastMonthProduct.cost - twoMonthsAgoProduct.cost
        }
      };
      
      // Calculate percentage changes
      changes.changes.impressionsChange = twoMonthsAgoProduct.impressions > 0 ? 
        ((lastMonthProduct.impressions - twoMonthsAgoProduct.impressions) / twoMonthsAgoProduct.impressions) * 100 : 0;
      changes.changes.clicksChange = twoMonthsAgoProduct.clicks > 0 ? 
        ((lastMonthProduct.clicks - twoMonthsAgoProduct.clicks) / twoMonthsAgoProduct.clicks) * 100 : 0;
      changes.changes.conversionsChange = twoMonthsAgoProduct.conversions > 0 ? 
        ((lastMonthProduct.conversions - twoMonthsAgoProduct.conversions) / twoMonthsAgoProduct.conversions) * 100 : 0;
      changes.changes.conversionValueChange = twoMonthsAgoProduct.conversionValue > 0 ? 
        ((lastMonthProduct.conversionValue - twoMonthsAgoProduct.conversionValue) / twoMonthsAgoProduct.conversionValue) * 100 : 0;
      changes.changes.convValuePerCostChange = twoMonthsAgoProduct.convValuePerCost > 0 ? 
        ((lastMonthProduct.convValuePerCost - twoMonthsAgoProduct.convValuePerCost) / twoMonthsAgoProduct.convValuePerCost) * 100 : 0;
      changes.changes.costChange = twoMonthsAgoProduct.cost > 0 ? 
        ((lastMonthProduct.cost - twoMonthsAgoProduct.cost) / twoMonthsAgoProduct.cost) * 100 : 0;
      
      continuingProducts.set(productId, changes);
    }
  }
  
  // Calculate summary metrics
  const summary = {
    totalProductsLastMonth: lastMonthProducts.size,
    totalProductsTwoMonthsAgo: twoMonthsAgoProducts.size,
    newProducts: newProducts.size,
    discontinuedProducts: discontinuedProducts.size,
    continuingProducts: continuingProducts.size,
    currency: ACCOUNT_CURRENCY,
    
    // Performance aggregates for newly active products
    newProductsImpressions: Array.from(newProducts.values()).reduce((sum, p) => sum + p.impressions, 0),
    newProductsClicks: Array.from(newProducts.values()).reduce((sum, p) => sum + p.clicks, 0),
    newProductsConversions: Array.from(newProducts.values()).reduce((sum, p) => sum + p.conversions, 0),
    newProductsConversionValue: Array.from(newProducts.values()).reduce((sum, p) => sum + p.conversionValue, 0),
    newProductsCost: Array.from(newProducts.values()).reduce((sum, p) => sum + p.cost, 0),
    
    // Performance aggregates for inactive products
    discontinuedProductsImpressions: Array.from(discontinuedProducts.values()).reduce((sum, p) => sum + p.impressions, 0),
    discontinuedProductsClicks: Array.from(discontinuedProducts.values()).reduce((sum, p) => sum + p.clicks, 0),
    discontinuedProductsConversions: Array.from(discontinuedProducts.values()).reduce((sum, p) => sum + p.conversions, 0),
    discontinuedProductsConversionValue: Array.from(discontinuedProducts.values()).reduce((sum, p) => sum + p.conversionValue, 0),
    discontinuedProductsCost: Array.from(discontinuedProducts.values()).reduce((sum, p) => sum + p.cost, 0)
  };
  
  Logger.log(`📊 Analysis complete: ${summary.newProducts} newly active, ${summary.discontinuedProducts} inactive, ${summary.continuingProducts} continuing`);
  
  return {
    summary: summary,
    newProducts: newProducts,
    discontinuedProducts: discontinuedProducts,
    continuingProducts: continuingProducts
  };
}

/**
 * Creates comprehensive analysis reports
 */
function createAnalysisReports(spreadsheet, analysis) {
  Logger.log('📋 Creating comprehensive analysis reports...');
  
  // Create main analysis summary
  createAnalysisSummary(spreadsheet, analysis);
  
  // Create detailed reports for each category
  createNewProductsReport(spreadsheet, analysis.newProducts);
  createDiscontinuedProductsReport(spreadsheet, analysis.discontinuedProducts);
  createPerformanceChangesReport(spreadsheet, analysis.continuingProducts);
  
  Logger.log('✅ All analysis reports created successfully');
}

/**
 * Creates the main analysis summary report
 */
function createAnalysisSummary(spreadsheet, analysis) {
  let summarySheet = spreadsheet.getSheetByName(OUTPUT_TABS.ANALYSIS);
  if (summarySheet) {
    summarySheet.clear();
  } else {
    summarySheet = spreadsheet.insertSheet(OUTPUT_TABS.ANALYSIS);
  }
  
  const summary = analysis.summary;
  const currencyFormat = getCurrencyFormat(ACCOUNT_CURRENCY);
  
  // Create summary data
  const summaryData = [
    [`Impression Activity Analysis Summary (Automated) - ${ACCOUNT_CURRENCY}`, ''],
    ['', ''],
    ['📊 IMPRESSION ACTIVITY OVERVIEW', ''],
    ['Total Products - Last Month', summary.totalProductsLastMonth],
    ['Total Products - Two Months Ago', summary.totalProductsTwoMonthsAgo],
    ['Net Change in Active Products', summary.totalProductsLastMonth - summary.totalProductsTwoMonthsAgo],
    ['', ''],
    ['📈 IMPRESSION ACTIVITY CHANGES', ''],
    ['Products That Started Getting Impressions', summary.newProducts],
    ['Products That Stopped Getting Impressions', summary.discontinuedProducts],
    ['Continuing Products', summary.continuingProducts],
    ['', ''],
    ['💰 NEWLY ACTIVE PRODUCTS PERFORMANCE', ''],
    ['Total Impressions', summary.newProductsImpressions],
    ['Total Clicks', summary.newProductsClicks],
    ['Total Conversions', summary.newProductsConversions],
    ['Total Conversion Value', summary.newProductsConversionValue],
    ['Total Cost', summary.newProductsCost],
    ['', ''],
    ['📉 INACTIVE PRODUCTS IMPACT', ''],
    ['Last Month Impressions', summary.discontinuedProductsImpressions],
    ['Last Month Clicks', summary.discontinuedProductsClicks],
    ['Last Month Conversions', summary.discontinuedProductsConversions],
    ['Last Month Conversion Value', summary.discontinuedProductsConversionValue],
    ['Last Month Cost', summary.discontinuedProductsCost]
  ];
  
  // Write summary data
  const summaryRange = summarySheet.getRange(1, 1, summaryData.length, 2);
  summaryRange.setValues(summaryData);
  
  // Format the summary sheet
  formatAnalysisSummary(summarySheet, summaryData.length, currencyFormat);
  
  Logger.log('✅ Analysis summary created');
}

/**
 * Formats the analysis summary sheet
 */
function formatAnalysisSummary(sheet, dataRows, currencyFormat) {
  // Clear any existing frozen columns to avoid merge conflicts
  sheet.setFrozenColumns(0);
  
  // Title formatting
  const titleRange = sheet.getRange('A1:B1');
  titleRange.setFontWeight('bold');
  titleRange.setFontSize(14);
  titleRange.setBackground('#4285f4');
  titleRange.setFontColor('#ffffff');
  titleRange.merge();
  
  // Section headers formatting
  const sectionRanges = ['A3', 'A8', 'A13', 'A20'];
  sectionRanges.forEach(range => {
    const sectionRange = sheet.getRange(range);
    sectionRange.setFontWeight('bold');
    sectionRange.setBackground('#e8f0fe');
    sectionRange.setFontColor('#1a73e8');
  });
  
  // Value formatting
  const valueColumns = ['B4:B6', 'B9:B11', 'B14:B16', 'B21:B23'];
  valueColumns.forEach(range => {
    const valueRange = sheet.getRange(range);
    valueRange.setHorizontalAlignment('right');
    valueRange.setNumberFormat('#,##0');
  });
  
  // Conversions formatting
  sheet.getRange('B16').setNumberFormat('#,##0.00'); // New products conversions
  sheet.getRange('B23').setNumberFormat('#,##0.00'); // Discontinued products conversions
  
  // Currency formatting for conversion value and cost
  sheet.getRange('B17').setNumberFormat(currencyFormat); // New products conversion value
  sheet.getRange('B18').setNumberFormat(currencyFormat); // New products cost
  sheet.getRange('B24').setNumberFormat(currencyFormat); // Discontinued products conversion value
  sheet.getRange('B25').setNumberFormat(currencyFormat); // Discontinued products cost
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, 2);
  
  // Set column widths
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 150);
}

/**
 * Creates new products report
 */
function createNewProductsReport(spreadsheet, newProducts) {
  let newSheet = spreadsheet.getSheetByName(OUTPUT_TABS.NEW_PRODUCTS);
  if (newSheet) {
    newSheet.clear();
  } else {
    newSheet = spreadsheet.insertSheet(OUTPUT_TABS.NEW_PRODUCTS);
  }
  
  // Create headers
  const headers = [
    'Product ID',
    'Product Title',
    'Impressions',
    'Clicks',
    'Product Type L1',
    'Product Type L2',
    'Product Type L3',
    'Conversions',
    'Conversion Value',
    'Cost',
    'Conv. Value / Cost',
    'Currency'
  ];
  
  // Create data rows
  const dataRows = [headers];
  for (const product of newProducts.values()) {
    dataRows.push([
      product.productId,
      product.productTitle,
      product.impressions,
      product.clicks,
      product.productTypeL1,
      product.productTypeL2,
      product.productTypeL3,
      product.conversions,
      product.conversionValue,
      product.cost,
      product.convValuePerCost,
      product.currency
    ]);
  }
  
  // Write data
  if (dataRows.length > 1) {
    const range = newSheet.getRange(1, 1, dataRows.length, headers.length);
    range.setValues(dataRows);
    
    // Format sheet
    formatReportSheet(newSheet, headers.length, dataRows.length, `Products That Started Getting Impressions Last Month (Automated) - ${ACCOUNT_CURRENCY}`);
  } else {
    newSheet.getRange('A1').setValue('No newly active products found');
  }
  
  Logger.log(`✅ Newly active products report created with ${newProducts.size} products`);
}

/**
 * Creates discontinued products report
 */
function createDiscontinuedProductsReport(spreadsheet, discontinuedProducts) {
  let discSheet = spreadsheet.getSheetByName(OUTPUT_TABS.DISCONTINUED);
  if (discSheet) {
    discSheet.clear();
  } else {
    discSheet = spreadsheet.insertSheet(OUTPUT_TABS.DISCONTINUED);
  }
  
  // Create headers
  const headers = [
    'Product ID',
    'Product Title',
    'Last Month Impressions',
    'Last Month Clicks',
    'Product Type L1',
    'Product Type L2',
    'Product Type L3',
    'Last Month Conversions',
    'Last Month Conversion Value',
    'Last Month Cost',
    'Last Month Conv. Value / Cost',
    'Currency'
  ];
  
  // Create data rows
  const dataRows = [headers];
  for (const product of discontinuedProducts.values()) {
    dataRows.push([
      product.productId,
      product.productTitle,
      product.impressions,
      product.clicks,
      product.productTypeL1,
      product.productTypeL2,
      product.productTypeL3,
      product.conversions,
      product.conversionValue,
      product.cost,
      product.convValuePerCost,
      product.currency
    ]);
  }
  
  // Write data
  if (dataRows.length > 1) {
    const range = discSheet.getRange(1, 1, dataRows.length, headers.length);
    range.setValues(dataRows);
    
    // Format sheet
    formatReportSheet(discSheet, headers.length, dataRows.length, `Products That Stopped Getting Impressions Since Two Months Ago (Automated) - ${ACCOUNT_CURRENCY}`);
  } else {
    discSheet.getRange('A1').setValue('No inactive products found');
  }
  
  Logger.log(`✅ Inactive products report created with ${discontinuedProducts.size} products`);
}

/**
 * Creates performance changes report for continuing products
 */
function createPerformanceChangesReport(spreadsheet, continuingProducts) {
  let perfSheet = spreadsheet.getSheetByName(OUTPUT_TABS.PERFORMANCE_CHANGES);
  if (perfSheet) {
    perfSheet.clear();
  } else {
    perfSheet = spreadsheet.insertSheet(OUTPUT_TABS.PERFORMANCE_CHANGES);
  }
  
  // Create headers
  const headers = [
    'Product ID',
    'Product Title',
    'Product Type L1',
    'Product Type L2',
    'Product Type L3',
    'Currency',
    'Impressions (Last Month)',
    'Impressions (Two Months Ago)',
    'Impressions Change %',
    'Clicks (Last Month)',
    'Clicks (Two Months Ago)', 
    'Clicks Change %',
    'Conversions (Last Month)',
    'Conversions (Two Months Ago)',
    'Conversions Change %',
    'Conv. Value (Last Month)',
    'Conv. Value (Two Months Ago)',
    'Conv. Value Change %',
    'Cost (Last Month)',
    'Cost (Two Months Ago)',
    'Cost Change %',
    'Conv. Value/Cost (Last Month)',
    'Conv. Value/Cost (Two Months Ago)',
    'Conv. Value/Cost Change %'
  ];
  
  // Create data rows
  const dataRows = [headers];
  for (const changes of continuingProducts.values()) {
    const lm = changes.lastMonth;
    const tma = changes.twoMonthsAgo;
    const ch = changes.changes;
    
    dataRows.push([
      changes.productId,
      lm.productTitle,
      lm.productTypeL1,
      lm.productTypeL2,
      lm.productTypeL3,
      lm.currency,
      lm.impressions,
      tma.impressions,
      ch.impressionsChange,
      lm.clicks,
      tma.clicks,
      ch.clicksChange,
      lm.conversions,
      tma.conversions,
      ch.conversionsChange,
      lm.conversionValue,
      tma.conversionValue,
      ch.conversionValueChange,
      lm.cost,
      tma.cost,
      ch.costChange,
      lm.convValuePerCost,
      tma.convValuePerCost,
      ch.convValuePerCostChange
    ]);
  }
  
  // Write data
  if (dataRows.length > 1) {
    const range = perfSheet.getRange(1, 1, dataRows.length, headers.length);
    range.setValues(dataRows);
    
    // Format sheet
    formatPerformanceChangesSheet(perfSheet, headers.length, dataRows.length);
  } else {
    perfSheet.getRange('A1').setValue('No continuing products found');
  }
  
  Logger.log(`✅ Performance changes report created with ${continuingProducts.size} products`);
}

/**
 * Formats standard report sheets
 */
function formatReportSheet(sheet, columnCount, rowCount, title) {
  // Clear any existing frozen columns to avoid merge conflicts
  sheet.setFrozenColumns(0);
  
  // Add title above the data
  sheet.insertRowBefore(1);
  const titleRange = sheet.getRange(1, 1, 1, columnCount);
  titleRange.setValue(title);
  titleRange.setFontWeight('bold');
  titleRange.setFontSize(12);
  titleRange.setBackground('#34a853');
  titleRange.setFontColor('#ffffff');
  titleRange.merge();
  titleRange.setHorizontalAlignment('center');
  
  // Format headers
  const headerRange = sheet.getRange(2, 1, 1, columnCount);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#e8f0fe');
  headerRange.setFontColor('#1a73e8');
  
  // Format data columns
  if (rowCount > 2) {
    const currencyFormat = getCurrencyFormat(ACCOUNT_CURRENCY);
    
    // Number formatting for metrics
    sheet.getRange(3, 3, rowCount - 2, 2).setNumberFormat('#,##0'); // Impressions, Clicks
    sheet.getRange(3, 8, rowCount - 2, 1).setNumberFormat('#,##0.00'); // Conversions (2 decimals)
    sheet.getRange(3, 9, rowCount - 2, 1).setNumberFormat(currencyFormat); // Conversion Value (currency)
    sheet.getRange(3, 10, rowCount - 2, 1).setNumberFormat(currencyFormat); // Cost (currency)
    sheet.getRange(3, 11, rowCount - 2, 1).setNumberFormat('#,##0.00'); // Conv Value / Cost (2 decimals)
  }
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, columnCount);
  
  // Freeze header rows
  sheet.setFrozenRows(2);
}

/**
 * Formats performance changes sheet with conditional formatting
 */
function formatPerformanceChangesSheet(sheet, columnCount, rowCount) {
  // Clear any existing frozen columns to avoid merge conflicts
  sheet.setFrozenColumns(0);
  
  // Add title
  sheet.insertRowBefore(1);
  const titleRange = sheet.getRange(1, 1, 1, columnCount);
  titleRange.setValue(`Performance Changes for Continuing Products (Automated) - ${ACCOUNT_CURRENCY}`);
  titleRange.setFontWeight('bold');
  titleRange.setFontSize(12);
  titleRange.setBackground('#ff9800');
  titleRange.setFontColor('#ffffff');
  titleRange.merge();
  titleRange.setHorizontalAlignment('center');
  
  // Format headers
  const headerRange = sheet.getRange(2, 1, 1, columnCount);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#fff3e0');
  headerRange.setFontColor('#e65100');
  
  if (rowCount > 2) {
    const currencyFormat = getCurrencyFormat(ACCOUNT_CURRENCY);
    
    // Number formatting
    sheet.getRange(3, 7, rowCount - 2, 2).setNumberFormat('#,##0'); // Impressions
    sheet.getRange(3, 9, rowCount - 2, 1).setNumberFormat('#,##0"%"'); // Impressions Change %
    sheet.getRange(3, 10, rowCount - 2, 2).setNumberFormat('#,##0'); // Clicks
    sheet.getRange(3, 12, rowCount - 2, 1).setNumberFormat('#,##0"%"'); // Clicks Change %
    sheet.getRange(3, 13, rowCount - 2, 2).setNumberFormat('#,##0.00'); // Conversions (2 decimals)
    sheet.getRange(3, 15, rowCount - 2, 1).setNumberFormat('#,##0"%"'); // Conversions Change %
    sheet.getRange(3, 16, rowCount - 2, 2).setNumberFormat(currencyFormat); // Conversion Value (currency)
    sheet.getRange(3, 18, rowCount - 2, 1).setNumberFormat('#,##0"%"'); // Conversion Value Change %
    sheet.getRange(3, 19, rowCount - 2, 2).setNumberFormat(currencyFormat); // Cost (currency)
    sheet.getRange(3, 21, rowCount - 2, 1).setNumberFormat('#,##0"%"'); // Cost Change %
    sheet.getRange(3, 22, rowCount - 2, 2).setNumberFormat('#,##0.00'); // Conv Value/Cost (2 decimals)
    sheet.getRange(3, 24, rowCount - 2, 1).setNumberFormat('#,##0"%"'); // Conv Value/Cost Change %
    
    // Conditional formatting for percentage changes
    const changeColumns = [9, 12, 15, 18, 21, 24]; // Change % columns
    changeColumns.forEach(col => {
      const changeRange = sheet.getRange(3, col, rowCount - 2, 1);
      
      // Positive changes (green)
      const positiveRule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(0)
        .setBackground('#c8e6c9')
        .setRanges([changeRange])
        .build();
      
      // Negative changes (red)
      const negativeRule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThan(0)
        .setBackground('#ffcdd2')
        .setRanges([changeRange])
        .build();
      
      const rules = sheet.getConditionalFormatRules();
      rules.push(positiveRule, negativeRule);
      sheet.setConditionalFormatRules(rules);
    });
  }
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, columnCount);
  
  // Freeze header rows
  sheet.setFrozenRows(2);
}

// AI Insights functions (similar to v1.3 but with currency awareness)
/**
 * Generates AI-powered business insights in plain language
 */
function generateExecutiveInsights(spreadsheet, analysis) {
  try {
    // Validate API key
    if (!OPENAI_API_KEY || OPENAI_API_KEY.trim() === '') {
      Logger.log('⚠️ WARNING: OpenAI API key not provided - skipping business insights');
      return;
    }
    
    Logger.log('🔍 Preparing data for AI analysis...');
    
    // Prepare structured data for AI analysis
    const analysisData = prepareAIAnalysisData(analysis);
    
    // Generate insights using OpenAI
    const insights = callOpenAIForInsights(analysisData);
    
    if (insights) {
      // Create the executive insights report
      createExecutiveInsightsReport(spreadsheet, insights);
      Logger.log('✅ Business insights generated successfully');
    } else {
      Logger.log('❌ Failed to generate business insights');
    }
    
  } catch (error) {
    Logger.log('❌ Error generating business insights: ' + error.toString());
  }
}

/**
 * Prepares structured data for AI analysis
 */
function prepareAIAnalysisData(analysis) {
  const summary = analysis.summary;
  
  // Calculate key ratios and trends
  const inactivityRate = summary.totalProductsTwoMonthsAgo > 0 ? 
    (summary.discontinuedProducts / summary.totalProductsTwoMonthsAgo * 100).toFixed(1) : 0;
  const replacementRate = summary.discontinuedProducts > 0 ? 
    (summary.newProducts / summary.discontinuedProducts * 100).toFixed(1) : 0;
  const netChange = summary.totalProductsLastMonth - summary.totalProductsTwoMonthsAgo;
  const netChangePercent = summary.totalProductsTwoMonthsAgo > 0 ? 
    (netChange / summary.totalProductsTwoMonthsAgo * 100).toFixed(1) : 0;
  
  // Calculate performance metrics
  const lostRevenueImpact = summary.discontinuedProductsConversionValue;
  const newRevenueGain = summary.newProductsConversionValue;
  const netRevenueImpact = newRevenueGain - lostRevenueImpact;
  
  // Get top 10 inactive products by revenue
  const topDiscontinuedProducts = Array.from(analysis.discontinuedProducts.values())
    .sort((a, b) => (b.conversionValue || 0) - (a.conversionValue || 0))
    .slice(0, 10)
    .map(product => ({
      title: product.productTitle,
      revenue: product.conversionValue || 0,
      itemId: product.productId
    }));
  
  // Get top 10 newly active products by revenue  
  const topNewProducts = Array.from(analysis.newProducts.values())
    .sort((a, b) => (b.conversionValue || 0) - (a.conversionValue || 0))
    .slice(0, 10)
    .map(product => ({
      title: product.productTitle,
      revenue: product.conversionValue || 0,
      itemId: product.productId
    }));
  
  return {
    currency: ACCOUNT_CURRENCY,
    impressionActivityMetrics: {
      totalProductsLastMonth: summary.totalProductsLastMonth,
      totalProductsTwoMonthsAgo: summary.totalProductsTwoMonthsAgo,
      netChange: netChange,
      netChangePercent: netChangePercent,
      inactivityRate: inactivityRate,
      replacementRate: replacementRate
    },
    impressionActivityChanges: {
      newProducts: summary.newProducts,
      discontinuedProducts: summary.discontinuedProducts,
      continuingProducts: summary.continuingProducts
    },
    performanceMetrics: {
      newProductsImpressions: summary.newProductsImpressions,
      newProductsClicks: summary.newProductsClicks,
      newProductsConversions: summary.newProductsConversions,
      newProductsConversionValue: summary.newProductsConversionValue,
      newProductsCost: summary.newProductsCost,
      lastMonthImpressions: summary.discontinuedProductsImpressions,
      lastMonthClicks: summary.discontinuedProductsClicks,
      lastMonthConversions: summary.discontinuedProductsConversions,
      lastMonthConversionValue: summary.discontinuedProductsConversionValue,
      lastMonthCost: summary.discontinuedProductsCost,
      netRevenueImpact: netRevenueImpact
    },
    topInactiveProducts: topDiscontinuedProducts,
    topNewlyActiveProducts: topNewProducts
  };
}

/**
 * Calls OpenAI API for executive insights
 */
function callOpenAIForInsights(analysisData) {
  const prompt = createInsightsPrompt();
  const dataString = JSON.stringify(analysisData, null, 2);
  
  const payload = {
    model: "gpt-4o-mini",
    messages: [
      {
        role: "system",
        content: prompt
      },
      {
        role: "user", 
        content: `Analyze this automatically-collected product inventory data and provide executive insights. 

IMPORTANT: This data was automatically pulled from Google Ads Shopping Product Stats API. The data includes "topDiscontinuedProducts" and "topNewProducts" arrays with real product information. Use the actual itemId, title, and revenue values from these arrays. All monetary values are in ${ACCOUNT_CURRENCY}.

Data:
${dataString}`
      }
    ],
    max_tokens: 1500,
    temperature: 0.3
  };
  
  const options = {
    method: "POST",
    headers: {
      "Authorization": `Bearer ${OPENAI_API_KEY}`,
      "Content-Type": "application/json"
    },
    payload: JSON.stringify(payload)
  };
  
  try {
    const response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", options);
    const responseData = JSON.parse(response.getContentText());
    
    if (responseData.error) {
      Logger.log(`❌ OpenAI API Error: ${responseData.error.message}`);
      return null;
    }
    
    return responseData.choices[0].message.content;
    
  } catch (error) {
    Logger.log(`❌ OpenAI API call failed: ${error.message}`);
    return null;
  }
}

/**
 * Creates the prompt for AI insights generation
 */
function createInsightsPrompt() {
  return `You are helping a small-to-medium business owner understand their Google Ads impression activity changes and what it means for their campaign performance and sales.

This data was automatically collected from Google Ads Shopping Product Stats for complete calendar months, comparing last month vs two months ago. The data represents products that were actively getting impressions and their performance metrics.

Write a simple, clear report about their impression activity data. Avoid technical jargon and explain things in plain business terms.

Format your response with these exact sections:

## 📊 WHAT'S HAPPENING
(2-3 sentences explaining what changed with their product impression activity last month)

## 🔍 WHAT IT MEANS

### Your Impression Activity Mix
(Simple explanation of how many products started getting impressions last month, stopped getting impressions since two months ago, or continued getting impressions. Use "products that stopped getting impressions" NOT "discontinued products")

### Sales Impact  
(How these impression activity changes affected revenue and website traffic in plain terms - compare last month vs two months ago)

### Business Impact
(What this means for the business - opportunities or problems to watch. Use "products that stopped getting impressions" NOT "discontinued products". Note: Products stopping impressions could be due to bidding changes, budget reallocation, competition, seasonal factors, performance issues, or actual discontinuation)

## 💰 TOP REVENUE CHANGES

### Biggest Revenue Losses (Products No Longer Getting Impressions)
Look for the "topInactiveProducts" array in the provided data. For each product object in this array, extract the itemId, title, and revenue fields and create a numbered list exactly like this:
1. [itemId from data] | [title from data] | [currency][revenue from data]

Note: These products stopped getting impressions, which could be due to bidding changes, budget reallocation, competition, seasonal factors, poor performance, or actual discontinuation.

### Best New Revenue Generators (Products That Started Getting Impressions)
Look for the "topNewlyActiveProducts" array in the provided data. For each product object in this array, extract the itemId, title, and revenue fields and create a numbered list exactly like this:
1. [itemId from data] | [title from data] | [currency][revenue from data]

## 🎯 WHAT TO DO NEXT
(3-4 simple action steps, starting with the most important)

## ⚠️ THINGS TO WATCH
(Important issues that need attention soon. Consider investigating why specific products stopped getting impressions - check bidding, budget allocation, competition, and seasonal factors)

Guidelines:
- Write like you're talking to a business owner, not a marketing expert
- Keep it under 800 words total
- Use real numbers from their data with the correct currency
- Focus on what they can actually do about it
- Explain how it affects their bottom line
- Avoid Google Ads technical terms
- Give practical, doable advice
- Note that this data represents products that were getting impressions for complete calendar months only, not entire inventory
- Clarify that "products stopping impressions" doesn't necessarily mean they were discontinued - they could be due to campaign management decisions, competition, or performance factors
- Use "products that stopped getting impressions" or "products no longer receiving impressions" instead of "discontinued products"
- Use actual product data from the arrays provided with proper currency formatting
- Emphasize that these are impression activity changes, not inventory changes
- When referring to time periods, use "last month" and "two months ago" consistently
- Explain that products stopping impressions could be due to: bidding changes, budget reallocation, increased competition, seasonal demand shifts, poor performance, or actual discontinuation
- Be specific about the difference between impression activity and inventory status
- NEVER use the term "discontinued products" - instead use "products that stopped getting impressions" or "products no longer receiving impressions"
- Avoid implying that products stopping impressions means they were removed from inventory`;
}

/**
 * Creates the executive insights report tab
 */
function createExecutiveInsightsReport(spreadsheet, insights) {
  let insightsSheet = spreadsheet.getSheetByName(OUTPUT_TABS.EXECUTIVE_INSIGHTS);
  if (insightsSheet) {
    insightsSheet.clear();
  } else {
    insightsSheet = spreadsheet.insertSheet(OUTPUT_TABS.EXECUTIVE_INSIGHTS);
  }
  
  // Split insights into lines for better formatting
  const insightLines = insights.split('\n');
  const formattedData = [];
  
  // Process each line for formatting
  for (let i = 0; i < insightLines.length; i++) {
    const line = insightLines[i];
    if (line.trim() === '') {
      formattedData.push(['']); // Empty line
    } else {
      formattedData.push([line]);
    }
  }
  
  // Write the insights to the sheet
  if (formattedData.length > 0) {
    const range = insightsSheet.getRange(1, 1, formattedData.length, 1);
    range.setValues(formattedData);
    
    // Format the insights sheet
    formatInsightsSheet(insightsSheet, formattedData.length);
  }
  
  Logger.log('✅ Business insights report created');
}

/**
 * Formats the executive insights sheet
 */
function formatInsightsSheet(sheet, rowCount) {
  // Set column width for readability
  sheet.setColumnWidth(1, 800);
  
  // Format different types of content
  for (let i = 1; i <= rowCount; i++) {
    const cellValue = sheet.getRange(i, 1).getValue();
    
    if (typeof cellValue === 'string') {
      const cellRange = sheet.getRange(i, 1);
      
      // Main headers (##)
      if (cellValue.startsWith('## ')) {
        cellRange.setFontWeight('bold');
        cellRange.setFontSize(14);
        cellRange.setBackground('#4285f4');
        cellRange.setFontColor('#ffffff');
      }
      // Sub-headers (###)
      else if (cellValue.startsWith('### ')) {
        cellRange.setFontWeight('bold');
        cellRange.setFontSize(12);
        cellRange.setBackground('#e8f0fe');
        cellRange.setFontColor('#1a73e8');
      }
      // Regular content
      else if (cellValue.trim() !== '') {
        cellRange.setFontSize(11);
        cellRange.setVerticalAlignment('top');
        cellRange.setWrap(true);
      }
    }
  }
  
  // Auto-resize row heights for wrapped text
  sheet.autoResizeRows(1, rowCount);
}
