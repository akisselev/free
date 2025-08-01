/**
 * ====================================
 * MONTHLY ADS MONITOR v1.1
 * ====================================
 * 
 * A Google Ads script that generates monthly performance trends
 * with year-over-year comparison and dynamic filtering capabilities.
 * 
 * Features:
 * - Monthly campaign performance data (2-year comparison)
 * - Campaign type filtering
 * - Brand campaign exclusion
 * - Professional formatting with conditional coloring
 * - Dynamic formulas for real-time filtering
 * - Conversion value by conversion time tracking
 * - ROAS by conversion time analysis
 * 
 * Installation Steps:
 * 1. Go to https://ads.google.com and sign into your Google Ads account
 * 2. Navigate to "Tools & Settings" > "Bulk Actions" > "Scripts"
 * 3. Click the "+" button to create a new script
 * 4. Copy and paste this entire script code
 * 5. Update the SHEET_URL variable (line 22) with your Google Sheet URL (optional - leave empty to auto-create)
 * 6. Update the BRAND_CAMPAIGNS array (line 28) with your brand campaign names
 * 7. Save the script with a descriptive name like "Monthly Ads Monitor v1.1"
 * 8. Click "Preview" to test the script
 * 9. Click "Run" to execute the script manually
 * 10. Create a schedule to run the script monthly on the first day of each month
 * 
 * Author: Andrey Kisselev
 * Version: 1.1
 * ====================================
 */

// Enter your Google Sheet URL here between the single quotes.
const SHEET_URL = '';

// *** ADD YOUR BRAND CAMPAIGN NAMES HERE ***
// Add the exact names of your brand campaigns below (case-sensitive)
// Example: 'Your Company Brand', 'Brand Search Campaign', 'Trademark Campaign'
// Leave empty [] if you don't have brand campaigns to exclude
const BRAND_CAMPAIGNS = [
  
];

// Calculate date range from beginning of last year to last complete month (2 years of data)
function getDateRange() {
  const today = new Date();
  
  // Calculate last complete month (end of previous month)
  const endDate = new Date(today.getFullYear(), today.getMonth(), 0); // Last day of previous month
  
  // Start from beginning of last year to get year-over-year comparison
  const currentYear = today.getFullYear();
  const startDate = new Date(currentYear - 1, 0, 1); // January 1st of last year
  
  const formatDate = (date) => {
    return Utilities.formatDate(date, AdsApp.currentAccount().getTimeZone(), 'yyyy-MM-dd');
  };
  
  return {
    start: formatDate(startDate),
    end: formatDate(endDate)
  };
}

const dateRange = getDateRange();
const START_DATE = dateRange.start;
const END_DATE = dateRange.end;

// Monthly campaign performance data with campaign type
const GAQL_QUERY = `
SELECT
  campaign.name,
  campaign.advertising_channel_type,
  segments.month,
  metrics.impressions,
  metrics.clicks,
  metrics.cost_micros,
  metrics.conversions,
  metrics.conversions_value,
  metrics.conversions_value_by_conversion_date
FROM campaign
WHERE metrics.impressions > 0
  AND segments.date BETWEEN "${START_DATE}" AND "${END_DATE}"
ORDER BY segments.month DESC, metrics.cost_micros DESC
`;

function main() {
  try {
    Logger.log('🚀 STARTING MONTHLY ADS MONITOR v1.1');
    Logger.log(`📅 Date Range (2-year comparison up to last complete month): ${START_DATE} to ${END_DATE}`);
    
    if (BRAND_CAMPAIGNS.length > 0) {
      Logger.log(`🏷️ Monthly Ads Monitor - Brand campaigns to identify: ${BRAND_CAMPAIGNS.join(', ')}`);
    } else {
      Logger.log('⚠️ Monthly Ads Monitor - No brand campaigns specified in BRAND_CAMPAIGNS array');
    }
    
    let spreadsheet;
    
    if (SHEET_URL === '') {
      // Create a new spreadsheet if no URL is provided
      spreadsheet = SpreadsheetApp.create('Monthly Ads Monitor - ' + 
        Utilities.formatDate(new Date(), AdsApp.currentAccount().getTimeZone(), 'yyyy-MM-dd'));
      Logger.log('✅ NEW SPREADSHEET CREATED: ' + spreadsheet.getUrl());
    } else {
      // Open existing spreadsheet
      spreadsheet = SpreadsheetApp.openByUrl(SHEET_URL);
      Logger.log('✅ Opened existing Monthly Ads Monitor spreadsheet');
    }
    
    // Get or create the "Raw" tab (renamed from Campaign Trends)
    let rawSheet = spreadsheet.getSheetByName('Raw');
    if (!rawSheet) {
      rawSheet = spreadsheet.insertSheet('Raw');
      Logger.log('✅ Created new "Raw" tab');
    } else {
      Logger.log('✅ Using existing "Raw" tab');
    }
    
    // Clear existing data
    rawSheet.clear();
    
    // Execute the GAQL query
    Logger.log('🔍 Executing query...');
    const rows = AdsApp.search(GAQL_QUERY);
    
    // Process data
    const rawData = processData(rows);
    
    if (rawData.length === 0) {
      Logger.log('⚠️ No data found for the specified date range.');
      return;
    }
    
    Logger.log(`📊 Found ${rawData.length} monthly campaign records`);
    
    // Create headers for raw data - 14 columns with campaign performance metrics including brand flag
    const rawHeaders = [
      'Month Start',
      'Month End',
      'Campaign Name',
      'Campaign Type',
      'Is Brand Campaign',
      'Impressions',
      'Clicks',
      'Cost',
      'Cost Per Click',
      'Conversions',
      'Cost Per Conversion',
      'Conversion Value',
      'Conversion Value (Time)',
      'ROAS (Time)'
    ];
    
    // Combine headers and data for raw sheet
    const allRawData = [rawHeaders, ...rawData];
    
    // Write to raw sheet
    Logger.log('📤 Writing raw data to spreadsheet...');
    const rawRange = rawSheet.getRange(1, 1, allRawData.length, rawHeaders.length);
    rawRange.setValues(allRawData);
    
    // Format raw sheet headers
    const rawHeaderRange = rawSheet.getRange(1, 1, 1, rawHeaders.length);
    rawHeaderRange.setFontWeight('bold');
    rawHeaderRange.setBackground('#4285f4');
    rawHeaderRange.setFontColor('#ffffff');
    
    // Auto-resize columns for raw sheet
    for (let i = 1; i <= rawHeaders.length; i++) {
      rawSheet.autoResizeColumn(i);
    }
    
    // Create aggregated monthly summary tab
    Logger.log('📊 Creating Monthly Ads Monitor dashboard...');
    createMonthlySummary(spreadsheet, rawData);
    
    Logger.log('🎉 SUCCESS! Monthly Ads Monitor v1.1 completed');
    Logger.log(`✅ Generated ${rawData.length} monthly campaign records with performance metrics`);
    Logger.log('🔗 Monthly Ads Monitor URL: ' + spreadsheet.getUrl());
    
  } catch (error) {
    Logger.log('❌ MONTHLY ADS MONITOR ERROR: ' + error.toString());
  }
}

function processData(rows) {
  let data = [];
  let count = 0;
  let brandCampaignCount = 0;
  
  Logger.log('⚙️ Processing monthly campaign performance data...');
  
  while (rows.hasNext()) {
    try {
      let row = rows.next();
      count++;
      
      const campaign = row.campaign || {};
      const segments = row.segments || {};
      const metrics = row.metrics || {};
      
      // Get campaign details
      let campaignName = campaign.name || '';
      let campaignType = campaign.advertisingChannelType || '';
      
      // Check if this campaign name is in the brand campaigns list
      let isBrand = BRAND_CAMPAIGNS.includes(campaignName);
      if (isBrand) brandCampaignCount++;
      
      // Format campaign type for readability
      switch(campaignType) {
        case 'SEARCH':
          campaignType = 'Search';
          break;
        case 'DISPLAY':
          campaignType = 'Display';
          break;
        case 'SHOPPING':
          campaignType = 'Shopping';
          break;
        case 'VIDEO':
          campaignType = 'Video';
          break;
        case 'PERFORMANCE_MAX':
          campaignType = 'Performance Max';
          break;
        default:
          campaignType = campaignType || 'Unknown';
      }
      
      let monthStart = segments.month || '';
      let impressions = Number(metrics.impressions) || 0;
      let clicks = Number(metrics.clicks) || 0;
      let costMicros = Number(metrics.costMicros) || 0;
      let conversions = Number(metrics.conversions) || 0;
      let conversionValue = Number(metrics.conversionsValue) || 0;
      let conversionValueByTime = Number(metrics.conversionsValueByConversionDate) || 0;
      
      // Calculate month end date (last day of the month)
      let monthEnd = '';
      if (monthStart) {
        try {
          let startDate = new Date(monthStart);
          let endDate = new Date(startDate.getFullYear(), startDate.getMonth() + 1, 0); // Last day of month
          monthEnd = Utilities.formatDate(endDate, AdsApp.currentAccount().getTimeZone(), 'yyyy-MM-dd');
        } catch (e) {
          monthEnd = monthStart; // fallback to start date if calculation fails
        }
      }
      
      // Convert cost from micros to currency and calculate derived metrics
      let cost = Math.round((costMicros / 1000000) * 100) / 100;
      conversionValue = Math.round(conversionValue * 100) / 100;
      conversionValueByTime = Math.round(conversionValueByTime * 100) / 100;
      
      // Calculate cost per click
      let costPerClick = clicks > 0 ? Math.round((cost / clicks) * 100) / 100 : 0;
      
      // Calculate cost per conversion
      let costPerConversion = conversions > 0 ? Math.round((cost / conversions) * 100) / 100 : 0;
      
      // Calculate ROAS (time)
      let roasTime = cost > 0 ? Math.round((conversionValueByTime / cost) * 100) / 100 : 0;
      
      // Create row - 14 columns with campaign performance metrics including brand flag
      let newRow = [
        monthStart,          // Month Start
        monthEnd,            // Month End
        campaignName,        // Campaign Name
        campaignType,        // Campaign Type
        isBrand,             // Is Brand Campaign
        impressions,         // Impressions
        clicks,              // Clicks
        cost,                // Cost
        costPerClick,        // Cost Per Click
        conversions,         // Conversions
        costPerConversion,   // Cost Per Conversion
        conversionValue,     // Conversion Value
        conversionValueByTime, // Conversion Value (Time)
        roasTime             // ROAS (Time)
      ];
      
      data.push(newRow);
      
      if (count % 500 === 0) {
        Logger.log(`📈 Processed ${count} monthly campaign records...`);
      }
      
    } catch (e) {
      Logger.log(`❌ Error processing row: ${e}`);
    }
  }
  
  Logger.log(`✅ Processing complete: ${data.length} records (${brandCampaignCount} brand campaign records found)`);
  return data;
}

function createMonthlySummary(spreadsheet, rawData) {
  try {
    // Get or create the "Monthly Summary" tab
    let summarySheet = spreadsheet.getSheetByName('Monthly Summary');
    if (!summarySheet) {
      summarySheet = spreadsheet.insertSheet('Monthly Summary');
      Logger.log('✅ Created new "Monthly Summary" tab');
    } else {
      Logger.log('✅ Using existing "Monthly Summary" tab');
    }
    
    summarySheet.clear();
    
    // Unfreeze any frozen rows/columns to avoid merge conflicts
    summarySheet.setFrozenRows(0);
    summarySheet.setFrozenColumns(0);
    
    // Get unique campaign types for the filter dropdown
    const allCampaignTypes = [...new Set(rawData.map(row => row[3]))].sort(); // Campaign Types
    const campaignTypeOptions = ['All', ...allCampaignTypes];
    
    // Debug: Log campaign types found in data
    Logger.log(`🔍 Campaign types found in raw data: ${allCampaignTypes.join(', ')}`);
    
    // Debug: Check data distribution by year and campaign type
    const debugCurrentYear = new Date().getFullYear();
    const debugPreviousYear = debugCurrentYear - 1;
    const currentYearData = rawData.filter(row => row[0].includes(debugCurrentYear.toString()));
    const previousYearData = rawData.filter(row => row[0].includes(debugPreviousYear.toString()));
    
    Logger.log(`📊 ${debugCurrentYear} data: ${currentYearData.length} records`);
    Logger.log(`📊 ${debugPreviousYear} data: ${previousYearData.length} records`);
    
    // Debug: Check brand campaign distribution
    const brandCampaigns = rawData.filter(row => row[4] === true); // Is Brand Campaign column
    const nonBrandCampaigns = rawData.filter(row => row[4] === false);
    
    Logger.log(`🏷️ Brand campaigns: ${brandCampaigns.length} records`);
    Logger.log(`🔍 Non-brand campaigns: ${nonBrandCampaigns.length} records`);
    
    // Add app title in first row
    summarySheet.getRange('A1').setValue('Monthly Ads Monitor v1.1');
    summarySheet.getRange('A1').setFontWeight('bold');
    summarySheet.getRange('A1').setFontSize(14);
    summarySheet.getRange('A1').setBackground('#f0f0f0');
    
    // Merge cells A1 and B1 for the app title
    summarySheet.getRange('A1:B1').merge();
    
    // Add campaign type filter controls
    summarySheet.getRange('A2').setValue('Campaign Type Filter:');
    summarySheet.getRange('A2').setFontWeight('bold');
    summarySheet.getRange('A2').setBackground('#f0f0f0');
    
    // Set up dropdown for campaign type selection
    const filterCell = summarySheet.getRange('B2');
    
    // Read the current filter selection BEFORE setting up dropdown
    const currentSelection = filterCell.getValue();
    
    // Set up dropdown validation
    const filterRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(campaignTypeOptions)
      .setAllowInvalid(false)
      .build();
    filterCell.setDataValidation(filterRule);
    filterCell.setBackground('#ffffff');
    filterCell.setBorder(true, true, true, true, true, true);
    
    // Only set default if cell is empty or invalid
    if (!currentSelection || !campaignTypeOptions.includes(currentSelection)) {
      filterCell.setValue('All');
    }
    
    // Add brand campaign filter controls
    summarySheet.getRange('A3').setValue('Exclude Brand Campaigns?');
    summarySheet.getRange('A3').setFontWeight('bold');
    summarySheet.getRange('A3').setBackground('#f0f0f0');
    
    // Add client name in row 4
    const clientName = AdsApp.currentAccount().getName();
    summarySheet.getRange('A4').setValue('Acc.: ' + clientName);
    summarySheet.getRange('A4').setFontWeight('bold');
    summarySheet.getRange('A4').setFontSize(12);
    summarySheet.getRange('A4').setBackground('#e8f0fe');
    summarySheet.getRange('A4').setFontColor('#1a73e8');
    
    // Merge cells A4 and B4 for the account name
    summarySheet.getRange('A4:B4').merge();
    
    // Set up dropdown for brand campaign exclusion
    const brandFilterCell = summarySheet.getRange('B3');
    
    // Read the current brand filter selection BEFORE setting up dropdown
    const currentBrandSelection = brandFilterCell.getValue();
    
    // Set up dropdown validation for brand filter
    const brandFilterRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['No', 'Yes'])
      .setAllowInvalid(false)
      .build();
    brandFilterCell.setDataValidation(brandFilterRule);
    brandFilterCell.setBackground('#ffffff');
    brandFilterCell.setBorder(true, true, true, true, true, true);
    
    // Only set default if cell is empty or invalid
    if (!currentBrandSelection || !['No', 'Yes'].includes(currentBrandSelection)) {
      brandFilterCell.setValue('No');
    }
    
    // Add instructions
    summarySheet.getRange('D2').setValue('🔄 Dynamic Filters: Select options from dropdowns → Data updates automatically!');
    summarySheet.getRange('D2').setFontStyle('italic');
    summarySheet.getRange('D2').setFontSize(9);
    summarySheet.getRange('D2').setFontColor('#0066cc');
    
    summarySheet.getRange('D3').setValue('💡 Brand Filter: "Yes" excludes brand campaigns from analysis');
    summarySheet.getRange('D3').setFontStyle('italic');
    summarySheet.getRange('D3').setFontSize(9);
    summarySheet.getRange('D3').setFontColor('#666666');
    
    // Read the final selected filter values
    const selectedCampaignType = filterCell.getValue() || 'All';
    const excludeBrandCampaigns = brandFilterCell.getValue() === 'Yes';
    Logger.log(`📊 Dynamic filtering enabled - Campaign type: ${selectedCampaignType}, Exclude brand: ${excludeBrandCampaigns}`);
    
    // Get current and previous year
    const currentYear = new Date().getFullYear();
    const previousYear = currentYear - 1;
    
    // Get all unique months from raw data and separate by year
    const allMonthDates = [...new Set(rawData.map(row => row[0]))].sort();
    const currentYearMonths = allMonthDates.filter(month => month.includes(currentYear.toString()));
    const allPreviousYearMonths = allMonthDates.filter(month => month.includes(previousYear.toString()));
    
    // Calculate current month number based on last complete month to limit previous year data
    const today = new Date();
    const lastCompleteMonth = new Date(today.getFullYear(), today.getMonth(), 0); // Last day of previous month
    const currentMonthNumber = lastCompleteMonth.getMonth() + 1; // 1-based month number
    
    // Filter previous year months to only include months up to current month number
    const previousYearMonths = allPreviousYearMonths.slice(0, currentMonthNumber);
    
    // Limit current year months to current month number as well
    const limitedCurrentYearMonths = currentYearMonths.slice(0, currentMonthNumber);
    
    Logger.log(`📅 Current month number: ${currentMonthNumber}`);
    Logger.log(`📅 Found ${limitedCurrentYearMonths.length} months in ${currentYear} and ${previousYearMonths.length} months in ${previousYear} (limited to month ${currentMonthNumber})`);
    
    // Create three-row header structure (rows 6-7)
    // Main headers (row 6) - now with 3 columns per metric
    const mainHeaders = ['Month Number', 'Month Start', 'Clicks', '', '', 'Average CPC', '', '', 'Cost', '', '', 'Conversions', '', '', 'Conversion Value', '', '', 'Conversion Value (Time)', '', '', 'ROAS', '', '', 'ROAS (Time)', '', ''];
    summarySheet.getRange('A6:Z6').setValues([mainHeaders]);
    
    // Year subheaders (row 7) - including Index YoY columns
    const yearHeaders = ['', '', currentYear, previousYear, 'Index YoY', currentYear, previousYear, 'Index YoY', currentYear, previousYear, 'Index YoY', currentYear, previousYear, 'Index YoY', currentYear, previousYear, 'Index YoY', currentYear, previousYear, 'Index YoY', currentYear, previousYear, 'Index YoY', currentYear, previousYear, 'Index YoY'];
    const yearHeaderRange = summarySheet.getRange('A7:Z7');
    yearHeaderRange.setValues([yearHeaders]);
    
    // Format year headers as plain numbers (not currency) for year columns only
    summarySheet.getRange('C7:D7').setNumberFormat('0'); // Clicks years
    summarySheet.getRange('F7:G7').setNumberFormat('0'); // CPC years
    summarySheet.getRange('I7:J7').setNumberFormat('0'); // Cost years
    summarySheet.getRange('L7:M7').setNumberFormat('0'); // Conversions years
    summarySheet.getRange('O7:P7').setNumberFormat('0'); // Conversion Value years
    summarySheet.getRange('R7:S7').setNumberFormat('0'); // Conversion Value (Time) years
    summarySheet.getRange('U7:V7').setNumberFormat('0'); // ROAS years
    summarySheet.getRange('X7:Y7').setNumberFormat('0'); // ROAS (Time) years
    
    // Merge cells for main headers
    try {
      summarySheet.getRange('A6:A7').merge(); // Month Number
      summarySheet.getRange('B6:B7').merge(); // Month Start
      summarySheet.getRange('C6:E6').merge(); // Clicks
      summarySheet.getRange('F6:H6').merge(); // Average CPC
      summarySheet.getRange('I6:K6').merge(); // Cost
      summarySheet.getRange('L6:N6').merge(); // Conversions
      summarySheet.getRange('O6:Q6').merge(); // Conversion Value
      summarySheet.getRange('R6:T6').merge(); // Conversion Value (Time)
      summarySheet.getRange('U6:W6').merge(); // ROAS
      summarySheet.getRange('X6:Z6').merge(); // ROAS (Time)
      Logger.log('✅ Header cells merged successfully');
    } catch (mergeError) {
      Logger.log('⚠️ Warning: Could not merge header cells: ' + mergeError.toString());
      Logger.log('📋 Headers will still function correctly without merging');
    }
    
    // Create dynamic formulas that reference the filter and raw data
    Logger.log('📤 Creating year-over-year dynamic formulas...');
    
    // Build formula-based data rows that update automatically
    const formulaData = [];
    
    // Use the actual number of complete months we have data for (limited to current complete months)
    const maxMonths = limitedCurrentYearMonths.length;
    
    for (let i = 0; i < maxMonths; i++) {
      const monthNumber = i + 1;
      const currentMonth = limitedCurrentYearMonths[i] || null;
      const previousMonth = previousYearMonths[i] || null;
      
      // Use current year month for display since we're only iterating through actual current year data
      const displayMonth = currentMonth || '';
      const monthStartQuoted = currentMonth ? `"${currentMonth}"` : '""';
      const previousMonthQuoted = previousMonth ? `"${previousMonth}"` : '""';
      
      // Current year formulas (only if current month exists) - now with brand campaign filter
      const clicksCurrentFormula = currentMonth ? 
        `=IF(B3="Yes",IF(B2="All",SUMIFS(Raw!G:G,Raw!A:A,${monthStartQuoted},Raw!E:E,FALSE),SUMIFS(Raw!G:G,Raw!A:A,${monthStartQuoted},Raw!D:D,B2,Raw!E:E,FALSE)),IF(B2="All",SUMIFS(Raw!G:G,Raw!A:A,${monthStartQuoted}),SUMIFS(Raw!G:G,Raw!A:A,${monthStartQuoted},Raw!D:D,B2)))` : 
        '0';
      const costCurrentFormula = currentMonth ? 
        `=IF(B3="Yes",IF(B2="All",SUMIFS(Raw!H:H,Raw!A:A,${monthStartQuoted},Raw!E:E,FALSE),SUMIFS(Raw!H:H,Raw!A:A,${monthStartQuoted},Raw!D:D,B2,Raw!E:E,FALSE)),IF(B2="All",SUMIFS(Raw!H:H,Raw!A:A,${monthStartQuoted}),SUMIFS(Raw!H:H,Raw!A:A,${monthStartQuoted},Raw!D:D,B2)))` : 
        '0';
      const conversionsCurrentFormula = currentMonth ? 
        `=IF(B3="Yes",IF(B2="All",SUMIFS(Raw!J:J,Raw!A:A,${monthStartQuoted},Raw!E:E,FALSE),SUMIFS(Raw!J:J,Raw!A:A,${monthStartQuoted},Raw!D:D,B2,Raw!E:E,FALSE)),IF(B2="All",SUMIFS(Raw!J:J,Raw!A:A,${monthStartQuoted}),SUMIFS(Raw!J:J,Raw!A:A,${monthStartQuoted},Raw!D:D,B2)))` : 
        '0';
      const conversionValueCurrentFormula = currentMonth ? 
        `=IF(B3="Yes",IF(B2="All",SUMIFS(Raw!L:L,Raw!A:A,${monthStartQuoted},Raw!E:E,FALSE),SUMIFS(Raw!L:L,Raw!A:A,${monthStartQuoted},Raw!D:D,B2,Raw!E:E,FALSE)),IF(B2="All",SUMIFS(Raw!L:L,Raw!A:A,${monthStartQuoted}),SUMIFS(Raw!L:L,Raw!A:A,${monthStartQuoted},Raw!D:D,B2)))` : 
        '0';
      const conversionValueTimeCurrentFormula = currentMonth ? 
        `=IF(B3="Yes",IF(B2="All",SUMIFS(Raw!M:M,Raw!A:A,${monthStartQuoted},Raw!E:E,FALSE),SUMIFS(Raw!M:M,Raw!A:A,${monthStartQuoted},Raw!D:D,B2,Raw!E:E,FALSE)),IF(B2="All",SUMIFS(Raw!M:M,Raw!A:A,${monthStartQuoted}),SUMIFS(Raw!M:M,Raw!A:A,${monthStartQuoted},Raw!D:D,B2)))` : 
        '0';
      
      // Previous year formulas (only if previous month exists) - now with brand campaign filter
      const clicksPreviousFormula = previousMonth ? 
        `=IF(B3="Yes",IF(B2="All",SUMIFS(Raw!G:G,Raw!A:A,${previousMonthQuoted},Raw!E:E,FALSE),SUMIFS(Raw!G:G,Raw!A:A,${previousMonthQuoted},Raw!D:D,B2,Raw!E:E,FALSE)),IF(B2="All",SUMIFS(Raw!G:G,Raw!A:A,${previousMonthQuoted}),SUMIFS(Raw!G:G,Raw!A:A,${previousMonthQuoted},Raw!D:D,B2)))` : 
        '0';
      const costPreviousFormula = previousMonth ? 
        `=IF(B3="Yes",IF(B2="All",SUMIFS(Raw!H:H,Raw!A:A,${previousMonthQuoted},Raw!E:E,FALSE),SUMIFS(Raw!H:H,Raw!A:A,${previousMonthQuoted},Raw!D:D,B2,Raw!E:E,FALSE)),IF(B2="All",SUMIFS(Raw!H:H,Raw!A:A,${previousMonthQuoted}),SUMIFS(Raw!H:H,Raw!A:A,${previousMonthQuoted},Raw!D:D,B2)))` : 
        '0';
      const conversionsPreviousFormula = previousMonth ? 
        `=IF(B3="Yes",IF(B2="All",SUMIFS(Raw!J:J,Raw!A:A,${previousMonthQuoted},Raw!E:E,FALSE),SUMIFS(Raw!J:J,Raw!A:A,${previousMonthQuoted},Raw!D:D,B2,Raw!E:E,FALSE)),IF(B2="All",SUMIFS(Raw!J:J,Raw!A:A,${previousMonthQuoted}),SUMIFS(Raw!J:J,Raw!A:A,${previousMonthQuoted},Raw!D:D,B2)))` : 
        '0';
      const conversionValuePreviousFormula = previousMonth ? 
        `=IF(B3="Yes",IF(B2="All",SUMIFS(Raw!L:L,Raw!A:A,${previousMonthQuoted},Raw!E:E,FALSE),SUMIFS(Raw!L:L,Raw!A:A,${previousMonthQuoted},Raw!D:D,B2,Raw!E:E,FALSE)),IF(B2="All",SUMIFS(Raw!L:L,Raw!A:A,${previousMonthQuoted}),SUMIFS(Raw!L:L,Raw!A:A,${previousMonthQuoted},Raw!D:D,B2)))` : 
        '0';
      const conversionValueTimePreviousFormula = previousMonth ? 
        `=IF(B3="Yes",IF(B2="All",SUMIFS(Raw!M:M,Raw!A:A,${previousMonthQuoted},Raw!E:E,FALSE),SUMIFS(Raw!M:M,Raw!A:A,${previousMonthQuoted},Raw!D:D,B2,Raw!E:E,FALSE)),IF(B2="All",SUMIFS(Raw!M:M,Raw!A:A,${previousMonthQuoted}),SUMIFS(Raw!M:M,Raw!A:A,${previousMonthQuoted},Raw!D:D,B2)))` : 
        '0';
      
      // Calculate Average CPC, ROAS, and Index YoY dynamically
      const rowNum = i + 8; // Starting at row 8 now (after app title + 2 filter rows + client name + 2 header rows)
      const avgCPCCurrentFormula = `=IF(C${rowNum}>0,I${rowNum}/C${rowNum},0)`;
      const avgCPCPreviousFormula = `=IF(D${rowNum}>0,J${rowNum}/D${rowNum},0)`;
      const roasCurrentFormula = `=IF(I${rowNum}>0,O${rowNum}/I${rowNum},0)`;
      const roasPreviousFormula = `=IF(J${rowNum}>0,P${rowNum}/J${rowNum},0)`;
      const roasTimeCurrentFormula = `=IF(I${rowNum}>0,R${rowNum}/I${rowNum},0)`;
      const roasTimePreviousFormula = `=IF(J${rowNum}>0,S${rowNum}/J${rowNum},0)`;
      
      // Index YoY formulas (Current Year / Previous Year * 100)
      const clicksIndexFormula = `=IF(D${rowNum}>0,(C${rowNum}/D${rowNum})*100,0)`;
      const cpcIndexFormula = `=IF(G${rowNum}>0,(F${rowNum}/G${rowNum})*100,0)`;
      const costIndexFormula = `=IF(J${rowNum}>0,(I${rowNum}/J${rowNum})*100,0)`;
      const conversionsIndexFormula = `=IF(M${rowNum}>0,(L${rowNum}/M${rowNum})*100,0)`;
      const conversionValueIndexFormula = `=IF(P${rowNum}>0,(O${rowNum}/P${rowNum})*100,0)`;
      const conversionValueTimeIndexFormula = `=IF(S${rowNum}>0,(R${rowNum}/S${rowNum})*100,0)`;
      const roasIndexFormula = `=IF(V${rowNum}>0,(U${rowNum}/V${rowNum})*100,0)`;
      const roasTimeIndexFormula = `=IF(Y${rowNum}>0,(X${rowNum}/Y${rowNum})*100,0)`;
      
      const formulaRow = [
        monthNumber,                        // A: Month Number
        displayMonth,                       // B: Month Start
        clicksCurrentFormula,               // C: Clicks Current Year (2025)
        clicksPreviousFormula,              // D: Clicks Previous Year (2024)
        clicksIndexFormula,                 // E: Clicks Index YoY
        avgCPCCurrentFormula,               // F: Avg CPC Current Year (2025)
        avgCPCPreviousFormula,              // G: Avg CPC Previous Year (2024)
        cpcIndexFormula,                    // H: Avg CPC Index YoY
        costCurrentFormula,                 // I: Cost Current Year (2025)
        costPreviousFormula,                // J: Cost Previous Year (2024)
        costIndexFormula,                   // K: Cost Index YoY
        conversionsCurrentFormula,          // L: Conversions Current Year (2025)
        conversionsPreviousFormula,         // M: Conversions Previous Year (2024)
        conversionsIndexFormula,            // N: Conversions Index YoY
        conversionValueCurrentFormula,      // O: Conversion Value Current Year (2025)
        conversionValuePreviousFormula,     // P: Conversion Value Previous Year (2024)
        conversionValueIndexFormula,        // Q: Conversion Value Index YoY
        conversionValueTimeCurrentFormula,  // R: Conversion Value (Time) Current Year (2025)
        conversionValueTimePreviousFormula, // S: Conversion Value (Time) Previous Year (2024)
        conversionValueTimeIndexFormula,    // T: Conversion Value (Time) Index YoY
        roasCurrentFormula,                 // U: ROAS Current Year (2025)
        roasPreviousFormula,                // V: ROAS Previous Year (2024)
        roasIndexFormula,                   // W: ROAS Index YoY
        roasTimeCurrentFormula,             // X: ROAS (Time) Current Year (2025)
        roasTimePreviousFormula,            // Y: ROAS (Time) Previous Year (2024)
        roasTimeIndexFormula                // Z: ROAS (Time) Index YoY
      ];
      
      formulaData.push(formulaRow);
    }
    
    // Write formula-based data to summary sheet (starting at row 8)
    if (formulaData.length > 0) {
      const summaryRange = summarySheet.getRange(8, 1, formulaData.length, 26);
      summaryRange.setValues(formulaData);
    }
    
    // Format main header (row 6 only)
    const mainHeaderRange = summarySheet.getRange(6, 1, 1, 26);
    mainHeaderRange.setFontWeight('bold');
    mainHeaderRange.setBackground('#34a853');
    mainHeaderRange.setFontColor('#ffffff');
    mainHeaderRange.setHorizontalAlignment('center');
    mainHeaderRange.setVerticalAlignment('middle');
    
    // Format year subheader (row 7) with light green background
    const yearSubHeaderRange = summarySheet.getRange(7, 1, 1, 26);
    yearSubHeaderRange.setFontWeight('bold');
    yearSubHeaderRange.setBackground('#c8e6c9'); // Light green background
    yearSubHeaderRange.setHorizontalAlignment('center');
    yearSubHeaderRange.setVerticalAlignment('middle');
    
    // Format data rows (starting at row 7)
    if (formulaData.length > 0) {
      const dataRowCount = formulaData.length;
      
      // Add alternating row backgrounds
      for (let i = 0; i < dataRowCount; i++) {
        const rowNum = i + 8;
        if (i % 2 === 1) { // Every other row (odd rows)
          const rowRange = summarySheet.getRange(rowNum, 1, 1, 26);
          rowRange.setBackground('#f8f9fa'); // Light gray background
        }
      }
      
      // Add thicker borders around each metric group (starting from row 6 for headers)
      const totalRows = dataRowCount + 2; // Include header rows 6-7
      
      // Clicks group border (C6:E through last data row)
      const clicksGroupRange = summarySheet.getRange(6, 3, totalRows, 3); // Columns C-E
      clicksGroupRange.setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      
      // Average CPC group border (F6:H through last data row)
      const cpcGroupRange = summarySheet.getRange(6, 6, totalRows, 3); // Columns F-H
      cpcGroupRange.setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      
      // Cost group border (I6:K through last data row)
      const costGroupRange = summarySheet.getRange(6, 9, totalRows, 3); // Columns I-K
      costGroupRange.setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      
      // Conversions group border (L6:N through last data row)
      const conversionsGroupRange = summarySheet.getRange(6, 12, totalRows, 3); // Columns L-N
      conversionsGroupRange.setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      
      // Conversion Value group border (O6:Q through last data row)
      const convValueGroupRange = summarySheet.getRange(6, 15, totalRows, 3); // Columns O-Q
      convValueGroupRange.setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      
      // Conversion Value (Time) group border (R6:T through last data row)
      const convValueTimeGroupRange = summarySheet.getRange(6, 18, totalRows, 3); // Columns R-T
      convValueTimeGroupRange.setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      
      // ROAS group border (U6:W through last data row)
      const roasGroupRange = summarySheet.getRange(6, 21, totalRows, 3); // Columns U-W
      roasGroupRange.setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      
      // ROAS (Time) group border (X6:Z through last data row)
      const roasTimeGroupRange = summarySheet.getRange(6, 24, totalRows, 3); // Columns X-Z
      roasTimeGroupRange.setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      
      // Month Number (Column A) - Plain numbers, center aligned
      const monthNumberRange = summarySheet.getRange(8, 1, dataRowCount, 1);
      monthNumberRange.setNumberFormat('0');
      monthNumberRange.setHorizontalAlignment('center');
      
      // Month Start (Column B) - Keep as date, center aligned
      const monthStartRange = summarySheet.getRange(8, 2, dataRowCount, 1);
      monthStartRange.setHorizontalAlignment('center');
      
      // Clicks (Columns C & D) - Number format, center aligned
      const clicksRange = summarySheet.getRange(8, 3, dataRowCount, 2);
      clicksRange.setNumberFormat('#,##0');
      clicksRange.setHorizontalAlignment('center');
      
      // Clicks Index (Column E) - Percentage format, center aligned
      const clicksIndexRange = summarySheet.getRange(8, 5, dataRowCount, 1);
      clicksIndexRange.setNumberFormat('#,##0"%"');
      clicksIndexRange.setHorizontalAlignment('center');
      
      // Average CPC (Columns F & G) - Currency format with 2 decimals, center aligned
      const cpcRange = summarySheet.getRange(8, 6, dataRowCount, 2);
      cpcRange.setNumberFormat('$#,##0.00');
      cpcRange.setHorizontalAlignment('center');
      
      // CPC Index (Column H) - Percentage format, center aligned
      const cpcIndexRange = summarySheet.getRange(8, 8, dataRowCount, 1);
      cpcIndexRange.setNumberFormat('#,##0"%"');
      cpcIndexRange.setHorizontalAlignment('center');
      
      // Cost (Columns I & J) - Currency format without decimals, center aligned
      const costRange = summarySheet.getRange(8, 9, dataRowCount, 2);
      costRange.setNumberFormat('$#,##0');
      costRange.setHorizontalAlignment('center');
      
      // Cost Index (Column K) - Percentage format, center aligned
      const costIndexRange = summarySheet.getRange(8, 11, dataRowCount, 1);
      costIndexRange.setNumberFormat('#,##0"%"');
      costIndexRange.setHorizontalAlignment('center');
      
      // Conversions (Columns L & M) - Number format with 2 decimals, center aligned
      const conversionsRange = summarySheet.getRange(8, 12, dataRowCount, 2);
      conversionsRange.setNumberFormat('#,##0.00');
      conversionsRange.setHorizontalAlignment('center');
      
      // Conversions Index (Column N) - Percentage format, center aligned
      const conversionsIndexRange = summarySheet.getRange(8, 14, dataRowCount, 1);
      conversionsIndexRange.setNumberFormat('#,##0"%"');
      conversionsIndexRange.setHorizontalAlignment('center');
      
      // Conversion Value (Columns O & P) - Currency format without decimals, center aligned
      const conversionValueRange = summarySheet.getRange(8, 15, dataRowCount, 2);
      conversionValueRange.setNumberFormat('$#,##0');
      conversionValueRange.setHorizontalAlignment('center');
      
      // Conversion Value Index (Column Q) - Percentage format, center aligned
      const conversionValueIndexRange = summarySheet.getRange(8, 17, dataRowCount, 1);
      conversionValueIndexRange.setNumberFormat('#,##0"%"');
      conversionValueIndexRange.setHorizontalAlignment('center');
      
      // Conversion Value (Time) (Columns R & S) - Currency format without decimals, center aligned
      const conversionValueTimeRange = summarySheet.getRange(8, 18, dataRowCount, 2);
      conversionValueTimeRange.setNumberFormat('$#,##0');
      conversionValueTimeRange.setHorizontalAlignment('center');
      
      // Conversion Value (Time) Index (Column T) - Percentage format, center aligned
      const conversionValueTimeIndexRange = summarySheet.getRange(8, 20, dataRowCount, 1);
      conversionValueTimeIndexRange.setNumberFormat('#,##0"%"');
      conversionValueTimeIndexRange.setHorizontalAlignment('center');
      
      // ROAS (Columns U & V) - Number format with 2 decimals, center aligned
      const roasRange = summarySheet.getRange(8, 21, dataRowCount, 2);
      roasRange.setNumberFormat('#,##0.00');
      roasRange.setHorizontalAlignment('center');
      
      // ROAS Index (Column W) - Percentage format, center aligned
      const roasIndexRange = summarySheet.getRange(8, 23, dataRowCount, 1);
      roasIndexRange.setNumberFormat('#,##0"%"');
      roasIndexRange.setHorizontalAlignment('center');
      
      // ROAS (Time) (Columns X & Y) - Number format with 2 decimals, center aligned
      const roasTimeRange = summarySheet.getRange(8, 24, dataRowCount, 2);
      roasTimeRange.setNumberFormat('#,##0.00');
      roasTimeRange.setHorizontalAlignment('center');
      
      // ROAS (Time) Index (Column Z) - Percentage format, center aligned
      const roasTimeIndexRange = summarySheet.getRange(8, 26, dataRowCount, 1);
      roasTimeIndexRange.setNumberFormat('#,##0"%"');
      roasTimeIndexRange.setHorizontalAlignment('center');
      
      // Apply conditional formatting to Index YoY columns with very light gradient colors
      const indexColumns = [5, 8, 11, 14, 17, 20, 23, 26]; // E, H, K, N, Q, T, W, Z
      
      indexColumns.forEach(colIndex => {
        const indexRange = summarySheet.getRange(8, colIndex, dataRowCount, 1);
        
        // Special handling for Average CPC (column H) - reverse colors since lower is better
        if (colIndex === 8) {
          // Apply reversed conditional formatting for Average CPC Index YoY
          applyReversedCPCFormatting(summarySheet, indexRange, dataRowCount);
          return; // Skip the normal formatting for this column
        }
        
        // Create multiple conditional formatting rules for very smooth gradient effect
        const rules = [];
        
        // GREEN GRADIENT SCALE (Above 100%) - Slightly more intense
        
        // Exceptional: >= 140% (Medium Light Green)
        const exceptionalRule = SpreadsheetApp.newConditionalFormatRule()
          .whenNumberGreaterThanOrEqualTo(140)
          .setBackground('#a5d6a7')
          .setRanges([indexRange])
          .build();
        rules.push(exceptionalRule);
        
        // Excellent: 130-139% (Light Green)
        const excellentRule = SpreadsheetApp.newConditionalFormatRule()
          .whenNumberBetween(130, 139)
          .setBackground('#c8e6c9')
          .setRanges([indexRange])
          .build();
        rules.push(excellentRule);
        
        // Very good: 125-129% (Light Medium Green)
        const veryGoodRule = SpreadsheetApp.newConditionalFormatRule()
          .whenNumberBetween(125, 129)
          .setBackground('#dcedc8')
          .setRanges([indexRange])
          .build();
        rules.push(veryGoodRule);
        
        // Good: 120-124% (Visible Light Green)
        const goodRule = SpreadsheetApp.newConditionalFormatRule()
          .whenNumberBetween(120, 124)
          .setBackground('#e8f5e8')
          .setRanges([indexRange])
          .build();
        rules.push(goodRule);
        
        // Above average: 115-119% (Noticeable Light Green)
        const aboveAverageRule = SpreadsheetApp.newConditionalFormatRule()
          .whenNumberBetween(115, 119)
          .setBackground('#f1f8e9')
          .setRanges([indexRange])
          .build();
        rules.push(aboveAverageRule);
        
        // Slightly good: 110-114% (Light Green - more visible)
        const slightlyGoodRule = SpreadsheetApp.newConditionalFormatRule()
          .whenNumberBetween(110, 114)
          .setBackground('#f4faf5')
          .setRanges([indexRange])
          .build();
        rules.push(slightlyGoodRule);
        
        // Just above baseline: 105-109% (Subtle Green - more visible)
        const justAboveRule = SpreadsheetApp.newConditionalFormatRule()
          .whenNumberBetween(105, 109)
          .setBackground('#f7fcf8')
          .setRanges([indexRange])
          .build();
        rules.push(justAboveRule);
        
        // Slightly above baseline: 100-104% (Very Subtle Green - more visible)
        const slightlyAboveRule = SpreadsheetApp.newConditionalFormatRule()
          .whenNumberBetween(100, 104)
          .setBackground('#f9fdf9')
          .setRanges([indexRange])
          .build();
        rules.push(slightlyAboveRule);
        
        // RED GRADIENT SCALE (Below 100%)
        
        // Just below baseline: 95-99% (Barely Pink)
        const justBelowRule = SpreadsheetApp.newConditionalFormatRule()
          .whenNumberBetween(95, 99)
          .setBackground('#fef9f7')
          .setRanges([indexRange])
          .build();
        rules.push(justBelowRule);
        
        // Slightly below: 90-94% (Very Faint Pink)
        const slightlyBelowRule = SpreadsheetApp.newConditionalFormatRule()
          .whenNumberBetween(90, 94)
          .setBackground('#fef5f5')
          .setRanges([indexRange])
          .build();
        rules.push(slightlyBelowRule);
        
        // Below average: 85-89% (Light Pink)
        const belowAverageRule = SpreadsheetApp.newConditionalFormatRule()
          .whenNumberBetween(85, 89)
          .setBackground('#ffebee')
          .setRanges([indexRange])
          .build();
        rules.push(belowAverageRule);
        
        // Concerning: 80-84% (Slightly More Pink)
        const concerningRule = SpreadsheetApp.newConditionalFormatRule()
          .whenNumberBetween(80, 84)
          .setBackground('#ffe0e1')
          .setRanges([indexRange])
          .build();
        rules.push(concerningRule);
        
        // Poor: 75-79% (Light Red)
        const poorRule = SpreadsheetApp.newConditionalFormatRule()
          .whenNumberBetween(75, 79)
          .setBackground('#ffcdd2')
          .setRanges([indexRange])
          .build();
        rules.push(poorRule);
        
        // Very poor: 70-74% (Light Medium Red)
        const veryPoorRule = SpreadsheetApp.newConditionalFormatRule()
          .whenNumberBetween(70, 74)
          .setBackground('#ffb3ba')
          .setRanges([indexRange])
          .build();
        rules.push(veryPoorRule);
        
        // Bad: 65-69% (Medium Light Red)
        const badRule = SpreadsheetApp.newConditionalFormatRule()
          .whenNumberBetween(65, 69)
          .setBackground('#ff9aa2')
          .setRanges([indexRange])
          .build();
        rules.push(badRule);
        
        // Very bad: 60-64% (Noticeable Red)
        const veryBadRule = SpreadsheetApp.newConditionalFormatRule()
          .whenNumberBetween(60, 64)
          .setBackground('#ff8a95')
          .setRanges([indexRange])
          .build();
        rules.push(veryBadRule);
        
        // Extremely poor: < 60% (Most Intense but still light)
        const extremelyPoorRule = SpreadsheetApp.newConditionalFormatRule()
          .whenNumberLessThan(60)
          .setBackground('#ff7782')
          .setRanges([indexRange])
          .build();
        rules.push(extremelyPoorRule);
        
        // Apply all rules to the sheet
        const existingRules = summarySheet.getConditionalFormatRules();
        summarySheet.setConditionalFormatRules(existingRules.concat(rules));
      });
    }
    
    // Function to apply reversed conditional formatting for Average CPC (lower is better)
    function applyReversedCPCFormatting(sheet, range, dataRowCount) {
      const rules = [];
      
      // RED GRADIENT SCALE (Above 100% - worse performance for CPC)
      
      // Very bad: >= 140% (Medium Light Red)
      const veryBadRule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThanOrEqualTo(140)
        .setBackground('#ff9aa2')
        .setRanges([range])
        .build();
      rules.push(veryBadRule);
      
      // Bad: 130-139% (Light Red)
      const badRule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberBetween(130, 139)
        .setBackground('#ffb3ba')
        .setRanges([range])
        .build();
      rules.push(badRule);
      
      // Poor: 125-129% (Light Pink)
      const poorRule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberBetween(125, 129)
        .setBackground('#ffcdd2')
        .setRanges([range])
        .build();
      rules.push(poorRule);
      
      // Concerning: 120-124% (Slightly Pink)
      const concerningRule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberBetween(120, 124)
        .setBackground('#ffe0e1')
        .setRanges([range])
        .build();
      rules.push(concerningRule);
      
      // Below average: 115-119% (Light Pink)
      const belowAverageRule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberBetween(115, 119)
        .setBackground('#ffebee')
        .setRanges([range])
        .build();
      rules.push(belowAverageRule);
      
      // Slightly concerning: 110-114% (Very Faint Pink)
      const slightlyConcerningRule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberBetween(110, 114)
        .setBackground('#fef5f5')
        .setRanges([range])
        .build();
      rules.push(slightlyConcerningRule);
      
      // Just above baseline: 105-109% (Barely Pink)
      const justAboveRule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberBetween(105, 109)
        .setBackground('#fef9f7')
        .setRanges([range])
        .build();
      rules.push(justAboveRule);
      
      // Slightly above baseline: 100-104% (Almost neutral)
      const slightlyAboveRule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberBetween(100, 104)
        .setBackground('#fefcfc')
        .setRanges([range])
        .build();
      rules.push(slightlyAboveRule);
      
      // GREEN GRADIENT SCALE (Below 100% - better performance for CPC)
      
      // Just below baseline: 95-99% (Very Subtle Green)
      const justBelowRule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberBetween(95, 99)
        .setBackground('#f9fdf9')
        .setRanges([range])
        .build();
      rules.push(justBelowRule);
      
      // Slightly below: 90-94% (Subtle Green)
      const slightlyBelowRule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberBetween(90, 94)
        .setBackground('#f7fcf8')
        .setRanges([range])
        .build();
      rules.push(slightlyBelowRule);
      
      // Below average: 85-89% (Light Green)
      const belowAverageGreenRule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberBetween(85, 89)
        .setBackground('#f4faf5')
        .setRanges([range])
        .build();
      rules.push(belowAverageGreenRule);
      
      // Good: 80-84% (Noticeable Light Green)
      const goodRule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberBetween(80, 84)
        .setBackground('#f1f8e9')
        .setRanges([range])
        .build();
      rules.push(goodRule);
      
      // Very good: 75-79% (Visible Light Green)
      const veryGoodRule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberBetween(75, 79)
        .setBackground('#e8f5e8')
        .setRanges([range])
        .build();
      rules.push(veryGoodRule);
      
      // Excellent: 70-74% (Light Medium Green)
      const excellentRule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberBetween(70, 74)
        .setBackground('#dcedc8')
        .setRanges([range])
        .build();
      rules.push(excellentRule);
      
      // Outstanding: 65-69% (Light Green)
      const outstandingRule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberBetween(65, 69)
        .setBackground('#c8e6c9')
        .setRanges([range])
        .build();
      rules.push(outstandingRule);
      
      // Exceptional: < 65% (Medium Light Green)
      const exceptionalRule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThan(65)
        .setBackground('#a5d6a7')
        .setRanges([range])
        .build();
      rules.push(exceptionalRule);
      
      // Apply the rules
      const existingRules = sheet.getConditionalFormatRules();
      sheet.setConditionalFormatRules(existingRules.concat(rules));
    }
    
    // Set column widths - make column A wider for app title and filter text
    summarySheet.setColumnWidth(1, 180);  // App Title / Month Number / Filter text (wider for full display)
    summarySheet.setColumnWidth(2, 100);  // Month Start
    
    // Set equal widths for Clicks columns (C, D, E)
    summarySheet.setColumnWidth(3, 90);   // Clicks 2025
    summarySheet.setColumnWidth(4, 90);   // Clicks 2024
    summarySheet.setColumnWidth(5, 80);   // Clicks Index YoY
    
    // Set equal widths for Average CPC columns (F, G, H)
    summarySheet.setColumnWidth(6, 90);   // Avg CPC 2025
    summarySheet.setColumnWidth(7, 90);   // Avg CPC 2024
    summarySheet.setColumnWidth(8, 80);   // Avg CPC Index YoY
    
    // Set equal widths for Cost columns (I, J, K)
    summarySheet.setColumnWidth(9, 90);   // Cost 2025
    summarySheet.setColumnWidth(10, 90);  // Cost 2024
    summarySheet.setColumnWidth(11, 80);  // Cost Index YoY
    
    // Set equal widths for Conversions columns (L, M, N)
    summarySheet.setColumnWidth(12, 100); // Conversions 2025
    summarySheet.setColumnWidth(13, 100); // Conversions 2024
    summarySheet.setColumnWidth(14, 80);  // Conversions Index YoY
    
    // Set equal widths for Conversion Value columns (O, P, Q)
    summarySheet.setColumnWidth(15, 110); // Conv Value 2025
    summarySheet.setColumnWidth(16, 110); // Conv Value 2024
    summarySheet.setColumnWidth(17, 80);  // Conv Value Index YoY
    
    // Set equal widths for Conversion Value (Time) columns (R, S, T)
    summarySheet.setColumnWidth(18, 110); // Conv Value (Time) 2025
    summarySheet.setColumnWidth(19, 110); // Conv Value (Time) 2024
    summarySheet.setColumnWidth(20, 80);  // Conv Value (Time) Index YoY
    
    // Set equal widths for ROAS columns (U, V, W)
    summarySheet.setColumnWidth(21, 80);  // ROAS 2025
    summarySheet.setColumnWidth(22, 80);  // ROAS 2024
    summarySheet.setColumnWidth(23, 80);  // ROAS Index YoY
    
    // Set equal widths for ROAS (Time) columns (X, Y, Z)
    summarySheet.setColumnWidth(24, 80);  // ROAS (Time) 2025
    summarySheet.setColumnWidth(25, 80);  // ROAS (Time) 2024
    summarySheet.setColumnWidth(26, 80);  // ROAS (Time) Index YoY
    
    // Freeze the filter and header rows for better navigation
    summarySheet.setFrozenRows(7); // Freeze rows 1-7 (app title + filters + client name + headers)
    summarySheet.setFrozenColumns(2); // Freeze columns A-B (month info)
    
    Logger.log(`✅ Monthly Ads Monitor dashboard created with ${maxMonths} months - filters update data automatically!`);
    
  } catch (error) {
    Logger.log('❌ Error creating Monthly Ads Monitor dashboard: ' + error.toString());
  }
} 