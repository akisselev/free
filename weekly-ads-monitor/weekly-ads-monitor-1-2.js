// Enter your Google Sheet URL here between the single quotes.
const SHEET_URL = 'https://docs.google.com/spreadsheets/d/1ikUbwYIxuc-sTq5y01U8pwQsf5khfEChiVLKArPOzSM/edit?gid=228289331#gid=228289331';

// Calculate date range from beginning of last year to last complete Sunday (2 years of data)
function getDateRange() {
  const today = new Date();
  const endDate = new Date(today);
  
  // Calculate last complete Sunday
  const dayOfWeek = today.getDay(); // 0 = Sunday, 1 = Monday, etc.
  
  if (dayOfWeek === 0) {
    // If today is Sunday, go back to previous Sunday (7 days ago) since today is not complete
    endDate.setDate(today.getDate() - 7);
  } else {
    // If today is not Sunday, go back to the most recent Sunday
    endDate.setDate(today.getDate() - dayOfWeek);
  }
  
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

// Weekly campaign performance data with campaign type
const GAQL_QUERY = `
SELECT
  campaign.name,
  campaign.advertising_channel_type,
  segments.week,
  metrics.impressions,
  metrics.clicks,
  metrics.cost_micros,
  metrics.conversions,
  metrics.conversions_value
FROM campaign
WHERE metrics.impressions > 0
  AND segments.date BETWEEN "${START_DATE}" AND "${END_DATE}"
ORDER BY segments.week DESC, metrics.cost_micros DESC
`;

function main() {
  try {
    Logger.log('🚀 STARTING CAMPAIGN TRENDS EXPORT');
    Logger.log(`📅 Date Range (2-year comparison up to last complete Sunday): ${START_DATE} to ${END_DATE}`);
    
    let spreadsheet;
    
    if (SHEET_URL === '') {
      // Create a new spreadsheet if no URL is provided
      spreadsheet = SpreadsheetApp.create('Campaign Trends Export - ' + 
        Utilities.formatDate(new Date(), AdsApp.currentAccount().getTimeZone(), 'yyyy-MM-dd'));
      Logger.log('✅ NEW SPREADSHEET CREATED: ' + spreadsheet.getUrl());
    } else {
      // Open existing spreadsheet
      spreadsheet = SpreadsheetApp.openByUrl(SHEET_URL);
      Logger.log('✅ Opened existing spreadsheet');
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
    
    Logger.log(`📊 Found ${rawData.length} weekly campaign records`);
    
    // Create headers for raw data - 11 columns with campaign performance metrics
    const rawHeaders = [
      'Week Start',
      'Week End',
      'Campaign Name',
      'Campaign Type',
      'Impressions',
      'Clicks',
      'Cost',
      'Cost Per Click',
      'Conversions',
      'Cost Per Conversion',
      'Conversion Value'
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
    
    // Create aggregated weekly summary tab
    Logger.log('📊 Creating weekly summary...');
    createWeeklySummary(spreadsheet, rawData);
    
    Logger.log('🎉 SUCCESS! Export completed');
    Logger.log(`✅ Exported ${rawData.length} weekly campaign records with performance metrics`);
    Logger.log('🔗 Spreadsheet URL: ' + spreadsheet.getUrl());
    
  } catch (error) {
    Logger.log('❌ ERROR: ' + error.toString());
  }
}

function processData(rows) {
  let data = [];
  let count = 0;
  
  Logger.log('⚙️ Processing weekly campaign performance data...');
  
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
      
      let weekStart = segments.week || '';
      let impressions = Number(metrics.impressions) || 0;
      let clicks = Number(metrics.clicks) || 0;
      let costMicros = Number(metrics.costMicros) || 0;
      let conversions = Number(metrics.conversions) || 0;
      let conversionValue = Number(metrics.conversionsValue) || 0;
      
      // Calculate week end date (add 6 days to week start)
      let weekEnd = '';
      if (weekStart) {
        try {
          let startDate = new Date(weekStart);
          let endDate = new Date(startDate);
          endDate.setDate(startDate.getDate() + 6);
          weekEnd = Utilities.formatDate(endDate, AdsApp.currentAccount().getTimeZone(), 'yyyy-MM-dd');
        } catch (e) {
          weekEnd = weekStart; // fallback to start date if calculation fails
        }
      }
      
      // Convert cost from micros to currency and calculate derived metrics
      let cost = Math.round((costMicros / 1000000) * 100) / 100;
      conversionValue = Math.round(conversionValue * 100) / 100;
      
      // Calculate cost per click
      let costPerClick = clicks > 0 ? Math.round((cost / clicks) * 100) / 100 : 0;
      
      // Calculate cost per conversion
      let costPerConversion = conversions > 0 ? Math.round((cost / conversions) * 100) / 100 : 0;
      
      // Create row - 11 columns with campaign performance metrics
      let newRow = [
        weekStart,           // Week Start
        weekEnd,             // Week End
        campaignName,        // Campaign Name
        campaignType,        // Campaign Type
        impressions,         // Impressions
        clicks,              // Clicks
        cost,                // Cost
        costPerClick,        // Cost Per Click
        conversions,         // Conversions
        costPerConversion,   // Cost Per Conversion
        conversionValue      // Conversion Value
      ];
      
      data.push(newRow);
      
      if (count % 500 === 0) {
        Logger.log(`📈 Processed ${count} weekly campaign records...`);
      }
      
    } catch (e) {
      Logger.log(`❌ Error processing row: ${e}`);
    }
  }
  
  Logger.log(`✅ Processing complete: ${data.length} records`);
  return data;
}

function createWeeklySummary(spreadsheet, rawData) {
  try {
    // Get or create the "Weekly Summary" tab
    let summarySheet = spreadsheet.getSheetByName('Weekly Summary');
    if (!summarySheet) {
      summarySheet = spreadsheet.insertSheet('Weekly Summary');
      Logger.log('✅ Created new "Weekly Summary" tab');
    } else {
      Logger.log('✅ Using existing "Weekly Summary" tab');
    }
    
    summarySheet.clear();
    
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
    
    // Debug: Check Shopping campaign data specifically
    const shoppingCurrentYear = currentYearData.filter(row => row[3] === 'Shopping');
    const shoppingPreviousYear = previousYearData.filter(row => row[3] === 'Shopping');
    
    Logger.log(`🛍️ Shopping campaigns ${debugCurrentYear}: ${shoppingCurrentYear.length} records`);
    Logger.log(`🛍️ Shopping campaigns ${debugPreviousYear}: ${shoppingPreviousYear.length} records`);
    
    if (shoppingPreviousYear.length === 0) {
      Logger.log(`⚠️ WARNING: No Shopping campaign data found for ${debugPreviousYear}`);
      Logger.log(`📋 Available campaign types in ${debugPreviousYear}: ${[...new Set(previousYearData.map(row => row[3]))].join(', ')}`);
    }
    
    // Add campaign type filter controls at the top
    summarySheet.getRange('A1').setValue('Campaign Type Filter:');
    summarySheet.getRange('A1').setFontWeight('bold');
    summarySheet.getRange('A1').setBackground('#f0f0f0');
    
    // Set up dropdown for campaign type selection
    const filterCell = summarySheet.getRange('B1');
    
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
    
    // Add instructions
    summarySheet.getRange('D1').setValue('🔄 Dynamic Filter: Select campaign type from dropdown → Data updates automatically!');
    summarySheet.getRange('D1').setFontStyle('italic');
    summarySheet.getRange('D1').setFontSize(9);
    summarySheet.getRange('D1').setFontColor('#0066cc');
    
    // Read the final selected campaign type filter
    const selectedCampaignType = filterCell.getValue() || 'All';
    Logger.log(`📊 Dynamic filtering enabled for campaign type: ${selectedCampaignType}`);
    
    // Get current and previous year
    const currentYear = new Date().getFullYear();
    const previousYear = currentYear - 1;
    
    // Get all unique weeks from raw data and separate by year
    const allWeekDates = [...new Set(rawData.map(row => row[0]))].sort();
    const currentYearWeeks = allWeekDates.filter(week => week.includes(currentYear.toString()));
    const allPreviousYearWeeks = allWeekDates.filter(week => week.includes(previousYear.toString()));
    
    // Calculate current week number based on last complete Sunday to limit previous year data
    const today = new Date();
    const lastCompleteSunday = new Date(today);
    const dayOfWeek = today.getDay(); // 0 = Sunday, 1 = Monday, etc.
    
    if (dayOfWeek === 0) {
      // If today is Sunday, go back to previous Sunday (7 days ago) since today is not complete
      lastCompleteSunday.setDate(today.getDate() - 7);
    } else {
      // If today is not Sunday, go back to the most recent Sunday
      lastCompleteSunday.setDate(today.getDate() - dayOfWeek);
    }
    
    const startOfYear = new Date(lastCompleteSunday.getFullYear(), 0, 1);
    const pastDaysOfYear = (lastCompleteSunday - startOfYear) / 86400000;
    const currentWeekNumber = Math.ceil((pastDaysOfYear + startOfYear.getDay() + 1) / 7);
    
    // Filter previous year weeks to only include weeks up to current week number
    const previousYearWeeks = allPreviousYearWeeks.slice(0, currentWeekNumber);
    
    // Limit current year weeks to current week number as well
    const limitedCurrentYearWeeks = currentYearWeeks.slice(0, currentWeekNumber);
    
    Logger.log(`📅 Current week number: ${currentWeekNumber}`);
    Logger.log(`📅 Found ${limitedCurrentYearWeeks.length} weeks in ${currentYear} and ${previousYearWeeks.length} weeks in ${previousYear} (limited to week ${currentWeekNumber})`);
    
    // Create two-row header structure (rows 3-4)
    // Main headers (row 3) - now with 3 columns per metric
    const mainHeaders = ['Week Number', 'Week Start', 'Clicks', '', '', 'Average CPC', '', '', 'Cost', '', '', 'Conversion Value', '', '', 'ROAS', '', ''];
    summarySheet.getRange('A3:Q3').setValues([mainHeaders]);
    
    // Year subheaders (row 4) - including Index YoY columns
    const yearHeaders = ['', '', currentYear, previousYear, 'Index YoY', currentYear, previousYear, 'Index YoY', currentYear, previousYear, 'Index YoY', currentYear, previousYear, 'Index YoY', currentYear, previousYear, 'Index YoY'];
    const yearHeaderRange = summarySheet.getRange('A4:Q4');
    yearHeaderRange.setValues([yearHeaders]);
    
    // Format year headers as plain numbers (not currency) for year columns only
    summarySheet.getRange('C4:D4').setNumberFormat('0'); // Clicks years
    summarySheet.getRange('F4:G4').setNumberFormat('0'); // CPC years
    summarySheet.getRange('I4:J4').setNumberFormat('0'); // Cost years
    summarySheet.getRange('L4:M4').setNumberFormat('0'); // Conversion Value years
    summarySheet.getRange('O4:P4').setNumberFormat('0'); // ROAS years
    
    // Merge cells for main headers
    summarySheet.getRange('A3:A4').merge(); // Week Number
    summarySheet.getRange('B3:B4').merge(); // Week Start
    summarySheet.getRange('C3:E3').merge(); // Clicks
    summarySheet.getRange('F3:H3').merge(); // Average CPC
    summarySheet.getRange('I3:K3').merge(); // Cost
    summarySheet.getRange('L3:N3').merge(); // Conversion Value
    summarySheet.getRange('O3:Q3').merge(); // ROAS
    
    // Create dynamic formulas that reference the filter and raw data
    Logger.log('📤 Creating year-over-year dynamic formulas...');
    
    // Build formula-based data rows that update automatically
    const formulaData = [];
    
    // Use the actual number of complete weeks we have data for (limited to current complete weeks)
    const maxWeeks = limitedCurrentYearWeeks.length;
    
    for (let i = 0; i < maxWeeks; i++) {
      const weekNumber = i + 1;
      const currentWeek = limitedCurrentYearWeeks[i] || null;
      const previousWeek = previousYearWeeks[i] || null;
      
      // Use current year week for display since we're only iterating through actual current year data
      const displayWeek = currentWeek || '';
      const weekStartQuoted = currentWeek ? `"${currentWeek}"` : '""';
      const previousWeekQuoted = previousWeek ? `"${previousWeek}"` : '""';
      
      // Current year formulas (only if current week exists)
      const clicksCurrentFormula = currentWeek ? 
        `=IF(B1="All",SUMIFS(Raw!F:F,Raw!A:A,${weekStartQuoted}),SUMIFS(Raw!F:F,Raw!A:A,${weekStartQuoted},Raw!D:D,B1))` : 
        '0';
      const costCurrentFormula = currentWeek ? 
        `=IF(B1="All",SUMIFS(Raw!G:G,Raw!A:A,${weekStartQuoted}),SUMIFS(Raw!G:G,Raw!A:A,${weekStartQuoted},Raw!D:D,B1))` : 
        '0';
      const conversionValueCurrentFormula = currentWeek ? 
        `=IF(B1="All",SUMIFS(Raw!K:K,Raw!A:A,${weekStartQuoted}),SUMIFS(Raw!K:K,Raw!A:A,${weekStartQuoted},Raw!D:D,B1))` : 
        '0';
      
      // Previous year formulas (only if previous week exists)
      const clicksPreviousFormula = previousWeek ? 
        `=IF(B1="All",SUMIFS(Raw!F:F,Raw!A:A,${previousWeekQuoted}),SUMIFS(Raw!F:F,Raw!A:A,${previousWeekQuoted},Raw!D:D,B1))` : 
        '0';
      const costPreviousFormula = previousWeek ? 
        `=IF(B1="All",SUMIFS(Raw!G:G,Raw!A:A,${previousWeekQuoted}),SUMIFS(Raw!G:G,Raw!A:A,${previousWeekQuoted},Raw!D:D,B1))` : 
        '0';
      const conversionValuePreviousFormula = previousWeek ? 
        `=IF(B1="All",SUMIFS(Raw!K:K,Raw!A:A,${previousWeekQuoted}),SUMIFS(Raw!K:K,Raw!A:A,${previousWeekQuoted},Raw!D:D,B1))` : 
        '0';
      
      // Calculate Average CPC, ROAS, and Index YoY dynamically
      const rowNum = i + 5; // Starting at row 5 now
      const avgCPCCurrentFormula = `=IF(C${rowNum}>0,I${rowNum}/C${rowNum},0)`;
      const avgCPCPreviousFormula = `=IF(D${rowNum}>0,J${rowNum}/D${rowNum},0)`;
      const roasCurrentFormula = `=IF(I${rowNum}>0,L${rowNum}/I${rowNum},0)`;
      const roasPreviousFormula = `=IF(J${rowNum}>0,M${rowNum}/J${rowNum},0)`;
      
      // Index YoY formulas (Current Year / Previous Year * 100)
      const clicksIndexFormula = `=IF(D${rowNum}>0,(C${rowNum}/D${rowNum})*100,0)`;
      const cpcIndexFormula = `=IF(G${rowNum}>0,(F${rowNum}/G${rowNum})*100,0)`;
      const costIndexFormula = `=IF(J${rowNum}>0,(I${rowNum}/J${rowNum})*100,0)`;
      const conversionValueIndexFormula = `=IF(M${rowNum}>0,(L${rowNum}/M${rowNum})*100,0)`;
      const roasIndexFormula = `=IF(P${rowNum}>0,(O${rowNum}/P${rowNum})*100,0)`;
      
      const formulaRow = [
        weekNumber,                          // A: Week Number
        displayWeek,                        // B: Week Start
        clicksCurrentFormula,               // C: Clicks Current Year (2025)
        clicksPreviousFormula,              // D: Clicks Previous Year (2024)
        clicksIndexFormula,                 // E: Clicks Index YoY
        avgCPCCurrentFormula,               // F: Avg CPC Current Year (2025)
        avgCPCPreviousFormula,              // G: Avg CPC Previous Year (2024)
        cpcIndexFormula,                    // H: Avg CPC Index YoY
        costCurrentFormula,                 // I: Cost Current Year (2025)
        costPreviousFormula,                // J: Cost Previous Year (2024)
        costIndexFormula,                   // K: Cost Index YoY
        conversionValueCurrentFormula,      // L: Conversion Value Current Year (2025)
        conversionValuePreviousFormula,     // M: Conversion Value Previous Year (2024)
        conversionValueIndexFormula,        // N: Conversion Value Index YoY
        roasCurrentFormula,                 // O: ROAS Current Year (2025)
        roasPreviousFormula,                // P: ROAS Previous Year (2024)
        roasIndexFormula                    // Q: ROAS Index YoY
      ];
      
      formulaData.push(formulaRow);
    }
    
    // Write formula-based data to summary sheet (starting at row 5)
    if (formulaData.length > 0) {
      const summaryRange = summarySheet.getRange(5, 1, formulaData.length, 17);
      summaryRange.setValues(formulaData);
    }
    
    // Format main header (row 3 only)
    const mainHeaderRange = summarySheet.getRange(3, 1, 1, 17);
    mainHeaderRange.setFontWeight('bold');
    mainHeaderRange.setBackground('#34a853');
    mainHeaderRange.setFontColor('#ffffff');
    mainHeaderRange.setHorizontalAlignment('center');
    mainHeaderRange.setVerticalAlignment('middle');
    
    // Format year subheader (row 4) with light green background
    const yearSubHeaderRange = summarySheet.getRange(4, 1, 1, 17);
    yearSubHeaderRange.setFontWeight('bold');
    yearSubHeaderRange.setBackground('#c8e6c9'); // Light green background
    yearSubHeaderRange.setHorizontalAlignment('center');
    yearSubHeaderRange.setVerticalAlignment('middle');
    
    // Format data rows (starting at row 5)
    if (formulaData.length > 0) {
      const dataRowCount = formulaData.length;
      
      // Add alternating row backgrounds
      for (let i = 0; i < dataRowCount; i++) {
        const rowNum = i + 5;
        if (i % 2 === 1) { // Every other row (odd rows)
          const rowRange = summarySheet.getRange(rowNum, 1, 1, 17);
          rowRange.setBackground('#f8f9fa'); // Light gray background
        }
      }
      
      // Add thicker borders around each metric group (starting from row 3 for headers)
      const totalRows = dataRowCount + 2; // Include header rows 3-4
      
      // Clicks group border (C3:E through last data row)
      const clicksGroupRange = summarySheet.getRange(3, 3, totalRows, 3); // Columns C-E
      clicksGroupRange.setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      
      // Average CPC group border (F3:H through last data row)
      const cpcGroupRange = summarySheet.getRange(3, 6, totalRows, 3); // Columns F-H
      cpcGroupRange.setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      
      // Cost group border (I3:K through last data row)
      const costGroupRange = summarySheet.getRange(3, 9, totalRows, 3); // Columns I-K
      costGroupRange.setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      
      // Conversion Value group border (L3:N through last data row)
      const convValueGroupRange = summarySheet.getRange(3, 12, totalRows, 3); // Columns L-N
      convValueGroupRange.setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      
      // ROAS group border (O3:Q through last data row)
      const roasGroupRange = summarySheet.getRange(3, 15, totalRows, 3); // Columns O-Q
      roasGroupRange.setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      
      // Week Number (Column A) - Plain numbers, center aligned
      const weekNumberRange = summarySheet.getRange(5, 1, dataRowCount, 1);
      weekNumberRange.setNumberFormat('0');
      weekNumberRange.setHorizontalAlignment('center');
      
      // Week Start (Column B) - Keep as date, center aligned
      const weekStartRange = summarySheet.getRange(5, 2, dataRowCount, 1);
      weekStartRange.setHorizontalAlignment('center');
      
      // Clicks (Columns C & D) - Number format, center aligned
      const clicksRange = summarySheet.getRange(5, 3, dataRowCount, 2);
      clicksRange.setNumberFormat('#,##0');
      clicksRange.setHorizontalAlignment('center');
      
      // Clicks Index (Column E) - Percentage format, center aligned
      const clicksIndexRange = summarySheet.getRange(5, 5, dataRowCount, 1);
      clicksIndexRange.setNumberFormat('#,##0"%"');
      clicksIndexRange.setHorizontalAlignment('center');
      
      // Average CPC (Columns F & G) - Currency format with 2 decimals, center aligned
      const cpcRange = summarySheet.getRange(5, 6, dataRowCount, 2);
      cpcRange.setNumberFormat('$#,##0.00');
      cpcRange.setHorizontalAlignment('center');
      
      // CPC Index (Column H) - Percentage format, center aligned
      const cpcIndexRange = summarySheet.getRange(5, 8, dataRowCount, 1);
      cpcIndexRange.setNumberFormat('#,##0"%"');
      cpcIndexRange.setHorizontalAlignment('center');
      
      // Cost (Columns I & J) - Currency format without decimals, center aligned
      const costRange = summarySheet.getRange(5, 9, dataRowCount, 2);
      costRange.setNumberFormat('$#,##0');
      costRange.setHorizontalAlignment('center');
      
      // Cost Index (Column K) - Percentage format, center aligned
      const costIndexRange = summarySheet.getRange(5, 11, dataRowCount, 1);
      costIndexRange.setNumberFormat('#,##0"%"');
      costIndexRange.setHorizontalAlignment('center');
      
      // Conversion Value (Columns L & M) - Currency format without decimals, center aligned
      const conversionValueRange = summarySheet.getRange(5, 12, dataRowCount, 2);
      conversionValueRange.setNumberFormat('$#,##0');
      conversionValueRange.setHorizontalAlignment('center');
      
      // Conversion Value Index (Column N) - Percentage format, center aligned
      const conversionValueIndexRange = summarySheet.getRange(5, 14, dataRowCount, 1);
      conversionValueIndexRange.setNumberFormat('#,##0"%"');
      conversionValueIndexRange.setHorizontalAlignment('center');
      
      // ROAS (Columns O & P) - Number format with 2 decimals, center aligned
      const roasRange = summarySheet.getRange(5, 15, dataRowCount, 2);
      roasRange.setNumberFormat('#,##0.00');
      roasRange.setHorizontalAlignment('center');
      
      // ROAS Index (Column Q) - Percentage format, center aligned
      const roasIndexRange = summarySheet.getRange(5, 17, dataRowCount, 1);
      roasIndexRange.setNumberFormat('#,##0"%"');
      roasIndexRange.setHorizontalAlignment('center');
      
      // Apply conditional formatting to Index YoY columns with very light gradient colors
      const indexColumns = [5, 8, 11, 14, 17]; // E, H, K, N, Q
      
      indexColumns.forEach(colIndex => {
        const indexRange = summarySheet.getRange(5, colIndex, dataRowCount, 1);
        
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
    
    // Set equal column widths for year comparison triplets
    summarySheet.setColumnWidth(1, 80);   // Week Number
    summarySheet.setColumnWidth(2, 100);  // Week Start
    
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
    
    // Set equal widths for Conversion Value columns (L, M, N)
    summarySheet.setColumnWidth(12, 110); // Conv Value 2025
    summarySheet.setColumnWidth(13, 110); // Conv Value 2024
    summarySheet.setColumnWidth(14, 80);  // Conv Value Index YoY
    
    // Set equal widths for ROAS columns (O, P, Q)
    summarySheet.setColumnWidth(15, 80);  // ROAS 2025
    summarySheet.setColumnWidth(16, 80);  // ROAS 2024
    summarySheet.setColumnWidth(17, 80);  // ROAS Index YoY
    
    Logger.log(`✅ Year-over-year weekly summary created with ${maxWeeks} weeks - filter changes automatically update data!`);
    
  } catch (error) {
    Logger.log('❌ Error creating weekly summary: ' + error.toString());
  }
} 