/**
 * =============================================================================
 * GOOGLE ADS UNIVERSAL PRODUCT TYPE OPTIMIZER - AI POWERED v2.0
 * =============================================================================
 * 
 * This script optimizes product types for ANY PRODUCT CATEGORY with AI-powered categorization
 * Configurable for any type of products you sell - works universally across all industries
 * 
 * v2.0 UPDATE: FULLY AUTOMATED GOOGLE ADS EXTRACTION!
 * - No longer requires external spreadsheet with product data
 * - API key configured directly in script (no spreadsheet setup needed)
 * - Automatically extracts product IDs, titles, and current types from your Google Ads account
 * - Uses last 7 days of data automatically (no manual date configuration needed)
 * - Single-run process - just run main() and you're done!
 * - Analyzes all products from Shopping and Performance Max campaigns
 * - Creates results spreadsheet automatically
 * 
 * Creates 3-level hierarchy: "[Your Category] > [Subcategory] > [Specific Type]"
 * Example: "Inflatable > Water Toys > Pool Floats" or "Clothing > Shirts > T-Shirts"
 * 
 * This script runs ONLY in Google Ads (not Google Sheets Apps Script)
 * 
 * SETUP:
 * 1. Go to Google Ads > Tools & Settings > Bulk Actions > Scripts
 * 2. Create new script and paste this code
 * 3. Update configuration below (especially PRODUCT_CATEGORY)
 * 4. Run testSetup() first, then main()
 */

// =============================================================================
// CONFIGURATION - UPDATE THESE BEFORE RUNNING
// =============================================================================

// Product category settings - CUSTOMIZE THIS FOR YOUR PRODUCTS
const PRODUCT_CATEGORY = "Inflatable";  // Examples: "Hardware", "Fasteners", "Metal Products", "Industrial", "Construction", etc.
// Note: PRODUCT_CATEGORY will always be the first level in the 3-level hierarchy
//
// HIERARCHY EXAMPLES:
// Hardware: "Hardware > Fasteners > Screws"
// Electronics: "Electronics > Computers > Laptops" 
// Clothing: "Clothing > Shirts > T-Shirts"
// Home & Garden: "Home & Garden > Furniture > Chairs"

// OpenAI API key - ADD YOUR API KEY HERE
const OPENAI_API_KEY = "sk-your-openai-api-key-here";  // Replace with your actual OpenAI API key

// Automatically uses last 7 days of data (no manual date configuration needed)
const DAYS_BACK = 7;  // Number of days to look back for product data

// Results spreadsheet settings (will be auto-created)
const RESULTS_SHEET_NAME = `${PRODUCT_CATEGORY} Product Type Optimization Results`;

// Processing settings
const BATCH_SIZE = 10;        // Increased batch size for efficiency
const DELAY_MS = 200;         // Reduced delay from 1000ms to 200ms
const MIN_IMPRESSIONS = 1;    // Minimum impressions to include a product

// =============================================================================
// MAIN FUNCTION
// =============================================================================

function main() {
  const startTime = Date.now();
  Logger.log(`🚀 Starting Google Ads ${PRODUCT_CATEGORY} Product Type Optimization v2.0`);
  Logger.log(`📊 Extracting data directly from Google Ads account`);
  
  try {
    // Step 1: Validate API key
    if (!OPENAI_API_KEY || OPENAI_API_KEY === "sk-your-openai-api-key-here") {
      Logger.log("❌ Please set your OpenAI API key in the OPENAI_API_KEY configuration variable");
      Logger.log("💡 Get your API key from: https://platform.openai.com/api-keys");
      return;
    }
    
    // Step 2: Extract products from Google Ads
    Logger.log("📥 Extracting product data from Google Ads campaigns...");
    const products = extractProductsFromGoogleAds();
    
    if (products.length === 0) {
      Logger.log(`❌ No ${PRODUCT_CATEGORY.toLowerCase()} products found in Shopping/Performance Max campaigns`);
      Logger.log("💡 Make sure you have Shopping or Performance Max campaigns with product data in the specified date range");
      return;
    }
    
    Logger.log(`📦 Found ${products.length} ${PRODUCT_CATEGORY.toLowerCase()} products to optimize`);
    
    // Step 3: Create results spreadsheet
    Logger.log("📋 Creating results spreadsheet...");
    const spreadsheet = createResultsSpreadsheet();
    const resultsSheet = spreadsheet.getActiveSheet();
    
    // Step 4: Set up spreadsheet with data
    setupResultsSheet(resultsSheet, products);
    
    Logger.log(`🔗 Results spreadsheet: ${spreadsheet.getUrl()}`);
    
    // Step 5: Process products with AI optimization
    Logger.log("🤖 Starting AI optimization process...");
    let successful = 0;
    let failed = 0;
    
    for (let i = 0; i < products.length; i += BATCH_SIZE) {
      const elapsedTime = Date.now() - startTime;
      const batch = products.slice(i, i + BATCH_SIZE);
      const batchNum = Math.floor(i / BATCH_SIZE) + 1;
      const totalBatches = Math.ceil(products.length / BATCH_SIZE);
      
      Logger.log(`🔄 Batch ${batchNum}/${totalBatches} (${batch.length} products) - ${Math.round(elapsedTime/60000)}min elapsed`);
      
      for (const product of batch) {
        try {
          const result = optimizeProductType(product.title, product.currentType, OPENAI_API_KEY, product.typeL1, product.typeL2, product.typeL3);
          writeOptimizedResult(resultsSheet, product.row, result.productType);
          successful++;
          Logger.log(`✅ Row ${product.row}: ${result.productType}`);
        } catch (error) {
          failed++;
          writeOptimizedResult(resultsSheet, product.row, `Error: ${error.message}`);
          Logger.log(`❌ Row ${product.row}: ${error.message}`);
        }
        
        Utilities.sleep(DELAY_MS);
      }
      
             if (i + BATCH_SIZE < products.length) {
         Utilities.sleep(500);
       }
     }
     
     const totalTime = Math.round((Date.now() - startTime) / 60000);
     Logger.log(`🎉 Optimization complete in ${totalTime} minutes! Success: ${successful}, Failed: ${failed}`);
     Logger.log(`📊 Results saved in results spreadsheet: ${spreadsheet.getUrl()}`);
     Logger.log("✅ Script finished successfully - safe to close");
     
   } catch (error) {
     const totalTime = Math.round((Date.now() - startTime) / 60000);
     Logger.log(`💥 Fatal error after ${totalTime} minutes: ${error.message}`);
     Logger.log("Script execution completed with errors. Check logs above for details.");
     return;
   }
   
   // Explicit completion
   Logger.log("✅ Script execution fully completed");
 }

// =============================================================================
// GOOGLE ADS DATA EXTRACTION
// =============================================================================

function getDateRange() {
  const today = new Date();
  const endDate = new Date(today);
  const startDate = new Date(today);
  
  // Set end date to yesterday (Google Ads data is typically 1 day behind)
  endDate.setDate(endDate.getDate() - 1);
  
  // Set start date to DAYS_BACK days before end date
  startDate.setDate(endDate.getDate() - (DAYS_BACK - 1));
  
  // Format dates as YYYY-MM-DD
  const formatDate = (date) => {
    return date.getFullYear() + '-' + 
           String(date.getMonth() + 1).padStart(2, '0') + '-' + 
           String(date.getDate()).padStart(2, '0');
  };
  
  return {
    startDate: formatDate(startDate),
    endDate: formatDate(endDate)
  };
}

function extractProductsFromGoogleAds() {
  Logger.log("🔍 Querying Google Ads for product data...");
  
  // Get dynamic date range
  const dateRange = getDateRange();
  const START_DATE = dateRange.startDate;
  const END_DATE = dateRange.endDate;
  
  const query = `
    SELECT
      segments.product_item_id,
      segments.product_title,
      segments.product_type_l1,
      segments.product_type_l2,
      segments.product_type_l3,
      metrics.impressions,
      metrics.clicks,
      campaign.name,
      campaign.advertising_channel_type
    FROM shopping_performance_view
    WHERE segments.date BETWEEN '${START_DATE}' AND '${END_DATE}'
      AND campaign.advertising_channel_type IN ("SHOPPING", "PERFORMANCE_MAX")
      AND metrics.impressions >= ${MIN_IMPRESSIONS}
    ORDER BY metrics.impressions DESC
  `;
  
  Logger.log(`📅 Date range: ${START_DATE} to ${END_DATE} (last ${DAYS_BACK} days)`);
  Logger.log(`📊 Minimum impressions: ${MIN_IMPRESSIONS}`);
  
  const rows = AdsApp.search(query);
  const products = [];
  const uniqueProducts = new Map(); // To deduplicate by product ID
  
  Logger.log("📥 Processing product data from Google Ads...");
  
  while (rows.hasNext()) {
    try {
      const row = rows.next();
      
      const productId = row.segments.productItemId || '';
      const title = row.segments.productTitle || '';
      const typeL1 = row.segments.productTypeL1 || '';
      const typeL2 = row.segments.productTypeL2 || '';
      const typeL3 = row.segments.productTypeL3 || '';
      const impressions = row.metrics.impressions || 0;
      const campaignName = row.campaign.name || '';
      const channelType = row.campaign.advertisingChannelType || '';
      
      if (productId && title) {
        // Build current type hierarchy
        const typeParts = [typeL1, typeL2, typeL3].filter(part => part && part.trim());
        const currentType = typeParts.join(' > ') || 'No Type';
        
        // Use product ID as key for deduplication
        if (!uniqueProducts.has(productId)) {
          uniqueProducts.set(productId, {
            id: productId,
            title: title,
            currentType: currentType,
            typeL1: typeL1,
            typeL2: typeL2,
            typeL3: typeL3,
            impressions: impressions,
            campaignName: campaignName,
            channelType: channelType
          });
        } else {
          // If duplicate, keep the one with higher impressions
          const existing = uniqueProducts.get(productId);
          if (impressions > existing.impressions) {
            uniqueProducts.set(productId, {
              id: productId,
              title: title,
              currentType: currentType,
              typeL1: typeL1,
              typeL2: typeL2,
              typeL3: typeL3,
              impressions: impressions,
              campaignName: campaignName,
              channelType: channelType
            });
          }
        }
      }
    } catch (error) {
      Logger.log(`⚠️ Error processing row: ${error.message}`);
    }
  }
  
  // Convert map to array and add row numbers
  const productsArray = Array.from(uniqueProducts.values());
  productsArray.forEach((product, index) => {
    product.row = index + 2; // +2 for header row
  });
  
  Logger.log(`✅ Extracted ${productsArray.length} unique products from Google Ads`);
  
  // Log sample data
  if (productsArray.length > 0) {
    const sample = productsArray[0];
    Logger.log(`📊 Sample product: "${sample.title}" (ID: ${sample.id}, Type: ${sample.currentType})`);
  }
  
  return productsArray;
}

// =============================================================================
// SPREADSHEET FUNCTIONS
// =============================================================================

function createResultsSpreadsheet() {
  const timestamp = Utilities.formatDate(new Date(), AdsApp.currentAccount().getTimeZone(), 'yyyy-MM-dd HH:mm');
  const sheetName = `${RESULTS_SHEET_NAME} - ${timestamp}`;
  
  const spreadsheet = SpreadsheetApp.create(sheetName);
  Logger.log(`✅ Created results spreadsheet: ${spreadsheet.getUrl()}`);
  
  return spreadsheet;
}



function setupResultsSheet(sheet, products) {
  // Rename the default sheet
  sheet.setName("Results");
  
  // Set up headers
  const headers = [
    'Product ID',
    'Product Title',
    'Current Product Type',
    'Optimized Product Type'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Format headers
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('#ffffff');
  
  // Add product data
  const data = products.map(product => [
    product.id,
    product.title,
    product.currentType,
    '' // Optimized type - to be filled
  ]);
  
  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, headers.length).setValues(data);
  }
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, headers.length);
  
  Logger.log(`✅ Set up results sheet with ${products.length} products`);
}



function writeOptimizedResult(sheet, row, optimizedType) {
  sheet.getRange(row, 4).setValue(optimizedType); // Column 4 is 'Optimized Product Type'
}

// =============================================================================
// AI OPTIMIZATION (UNCHANGED FROM ORIGINAL)
// =============================================================================

function optimizeProductType(title, currentType, apiKey, typeL1 = '', typeL2 = '', typeL3 = '') {
  const prompt = `You are a Google Merchant Center expert. Create a 3-level product type hierarchy for ${PRODUCT_CATEGORY.toUpperCase()} PRODUCTS.

HIERARCHY STRUCTURE:
Level 1: "${PRODUCT_CATEGORY}" (FIXED - always use this exactly)
Level 2: Subcategory or attribute (identify the most important categorizing feature)
Level 3: Specific product type (you generate this)

RULES:
- Exactly 3 levels: "${PRODUCT_CATEGORY} > Subcategory > Specific Type"
- Level 1 is ALWAYS "${PRODUCT_CATEGORY}" exactly as specified
- Level 2: Most important subcategory, material, style, or attribute that helps categorize the product
- Level 3: Specific product type within that category
- Be logical and intuitive for shoppers
- No brand names
- Be specific but not overly granular

EXAMPLES FOR ${PRODUCT_CATEGORY}:
- "${PRODUCT_CATEGORY} > Water Toys > Pool Floats"
- "${PRODUCT_CATEGORY} > Party Supplies > Bounce Houses"
- "${PRODUCT_CATEGORY} > Outdoor Recreation > Water Slides"
- "${PRODUCT_CATEGORY} > Sports Equipment > Obstacle Courses"
- "${PRODUCT_CATEGORY} > Furniture > Chairs"

RETURN JSON:
{"productType": "${PRODUCT_CATEGORY} > Subcategory > Specific Type"}`;

  // Build context string with existing type structure
  let contextString = `Title: ${title}`;
  
  if (typeL1 || typeL2 || typeL3) {
    contextString += `\nExisting Type Structure:`;
    if (typeL1) contextString += `\n  Level 1: ${typeL1}`;
    if (typeL2) contextString += `\n  Level 2: ${typeL2}`;
    if (typeL3) contextString += `\n  Level 3: ${typeL3}`;
  }
  
  if (currentType && currentType !== [typeL1, typeL2, typeL3].filter(t => t).join(' > ')) {
    contextString += `\nCurrent Type: ${currentType}`;
  }
  
  contextString += `\n\nNote: This is a ${PRODUCT_CATEGORY.toLowerCase()} product. Please identify the most appropriate subcategory for Level 2.`;

  const payload = {
    model: "gpt-4o-mini",
    messages: [
      { role: "system", content: prompt },
      { role: "user", content: contextString }
    ],
    max_tokens: 100,
    temperature: 0.2
  };
  
  const options = {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${apiKey}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload)
  };
  
  const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', options);
  const data = JSON.parse(response.getContentText());
  
  if (data.error) {
    throw new Error(`API Error: ${data.error.message}`);
  }
  
  return parseResponse(data.choices[0].message.content);
}

function parseResponse(text) {
  try {
    // Try JSON parsing first
    const jsonMatch = text.match(/\{[^}]*\}/);
    if (jsonMatch) {
      const parsed = JSON.parse(jsonMatch[0]);
      return {
        productType: parsed.productType || 'Parse Error'
      };
    }
    
    // Fallback parsing
    const lines = text.split('\n').map(l => l.trim()).filter(l => l);
    let productType = 'Parse Error';
    
    for (const line of lines) {
      if (line.includes('>')) {
        productType = line.replace(/^[^:]*:?\s*/, '');
        break;
      }
    }
    
    return { productType };
    
  } catch (error) {
    return {
      productType: 'Parse Error'
    };
  }
}

// =============================================================================
// UTILITY FUNCTIONS
// =============================================================================



// =============================================================================
// TEST FUNCTIONS
// =============================================================================

function testSetup() {
  Logger.log("🧪 Testing Google Ads data extraction...");
  
  try {
    // Get dynamic date range
    const dateRange = getDateRange();
    const START_DATE = dateRange.startDate;
    const END_DATE = dateRange.endDate;
    
    // Test Google Ads query with limit
    const testQuery = `
      SELECT
        segments.product_item_id,
        segments.product_title,
        segments.product_type_l1,
        segments.product_type_l2,
        segments.product_type_l3,
        metrics.impressions,
        campaign.name,
        campaign.advertising_channel_type
      FROM shopping_performance_view
      WHERE segments.date BETWEEN '${START_DATE}' AND '${END_DATE}'
        AND campaign.advertising_channel_type IN ("SHOPPING", "PERFORMANCE_MAX")
        AND metrics.impressions >= ${MIN_IMPRESSIONS}
      ORDER BY metrics.impressions DESC
      LIMIT 5
    `;
    
    Logger.log(`🔍 Testing Google Ads query for date range: ${START_DATE} to ${END_DATE} (last ${DAYS_BACK} days)...`);
    const rows = AdsApp.search(testQuery);
    
    let count = 0;
    while (rows.hasNext() && count < 5) {
      const row = rows.next();
      count++;
      
      Logger.log(`✅ Product ${count}:`);
      Logger.log(`   ID: ${row.segments.productItemId || 'N/A'}`);
      Logger.log(`   Title: ${row.segments.productTitle || 'N/A'}`);
      Logger.log(`   Type L1: ${row.segments.productTypeL1 || 'N/A'}`);
      Logger.log(`   Type L2: ${row.segments.productTypeL2 || 'N/A'}`);
      Logger.log(`   Type L3: ${row.segments.productTypeL3 || 'N/A'}`);
      Logger.log(`   Impressions: ${row.metrics.impressions || 0}`);
      Logger.log(`   Campaign: ${row.campaign.name || 'N/A'}`);
      Logger.log(`   Channel Type: ${row.campaign.advertisingChannelType || 'N/A'}`);
    }
    
    if (count === 0) {
      Logger.log("❌ No products found. Check:");
      Logger.log("   - Date range contains data");
      Logger.log("   - You have Shopping or Performance Max campaigns");
      Logger.log("   - Products have impressions >= " + MIN_IMPRESSIONS);
    } else {
      Logger.log(`🎉 Test successful! Found ${count} sample products`);
      Logger.log("✅ Ready to run main() function");
    }
    
  } catch (error) {
    Logger.log(`❌ Test failed: ${error.message}`);
  }
}

function testSingleOptimization() {
  Logger.log(`🧪 Testing single ${PRODUCT_CATEGORY.toLowerCase()} product optimization...`);
  
  try {
    // Check if API key is configured
    if (!OPENAI_API_KEY || OPENAI_API_KEY === "sk-your-openai-api-key-here") {
      Logger.log("❌ Please set your OpenAI API key in the OPENAI_API_KEY configuration variable");
      Logger.log("💡 Get your API key from: https://platform.openai.com/api-keys");
      return;
    }
    
    // Use category-appropriate test product
    const testTitle = PRODUCT_CATEGORY === "Inflatable" ? "Large Inflatable Water Slide with Pool" : "Sample Product for Testing";
    const testCurrentType = PRODUCT_CATEGORY === "Inflatable" ? "Water Toys" : "General Category";
    
    const result = optimizeProductType(
      testTitle, 
      testCurrentType,
      OPENAI_API_KEY,
      testCurrentType,
      '',
      ''
    );
    
    Logger.log(`✅ Test successful! Result: ${result.productType}`);
    
  } catch (error) {
    Logger.log(`❌ Test failed: ${error.message}`);
  }
}

// =============================================================================
// INSTRUCTIONS
// =============================================================================

/*
HOW TO USE - NEW v2.0 DIRECT GOOGLE ADS EXTRACTION:

1. UPDATE CONFIGURATION:
   - Set your PRODUCT_CATEGORY (e.g., "Electronics", "Clothing", "Home & Garden", etc.)
   - Set your OPENAI_API_KEY (get from https://platform.openai.com/api-keys)
   - Optionally adjust DAYS_BACK (defaults to 7 days of recent data)
   - This will be the first level in your 3-level hierarchy

2. NO SPREADSHEET SETUP REQUIRED:
   - The script now extracts data directly from your Google Ads account
   - Looks at Shopping and Performance Max campaigns
   - Extracts product IDs, titles, and current types
   - No need for manual data export or Google Sheets setup

3. SIMPLIFIED WORKFLOW:
   - Just run: main()
   - That's it! The script will:
     * Extract product data from Google Ads
     * Create results spreadsheet automatically
     * Process all products with AI optimization
     * Save optimized product types to the spreadsheet

4. DATA EXTRACTION DETAILS:
   - Queries shopping_performance_view
   - Automatically uses last 7 days of data (ending yesterday)
   - Includes products from Shopping and Performance Max campaigns
   - Filters by minimum impressions threshold
   - Deduplicates products by ID (keeps highest impression version)
   - Extracts: Product ID, Title, Type L1/L2/L3, Impressions, Campaign

5. RUN FUNCTIONS:
   - First: testSetup() - validates Google Ads data access
   - Optional: testSingleOptimization() - tests API key and AI optimization
   - Then: main() - runs complete process (data extraction + AI optimization)

6. MONITOR LOGS:
   - Check Google Ads Scripts logs for progress
   - Look for ✅ (success) and ❌ (error) indicators
   - Results will be saved in the auto-created spreadsheet

ADVANTAGES OF v2.0:
- No manual data export required
- Always uses fresh, current product data (last 7 days)
- Single-run process - no multi-step workflow
- API key configured directly in script (more secure)
- Automatically handles Shopping and Performance Max campaigns
- Deduplicates products across campaigns
- Shows impression data for prioritization
- Creates clean, organized results spreadsheet

TROUBLESHOOTING:
- "Please set your OpenAI API key" = Update OPENAI_API_KEY in script configuration
- "No products found" = Check if you have Shopping/Performance Max campaigns with recent data (last 7 days)
- "API Error" = Check OpenAI API key validity and credits
- "Query failed" = Check if you have Shopping/Performance Max campaigns with product data
- "No recent data" = Increase DAYS_BACK if your campaigns have lower activity

SECURITY:
- API key stored directly in script configuration
- Easy to update and manage
- No external spreadsheet dependencies for API key

AUTOMATIC DATE RANGE:
- Script automatically uses last 7 days of data (ending yesterday)
- Ensures fresh, recent product performance data
- Adjust DAYS_BACK if you need more/fewer days of data
- Google Ads data is typically 1 day behind, so "yesterday" is the most recent available
*/ 