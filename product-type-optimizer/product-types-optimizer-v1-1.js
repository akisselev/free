/**
 * =============================================================================
 * GOOGLE ADS UNIVERSAL PRODUCT TYPE OPTIMIZER - AI POWERED v1.1
 * =============================================================================
 * 
 * This script optimizes product types for ANY PRODUCT CATEGORY with AI-powered categorization
 * Configurable for any type of products you sell - works universally across all industries
 * 
 * v1.1 UPDATE: Fixed incorrect categorization in product type level 2 (steel) - now uses 
 * intelligent subcategorization based on actual product attributes instead of forcing metal types
 * 
 * Creates 3-level hierarchy: "[Your Category] > [Subcategory] > [Specific Type]"
 * Example: "Inflatable > Water Toys > Pool Floats" or "Clothing > Shirts > T-Shirts"
 * 
 * This script runs ONLY in Google Ads (not Google Sheets Apps Script)
 * 
 * SETUP INSTRUCTIONS VIDEO: https://youtu.be/4t1cY1U8ZYY
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

const SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/YOUR_SPREADSHEET_ID/edit#gid=0";

// Product category settings - CUSTOMIZE THIS FOR YOUR PRODUCTS
const PRODUCT_CATEGORY = "Inflatable";  // Examples: "Hardware", "Fasteners", "Metal Products", "Industrial", "Construction", etc.
// Note: PRODUCT_CATEGORY will always be the first level in the 3-level hierarchy
//
// HIERARCHY EXAMPLES:
// Hardware: "Hardware > Fasteners > Screws"
// Electronics: "Electronics > Computers > Laptops" 
// Clothing: "Clothing > Shirts > T-Shirts"
// Home & Garden: "Home & Garden > Furniture > Chairs"

// Sheet settings
const SOURCE_SHEET_NAME = "products";        // Source sheet with original data
const TARGET_SHEET_NAME = "Product Types";   // New sheet for processing
const HEADER_ROW = 1;

// Named range for API key (create this in your spreadsheet)
const API_KEY_NAMED_RANGE = "openaiApiKey";

// Column mapping (will be dynamically detected)
const COLUMNS = {
  ID: null,                // Will be found by header name "ID"
  PRODUCT_TITLE: null,     // Will be found by header name
  PRODUCT_TYPE: null,      // Will be found by header name "product type"
  OPTIMIZED_TYPE: null     // Will be created in the new sheet
};

// Processing settings
const BATCH_SIZE = 10;        // Increased batch size for efficiency
const DELAY_MS = 200;         // Reduced delay from 1000ms to 200ms

// =============================================================================
// MAIN FUNCTION
// =============================================================================

function main() {
  const startTime = Date.now();
  Logger.log(`🚀 Starting Google Ads ${PRODUCT_CATEGORY} Product Type Optimization`);
  
  try {
    // Step 1: Get spreadsheet and validate setup
    const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
    
    if (!validateSetup(spreadsheet)) {
      Logger.log("❌ Setup validation failed - check configuration");
      return;
    }
    
    // Step 2: Get API key from named range
    const apiKey = getApiKey(spreadsheet);
    if (!apiKey) {
      Logger.log("❌ Failed to get API key from named range");
      return;
    }
    
    // Step 3: Get source sheet
    const sourceSheet = spreadsheet.getSheetByName(SOURCE_SHEET_NAME);
    
    if (!sourceSheet) {
      Logger.log(`❌ Source sheet '${SOURCE_SHEET_NAME}' not found`);
      return;
    }
    
    // Step 4: Detect columns in source sheet
    if (!detectColumns(sourceSheet)) {
      Logger.log("❌ Could not detect required columns in source sheet - check your spreadsheet headers");
      return;
    }
    
    // Step 5: Create new "Product Types" sheet with extracted data
    const targetSheet = createProductTypesSheet(spreadsheet, sourceSheet);
    
    // Step 6: Get data from the new sheet
    const data = targetSheet.getDataRange().getValues();
    const products = extractProducts(data);
    
    Logger.log(`📦 Found ${products.length} ${PRODUCT_CATEGORY.toLowerCase()} products to optimize`);
    
    if (products.length === 0) {
      Logger.log(`❌ No valid ${PRODUCT_CATEGORY.toLowerCase()} products found in '${TARGET_SHEET_NAME}' sheet`);
      return;
    }
    
    // Step 7: Process products
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
          const result = optimizeProductType(product.title, product.currentType, apiKey, product.typeL1, product.typeL2, product.typeL3);
          writeToSheet(targetSheet, product.row, result.productType);
          successful++;
          Logger.log(`✅ Row ${product.row}: ${result.productType}`);
        } catch (error) {
          failed++;
          writeToSheet(targetSheet, product.row, `Error: ${error.message}`);
          Logger.log(`❌ Row ${product.row}: ${error.message}`);
        }
        
        // Minimal delay between products (reduced from 1000ms to 200ms)
        Utilities.sleep(DELAY_MS);
      }
      
      // Only brief delay between batches (reduced from 3000ms to 500ms)
      if (i + BATCH_SIZE < products.length) {
        Utilities.sleep(500);
      }
    }
    
    const totalTime = Math.round((Date.now() - startTime) / 60000);
    Logger.log(`🎉 Optimization complete in ${totalTime} minutes! Success: ${successful}, Failed: ${failed}`);
    Logger.log(`📊 Results saved in '${TARGET_SHEET_NAME}' sheet`);
    Logger.log("✅ Script finished successfully - safe to close");
    
  } catch (error) {
    const totalTime = Math.round((Date.now() - startTime) / 60000);
    Logger.log(`💥 Fatal error after ${totalTime} minutes: ${error.message}`);
    Logger.log("Script execution completed with errors. Check logs above for details.");
    return; // Explicit exit
  }
  
  // Explicit completion
  Logger.log("✅ Script execution fully completed");
}

// =============================================================================
// CORE FUNCTIONS
// =============================================================================

function createProductTypesSheet(spreadsheet, sourceSheet) {
  Logger.log("📋 Creating 'Product Types' sheet...");
  
  // Check if sheet already exists
  let targetSheet = spreadsheet.getSheetByName(TARGET_SHEET_NAME);
  if (targetSheet) {
    Logger.log(`⚠️ Sheet '${TARGET_SHEET_NAME}' already exists - clearing content`);
    targetSheet.clear();
  } else {
    targetSheet = spreadsheet.insertSheet(TARGET_SHEET_NAME);
    Logger.log(`✅ Created new sheet '${TARGET_SHEET_NAME}'`);
  }
  
  // Set up columns in the new sheet
  const newColumns = {
    ID: 1,
    PRODUCT_TITLE: 2,
    PRODUCT_TYPE: 3,
    OPTIMIZED_TYPE: 4
  };
  
  // Set headers
  targetSheet.getRange(1, newColumns.ID).setValue("ID");
  targetSheet.getRange(1, newColumns.PRODUCT_TITLE).setValue("Product Title");
  targetSheet.getRange(1, newColumns.PRODUCT_TYPE).setValue("Product Type");
  targetSheet.getRange(1, newColumns.OPTIMIZED_TYPE).setValue("Optimized Type");
  
  // Copy data from source sheet
  const sourceData = sourceSheet.getDataRange().getValues();
  const extractedData = [];
  
  for (let i = HEADER_ROW; i < sourceData.length; i++) {
    const row = sourceData[i];
    const id = getString(row[COLUMNS.ID - 1]);
    const title = getString(row[COLUMNS.PRODUCT_TITLE - 1]);
    const productType = getString(row[COLUMNS.PRODUCT_TYPE - 1]);
    
    if (id && title && productType) {
      extractedData.push([id, title, productType, '']); // Empty optimized type initially
    }
  }
  
  if (extractedData.length > 0) {
    targetSheet.getRange(2, 1, extractedData.length, 4).setValues(extractedData);
    Logger.log(`✅ Copied ${extractedData.length} products to '${TARGET_SHEET_NAME}' sheet`);
  } else {
    Logger.log("❌ No valid data found to copy");
  }
  
  // Update COLUMNS object for the new sheet
  COLUMNS.ID = newColumns.ID;
  COLUMNS.PRODUCT_TITLE = newColumns.PRODUCT_TITLE;
  COLUMNS.PRODUCT_TYPE = newColumns.PRODUCT_TYPE;
  COLUMNS.OPTIMIZED_TYPE = newColumns.OPTIMIZED_TYPE;
  
  return targetSheet;
}

function detectColumns(sheet) {
  const headers = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  Logger.log(`🔍 Detecting columns from headers: ${headers.join(', ')}`);
  
  // Find columns by header names (case-insensitive)
  for (let i = 0; i < headers.length; i++) {
    const header = String(headers[i]).toLowerCase().trim();
    const originalHeader = String(headers[i]).trim();
    
    // More precise ID detection to avoid false matches
    if (header === 'id' || header === 'product id' || header === 'sku' || header === 'product_id') {
      COLUMNS.ID = i + 1;
      Logger.log(`✅ Found ID column at position ${i + 1}: "${originalHeader}"`);
    } else if (header === 'product title' || header === 'title' || header === 'name' || header === 'product name' || header === 'product_title') {
      COLUMNS.PRODUCT_TITLE = i + 1;
      Logger.log(`✅ Found Title column at position ${i + 1}: "${originalHeader}"`);
    } else if (header === 'product type' || header === 'type' || header === 'category' || header === 'product_type') {
      COLUMNS.PRODUCT_TYPE = i + 1;
      Logger.log(`✅ Found Product Type column at position ${i + 1}: "${originalHeader}"`);
    }
  }
  
  // Set optimized type column next to product type
  if (COLUMNS.PRODUCT_TYPE) {
    COLUMNS.OPTIMIZED_TYPE = COLUMNS.PRODUCT_TYPE + 1;
  }
  
  Logger.log(`📋 Final column mapping: ID=${COLUMNS.ID}, Title=${COLUMNS.PRODUCT_TITLE}, Product Type=${COLUMNS.PRODUCT_TYPE}, Optimized Type=${COLUMNS.OPTIMIZED_TYPE}`);
  
  const hasRequiredColumns = COLUMNS.ID && COLUMNS.PRODUCT_TITLE && COLUMNS.PRODUCT_TYPE;
  
  if (!hasRequiredColumns) {
    Logger.log(`❌ Missing required columns:`);
    if (!COLUMNS.ID) Logger.log(`   - ID column not found (looking for: "id", "product id", "sku")`);
    if (!COLUMNS.PRODUCT_TITLE) Logger.log(`   - Title column not found (looking for: "product title", "title", "name")`);
    if (!COLUMNS.PRODUCT_TYPE) Logger.log(`   - Product Type column not found (looking for: "product type", "type", "category")`);
    
    // Show what headers we actually found
    Logger.log(`📋 Available headers: ${headers.join(', ')}`);
  }
  
  return hasRequiredColumns;
}

function debugSourceData() {
  Logger.log("🔍 Debugging source data...");
  
  try {
    const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
    const sourceSheet = spreadsheet.getSheetByName(SOURCE_SHEET_NAME);
    
    if (!sourceSheet) {
      Logger.log(`❌ Source sheet '${SOURCE_SHEET_NAME}' not found`);
      return;
    }
    
    // Show sheet info
    Logger.log(`📊 Sheet '${SOURCE_SHEET_NAME}' has ${sourceSheet.getLastRow()} rows and ${sourceSheet.getLastColumn()} columns`);
    
    // Show headers
    const headers = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
    Logger.log(`📋 Headers: ${headers.join(' | ')}`);
    
    // Try to detect columns
    if (!detectColumns(sourceSheet)) {
      Logger.log("❌ Column detection failed");
      return;
    }
    
    // Show sample data
    const sampleData = sourceSheet.getRange(1, 1, Math.min(6, sourceSheet.getLastRow()), sourceSheet.getLastColumn()).getValues();
    Logger.log("📋 Sample data (first 5 rows):");
    
    for (let i = 0; i < sampleData.length; i++) {
      const row = sampleData[i];
      Logger.log(`   Row ${i + 1}: ${row.join(' | ')}`);
    }
    
    // Check specific columns
    if (COLUMNS.ID && COLUMNS.PRODUCT_TITLE && COLUMNS.PRODUCT_TYPE) {
      Logger.log("🔍 Checking specific column data:");
      
      for (let i = 1; i < Math.min(6, sourceSheet.getLastRow()); i++) {
        const row = sourceSheet.getRange(i + 1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
        const id = getString(row[COLUMNS.ID - 1]);
        const title = getString(row[COLUMNS.PRODUCT_TITLE - 1]);
        const productType = getString(row[COLUMNS.PRODUCT_TYPE - 1]);
        
        Logger.log(`   Row ${i + 1}: ID="${id}", Title="${title}", Type="${productType}"`);
        
        if (id && title && productType) {
          Logger.log(`   ✅ Row ${i + 1} has all required data`);
        } else {
          Logger.log(`   ❌ Row ${i + 1} missing data - ID:${!!id}, Title:${!!title}, Type:${!!productType}`);
        }
      }
    }
    
  } catch (error) {
    Logger.log(`❌ Debug failed: ${error.message}`);
  }
}

function testDataExtraction() {
  Logger.log("🧪 Testing data extraction...");
  
  try {
    const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
    const sourceSheet = spreadsheet.getSheetByName(SOURCE_SHEET_NAME);
    
    if (!sourceSheet) {
      Logger.log(`❌ Source sheet '${SOURCE_SHEET_NAME}' not found`);
      return;
    }
    
    if (!detectColumns(sourceSheet)) {
      Logger.log("❌ Could not detect required columns");
      return;
    }
    
    const sourceData = sourceSheet.getDataRange().getValues();
    const extractedData = [];
    let validRows = 0;
    let invalidRows = 0;
    
    Logger.log(`📊 Processing ${sourceData.length - 1} data rows...`);
    
    for (let i = HEADER_ROW; i < sourceData.length; i++) {
      const row = sourceData[i];
      const id = getString(row[COLUMNS.ID - 1]);
      const title = getString(row[COLUMNS.PRODUCT_TITLE - 1]);
      const productType = getString(row[COLUMNS.PRODUCT_TYPE - 1]);
      
      if (id && title && productType) {
        extractedData.push([id, title, productType, '']);
        validRows++;
        if (validRows <= 3) {
          Logger.log(`✅ Valid row ${i + 1}: ID="${id}", Title="${title}", Type="${productType}"`);
        }
      } else {
        invalidRows++;
        if (invalidRows <= 3) {
          Logger.log(`❌ Invalid row ${i + 1}: ID="${id}", Title="${title}", Type="${productType}"`);
        }
      }
    }
    
    Logger.log(`📊 Results: ${validRows} valid rows, ${invalidRows} invalid rows`);
    
    if (validRows > 0) {
      Logger.log("✅ Data extraction working - issue may be elsewhere");
    } else {
      Logger.log("❌ No valid data found - check your source data");
    }
    
  } catch (error) {
    Logger.log(`❌ Test failed: ${error.message}`);
  }
}

function extractProducts(data) {
  const products = [];
  
  for (let i = HEADER_ROW; i < data.length; i++) {
    const row = data[i];
    const id = getString(row[COLUMNS.ID - 1]); // Convert to 0-based index
    const title = getString(row[COLUMNS.PRODUCT_TITLE - 1]); // Convert to 0-based index
    const currentType = getString(row[COLUMNS.PRODUCT_TYPE - 1]); // Convert to 0-based index
    
    // Only process if we have all required data
    if (id && title && currentType) {
      // Parse hierarchical product type
      const typeParts = currentType.split('>').map(part => part.trim());
      const typeL1 = typeParts[0] || '';
      const typeL2 = typeParts[1] || '';
      const typeL3 = typeParts[2] || '';
      
      products.push({
        row: i + 1,
        id: id,
        title: title,
        typeL1: typeL1,
        typeL2: typeL2,
        typeL3: typeL3,
        currentType: currentType
      });
    }
  }
  
  return products;
}

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

function writeToSheet(sheet, row, productType) {
  sheet.getRange(row, COLUMNS.OPTIMIZED_TYPE).setValue(productType);
}

function setColumnHeaders(sheet) {
  try {
    // Detect column headers first if not already done
    if (!COLUMNS.ID || !COLUMNS.PRODUCT_TITLE || !COLUMNS.PRODUCT_TYPE) {
      detectColumns(sheet);
    }
    
    // Set headers for the columns
    if (COLUMNS.ID) {
      sheet.getRange(HEADER_ROW, COLUMNS.ID).setValue("ID");
    }
    if (COLUMNS.PRODUCT_TITLE) {
      sheet.getRange(HEADER_ROW, COLUMNS.PRODUCT_TITLE).setValue("Product Title");
    }
    if (COLUMNS.PRODUCT_TYPE) {
      sheet.getRange(HEADER_ROW, COLUMNS.PRODUCT_TYPE).setValue("Product Type");
    }
    if (COLUMNS.OPTIMIZED_TYPE) {
      sheet.getRange(HEADER_ROW, COLUMNS.OPTIMIZED_TYPE).setValue("Optimized Type");
    }
    
    Logger.log(`✅ Column headers set: 'ID' (col ${COLUMNS.ID}), 'Product Title' (col ${COLUMNS.PRODUCT_TITLE}), 'Product Type' (col ${COLUMNS.PRODUCT_TYPE}), 'Optimized Type' (col ${COLUMNS.OPTIMIZED_TYPE})`);
  } catch (error) {
    Logger.log(`⚠️ Warning: Could not set headers - ${error.message}`);
  }
}

// =============================================================================
// UTILITY FUNCTIONS
// =============================================================================

function getString(value) {
  return value ? String(value).trim() : '';
}

function getApiKey(spreadsheet) {
  try {
    const namedRange = spreadsheet.getRangeByName(API_KEY_NAMED_RANGE);
    if (!namedRange) {
      Logger.log(`❌ Named range '${API_KEY_NAMED_RANGE}' not found in spreadsheet`);
      return null;
    }
    
    const apiKey = namedRange.getValue();
    if (!apiKey || typeof apiKey !== 'string' || apiKey.trim() === '') {
      Logger.log(`❌ API key in named range '${API_KEY_NAMED_RANGE}' is empty or invalid`);
      return null;
    }
    
    Logger.log(`✅ API key retrieved from named range '${API_KEY_NAMED_RANGE}'`);
    return apiKey.trim();
    
  } catch (error) {
    Logger.log(`❌ Error getting API key: ${error.message}`);
    return null;
  }
}

function validateSetup(spreadsheet) {
  let valid = true;
  
  if (!SPREADSHEET_URL || SPREADSHEET_URL.includes('YOUR_SPREADSHEET_ID')) {
    Logger.log('❌ Update SPREADSHEET_URL in the script configuration');
    valid = false;
  } else {
    Logger.log('✅ Spreadsheet URL configured');
  }
  
  // Check if API key named range exists
  try {
    const namedRange = spreadsheet.getRangeByName(API_KEY_NAMED_RANGE);
    if (!namedRange) {
      Logger.log(`❌ Create named range '${API_KEY_NAMED_RANGE}' with your OpenAI API key`);
      valid = false;
    } else {
      const apiKey = namedRange.getValue();
      if (!apiKey || typeof apiKey !== 'string' || apiKey.trim() === '') {
        Logger.log(`❌ Named range '${API_KEY_NAMED_RANGE}' is empty - add your API key`);
        valid = false;
      } else {
        Logger.log(`✅ API key found in named range '${API_KEY_NAMED_RANGE}'`);
      }
    }
  } catch (error) {
    Logger.log(`❌ Error checking API key named range: ${error.message}`);
    valid = false;
  }
  
  if (valid) {
    Logger.log('✅ Configuration validation passed');
  }
  
  return valid;
}

// =============================================================================
// TEST FUNCTIONS
// =============================================================================

function testSetup() {
  Logger.log("🧪 Testing setup...");
  
  try {
    // Test spreadsheet access
    const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
    
    // Test configuration
    if (!validateSetup(spreadsheet)) {
      return;
    }
    
    // Get API key
    const apiKey = getApiKey(spreadsheet);
    if (!apiKey) {
      Logger.log("❌ Cannot get API key - test aborted");
      return;
    }
    
    // Test source sheet access
    const sourceSheet = spreadsheet.getSheetByName(SOURCE_SHEET_NAME);
    
    if (!sourceSheet) {
      Logger.log(`❌ Source sheet '${SOURCE_SHEET_NAME}' not found`);
      return;
    }
    
    Logger.log("✅ Spreadsheet access OK");
    
    // Test column detection
    if (!detectColumns(sourceSheet)) {
      Logger.log("❌ Could not detect required columns - check your spreadsheet headers");
      return;
    }
    
    // Test creating Product Types sheet
    const targetSheet = createProductTypesSheet(spreadsheet, sourceSheet);
    Logger.log("✅ Product Types sheet creation OK");
    
    // Test API with sample product
    const sampleTitle = PRODUCT_CATEGORY === "Inflatable" ? "Giant Inflatable Pool Float" : "Sample Test Product";
    const sampleCurrentType = PRODUCT_CATEGORY === "Inflatable" ? "Pool Toys" : "Test Category";
    const result = optimizeProductType(sampleTitle, sampleCurrentType, apiKey, sampleCurrentType, '', '');
    Logger.log(`✅ API test result: ${result.productType}`);
    
    Logger.log("🎉 All tests passed!");
    Logger.log(`📊 '${TARGET_SHEET_NAME}' sheet is ready for processing`);
    
  } catch (error) {
    Logger.log(`❌ Test failed: ${error.message}`);
  }
}

function testSingle() {
  Logger.log(`🧪 Testing single ${PRODUCT_CATEGORY.toLowerCase()} product...`);
  
  try {
    // Get spreadsheet and API key
    const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
    const apiKey = getApiKey(spreadsheet);
    
    if (!apiKey) {
      Logger.log("❌ Cannot get API key - test aborted");
      return;
    }
    
    // Use category-appropriate test product
    const testTitle = PRODUCT_CATEGORY === "Inflatable" ? "Large Inflatable Water Slide with Pool" : "Sample Product for Testing";
    const testCurrentType = PRODUCT_CATEGORY === "Inflatable" ? "Water Toys" : "General Category";
    
    const result = optimizeProductType(
      testTitle, 
      testCurrentType,
      apiKey,
      testCurrentType,
      '',
      ''
    );
    
    Logger.log(`Result: ${result.productType}`);
    
  } catch (error) {
    Logger.log(`❌ Test failed: ${error.message}`);
  }
}

function setupHeaders() {
  Logger.log("📝 Setting up 'Product Types' sheet...");
  
  try {
    const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
    const sourceSheet = spreadsheet.getSheetByName(SOURCE_SHEET_NAME);
    
    if (!sourceSheet) {
      Logger.log(`❌ Source sheet '${SOURCE_SHEET_NAME}' not found`);
      return;
    }
    
    if (!detectColumns(sourceSheet)) {
      Logger.log("❌ Could not detect required columns - check your spreadsheet headers");
      return;
    }
    
    // Create new "Product Types" sheet with extracted data
    const targetSheet = createProductTypesSheet(spreadsheet, sourceSheet);
    
    Logger.log("✅ Product Types sheet setup complete!");
    
  } catch (error) {
    Logger.log(`❌ Setup failed: ${error.message}`);
  }
}

function createProductTypesSheetOnly() {
  Logger.log("📋 Creating 'Product Types' sheet only...");
  
  try {
    const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
    const sourceSheet = spreadsheet.getSheetByName(SOURCE_SHEET_NAME);
    
    if (!sourceSheet) {
      Logger.log(`❌ Source sheet '${SOURCE_SHEET_NAME}' not found`);
      return;
    }
    
    if (!detectColumns(sourceSheet)) {
      Logger.log("❌ Could not detect required columns - check your spreadsheet headers");
      return;
    }
    
    const targetSheet = createProductTypesSheet(spreadsheet, sourceSheet);
    
    Logger.log(`✅ '${TARGET_SHEET_NAME}' sheet created successfully!`);
    Logger.log("📊 You can now run main() to process the optimization");
    
  } catch (error) {
    Logger.log(`❌ Failed to create sheet: ${error.message}`);
  }
}

function processInChunks() {
  Logger.log("🔄 Processing in chunks (for large datasets)...");
  
  const CHUNK_SIZE = 50; // Process 50 products at a time
  const startTime = Date.now();
  
  try {
    const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
    
    if (!validateSetup(spreadsheet)) {
      Logger.log("❌ Setup validation failed");
      return;
    }
    
    const apiKey = getApiKey(spreadsheet);
    if (!apiKey) {
      Logger.log("❌ Failed to get API key");
      return;
    }
    
    const sourceSheet = spreadsheet.getSheetByName(SOURCE_SHEET_NAME);
    
    if (!sourceSheet) {
      Logger.log(`❌ Source sheet '${SOURCE_SHEET_NAME}' not found`);
      return;
    }
    
    if (!detectColumns(sourceSheet)) {
      Logger.log("❌ Could not detect required columns - check your spreadsheet headers");
      return;
    }
    
    // Create new "Product Types" sheet with extracted data
    const targetSheet = createProductTypesSheet(spreadsheet, sourceSheet);
    
    const data = targetSheet.getDataRange().getValues();
    const allProducts = extractProducts(data);
    
    Logger.log(`📦 Found ${allProducts.length} ${PRODUCT_CATEGORY.toLowerCase()} products. Processing in chunks of ${CHUNK_SIZE}...`);
    
    let totalProcessed = 0;
    let chunkNum = 1;
    
    for (let start = 0; start < allProducts.length; start += CHUNK_SIZE) {
      const elapsedTime = Date.now() - startTime;
      const chunk = allProducts.slice(start, start + CHUNK_SIZE);
      Logger.log(`🔄 Chunk ${chunkNum}: Processing products ${start + 1} to ${start + chunk.length} - ${Math.round(elapsedTime/60000)}min elapsed`);
      
      let successful = 0;
      for (const product of chunk) {
        try {
          const result = optimizeProductType(product.title, product.currentType, apiKey, product.typeL1, product.typeL2, product.typeL3);
          writeToSheet(targetSheet, product.row, result.productType);
          successful++;
          totalProcessed++;
        } catch (error) {
          writeToSheet(targetSheet, product.row, `Error: ${error.message}`);
          Logger.log(`❌ Row ${product.row}: ${error.message}`);
        }
        
        Utilities.sleep(100); // Minimal delay
      }
      
      Logger.log(`✅ Chunk ${chunkNum} complete: ${successful}/${chunk.length} successful`);
      chunkNum++;
      
      if (start + CHUNK_SIZE < allProducts.length) {
        Utilities.sleep(1000); // Brief pause between chunks
      }
    }
    
    const totalTime = Math.round((Date.now() - startTime) / 60000);
    Logger.log(`🎉 All chunks complete in ${totalTime} minutes! Processed ${totalProcessed} products.`);
    Logger.log(`📊 Results saved in '${TARGET_SHEET_NAME}' sheet`);
    
  } catch (error) {
    Logger.log(`❌ Chunk processing failed: ${error.message}`);
  }
}

// =============================================================================
// INSTRUCTIONS
// =============================================================================

/*
HOW TO USE:

1. UPDATE CONFIGURATION:
   - Set your SPREADSHEET_URL in the script
   - Set your PRODUCT_CATEGORY (e.g., "Electronics", "Clothing", "Home & Garden", etc.)
   - This will be the first level in your 3-level hierarchy

2. PREPARE SPREADSHEET:
   - Source sheet name: "products" (the script will look for this sheet name)
   - Required columns in source sheet (headers will be detected automatically):
     * "ID" - unique identifier for each product
     * "Product Title" or "Title" - contains product titles
     * "Product Type" - contains existing hierarchical product types (e.g., "Hardware > Steel > Fasteners")
   - CREATE NAMED RANGE: "openaiApiKey" with your OpenAI API key

3. CREATE NAMED RANGE FOR API KEY:
   - In Google Sheets: Data > Named ranges
   - Name: "openaiApiKey"
   - Range: Single cell containing your OpenAI API key (e.g., A1)
   - Value: sk-your-actual-openai-api-key-here

4. WORKFLOW:
   The script will automatically:
   - Extract ID, Product Title, and Product Type from "products" sheet
   - Create a new "Product Types" sheet with these columns
   - Add an "Optimized Type" column for the results
   - Process all products and optimize their product types

5. RUN FUNCTIONS:
   - First: testSetup() - validates setup and creates Product Types sheet
   - **DEBUGGING FUNCTIONS** (if you get errors):
     * debugSourceData() - shows your sheet headers and sample data
     * testDataExtraction() - tests if data can be extracted from your sheet
   - Optional: createProductTypesSheetOnly() - just creates the sheet without processing
   - Optional: setupHeaders() - same as above
   - Then: testSingle() - tests with one product
   - For regular datasets: main() - processes all products (optimized for speed)
   - For large datasets (100+ products): processInChunks() - safer for large batches

6. MONITOR LOGS:
   - Check Google Ads Scripts logs for progress
   - Look for ✅ (success) and ❌ (error) indicators
   - Results will be saved in the "Product Types" sheet

TROUBLESHOOTING:
- "Named range 'openaiApiKey' not found" = Create the named range in your spreadsheet
- "API key is empty" = Make sure the named range contains your actual API key
- "Source sheet 'products' not found" = Make sure your source sheet is named "products"
- "Could not detect required columns" = Make sure you have "ID", "Product Title", and "Product Type" columns
- "Spreadsheet not found" = Check URL and permissions
- "API Error" = Check OpenAI API key validity and credits
- "Parse Error" = AI response format issue (usually temporary)
- "Exceeded execution time" = May occur with very large datasets (Google Ads has 30-minute script limit)

PERFORMANCE OPTIMIZATIONS:
- Reduced delays: 200ms between products (was 1000ms)
- Larger batches: 10 products per batch (was 5)
- Progress tracking: Shows elapsed time and remaining products
- Explicit completion: Clear finish indicators

UNIVERSAL FEATURES:
- Configurable for any product category via PRODUCT_CATEGORY constant
- 3-level hierarchy: "[Your Category] > [Subcategory] > [Specific Product Type]"
- Intelligent Level 2: Automatically identifies the most relevant subcategory or attribute
- Dynamic column detection: Automatically finds required columns in source sheet
- Hierarchical parsing: Parses existing product types with ">" separators
- Automated sheet creation: Creates dedicated "Product Types" sheet for processing
- Clean data extraction: Only copies relevant columns to new sheet
- Consistent first-level categorization with your business focus
- AI-powered categorization that adapts to any product type

SECURITY BENEFIT:
- API key is stored in spreadsheet, not in script code
- More secure and easier to update

SHEET STRUCTURE:
SOURCE SHEET ("products"):
- The script will automatically detect columns by their headers
- "ID" column: Unique identifier for each product
- "Product Title" or "Title" column: Contains product titles
- "Product Type" column: Contains existing hierarchical product types (e.g., "Hardware > Steel > Fasteners")

TARGET SHEET ("Product Types"):
- Automatically created by the script
- Column A: ID (copied from source)
- Column B: Product Title (copied from source)
- Column C: Product Type (copied from source)
- Column D: Optimized Type (filled by the script)

WORKFLOW SUMMARY:
1. Run testSetup() to validate and create Product Types sheet
2. Run main() to process all products
3. Check results in the "Product Types" sheet
4. Use the optimized product types in your Google Ads campaigns
*/ 