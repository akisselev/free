# Google Ads Product Type Optimizer - Setup Guide

## 🎯 Overview
This script runs directly in Google Ads to optimize product types for Google Merchant Center using OpenAI API. It reads product data from Google Sheets and updates it with optimized two-level product type hierarchies.

## 🚀 Setup Instructions

### Step 1: Access Google Ads Scripts
1. Go to your **Google Ads account**
2. Navigate to **Tools & Settings** > **Bulk Actions** > **Scripts**
3. Click **"+ New Script"**
4. Give it a name like "Product Type Optimizer"

### Step 2: Configure the Script
1. Copy the code from `google-ads-product-types.js`
2. Paste it into the Google Ads script editor
3. Update the configuration at the top:

```javascript
// INSERT YOUR GOOGLE SHEETS URL HERE
const SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/YOUR_ACTUAL_SPREADSHEET_ID/edit#gid=0";

// INSERT YOUR OPENAI API KEY HERE
const OPENAI_API_KEY = "sk-your-actual-openai-api-key-here";
```

### Step 3: Prepare Your Google Sheets
Set up your spreadsheet with this structure:

| Column A | Column B (Product Title) | Column C (Current Type) | Column D (Optimized Type) | Column E (Confidence) |
|----------|-------------------------|-------------------------|---------------------------|----------------------|
| ID       | Bluetooth Headphones    | Electronics             | *Will be filled*          | *Will be filled*     |
| ID       | Running Shoes           | Footwear               | *Will be filled*          | *Will be filled*     |

### Step 4: Authorize and Test
1. Click **"Authorize"** to grant permissions
2. Run the **`testSetup()`** function first to validate everything
3. If successful, run **`testSingleProduct()`** to test with sample data
4. Finally, run **`main()`** to process all products

## ⚙️ Configuration Options

### Sheet Settings
```javascript
const SHEET_NAME = "Sheet1";         // Name of your sheet tab
const HEADER_ROW = 1;                // Row number where headers are located
```

### Column Mapping
```javascript
const COLUMNS = {
  PRODUCT_TITLE: 'B',        // Column B - Product titles
  CURRENT_TYPE: 'C',         // Column C - Current product types
  OPTIMIZED_TYPE: 'D',       // Column D - New optimized product types
  CONFIDENCE_SCORE: 'E'      // Column E - Confidence score
};
```

### Processing Settings
```javascript
const BATCH_SIZE = 10;                // Process products in batches
const DELAY_BETWEEN_BATCHES = 5000;   // 5 seconds delay between batches
```

## 🔧 Available Functions

### Main Functions
- **`main()`** - Process all products in your spreadsheet
- **`testSetup()`** - Validate configuration and test connections
- **`testSingleProduct()`** - Test with a single sample product

### How to Run
1. **In Google Ads Scripts**: Select the function from the dropdown and click "Run"
2. **Preview before running**: Use "Preview" to see what will happen
3. **Schedule automated runs**: Set up recurring schedules in Google Ads

## 📊 Expected Output

The script will generate two-level product type hierarchies:

**Input:**
- Title: "Sony WH-1000XM4 Wireless Headphones"
- Current Type: "Electronics"

**Output:**
- Optimized Type: "Electronics > Wireless Headphones"
- Confidence: "High"

## 🔍 Monitoring and Logs

### View Logs
1. In Google Ads Scripts, click **"Logs"** after running
2. Check for success/error messages
3. Monitor processing progress

### Sample Log Output
```
Starting Google Ads Product Type Optimization
✅ Configuration validated successfully
Found 25 products to process
Processing batch 1 of 3
✅ Row 2: Electronics > Wireless Headphones (High)
✅ Row 3: Apparel & Accessories > Running Shoes (High)
❌ Row 4: Error: Invalid product title
✅ Optimization complete! 24/25 products processed successfully
```

## 🛠️ Troubleshooting

### Common Issues

#### 1. "Authorization required"
**Solution:** Click "Authorize" and grant all requested permissions

#### 2. "Spreadsheet not found"
**Solution:** Check your SPREADSHEET_URL is correct and accessible

#### 3. "OpenAI API error"
**Solution:** Verify your API key is valid and has credits

#### 4. "Sheet not found"
**Solution:** Update SHEET_NAME to match your actual sheet tab name

#### 5. "No data found"
**Solution:** Make sure your spreadsheet has data starting from the header row

### Performance Tips

1. **Batch Processing**: The script processes products in batches to avoid timeouts
2. **Rate Limiting**: Built-in delays prevent API rate limit issues
3. **Error Handling**: Individual product errors won't stop the entire process

### API Limits
- **OpenAI API**: Has rate limits based on your plan
- **Google Sheets API**: Has quotas for read/write operations
- **Google Ads Scripts**: Have execution time limits (30 minutes max)

## 📈 Advanced Usage

### Scheduling
1. In Google Ads Scripts, set up **"Frequency"** to run automatically
2. Choose daily, weekly, or monthly schedules
3. Monitor performance and adjust as needed

### Integration with Google Ads
The script includes a placeholder for Google Ads integration:
```javascript
function updateGoogleAdsProductTypes() {
  // Future enhancement: Update product types in Shopping campaigns
}
```

### Custom Prompts
Modify the `getOptimizationPrompt()` function to customize AI behavior:
```javascript
function getOptimizationPrompt() {
  return `Your custom prompt here...`;
}
```

## 🔐 Security Notes

- **API Keys**: Never share your OpenAI API key
- **Permissions**: Only grant necessary Google Ads permissions
- **Data**: Your product data is processed securely through Google's systems

## 📞 Support

If you encounter issues:
1. Check the logs for detailed error messages
2. Run `testSetup()` to diagnose configuration problems
3. Verify your Google Sheets structure matches the expected format
4. Ensure your OpenAI API key has sufficient credits

## 🎉 Next Steps

After successful optimization:
1. **Review Results**: Check the optimized product types in your spreadsheet
2. **Apply to Merchant Center**: Use the optimized types in your Google Merchant Center feed
3. **Monitor Performance**: Track improvements in your Google Ads Shopping campaigns
4. **Iterate**: Run the script regularly to optimize new products 