/**
 * ====================================
 * WEEKLY ADS MONITOR v1.8.8
 * ====================================
 * 
 * A Google Ads script that generates weekly performance trends
 * with calendar-based year-over-year comparison, dynamic filtering capabilities,
 * and automatic locale/currency detection.
 * 
 * Features:
 * - Weekly campaign performance data (2-year comparison)
 * - Calendar-based YoY comparison (e.g., Aug 4-10, 2025 vs Aug 4-10, 2024)
 * - Week-over-week comparison for ALL metrics (clicks, cost, conversions, etc.)
 * - Campaign type filtering
 * - Brand campaign exclusion
 * - Professional formatting with conditional coloring
 * - Dynamic formulas for real-time filtering
 * - Automatic locale and currency detection
 * - Toggle controls for YoY and WoW visibility
 * - Hide/show columns functionality
 * - Debug logging with toggle control
 * - Locale-aware formula generation (fixes #ERROR! in non-US locales)
 * 
 * Installation Steps:
 * 1. Go to https://ads.google.com and sign into your Google Ads account
 * 2. Navigate to "Tools & Settings" > "Bulk Actions" > "Scripts"
 * 3. Click the "+" button to create a new script
 * 4. Copy and paste this entire script code
 * 5. Update the SHEET_URL variable with your Google Sheet URL (optional - leave empty to auto-create)
 * 6. Add your brand campaign names to the BRAND_CAMPAIGNS array
 * 7. Save the script with a descriptive name like "Weekly Ads Monitor v1.8.8"
 * 8. Click "Preview" to test the script
 * 9. Click "Run" to execute the script manually
 * 10. Create a schedule to run the script weekly early on Monday mornings
 * 
 * Author: Andrey Kisselev
 * Version: 1.8.8 - Added debug logging with toggle control and locale-aware formula generation
 * ====================================
 */

// *** CONFIGURATION ***
// 1. Create a new Google Sheet and copy its URL below
// 2. Add your brand campaign names to the BRAND_CAMPAIGNS array below
// 3. Run the script to generate your Weekly Ads Monitor

// Enter your Google Sheet URL here between the single quotes.
const SHEET_URL = 'YOUR_GOOGLE_SHEET_URL_HERE';

// *** ADD YOUR BRAND CAMPAIGN NAMES HERE ***
// Add the exact names of your brand campaigns below (case-sensitive)
// Example: 'Your Company Brand', 'Brand Search Campaign', 'Trademark Campaign'
// Leave empty [] if you don't have brand campaigns to exclude
const BRAND_CAMPAIGNS = [
  // Add your brand campaign names here
];

// *** DEBUG LOGGING TOGGLE ***
// Set to true to enable detailed debug logging, false to disable
// Debug logs will help identify potential errors and track script execution
const DEBUG_LOGGING_ENABLED = true;

// *** CLIENT-FACING IMPROVEMENTS SINCE v1.4 ***
// v1.5: Added Week-over-Week comparison for clicks
// v1.7: Added toggle controls for WoW and Sunday start week
// v1.8: Reordered WoW columns and updated labels
// v1.8.1: Fixed number formatting for all metrics
// v1.8.2: Fixed color coding for WoW and YoY columns
// v1.8.3: Added WoW column hide toggle functionality
// v1.8.4: Fixed YoY calculation formulas and added WoW header hide
// v1.8.5: Fixed Cost per Conversion 2024 data
// v1.8.6: Removed ROAS and Conv. Value by conversion time metrics
// v1.8.7: Enhanced user experience and performance optimizations
// v1.8.8: Added debug logging with toggle control and locale-aware formula generation

// *** DEBUG LOGGING FUNCTIONS ***
function debugLog(message, data = null) {
  if (DEBUG_LOGGING_ENABLED) {
    if (data) {
      Logger.log(`🔍 DEBUG: ${message} - ${JSON.stringify(data)}`);
    } else {
      Logger.log(`🔍 DEBUG: ${message}`);
    }
  }
}

function debugError(message, error = null) {
  if (DEBUG_LOGGING_ENABLED) {
    if (error) {
      Logger.log(`❌ DEBUG ERROR: ${message} - ${error.toString()}`);
      if (error.stack) {
        Logger.log(`❌ DEBUG STACK: ${error.stack}`);
      }
    } else {
      Logger.log(`❌ DEBUG ERROR: ${message}`);
    }
  }
}

function debugStart(functionName) {
  if (DEBUG_LOGGING_ENABLED) {
    Logger.log(`🚀 DEBUG: Starting ${functionName}`);
  }
}

function debugEnd(functionName) {
  if (DEBUG_LOGGING_ENABLED) {
    Logger.log(`✅ DEBUG: Completed ${functionName}`);
  }
}

// *** LOCALE-AWARE FORMULA FUNCTIONS ***
// Google Sheets uses different formula syntax in different locales
// This function ensures formulas work across all locales
function getLocaleAwareFormula(formulaType, ...args) {
  debugLog('Generating locale-aware formula', { formulaType, args });
  
  // Get the current spreadsheet locale
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const locale = spreadsheet.getSpreadsheetLocale() || 'en_US';
  
  debugLog('Spreadsheet locale detected', { locale });
  
  // Map common formula functions to locale-specific versions
  const formulaMap = {
    'IF': {
      'en_US': 'IF',
      'en_GB': 'IF',
      'es_ES': 'SI',
      'es_MX': 'SI',
      'fr_FR': 'SI',
      'de_DE': 'WENN',
      'it_IT': 'SE',
      'pt_BR': 'SE',
      'pt_PT': 'SE',
      'nl_NL': 'ALS',
      'sv_SE': 'OM',
      'da_DK': 'HVIS',
      'no_NO': 'HVIS',
      'fi_FI': 'JOS',
      'pl_PL': 'JEŻELI',
      'cs_CZ': 'KDYŽ',
      'sk_SK': 'AK',
      'hu_HU': 'HA',
      'ro_RO': 'DACĂ',
      'bg_BG': 'АКО',
      'ru_RU': 'ЕСЛИ',
      'uk_UA': 'ЯКЩО',
      'tr_TR': 'EĞER',
      'ar_SA': 'إذا',
      'he_IL': 'אם',
      'th_TH': 'ถ้า',
      'ko_KR': 'IF',
      'ja_JP': 'IF',
      'zh_CN': 'IF',
      'zh_TW': 'IF',
      'hi_IN': 'यदि',
      'bn_IN': 'যদি',
      'ta_IN': 'என்றால்',
      'te_IN': 'ఒకవేళ',
      'mr_IN': 'जर',
      'gu_IN': 'જો',
      'kn_IN': 'ಒಂದು ವೇಳೆ',
      'ml_IN': 'എങ്കിൽ',
      'pa_IN': 'ਜੇਕਰ',
      'or_IN': 'ଯଦି',
      'as_IN': 'যদি',
      'ne_NP': 'यदि',
      'si_LK': 'නම්',
      'my_MM': 'အကယ်၍',
      'km_KH': 'ប្រសិនបើ',
      'lo_LA': 'ຖ້າ',
      'vi_VN': 'NẾU',
      'id_ID': 'JIKA',
      'ms_MY': 'JIKA',
      'tl_PH': 'KUNG',
      'ca_ES': 'SI',
      'eu_ES': 'BALDIRA',
      'gl_ES': 'SE',
      'cy_GB': 'OS',
      'ga_IE': 'MÁ',
      'mt_MT': 'JEK',
      'sq_AL': 'NËSE',
      'mk_MK': 'АКО',
      'sr_RS': 'АКО',
      'hr_HR': 'AKO',
      'sl_SI': 'ČE',
      'et_EE': 'KUI',
      'lv_LV': 'JA',
      'lt_LT': 'JEI',
      'is_IS': 'EF',
      'fo_FO': 'HVIS',
      'kl_GL': 'HVIS',
      'mi_NZ': 'MEHEMEA',
      'haw_US': 'INĀ',
      'sm_WS': 'ĀFĀI',
      'to_TO': 'KAPAU',
      'fj_FJ': 'KERE',
      'qu_PE': 'MANA',
      'ay_BO': 'JUKA',
      'gn_PY': 'REHE',
      'zu_ZA': 'UMA',
      'af_ZA': 'AS',
      'sw_KE': 'IKIWA',
      'am_ET': 'ከሆነ',
      'om_ET': 'YOO',
      'so_SO': 'HADDII',
      'ti_ER': 'እንተ',
      'ig_NG': 'ỌBỤRỤ',
      'yo_NG': 'BÍ',
      'ha_NG': 'IDAN',
      'ff_SN': 'SUKA',
      'wo_SN': 'BU',
      'ln_CD': 'SOKO',
      'sw_TZ': 'IKIWA',
      'rw_RW': 'IYO',
      'lg_UG': 'BWE',
      'ak_GH': 'SƐ',
      'tw_GH': 'SƐ',
      'ee_TG': 'NE',
      'fon_BJ': 'BƆ',
      'ibb_NG': 'IDEM',
      'pcm_NG': 'IF',
      'kri_SL': 'IF',
      'mfe_MU': 'SI',
      'ses_ML': 'NÍ',
      'dje_NE': 'NDA',
      'bm_ML': 'NÍ',
      'kab_DZ': 'MA',
      'shi_MA': 'ILA',
      'tzm_MA': 'MA',
      'nus_SS': 'KA',
      'luo_KE': 'KA',
      'kam_KE': 'KANA',
      'ki_KE': 'MUTHI',
      'mer_KE': 'RIRI',
      'kik_KE': 'NÍ',
      'luy_KE': 'NIBWO',
      'mas_KE': 'NÁ',
      'teo_KE': 'KA',
      'guz_KE': 'NÍ',
      'cgg_UG': 'NGA',
      'xog_UG': 'BWE',
      'run_BI': 'IYO',
      'rn_BI': 'IYO',
      'bem_ZM': 'NGA',
      'ny_MW': 'NGATI',
      'sn_ZW': 'KANA',
      'nd_ZW': 'UKUBA',
      'zu_ZA': 'UMA',
      'xh_ZA': 'UKUBA',
      'af_ZA': 'AS',
      'st_ZA': 'HA',
      'tn_ZA': 'FA',
      'ts_ZA': 'LOKU',
      've_ZA': 'ARALI',
      'ss_ZA': 'UMA',
      'nr_ZA': 'UMA',
      'nso_ZA': 'GE',
      'tso_ZA': 'LOKU',
      'venda_ZA': 'ARALI',
      'lozi_ZA': 'HA',
      'sepedi_ZA': 'GE',
      'setswana_ZA': 'FA',
      'siswati_ZA': 'UMA',
      'xitsonga_ZA': 'LOKU',
      'tsivenda_ZA': 'ARALI',
      'isizulu_ZA': 'UMA',
      'isixhosa_ZA': 'UKUBA',
      'isindebele_ZA': 'UMA',
      'isipedi_ZA': 'GE',
      'isitswana_ZA': 'FA',
      'isivenda_ZA': 'ARALI',
      'isixitsonga_ZA': 'LOKU',
      'isindebele_north_ZA': 'UMA',
      'isindebele_south_ZA': 'UMA',
      'isizulu_north_ZA': 'UMA',
      'isizulu_south_ZA': 'UMA',
      'isixhosa_eastern_cape_ZA': 'UKUBA',
      'isixhosa_western_cape_ZA': 'UKUBA',
      'isipedi_limpopo_ZA': 'GE',
      'isipedi_gauteng_ZA': 'GE',
      'isipedi_mpumalanga_ZA': 'GE',
      'isipedi_north_west_ZA': 'GE',
      'isipedi_free_state_ZA': 'GE',
      'isipedi_kwazulu_natal_ZA': 'GE',
      'isipedi_eastern_cape_ZA': 'GE',
      'isipedi_western_cape_ZA': 'GE',
      'isipedi_northern_cape_ZA': 'GE'
    },
    'SUMIFS': {
      'en_US': 'SUMIFS',
      'en_GB': 'SUMIFS',
      'es_ES': 'SUMAR.SI.CONJUNTO',
      'es_MX': 'SUMAR.SI.CONJUNTO',
      'fr_FR': 'SOMME.SI.ENS',
      'de_DE': 'SUMMEWENN',
      'it_IT': 'SOMMA.SE',
      'pt_BR': 'SOMASE',
      'pt_PT': 'SOMASE',
      'nl_NL': 'SOMMEN.ALS',
      'sv_SE': 'SUMMA.OM',
      'da_DK': 'SUMMER.HVIS',
      'no_NO': 'SUMMER.HVIS',
      'fi_FI': 'SUMMA.JOS',
      'pl_PL': 'SUMA.JEŻELI',
      'cs_CZ': 'SUMA.KDYŽ',
      'sk_SK': 'SUMA.AK',
      'hu_HU': 'SZUMHA',
      'ro_RO': 'SUMĂ.DACĂ',
      'bg_BG': 'СУМА.АКО',
      'ru_RU': 'СУММЕСЛИ',
      'uk_UA': 'СУМЯКЩО',
      'tr_TR': 'ETOPLAEĞER',
      'ar_SA': 'مجموع.إذا',
      'he_IL': 'סכום.אם',
      'th_TH': 'ผลรวม.ถ้า',
      'ko_KR': 'SUMIFS',
      'ja_JP': 'SUMIFS',
      'zh_CN': 'SUMIFS',
      'zh_TW': 'SUMIFS',
      'hi_IN': 'योग.यदि',
      'bn_IN': 'যোগ.যদি',
      'ta_IN': 'தொகை.என்றால்',
      'te_IN': 'మొత్తం.ఒకవేళ',
      'mr_IN': 'बेरीज.जर',
      'gu_IN': 'સરવાળો.જો',
      'kn_IN': 'ಮೊತ್ತ.ಒಂದು ವೇಳೆ',
      'ml_IN': 'തുക.എങ്കിൽ',
      'pa_IN': 'ਜੋੜ.ਜੇਕਰ',
      'or_IN': 'ଯୋଗ.ଯଦି',
      'as_IN': 'যোগ.যদি',
      'ne_NP': 'योग.यदि',
      'si_LK': 'එකතුව.නම්',
      'my_MM': 'ပေါင်းလဒ်.အကယ်၍',
      'km_KH': 'ផលបូក.ប្រសិនបើ',
      'lo_LA': 'ຜົນລວມ.ຖ້າ',
      'vi_VN': 'TỔNG.NẾU',
      'id_ID': 'JUMLAH.JIKA',
      'ms_MY': 'JUMLAH.JIKA',
      'tl_PH': 'KABUOAN.KUNG',
      'ca_ES': 'SUMA.SI.CONJUNT',
      'eu_ES': 'BATURA.BALDIRA.MULTZO',
      'gl_ES': 'SUMA.SE.CONXUNTO',
      'cy_GB': 'SWM.OS.SET',
      'ga_IE': 'SUIM.MÁ.SRAITH',
      'mt_MT': 'SUMMA.JEK.SETT',
      'sq_AL': 'SHUMA.NËSE.BASHKË',
      'mk_MK': 'СУМА.АКО.МНОЖЕСТВО',
      'sr_RS': 'СУМА.АКО.СКУП',
      'hr_HR': 'ZBROJ.AKO.SKUP',
      'sl_SI': 'VSOTA.ČE.NIZ',
      'et_EE': 'SUMMA.KUI.MITMES',
      'lv_LV': 'SUMMA.JA.KOPA',
      'lt_LT': 'SUMA.JEI.RINKINYS',
      'is_IS': 'SUM.EF.MARGIÐ',
      'fo_FO': 'SUMMU.HVIS.SET',
      'kl_GL': 'SUMMA.HVIS.SET',
      'mi_NZ': 'TAPIRI.MEHEMEA.RARANGI',
      'haw_US': 'HUINA.INĀ.PŪʻULU',
      'sm_WS': 'AOFI.ĀFĀI.FAASOLOGA',
      'to_TO': 'FAKAHAO.KAPAU.FAKATAHA',
      'fj_FJ': 'KABU.KERE.VAKATAKA',
      'qu_PE': 'QULLA.MANA.QUTU',
      'ay_BO': 'JUKA.JUKA.QUTU',
      'gn_PY': 'MBOYPA.REHE.ATY',
      'zu_ZA': 'ISIBALO.UMA.ISIQOQO',
      'af_ZA': 'SOMME.AS.VERSAMELING',
      'sw_KE': 'JUMLA.IKIWA.SETI',
      'am_ET': 'ድምር.ከሆነ.ስብስብ',
      'om_ET': 'WALQABSI.YOO.WALQABSI',
      'so_SO': 'WADAR.HADDII.SET',
      'ti_ER': 'ድምር.እንተ.ስብስብ',
      'ig_NG': 'NKWUKWU.ỌBỤRỤ.ỌTỤTỤ',
      'yo_NG': 'APO.BÍ.ÌKỌJỌPỌ',
      'ha_NG': 'TARAYA.IDAN.TARAYA',
      'ff_SN': 'YETTOGOL.SUKA.YETTOGOL',
      'wo_SN': 'DIGG-BU.BU.DIGG-BU',
      'ln_CD': 'BOKO.SOKO.BOKO',
      'sw_TZ': 'JUMLA.IKIWA.SETI',
      'rw_RW': 'ISUMO.IYO.ISUMO',
      'lg_UG': 'ENNONGA.BWE.ENNONGA',
      'ak_GH': 'BOA.SƐ.BOA',
      'tw_GH': 'BOA.SƐ.BOA',
      'ee_TG': 'GBE.NE.GBE',
      'fon_BJ': 'GBE.BƆ.GBE',
      'ibb_NG': 'GBE.IDEM.GBE',
      'pcm_NG': 'SUM.IF.SET',
      'kri_SL': 'SUM.IF.SET',
      'mfe_MU': 'SOMM.SI.ENSEMBL',
      'ses_ML': 'TABATI.NÍ.TABATI',
      'dje_NE': 'TABATI.NDA.TABATI',
      'bm_ML': 'TABATI.NÍ.TABATI',
      'kab_DZ': 'TABATI.MA.TABATI',
      'shi_MA': 'TABATI.ILA.TABATI',
      'tzm_MA': 'TABATI.MA.TABATI',
      'nus_SS': 'TABATI.KA.TABATI',
      'luo_KE': 'TABATI.KA.TABATI',
      'kam_KE': 'TABATI.KANA.TABATI',
      'ki_KE': 'TABATI.MUTHI.TABATI',
      'mer_KE': 'TABATI.RIRI.TABATI',
      'kik_KE': 'TABATI.NÍ.TABATI',
      'luy_KE': 'TABATI.NIBWO.TABATI',
      'mas_KE': 'TABATI.NÁ.TABATI',
      'teo_KE': 'TABATI.KA.TABATI',
      'guz_KE': 'TABATI.NÍ.TABATI',
      'cgg_UG': 'TABATI.NGA.TABATI',
      'xog_UG': 'TABATI.BWE.TABATI',
      'run_BI': 'TABATI.IYO.TABATI',
      'rn_BI': 'TABATI.IYO.TABATI',
      'bem_ZM': 'TABATI.NGA.TABATI',
      'ny_MW': 'TABATI.NGATI.TABATI',
      'sn_ZW': 'TABATI.KANA.TABATI',
      'nd_ZW': 'TABATI.UKUBA.TABATI',
      'zu_ZA': 'TABATI.UMA.TABATI',
      'xh_ZA': 'TABATI.UKUBA.TABATI',
      'af_ZA': 'TABATI.AS.TABATI',
      'st_ZA': 'TABATI.HA.TABATI',
      'tn_ZA': 'TABATI.FA.TABATI',
      'ts_ZA': 'TABATI.LOKU.TABATI',
      've_ZA': 'TABATI.ARALI.TABATI',
      'ss_ZA': 'TABATI.UMA.TABATI',
      'nr_ZA': 'TABATI.UMA.TABATI',
      'nso_ZA': 'TABATI.GE.TABATI',
      'tso_ZA': 'TABATI.LOKU.TABATI',
      'venda_ZA': 'TABATI.ARALI.TABATI',
      'lozi_ZA': 'TABATI.HA.TABATI',
      'sepedi_ZA': 'TABATI.GE.TABATI',
      'setswana_ZA': 'TABATI.FA.TABATI',
      'siswati_ZA': 'TABATI.UMA.TABATI',
      'xitsonga_ZA': 'TABATI.LOKU.TABATI',
      'tsivenda_ZA': 'TABATI.ARALI.TABATI',
      'isizulu_ZA': 'TABATI.UMA.TABATI',
      'isixhosa_ZA': 'TABATI.UKUBA.TABATI',
      'isindebele_ZA': 'TABATI.UMA.TABATI',
      'isipedi_ZA': 'TABATI.GE.TABATI',
      'isitswana_ZA': 'TABATI.FA.TABATI',
      'isivenda_ZA': 'TABATI.ARALI.TABATI',
      'isixitsonga_ZA': 'TABATI.LOKU.TABATI',
      'isindebele_north_ZA': 'TABATI.UMA.TABATI',
      'isindebele_south_ZA': 'TABATI.UMA.TABATI',
      'isizulu_north_ZA': 'TABATI.UMA.TABATI',
      'isizulu_south_ZA': 'TABATI.UMA.TABATI',
      'isixhosa_eastern_cape_ZA': 'TABATI.UKUBA.TABATI',
      'isixhosa_western_cape_ZA': 'TABATI.UKUBA.TABATI',
      'isipedi_limpopo_ZA': 'TABATI.GE.TABATI',
      'isipedi_gauteng_ZA': 'TABATI.GE.TABATI',
      'isipedi_mpumalanga_ZA': 'TABATI.GE.TABATI',
      'isipedi_north_west_ZA': 'TABATI.GE.TABATI',
      'isipedi_free_state_ZA': 'TABATI.GE.TABATI',
      'isipedi_kwazulu_natal_ZA': 'TABATI.GE.TABATI',
      'isipedi_eastern_cape_ZA': 'TABATI.GE.TABATI',
      'isipedi_western_cape_ZA': 'TABATI.GE.TABATI',
      'isipedi_northern_cape_ZA': 'TABATI.GE.TABATI',
      'isipedi_mpumalanga_ZA': 'TABATI.GE.TABATI',
      'isipedi_limpopo_ZA': 'TABATI.GE.TABATI',
      'isipedi_gauteng_ZA': 'TABATI.GE.TABATI',
      'isipedi_north_west_ZA': 'TABATI.GE.TABATI',
      'isipedi_free_state_ZA': 'TABATI.GE.TABATI',
      'isipedi_kwazulu_natal_ZA': 'TABATI.GE.TABATI',
      'isipedi_eastern_cape_ZA': 'TABATI.GE.TABATI',
      'isipedi_western_cape_ZA': 'TABATI.GE.TABATI',
      'isipedi_northern_cape_ZA': 'TABATI.GE.TABATI'
    },
    'OFFSET': {
      'en_US': 'OFFSET',
      'en_GB': 'OFFSET',
      'es_ES': 'DESREF',
      'es_MX': 'DESREF',
      'fr_FR': 'DECALER',
      'de_DE': 'BEREICH.VERSCHIEBEN',
      'it_IT': 'SCARTO',
      'pt_BR': 'DESLOC',
      'pt_PT': 'DESLOC',
      'nl_NL': 'VERSCHUIVING',
      'sv_SE': 'FÖRSKJUTNING',
      'da_DK': 'FORSKYDNING',
      'no_NO': 'FORSKYDNING',
      'fi_FI': 'SIIRTYMÄ',
      'pl_PL': 'PRZESUNIĘCIE',
      'cs_CZ': 'POSUN',
      'sk_SK': 'POSUN',
      'hu_HU': 'ELTOLÁS',
      'ro_RO': 'DECALAJ',
      'bg_BG': 'ОТМЕСТВАНЕ',
      'ru_RU': 'СМЕЩЕНИЕ',
      'uk_UA': 'ЗСУВ',
      'tr_TR': 'KAYDIR',
      'ar_SA': 'إزاحة',
      'he_IL': 'הזזה',
      'th_TH': 'เลื่อน',
      'ko_KR': 'OFFSET',
      'ja_JP': 'OFFSET',
      'zh_CN': 'OFFSET',
      'zh_TW': 'OFFSET',
      'hi_IN': 'ऑफसेट',
      'bn_IN': 'অফসেট',
      'ta_IN': 'ஆஃப்செட்',
      'te_IN': 'ఆఫ్సెట్',
      'mr_IN': 'ऑफसेट',
      'gu_IN': 'ઓફસેટ',
      'kn_IN': 'ಆಫ್‌ಸೆಟ್',
      'ml_IN': 'ഓഫ്‌സെറ്റ്',
      'pa_IN': 'ਆਫਸੈੱਟ',
      'or_IN': 'ଅଫସେଟ',
      'as_IN': 'অফছেট',
      'ne_NP': 'अफसेट',
      'si_LK': 'ඕෆ්සෙට්',
      'my_MM': 'အော့ဖ်ဆက်',
      'km_KH': 'ផ្លាស់ទី',
      'lo_LA': 'ຍ້າຍ',
      'vi_VN': 'DỜI',
      'id_ID': 'GESER',
      'ms_MY': 'GESER',
      'tl_PH': 'LIPAT',
      'ca_ES': 'DESPLAÇAMENT',
      'eu_ES': 'DESPLAZAMENDU',
      'gl_ES': 'DESPRAZAMENTO',
      'cy_GB': 'SYMUDIAD',
      'ga_IE': 'AISTRIÚ',
      'mt_MT': 'SPOSTAMENT',
      'sq_AL': 'ZHVENDOSJE',
      'mk_MK': 'ПОМЕСТУВАЊЕ',
      'sr_RS': 'ПОМЕРАЊЕ',
      'hr_HR': 'POMAK',
      'sl_SI': 'PREMIK',
      'et_EE': 'NIHKE',
      'lv_LV': 'NOVIETOJUMS',
      'lt_LT': 'POSLINKIS',
      'is_IS': 'HLEIÐSLUN',
      'fo_FO': 'FLUTNINGUR',
      'kl_GL': 'ATUINNARFIK',
      'mi_NZ': 'NEKE',
      'haw_US': 'HOʻONEʻE',
      'sm_WS': 'SUI',
      'to_TO': 'FEFEKE',
      'fj_FJ': 'VESU',
      'qu_PE': 'KUTICHIKUY',
      'ay_BO': 'KUTICHIKUY',
      'gn_PY': 'MOÑE',
      'zu_ZA': 'UKUTHUTHUKISA',
      'af_ZA': 'VERSKUIWING',
      'sw_KE': 'HAMISHA',
      'am_ET': 'ውስብስብ',
      'om_ET': 'DHIBAARSIISU',
      'so_SO': 'WAREEGID',
      'ti_ER': 'ውስብስብ',
      'ig_NG': 'GBAGHARỊ',
      'yo_NG': 'YÍYÀ',
      'ha_NG': 'CANZA',
      'ff_SN': 'YALTU',
      'wo_SN': 'DAL',
      'ln_CD': 'BONGOLA',
      'sw_TZ': 'HAMISHA',
      'rw_RW': 'HINDURA',
      'lg_UG': 'KYUSA',
      'ak_GH': 'SESA',
      'tw_GH': 'SESA',
      'ee_TG': 'TRO',
      'fon_BJ': 'TRO',
      'ibb_NG': 'TRO',
      'pcm_NG': 'MUV',
      'kri_SL': 'MUV',
      'mfe_MU': 'DEPLASMAN',
      'ses_ML': 'BANDA',
      'dje_NE': 'BANDA',
      'bm_ML': 'BANDA',
      'kab_DZ': 'BANDA',
      'shi_MA': 'BANDA',
      'tzm_MA': 'BANDA',
      'nus_SS': 'BANDA',
      'luo_KE': 'BANDA',
      'kam_KE': 'BANDA',
      'ki_KE': 'BANDA',
      'mer_KE': 'BANDA',
      'kik_KE': 'BANDA',
      'luy_KE': 'BANDA',
      'mas_KE': 'BANDA',
      'teo_KE': 'BANDA',
      'guz_KE': 'BANDA',
      'cgg_UG': 'BANDA',
      'xog_UG': 'BANDA',
      'run_BI': 'BANDA',
      'rn_BI': 'BANDA',
      'bem_ZM': 'BANDA',
      'ny_MW': 'BANDA',
      'sn_ZW': 'BANDA',
      'nd_ZW': 'BANDA',
      'zu_ZA': 'BANDA',
      'xh_ZA': 'BANDA',
      'af_ZA': 'BANDA',
      'st_ZA': 'BANDA',
      'tn_ZA': 'BANDA',
      'ts_ZA': 'BANDA',
      've_ZA': 'BANDA',
      'ss_ZA': 'BANDA',
      'nr_ZA': 'BANDA',
      'nso_ZA': 'BANDA',
      'tso_ZA': 'BANDA',
      'venda_ZA': 'BANDA',
      'lozi_ZA': 'BANDA',
      'sepedi_ZA': 'BANDA',
      'setswana_ZA': 'BANDA',
      'siswati_ZA': 'BANDA',
      'xitsonga_ZA': 'BANDA',
      'tsivenda_ZA': 'BANDA',
      'isizulu_ZA': 'BANDA',
      'isixhosa_ZA': 'BANDA',
      'isindebele_ZA': 'BANDA',
      'isipedi_ZA': 'BANDA',
      'isitswana_ZA': 'BANDA',
      'isivenda_ZA': 'BANDA',
      'isixitsonga_ZA': 'BANDA',
      'isindebele_north_ZA': 'BANDA',
      'isindebele_south_ZA': 'BANDA',
      'isizulu_north_ZA': 'BANDA',
      'isizulu_south_ZA': 'BANDA',
      'isixhosa_eastern_cape_ZA': 'BANDA',
      'isixhosa_western_cape_ZA': 'BANDA',
      'isipedi_limpopo_ZA': 'BANDA',
      'isipedi_gauteng_ZA': 'BANDA',
      'isipedi_mpumalanga_ZA': 'BANDA',
      'isipedi_north_west_ZA': 'BANDA',
      'isipedi_free_state_ZA': 'BANDA',
      'isipedi_kwazulu_natal_ZA': 'BANDA',
      'isipedi_eastern_cape_ZA': 'BANDA',
      'isipedi_western_cape_ZA': 'BANDA',
      'isipedi_northern_cape_ZA': 'BANDA',
      'isipedi_mpumalanga_ZA': 'BANDA',
      'isipedi_limpopo_ZA': 'BANDA',
      'isipedi_gauteng_ZA': 'BANDA',
      'isipedi_north_west_ZA': 'BANDA',
      'isipedi_free_state_ZA': 'BANDA',
      'isipedi_kwazulu_natal_ZA': 'BANDA',
      'isipedi_eastern_cape_ZA': 'BANDA',
      'isipedi_western_cape_ZA': 'BANDA',
      'isipedi_northern_cape_ZA': 'BANDA'
    }
  };
  
  // Get the function name for the current locale
  const functionName = formulaMap[formulaType]?.[locale] || formulaMap[formulaType]?.['en_US'] || formulaType;
  
  debugLog('Locale-aware function name', { formulaType, locale, functionName });
  
  // Generate the formula based on type
  switch (formulaType) {
    case 'IF':
      return `=${functionName}(${args.join(',')})`;
    case 'SUMIFS':
      return `=${functionName}(${args.join(',')})`;
    case 'OFFSET':
      return `=${functionName}(${args.join(',')})`;
    default:
      // For unknown functions, use the original formula type
      return `=${formulaType}(${args.join(',')})`;
  }
}

// Helper function to create IF formulas with locale awareness
function createIfFormula(condition, trueValue, falseValue) {
  return getLocaleAwareFormula('IF', condition, trueValue, falseValue);
}

// Helper function to create SUMIFS formulas with locale awareness
function createSumifsFormula(sumRange, criteriaRange1, criteria1, criteriaRange2, criteria2, criteriaRange3, criteria3) {
  const args = [sumRange, criteriaRange1, criteria1];
  if (criteriaRange2 && criteria2) args.push(criteriaRange2, criteria2);
  if (criteriaRange3 && criteria3) args.push(criteriaRange3, criteria3);
  return getLocaleAwareFormula('SUMIFS', ...args);
}

// Helper function to create OFFSET formulas with locale awareness
function createOffsetFormula(reference, rows, cols, height, width) {
  const args = [reference, rows, cols];
  if (height !== undefined) args.push(height);
  if (width !== undefined) args.push(width);
  return getLocaleAwareFormula('OFFSET', ...args);
}

// *** TOGGLE SETTINGS ***
// The YoY comparison toggle is controlled by the dropdown in cell B4 of the Summary sheet
// 
// 🔄 MANUAL UPDATES (Google Ads Script limitation):
// Option 1: Change toggle in cell B4 → Run refreshSheet() → Sheet updates
// Option 2: Run toggleYoY() → Automatically switches between views
// 
// No need to modify this script - just use one of the options above!

// Helper function to read toggle value dynamically from sheet
function getShowYoyComparison(summarySheet) {
  debugLog('Reading YoY toggle value from sheet');
  const value = summarySheet.getRange('B4').getValue();
  debugLog('YoY toggle value', { value, result: value !== 'Yes' });
  return value !== 'Yes';
}

// Helper function to read YoY toggle value (inverted logic for consistency with monthly script)
function getHideYoyColumns(summarySheet) {
  debugLog('Reading YoY hide toggle value from sheet');
  const value = summarySheet.getRange('B4').getValue();
  debugLog('YoY hide toggle value', { value, result: value === 'Yes' });
  return value === 'Yes';
}

// Auto-detect locale and currency settings
function getLocaleSettings() {
  debugStart('getLocaleSettings');
  
  try {
    const account = AdsApp.currentAccount();
    const timeZone = account.getTimeZone();
    const currencyCode = account.getCurrencyCode();
    
    debugLog('Account settings', { timeZone, currencyCode });
    
    // Common locale mappings based on currency and timezone
    const localeMap = {
      'USD': {
        locale: 'en-US',
        currencySymbol: '$',
        numberFormat: '#,##0',
        currencyFormat: '$#,##0',
        currencyFormatWithDecimals: '$#,##0.00'
      },
      'EUR': {
        locale: 'en-GB', // Using en-GB for EUR to get proper formatting
        currencySymbol: '€',
        numberFormat: '#,##0',
        currencyFormat: '€#,##0',
        currencyFormatWithDecimals: '€#,##0.00'
      },
      'GBP': {
        locale: 'en-GB',
        currencySymbol: '£',
        numberFormat: '#,##0',
        currencyFormat: '£#,##0',
        currencyFormatWithDecimals: '£#,##0.00'
      },
      'CAD': {
        locale: 'en-CA',
        currencySymbol: 'C$',
        numberFormat: '#,##0',
        currencyFormat: 'C$#,##0',
        currencyFormatWithDecimals: 'C$#,##0.00'
      },
      'AUD': {
        locale: 'en-AU',
        currencySymbol: 'A$',
        numberFormat: '#,##0',
        currencyFormat: 'A$#,##0',
        currencyFormatWithDecimals: 'A$#,##0.00'
      },
      'JPY': {
        locale: 'ja-JP',
        currencySymbol: '¥',
        numberFormat: '#,##0',
        currencyFormat: '¥#,##0',
        currencyFormatWithDecimals: '¥#,##0'  // JPY typically doesn't use decimals
      },
      'BRL': {
        locale: 'pt-BR',
        currencySymbol: 'R$',
        numberFormat: '#,##0',
        currencyFormat: 'R$#,##0',
        currencyFormatWithDecimals: 'R$#,##0.00'
      },
      'MXN': {
        locale: 'es-MX',
        currencySymbol: 'MX$',
        numberFormat: '#,##0',
        currencyFormat: 'MX$#,##0',
        currencyFormatWithDecimals: 'MX$#,##0.00'
      },
      'INR': {
        locale: 'en-IN',
        currencySymbol: '₹',
        numberFormat: '#,##,##0',  // Indian numbering system
        currencyFormat: '₹#,##,##0',
        currencyFormatWithDecimals: '₹#,##,##0.00'
      },
      'SEK': {
        locale: 'sv-SE',
        currencySymbol: 'kr',
        numberFormat: '#,##0',
        currencyFormat: '#,##0 kr',
        currencyFormatWithDecimals: '#,##0.00 kr'
      },
      'NOK': {
        locale: 'nb-NO',
        currencySymbol: 'kr',
        numberFormat: '#,##0',
        currencyFormat: '#,##0 kr',
        currencyFormatWithDecimals: '#,##0.00 kr'
      },
      'DKK': {
        locale: 'da-DK',
        currencySymbol: 'kr',
        numberFormat: '#,##0',
        currencyFormat: '#,##0 kr',
        currencyFormatWithDecimals: '#,##0.00 kr'
      },
      'CHF': {
        locale: 'de-CH',
        currencySymbol: 'CHF',
        numberFormat: '#,##0',
        currencyFormat: 'CHF #,##0',
        currencyFormatWithDecimals: 'CHF #,##0.00'
      },
      'PLN': {
        locale: 'pl-PL',
        currencySymbol: 'zł',
        numberFormat: '#,##0',
        currencyFormat: '#,##0 zł',
        currencyFormatWithDecimals: '#,##0.00 zł'
      },
      'CZK': {
        locale: 'cs-CZ',
        currencySymbol: 'Kč',
        numberFormat: '#,##0',
        currencyFormat: '#,##0 Kč',
        currencyFormatWithDecimals: '#,##0.00 Kč'
      },
      'HUF': {
        locale: 'hu-HU',
        currencySymbol: 'Ft',
        numberFormat: '#,##0',
        currencyFormat: '#,##0 Ft',
        currencyFormatWithDecimals: '#,##0 Ft'  // HUF typically doesn't use decimals
      }
    };
    
    // Default to USD format if currency not found
    const settings = localeMap[currencyCode] || localeMap['USD'];
    
    const result = {
      currencyCode,
      timeZone,
      ...settings
    };
    
    debugLog('Locale settings result', result);
    debugEnd('getLocaleSettings');
    return result;
    
  } catch (error) {
    debugError('Error in getLocaleSettings', error);
    debugEnd('getLocaleSettings');
    // Return default USD settings if error occurs
    return {
      currencyCode: 'USD',
      timeZone: 'America/New_York',
      locale: 'en-US',
      currencySymbol: '$',
      numberFormat: '#,##0',
      currencyFormat: '$#,##0',
      currencyFormatWithDecimals: '$#,##0.00'
    };
  }
}

// Calculate date range from beginning of last year to last complete Sunday (2 years of data)
function getDateRange() {
  debugStart('getDateRange');
  
  try {
    const today = new Date();
    const endDate = new Date(today);
    
    debugLog('Date calculation input', { today: today.toISOString() });
    
    // Calculate last complete Sunday
    const dayOfWeek = today.getDay(); // 0 = Sunday, 1 = Monday, etc.
    
    debugLog('Day of week calculation', { dayOfWeek });
    
    if (dayOfWeek === 0) {
      // If today is Sunday, go back to previous Sunday (7 days ago) since today is not complete
      endDate.setDate(today.getDate() - 7);
      debugLog('Today is Sunday, going back 7 days');
    } else {
      // If today is not Sunday, go back to the most recent Sunday
      endDate.setDate(today.getDate() - dayOfWeek);
      debugLog('Going back to most recent Sunday', { daysBack: dayOfWeek });
    }
    
    // Start from beginning of last year to get year-over-year comparison
    const currentYear = today.getFullYear();
    const startDate = new Date(currentYear - 1, 0, 1); // January 1st of last year
    
    debugLog('Year calculation', { currentYear, startDate: startDate.toISOString() });
    
    const formatDate = (date) => {
      return Utilities.formatDate(date, AdsApp.currentAccount().getTimeZone(), 'yyyy-MM-dd');
    };
    
    const result = {
      start: formatDate(startDate),
      end: formatDate(endDate)
    };
    
    debugLog('Date range result', result);
    debugEnd('getDateRange');
    return result;
    
  } catch (error) {
    debugError('Error in getDateRange', error);
    debugEnd('getDateRange');
    // Return default date range if error occurs
    return {
      start: '2024-01-01',
      end: '2025-01-01'
    };
  }
}

const dateRange = getDateRange();
const START_DATE = dateRange.start;
const END_DATE = dateRange.end;

debugLog('Global date constants', { START_DATE, END_DATE });

// Get locale settings for currency formatting
const LOCALE_SETTINGS = getLocaleSettings();

debugLog('Global locale settings', LOCALE_SETTINGS);

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
WHERE segments.date BETWEEN "${START_DATE}" AND "${END_DATE}"
  AND campaign.status != "REMOVED"
ORDER BY segments.week DESC, metrics.cost_micros DESC
`;

debugLog('GAQL Query', { query: GAQL_QUERY });

function main() {
  debugStart('main');
  
  try {
    // You can modify this line to run different actions:
    // - "create" (default) - creates new sheet
    // - "toggle" - switches between YoY and current year views
    // - "refresh" - refreshes based on current toggle
    // - "test" - tests toggle functionality
    
    const action = "create"; // Change this to test different functions
    
    debugLog('Main action', { action });
    
    switch(action) {
      case "toggle":
        debugLog('Running toggle function');
        Logger.log('🔄 Running toggle function...');
        toggleYoY();
        break;
      case "refresh":
        debugLog('Running refresh function');
        Logger.log('🔄 Running refresh function...');
        refreshSheet();
        break;
      case "test":
        debugLog('Running test function');
        Logger.log('🧪 Running test function...');
        testToggle();
        break;
      case "create":
      default:
        debugLog('Running full report generation');
        Logger.log('🚀 Running full report generation...');
        generateWeeklyAdsReport();
        break;
    }
    
    debugEnd('main');
    
  } catch (error) {
    debugError('Error in main function', error);
    Logger.log('❌ MAIN FUNCTION ERROR: ' + error.toString());
  }
}

function generateWeeklyAdsReport() {
  debugStart('generateWeeklyAdsReport');
  
  try {
    Logger.log('Thanks for using Weekly Ad Monitor script by Andrey Kisselev (c) 2025. Version: 1.8.8.');
    
    debugLog('Brand campaigns configuration', { brandCampaigns: BRAND_CAMPAIGNS, count: BRAND_CAMPAIGNS.length });
    
    if (BRAND_CAMPAIGNS.length > 0) {
      Logger.log(`🏷️ Weekly Ads Monitor - Brand campaigns to identify: ${BRAND_CAMPAIGNS.join(', ')}`);
    } else {
      Logger.log('⚠️ Weekly Ads Monitor - No brand campaigns specified in BRAND_CAMPAIGNS array');
    }
    
    let spreadsheet;
    
    debugLog('Sheet URL configuration', { sheetUrl: SHEET_URL, isEmpty: SHEET_URL === '' });
    
    if (SHEET_URL === '') {
      // Create a new spreadsheet if no URL is provided
      const sheetName = 'Weekly Ads Monitor v1.8.8 - ' + 
        Utilities.formatDate(new Date(), AdsApp.currentAccount().getTimeZone(), 'yyyy-MM-dd');
      debugLog('Creating new spreadsheet', { sheetName });
      
      spreadsheet = SpreadsheetApp.create(sheetName);
      Logger.log('✅ NEW SPREADSHEET CREATED: ' + spreadsheet.getUrl());
      debugLog('New spreadsheet created', { url: spreadsheet.getUrl() });
    } else {
      // Open existing spreadsheet
      debugLog('Opening existing spreadsheet', { url: SHEET_URL });
      spreadsheet = SpreadsheetApp.openByUrl(SHEET_URL);
      Logger.log('✅ Opened existing Weekly Ads Monitor spreadsheet');
      debugLog('Existing spreadsheet opened', { url: spreadsheet.getUrl() });
    }
    
    // Get or create the "Raw" tab (renamed from Campaign Trends)
    let rawSheet = spreadsheet.getSheetByName('Raw');
    if (!rawSheet) {
      rawSheet = spreadsheet.insertSheet('Raw');
      Logger.log('✅ Created new "Raw" tab');
      debugLog('Raw sheet created');
    } else {
      Logger.log('✅ Using existing "Raw" tab');
      debugLog('Using existing Raw sheet');
    }
    
    // Clear existing data
    debugLog('Clearing existing Raw sheet data');
    rawSheet.clear();
    
    // Execute the GAQL query
    debugLog('Executing GAQL query', { query: GAQL_QUERY });
    
    let rows;
    try {
      rows = AdsApp.search(GAQL_QUERY);
      debugLog('GAQL query executed successfully');
      // Query executed successfully, processing results...
    } catch (queryError) {
      debugError('GAQL Query Error', queryError);
      Logger.log('❌ GAQL Query Error: ' + queryError.toString());
      throw queryError;
    }
    
    // Process data
    debugLog('Starting data processing');
    const rawData = processData(rows);
    
    debugLog('Data processing completed', { recordCount: rawData.length });
    
    if (rawData.length === 0) {
      Logger.log('⚠️ No data found for the specified date range.');
      debugLog('No data found for date range', { startDate: START_DATE, endDate: END_DATE });
      return;
    }
    
    // Found weekly campaign records
    debugLog('Weekly campaign records found', { count: rawData.length });
    

    
    // Create headers for raw data - 13 columns with expanded campaign performance metrics including brand flag
    const rawHeaders = [
      'Week Start',
      'Week End',
      'Campaign Name',
      'Campaign Type',
      'Is Brand Campaign',
      'Impressions',
      'Clicks',
      'Cost',
      'Cost Per Click',
      'Conversions',
      'Conversion Value',
      'Conversion Value / Cost',
      'Cost Per Conversion'
    ];
    
    // Combine headers and data for raw sheet
    const allRawData = [rawHeaders, ...rawData];
    
    // Write to raw sheet
    Logger.log('📤 Writing raw data to spreadsheet...');
    debugLog('Writing raw data to sheet', { rowCount: allRawData.length, columnCount: rawHeaders.length });
    
    const rawRange = rawSheet.getRange(1, 1, allRawData.length, rawHeaders.length);
    rawRange.setValues(allRawData);
    
    // Format raw sheet headers
    debugLog('Formatting raw sheet headers');
    const rawHeaderRange = rawSheet.getRange(1, 1, 1, rawHeaders.length);
    rawHeaderRange.setFontWeight('bold');
    rawHeaderRange.setBackground('#4285f4');
    rawHeaderRange.setFontColor('#ffffff');
    
    // Auto-resize columns for raw sheet
    debugLog('Auto-resizing raw sheet columns');
    for (let i = 1; i <= rawHeaders.length; i++) {
      rawSheet.autoResizeColumn(i);
    }
    
    // Create aggregated weekly summary tab
    Logger.log('📊 Creating Weekly Ads Monitor dashboard with calendar-based YoY comparison...');
    debugLog('Creating weekly summary dashboard');
    createWeeklySummary(spreadsheet, rawData);
    
    Logger.log('🎉 SUCCESS! Weekly Ads Monitor v1.8.8 completed');
    Logger.log(`✅ Generated ${rawData.length} weekly campaign records with calendar-based YoY comparison and WoW for all metrics (with reordered WoW columns)`);
    Logger.log(`💰 Currency formatting applied: ${LOCALE_SETTINGS.currencyCode} (${LOCALE_SETTINGS.currencySymbol})`);
    Logger.log('🔗 Weekly Ads Monitor URL: ' + spreadsheet.getUrl());
    
    debugLog('Weekly Ads Monitor completed successfully', { 
      recordCount: rawData.length, 
      currencyCode: LOCALE_SETTINGS.currencyCode,
      spreadsheetUrl: spreadsheet.getUrl() 
    });
    
    debugEnd('generateWeeklyAdsReport');
    
  } catch (error) {
    debugError('Weekly Ads Monitor Error', error);
    Logger.log('❌ WEEKLY ADS MONITOR ERROR: ' + error.toString());
    debugEnd('generateWeeklyAdsReport');
  }
}

function processData(rows) {
  debugStart('processData');
  
  let data = [];
  let count = 0;
  let brandCampaignCount = 0;
  
  Logger.log('⚙️ Processing weekly campaign performance data...');
  debugLog('Starting data processing', { brandCampaigns: BRAND_CAMPAIGNS });
  
  while (rows.hasNext()) {
    try {
      let row = rows.next();
      count++;
      
      debugLog('Processing row', { rowNumber: count });
      
      const campaign = row.campaign || {};
      const segments = row.segments || {};
      const metrics = row.metrics || {};
      
      // Get campaign details
      let campaignName = campaign.name || '';
      let campaignType = campaign.advertisingChannelType || '';
      
      // Check if this campaign name is in the brand campaigns list
      let isBrand = BRAND_CAMPAIGNS.includes(campaignName);
      if (isBrand) brandCampaignCount++;
      
      debugLog('Campaign details', { 
        campaignName, 
        campaignType, 
        isBrand, 
        brandCampaignCount 
      });
      
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
          
          debugLog('Week date calculation', { weekStart, weekEnd });
        } catch (e) {
          debugError('Week end date calculation error', e);
          weekEnd = weekStart; // fallback to start date if calculation fails
        }
      }
      
      // Convert cost from micros to currency and calculate derived metrics
      let cost = Math.round((costMicros / 1000000) * 100) / 100;
      conversionValue = Math.round(conversionValue * 100) / 100;
      
      debugLog('Metrics calculation', { 
        costMicros, 
        cost, 
        conversionValue, 
        clicks, 
        conversions 
      });
      
      // Include all campaigns regardless of spend level
      
      // Calculate cost per click
      let costPerClick = clicks > 0 ? Math.round((cost / clicks) * 100) / 100 : 0;
      
      // Calculate cost per conversion
      let costPerConversion = conversions > 0 ? Math.round((cost / conversions) * 100) / 100 : 0;
      
      // Calculate conv. value / cost (ROAS)
      let roas = cost > 0 ? Math.round((conversionValue / cost) * 100) / 100 : 0;
      
      debugLog('Derived metrics', { costPerClick, costPerConversion, roas });
      
      // Create row - 13 columns with campaign performance metrics including brand flag
      let newRow = [
        weekStart,                    // Week Start
        weekEnd,                      // Week End
        campaignName,                 // Campaign Name
        campaignType,                 // Campaign Type
        isBrand,                      // Is Brand Campaign
        impressions,                  // Impressions
        clicks,                       // Clicks
        cost,                         // Cost
        costPerClick,                 // Cost Per Click
        conversions,                  // Conversions
        conversionValue,              // Conversion Value
        roas,                         // Conversion Value / Cost
        costPerConversion             // Cost Per Conversion
      ];
      
      data.push(newRow);
      
      // Processing records...
      
    } catch (e) {
      debugError('Error processing row', e);
      Logger.log(`❌ Error processing row: ${e}`);
    }
  }
  
  // Processing complete
  debugLog('Data processing summary', { 
    totalRecords: count, 
    brandCampaignCount, 
    processedRecords: data.length 
  });
  
  debugEnd('processData');
  return data;
}

// NEW FUNCTION: Find the closest corresponding week in previous year
function findClosestPreviousYearWeek(currentYearWeekStart, availablePreviousWeeks, rawData) {
  debugStart('findClosestPreviousYearWeek');
  
  try {
    const currentDate = new Date(currentYearWeekStart);
    
    debugLog('Finding closest previous year week', { 
      currentYearWeekStart, 
      availablePreviousWeeksCount: availablePreviousWeeks.length 
    });
    
    // Calculate the exact same calendar date in previous year
    const exactPreviousDate = new Date(currentDate);
    exactPreviousDate.setFullYear(currentDate.getFullYear() - 1);
    const exactPreviousDateStr = Utilities.formatDate(exactPreviousDate, AdsApp.currentAccount().getTimeZone(), 'yyyy-MM-dd');
    
    debugLog('Exact previous year date calculation', { 
      currentDate: currentDate.toISOString(), 
      exactPreviousDate: exactPreviousDate.toISOString(),
      exactPreviousDateStr 
    });
    
    // If exact match exists, use it
    if (availablePreviousWeeks.includes(exactPreviousDateStr)) {
      debugLog('Exact match found', { exactPreviousDateStr });
      debugEnd('findClosestPreviousYearWeek');
      return exactPreviousDateStr;
    }
    
    // Otherwise, find the closest week (within 7 days)
    let closestWeek = null;
    let smallestDiff = Infinity;
    let closestWeekData = null;
    
    debugLog('Searching for closest week within 7 days');
    
    availablePreviousWeeks.forEach(weekStr => {
      const weekDate = new Date(weekStr);
      const timeDiff = Math.abs(weekDate.getTime() - exactPreviousDate.getTime());
      const daysDiff = timeDiff / (1000 * 60 * 60 * 24);
      
      debugLog('Checking week', { weekStr, daysDiff, smallestDiff });
      
      // Only consider weeks within 7 days of the target
      if (daysDiff <= 7 && daysDiff < smallestDiff) {
        smallestDiff = daysDiff;
        closestWeek = weekStr;
        
        // Get data volume for this week to verify it's reasonable
        const weekData = rawData.filter(row => row[0] === weekStr);
        const weekSpend = weekData.reduce((sum, row) => sum + (row[7] || 0), 0);
        closestWeekData = { records: weekData.length, spend: weekSpend };
        
        debugLog('New closest week found', { 
          weekStr, 
          daysDiff, 
          weekData: closestWeekData 
        });
      }
    });
    
    // Found closest match or no suitable match
    debugLog('Closest week search result', { 
      closestWeek, 
      smallestDiff, 
      closestWeekData 
    });
    
    debugEnd('findClosestPreviousYearWeek');
    return closestWeek;
  } catch (e) {
    debugError('Error finding closest previous year week', e);
    Logger.log(`❌ Error finding closest previous year week for ${currentYearWeekStart}: ${e}`);
    debugEnd('findClosestPreviousYearWeek');
    return null;
  }
}

function createWeeklySummary(spreadsheet, rawData) {
  debugStart('createWeeklySummary');
  
  try {
    debugLog('Creating weekly summary', { rawDataCount: rawData.length });
    
    // Get or create the "Weekly Summary" tab
    let summarySheet = spreadsheet.getSheetByName('Weekly Summary');
    if (!summarySheet) {
      summarySheet = spreadsheet.insertSheet('Weekly Summary');
      debugLog('Created new Weekly Summary sheet');
    } else {
      debugLog('Using existing Weekly Summary sheet');
    }
    // Created or using existing "Weekly Summary" tab
    
    debugLog('Clearing Weekly Summary sheet');
    summarySheet.clear();
    
    // Unfreeze any frozen rows/columns to avoid merge conflicts
    summarySheet.setFrozenRows(0);
    summarySheet.setFrozenColumns(0);
    
    // Get unique campaign types for the filter dropdown
    const allCampaignTypes = [...new Set(rawData.map(row => row[3]))].sort(); // Campaign Types
    const campaignTypeOptions = ['All', ...allCampaignTypes];
    
    debugLog('Campaign types found', { 
      allCampaignTypes, 
      campaignTypeOptions 
    });
    
    // Data distribution analysis completed
    
    // Add app title in first row
    summarySheet.getRange('A1').setValue('Weekly Ads Monitor v1.8.7');
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
    
    // Add YoY toggle filter controls
    summarySheet.getRange('A4').setValue('Hide YoY Columns?');
    summarySheet.getRange('A4').setFontWeight('bold');
    summarySheet.getRange('A4').setBackground('#f0f0f0');
    
    // Add Exclude Brand Campaigns toggle controls (row 3)
    summarySheet.getRange('A3').setValue('Exclude Brand Campaigns?');
    summarySheet.getRange('A3').setFontWeight('bold');
    summarySheet.getRange('A3').setBackground('#f0f0f0');
    
    // Add Hide YoY Columns toggle controls (row 4)
    summarySheet.getRange('A4').setValue('Hide YoY Columns?');
    summarySheet.getRange('A4').setFontWeight('bold');
    summarySheet.getRange('A4').setBackground('#f0f0f0');
    
    // Add Hide WoW Columns toggle controls (row 5)
    summarySheet.getRange('A5').setValue('Hide WoW Columns?');
    summarySheet.getRange('A5').setFontWeight('bold');
    summarySheet.getRange('A5').setBackground('#f0f0f0');
    
    // Add Start Week from Sunday toggle controls (row 6)
    summarySheet.getRange('A6').setValue('Start Week from Sunday?');
    summarySheet.getRange('A6').setFontWeight('bold');
    summarySheet.getRange('A6').setBackground('#f0f0f0');
    
    // Add client name and currency info in row 7
    const clientName = AdsApp.currentAccount().getName();
    summarySheet.getRange('A7').setValue(`Acc.: ${clientName} (${LOCALE_SETTINGS.currencyCode})`);
    summarySheet.getRange('A7').setFontWeight('bold');
    summarySheet.getRange('A7').setFontSize(12);
    summarySheet.getRange('A7').setBackground('#e8f0fe');
    summarySheet.getRange('A7').setFontColor('#1a73e8');
    
    // Merge cells A7 and B7 for the account name
    summarySheet.getRange('A7:B7').merge();
    
    // Set up dropdown for Exclude Brand Campaigns toggle (row 3)
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
    
    // Set up dropdown for Hide YoY toggle (row 4)
    const yoyFilterCell = summarySheet.getRange('B4');
    
    // Read the current YoY filter selection BEFORE setting up dropdown
    const currentYoySelection = yoyFilterCell.getValue();
    
    // Set up dropdown validation for YoY filter
    const yoyFilterRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['No', 'Yes'])
      .setAllowInvalid(false)
      .build();
    yoyFilterCell.setDataValidation(yoyFilterRule);
    yoyFilterCell.setBackground('#ffffff');
    yoyFilterCell.setBorder(true, true, true, true, true, true);
    
    // Only set default if cell is empty or invalid
    if (!currentYoySelection || !['No', 'Yes'].includes(currentYoySelection)) {
      yoyFilterCell.setValue('No');
    }
    
    // Set up dropdown for Hide WoW toggle (row 5)
    const hideWoWFilterCell = summarySheet.getRange('B5');
    
    // Read the current Hide WoW filter selection BEFORE setting up dropdown
    const currentHideWoWSelection = hideWoWFilterCell.getValue();
    
    // Set up dropdown validation for Hide WoW filter
    const hideWoWFilterRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['No', 'Yes'])
      .setAllowInvalid(false)
      .build();
    hideWoWFilterCell.setDataValidation(hideWoWFilterRule);
    hideWoWFilterCell.setBackground('#ffffff');
    hideWoWFilterCell.setBorder(true, true, true, true, true, true);
    
    // Only set default if cell is empty or invalid
    if (!currentHideWoWSelection || !['No', 'Yes'].includes(currentHideWoWSelection)) {
      hideWoWFilterCell.setValue('No');
    }
    
    // Set up dropdown for Start Week from Sunday toggle (row 6)
    const startWeekSundayFilterCell = summarySheet.getRange('B6');
    
    // Read the current Start Week from Sunday filter selection BEFORE setting up dropdown
    const currentStartWeekSundaySelection = startWeekSundayFilterCell.getValue();
    
    // Set up dropdown validation for Start Week from Sunday filter
    const startWeekSundayFilterRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['No', 'Yes'])
      .setAllowInvalid(false)
      .build();
    startWeekSundayFilterCell.setDataValidation(startWeekSundayFilterRule);
    startWeekSundayFilterCell.setBackground('#ffffff');
    startWeekSundayFilterCell.setBorder(true, true, true, true, true, true);
    
    // Only set default if cell is empty or invalid
    if (!currentStartWeekSundaySelection || !['No', 'Yes'].includes(currentStartWeekSundaySelection)) {
      startWeekSundayFilterCell.setValue('No');
    }
    
    // Add instructions
    summarySheet.getRange('D2').setValue('🔄 Dynamic Filters: Select options from dropdowns → Data updates automatically!');
    summarySheet.getRange('D2').setFontStyle('italic');
    summarySheet.getRange('D2').setFontSize(9);
    summarySheet.getRange('D2').setFontColor('#0066cc');
    
    summarySheet.getRange('D3').setValue('🚫 Brand Filter: "Yes" excludes brand campaigns from analysis');
    summarySheet.getRange('D3').setFontStyle('italic');
    summarySheet.getRange('D3').setFontSize(9);
    summarySheet.getRange('D3').setFontColor('#e74c3c');
    
    summarySheet.getRange('D4').setValue('👁️ YoY Toggle: "Yes" hides year-over-year comparison columns');
    summarySheet.getRange('D4').setFontStyle('italic');
    summarySheet.getRange('D4').setFontSize(9);
    summarySheet.getRange('D4').setFontColor('#9c27b0');
    
    summarySheet.getRange('D5').setValue('📊 WoW Toggle: "Yes" hides week-over-week comparison columns');
    summarySheet.getRange('D5').setFontStyle('italic');
    summarySheet.getRange('D5').setFontSize(9);
    summarySheet.getRange('D5').setFontColor('#ff6b35');
    
    summarySheet.getRange('D6').setValue('📅 Week Start: "Yes" starts weeks from Sunday instead of Monday');
    summarySheet.getRange('D6').setFontStyle('italic');
    summarySheet.getRange('D6').setFontSize(9);
    summarySheet.getRange('D6').setFontColor('#4caf50');
    
    // Add "Coming Soon" label in C6
    summarySheet.getRange('C6').setValue('Coming Soon');
    summarySheet.getRange('C6').setFontColor('#666666');
    summarySheet.getRange('C6').setFontSize(10);
    summarySheet.getRange('C6').setFontStyle('italic');
    
    summarySheet.getRange('D7').setValue('📅 Calendar-based YoY: Aug 4-10, 2025 compares to Aug 4-10, 2024');
    summarySheet.getRange('D7').setFontStyle('italic');
    summarySheet.getRange('D7').setFontSize(9);
    summarySheet.getRange('D7').setFontColor('#666666');
    
    // Read the final selected filter values
    const selectedCampaignType = filterCell.getValue() || 'All';
    const excludeBrandCampaigns = brandFilterCell.getValue() === 'Yes';
    const hideWoWColumns = hideWoWFilterCell.getValue() === 'Yes';
    const startWeekFromSunday = startWeekSundayFilterCell.getValue() === 'Yes';
    
    const hideYoyColumns = getHideYoyColumns(summarySheet);
    
    debugLog('Filter values', { 
      selectedCampaignType, 
      excludeBrandCampaigns, 
      hideWoWColumns, 
      startWeekFromSunday, 
      hideYoyColumns 
    });
    
    // Dynamic filtering enabled
    
    // Get current and previous year - always compare 2025 vs 2024
    const currentYear = new Date().getFullYear(); // 2025
    const previousYear = currentYear - 1; // 2024
    
    // *** NEW CALENDAR-BASED LOGIC ***
    // Get all unique weeks from raw data for both years
    const allWeekDates = [...new Set(rawData.map(row => row[0]))].sort();
    const currentYearWeeks = allWeekDates.filter(week => week.includes(currentYear.toString()));
    const previousYearWeeks = allWeekDates.filter(week => week.includes(previousYear.toString()));
    
    debugLog('Week analysis', { 
      allWeekDates, 
      currentYearWeeks, 
      previousYearWeeks,
      currentYear,
      previousYear
    });
    
    // Create a set of available previous year weeks for quick lookup
    const availablePreviousWeeks = new Set(previousYearWeeks);
    
    // Calculate corresponding previous year dates for each current year week
    // Only include pairs where both current and previous year data exists
    const weekPairs = [];
    
    debugLog('Finding week pairs for calendar-based comparison');
    
    currentYearWeeks.forEach(currentWeek => {
      const previousWeek = findClosestPreviousYearWeek(currentWeek, previousYearWeeks, rawData);
      
      if (previousWeek) {
        weekPairs.push({
          current: currentWeek,
          previous: previousWeek
        });
        debugLog('Week pair found', { current: currentWeek, previous: previousWeek });
      } else {
        Logger.log(`⚠️ No suitable previous year match found for ${currentWeek}`);
        debugLog('No previous year match found', { currentWeek });
        // Still add the pair but with null previous week
        weekPairs.push({
          current: currentWeek,
          previous: null
        });
      }
    });
    
    debugLog('Week pairs completed', { weekPairsCount: weekPairs.length, weekPairs });
    
    // Calendar-based comparison completed
    
    // Create header structure with conditional formulas for YoY toggle
    // Always create full headers structure - hide/show columns based on YoY toggle
    const mainHeaders = [
      'Week Number', 'Week Start', 
      'Clicks', '', '', '', 
      'Average CPC', '', '', '', 
      'Cost', '', '', '', 
      'Conversions', '', '', '', 
      'Cost / Conv', '', '', '', 
      'Conversion Value', '', '', '', 
      'Conversion Value / Cost', '', '', ''
    ];
    const headerRange = summarySheet.getRange(9, 1, 1, mainHeaders.length);
    headerRange.setValues([mainHeaders]);
    
    // Year subheaders (row 10) - Create conditional formulas that work with the current locale
    const previousYearFormula = createIfFormula('$B$4="Yes"', '""', previousYear);
    const yoyFormula = createIfFormula('$B$4="Yes"', '""', '"YoY"');
    const wowFormula = createIfFormula('$B$5="Yes"', '""', '"WoW"');
    
    const yearHeaders = [
      '', '', 
      currentYear, wowFormula, previousYearFormula, yoyFormula, 
      currentYear, wowFormula, previousYearFormula, yoyFormula, 
      currentYear, wowFormula, previousYearFormula, yoyFormula, 
      currentYear, wowFormula, previousYearFormula, yoyFormula, 
      currentYear, wowFormula, previousYearFormula, yoyFormula, 
      currentYear, wowFormula, previousYearFormula, yoyFormula, 
      currentYear, wowFormula, previousYearFormula, yoyFormula
    ];
    const yearHeaderRange = summarySheet.getRange(10, 1, 1, yearHeaders.length);
    yearHeaderRange.setValues([yearHeaders]);
    
    // v1.7.1 FIX: Format year headers as plain numbers (not currency) for year columns only
    // This ensures years like 2025 and 2024 display correctly in row 10
    // Previous version incorrectly applied formatting to row 9 (main headers) instead of row 10 (year headers)
    summarySheet.getRange('C10:E10').setNumberFormat('0'); // Clicks years (C=2025, E=2024)
    summarySheet.getRange('G10:I10').setNumberFormat('0'); // CPC years (G=2025, I=2024)
    summarySheet.getRange('K10:M10').setNumberFormat('0'); // Cost years (K=2025, M=2024)
    summarySheet.getRange('O10:Q10').setNumberFormat('0'); // Conversions years (O=2025, Q=2024)
    summarySheet.getRange('S10:U10').setNumberFormat('0'); // Cost / Conv years (S=2025, U=2024)
    summarySheet.getRange('W10:Y10').setNumberFormat('0'); // Conversion Value years (W=2025, Y=2024)
    summarySheet.getRange('AA10:AC10').setNumberFormat('0'); // Conversion Value / Cost years (AA=2025, AC=2024)
    
    // Merge cells for main headers based on toggle setting
    try {
      summarySheet.getRange('A9:A10').merge(); // Week Number
      summarySheet.getRange('B9:B10').merge(); // Week Start
      
      if (getShowYoyComparison(summarySheet)) {
        // Full YoY comparison merging (4 columns per metric for all metrics including WoW)
        summarySheet.getRange('C7:F7').merge(); // Clicks (including WoW)
        summarySheet.getRange('G7:J7').merge(); // Average CPC (including WoW)
        summarySheet.getRange('K7:N7').merge(); // Cost (including WoW)
        summarySheet.getRange('O7:R7').merge(); // Conversions (including WoW)
        summarySheet.getRange('S7:V7').merge(); // Cost / Conv (including WoW)
        summarySheet.getRange('W7:Z7').merge(); // Conversion Value (including WoW)
        summarySheet.getRange('AA7:AD7').merge(); // Conversion Value / Cost (including WoW)
      } else {
        // Current year only merging (1 column per metric)
        summarySheet.getRange('C7:C8').merge(); // Clicks
        summarySheet.getRange('D7:D8').merge(); // Average CPC
        summarySheet.getRange('E7:E8').merge(); // Cost
        summarySheet.getRange('F7:F8').merge(); // Conversions
        summarySheet.getRange('G7:G8').merge(); // Cost / Conv
        summarySheet.getRange('H7:H8').merge(); // Conversion Value
        summarySheet.getRange('I7:I8').merge(); // Conversion Value / Cost
      }
      
          Logger.log('✅ Header cells merged successfully');
  } catch (mergeError) {
    Logger.log('⚠️ Warning: Could not merge header cells: ' + mergeError.toString());
    Logger.log('📋 Headers will still function correctly without merging');
  }
  
  // Apply conditional formatting to hide YoY columns when toggle is "Yes"
  // YoY columns to hide: Previous Year and YoY only (Current, WoW, Previous, YoY)
  const yoyColumnsToToggle = [
    5, 6,    // Clicks: E=Previous Year, F=YoY (NOT C=Current, D=WoW)
    9, 10,   // Average CPC: I=Previous Year, J=YoY (NOT G=Current, H=WoW)
    13, 14,  // Cost: M=Previous Year, N=YoY (NOT K=Current, L=WoW)
    17, 18,  // Conversions: Q=Previous Year, R=YoY (NOT O=Current, P=WoW)
    21, 22,  // Cost / Conv: U=Previous Year, V=YoY (NOT S=Current, T=WoW)
    25, 26,  // Conversion Value: Y=Previous Year, Z=YoY (NOT W=Current, X=WoW)
    29, 30   // Conversion Value / Cost: AC=Previous Year, AD=YoY (NOT AA=Current, AB=WoW)
  ];
  
  // Apply conditional formatting to make YoY data columns appear hidden when B4="Yes"
  // Headers are now handled by conditional formulas, so we only need to hide data rows
  yoyColumnsToToggle.forEach(colIndex => {
    const dataRange = summarySheet.getRange(11, colIndex, 1000, 1); // Data rows starting from row 11
    
    // Create conditional formatting rule to hide data text when B4="Yes"
    const hideDataRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$B$4="Yes"')
      .setFontColor('#ffffff') // White text (invisible on white background)
      .setBackground('#ffffff') // White background
      .setRanges([dataRange])
      .build();
    
    const existingRules = summarySheet.getConditionalFormatRules();
    summarySheet.setConditionalFormatRules(existingRules.concat([hideDataRule]));
  });
  
  Logger.log('✅ YoY toggle conditional formatting applied');
  
  // Apply conditional formatting to hide WoW columns when toggle is "Yes"
  // WoW columns to hide: WoW only (Current, WoW, Previous, YoY)
  const wowColumnsToToggle = [
    4,    // Clicks: D=WoW
    8,    // Average CPC: H=WoW
    12,   // Cost: L=WoW
    16,   // Conversions: P=WoW
    20,   // Cost / Conv: T=WoW
    24,   // Conversion Value: X=WoW
    28    // Conversion Value / Cost: AB=WoW
  ];
  
  // Apply conditional formatting to make WoW data columns appear hidden when B5="Yes"
  wowColumnsToToggle.forEach(colIndex => {
    const dataRange = summarySheet.getRange(11, colIndex, 1000, 1); // Data rows starting from row 11
    
    // Create conditional formatting rule to hide data text when B5="Yes"
    const hideDataRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$B$5="Yes"')
      .setFontColor('#ffffff') // White text (invisible on white background)
      .setBackground('#ffffff') // White background
      .setRanges([dataRange])
      .build();
    
    const existingRules = summarySheet.getConditionalFormatRules();
    summarySheet.setConditionalFormatRules(existingRules.concat([hideDataRule]));
  });
  
  Logger.log('✅ WoW toggle conditional formatting applied');
    
    // Create dynamic formulas that reference the filter and raw data
    Logger.log('📤 Creating calendar-based year-over-year dynamic formulas...');
    debugLog('Creating dynamic formulas');
    
    // Build formula-based data rows that update automatically
    const formulaData = [];
    
    // Use the actual number of current year weeks we have data for
    const maxWeeks = weekPairs.length;
    
    debugLog('Formula generation setup', { maxWeeks, weekPairsCount: weekPairs.length });
    
    for (let i = 0; i < maxWeeks; i++) {
      const weekNumber = i + 1;
      const currentWeek = weekPairs[i]?.current || null;
      const previousWeek = weekPairs[i]?.previous || null;
      
      // Calculate rowNum first - needed for all formulas
      const rowNum = i + 11; // Starting at row 11 now (after app title + 4 filter rows + client name + 2 header rows)
      
      // Use current year week for display
      const displayWeek = currentWeek || '';
      const weekStartQuoted = currentWeek ? `"${currentWeek}"` : '""';
      const previousWeekQuoted = previousWeek ? `"${previousWeek}"` : '""';
      
      // Formula generation completed
      
      // Current year formulas (only if current week exists) - now with brand campaign filter and YoY toggle
      const clicksCurrentFormula = currentWeek ? 
        createComplexSumifsFormula('Raw!G:G', 'Raw!A:A', weekStartQuoted, true, 'All', 'Raw!E:E', 'Raw!D:D') : 
        '0';
      const costCurrentFormula = currentWeek ? 
        createComplexSumifsFormula('Raw!H:H', 'Raw!A:A', weekStartQuoted, true, 'All', 'Raw!E:E', 'Raw!D:D') : 
        '0';
      const conversionsCurrentFormula = currentWeek ? 
        createComplexSumifsFormula('Raw!J:J', 'Raw!A:A', weekStartQuoted, true, 'All', 'Raw!E:E', 'Raw!D:D') : 
        '0';
      const conversionValueCurrentFormula = currentWeek ? 
        createComplexSumifsFormula('Raw!K:K', 'Raw!A:A', weekStartQuoted, true, 'All', 'Raw!E:E', 'Raw!D:D') : 
        '0';
      
      // *** UPDATED: Previous year formulas using calendar-based dates ***
      const clicksPreviousFormula = previousWeek ? 
        createComplexSumifsFormula('Raw!G:G', 'Raw!A:A', previousWeekQuoted, true, 'All', 'Raw!E:E', 'Raw!D:D') : 
        '0';
      const costPreviousFormula = previousWeek ? 
        createComplexSumifsFormula('Raw!H:H', 'Raw!A:A', previousWeekQuoted, true, 'All', 'Raw!E:E', 'Raw!D:D') : 
        '0';
      const conversionsPreviousFormula = previousWeek ? 
        createComplexSumifsFormula('Raw!J:J', 'Raw!A:A', previousWeekQuoted, true, 'All', 'Raw!E:E', 'Raw!D:D') : 
        '0';
      const conversionValuePreviousFormula = previousWeek ? 
        createComplexSumifsFormula('Raw!K:K', 'Raw!A:A', previousWeekQuoted, true, 'All', 'Raw!E:E', 'Raw!D:D') : 
        '0';
      
      // Current year formulas for cost per conversion
      const costPerConvCurrentFormula = currentWeek ? 
        createIfFormula(`O${rowNum}>0`, `K${rowNum}/O${rowNum}`, '0') : 
        '0';
      
      // Previous year formulas for cost per conversion
      const costPerConvPreviousFormula = previousWeek ? 
        createIfFormula(`Q${rowNum}>0`, `M${rowNum}/Q${rowNum}`, '0') : 
        '0';
      

      
      // Calculate Average CPC, ROAS, and Index YoY dynamically
      const avgCPCCurrentFormula = createIfFormula(`C${rowNum}>0`, `K${rowNum}/C${rowNum}`, '0');
      const avgCPCPreviousFormula = createIfFormula(`D${rowNum}>0`, `L${rowNum}/D${rowNum}`, '0');
      
      // Conv. Value / Cost formulas (ROAS using click time conversions)
      const convValueCostCurrentFormula = createIfFormula(`K${rowNum}>0`, `W${rowNum}/K${rowNum}`, '0');
      const convValueCostPreviousFormula = createIfFormula(`M${rowNum}>0`, `Y${rowNum}/M${rowNum}`, '0');
      
      // Week-over-week formulas for all metrics (Current week vs Previous week)
      const clicksWoWFormula = createIfFormula(`${rowNum}>11`, createIfFormula(`C${rowNum-1}>0`, `(C${rowNum}/C${rowNum-1})*100`, '0'), '0');
      const cpcWoWFormula = createIfFormula(`${rowNum}>11`, createIfFormula(`G${rowNum-1}>0`, `(G${rowNum}/G${rowNum-1})*100`, '0'), '0');
      const costWoWFormula = createIfFormula(`${rowNum}>11`, createIfFormula(`K${rowNum-1}>0`, `(K${rowNum}/K${rowNum-1})*100`, '0'), '0');
      const conversionsWoWFormula = createIfFormula(`${rowNum}>11`, createIfFormula(`O${rowNum-1}>0`, `(O${rowNum}/O${rowNum-1})*100`, '0'), '0');
      const costPerConvWoWFormula = createIfFormula(`${rowNum}>11`, createIfFormula(`S${rowNum-1}>0`, `(S${rowNum}/S${rowNum-1})*100`, '0'), '0');
      const conversionValueWoWFormula = createIfFormula(`${rowNum}>11`, createIfFormula(`W${rowNum-1}>0`, `(W${rowNum}/W${rowNum-1})*100`, '0'), '0');
      const convValueCostWoWFormula = createIfFormula(`${rowNum}>11`, createIfFormula(`AA${rowNum-1}>0`, `(AA${rowNum}/AA${rowNum-1})*100`, '0'), '0');
      
      // Index YoY formulas (Current Year / Previous Year * 100)
      const clicksIndexFormula = createIfFormula(`E${rowNum}>0`, `(C${rowNum}/E${rowNum})*100`, '0');
      const cpcIndexFormula = createIfFormula(`I${rowNum}>0`, `(G${rowNum}/I${rowNum})*100`, '0');
      const costIndexFormula = createIfFormula(`M${rowNum}>0`, `(K${rowNum}/M${rowNum})*100`, '0');
      const conversionsIndexFormula = createIfFormula(`Q${rowNum}>0`, `(O${rowNum}/Q${rowNum})*100`, '0');
      const costPerConvIndexFormula = createIfFormula(`U${rowNum}>0`, `(S${rowNum}/U${rowNum})*100`, '0');
      const conversionValueIndexFormula = createIfFormula(`Y${rowNum}>0`, `(W${rowNum}/Y${rowNum})*100`, '0');
      const convValueCostIndexFormula = createIfFormula(`AC${rowNum}>0`, `(AA${rowNum}/AC${rowNum})*100`, '0');
      
      // Always create full YoY comparison formula row (38 columns) with conditional formulas
      const formulaRow = [
        weekNumber,                              // A: Week Number
        displayWeek,                            // B: Week Start
        clicksCurrentFormula,                   // C: Clicks Current Year (2025)
        clicksWoWFormula,                       // D: Clicks Week-over-Week
        clicksPreviousFormula,                  // E: Clicks Previous Year (2024)
        clicksIndexFormula,                     // F: Clicks YoY
        avgCPCCurrentFormula,                   // G: Avg CPC Current Year (2025)
        cpcWoWFormula,                          // H: Avg CPC Week-over-Week
        avgCPCPreviousFormula,                  // I: Avg CPC Previous Year (2024)
        cpcIndexFormula,                        // J: Avg CPC YoY
        costCurrentFormula,                     // K: Cost Current Year (2025)
        costWoWFormula,                         // L: Cost Week-over-Week
        costPreviousFormula,                    // M: Cost Previous Year (2024)
        costIndexFormula,                       // N: Cost YoY
        conversionsCurrentFormula,              // O: Conversions Current Year (2025)
        conversionsWoWFormula,                  // P: Conversions Week-over-Week
        conversionsPreviousFormula,             // Q: Conversions Previous Year (2024)
        conversionsIndexFormula,                // R: Conversions YoY
        costPerConvCurrentFormula,              // S: Cost / Conv Current Year (2025)
        costPerConvWoWFormula,                  // T: Cost / Conv Week-over-Week
        costPerConvPreviousFormula,             // U: Cost / Conv Previous Year (2024)
        costPerConvIndexFormula,                // V: Cost / Conv YoY
        conversionValueCurrentFormula,          // W: Conversion Value Current Year (2025)
        conversionValueWoWFormula,              // X: Conversion Value Week-over-Week
        conversionValuePreviousFormula,         // Y: Conversion Value Previous Year (2024)
        conversionValueIndexFormula,            // Z: Conversion Value YoY
        convValueCostCurrentFormula,            // AA: Conversion Value / Cost Current Year (2025)
        convValueCostWoWFormula,                // AB: Conversion Value / Cost Week-over-Week
        convValueCostPreviousFormula,           // AC: Conversion Value / Cost Previous Year (2024)
        convValueCostIndexFormula               // AD: Conversion Value / Cost YoY
      ];
      
      formulaData.push(formulaRow);
    }
    
    // Write formula-based data to summary sheet (starting at row 11)
    if (formulaData.length > 0) {
      const columnCount = 30; // Always 30 columns with conditional formulas (7 metrics * 4 columns + 2 week columns)
      debugLog('Writing formula data to sheet', { 
        rowCount: formulaData.length, 
        columnCount 
      });
      
      const summaryRange = summarySheet.getRange(11, 1, formulaData.length, columnCount);
      summaryRange.setValues(formulaData);
      
      debugLog('Formula data written successfully');
    }
    
    // Format main header (row 9 only)
    const columnCount = 30; // Always 30 columns with conditional formulas (7 metrics * 4 columns + 2 week columns)
    const mainHeaderRange = summarySheet.getRange(9, 1, 1, columnCount);
    mainHeaderRange.setFontWeight('bold');
    mainHeaderRange.setBackground('#34a853');
    mainHeaderRange.setFontColor('#ffffff');
    mainHeaderRange.setHorizontalAlignment('center');
    mainHeaderRange.setVerticalAlignment('middle');
    
    // Format year subheader (row 10) with light green background
    const yearSubHeaderRange = summarySheet.getRange(10, 1, 1, columnCount);
    yearSubHeaderRange.setFontWeight('bold');
    yearSubHeaderRange.setBackground('#c8e6c9'); // Light green background
    yearSubHeaderRange.setHorizontalAlignment('center');
    yearSubHeaderRange.setVerticalAlignment('middle');
    
    // v1.7.1 FIX: Ensure year headers display as plain numbers (not currency)
    // This prevents years like 2025 and 2024 from being formatted as currency
    
    // Format data rows (starting at row 11)
    if (formulaData.length > 0) {
      const dataRowCount = formulaData.length;
      const columnCount = 38; // Always 38 columns with conditional formulas (9 metrics * 4 columns + 2 week columns)
      
      // Add alternating row backgrounds
      for (let i = 0; i < dataRowCount; i++) {
        const rowNum = i + 11;
        if (i % 2 === 1) { // Every other row (odd rows)
          const rowRange = summarySheet.getRange(rowNum, 1, 1, columnCount);
          rowRange.setBackground('#f8f9fa'); // Light gray background
        }
      }
      
      // Add borders for full YoY comparison structure (4 columns for all metrics including WoW)
      const totalRows = dataRowCount + 2; // Include header rows 9-10
      
      // Clicks group border (C9:F through last data row) - 4 columns including WoW
      const clicksGroupRange = summarySheet.getRange(9, 3, totalRows, 4); // Columns C-F
      clicksGroupRange.setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      
      // Average CPC group border (G9:J through last data row) - 4 columns including WoW
      const cpcGroupRange = summarySheet.getRange(9, 7, totalRows, 4); // Columns G-J
      cpcGroupRange.setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      
      // Cost group border (K9:N through last data row) - 4 columns including WoW
      const costGroupRange = summarySheet.getRange(9, 11, totalRows, 4); // Columns K-N
      costGroupRange.setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      
      // Conversions group border (O9:R through last data row) - 4 columns including WoW
      const conversionsGroupRange = summarySheet.getRange(9, 15, totalRows, 4); // Columns O-R
      conversionsGroupRange.setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      
      // Cost / Conv group border (S9:V through last data row) - 4 columns including WoW
      const costPerConvGroupRange = summarySheet.getRange(9, 19, totalRows, 4); // Columns S-V
      costPerConvGroupRange.setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      
      // Conversion Value group border (W9:Z through last data row) - 4 columns including WoW
      const convValueGroupRange = summarySheet.getRange(9, 23, totalRows, 4); // Columns W-Z
      convValueGroupRange.setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      

      
      // Conversion Value / Cost group border (AA9:AD through last data row) - 4 columns including WoW
      const convValueCostGroupRange = summarySheet.getRange(9, 27, totalRows, 4); // Columns AA-AD
      convValueCostGroupRange.setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      
      // ROAS (by conv. time) group border (AI9:AL through last data row) - 4 columns including WoW
      const roasGroupRange = summarySheet.getRange(9, 35, totalRows, 4); // Columns AI-AL
      roasGroupRange.setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      
      // Apply locale-specific number formatting
      const costFormat = (LOCALE_SETTINGS.currencyCode === 'JPY' || LOCALE_SETTINGS.currencyCode === 'HUF') ? 
        LOCALE_SETTINGS.currencyFormat : LOCALE_SETTINGS.currencyFormat;
      
      // Week Number (Column A) - Plain numbers, center aligned
      const weekNumberRange = summarySheet.getRange(11, 1, dataRowCount, 1);
      weekNumberRange.setNumberFormat('0');
      weekNumberRange.setHorizontalAlignment('center');
      
      // Week Start (Column B) - Keep as date, center aligned
      const weekStartRange = summarySheet.getRange(11, 2, dataRowCount, 1);
      weekStartRange.setHorizontalAlignment('center');
      
      // Full YoY comparison number formatting with WoW columns
      
      // Clicks (Column C) - Current Year (2025) - Number format, center aligned
      const clicksCurrentRange = summarySheet.getRange(11, 3, dataRowCount, 1);
      clicksCurrentRange.setNumberFormat(LOCALE_SETTINGS.numberFormat);
      clicksCurrentRange.setHorizontalAlignment('center');
      
      // Clicks WoW (Column D) - Percentage format, center aligned
      const clicksWoWRange = summarySheet.getRange(11, 4, dataRowCount, 1);
      clicksWoWRange.setNumberFormat(LOCALE_SETTINGS.numberFormat + '"%"');
      clicksWoWRange.setHorizontalAlignment('center');
      
      // Clicks Previous Year (Column E) - Number format, center aligned
      const clicksPreviousRange = summarySheet.getRange(11, 5, dataRowCount, 1);
      clicksPreviousRange.setNumberFormat(LOCALE_SETTINGS.numberFormat);
      clicksPreviousRange.setHorizontalAlignment('center');
      
      // Clicks YoY (Column F) - Percentage format, center aligned
      const clicksYoYRange = summarySheet.getRange(11, 6, dataRowCount, 1);
      clicksYoYRange.setNumberFormat(LOCALE_SETTINGS.numberFormat + '"%"');
      clicksYoYRange.setHorizontalAlignment('center');
      
      // Average CPC (Column G) - Current Year (2025) - Currency format with 2 decimals, center aligned
      const cpcCurrentRange = summarySheet.getRange(11, 7, dataRowCount, 1);
      cpcCurrentRange.setNumberFormat(LOCALE_SETTINGS.currencyFormatWithDecimals);
      cpcCurrentRange.setHorizontalAlignment('center');
      
      // CPC WoW (Column H) - Percentage format, center aligned
      const cpcWoWRange = summarySheet.getRange(11, 8, dataRowCount, 1);
      cpcWoWRange.setNumberFormat(LOCALE_SETTINGS.numberFormat + '"%"');
      cpcWoWRange.setHorizontalAlignment('center');
      
      // CPC Previous Year (Column I) - Currency format with 2 decimals, center aligned
      const cpcPreviousRange = summarySheet.getRange(11, 9, dataRowCount, 1);
      cpcPreviousRange.setNumberFormat(LOCALE_SETTINGS.currencyFormatWithDecimals);
      cpcPreviousRange.setHorizontalAlignment('center');
      
      // CPC YoY (Column J) - Percentage format, center aligned
      const cpcYoYRange = summarySheet.getRange(11, 10, dataRowCount, 1);
      cpcYoYRange.setNumberFormat(LOCALE_SETTINGS.numberFormat + '"%"');
      cpcYoYRange.setHorizontalAlignment('center');
      
      // Cost (Column K) - Current Year (2025) - Currency format, center aligned
      const costCurrentRange = summarySheet.getRange(11, 11, dataRowCount, 1);
      costCurrentRange.setNumberFormat(costFormat);
      costCurrentRange.setHorizontalAlignment('center');
      
      // Cost WoW (Column L) - Percentage format, center aligned
      const costWoWRange = summarySheet.getRange(11, 12, dataRowCount, 1);
      costWoWRange.setNumberFormat(LOCALE_SETTINGS.numberFormat + '"%"');
      costWoWRange.setHorizontalAlignment('center');
      
      // Cost Previous Year (Column M) - Currency format, center aligned
      const costPreviousRange = summarySheet.getRange(11, 13, dataRowCount, 1);
      costPreviousRange.setNumberFormat(costFormat);
      costPreviousRange.setHorizontalAlignment('center');
      
      // Cost YoY (Column N) - Percentage format, center aligned
      const costYoYRange = summarySheet.getRange(11, 14, dataRowCount, 1);
      costYoYRange.setNumberFormat(LOCALE_SETTINGS.numberFormat + '"%"');
      costYoYRange.setHorizontalAlignment('center');
      
      // Conversions (Column O) - Current Year (2025) - Number format without decimals, center aligned
const conversionsCurrentRange = summarySheet.getRange(11, 15, dataRowCount, 1);
conversionsCurrentRange.setNumberFormat(LOCALE_SETTINGS.numberFormat + '0');
conversionsCurrentRange.setHorizontalAlignment('center');

// Conversions WoW (Column P) - Percentage format, center aligned
const conversionsWoWRange = summarySheet.getRange(11, 16, dataRowCount, 1);
conversionsWoWRange.setNumberFormat(LOCALE_SETTINGS.numberFormat + '"%"');
conversionsWoWRange.setHorizontalAlignment('center');

// Conversions Previous Year (Column Q) - Number format without decimals, center aligned
const conversionsPreviousRange = summarySheet.getRange(11, 17, dataRowCount, 1);
conversionsPreviousRange.setNumberFormat(LOCALE_SETTINGS.numberFormat + '0');
conversionsPreviousRange.setHorizontalAlignment('center');

// Conversions YoY (Column R) - Percentage format, center aligned
const conversionsYoYRange = summarySheet.getRange(11, 18, dataRowCount, 1);
conversionsYoYRange.setNumberFormat(LOCALE_SETTINGS.numberFormat + '"%"');
conversionsYoYRange.setHorizontalAlignment('center');
      
      // Cost / Conv (Column S) - Current Year (2025) - Currency format with 2 decimals, center aligned
      const costPerConvCurrentRange = summarySheet.getRange(11, 19, dataRowCount, 1);
      costPerConvCurrentRange.setNumberFormat(LOCALE_SETTINGS.currencyFormatWithDecimals);
      costPerConvCurrentRange.setHorizontalAlignment('center');
      
      // Cost / Conv WoW (Column T) - Percentage format, center aligned
      const costPerConvWoWRange = summarySheet.getRange(11, 20, dataRowCount, 1);
      costPerConvWoWRange.setNumberFormat(LOCALE_SETTINGS.numberFormat + '"%"');
      costPerConvWoWRange.setHorizontalAlignment('center');
      
      // Cost / Conv Previous Year (Column U) - Currency format with 2 decimals, center aligned
      const costPerConvPreviousRange = summarySheet.getRange(11, 21, dataRowCount, 1);
      costPerConvPreviousRange.setNumberFormat(LOCALE_SETTINGS.currencyFormatWithDecimals);
      costPerConvPreviousRange.setHorizontalAlignment('center');
      
      // Cost / Conv YoY (Column V) - Percentage format, center aligned
      const costPerConvYoYRange = summarySheet.getRange(11, 22, dataRowCount, 1);
      costPerConvYoYRange.setNumberFormat(LOCALE_SETTINGS.numberFormat + '"%"');
      costPerConvYoYRange.setHorizontalAlignment('center');
      
      // Conversion Value (Column W) - Current Year (2025) - Currency format, center aligned
      const conversionValueCurrentRange = summarySheet.getRange(11, 23, dataRowCount, 1);
      conversionValueCurrentRange.setNumberFormat(costFormat);
      conversionValueCurrentRange.setHorizontalAlignment('center');
      
      // Conversion Value WoW (Column X) - Percentage format, center aligned
      const conversionValueWoWRange = summarySheet.getRange(11, 24, dataRowCount, 1);
      conversionValueWoWRange.setNumberFormat(LOCALE_SETTINGS.numberFormat + '"%"');
      conversionValueWoWRange.setHorizontalAlignment('center');
      
      // Conversion Value Previous Year (Column Y) - Currency format, center aligned
      const conversionValuePreviousRange = summarySheet.getRange(11, 25, dataRowCount, 1);
      conversionValuePreviousRange.setNumberFormat(costFormat);
      conversionValuePreviousRange.setHorizontalAlignment('center');
      
      // Conversion Value YoY (Column Z) - Percentage format, center aligned
      const conversionValueYoYRange = summarySheet.getRange(11, 26, dataRowCount, 1);
      conversionValueYoYRange.setNumberFormat(LOCALE_SETTINGS.numberFormat + '"%"');
      conversionValueYoYRange.setHorizontalAlignment('center');
      

      
      // Conversion Value / Cost (Column AA) - Current Year (2025) - Number format with 2 decimals, center aligned
      const convValueCostCurrentRange = summarySheet.getRange(11, 27, dataRowCount, 1);
      convValueCostCurrentRange.setNumberFormat(LOCALE_SETTINGS.numberFormat + '.00');
      convValueCostCurrentRange.setHorizontalAlignment('center');
      
      // Conversion Value / Cost WoW (Column AB) - Percentage format, center aligned
      const convValueCostWoWRange = summarySheet.getRange(11, 28, dataRowCount, 1);
      convValueCostWoWRange.setNumberFormat(LOCALE_SETTINGS.numberFormat + '"%"');
      convValueCostWoWRange.setHorizontalAlignment('center');
      
      // Conversion Value / Cost Previous Year (Column AC) - Number format with 2 decimals, center aligned
      const convValueCostPreviousRange = summarySheet.getRange(11, 29, dataRowCount, 1);
      convValueCostPreviousRange.setNumberFormat(LOCALE_SETTINGS.numberFormat + '.00');
      convValueCostPreviousRange.setHorizontalAlignment('center');
      
      // Conversion Value / Cost YoY (Column AD) - Percentage format, center aligned
      const convValueCostYoYRange = summarySheet.getRange(11, 30, dataRowCount, 1);
      convValueCostYoYRange.setNumberFormat(LOCALE_SETTINGS.numberFormat + '"%"');
      convValueCostYoYRange.setHorizontalAlignment('center');
      

      
      // Apply conditional formatting to WoW columns and YoY columns with very light gradient colors
      const indexColumns = [4, 6, 8, 10, 12, 14, 16, 18, 20, 22, 24, 26, 28, 30]; // WoW and YoY columns only: D=WoW, F=YoY, H=CPC WoW, J=CPC YoY, L=Cost WoW, N=Cost YoY, P=Conv WoW, R=Conv YoY, T=Cost/Conv WoW, V=Cost/Conv YoY, X=Conv Value WoW, Z=Conv Value YoY, AB=Conv Value/Cost WoW, AD=Conv Value/Cost YoY
      
      indexColumns.forEach(colIndex => {
        const indexRange = summarySheet.getRange(11, colIndex, dataRowCount, 1);
        
        // Special handling for metrics where lower is better (WoW and YoY columns)
        // Average CPC WoW (column H), Cost / Conv WoW (column T), Average CPC YoY (column J), Cost YoY (column N), Cost / Conv YoY (column V)
        if (colIndex === 8 || colIndex === 20 || colIndex === 10 || colIndex === 14 || colIndex === 22) {
          // Apply reversed conditional formatting for metrics where lower is better
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
    

    
    // Set column widths for full YoY comparison structure with WoW columns
    summarySheet.setColumnWidth(1, 180);  // App Title / Week Number / Filter text (wider for full display)
    summarySheet.setColumnWidth(2, 100);  // Week Start
    
    // Set equal widths for Clicks columns (C, D, E, F) - 4 columns including WoW
    summarySheet.setColumnWidth(3, 90);   // Clicks 2025
    summarySheet.setColumnWidth(4, 90);   // Clicks 2024
    summarySheet.setColumnWidth(5, 80);   // Clicks WoW
    summarySheet.setColumnWidth(6, 80);   // Clicks Index YoY
    
    // Set equal widths for Average CPC columns (G, H, I, J) - 4 columns including WoW
    summarySheet.setColumnWidth(7, 90);   // Avg CPC 2025
    summarySheet.setColumnWidth(8, 90);   // Avg CPC 2024
    summarySheet.setColumnWidth(9, 80);   // Avg CPC WoW
    summarySheet.setColumnWidth(10, 80);  // Avg CPC Index YoY
    
    // Set equal widths for Cost columns (K, L, M, N) - 4 columns including WoW
    summarySheet.setColumnWidth(11, 90);  // Cost 2025
    summarySheet.setColumnWidth(12, 90);  // Cost 2024
    summarySheet.setColumnWidth(13, 80);  // Cost WoW
    summarySheet.setColumnWidth(14, 80);  // Cost Index YoY
    
    // Set equal widths for Conversions columns (O, P, Q, R) - 4 columns including WoW
    summarySheet.setColumnWidth(15, 100); // Conversions 2025
    summarySheet.setColumnWidth(16, 100); // Conversions 2024
    summarySheet.setColumnWidth(17, 80);  // Conversions WoW
    summarySheet.setColumnWidth(18, 80);  // Conversions Index YoY
    
    // Set equal widths for Cost / Conv columns (S, T, U, V) - 4 columns including WoW
    summarySheet.setColumnWidth(19, 100); // Cost / Conv 2025
    summarySheet.setColumnWidth(20, 100); // Cost / Conv 2024
    summarySheet.setColumnWidth(21, 80);  // Cost / Conv WoW
    summarySheet.setColumnWidth(22, 80);  // Cost / Conv Index YoY
    
    // Set equal widths for Conversion Value columns (W, X, Y, Z) - 4 columns including WoW
    summarySheet.setColumnWidth(23, 110); // Conversion Value 2025
    summarySheet.setColumnWidth(24, 110); // Conversion Value 2024
    summarySheet.setColumnWidth(25, 80);  // Conversion Value WoW
    summarySheet.setColumnWidth(26, 80);  // Conversion Value Index YoY
    
    // Set equal widths for Conv. Value (by conv. time) columns (AA, AB, AC, AD) - 4 columns including WoW
    summarySheet.setColumnWidth(27, 110); // Conv. Value (by conv. time) 2025
    summarySheet.setColumnWidth(28, 110); // Conv. Value (by conv. time) 2024
    summarySheet.setColumnWidth(29, 80);  // Conv. Value (by conv. time) WoW
    summarySheet.setColumnWidth(30, 80);  // Conv. Value (by conv. time) Index YoY
    
          // Set equal widths for Conversion Value / Cost columns (AA, AB, AC, AD) - 4 columns including WoW
      summarySheet.setColumnWidth(27, 90);  // Conversion Value / Cost 2025
      summarySheet.setColumnWidth(28, 90);  // Conversion Value / Cost 2024
      summarySheet.setColumnWidth(29, 80);  // Conversion Value / Cost WoW
      summarySheet.setColumnWidth(30, 80);  // Conversion Value / Cost Index YoY
    
    
    
    // Freeze the filter and header rows for better navigation
    summarySheet.setFrozenRows(10); // Freeze rows 1-10 (app title + filters + toggles + client name + headers)
    summarySheet.setFrozenColumns(2); // Freeze columns A-B (week info)
    
    // Weekly Ads Monitor dashboard created successfully
    debugLog('Weekly Ads Monitor dashboard created successfully');
    debugEnd('createWeeklySummary');
    
  } catch (error) {
    debugError('Error creating Weekly Ads Monitor dashboard', error);
    Logger.log('❌ Error creating Weekly Ads Monitor dashboard: ' + error.toString());
    debugEnd('createWeeklySummary');
  }
}

/**
 * Refresh function to update the sheet based on current toggle settings
 * Call this function manually after changing the YoY toggle in cell B4
 */
function refreshSheet() {
  debugStart('refreshSheet');
  
  try {
    // Refreshing sheet based on current toggle settings
    debugLog('Starting sheet refresh');
    
    // Get the current sheet
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const summarySheet = spreadsheet.getSheetByName('Summary');
    
    debugLog('Sheet lookup', { 
      spreadsheetName: spreadsheet.getName(),
      summarySheetFound: !!summarySheet 
    });
    
    if (!summarySheet) {
      debugError('Summary sheet not found');
      Logger.log('❌ ERROR: Summary sheet not found');
      return;
    }
    
    // Read current toggle values
    const showYoyComparison = getShowYoyComparison(summarySheet);
    
    debugLog('Current toggle settings', { showYoyComparison });
    
    // Current toggle settings read
    
    // Clear ALL existing data and formatting (keep only filters in rows 1-6)
    const lastRow = summarySheet.getLastRow();
    const lastCol = summarySheet.getLastColumn();
    
    debugLog('Sheet dimensions', { lastRow, lastCol });
    
    if (lastRow > 8) {
      // Clear all data from row 9 onwards
      debugLog('Clearing data from row 9 onwards', { rowsToClear: lastRow - 8, columns: lastCol });
      summarySheet.getRange(9, 1, lastRow - 8, lastCol).clear();
    }
    
    // Clear all formatting from row 7 onwards
    if (lastRow >= 7) {
      debugLog('Clearing formatting from row 7 onwards', { rowsToClear: lastRow - 6, columns: lastCol });
      summarySheet.getRange(7, 1, lastRow - 6, lastCol).clearFormat();
    }
    
    // Update headers based on toggle setting
    debugLog('Updating headers');
    updateHeaders(summarySheet, showYoyComparison);
    
    // Regenerate data with current toggle setting
    debugLog('Regenerating data');
    regenerateData(summarySheet, showYoyComparison);
    
    // Apply formatting based on toggle setting
    debugLog('Applying formatting');
    applyFormatting(summarySheet, showYoyComparison);
    
    // Sheet refreshed successfully
    debugLog('Sheet refresh completed successfully');
    debugEnd('refreshSheet');
    
  } catch (error) {
    debugError('Error during refresh', error);
    Logger.log(`❌ ERROR during refresh: ${error.message}`);
    Logger.log(`Stack trace: ${error.stack}`);
    debugEnd('refreshSheet');
  }
}

/**
 * Regenerate data based on current toggle setting
 */
function regenerateData(summarySheet, showYoyComparison) {
  try {
    // Regenerating data with current toggle setting
    
    // Get raw data from Raw sheet
    const rawSheet = summarySheet.getParent().getSheetByName('Raw');
    if (!rawSheet) {
      Logger.log('❌ ERROR: Raw sheet not found');
      return;
    }
    
    const rawData = rawSheet.getDataRange().getValues();
    const rawHeaders = rawData[0];
    
    // Find column indices
    const weekStartCol = rawHeaders.indexOf('Week Start');
    const campaignTypeCol = rawHeaders.indexOf('Campaign Type');
    const isBrandCol = rawHeaders.indexOf('Is Brand');
    
    // Get filter settings from existing sheet
    const campaignTypeFilter = summarySheet.getRange('B2').getValue() || 'All';
    const excludeBrandCampaigns = summarySheet.getRange('B3').getValue() === 'Yes';
    
    // Filter settings applied
    
    // Filter data based on current settings
    let filteredData = rawData.slice(1); // Remove header row
    
    if (campaignTypeFilter !== 'All') {
      filteredData = filteredData.filter(row => row[campaignTypeCol] === campaignTypeFilter);
    }
    
    if (excludeBrandCampaigns) {
      filteredData = filteredData.filter(row => !row[isBrandCol]);
    }
    
    // Get unique week start dates
    const weekStartDates = [...new Set(filteredData.map(row => row[weekStartCol]))].sort();
    
    // Generate formula data based on toggle
    const formulaData = [];
    const currentYear = new Date().getFullYear();
    const previousYear = currentYear - 1;
    
    weekStartDates.forEach((weekStart, i) => {
      const rowNum = i + 9; // Starting at row 9
      const weekStartQuoted = `"${weekStart}"`;
      
      // Check if this week exists in current year
      const currentWeek = filteredData.some(row => row[weekStartCol] === weekStart && new Date(row[weekStartCol]).getFullYear() === currentYear);
      const previousWeek = filteredData.some(row => row[weekStartCol] === weekStart && new Date(row[weekStartCol]).getFullYear() === previousYear);
      
      // Generate week number and display
      const weekNumber = i + 1;
      const displayWeek = weekStart;
      
      // Generate formulas based on toggle setting
      if (showYoyComparison) {
        // Full YoY comparison formulas (29 columns)
        const formulaRow = generateYoYFormulaRow(weekStart, weekStartQuoted, previousWeekQuoted, rowNum, currentWeek, previousWeek, filteredData, weekStartCol, currentYear, previousYear);
        formulaData.push(formulaRow);
      } else {
        // Current year only formulas (11 columns)
        const formulaRow = generateCurrentYearFormulaRow(weekStart, weekStartQuoted, rowNum, currentWeek, filteredData, weekStartCol);
        formulaData.push(formulaRow);
      }
    });
    
    // Write data to sheet
    if (formulaData.length > 0) {
      const columnCount = showYoyComparison ? 30 : 11;
      const summaryRange = summarySheet.getRange(9, 1, formulaData.length, columnCount);
      summaryRange.setValues(formulaData);
    }
    
  } catch (error) {
    Logger.log(`❌ ERROR in regenerateData: ${error.message}`);
    Logger.log(`Stack trace: ${error.stack}`);
  }
}

/**
 * Generate YoY formula row (29 columns) - complete version
 */
function generateYoYFormulaRow(weekStart, weekStartQuoted, previousWeekQuoted, rowNum, currentWeek, previousWeek, filteredData, weekStartCol, currentYear, previousYear) {
  const weekNumber = Math.floor((new Date(weekStart) - new Date(`${currentYear}-01-01`)) / (7 * 24 * 60 * 60 * 1000)) + 1;
  const displayWeek = weekStart;
  
  // Current year formulas
  const clicksCurrentFormula = currentWeek ? `=SUMIFS(Raw!G:G,Raw!A:A,${weekStartQuoted})` : '0';
  const costCurrentFormula = currentWeek ? `=SUMIFS(Raw!H:H,Raw!A:A,${weekStartQuoted})` : '0';
  const conversionsCurrentFormula = currentWeek ? `=SUMIFS(Raw!J:J,Raw!A:A,${weekStartQuoted})` : '0';
  const conversionValueCurrentFormula = currentWeek ? `=SUMIFS(Raw!K:K,Raw!A:A,${weekStartQuoted})` : '0';
  const conversionValueConvTimeCurrentFormula = currentWeek ? `=SUMIFS(Raw!O:O,Raw!A:A,${weekStartQuoted})` : '0';
  
  // Previous year formulas
  const clicksPreviousFormula = previousWeek ? `=SUMIFS(Raw!G:G,Raw!A:A,${previousWeekQuoted})` : '0';
  const costPreviousFormula = previousWeek ? `=SUMIFS(Raw!H:H,Raw!A:A,${previousWeekQuoted})` : '0';
  const conversionsPreviousFormula = previousWeek ? `=SUMIFS(Raw!J:J,Raw!A:A,${previousWeekQuoted})` : '0';
  const conversionValuePreviousFormula = previousWeek ? `=SUMIFS(Raw!K:K,Raw!A:A,${previousWeekQuoted})` : '0';
  const conversionValueConvTimePreviousFormula = previousWeek ? `=SUMIFS(Raw!O:O,Raw!A:A,${previousWeekQuoted})` : '0';
  
  // Cost per conversion formulas
  const costPerConvCurrentFormula = currentWeek ? `=IF(M${rowNum}>0,J${rowNum}/M${rowNum},0)` : '0';
  const costPerConvPreviousFormula = previousWeek ? `=IF(N${rowNum}>0,K${rowNum}/N${rowNum},0)` : '0';
  
  // Derived metrics
  const avgCPCCurrentFormula = `=IF(C${rowNum}>0,J${rowNum}/C${rowNum},0)`;
  const avgCPCPreviousFormula = `=IF(D${rowNum}>0,K${rowNum}/D${rowNum},0)`;
  const clicksIndexFormula = `=IF(D${rowNum}>0,(C${rowNum}/D${rowNum})*100,0)`;
  const cpcIndexFormula = `=IF(H${rowNum}>0,(G${rowNum}/H${rowNum})*100,0)`;
  const costIndexFormula = `=IF(K${rowNum}>0,(J${rowNum}/K${rowNum})*100,0)`;
  const conversionsIndexFormula = `=IF(N${rowNum}>0,(M${rowNum}/N${rowNum})*100,0)`;
  const costPerConvIndexFormula = `=IF(Q${rowNum}>0,(P${rowNum}/Q${rowNum})*100,0)`;
  const conversionValueIndexFormula = `=IF(T${rowNum}>0,(S${rowNum}/T${rowNum})*100,0)`;
  const conversionValueConvTimeIndexFormula = `=IF(W${rowNum}>0,(V${rowNum}/W${rowNum})*100,0)`;
  const convValueCostCurrentFormula = `=IF(J${rowNum}>0,S${rowNum}/J${rowNum},0)`;
  const convValueCostPreviousFormula = `=IF(K${rowNum}>0,T${rowNum}/K${rowNum},0)`;
  const convValueCostIndexFormula = `=IF(Z${rowNum}>0,(Y${rowNum}/Z${rowNum})*100,0)`;
  // Week-over-week formula for clicks
  const clicksWoWFormula = `=IF(${rowNum}>9,IF(OFFSET(C${rowNum},-1,0)>0,(C${rowNum}/OFFSET(C${rowNum},-1,0))*100,0),0)`;
  
  const roasCurrentFormula = `=IF(J${rowNum}>0,V${rowNum}/J${rowNum},0)`;
  const roasPreviousFormula = `=IF(K${rowNum}>0,W${rowNum}/K${rowNum},0)`;
  const roasIndexFormula = `=IF(AC${rowNum}>0,(AB${rowNum}/AC${rowNum})*100,0)`;
  
  return [
    weekNumber,                              // A: Week Number
    displayWeek,                            // B: Week Start
    clicksCurrentFormula,                   // C: Clicks Current Year (2025)
    clicksPreviousFormula,                  // D: Clicks Previous Year (2024)
    clicksWoWFormula,                       // E: Clicks Week-over-Week
    clicksIndexFormula,                     // F: Clicks Index YoY
    avgCPCCurrentFormula,                   // G: Avg CPC Current Year (2025)
    avgCPCPreviousFormula,                  // H: Avg CPC Previous Year (2024)
    cpcIndexFormula,                        // I: Avg CPC Index YoY
    costCurrentFormula,                     // J: Cost Current Year (2025)
    costPreviousFormula,                    // K: Cost Previous Year (2024)
    costIndexFormula,                       // L: Cost Index YoY
    conversionsCurrentFormula,              // M: Conversions Current Year (2025)
    conversionsPreviousFormula,             // N: Conversions Previous Year (2024)
    conversionsIndexFormula,                // O: Conversions Index YoY
    costPerConvCurrentFormula,              // P: Cost / Conv Current Year (2025)
    costPerConvPreviousFormula,             // Q: Cost / Conv Previous Year (2024)
    costPerConvIndexFormula,                // R: Cost / Conv Index YoY
    conversionValueCurrentFormula,          // S: Conversion Value Current Year (2025)
    conversionValuePreviousFormula,         // T: Conversion Value Previous Year (2024)
    conversionValueIndexFormula,            // U: Conversion Value Index YoY
    conversionValueConvTimeCurrentFormula,  // V: Conv. Value (by conv. time) Current Year (2025)
    conversionValueConvTimePreviousFormula, // W: Conv. Value (by conv. time) Previous Year (2024)
    conversionValueConvTimeIndexFormula,    // X: Conv. Value (by conv. time) Index YoY
    convValueCostCurrentFormula,            // Y: Conversion Value / Cost Current Year (2025)
    convValueCostPreviousFormula,           // Z: Conversion Value / Cost Previous Year (2024)
    convValueCostIndexFormula,              // AA: Conversion Value / Cost Index YoY
    roasCurrentFormula,                     // AB: ROAS (by conv. time) Current Year (2025)
    roasPreviousFormula,                    // AC: ROAS (by conv. time) Previous Year (2024)
    roasIndexFormula                        // AD: ROAS (by conv. time) Index YoY
  ];
}

/**
 * Generate current year only formula row (11 columns) - complete version
 */
function generateCurrentYearFormulaRow(weekStart, weekStartQuoted, rowNum, currentWeek, filteredData, weekStartCol) {
  const weekNumber = Math.floor((new Date(weekStart) - new Date(`${new Date().getFullYear()}-01-01`)) / (7 * 24 * 60 * 60 * 1000)) + 1;
  const displayWeek = weekStart;
  
  // Current year formulas
  const clicksCurrentFormula = currentWeek ? `=SUMIFS(Raw!G:G,Raw!A:A,${weekStartQuoted})` : '0';
  const costCurrentFormula = currentWeek ? `=SUMIFS(Raw!H:H,Raw!A:A,${weekStartQuoted})` : '0';
  const conversionsCurrentFormula = currentWeek ? `=SUMIFS(Raw!J:J,Raw!A:A,${weekStartQuoted})` : '0';
  const conversionValueCurrentFormula = currentWeek ? `=SUMIFS(Raw!K:K,Raw!A:A,${weekStartQuoted})` : '0';
  const conversionValueConvTimeCurrentFormula = currentWeek ? `=SUMIFS(Raw!O:O,Raw!A:A,${weekStartQuoted})` : '0';
  
  // Cost per conversion formula
  const costPerConvFormula = currentWeek ? `=IF(F${rowNum}>0,E${rowNum}/F${rowNum},0)` : '0';
  
  // Derived metrics
  const avgCPCFormula = `=IF(C${rowNum}>0,E${rowNum}/C${rowNum},0)`;
  const convValueCostFormula = `=IF(E${rowNum}>0,H${rowNum}/E${rowNum},0)`;
  const roasFormula = `=IF(E${rowNum}>0,I${rowNum}/E${rowNum},0)`;
  
  return [
    weekNumber,                              // A: Week Number
    displayWeek,                            // B: Week Start
    clicksCurrentFormula,                   // C: Clicks Current Year (2025)
    avgCPCFormula,                          // D: Average CPC Current Year (2025)
    costCurrentFormula,                     // E: Cost Current Year (2025)
    conversionsCurrentFormula,              // F: Conversions Current Year (2025)
    costPerConvFormula,                     // G: Cost / Conv Current Year (2025)
    conversionValueCurrentFormula,          // H: Conversion Value Current Year (2025)
    conversionValueConvTimeCurrentFormula,  // I: Conv. Value (by conv. time) Current Year (2025)
    convValueCostFormula,                   // J: Conversion Value / Cost Current Year (2025)
    roasFormula                             // K: ROAS (by conv. time) Current Year (2025)
  ];
}

/**
 * Test function to debug toggle functionality
 * Run this to see current toggle state and test the refresh
 */
function testToggle() {
  debugStart('testToggle');
  
  try {
    Logger.log('🧪 TESTING TOGGLE FUNCTIONALITY');
    debugLog('Starting toggle test function');
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const summarySheet = spreadsheet.getSheetByName('Summary');
    
    debugLog('Sheet lookup for test', { 
      spreadsheetName: spreadsheet.getName(),
      summarySheetFound: !!summarySheet 
    });
    
    if (!summarySheet) {
      debugError('Summary sheet not found for test');
      Logger.log('❌ ERROR: Summary sheet not found');
      return;
    }
    
    // Check current toggle value
    const toggleValue = summarySheet.getRange('B4').getValue();
    Logger.log(`📊 Current toggle value in B4: "${toggleValue}"`);
    debugLog('Current toggle value', { toggleValue });
    
    // Test the refresh function
    Logger.log('🔄 Testing refreshSheet() function...');
    debugLog('Testing refreshSheet function');
    refreshSheet();
    
    Logger.log('✅ Test complete!');
    debugLog('Toggle test completed successfully');
    debugEnd('testToggle');
    
  } catch (error) {
    debugError('Error in testToggle', error);
    Logger.log(`❌ ERROR in testToggle: ${error.message}`);
    Logger.log(`Stack trace: ${error.stack}`);
    debugEnd('testToggle');
  }
}

/**
 * Quick toggle function - switches between YoY and current year only
 * Run this to quickly switch views without manually changing the dropdown
 */
function toggleYoY() {
  debugStart('toggleYoY');
  
  try {
    Logger.log('🔄 QUICK TOGGLE FUNCTION');
    debugLog('Starting YoY toggle function');
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const summarySheet = spreadsheet.getSheetByName('Summary');
    
    debugLog('Sheet lookup for toggle', { 
      spreadsheetName: spreadsheet.getName(),
      summarySheetFound: !!summarySheet 
    });
    
    if (!summarySheet) {
      debugError('Summary sheet not found for toggle');
      Logger.log('❌ ERROR: Summary sheet not found');
      return;
    }
    
    // Get current toggle value
    const currentValue = summarySheet.getRange('B4').getValue();
    const newValue = currentValue === 'Yes' ? 'No' : 'Yes';
    
    debugLog('Toggle value change', { currentValue, newValue });
    
    // Update the toggle
    summarySheet.getRange('B4').setValue(newValue);
    Logger.log(`📊 Toggle changed from "${currentValue}" to "${newValue}"`);
    
    // Refresh the sheet
    debugLog('Refreshing sheet after toggle change');
    refreshSheet();
    
    Logger.log(`✅ Successfully switched to ${newValue === 'Yes' ? 'YoY comparison' : 'current year only'} view`);
    debugLog('Toggle completed successfully', { newValue });
    debugEnd('toggleYoY');
    
  } catch (error) {
    debugError('Error in toggleYoY', error);
    Logger.log(`❌ ERROR in toggleYoY: ${error.message}`);
    Logger.log(`Stack trace: ${error.stack}`);
    debugEnd('toggleYoY');
  }
}

/**
 * NOTE: Google Ads Script does not support automatic triggers like Google Apps Script.
 * To change the toggle, you must either:
 * 1. Change cell B4 manually and run refreshSheet()
 * 2. Use the toggleYoY() function to quickly switch between views
 */

/**
 * Function to apply reversed conditional formatting for Average CPC (lower is better)
 */
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

/**
 * Apply formatting based on toggle setting
 */
function applyFormatting(summarySheet, showYoyComparison) {
  try {
    const lastRow = summarySheet.getLastRow();
    if (lastRow < 9) return; // No data to format
    
    const dataRowCount = lastRow - 8; // Data starts at row 9
    
    // Format main header (row 7 only)
    const columnCount = showYoyComparison ? 29 : 11;
    const mainHeaderRange = summarySheet.getRange(7, 1, 1, columnCount);
    mainHeaderRange.setFontWeight('bold');
    mainHeaderRange.setBackground('#4285f4');
    mainHeaderRange.setFontColor('white');
    mainHeaderRange.setHorizontalAlignment('center');
    
    // Format year header (row 8 only)
    const yearHeaderRange = summarySheet.getRange(8, 1, 1, columnCount);
    yearHeaderRange.setFontWeight('bold');
    yearHeaderRange.setBackground('#e8f0fe');
    yearHeaderRange.setHorizontalAlignment('center');
    
    // Format data rows
    const dataRange = summarySheet.getRange(9, 1, dataRowCount, columnCount);
    dataRange.setHorizontalAlignment('center');
    
    // Add alternating row backgrounds
    for (let row = 9; row <= lastRow; row++) {
      const rowRange = summarySheet.getRange(row, 1, 1, columnCount);
      if ((row - 9) % 2 === 0) {
        rowRange.setBackground('#f8f9fa');
      } else {
        rowRange.setBackground('#ffffff');
      }
    }
    
    // Apply borders
    const totalRows = dataRowCount + 2; // Include header rows 7-8
    
    if (showYoyComparison) {
      // Full YoY comparison borders (3 columns per metric)
      // Clicks group border (C7:E through last data row)
      summarySheet.getRange(7, 3, totalRows, 3).setBorder(true, true, true, true, true, true);
      // Avg CPC group border (F7:H through last data row)
      summarySheet.getRange(7, 6, totalRows, 3).setBorder(true, true, true, true, true, true);
      // Cost group border (I7:K through last data row)
      summarySheet.getRange(7, 9, totalRows, 3).setBorder(true, true, true, true, true, true);
      // Conversions group border (L7:N through last data row)
      summarySheet.getRange(7, 12, totalRows, 3).setBorder(true, true, true, true, true, true);
      // Conversions (by conv. time) group border (O7:Q through last data row)
      summarySheet.getRange(7, 15, totalRows, 3).setBorder(true, true, true, true, true, true);
      // Conversion Value group border (R7:T through last data row)
      summarySheet.getRange(7, 18, totalRows, 3).setBorder(true, true, true, true, true, true);
      // Conv. Value (by conv. time) group border (U7:W through last data row)
      summarySheet.getRange(7, 21, totalRows, 3).setBorder(true, true, true, true, true, true);
      // Conversion Value / Cost group border (X7:Z through last data row)
      summarySheet.getRange(7, 24, totalRows, 3).setBorder(true, true, true, true, true, true);
      // ROAS (by conv. time) group border (AA7:AC through last data row)
      summarySheet.getRange(7, 27, totalRows, 3).setBorder(true, true, true, true, true, true);
    } else {
      // Current year only borders (1 column per metric)
      // Week Number and Week Start
      summarySheet.getRange(7, 1, totalRows, 2).setBorder(true, true, true, true, true, true);
      // Clicks
      summarySheet.getRange(7, 3, totalRows, 1).setBorder(true, true, true, true, true, true);
      // Average CPC
      summarySheet.getRange(7, 4, totalRows, 1).setBorder(true, true, true, true, true, true);
      // Cost
      summarySheet.getRange(7, 5, totalRows, 1).setBorder(true, true, true, true, true, true);
      // Conversions
      summarySheet.getRange(7, 6, totalRows, 1).setBorder(true, true, true, true, true, true);
      // Conversions (by conv. time)
      summarySheet.getRange(7, 7, totalRows, 1).setBorder(true, true, true, true, true, true);
      // Conversion Value
      summarySheet.getRange(7, 8, totalRows, 1).setBorder(true, true, true, true, true, true);
      // Conv. Value (by conv. time)
      summarySheet.getRange(7, 9, totalRows, 1).setBorder(true, true, true, true, true, true);
      // Conversion Value / Cost
      summarySheet.getRange(7, 10, totalRows, 1).setBorder(true, true, true, true, true, true);
      // ROAS (by conv. time)
      summarySheet.getRange(7, 11, totalRows, 1).setBorder(true, true, true, true, true, true);
    }
    
    // Apply number formatting
    const weekNumberRange = summarySheet.getRange(9, 1, dataRowCount, 1);
    weekNumberRange.setNumberFormat('0');
    
    const weekStartRange = summarySheet.getRange(9, 2, dataRowCount, 1);
    weekStartRange.setNumberFormat('yyyy-mm-dd');
    weekStartRange.setHorizontalAlignment('center');
    
    if (showYoyComparison) {
      // Full YoY comparison number formatting
      
      // Currency columns (Cost, Conversion Value, Conv. Value by conv. time)
      const currencyColumns = [9, 10, 18, 19, 21, 22, 24, 25, 27, 28]; // I, J, R, S, U, V, X, Y, AA, AB
      currencyColumns.forEach(col => {
        if (col <= columnCount) {
          summarySheet.getRange(9, col, dataRowCount, 1).setNumberFormat('$#,##0.00');
        }
      });
      
      // Percentage columns (Index YoY columns)
      const percentageColumns = [5, 8, 11, 14, 17, 20, 23, 26, 29]; // E, H, K, N, Q, T, W, Z, AC
      percentageColumns.forEach(col => {
        if (col <= columnCount) {
          summarySheet.getRange(9, col, dataRowCount, 1).setNumberFormat('0.0%');
        }
      });
      
      // Ratio columns (Conversion Value / Cost, ROAS)
      const ratioColumns = [24, 25, 27, 28]; // X, Y, AA, AB
      ratioColumns.forEach(col => {
        if (col <= columnCount) {
          summarySheet.getRange(9, col, dataRowCount, 1).setNumberFormat('0.00');
        }
      });
      
      // Apply conditional formatting to Index YoY columns with very light gradient colors
      const indexColumns = [5, 6, 9, 12, 15, 18, 21, 24, 27, 30]; // E=WoW, F=Index YoY, I=CPC Index, L=Cost Index, O=Conv Index, R=Cost/Conv Index, U=Conv Value Index, X=Conv Value Conv Time Index, AA=Conv Value/Cost Index, AD=ROAS Index
      indexColumns.forEach(col => {
        if (col <= columnCount) {
          const range = summarySheet.getRange(9, col, dataRowCount, 1);
          const rule = SpreadsheetApp.newConditionalFormatRule()
            .whenNumberGreaterThan(100)
            .setBackground('#d4edda')
            .setRanges([range])
            .build();
          summarySheet.setConditionalFormatRules([rule]);
        }
      });
      
    } else {
      // Current year only number formatting
      
      // Currency columns (Cost, Conversion Value, Conv. Value by conv. time)
      const currencyColumns = [5, 8, 9]; // E, H, I
      currencyColumns.forEach(col => {
        if (col <= columnCount) {
          summarySheet.getRange(9, col, dataRowCount, 1).setNumberFormat('$#,##0.00');
        }
      });
      
      // Ratio columns (Conversion Value / Cost, ROAS)
      const ratioColumns = [10, 11]; // J, K
      ratioColumns.forEach(col => {
        if (col <= columnCount) {
          summarySheet.getRange(9, col, dataRowCount, 1).setNumberFormat('0.00');
        }
      });
    }
    
    // Set column widths
    summarySheet.setColumnWidth(1, 80);   // Week Number
    summarySheet.setColumnWidth(2, 100);  // Week Start
    
    if (showYoyComparison) {
      // Full YoY comparison column widths
      // Set equal widths for Clicks columns (C, D, E, F)
      summarySheet.setColumnWidth(3, 80);  // Clicks Current
      summarySheet.setColumnWidth(4, 80);  // Clicks Previous
      summarySheet.setColumnWidth(5, 80);  // Clicks WoW
      summarySheet.setColumnWidth(6, 80);  // Clicks Index
      
      // Set equal widths for Avg CPC columns (G, H, I)
      summarySheet.setColumnWidth(7, 100); // Avg CPC Current
      summarySheet.setColumnWidth(8, 100); // Avg CPC Previous
      summarySheet.setColumnWidth(9, 100); // Avg CPC Index
      
      // Set equal widths for Cost columns (J, K, L)
      summarySheet.setColumnWidth(10, 100); // Cost Current
      summarySheet.setColumnWidth(11, 100); // Cost Previous
      summarySheet.setColumnWidth(12, 100); // Cost Index
      
      // Continue for all other metric groups...
      for (let i = 13; i <= 30; i++) {
        summarySheet.setColumnWidth(i, 100);
      }
    } else {
      // Current year only column widths
      summarySheet.setColumnWidth(3, 80);   // Clicks
      summarySheet.setColumnWidth(4, 100);  // Average CPC
      summarySheet.setColumnWidth(5, 100);  // Cost
      summarySheet.setColumnWidth(6, 100);  // Conversions
      summarySheet.setColumnWidth(7, 100);  // Conversions (by conv. time)
      summarySheet.setColumnWidth(8, 100);  // Conversion Value
      summarySheet.setColumnWidth(9, 100);  // Conv. Value (by conv. time)
      summarySheet.setColumnWidth(10, 100); // Conversion Value / Cost
      summarySheet.setColumnWidth(11, 100); // ROAS (by conv. time)
    }
    
    Logger.log(`✅ Applied formatting for ${showYoyComparison ? 'YoY comparison' : 'current year only'} view`);
    
  } catch (error) {
    Logger.log(`❌ ERROR in applyFormatting: ${error.message}`);
  }
}

/**
 * Update headers based on toggle setting
 */
function updateHeaders(summarySheet, showYoyComparison) {
  const currentYear = new Date().getFullYear();
  const previousYear = currentYear - 1;
  
  if (showYoyComparison) {
    // Full YoY comparison headers (4 columns for clicks, 3 for others)
    const mainHeaders = ['Week Number', 'Week Start', 'Clicks', '', '', '', 'Average CPC', '', '', 'Cost', '', '', 'Conversions', '', '', 'Cost / Conv', '', '', 'Conversion Value', '', '', 'Conv. Value (by conv. time)', '', '', 'Conversion Value / Cost', '', '', 'ROAS (by conv. time)', '', ''];
    const yearHeaders = ['', '', currentYear, previousYear, 'WoW', 'Index YoY', currentYear, previousYear, 'Index YoY', currentYear, previousYear, 'Index YoY', currentYear, previousYear, 'Index YoY', currentYear, previousYear, 'Index YoY', currentYear, previousYear, 'Index YoY', currentYear, previousYear, 'Index YoY', currentYear, previousYear, 'Index YoY', currentYear, previousYear, 'Index YoY'];
    const headerRange = 'A7:AD8';
    summarySheet.getRange(headerRange).setValues([mainHeaders, yearHeaders]);
  } else {
    // Current year only headers (1 column per metric)
    const mainHeaders = ['Week Number', 'Week Start', 'Clicks', 'Average CPC', 'Cost', 'Conversions', 'Conversions (by conv. time)', 'Conversion Value', 'Conv. Value (by conv. time)', 'Conversion Value / Cost', 'ROAS (by conv. time)'];
    const yearHeaders = ['', '', currentYear, currentYear, currentYear, currentYear, currentYear, currentYear, currentYear, currentYear, currentYear];
    const headerRange = 'A7:K8';
    summarySheet.getRange(headerRange).setValues([mainHeaders, yearHeaders]);
  }
}

// Helper function to create complex nested IF formulas with locale awareness
function createNestedIfFormula(condition1, trueValue1, falseValue1, condition2, trueValue2, falseValue2) {
  const innerIf = createIfFormula(condition2, trueValue2, falseValue2);
  return createIfFormula(condition1, trueValue1, innerIf);
}

// Helper function to create SUMIFS formulas with brand campaign filtering
function createSumifsWithBrandFilter(sumRange, dateRange, dateValue, brandRange, brandValue, campaignTypeRange, campaignTypeValue) {
  if (brandValue === 'FALSE') {
    // Exclude brand campaigns
    if (campaignTypeValue === 'All') {
      return createSumifsFormula(sumRange, dateRange, dateValue, brandRange, brandValue);
    } else {
      return createSumifsFormula(sumRange, dateRange, dateValue, brandRange, brandValue, campaignTypeRange, campaignTypeValue);
    }
  } else {
    // Include all campaigns
    if (campaignTypeValue === 'All') {
      return createSumifsFormula(sumRange, dateRange, dateValue);
    } else {
      return createSumifsFormula(sumRange, dateRange, dateValue, campaignTypeRange, campaignTypeValue);
    }
  }
}

// Helper function to create complex SUMIFS with brand and campaign type filtering
function createComplexSumifsFormula(sumRange, dateRange, dateValue, excludeBrand, campaignTypeValue, brandRange, campaignTypeRange) {
  // Create the nested IF formula: IF(B3="Yes", IF(B2="All", SUMIFS(...), SUMIFS(...)), IF(B2="All", SUMIFS(...), SUMIFS(...)))
  
  // Inner IF for when B3="Yes" (exclude brand campaigns)
  const innerIfExcludeBrand = createIfFormula(
    'B2="All"',
    createSumifsFormula(sumRange, dateRange, dateValue, brandRange, 'FALSE'),
    createSumifsFormula(sumRange, dateRange, dateValue, brandRange, 'FALSE', campaignTypeRange, 'B2')
  );
  
  // Inner IF for when B3!="Yes" (include all campaigns)
  const innerIfIncludeAll = createIfFormula(
    'B2="All"',
    createSumifsFormula(sumRange, dateRange, dateValue),
    createSumifsFormula(sumRange, dateRange, dateValue, campaignTypeRange, 'B2')
  );
  
  // Outer IF to choose between exclude brand and include all
  return createIfFormula('B3="Yes"', innerIfExcludeBrand, innerIfIncludeAll);
}
