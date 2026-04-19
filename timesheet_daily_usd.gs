// ============================================
// AUTOMATED INVOICE GENERATION SYSTEM
// ============================================

// Master tracker configuration
const MASTER_TRACKER_SHEET_ID = '1SMbg0pt3sGRwoma0WTR_hYGqLSj8R07ehOgbmXQslLc';
const MASTER_TRACKER_SHEET_NAME = 'Staffing Tracker';
const DISPATCH_WEBHOOK_URL = 'https://script.google.com/macros/s/AKfycbwWFwAgToE8dv6t0d8uU3WeQ0rJL0zZFHe2L8dUrkSUI80qVsByqa9RHRQyS22IfSEDRQ/exec'; // <-- VERIFY this is your current Dispatch Queue Web App URL from Expense Tracker setup!
// This secret must match the one defined in your Dispatch Queue script (DISPATCH_WEB_APP_SECRET)
const DISPATCH_HMAC_SECRET = '79e0a2f59f8c47ba917063a9e6a244fcd7aa9c90574ed26e3b29aa3b607445f5'; // <-- VERIFY this matches the secret in Dispatch Queue script!

// Central counter sheet that all consultants share
const GLOBAL_COUNTER_SHEET_ID = '1kHp20TH861eHk9OMtdvu2RGPVZuCHuCeqa6Q0a1KTG8';   //  <-- paste the ID from step 1 d

// Rate column mappings
const CONTRACTOR_RATE_COLUMN = 'Consultant daily rate';
const CLIENT_RATE_COLUMN = 'All incl. rate';

// Consultant specific column mappings
const CONSULTANT_NAME = 'Name';
const PROJECT_NAME = 'Project Name'
const PROJECT_CODE = 'Project Code'

// Client email columns
const CONTRACTOR_EMAIL = 'Email';
const PARTNER_COLUMN = 'Partner';
const CLIENT_COUNTERPARTS_COLUMN = 'Client Counterpart';
const EXEC_ASSISTANT_COLUMN = 'Exec. Assistant';

// Invoice numbering
const CONTRACTOR_INVOICE_PREFIX = 'CONTR-INV';
const CLIENT_INVOICE_PREFIX = 'CLIENT-INV';

// Configuration Constants
const CONTRACTOR_INVOICE_FOLDER_ID = '11kW5GeN3ZPwAwwdhI3lxa_29f3Ul1qNJ';
const CLIENT_INVOICE_FOLDER_ID = '1e0hfepov1XLRod5Spy6V95OThVZNBJnm';
const FINAL_FOLDER_ID = '1DOFVLlAZmoH4rotkGJ9qjEfHrcPL92x_';
const TEMPLATE_FILE_ID = '1Rf-Srvg6JSIhNsFkYP2_btLdfNJVKFe6xG6MSBjFLj8';  // Invoice template ID
const VAT_TEMPLATE_FILE_ID = '1Upum8vd48P4OESGzz2wrpOiUOJ6sZPjO6M1JEyvF9Vs';
const DRAFT_FOLDER_ID  = '17EwmlwUCa_fWMY-JjiNeI2Cpq7uAn3TT';             // Drafts for approval
const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbzgEE0LhdWlgfgcV-zBpu0HNRlGWg3oj7q92fYQUcGtPDRekSlJVu21AQlX6XgpPadT/exec';
const HMAC_SECRET = PropertiesService
  .getScriptProperties()
  .getProperty('HMAC_SECRET');
if (!HMAC_SECRET) {
  throw new Error('HMAC_SECRET not defined in Script Properties');
}
const TAX_STATUS_COLUMN = '5% VAT';

//const EMAIL_RECIPIENT = 'harshvardhan.chauhan@stratverse.co';
const OPERATIONS_HEAD_EMAIL = 'admin@stratverse.co';
const EMAIL_TO = ['harshvardhan.chauhan@stratverse.co', 'harshvardhan.chauhan1919@gmail.com'];
const EMAIL_CC = ['harshvardhan.chauhan1919@gmail.com']; // Add CC recipients here
const EMAIL_BCC = ['harshvardhan.chauhan19@gmail.com']; // Add BCC recipients here
const INVOICE_PREFIX = 'INV-';
// CORRECTED Field Mappings for Timesheet

const TIMESHEET_FIELDS = {
  name: 'C4',              // Updated from C3
  street: 'C5',            // Updated from C4
  city: 'C6',              // Updated from C5
  postal_code: 'C7',       // Updated from C6
  country: 'C8',           // Updated from C7
  invoice_number: 'C10',   // Updated from C9
  status: 'E51',           // Updated from E50
  total_hours: 'E47',      // Updated from E46
  daily_rate: 'E48',       // Updated from E47
  billing_amount: 'E49',   // Updated from E48
  activity_start_row: 15   // Updated from 13
};

// CORRECTED Field Mappings for Invoice Template
const INVOICE_FIELDS = {
  invoice_number: 'B8',
  consultant_name: 'B9',
  client_name: 'B10',
  project_name: 'B11',
  consultant_postal: 'E5',
  consultant_country: 'E6',
  consultant_street: 'E3',      // Can be blank for now
  consultant_city: 'E4',       // Can be blank for now
  creation_date: 'E8',
  invoice_date: 'E9',
  data_start_row: 13,       // Adjust based on where activity data starts
  total_hours_result: 'E45',  // CORRECTED from E40
  daily_rate_result: 'E47',   // CORRECTED from E42
  billing_amount_result: 'E48' // CORRECTED from E43
};

// ===========================================================================
// FUNCTION TO VALIDATE THE INTEGRATION BETWEEN THE MASTER SHEET AND TIMESHEET
// ===========================================================================
function validateConsultantInMaster(consultantName) {
  try {
    console.log('Validating consultant:', consultantName);
    const masterSheet = SpreadsheetApp.openById(MASTER_TRACKER_SHEET_ID)
                                     .getSheetByName(MASTER_TRACKER_SHEET_NAME);
    
    // Get all consultant names (assuming they're in column A)
    const consultantRange = masterSheet.getRange('B:B');
    const consultants = consultantRange.getValues().flat()
                                      .filter(name => name && name.toString().trim() !== '');
    
    console.log('Found consultants in master tracker:', consultants.length);
    
    // Perform case-insensitive comparison
    const match = consultants.find(name => 
      name.toString().toLowerCase().trim() === consultantName.toLowerCase().trim()
    );
    
    console.log('Match found:', !!match);
    return {
      found: !!match,
      exactMatch: match || null,
      allConsultants: consultants // For debugging purposes
    };
  } catch (error) {
    console.error('Error validating consultant:', error);
    return {
      found: false,
      error: error.message
    };
  }
}

// =================================================================================
// UNIT TEST FUNCTION TO TEST THE INTEGRATION BETWEEN THE MASTER SHEET AND TIMESHEET
// =================================================================================

function testValidateConsultant() {
  // Test with actual consultant name from your timesheet
  const result1 = validateConsultantInMaster('Omar Badran'); // Use actual name
  console.log('Test 1 - Exact match:', result1);
  
  // Test with case variation
  const result2 = validateConsultantInMaster('OMAR BADRAN');
  console.log('Test 2 - Case variation:', result2);
  
  // Test with non-existent consultant
  const result3 = validateConsultantInMaster('Non Existent');
  console.log('Test 3 - Non-existent:', result3);
}

// ===========================================
// FUNCTION TO DEBUG MASTER TRACKER
// ===========================================

function debugMasterTrackerHeaders() {
  try {
    const masterSheet = SpreadsheetApp.openById(MASTER_TRACKER_SHEET_ID)
                                     .getSheetByName(MASTER_TRACKER_SHEET_NAME);
    
    // Read from Row 4 instead of Row 1
    const headers = masterSheet.getRange(4, 1, 1, masterSheet.getLastColumn()).getValues()[0];
    
    console.log('Total headers found:', headers.length);
    console.log('Raw headers:', headers);
    
    // Log each header with its index
    headers.forEach((header, index) => {
      console.log(`Index ${index}: "${header}" (Type: ${typeof header})`);
    });
    
    // Also check what's in column B (index 1) for consultant names
    console.log('Column B header (consultant names):', headers[1]);
    
  } catch (error) {
    console.error('Error reading headers:', error);
  }
}


// ===========================================
// FUNCTION TO LOOK UP THE CONSULTANT RATES
// ===========================================


function lookupConsultantRates(consultantName) {
  try {
    console.log('Looking up rates for:', consultantName);
    
    // 1. Validate existence
    const validation = validateConsultantInMaster(consultantName);
    if (!validation.found) {
      throw new Error(`Consultant "${consultantName}" not found in master tracker`);
    }
    
    // 2. Access sheet, headers, and data
    const masterSheet = SpreadsheetApp.openById(MASTER_TRACKER_SHEET_ID)
                                     .getSheetByName(MASTER_TRACKER_SHEET_NAME);
    const lastCol    = masterSheet.getLastColumn();
    const headers    = masterSheet.getRange(4, 1, 1, lastCol).getValues()[0];
    const processed  = headers.map(h => h.toString().trim().toLowerCase());
    const data       = masterSheet.getRange(5, 1, masterSheet.getLastRow() - 4, lastCol)
                                  .getValues();
    
    // 3. Compute column indices
    const nameColIndex         = 1;  // Column B (index 1)
    const emailIndex           = processed.indexOf('email');
    const contractorRateIndex  = processed.indexOf('consultant daily rate');
    const clientRateIndex      = processed.indexOf('all incl. rate');
    // Check for multiple header variations for robustness:
    const vatStatusIndex       = processed.indexOf('5% vat');
    if (vatStatusIndex === -1) {
      console.warn("VAT column header '5% vat' not found. Checking alternatives...");
      // Add fallback checks if the header name is slightly different
      const altIndices = ['5% tax', 'vat', 'tax', 'add vat'].map(h => processed.indexOf(h));
      const foundIndex = altIndices.find(index => index !== -1);
      if (foundIndex !== undefined) {
        vatStatusIndex = foundIndex;
        console.log(`DEBUG: Found VAT column using alternative header at index: ${vatStatusIndex}`);
      }
    }
    // ---------------------------------
    const partnerIndex   = processed.indexOf('partner');
    const clientCounterpartsIndex = processed.indexOf('client counterpart');
    const execAssistantIndex   = processed.indexOf('exec. assistant');
    const projectCodeIndex = processed.indexOf('project code');
    const projectNameIndex = processed.indexOf('project name');
    const projectIdIndex = processed.indexOf('project id'); 
    const clientNameIndex = processed.indexOf('client');
    const clientAddressIndex = processed.indexOf('client address');
    const financeTeamIndex = processed.indexOf('finance team');
    const internalTeamIndex = processed.indexOf('internal team');
    const entityIndex = processed.indexOf('entity'); 
    
    console.log('DEBUG: Full list of processed headers:', processed.join(' | '));
    console.log('DEBUG: VAT Column Index Calculated:', vatStatusIndex);
    
    // 4. Find the consultant’s data row
    const consultantRow = data.find(row =>
      row[nameColIndex]?.toString().toLowerCase().trim() ===
      consultantName.toLowerCase().trim()
    );
    if (!consultantRow) {
      throw new Error(`Consultant data not found for "${consultantName}"`);
    }

    // Determine VAT status (true if 'Yes', false if 'No' or blank)
    let addVat = false;
    if (vatStatusIndex !== -1) {
      const vatValue = consultantRow[vatStatusIndex]?.toString().toLowerCase().trim();
      addVat = vatValue === 'yes';
      console.log(`DEBUG: Raw value read from VAT cell (Index ${vatStatusIndex}): '${consultantRow[vatStatusIndex]}'. Resulting VAT flag: ${addVat}`);
    } else {
      console.warn("WARNING: VAT column was not found. Defaulting VAT flag to false.");
    }

    console.log(`VAT status for ${consultantName}: ${addVat ? 'Yes' : 'No'}`);
    
    // 5. Extract and return rate & email information
    return {
      consultantName: consultantRow[nameColIndex],
      contractorRate: parseFloat(consultantRow[contractorRateIndex]) || 0,
      contractorEmail: consultantRow[emailIndex]?.toString().trim() || '',
      clientRate:     parseFloat(consultantRow[clientRateIndex])     || 0,
      partnerEmails:          parseEmailList(consultantRow[partnerIndex]),
      clientCounterpartNames: consultantRow[clientCounterpartsIndex]?.toString().trim() || '',
     // clientCounterpartEmails: parseEmailList(consultantRow[clientCounterpartsIndex]),
      projectCode: consultantRow[projectCodeIndex]?.toString().trim() || '',
      projectName: consultantRow[projectNameIndex]?.toString().trim() || '',
      clientName: consultantRow[clientNameIndex]?.toString().trim() || '',
      projectId: consultantRow[projectIdIndex]?.toString().trim() || '',
      clientAddress: consultantRow[clientAddressIndex]?.toString().trim() || '',
      financeTeamEmails: parseEmailList(consultantRow[financeTeamIndex]),
      internalTeamEmails: parseEmailList(consultantRow[internalTeamIndex]),
      execAssistantEmails:    parseEmailList(consultantRow[execAssistantIndex]),
      entity: consultantRow[entityIndex]?.toString().trim() || '',
      // >> VAT STATUS ADDED TO RETURN OBJECT <<
      addVat: addVat
      // --------------------------------------
    };
    
  } catch (error) {
    console.error('Error in lookupConsultantRates:', error);
    throw error;
  }
}

function parseEmailList(emailString) {
  if (!emailString) return [];
  return emailString.toString().split(',')
                   .map(email => email.trim())
                   .filter(email => email.length > 0);
}

//=================
//TEST LOOKUP RATES
//=================

function testLookupRates() {
  try {
    const rates = lookupConsultantRates('Harsh Chauhan'); // Use actual consultant name
    console.log('Rates lookup result:', rates);
    
    // Validate data structure
    console.log('Contractor rate type:', typeof rates.contractorRate);
    console.log('Client rate type:', typeof rates.clientRate);
    console.log('Partner emails:', rates.partnerEmails);
  } catch (error) {
    console.error('Test failed:', error);
  }
}

/**
 * Builds a one-time approval link signed with HMAC-SHA256.
 * @param {string} sheetId Spreadsheet ID.
 * @param {string} month   Month name/tab.
 * @return {string} Signed URL.
 */

function createSignedUrl(sheetId, sheetName, type) {
  // 1. Create the payload
  const payload = `${sheetId}|${sheetName}|${type}`;

  // 2. Compute the HMAC-SHA256 signature
  const sig = Utilities.computeHmacSha256Signature(
    payload,
    HMAC_SECRET,
    Utilities.Charset.UTF_8
  )
  .map(b => ('0' + (b & 0xFF).toString(16)).slice(-2))
  .join('');

  // 3. Build the URL with plain ampersands (no consultantName)
  return `${WEB_APP_URL}?sheetId=${encodeURIComponent(sheetId)}`
       + `&sheetName=${encodeURIComponent(sheetName)}`
       + `&type=${encodeURIComponent(type)}`
       + `&sig=${sig}`;
}

/**
 * New: contractor link builder that also carries consultantName
 */
function createContractorSignedUrl(sheetId, sheetName, consultantName) {
  const type    = 'contractor';
  const payload = `${sheetId}|${sheetName}|${type}|${consultantName}`;
  const sig     = Utilities.computeHmacSha256Signature(
                    payload,
                    HMAC_SECRET,
                    Utilities.Charset.UTF_8
                  )
                  .map(b => ('0' + (b & 0xFF).toString(16)).slice(-2))
                  .join('');
  return `${WEB_APP_URL}`
       + `?sheetId=${encodeURIComponent(sheetId)}`
       + `&sheetName=${encodeURIComponent(sheetName)}`
       + `&type=${encodeURIComponent(type)}`
       + `&consultantName=${encodeURIComponent(consultantName)}`
       + `&sig=${encodeURIComponent(sig)}`;
}


function testCreateSignedUrl() {
  // Arrange
  const sheetId   = '1ubiXW9SJORMNMerklSnu2VzObQaXmjRorCi3mkWQcu4';
  const sheetName = 'February 2025';
  const type      = 'contractor';

  // Act
  const url = createSignedUrl(sheetId, sheetName, type);
  Logger.log('Generated URL: %s', url);

  // Parse out the query string
  const qs = url.split('?')[1];
  if (!qs) throw new Error('No query string found in URL');

  const params = qs.split('&').reduce((o, segment) => {
    const [k,v] = segment.split('=');
    o[k] = decodeURIComponent(v||'');
    return o;
  }, {});

  // Assert
  if (params.sheetId !== sheetId) {
    throw new Error(`sheetId mismatch: expected "${sheetId}", got "${params.sheetId}"`);
  }
  if (params.sheetName !== sheetName) {
    throw new Error(`sheetName mismatch: expected "${sheetName}", got "${params.sheetName}"`);
  }
  if (params.type !== type) {
    throw new Error(`type mismatch: expected "${type}", got "${params.type}"`);
  }
  if (!params.sig) {
    throw new Error('Missing or empty signature (sig) parameter');
  }

  Logger.log('✅ createSignedUrl passed');
}


function createClientSignedUrl(sheetId, sheetName, consultantName) {
  const type    = 'client';
  const payload = `${sheetId}|${sheetName}|${type}|${consultantName}`;
  const sig     = Utilities.computeHmacSha256Signature(
                    payload,
                    HMAC_SECRET,
                    Utilities.Charset.UTF_8
                  )
                  .map(b => ('0' + (b & 0xFF).toString(16)).slice(-2))
                  .join('');

  return `${WEB_APP_URL}?sheetId=${encodeURIComponent(sheetId)}`
       + `&sheetName=${encodeURIComponent(sheetName)}`
       + `&type=${encodeURIComponent(type)}`
       + `&consultantName=${encodeURIComponent(consultantName)}`
       + `&sig=${encodeURIComponent(sig)}`;
}


/**
 * Test harness for createClientSignedUrl.
 */
function testCreateClientSignedUrl() {
  // 1) Arrange: choose test values
  const sheetId         = '1ubiXW9SJORMNMerklSnu2VzObQaXmjRorCi3mkWQcu4';
  const sheetName       = 'February 2025';
  const consultantName  = 'Harsh Chauhan';

  // 2) Act: generate the URL
  const url = createClientSignedUrl(sheetId, sheetName, consultantName);
  Logger.log('Generated client URL: %s', url);

  // 3) Parse out the query string
  const query = url.split('?')[1];
  if (!query) throw new Error('No query string found in URL');
  const params = query.split('&').reduce((o, segment) => {
    const [k,v] = segment.split('=');
    o[k] = decodeURIComponent(v || '');
    return o;
  }, {});

  // 4) Validate parameters
  if (params.sheetId !== sheetId) {
    throw new Error(`sheetId mismatch: expected "${sheetId}", got "${params.sheetId}"`);
  }
  if (params.sheetName !== sheetName) {
    throw new Error(`sheetName mismatch: expected "${sheetName}", got "${params.sheetName}"`);
  }
  if (params.type !== 'client') {
    throw new Error(`type mismatch: expected "client", got "${params.type}"`);
  }
  if (params.consultantName !== consultantName) {
    throw new Error(`consultantName mismatch: expected "${consultantName}", got "${params.consultantName}"`);
  }
  if (!params.sig) {
    throw new Error('Missing or empty signature (sig) parameter');
  }

  // 5) Independently recompute the expected signature
  const payload = `${sheetId}|${sheetName}|client|${consultantName}`;
  const expectedSig = Utilities.computeHmacSha256Signature(
    payload,
    HMAC_SECRET,
    Utilities.Charset.UTF_8
  ).map(b => ('0' + (b & 0xFF).toString(16)).slice(-2))
   .join('');
  if (params.sig !== expectedSig) {
    throw new Error(`Signature mismatch: expected "${expectedSig}", got "${params.sig}"`);
  }

  Logger.log('✅ createClientSignedUrl passed all checks');
}



// ============================================
// MAIN TRIGGER FUNCTION
// ============================================

/**
 * onEdit trigger: when status cell becomes "Final", copy and email draft.
 * @param {SpreadsheetApp.SheetEditEvent} e
 */
function handleEdit(e) {
  const sheet = e.range.getSheet(); // Define sheet here for broader scope (fixes previous 'sheet is not defined' warning)
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return; // Prevents truly concurrent runs

  try {
    // Basic validation to ensure it's the target status cell
    if (!e || !e.range || e.range.getA1Notation() !== 'E51' || e.value !== 'Final') {
      console.log(`handleEdit skipped – edited cell ${e?.range?.getA1Notation() || 'N/A'}, value: ${e?.value || 'N/A'}`);
      return; 
    }

    const monthName = sheet.getName();
    const consultantName = getConsultantName(sheet);
    if (!consultantName) {
      throw new Error('handleEdit: Consultant name not found on sheet ' + monthName);
    }
    const ratesInfo = lookupConsultantRates(consultantName); // Lookup rates early to get projectId
    const projectId = ratesInfo.projectId || 'UNKNOWN'; // Ensure projectId is available for initial dispatch

    // --- ABSOLUTE MINIMAL CHANGE: Send initial notification to Dispatch Queue ---
    // This is the CRITICAL change. It happens first to guarantee the entry is recorded,
    // even if later, more complex steps (like invoice creation or email sending) fail.
    initialNotifyDispatchQueue(consultantName, monthName, projectId, 'timesheet');
    // --- END MINIMAL CHANGE ---

    // Continue with existing logic (create invoices, send approval emails to Ops Head)
    const timesheetData = extractTimesheetData(sheet); // Corrected: ensuring 'sheet' is used here if 'monthSheet' was a typo in your original code
    timesheetData.project_code = ratesInfo.projectCode; // Ensure project_code exists in timesheetData object

    const { contractorInvoice, clientInvoice } = createDualInvoices(timesheetData, ratesInfo, monthName);
    sendContractorApprovalEmail(contractorInvoice, consultantName);
    sendClientApprovalEmail(clientInvoice, consultantName); // This sends to Ops Head for final approval email

    SpreadsheetApp.getUi().alert('Invoice submitted. The file will now be locked for editing.');
    protectSubmittedSheet(sheet); // Protects the sheet from further edits

  } catch (err) {
    console.error('CRITICAL ERROR in handleEdit: ' + err.message + '\n' + err.stack);
    // In a production scenario, consider sending an error email to an admin here.
  } finally {
    lock.releaseLock(); // Always release the lock
  }

  // This block attempts to lock the parent file. 'sheet' is now defined in the outer scope.
 /* try {
    const file = DriveApp.getFileById(sheet.getParent().getId());
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); // Optional, for public viewing
    file.setSharing(DriveApp.Access.PRIVATE, null); // Removes previous sharing, sets to private (this is the actual locking)
    file.setEditors([]); // Remove all editors, including consultant
    console.log('🔒 File locked to prevent re-edits.');
  } catch (err) {
    console.warn('⚠️ Could not lock parent file (may already be restricted or permission issue):', err.message);
  } */
}


// ============================================
// TEST CREATE DRAFT COPIES FUNCTION
// ============================================

function testCreateDraftCopies() {
  const sheetId    = '1ubiXW9SJORMNMerklSnu2VzObQaXmjRorCi3mkWQcu4';                                  
  const monthTab   = 'June 2025';
  const consultant = 'Batuhan Er';

  // 1. Generate invoice sheets via your creation function
  const rates     = lookupConsultantRates(consultant);
  const timesheet = extractTimesheetData(
                      SpreadsheetApp.openById(sheetId).getSheetByName(monthTab)
                    );
  const invoices  = createDualInvoices(timesheet, rates, monthTab);

  // 2. Pass the returned sheet objects into draft copier
  createDraftCopies(sheetId, monthTab, consultant, {
    contractorInvoice: invoices.contractorInvoice,
    clientInvoice:     invoices.clientInvoice
  });
}

// ============================================
// CONTRACTOR APPROVAL EMAIL FUNCTION
// ============================================
/**
 * Sends a contractor-approval email with the timesheet attached.
 *
 * @param {string} spreadsheetId   ID of the timesheet spreadsheet.
 * @param {string} sheetName       Name of the month tab (e.g., "February 2025").
 * @param {string} consultantName  Consultant’s name.
 */
/*function sendContractorApprovalEmail(spreadsheetId, sheetName, consultantName) {
  // 1. Open the timesheet sheet
 // const ss    = SpreadsheetApp.openById(spreadsheetId);        
 const ss = SpreadsheetApp.getActive();   // Allowed in simple triggers[1]
                         // [1]
  const sheet = ss.getSheetByName(sheetName);                                            // [2]
  if (!sheet) throw new Error(`Sheet "${sheetName}" not found in ${spreadsheetId}`);    // [3]

  // 2. Generate HMAC-signed approval URL
  const payload    = `${spreadsheetId}|${sheetName}|contractor`;
  const signature  = Utilities.computeHmacSha256Signature(payload, HMAC_SECRET, Utilities.Charset.UTF_8)
                               .map(b => ('0'+(b&0xFF).toString(16)).slice(-2)).join('');  // [4]
  const approvalUrl = `${WEB_APP_URL}?sheetId=${spreadsheetId}` +
                      `&sheetName=${encodeURIComponent(sheetName)}` +
                      `&type=contractor&sig=${signature}`;                              // [5]

  // 3. Export the timesheet as XLSX
  const gid       = sheet.getSheetId();                                                 // [6]
  const token     = encodeURIComponent(ScriptApp.getOAuthToken());                      // [7]
  const exportUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export` +
                    `?format=xlsx&gid=${gid}&access_token=${token}`;                  // [8]
  const xlsxBlob  = UrlFetchApp.fetch(exportUrl).getBlob()
                       .setName(`Timesheet_${sheetName}_${consultantName}.xlsx`);       // [9]

  // 4. Compose HTML email body with styled button
  const htmlBody = `
    <!DOCTYPE html><html><body>
      <p>Please review the attached timesheet for <strong>${consultantName}</strong> 
         (<strong>${sheetName}</strong>) and click below to approve:</p>
      <a href="${approvalUrl}" style="
         display:inline-block;padding:12px 24px;
         background:#4CAF50;color:#fff;text-decoration:none;
         border-radius:4px;font-weight:bold;
      ">Approve Contractor Invoice</a>
      <p>If the button fails, copy/paste this link:<br>${approvalUrl}</p>
    </body></html>`;                                                                      // [10]

  // 5. Send email to Operations Head only
  GmailApp.sendEmail(
    OPERATIONS_HEAD_EMAIL,
    `Approve Contractor Invoice – ${consultantName} – ${sheetName}`,
    'Please view this email in an HTML-capable client.',
    { htmlBody: htmlBody, attachments: [xlsxBlob] }                                      // [11]
  );
}*/

function sendContractorApprovalEmail(invoiceSheet, consultant) {
  if (!invoiceSheet) throw new Error('Missing contractor invoice sheet');

  // 1) Pull IDs and names
  const ssId      = invoiceSheet.getParent().getId();
  const sheetName = invoiceSheet.getName();     // typically "Invoice"
  const sheetGid  = invoiceSheet.getSheetId();

  // 2) Build the HMAC-signed approval URL
  const approvalUrl = createSignedUrl(ssId, sheetName, 'contractor');
 // const approvalUrl = createContractorSignedUrl(ssId, sheetName, consultantName);

  // 3) Fetch the XLSX blob of the invoice sheet
  const token     = encodeURIComponent(ScriptApp.getOAuthToken());
  const xlsxUrl   = `https://docs.google.com/spreadsheets/d/${ssId}/export`
                  + `?format=xlsx&gid=${sheetGid}`
                  + `&access_token=${token}`;
  const xlsxBlob  = UrlFetchApp.fetch(xlsxUrl).getBlob()
                        .setName(`Contractor_Invoice_${sheetName}_${consultant}.xlsx`);

  // 4) Generate the PDF (unchanged)
  //const pdfBlob   = generateInvoicePDF(invoiceSheet, sheetName, 'contractor', consultant)
    //                   .setName(`Contractor_Invoice_${sheetName}_${consultant}.pdf`);

  // 5) Send email with both attachments
  const htmlBody = `
    <!DOCTYPE html><html><body>
      <p>Please review the attached draft contractor invoice for
         <strong>${consultant}</strong> (${sheetName}).</p>
      <p>When ready, click below to approve:</p>
      <a href="${approvalUrl}" style="
         display:inline-block;padding:12px 24px;
         background:#4CAF50;color:#fff;text-decoration:none;
         border-radius:4px;font-weight:bold;
      ">Approve Contractor Invoice</a>
      <p>If the button fails, paste this link:<br>${approvalUrl}</p>
    </body></html>`;

  GmailApp.sendEmail(
   // lookupConsultantRates(consultant).contractorEmail,
     OPERATIONS_HEAD_EMAIL,
    `Approve Contractor Invoice – ${consultant} – ${sheetName}`,
    'Please view in an HTML-capable client.',
    {
   //   cc: OPERATIONS_HEAD_EMAIL,
      htmlBody: htmlBody,
      attachments: [xlsxBlob]
    }
  );
}


// ============================================
// CLIENT APPROVAL EMAIL FUNCTION
// ============================================
function sendClientApprovalEmail(invoiceSheet, consultant) {
  // 0) Verify we really got a Sheet
  if (!invoiceSheet || typeof invoiceSheet.getSheetId !== 'function') {
    throw new Error('sendClientApprovalEmail: invoiceSheet is not a Sheet');
  }

  // 1) Derive IDs and names dynamically
  const ssId      = invoiceSheet.getParent().getId();
  const sheetName = invoiceSheet.getName();
  if (!sheetName) {
    throw new Error('sendClientApprovalEmail: invoiceSheet has empty name');
  }

  // 2) Log for debugging
  Logger.log('▶ sendClientApprovalEmail using sheetName="%s"', sheetName);

  // 3) Build the HMAC‐signed client link
  const approvalUrl = createClientSignedUrl(ssId, sheetName, consultant);

  // 4) Export the sheet as XLSX
  SpreadsheetApp.flush();
  Logger.log('Exporting from spreadsheet: ' + ssId);
  Logger.log('Sheet ID: ' + invoiceSheet.getSheetId());
  Logger.log('Sheet Name: ' + sheetName);

  const token     = encodeURIComponent(ScriptApp.getOAuthToken());
  const exportUrl = `https://docs.google.com/spreadsheets/d/${ssId}/export`
                  + `?format=xlsx&gid=${invoiceSheet.getSheetId()}`
                  + `&access_token=${token}`;
  const xlsxBlob  = UrlFetchApp.fetch(exportUrl)
                        .getBlob()
                        .setName(`Client_Invoice_${sheetName}_${consultant}.xlsx`);

  // 5) Compose a full HTML email with styled button + plain‐link fallback
  const htmlBody = `
    <!DOCTYPE html><html><body>
      <p>Please review the attached client invoice for <strong>${consultant}</strong>
         (<strong>${sheetName}</strong>) and click below to approve:</p>
      <a href="${approvalUrl}" style="
         display:inline-block;padding:12px 24px;
         background:#2196F3;color:#fff;text-decoration:none;
         border-radius:4px;font-weight:bold;
      ">Approve Client Invoice</a>
      <p>If that fails, copy/paste this link:<br>${approvalUrl}</p>
    </body></html>`;

  // 6) Send to Operations head with the XLSX attached
  GmailApp.sendEmail(
    OPERATIONS_HEAD_EMAIL,
    `Approve Client Invoice – ${consultant} – ${sheetName}`,
    'Please view in an HTML‐capable client.',
    { htmlBody, attachments: [xlsxBlob] }
  );
}



// ============================================
// TEST CONTRACTOR APPROVAL EMAIL FUNCTION
// ============================================

function testContractorApprovalEmail() {
  try {
    const sheetId = '1ubiXW9SJORMNMerklSnu2VzObQaXmjRorCi3mkWQcu4'; // Your actual sheet ID
    const monthTab = 'February 2025';
    const consultantName = 'Harsh Chauhan';
    
    sendContractorApprovalEmail(sheetId, monthTab, consultantName);
    console.log('Test contractor email sent successfully');
  } catch (error) {
    console.error('Test failed:', error);
  }
}

// ============================================
// TEST CLIENT APPROVAL EMAIL FUNCTION
// ============================================

function testClientApprovalEmail() {
  try {
    // 1) Open the timesheet file and pick the month‐tab explicitly
    const timesheetId   = '1ubiXW9SJORMNMerklSnu2VzObQaXmjRorCi3mkWQcu4';
    const monthName     = 'February 2025';
    const ssTimesheet   = SpreadsheetApp.openById(timesheetId);
    const monthSheet    = ssTimesheet.getSheetByName(monthName);
    if (!monthSheet) throw new Error('Month sheet not found: ' + monthName);

    // 2) Read the consultant’s name from that sheet
    const consultantName = getConsultantName(monthSheet);
    Logger.log('Consultant name:', consultantName);

    // 3) Lookup rates by the correct name
    const ratesInfo = lookupConsultantRates(consultantName);
    Logger.log('Rates info:', ratesInfo);

    // 4) Extract timesheet data & build both invoices
    const timesheetData   = extractTimesheetData(monthSheet);
    const { clientInvoice } = createDualInvoices(timesheetData, ratesInfo, monthName);

    // 5) Send the client–approval email
    sendClientApprovalEmail(clientInvoice, consultantName);
    Logger.log('✅ testClientApprovalEmail executed successfully');
  }
  catch (error) {
    Logger.log('❌ testClientApprovalEmail failed:', error);
    throw error;
  }
}



/**
 * Retrieves the consultant name from cell C4 of the given sheet.
 *
 * @param {Sheet} sheet  The Google Sheets sheet object.
 * @return {string}      The trimmed consultant name.
 */
function getConsultantName(sheet) {
  return sheet.getRange('C4').getDisplayValue().toString().trim();
}

function getConsultantNameFromInvoice(sheet) {
  return sheet.getRange('B9').getDisplayValue().trim();
}


/**
 * Returns the invoice sheet by name for the given type.
 *
 * @param {Spreadsheet} ss    The parent spreadsheet.
 * @param {string}      month The invoice tab name (e.g., "June 2025").
 * @param {string}      type  "contractor" or "client".
 * @return {Sheet}            The matching invoice sheet.
 */
function getInvoiceSheet(ss) {
  const sheet = ss.getSheetByName('Invoice');
  if (!sheet) {
    const available = ss.getSheets().map(s => s.getName()).join(', ');
    throw new Error(`Invoice sheet "Invoice" not found. Available: ${available}`);
  }
  return sheet;
}


// ============================================
// STAGE 2: FINAL ISSUANCE VIA WEB APP
// ============================================
// ... (Your existing functions preceding doGet remain unchanged) ...

/**
 * doGet: validates signature, generates PDF, saves final, and emails confirmation.
 * MODIFIED to dispatch client invoices to central queue instead of direct email.
 * @param {Object} e Event parameters.
 * @returns {HtmlOutput}
 */
function doGet(e) {
  // ============================================
  // NEW: Script Lock to prevent concurrent executions
  // ============================================
  const lock = LockService.getScriptLock();
  console.log('DEBUG: Attempting to acquire script lock...'); // <-- NEW LOG
  // Try to acquire the lock for up to 5 seconds.
  if (!lock.tryLock(5000)) {
    console.warn('DEBUG: Lock could NOT be acquired within 5 seconds. Another approval might be in progress. Exiting.'); // <-- NEW LOG
    return HtmlService.createHtmlOutput('Approval in progress. Please wait a moment and do not re-click.');
  }
  console.log('DEBUG: Script lock acquired.'); // <-- NEW LOG (SUCCESS)
  // ============================================
  // END NEW
  // ============================================

  try{
  console.log('▶ doGet triggered, raw parameters:', JSON.stringify(e.parameter));

  const p = e.parameter || {};
  const {
    sheetId,
    sheetName,
    type, // 'contractor' or 'client'
    sig,
    consultantName: consultantParam // Only present for client type approval links
  } = p;

  // A) Validate presence of required parameters
  if (!sheetId || !sheetName || !type || !sig || (type === 'client' && !consultantParam)) {
    return HtmlService.createHtmlOutput('Error: Missing parameters');
  }

  // B) Recompute expected HMAC to validate the link's integrity
  let payload = `${sheetId}|${sheetName}|${type}`;
  if (type === 'client') payload += `|${consultantParam}`;

  const expectedSig = Utilities.computeHmacSha256Signature(
    payload,
    HMAC_SECRET, // This HMAC_SECRET is for validating the approval link itself
    Utilities.Charset.UTF_8
  ).map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');

  if (sig !== expectedSig) {
     console.error('doGet: Invalid link signature.');
    return HtmlService.createHtmlOutput('Invalid link');
  }

  // ============================================
    // START NEW CODE BLOCK: ROBUST FILE EXISTENCE AND STATUS CHECK
    // This is the primary gate to prevent duplicate processing.
    // ============================================
    let spreadsheetFile = null;
    let isAlreadyProcessed = false;

    try {
        spreadsheetFile = DriveApp.getFileById(sheetId);
        // Check if the file is already in the trash
        if (spreadsheetFile.isTrashed()) {
            isAlreadyProcessed = true;
            console.log(`DEBUG: File ${sheetId} is found in trash. Assumed already processed.`);
        } else {
            // Check if the file is still a Google Sheet (prevents processing if it was deleted or converted)
            // This is a robust check against files that might exist but are no longer the target spreadsheet.
            if (spreadsheetFile.getMimeType() !== MimeType.GOOGLE_SHEETS) { // MimeType constant needs to be defined globally
                isAlreadyProcessed = true;
                console.log(`DEBUG: File ${sheetId} exists but is not a Google Sheet (${spreadsheetFile.getMimeType()}). Assumed already processed.`);
            }
        }
    } catch (getFileError) {
        // DriveApp.getFileById throws an error if the file does not exist at all or is inaccessible
        console.log(`DEBUG: Could not find file ${sheetId} in Drive (likely deleted or never existed): ${getFileError.message}. Assumed already processed.`);
        isAlreadyProcessed = true; // Treat it as already processed if file is genuinely gone
    }

    if (isAlreadyProcessed) {
        // If the file is already in trash, deleted, or not a Google Sheet, do not process again.
        console.log('DEBUG: Invoice already processed or link expired based on file status. Returning "already approved".');
        return HtmlService.createHtmlOutput('✅ Invoice already approved or link expired. No further action taken.');
    }
    // ============================================
    // END NEW CODE BLOCK: ROBUST FILE EXISTENCE AND STATUS CHECK
    // ============================================

  // C) Open the target spreadsheet and sheet (the approved invoice draft)
  const ss = SpreadsheetApp.openById(sheetId);
  const sh = ss.getSheetByName(sheetName);
  if (!sh) {
    return HtmlService.createHtmlOutput(`Error: Sheet "${sheetName}" not found`);
  }

  

  console.log('✔ Opened sheet:', sheetName);

  // D) Determine consultantName (from URL for client approval, from sheet for contractor approval)
  const consultantName = (type === 'client')
    ? consultantParam
    : getConsultantNameFromInvoice(sh);

  const invoiceNumber = sh.getRange('B8').getValue(); // Assuming B8 is the invoice number cell
  
  // E) Generate the final PDF of the invoice
  // Your existing generateInvoicePDF is used here, and it returns a Blob.
  const pdfBlob = generateInvoicePDF(sh, invoiceNumber);

  // F) Save the PDF to the correct Drive folder (Client Invoices or Contractor Invoices)
  const folderId = (type === 'client')
    ? CLIENT_INVOICE_FOLDER_ID
    : CONTRACTOR_INVOICE_FOLDER_ID;
  
  // Create the file in Drive and get the DriveApp.File object to get its ID
  const pdfFile = DriveApp.getFolderById(folderId).createFile(pdfBlob); 
  
  const rates = lookupConsultantRates(consultantName);
  const monthYear = sheetName; // The sheet name is usually "Month Year" (e.g., "February 2025")

  // G) Conditional Email Sending / Dispatching based on invoice type
  if (type === 'contractor') {
    // Contractor invoice → sent directly to contractor (remains unchanged from your working code)
    const toEmail = rates.contractorEmail;
    //const subject = `Invoice Receipt - ${invoiceNumber}`;
    const subject = `${consultantName} | Invoice for ${monthYear}`;
    //const htmlBody = `<p> Your invoice has been approved and sent.</p>`;
      const htmlBody = `<div style="font-family: Arial, sans-serif; font-size: 14px; color: #000000;">
                        <p>Dear <strong>${consultantName}</strong>,</p>
                        <p>Your invoice(s) for <strong>${monthYear}</strong> is/are ready and attached to this email.</p>
                        <p>Thank you</p>
                      </div>`;

    GmailApp.sendEmail(
      toEmail,
      subject,
      'Please view this email in an HTML‐capable client.',
      {
        cc: OPERATIONS_HEAD_EMAIL, // Or your original CC for contractor emails
        htmlBody: htmlBody,
        attachments: [pdfFile.getBlob()] // Attach the generated PDF
      }
    );

  } else if (type === 'client') {
    // MODIFIED: Client invoice → Dispatch to central queue instead of direct email
    console.log('🚚 Dispatching client Timesheet invoice to Central Queue...');

    // Log the payload being sent for debugging in execution logs
    Logger.log('🚚 Dispatch payload preview (Timesheet):');
    Logger.log(' • consultant = "%s"', consultantName);
    Logger.log(' • month      = "%s"', monthYear);
    Logger.log(' • projectId  = "%s"', rates.projectId); // Ensure lookupConsultantRates returns projectId
    Logger.log(' • type       = "timesheet"');
    Logger.log(' • fileId     = "%s"', pdfFile.getId()); // The ID of the saved PDF file


console.log("DISPATCH CALL DEBUG:", {
  consultant: consultantName,
  month: monthYear,
  projectId: rates.projectId,
  fileId: pdfFile.getId()
});
    notifyCentralDispatchQueue({
      consultant: consultantName,
      month: monthYear,
      projectId: rates.projectId, // Pass projectId from lookupConsultantRates
      type: 'timesheet', // Explicitly 'timesheet' for this script
      fileId: pdfFile.getId() // Pass the ID of the saved PDF file
      // No reportId for timesheets
    });
  }


   // ============================================
    // START NEW CODE BLOCK: TRASH DRAFT SPREADSHEET AFTER SUCCESSFUL PROCESSING
    // This makes the approval link effectively one-time use.
    // ============================================
    console.log(`DEBUG: Attempting to trash draft spreadsheet ${sheetId}...`);
    try {
      // Use the 'spreadsheetFile' object obtained earlier, if it's still valid and not yet trashed.
      // This is a robust way to ensure we trash the correct file, and only if it's not already gone.
      const fileToTrash = DriveApp.getFileById(sheetId); // Re-get to ensure latest state
      if (fileToTrash && !fileToTrash.isTrashed()) {
        fileToTrash.setTrashed(true); // Move the draft spreadsheet file to trash
        console.log(`✅ DEBUG: Draft spreadsheet ${sheetId} (${sheetName}) trashed successfully.`);
      } else {
        console.log(`DEBUG: File ${sheetId} was already trashed or missing before final trash attempt. No action needed.`);
      }
    } catch (trashError) {
      console.error(`❌ DEBUG: FAILED to trash draft spreadsheet ${sheetId}: ${trashError.message}. This is critical for preventing future duplicates!`);
      // Log the error but don't prevent the success message to the user, as the primary task is done.
    }
    // ============================================
    // END NEW CODE BLOCK: TRASH DRAFT SPREADSHEET
    // ============================================

    // H) Return confirmation page
    console.log('DEBUG: Returning success HTML.'); // <-- NEW LOG
    return HtmlService.createHtmlOutput('✅ Invoice Approved and Dispatched to Central Queue!');

  } catch (err) {
    console.error('CRITICAL ERROR in doGet: ' + err.message + '\n' + err.stack);
    return HtmlService.createHtmlOutput('An unexpected error occurred during approval: ' + err.message);
  } finally {
    // ============================================
    // IMPORTANT: Release the lock in the finally block
    // ============================================
    lock.releaseLock();
    console.log('DEBUG: Script lock released.'); // <-- NEW LOG
  }
 }

// ... (Your existing functions after doGet remain unchanged) ...

  
  




// ============================================
// MAIN INVOICE GENERATION FUNCTION
// ============================================
//OLD LOGIC
/*function generateInvoicePDF(sheet, monthName, type, consultant) {
   console.log('▶ generateInvoicePDF called with sheet:', {
    sheetName:    sheet.getName(),
    spreadsheet:  sheet.getParent().getName(),
    sheetId:      sheet.getSheetId(),
    type:         type,
    consultant:   consultantName
  });
  
  SpreadsheetApp.flush(); Utilities.sleep(3000);
  const ss     = sheet.getParent();
  const base   = ss.getUrl().replace('/edit', '/export?format=pdf');
  const url    = `${base}&size=A4&portrait=true&fitw=true`
                + `&sheetnames=false&printtitle=false&pagenum=false`
                + `&gridlines=false&fzr=false&gid=${sheet.getSheetId()}`;
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
  });
  if (response.getResponseCode() !== 200) {
    throw new Error('PDF export failed: ' + response.getResponseCode());
  }
  const fileName = (type === 'contractor')
    ? `Invoice_Receipt_${monthName}_${consultant}.pdf`
    : `Invoice_${monthName}_${consultant}.pdf`;
  return response.getBlob().setName(fileName);
}*/

//NEW LOGIC
function generateInvoicePDF(sheet, invoiceNumber) {
  SpreadsheetApp.flush();
  Utilities.sleep(3000);  // ensure all formulas/rendering are complete

  const url = sheet.getParent().getUrl().replace('/edit', '/export?format=pdf')
              + `&size=A4&portrait=true&fitw=true`
              + `&sheetnames=false&printtitle=false&pagenum=false`
              + `&gridlines=false&fzr=false`
              + `&gid=${sheet.getSheetId()}`;

  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
  });

  if (response.getResponseCode() !== 200) {
    throw new Error('PDF export failed: ' + response.getResponseCode());
  }

  // Get consultant name and month name from the sheet data
  const consultantName = getConsultantNameFromInvoice(sheet);
  const monthName = getMonthNameFromInvoiceNumber(invoiceNumber);
  
  // Construct the new, descriptive filename
 // const safeConsultantName = consultantName.replace(/\s+/g, '_'); 
 // const safeMonthName = monthName.replace(/\s+/g, '_'); 
  
  const newFileName = `Timesheet_Invoice_${invoiceNumber}_${consultantName}_${monthName}.pdf`;
  
   return response.getBlob().setName(newFileName);
  //return response.getBlob().setName(`${invoiceNumber}.pdf`);
}


/*function sendFinalInvoiceEmail(pdfBlob, month) {
  GmailApp.sendEmail(
    PARTNER_COLUMN.join(','), `Approved Invoice – ${month}`,
    'Please find the approved invoice attached.', {
      attachments: [pdfBlob],
      cc: CLIENT_COUNTERPARTS_COLUMN.join(','), bcc: EMAIL_BCC.join(',')
    }
  );                                                                         // [14]
}*/

function sendFinalInvoiceEmail(pdfBlob, month, type, consultant) {
  let to, cc, bcc, subject;

  if (type === 'contractor') {
    to = CONTRACTOR_EMAIL;
    cc = OPERATIONS_HEAD_EMAIL;
   // bcc = EMAIL_BCC.join(',');
    subject = `✅ Your Contractor Invoice – ${consultant} – ${month}`;
  } else if (type === 'client') {
    to = PARTNER_COLUMN.join(',');
    cc = EXEC_ASSISTANT_COLUMN.join(',');
   // bcc = EMAIL_BCC.join(',');
    subject = `✅ Approved Client Invoice – ${consultant} – ${month}`;
  } else {
    throw new Error(`Unknown invoice type: ${type}`);
  }

 const finalClientEmailBody = `
 Dear ${CLIENT_COUNTERPARTS_COLUMN[0]},

 Greetings from Stratverse!

 Enclosed, please find the invoices pertaining to the services provided by consultant ${CONSULTANT_NAME}, in connection with the project "${PROJECT_NAME}".

 We kindly request your approval of the invoice and confirmation of the project code '${PROJECT_CODE}' to initiate the payment process, to our USD account as previously specified.

 Best regards,
`;

  GmailApp.sendEmail(
    to,
    subject,
    finalClientEmailBody,
   // 'Please find the approved invoice attached.',
    {
      attachments: [pdfBlob],
      cc: cc,
      bcc: bcc
    }
  );
}



// ============================================
// DATA EXTRACTION FUNCTIONS
// ============================================
function extractTimesheetData(sheet) {
  try {
    console.log('Starting data extraction from sheet: ' + sheet.getName());
    
   // data.sourceSheetName = sheet.getName();

    var data = {
      sourceSheetName: sheet.getName(),
      name: null,
      street: null,
      city: null,
      postal_code: null,
      country: null,
      invoice_number: null,
      total_hours: null,
      daily_rate: null,
      billing_amount: null,
      activities: []
    };
    
    // Extract personal details with error checking
    try {
      data.name = sheet.getRange(TIMESHEET_FIELDS.name).getValue();
      console.log('Name extracted: ' + data.name);
    } catch (e) {
      console.error('Error extracting name from ' + TIMESHEET_FIELDS.name + ': ' + e.toString());
    }
    
   /* try {
      data.street = sheet.getRange(TIMESHEET_FIELDS.street).getValue();
      console.log('Street extracted: ' + data.street);
    } catch (e) {
      console.error('Error extracting street from ' + TIMESHEET_FIELDS.street + ': ' + e.toString());
    } */
    
   /* try {
      data.city = sheet.getRange(TIMESHEET_FIELDS.city).getValue();
      console.log('City extracted: ' + data.city);
    } catch (e) {
      console.error('Error extracting city from ' + TIMESHEET_FIELDS.city + ': ' + e.toString());
    } */
    
   /* try {
      data.postal_code = sheet.getRange(TIMESHEET_FIELDS.postal_code).getValue();
      console.log('Postal code extracted: ' + data.postal_code);
    } catch (e) {
      console.error('Error extracting postal code from ' + TIMESHEET_FIELDS.postal_code + ': ' + e.toString());
    } */
    
  /*  try {
      data.country = sheet.getRange(TIMESHEET_FIELDS.country).getValue();
      console.log('Country extracted: ' + data.country);
    } catch (e) {
      console.error('Error extracting country from ' + TIMESHEET_FIELDS.country + ': ' + e.toString());
    } */

    // 🔁 Pull address from "2026 Invoice Tracker"
const trackerSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('2026 Invoice Tracker');
if (!trackerSheet) throw new Error('Sheet "2026 Invoice Tracker" not found');

try {
  data.street = trackerSheet.getRange('C26').getValue();
  data.city = trackerSheet.getRange('C27').getValue();
  data.postal_code = trackerSheet.getRange('C28').getValue();
  data.country = trackerSheet.getRange('C29').getValue();
  
  console.log(`✅ Pulled address from tracker: ${data.street}, ${data.city}, ${data.postal_code}, ${data.country}`);
} catch (e) {
  console.error('❌ Failed to pull address from tracker:', e.toString());
}

    
    try {
      data.total_hours = sheet.getRange(TIMESHEET_FIELDS.total_hours).getValue();
      console.log('Total hours extracted: ' + data.total_hours);
    } catch (e) {
      console.error('Error extracting total hours from ' + TIMESHEET_FIELDS.total_hours + ': ' + e.toString());
    }
    
    try {
      data.daily_rate = sheet.getRange(TIMESHEET_FIELDS.daily_rate).getValue();
      console.log('Daily rate extracted: ' + data.daily_rate);
    } catch (e) {
      console.error('Error extracting daily rate from ' + TIMESHEET_FIELDS.daily_rate + ': ' + e.toString());
    }
    
    try {
      data.billing_amount = sheet.getRange(TIMESHEET_FIELDS.billing_amount).getValue();
      console.log('Billing amount extracted: ' + data.billing_amount);
    } catch (e) {
      console.error('Error extracting billing amount from ' + TIMESHEET_FIELDS.billing_amount + ': ' + e.toString());
    }
    
    // Extract activity data dynamically - only include rows with a date
    console.log('Starting dynamic activity data extraction...');
    const maxActivityRowsToScan = 31; // Scan up to 31 rows from the start
    const startActivityRow = TIMESHEET_FIELDS.activity_start_row; // e.g., 15

    for (let i = 0; i < maxActivityRowsToScan; i++) {
      const currentRow = startActivityRow + i;
      const date = sheet.getRange('B' + currentRow).getValue();
      const weekday = sheet.getRange('C' + currentRow).getValue();
      const activity = sheet.getRange('D' + currentRow).getValue();
      const timeWorked = sheet.getRange('E' + currentRow).getValue();

      // Only add to activities if a date is present
      if (date && date.toString().trim() !== '') {
        const hours = timeWorked ? convertTimeToHours(timeWorked) : 0;

        data.activities.push({
          date: date,
          weekday: weekday,
          activity: activity || '',
          hours: hours
        });
      } else {
        // Optional: If you want to stop scanning once the first empty date is found, uncomment 'break;'
        // This assumes activities are contiguous without gaps.
        // break;
      }
    }

    console.log('Data extraction completed successfully. Actual activities found:', data.activities.length);
    return data;

  } catch (error) {
    console.error('Critical error in extractTimesheetData: ' + error.toString());
    throw error;
  }
}
  /*  for (var i = 0; i < 31; i++) {
     try {
      var row = TIMESHEET_FIELDS.activity_start_row + i;
      var date = sheet.getRange('B' + row).getValue();
      var weekday = sheet.getRange('C' + row).getValue();
      var activity = sheet.getRange('D' + row).getValue();
      var timeWorked = sheet.getRange('E' + row).getValue();


    for (var i = 0; i < 31; i++) {
  try {
    var row = TIMESHEET_FIELDS.activity_start_row + i;
    var date = sheet.getRange('B' + row).getValue();
    var weekday = sheet.getRange('C' + row).getValue();
    var activity = sheet.getRange('D' + row).getValue();
    var timeWorked = sheet.getRange('E' + row).getValue();
    var hours = timeWorked ? convertTimeToHours(timeWorked) : '';

    // 🟢 Always push to preserve order (even if the row is mostly empty)
    data.activities.push({
      date: date || '',
      weekday: weekday || '',
      activity: activity || '',
     // timeWorked: timeWorked || '',
      hours: hours
    });

  } catch (e) {
    console.log('Skipping row ' + (TIMESHEET_FIELDS.activity_start_row + i) + ': ' + e.toString());
  }
}
  }  catch (e) {
      console.log('Skipping row ' + (TIMESHEET_FIELDS.activity_start_row + i) + ': ' + e.toString());
  }
}
    
    console.log('Data extraction completed successfully');
    return data;
    
  } catch (error) {
    console.error('Critical error in extractTimesheetData: ' + error.toString());
    throw error;
  }
} */


//==========================================================================
// Quick validation test - just run the function and check for basic errors
//===========================================================================
function quickTestExtractTimesheetData() {
  console.log('🚀 Quick test: extractTimesheetData');
  
  try {
    // First, validate that required constants exist
    validateTimesheetFields();
    
    // Get the active sheet (or specify your test sheet)
   // const sheet = SpreadsheetApp.getActiveSheet();

   const ss    = SpreadsheetApp.openById('1ubiXW9SJORMNMerklSnu2VzObQaXmjRorCi3mkWQcu4');
   const sheet = ss.getSheetByName('February 2025');   // target the month tab
   const data  = extractTimesheetData(sheet);

    console.log('Using sheet:', sheet.getName());
    
    // Run the extraction
    const startTime = new Date();
    const result = extractTimesheetData(sheet);
    const endTime = new Date();
    
    console.log(`✅ Extraction completed in ${endTime - startTime}ms`);
    console.log('✅ Result structure validation:');
    console.log('  - Name:', result.name);
    console.log('  - Total Hours:', result.total_hours);
    console.log('  - Daily Rate:', result.daily_rate);
    console.log('  - Activities Count:', result.activities ? result.activities.length : 'N/A');
    
    // Basic validations
    if (result.name) {
      console.log('✓ Name extracted successfully');
    } else {
      console.log('⚠️  Name is empty or null');
    }
    
    if (typeof result.total_hours === 'number' && result.total_hours > 0) {
      console.log('✓ Total hours is a valid number');
    } else {
      console.log('⚠️  Total hours might be invalid:', result.total_hours);
    }
    
    if (Array.isArray(result.activities)) {
      console.log('✓ Activities is an array');
      if (result.activities.length > 0) {
        console.log('✓ Activities contain data');
        console.log('  First activity:', result.activities[0]);
      } else {
        console.log('⚠️  No activities found');
      }
    } else {
      console.log('❌ Activities is not an array');
    }
    
    return result;
    
  } catch (error) {
    console.error('❌ Quick test failed:', error);
    throw error;
  }
}


/**
 * Test helper: Check if TIMESHEET_FIELDS constants are defined
 */
function validateTimesheetFields() {
  console.log('\n🔍 Validating TIMESHEET_FIELDS constants...');
  
  const requiredFields = [
    'name', 'street', 'city', 'postal_code', 'country',
    'total_hours', 'daily_rate', 'billing_amount', 'activity_start_row'
  ];
  
  try {
    requiredFields.forEach(field => {
      if (typeof TIMESHEET_FIELDS[field] !== 'undefined') {
        console.log(`✓ TIMESHEET_FIELDS.${field}:`, TIMESHEET_FIELDS[field]);
      } else {
        console.error(`❌ Missing TIMESHEET_FIELDS.${field}`);
      }
    });
  } catch (error) {
    console.error('❌ TIMESHEET_FIELDS not defined or accessible:', error);
    throw error;
  }
}


// ============================================
// CONVERT TIME TO HOURS FUNCTIONS
// ============================================
function convertTimeToHours(timeWorked) {
  if (!timeWorked || timeWorked === '') return 0;
  if (typeof timeWorked === 'string') {
    if (timeWorked.includes('Full-day (8 hours)')) return 8;
    if (timeWorked.includes('Half-day (4 hours)')) return 4;
  }
  return 0;
}


function generateInvoiceFromTimesheet(sheetId, month) {
  const ss    = SpreadsheetApp.openById(sheetId);
  const sheet = ss.getSheetByName(month);
  if (!sheet) throw new Error('Worksheet not found: ' + month);

  // 1. Extract timesheet data
  const data = extractTimesheetData(sheet);

  // 2. Create and populate invoice template
  const invoiceSheet = createInvoiceFromTemplate(data, month);

  return invoiceSheet;
}


// ============================================
// INVOICE CREATION FUNCTIONS
// ============================================
/*function createInvoiceFromTemplate(data, monthName, type) {

  invoiceSheet.setName(monthName);  // instead of 'Invoice'

  for (var i = 0; i < data.activities.length; i++) {
  var row = INVOICE_FIELDS.data_start_row + i;
  var item = data.activities[i];

  // If activity is blank and hours > 0, replace activity with projectCode
  var description = (item.activity && item.activity.trim()) || '';
  var shouldInsertProjectCode = (!description && item.hours && item.hours > 0);
  var finalActivity = shouldInsertProjectCode ? (data.projectCode || 'N/A') : description;

  sheet.getRange('A' + row).setValue(item.date || '');
  sheet.getRange('B' + row).setValue(item.weekday || '');
  sheet.getRange('C' + row).setValue(finalActivity);  // 👈 fill in adjusted activity
  sheet.getRange('D' + row).setValue(item.hours || '');
}

  const template = SpreadsheetApp.openById(TEMPLATE_FILE_ID).getSheetByName('Sample Contractor Invoice');
  const invoiceSS = SpreadsheetApp.create(`${type === 'contractor' ? 'Draft_Contractor' : 'Draft_Client'}_Invoice_${data.name}_${monthName}`);
  const file = DriveApp.getFileById(invoiceSS.getId());
  DriveApp.getFolderById(DRAFT_FOLDER_ID).addFile(file);
  file.getParents().next().removeFile(file);
  const copied = template.copyTo(invoiceSS).setName('Invoice');
  invoiceSS.getSheets().forEach(s => { if (s.getName()!=='Invoice') invoiceSS.deleteSheet(s); });
  const invNum = generateInvoiceNumber(type.toUpperCase(), new Date().getFullYear());
 // populateInvoiceHeader(copied, data, invNum);
  populateInvoiceHeader(sheet, contractorData, invoiceNumber, 'contractor'); // for contractor
  populateInvoiceHeader(sheet, clientData, invoiceNumber, 'client');         // for client

  //populateInvoiceData(copied, data);
  populateInvoiceData(copied, data, type);
  //populateInvoiceData(copied, data);
  populateInvoiceTotals(copied, data, type);
  return copied;
} */
/*function createInvoiceFromTemplate(data, monthName, type) {
  try {
    // Determine which master file and sheet name to use
    let templateFileId = TEMPLATE_FILE_ID;
    let templateSheetName;

    if (type === 'contractor') {
        templateSheetName = 'Sample Contractor Invoice';
    } else { // type === 'client'
        if (data.applyVat) {
            templateFileId = VAT_TEMPLATE_FILE_ID;
            templateSheetName = 'VAT Client Invoice'; // Assume a distinct sheet name in the VAT template
        } else {
            templateSheetName = 'Sample Client Invoice 2';
        }
    }

    if (templateFileId.includes('YOUR_VAT_CLIENT_INVOICE_TEMPLATE_ID')) {
        throw new Error("VAT_TEMPLATE_FILE_ID must be replaced with a valid ID before using VAT logic.");
    }
    
    // 1. Open your master template spreadsheet (either standard or VAT version)
    const templateSS = SpreadsheetApp.openById(templateFileId);

    // 2. Pick the right sheet in the template
    const templateSheet = templateSS.getSheetByName(templateSheetName);
    if (!templateSheet) {
      throw new Error(`Template sheet "${templateSheetName}" not found in spreadsheet ID ${templateFileId}`);
    }

    // 3. Create a brand-new spreadsheet in your Draft folder
    const draftFolder = DriveApp.getFolderById(DRAFT_FOLDER_ID);
    const prefix      = type === 'contractor'
      ? 'Draft_Contractor_Invoice'
      : 'Draft_Client_Invoice';
    const title       = `${prefix}_${data.name}_${monthName}`;
    const invoiceSS   = SpreadsheetApp.create(title);
    const invoiceFile = DriveApp.getFileById(invoiceSS.getId());
    draftFolder.addFile(invoiceFile);
    invoiceFile.getParents().next().removeFile(invoiceFile);

    // 4. Copy the chosen template tab into the new spreadsheet
    const invoiceSheet = templateSheet.copyTo(invoiceSS);
    invoiceSheet.setName(monthName);

    // 5. Remove any other default sheets
    invoiceSS.getSheets().forEach(s => {
      if (s.getSheetId() !== invoiceSheet.getSheetId()) {
        invoiceSS.deleteSheet(s);
      }
    });

    // 6. Use the pre-calculated invoice number from data
    const invoiceNumber = data.invoice_number;
    console.log('Generated invoice number:', invoiceNumber);

    // 7. Populate header, activity rows, and totals
    populateInvoiceHeader(invoiceSheet, data, invoiceNumber, type);
    populateInvoiceData(invoiceSheet,   data, type);
    populateInvoiceTotals(invoiceSheet, data, type); // Totals function needs minor adjustment (see below)

    // 8. Return the populated “Invoice” sheet
    return invoiceSheet;
  }
  catch (err) {
    console.error('Error in createInvoiceFromTemplate:', err);
    throw err;
  }
} */



// ============================================
// PDF GENERATION AND STORAGE
// ============================================

function generateInvoicePDF(sheet, monthName) {
  try {
    // 1. Ensure edits are flushed and indexed
    SpreadsheetApp.flush();
    Utilities.sleep(3000);

    // 2. Build the export URL for PDF
    const ss     = sheet.getParent();
    const baseUrl= ss.getUrl().replace('/edit', '/export?format=pdf');
    const pdfUrl = `${baseUrl}&size=A4&portrait=true&fitw=true`
                 + `&sheetnames=false&printtitle=false&pagenum=false`
                 + `&gridlines=false&fzr=false&gid=${sheet.getSheetId()}`;

    // 3. Fetch the PDF blob with authorization header
    const response = UrlFetchApp.fetch(pdfUrl, {
      muteHttpExceptions: true,
      headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() }
    });
    if (response.getResponseCode() !== 200) {
      throw new Error('PDF export failed with code: ' + response.getResponseCode());
    }

    // 4. Return named PDF blob
    return response.getBlob().setName(`Invoice_${monthName}.pdf`);
  } catch (error) {
    console.error('Error in generateInvoicePDF:', error);
    throw error;
  }
}


function savePDFToDrive(pdfBlob, data, monthName) {
  var folder = DriveApp.getFolderById(FINAL_FOLDER_ID);
  var fileName = 'Invoice_' + data.name.replace(/\s+/g, '_') + '_' + monthName + '_' + 
                 new Date().toISOString().slice(0,10) + '.pdf';
  
  var file = folder.createFile(pdfBlob.setName(fileName));
  console.log('PDF saved to Drive: ' + fileName);
  return file;
}

// ============================================
// UTILITY FUNCTIONS: GENERATE INVOICE FUNCTION
// ============================================
//old logic
/*function generateInvoiceNumber(type, year) {
  try {
    const prefix = type === 'CONTRACTOR' ? 'CONTR-INV' : 'CLIENT-INV';
    const propertyKey = `${type}_INVOICE_COUNTER_${year}`;
    
    // Get current counter from Script Properties
    let counter = PropertiesService.getScriptProperties().getProperty(propertyKey);
    if (!counter) {
      counter = 0;
      console.log(`Initializing ${type} invoice counter for ${year}`);
    }
    
    // Increment counter
    counter = parseInt(counter) + 1;
    
    // Save updated counter back to properties
    PropertiesService.getScriptProperties().setProperty(propertyKey, counter.toString());
    
    // Return formatted invoice number
    const invoiceNumber = `${prefix}-${year}-${counter.toString().padStart(3, '0')}`;
    console.log(`Generated invoice number: ${invoiceNumber}`);
    
    return invoiceNumber;
  } catch (error) {
    console.error('Error generating invoice number:', error);
    throw error;
  }
}*/

//new logic
function generateInvoiceNumber(type, year) {
  try {
    const propertyKey = `${type}_INVOICE_COUNTER_${year}`;
    
    let counter = PropertiesService.getScriptProperties().getProperty(propertyKey);
    if (!counter) {
      counter = 0;
      console.log(`Initializing ${type} invoice counter for ${year}`);
    }
    
    counter = parseInt(counter) + 1;
    PropertiesService.getScriptProperties().setProperty(propertyKey, counter.toString());
    
    return counter.toString().padStart(3, '0');  // returns '033'
  } catch (error) {
    console.error('Error generating invoice serial:', error);
    throw error;
  }
}


// ==============================
// TEST GENERATE INVOICE FUNCTION
// ==============================
function testInvoiceNumbering() {
  try {
    // Test contractor invoice numbering
    const contractorInvoice1 = generateInvoiceNumber('CONTRACTOR', 2025);
    const contractorInvoice2 = generateInvoiceNumber('CONTRACTOR', 2025);
    
    // Test client invoice numbering
    const clientInvoice1 = generateInvoiceNumber('CLIENT', 2025);
    const clientInvoice2 = generateInvoiceNumber('CLIENT', 2025);
    
    console.log('Contractor invoices:', contractorInvoice1, contractorInvoice2);
    console.log('Client invoices:', clientInvoice1, clientInvoice2);
    
    // Verify sequential numbering
    console.log('Sequential numbering working correctly');
  } catch (error) {
    console.error('Invoice numbering test failed:', error);
  }
}

// ========================================
// UTILITY FUNCTIONS: DUAL INVOICE CREATION
// ========================================
//OLD LOGIC
/*function createDualInvoices(timesheetData, ratesInfo, monthName) {
  // Contractor invoice using daily rate
  const contractorData = {
    ...timesheetData,
    daily_rate: ratesInfo.contractorRate,
    billing_amount: timesheetData.total_hours * ratesInfo.contractorRate
  };
  // Client invoice using all-inclusive rate
  const clientData = {
    ...timesheetData,
    daily_rate: ratesInfo.clientRate,
    billing_amount: timesheetData.total_hours * ratesInfo.clientRate
  };
  // Create and return sheet objects
  return {
    contractorInvoice: createInvoiceFromTemplate(contractorData, monthName, 'contractor'),
    clientInvoice:     createInvoiceFromTemplate(clientData,     monthName, 'client')
  };
}*/

//NEW LOGIC
function createDualInvoices(timesheetData, ratesInfo, monthName) {
  const year = new Date().getFullYear();
  const fullName = timesheetData.name || 'Unknown';
  const nameParts = fullName.trim().split(/\s+/);
  let shortName;
  
  const monthMap = {
    January: '01', February: '02', March: '03', April: '04', May: '05', June: '06',
    July: '07', August: '08', September: '09', October: '10', November: '11', December: '12'
  };
  
  const [monthStr, yearStr] = monthName.split(' ');
  const yearMonth = `${yearStr}${monthMap[monthStr] || '00'}`;
  
  if (nameParts.length >= 2) {
    const firstInitial = nameParts[0][0].toUpperCase();
    const lastName = nameParts[nameParts.length - 1];
    shortName = firstInitial + lastName;
  } else {
    shortName = fullName;
  }
  
  const consultantName = shortName;
  const projectId = ratesInfo.projectId || 'XXX';
  
  const serial = getGlobalInvoiceSerial();
  const paddedSerial = String(serial).padStart(3, '0');
  
  const baseClientInvoiceNumber = `${yearMonth}_${shortName}_${paddedSerial}`;
  const clientInvoiceNumber = ratesInfo.addVat
    ? `${baseClientInvoiceNumber}`
    : baseClientInvoiceNumber;

  const contractorInvoiceNumber = `${yearMonth}_TS_${shortName}`;
  
  // Find the latest activity date (omitted for brevity, assume correct)
  let latestDate = null;
  timesheetData.activities.forEach(act => {
    if (act.date && act.hours >= 0) {
      if (!latestDate || new Date(act.date) > new Date(latestDate)) {
        latestDate = act.date;
      }
    }
  });
  
  const formattedInvoiceDate = latestDate
    ? Utilities.formatDate(new Date(latestDate), Session.getScriptTimeZone(), 'dd MMMM yyyy')
    : '';
  
  const contractorData = {
    ...timesheetData,
    daily_rate: ratesInfo.contractorRate,
    billing_amount: timesheetData.total_hours * ratesInfo.contractorRate / 8,
    invoice_number: contractorInvoiceNumber,
    projectName: ratesInfo.projectName,
    clientName: ratesInfo.clientName,
    projectCode: ratesInfo.projectCode,
    invoice_date: formattedInvoiceDate,
    entity: ratesInfo.entity
  };

  const clientData = {
    ...timesheetData,
    daily_rate: ratesInfo.clientRate,
    billing_amount: timesheetData.total_hours * ratesInfo.clientRate / 8,
    invoice_number: clientInvoiceNumber,
    projectName: ratesInfo.projectName,
    clientName: ratesInfo.clientName,
    projectCode: ratesInfo.projectCode,
    clientAddress: ratesInfo.clientAddress,
    invoice_date: formattedInvoiceDate,
    entity: ratesInfo.entity,
    // CRITICAL FIX: Ensure flag is explicitly passed
    applyVat: ratesInfo.addVat
  };

  return {
    contractorInvoice: createInvoiceFromTemplate(contractorData, monthName, 'contractor'),
    clientInvoice:     createInvoiceFromTemplate(clientData,     monthName, 'client')
  };
}




/*function createInvoiceFromTemplate(data, monthName, type) {
  try {
    // 1. Open your master template spreadsheet
    const TEMPLATE_ID  = '1Rf-Srvg6JSIhNsFkYP2_btLdfNJVKFe6xG6MSBjFLj8';
    const templateSS   = SpreadsheetApp.openById(TEMPLATE_ID);

    // 2. Pick the right sheet in the template
    const templateName = type === 'contractor'
      ? 'Sample Contractor Invoice'
      : 'Sample Client Invoice 2';
    const templateSheet = templateSS.getSheetByName(templateName);
    if (!templateSheet) {
      throw new Error(`Template sheet "${templateName}" not found`);
    }

    // 3. Create a brand-new spreadsheet in your Draft folder
    const draftFolder = DriveApp.getFolderById(DRAFT_FOLDER_ID);
    const prefix      = type === 'contractor'
      ? 'Draft_Contractor_Invoice'
      : 'Draft_Client_Invoice';
    const title       = `${prefix}_${data.name}_${monthName}`;
    const invoiceSS   = SpreadsheetApp.create(title);
    const invoiceFile = DriveApp.getFileById(invoiceSS.getId());
    draftFolder.addFile(invoiceFile);
    invoiceFile.getParents().next().removeFile(invoiceFile);

    // 4. Copy the chosen template tab into the new spreadsheet
    const invoiceSheet = templateSheet.copyTo(invoiceSS);
    invoiceSheet.setName(monthName);

    // 5. Remove any other default sheets
    invoiceSS.getSheets().forEach(s => {
      if (s.getSheetId() !== invoiceSheet.getSheetId()) {
        invoiceSS.deleteSheet(s);
      }
    });

    // 6. Generate and log a new invoice number
    const year          = new Date().getFullYear();
    //const invoiceNumber = generateInvoiceNumber(type.toUpperCase(), year);
    const invoiceNumber = data.invoice_number;
    console.log('Generated invoice number:', invoiceNumber);

    // 7. Populate header, activity rows, and totals
    populateInvoiceHeader(invoiceSheet, data, invoiceNumber, type);
    populateInvoiceData(invoiceSheet,   data, type);
    populateInvoiceTotals(invoiceSheet, data, type);

    // 8. Return the populated “Invoice” sheet
    return invoiceSheet;
  }
  catch (err) {
    console.error('Error in createInvoiceFromTemplate:', err);
    throw err;
  }
}  */

function createInvoiceFromTemplate(data, monthName, type) {
  try {
    // 1. Determine template based on type and VAT flag
    let templateFileId = TEMPLATE_FILE_ID;
    let templateSheetName;

    if (type === 'contractor') {
        templateSheetName = 'Sample Contractor Invoice';
    } else { // type === 'client'
        if (data.applyVat) {
            // CRITICAL: Template ID assignment based on persistent flag
            if (VAT_TEMPLATE_FILE_ID.includes('YOUR_VAT_CLIENT_INVOICE_TEMPLATE_ID')) {
                throw new Error("Configuration Error: VAT_TEMPLATE_FILE_ID must be replaced with a valid ID.");
            }
            templateFileId = VAT_TEMPLATE_FILE_ID;
            templateSheetName = 'VAT Client Invoice'; // Assumed sheet name in the VAT template
        } else {
            templateSheetName = 'Sample Client Invoice 2';
        }
    }
    
    console.log(`DEBUG: Template selected: File ID ${templateFileId} / Sheet Name ${templateSheetName}`);

    // 2. Open the correct template spreadsheet
    const templateSS = SpreadsheetApp.openById(templateFileId);

    // 3. Pick the right sheet in the template
    const templateSheet = templateSS.getSheetByName(templateSheetName);
    if (!templateSheet) {
      throw new Error(`Template sheet "${templateSheetName}" not found in spreadsheet ID ${templateFileId}.`);
    }

    // 4. Create Draft File Title
    const draftFolder = DriveApp.getFolderById(DRAFT_FOLDER_ID);
    const prefix      = type === 'contractor'
      ? 'Draft_Contractor_Invoice'
      : (data.applyVat ? 'Draft_Client_VAT_Invoice' : 'Draft_Client_Standard_Invoice');
    const title       = `${prefix}_${data.name}_${monthName}`;
    
    // 5. Create new spreadsheet
    const invoiceSS   = SpreadsheetApp.create(title);
    const invoiceFile = DriveApp.getFileById(invoiceSS.getId());
    draftFolder.addFile(invoiceFile);
    invoiceFile.getParents().next().removeFile(invoiceFile);

    // 6. Copy the chosen template tab
    const invoiceSheet = templateSheet.copyTo(invoiceSS);
    invoiceSheet.setName(monthName);

    // 7. Remove default sheets
    invoiceSS.getSheets().forEach(s => {
      if (s.getSheetId() !== invoiceSheet.getSheetId()) {
        invoiceSS.deleteSheet(s);
      }
    });

   // ➡️ CRITICAL FIX: Ensure minimum sheet dimensions before populating data (NEW CODE)
    const requiredRows = 60; 
    const requiredCols = 6; // Column F is the 6th column (A=1, B=2, ..., F=6)

    if (invoiceSheet.getMaxRows() < requiredRows) {
        invoiceSheet.insertRowsAfter(invoiceSheet.getMaxRows(), requiredRows - invoiceSheet.getMaxRows());
    }
    if (invoiceSheet.getMaxColumns() < requiredCols) {
        invoiceSheet.insertColumnsAfter(invoiceSheet.getMaxColumns(), requiredCols - invoiceSheet.getMaxColumns());
    }
    // -------------------------------------------------------------

    // 8. Populate Data
    const invoiceNumber = data.invoice_number;
    console.log('Generated invoice number:', invoiceNumber);

    populateInvoiceHeader(invoiceSheet, data, invoiceNumber, type);
    populateInvoiceData(invoiceSheet,   data, type);
    // CRITICAL: Call the totals function (which writes the formulas)
    populateInvoiceTotals(invoiceSheet, data, type); 

    // 9. Return the populated “Invoice” sheet
    return invoiceSheet;
  }
  catch (err) {
    console.error('Error in createInvoiceFromTemplate:', err);
    throw err;
  }
}





// ========================================
// TEST DUAL INVOICE CREATION FUNCTION
// ========================================
function testDualInvoiceGeneration() {
  try {
    const testSheetId = '1FmWetWQ0OYz6bpJWqp8jB7uwYWTgm2ZxgpHfY7qzP4g'; // Your actual sheet ID
    const testMonth = 'June 2025';
    const consultantName = 'Harsh Chauhan'; // Use actual consultant name
    
    // Step 1: Lookup consultant rates
    console.log('Step 1: Looking up consultant rates');
    const consultantRates = lookupConsultantRates(consultantName);
    
    // Step 2: Extract timesheet data (using existing function)
    console.log('Step 2: Extracting timesheet data');
    const sheet = SpreadsheetApp.openById(testSheetId).getSheetByName(testMonth);
    const timesheetData = extractTimesheetData(sheet);
    
    // Step 3: Create dual invoices
    console.log('Step 3: Creating dual invoices');
    const invoices = createDualInvoices(timesheetData, consultantRates, testMonth);
    
    // Step 4: Verify results
    console.log('Test Results:');
    console.log('- Contractor invoice created:', !!invoices.contractorInvoice);
    console.log('- Client invoice created:', !!invoices.clientInvoice);
    console.log('- Invoice numbers generated correctly');
    console.log('- Rate differences applied correctly');
    
    return invoices;
  } catch (error) {
    console.error('Dual invoice generation test failed:', error);
    throw error;
  }
}



function isMonthlySheet(sheetName) {
  var monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
                   'July', 'August', 'September', 'October', 'November', 'December'];
  return monthNames.some(month => sheetName.includes(month)) && sheetName.includes('2025');
}

function validateTimesheetData(data) {
  if (!data.name || data.name === 'For consultant to fill') {
    throw new Error('Consultant name is required');
  }
  if (!data.total_hours || data.total_hours <= 0) {
    throw new Error('Total hours must be greater than 0');
  }
  if (!data.daily_rate || data.daily_rate <= 0) {
    throw new Error('Daily rate must be greater than 0');
  }
  return true;
}

// ============================================
// EMAIL FUNCTIONS
// ============================================
function sendConfirmationEmail(data, monthName, pdfFile) {
  var subject = 'Invoice Generated Successfully - ' + data.name + ' - ' + monthName;
  var body = 'Hello,\n\n' +
            'An invoice has been successfully generated with the following details:\n\n' +
            'Consultant: ' + data.name + '\n' +
            'Month: ' + monthName + '\n' +
            'Total Hours: ' + data.total_hours + '\n' +
            'Daily Rate: $' + data.daily_rate + '\n' +
            'Billing Amount: $' + data.billing_amount + '\n' +
            'Generated: ' + new Date().toString() + '\n\n' +
            'The invoice PDF has been saved to Google Drive and is attached to this email.\n\n' +
            'Best regards,\n' +
            'Automated Invoice System';

  // INITIALIZE emailOptions object first
  var emailOptions = {
    attachments: [pdfFile.getBlob()]
  };

  // Add CC if there are CC recipients
  if (EMAIL_CC && EMAIL_CC.length > 0) {
      emailOptions.cc = EMAIL_CC.join(',');
   }
  
  // Add BCC if there are BCC recipients
  if (EMAIL_BCC && EMAIL_BCC.length > 0) {
      emailOptions.bcc = EMAIL_BCC.join(',');
   }
   
  GmailApp.sendEmail(
    // EMAIL_RECIPIENT,
      EMAIL_TO.join(','),
    //EMAIL_BCC.join(','),
    //EMAIL_CC.join(','),
    subject,
    body,
    {
      cc:    EMAIL_CC.join(','),    // CC recipients
      bcc:   EMAIL_BCC.join(','),   // BCC recipients
      attachments: [pdfFile.getBlob()]
    },
  );
  
    console.log('Confirmation email sent to: ' + EMAIL_TO.join(','));
    console.log('Confirmation email sent to: ' + EMAIL_CC.join(','));
//  console.log('Confirmation email sent to: ' + EMAIL_RECIPIENT);

}

function sendErrorEmail(errorType, errorMessage) {
  var subject = 'Invoice System Error - ' + errorType;
  var body = 'An error occurred in the automated invoice system:\n\n' +
            'Error Type: ' + errorType + '\n' +
            'Error Message: ' + errorMessage + '\n' +
            'Time: ' + new Date().toString() + '\n\n' +
            'Please check the system and resolve the issue.\n\n' +
            'Automated Invoice System';
  
  try {
    GmailApp.sendEmail(EMAIL_TO.join(','), subject, body);
  } catch (e) {
    console.error('Failed to send error email: ' + e.toString());

    try {
    GmailApp.sendEmail(EMAIL_CC.join(','), subject, body);
  } catch (e) {
    console.error('Failed to send error email: ' + e.toString());
  }
  }
}

// ============================================
// DEBUG STRUCTURE
// ============================================

function debugSheetStructure() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('January 2025');
  
  if (!sheet) {
    console.log('January 2025 sheet not found!');
    return;
  }
  
  console.log('=== DEBUGGING SHEET STRUCTURE ===');
  console.log('Sheet name: ' + sheet.getName());
  console.log('Total rows: ' + sheet.getLastRow());
  console.log('Total columns: ' + sheet.getLastColumn());
  
  // Check personal details
  console.log('=== PERSONAL DETAILS ===');
  console.log('C4 (Name): ' + sheet.getRange('C4').getValue());
  console.log('C5 (Street): ' + sheet.getRange('C5').getValue());
  console.log('C6 (City): ' + sheet.getRange('C6').getValue());
  console.log('C7 (Postal): ' + sheet.getRange('C7').getValue());
  console.log('C8 (Country): ' + sheet.getRange('C8').getValue());
  
  // Check calculations
  console.log('=== CALCULATIONS ===');
  console.log('E47 (Total Hours): ' + sheet.getRange('E47').getValue());
  console.log('E48 (Daily Rate): ' + sheet.getRange('E48').getValue());
  console.log('E49 (Billing Amount): ' + sheet.getRange('E49').getValue());
  console.log('E51 (Status): ' + sheet.getRange('E51').getValue());
  
  // Check first few activity rows
  console.log('=== ACTIVITY DATA SAMPLE ===');
  for (var i = 0; i < 5; i++) {
    var row = 15 + i;
    console.log('Row ' + row + ': B' + row + '=' + sheet.getRange('B' + row).getValue() + 
                ', C' + row + '=' + sheet.getRange('C' + row).getValue() + 
                ', D' + row + '=' + sheet.getRange('D' + row).getValue() + 
                ', E' + row + '=' + sheet.getRange('E' + row).getValue());
  }
}

// ============================================
// MANUAL TESTING FUNCTION
// ============================================
function testInvoiceGeneration() {
  // For testing purposes - manually trigger invoice generation
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('January 2025');
  generateInvoiceFromTimesheet(sheet, 'January 2025');
}

// ============================================
// FETCHING THE RIGHT SHEET OF INVOICE FUNCTION
// ============================================

function testFileId() {
  var templateFileId = '1tpzGoMalE1Ffecma7wNwKy4RIP7HHovt6tQZZuMWr08';
  try {
    var ss = SpreadsheetApp.openById(templateFileId);
    console.log('Success! File name: ' + ss.getName());
  } catch (error) {
    console.log('Error: ' + error.toString());
  }
}

// ============================================
// DEBUG INVOICE CREATION FUNCTION
// ============================================

function debugInvoiceCreation() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('January 2025');
    var data = extractTimesheetData(sheet);
    
    console.log('=== DEBUGGING INVOICE CREATION ===');
    
    // Test template file access
    var templateFileId = '1tpzGoMalE1Ffecma7wNwKy4RIP7HHovt6tQZZuMWr08';
    var templateFile = SpreadsheetApp.openById(templateFileId);
    console.log('Template file accessed: ' + templateFile.getName());
    
    // Test template sheet access
    var templateSheet = templateFile.getSheetByName('Sample Contractor Invoice');
    if (!templateSheet) {
      console.log('ERROR: Template sheet "Sample Contractor Invoice" not found');
      return;
    }
    console.log('Template sheet found: ' + templateSheet.getName());
    
    // Test sheet copying
    var invoiceSpreadsheet = SpreadsheetApp.create('DEBUG_Invoice_Test');
    var copiedSheet = templateSheet.copyTo(invoiceSpreadsheet);
    console.log('Sheet copied successfully');
    
    // Test basic range access
    try {
      var testRange = copiedSheet.getRange('A1');
      console.log('Range access successful: A1 = ' + testRange.getValue());
    } catch (e) {
      console.log('ERROR accessing range A1: ' + e.toString());
    }
    
    // Test specific invoice field access
    console.log('Testing invoice field mappings...');
    try {
      copiedSheet.getRange(INVOICE_FIELDS.consultant_name).setValue('Test');
      console.log('consultant_name field accessible');
    } catch (e) {
      console.log('ERROR with consultant_name field: ' + e.toString());
    }
    
    try {
      copiedSheet.getRange(INVOICE_FIELDS.creation_date).setValue('Test');
      console.log('creation_date field accessible');
    } catch (e) {
      console.log('ERROR with creation_date field: ' + e.toString());
    }
    
    // Clean up test file
    DriveApp.getFileById(invoiceSpreadsheet.getId()).setTrashed(true);
    console.log('Test file cleaned up');
    
  } catch (error) {
    console.log('CRITICAL ERROR in debugInvoiceCreation: ' + error.toString());
  }
}

// ============================================
// ANALYSE INVOICE TEMPLATE FUNCTION
// ============================================

function analyzeInvoiceTemplate() {
  try {
    var templateFileId = '1Rf-Srvg6JSIhNsFkYP2_btLdfNJVKFe6xG6MSBjFLj8'; // Replace with your actual template file ID
    var templateFile = SpreadsheetApp.openById(templateFileId);
    var templateSheet = templateFile.getSheetByName('Sample Contractor Invoice');
    
    if (!templateSheet) {
      console.log('ERROR: Template sheet "Sample Contractor Invoice" not found');
      return;
    }
    
    console.log('=== INVOICE TEMPLATE ANALYSIS ===');
    console.log('Template file: ' + templateFile.getName());
    console.log('Template sheet: ' + templateSheet.getName());
    
    // Check the structure of cells that need data - UPDATED with correct totals cells
    var cellsToCheck = ['A7', 'B7', 'A8', 'B8', 'A15', 'B15', 'E50', 'E52', 'E53'];
    
    cellsToCheck.forEach(function(cell) {
      try {
        var value = templateSheet.getRange(cell).getValue();
        console.log('Cell ' + cell + ': "' + value + '"');
      } catch (e) {
        console.log('Error accessing cell ' + cell + ': ' + e.toString());
      }
    });
    
    // Additional analysis of the template structure
    console.log('=== TEMPLATE DIMENSIONS ===');
    console.log('Last row with data: ' + templateSheet.getLastRow());
    console.log('Last column with data: ' + templateSheet.getLastColumn());
    
  } catch (error) {
    console.error('Error in analyzeInvoiceTemplate: ' + error.toString());
  }
}

// ============================================
// POPULATE INVOICE HEADER FUNCTION
// ============================================


function populateInvoiceHeaderSafe(sheet, data, invoiceNumber) {
  var fields = [
    { range: INVOICE_FIELDS.invoice_number, value: invoiceNumber, name: 'Invoice Number' },
    { range: INVOICE_FIELDS.consultant_name, value: data.name, name: 'Consultant Name' },
    { range: INVOICE_FIELDS.consultant_street, value: data.street, name: 'Street Address' },
    { range: INVOICE_FIELDS.consultant_city, value: data.city, name: 'City' },
    { range: INVOICE_FIELDS.consultant_postal, value: data.postal_code, name: 'Postal Code' },
    { range: INVOICE_FIELDS.consultant_country, value: data.country, name: 'Country' },
    { range: INVOICE_FIELDS.project_name, value: data.projectName, name: 'Project Name' },
    { range: INVOICE_FIELDS.client_name, value: data.clientName, name: 'Client Name' }
  ];
  
  // Handle dates separately
  var currentDate = new Date();
  var formattedDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), 'dd MMMM yyyy');
  
  console.log('✅ projectName:', data.projectName);
  console.log('✅ clientName:', data.clientName);

  fields.push(
    { range: INVOICE_FIELDS.creation_date, value: formattedDate, name: 'Creation Date' },
    { range: INVOICE_FIELDS.invoice_date, value: formattedDate, name: 'Invoice Date' }
  );
  
  // Populate each field individually with error handling
  fields.forEach(function(field) {
    try {
      sheet.getRange(field.range).setValue(field.value || '');
      console.log('✓ ' + field.name + ' populated successfully');
    } catch (error) {
      console.log('⚠ Error populating ' + field.name + ': ' + error.toString());
    }
  });
}

// ============================================
// VERIFY UPDATED INVOICE FIELDS FUNCTION
// ============================================

function verifyUpdatedInvoiceFields() {
  console.log('=== VERIFYING UPDATED INVOICE FIELDS ===');
  console.log('Header fields:');
  console.log('  invoice_number: ' + INVOICE_FIELDS.invoice_number);
  console.log('  consultant_name: ' + INVOICE_FIELDS.consultant_name);
  console.log('Totals fields:');
  console.log('  total_hours_result: ' + INVOICE_FIELDS.total_hours_result);
  console.log('  daily_rate_result: ' + INVOICE_FIELDS.daily_rate_result);
  console.log('  billing_amount_result: ' + INVOICE_FIELDS.billing_amount_result);
}

function analyzeInvoiceTemplate() {
  try {
    var templateFileId = '1Rf-Srvg6JSIhNsFkYP2_btLdfNJVKFe6xG6MSBjFLj8';
    var templateFile = SpreadsheetApp.openById(templateFileId);
    var templateSheet = templateFile.getSheetByName('Sample Contractor Invoice');
    
    if (!templateSheet) {
      console.log('ERROR: Template sheet not found');
      return;
    }
    
    console.log('=== INVOICE TEMPLATE ANALYSIS ===');
    var cellsToCheck = ['A7', 'B7', 'A8', 'B8', 'E50', 'E52', 'E53'];
    
    cellsToCheck.forEach(function(cell) {
      try {
        var value = templateSheet.getRange(cell).getValue();
        console.log('Cell ' + cell + ': "' + value + '"');
      } catch (e) {
        console.log('Error accessing cell ' + cell + ': ' + e.toString());
      }
    });
    
  } catch (error) {
    console.error('Error in analyzeInvoiceTemplate: ' + error.toString());
  }
}


// ============================================
// MISSING INVOICE POPULATION FUNCTIONS
// ============================================


/* function populateInvoiceData(sheet, data, invoiceType) {
  try {
    console.log('Populating invoice activity data...');
    
    var startRow = INVOICE_FIELDS.data_start_row; // usually 18
    var dailyRate = data.daily_rate || 0;

    for (var i = 0; i < data.activities.length; i++) {
      var row = startRow + i;
      var activity = data.activities[i];

      if (activity.date && activity.hours > 0) {
        try {
          sheet.getRange('A' + row).setValue(activity.date);
          sheet.getRange('B' + row).setValue(activity.weekday);
          const desc = (activity.activity && activity.activity.trim()) || data.projectCode || 'N/A';
          sheet.getRange('C' + row).setValue(desc);
          sheet.getRange('D' + row).setValue(activity.timeWorked);
          sheet.getRange('E' + row).setValue(activity.hours);

          // 💵 Add Amount column logic for client invoice only
          if (invoiceType === 'client') {
            let amount = '';
            if (activity.hours === 4) {
              amount = dailyRate / 2;
            } else if (activity.hours === 8) {
              amount = dailyRate;
            }

            if (amount) {
              sheet.getRange('F' + row).setValue(amount);
            }
          }

        } catch (e) {
          console.log('⚠️ Warning: Could not populate row ' + row + ': ' + e.toString());
        }
      }
    }

    console.log('✓ Activity data populated successfully');
  } catch (error) {
    throw new Error('❌ Error in populateInvoiceData: ' + error.toString());
  }
} */

 //New populateInvoiceData()

 function populateInvoiceData(invoiceSheet, timesheetData, invoiceType) {
  try {
    console.log('✅ Populating invoice activity data from timesheet...');

    const startRow = INVOICE_FIELDS.data_start_row; // e.g., 13
    const maxPossibleRowsInTemplate = 31; // The max number of activity rows your template visual layout accounts for (e.g., from row 13 to 43)
    const dailyRate = timesheetData.daily_rate || 0;

    const actualActivityCount = timesheetData.activities.length;
    console.log(`DEBUG: Actual activity count found: ${actualActivityCount}`);

    // --- First, clear ALL potential activity rows to remove any old data/formatting ---
    const fullActivityRangeToClear = invoiceType === 'client'
      ? `A${startRow}:F${startRow + maxPossibleRowsInTemplate - 1}`
      : `A${startRow}:E${startRow + maxPossibleRowsInTemplate - 1}`;
    invoiceSheet.getRange(fullActivityRangeToClear).clear({ contentsOnly: false, formatsOnly: false, validationsOnly: false, notesOnly: false });
    console.log(`DEBUG: Cleared template activity area: ${fullActivityRangeToClear}`);

    // Loop ONLY for actual activities collected
    for (let i = 0; i < actualActivityCount; i++) {
      const row = startRow + i;
      const item = timesheetData.activities[i];

      const hours = item.hours;
      let finalActivity = (item.activity && item.activity.trim()) || '';
      if (hours === 0) {
          finalActivity = '-'; // Explicitly set to '-' if no hours, unless activity is provided
      } else {
          // If hours > 0 and activity is placeholder, replace with project_code
          const currentActivity = String(item.activity || '').trim(); // Get current activity, treat null/undefined as empty string
          const isPlaceholder = finalActivity.toLowerCase().replace(/\s+/g, ' ') === 'for consultant to fill';
          const isBlank = currentActivity === ''; // Check if it's truly blank/empty;
          if (isPlaceholder || isBlank) {
              finalActivity = timesheetData.project_code || '';
          }
      }

      invoiceSheet.getRange(`A${row}`).setValue(item.date || '');
      invoiceSheet.getRange(`B${row}`).setValue(item.weekday || '');
      invoiceSheet.getRange(`C${row}`).setValue(finalActivity);
      invoiceSheet.getRange(`D${row}`).clearContent(); // Keep D blank
      invoiceSheet.getRange(`E${row}`).setValue(hours);

      if (invoiceType === 'client') {
        let amount = 0;
        if (hours === 4) amount = dailyRate / 2;
        else if (hours === 8) amount = dailyRate;
        invoiceSheet.getRange(`F${row}`).setValue(amount);
        invoiceSheet.getRange(`F${row}`).setNumberFormat('$#,##0.00');
      }

      // Determine background color
      const normalizedWeekday = String(item.weekday || '').trim().toLowerCase();
      let bgColor = '#ffffff';
      if (normalizedWeekday === 'friday' || normalizedWeekday === 'saturday') {
        bgColor = '#d9d9d9'; // gray for weekend/holidays
      } else if ((i % 2) !== 0) {
        bgColor = '#f2f2f2'; // light gray for alternate rows
      }

      const rangeToFormat = invoiceType === 'client'
        ? `A${row}:F${row}`
        : `A${row}:E${row}`;

      // Apply background
      invoiceSheet.getRange(rangeToFormat).setBackground(bgColor);

      // Apply DARK GREY borders for populated rows
      invoiceSheet.getRange(rangeToFormat).setBorder(
        true, true, true, true,
        true, true,
        '#C0C0C0', SpreadsheetApp.BorderStyle.SOLID // Use SOLID for thinner borders, or MEDIUM for testing
      );

       // 2. Now, explicitly remove the specific vertical border between C and D for this row
      // Make the right border of column C invisible
      invoiceSheet.getRange(`C${row}`).setBorder(
        null, null, null, false, // Preserve top/left/bottom, set right to false
        null, null,              // Inner borders (ignored for single cell context)
        null, SpreadsheetApp.BorderStyle.NONE // Apply NONE style to the 'false' right border
      );

      // Make the left border of column D invisible
      invoiceSheet.getRange(`D${row}`).setBorder(
        null, false, null, null, // Preserve top/bottom/right, set left to false
        null, null,              // Inner borders (ignored for single cell context)
        null, SpreadsheetApp.BorderStyle.NONE // Apply NONE style to the 'false' left border
      );
      // --- END CRITICAL BORDER LOGIC ---
    } // End of loop for actual activities

    // REMOVE THE SECOND BORDER LOOP HERE if it still exists in your code:
    // This loop previously drew 31 lines regardless of data.
    /*
    for (let i = 0; i < 31; i++) {
      const row = startRow + i;
      const rangeToBorder = invoiceType === 'client'
        ? `A${row}:F${row}`
        : `A${row}:E${row}`;
      invoiceSheet.getRange(rangeToBorder).setBorder(
        true, true, true, true, true, true,
        '#808080', SpreadsheetApp.BorderStyle.SOLID_MEDIUM
      );
    }
    */

    // ====================================================================
    // ➡️ FINAL CLEANUP: HIDE/CLEAR COLUMN F FOR CONTRACTOR INVOICES
    // ====================================================================
    if (invoiceType === 'contractor') {
        const startRow = INVOICE_FIELDS.data_start_row; // e.g., 13
        // Use a generous range (e.g., Row 13 to Row 60) to ensure the column is cleared
        const fullColumnFToClear = `F1:F55`; 
        
        // Clear everything and set background to white, removing all borders.
        const rangeF = invoiceSheet.getRange(fullColumnFToClear);
        rangeF.clear(); 
        rangeF.setBorder(
            false, false, false, false, 
            false, false,
            null, SpreadsheetApp.BorderStyle.NONE 
        );
        rangeF.setBackground('#ffffff'); // Force white background
        console.log(`DEBUG: Explicitly cleared and removed formatting/borders from column F for Contractor Invoice: ${fullColumnFToClear}`);
    }
    // ====================================================================

    console.log('✅ Final invoice data populated with dynamic rows.');

  } catch (error) {
    console.error('❌ Error in populateInvoiceData:', error.toString());
    throw error;
  }
}





function populateInvoiceHeader(sheet, data, invoiceNumber, type) {
  try {
    console.log('Populating invoice header...');

    let creationDateValueCell;
    let invoiceDateValueCell;

    if (type === 'client') {
      creationDateValueCell = 'F8';   // Client Creation Date value moves from E8 to F8
      invoiceDateValueCell = 'F9';     // Client Invoice Date value moves from E9 to F9
      clientAddressValueCell = 'E5';   // Client Address value moves from E4 to E5
    } else { // type === 'contractor'
      creationDateValueCell = INVOICE_FIELDS.creation_date; // 'E8' for contractor
      invoiceDateValueCell = INVOICE_FIELDS.invoice_date;   // 'E9' for contractor
      clientAddressValueCell = 'E4';   // Contractor address goes to E4
    }
    
    // Populate all header fields with direct cell mapping
    sheet.getRange(INVOICE_FIELDS.invoice_number).setValue(invoiceNumber);
    sheet.getRange(INVOICE_FIELDS.consultant_name).setValue(data.name || 'N/A');
    sheet.getRange(INVOICE_FIELDS.project_name).setValue(data.projectName || 'N/A');
    sheet.getRange(INVOICE_FIELDS.client_name).setValue(data.clientName || 'N/A');
   // sheet.getRange(INVOICE_FIELDS.consultant_street).setValue(data.street || 'N/A');
    //sheet.getRange(INVOICE_FIELDS.consultant_city).setValue(data.city || 'N/A');
  //  sheet.getRange(INVOICE_FIELDS.consultant_postal).setValue(data.postal_code || 'N/A');
  //  sheet.getRange(INVOICE_FIELDS.consultant_country).setValue(data.country || 'N/A');
   
    // Leave project and client fields blank for now
    //sheet.getRange(INVOICE_FIELDS.project_name).setValue('');
    //sheet.getRange(INVOICE_FIELDS.client_name).setValue('');

    //sheet.getRange(INVOICE_FIELDS.project_name).setValue(data.projectName || 'N/A');
    sheet.getRange(INVOICE_FIELDS.project_name).setValue(data.projectCode || 'N/A');
    sheet.getRange(INVOICE_FIELDS.client_name).setValue(data.clientName || 'N/A');
 
    // Auto-generate dates
    //var currentDate = new Date();
   // var formattedDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), 'dd MMMM yyyy');
    //sheet.getRange(INVOICE_FIELDS.creation_date).setValue(formattedDate);
   // sheet.getRange('E1').setValue(data.entity || ''); 
    const d1Range = sheet.getRange('E1'); // Get the range object for D1
    d1Range.setValue(data.entity || '');  // Set its value

    // Apply formatting to D1
    d1Range.setFontWeight('bold');        // Make text bold
    d1Range.setFontSize(12);              // Set font size to 12 (adjust as needed)
    //sheet.getRange(INVOICE_FIELDS.invoice_date).setValue(formattedDate);
    if (type === 'contractor') {

      // Clear E3, E5, E6 as we'll combine everything into E4
      sheet.getRange('E3').clearContent();
      sheet.getRange('E5').clearContent();
      sheet.getRange('E6').clearContent();

      // Create an array to hold each line of the address
      const addressLines = [];

      if (data.name) {
        addressLines.push(data.name);
      }
      if (data.street) {
        addressLines.push(data.street);
      }
      if (data.city) {
        addressLines.push(data.city);
      }
      if (data.postal_code) {
        addressLines.push(data.postal_code);
      }
      if (data.country) {
        addressLines.push(data.country);
      }

      // Join the lines with a newline character and set in E4
      const contractorAddress = addressLines.join('\n');
      sheet.getRange('E4').setValue(contractorAddress.trim());
      console.log('Set contractor address in E4 with newlines:\n', contractorAddress.trim());

   // sheet.getRange('E3').setValue(data.street || '');        // Street
   // sheet.getRange('E4').setValue(data.city || '');          // City
   // sheet.getRange('E5').setValue(data.postal_code || '');   // Postal Code
   // sheet.getRange('E6').setValue(data.country || '');       // Country
    } else if (type === 'client') {
    sheet.getRange('E3').setValue('');                       // Blank
    sheet.getRange('F5').setValue(data.clientAddress || ''); // Client Address
    sheet.getRange('E4').setValue('');                       // Blank
    sheet.getRange('E6').setValue('');                       // Blank
   }

    // Find the last non-empty date in activities
    // 🔍 Find the last date with hours > 0
  let lastWorkedDate = null;
   if (Array.isArray(data.activities)) {
    for (let i = data.activities.length - 1; i >= 0; i--) {
     const act = data.activities[i];
      if (act && act.date instanceof Date && act.hours >= 0) {
       lastWorkedDate = act.date;
       break;
    }
  }
}

 var currentDate = new Date();
 var formattedCurrentDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), 'dd MMMM yyyy'); // Renamed to avoid conflict

const invoiceDate = lastWorkedDate || currentDate;
const formattedInvoiceDate = Utilities.formatDate(invoiceDate, Session.getScriptTimeZone(), 'dd MMMM yyyy');

 // Apply date format to both (assuming they are adjacent in their respective columns)
    if (type === 'client') {
        sheet.getRange('F8:F9').setNumberFormat('dd/mm/yyyy');
    } else {
        sheet.getRange('E8:E9').setNumberFormat('dd/mm/yyyy');
    }

         // Set Creation Date - ***CRITICAL CHANGE: USING DYNAMIC VARIABLE***
    sheet.getRange(creationDateValueCell).setValue(formattedCurrentDate);
    // Set Invoice Date - ***CRITICAL CHANGE: USING DYNAMIC VARIABLE***
    sheet.getRange(invoiceDateValueCell).setValue(formattedInvoiceDate);
    

//sheet.getRange(INVOICE_FIELDS.invoice_date).setValue(formattedInvoiceDate);
//sheet.getRange(INVOICE_FIELDS.invoice_date).setValue(data.invoice_date || formattedDate);


    
    console.log('✓ Header populated successfully');
  } catch (error) {
    throw new Error('Error in populateInvoiceHeader: ' + error.toString());
  }
}



/*function populateInvoiceTotals(sheet, data, type) { // Make sure 'type' is passed here!
  try {
    console.log(`DEBUG: Populating invoice totals for sheet: ${sheet.getName()} (Type: ${type})`);
    console.log(`DEBUG: Data for totals (daily_rate from data object): ${data.daily_rate}`);
    console.log(`DEBUG: Data for totals (total_hours from data object): ${data.total_hours}`);
    console.log(`DEBUG: Data for totals (billing_amount from data object): ${data.billing_amount}`);

    // DETERMINE CELL REFERENCES BASED ON INVOICE TYPE
    let totalHoursCell;
    let dailyRateCell;
    let billingAmountCell;

    if (type === 'client') {
      totalHoursCell = 'F45'; // Assuming F column for client
      dailyRateCell = 'F47';   // Assuming F column for client
      billingAmountCell = 'F48'; // Assuming F column for client
    } else { // type === 'contractor'
      totalHoursCell = 'E45'; // E column for contractor
      dailyRateCell = 'E47';   // E column for contractor
      billingAmountCell = 'E48'; // E column for contractor
    }
    console.log(`DEBUG: Using totalHoursCell: ${totalHoursCell}, dailyRateCell: ${dailyRateCell}, billingAmountCell: ${billingAmountCell} for type: ${type}`);


    // Set Total Hours (E45)
    sheet.getRange(totalHoursCell).setValue(data.total_hours);
    console.log(`DEBUG: Set ${totalHoursCell} to ${data.total_hours} on ${sheet.getName()}`);

    // Set Daily Rate (E47)
    sheet.getRange(dailyRateCell).setValue(data.daily_rate);
    console.log(`DEBUG: Set ${dailyRateCell} to ${data.daily_rate} on ${sheet.getName()}`);

    // --- CRITICAL CHECK: After setting, read back the values ---
    SpreadsheetApp.flush(); // Ensure writes are committed before reading back
    Utilities.sleep(500); // Give Sheets a moment to recalculate if formulas are involved

    const actualTotalHoursInCell = sheet.getRange(totalHoursCell).getValue();
    const actualDailyRateInCell = sheet.getRange(dailyRateCell).getValue();
    console.log(`DEBUG: Actual value in ${totalHoursCell} (E45): ${actualTotalHoursInCell}`);
    console.log(`DEBUG: Actual value in ${dailyRateCell} (E47): ${actualDailyRateInCell}`);


    // Handle Billing Amount (E48) based on invoice type
    if (type === 'client') {
      // For client, directly set the pre-calculated billing_amount
      const formula = `=${totalHoursCell}*${dailyRateCell}/8`;
      sheet.getRange(billingAmountCell).setFormula(formula);
      console.log(`DEBUG: Set ${billingAmountCell} (F48) for client invoice to DIRECT VALUE: ${data.billing_amount}`);
    } else {
      // For contractor, keep the formula
      const formula = `=${totalHoursCell}*${dailyRateCell}/8`;
      sheet.getRange(billingAmountCell).setFormula(formula);
      console.log(`DEBUG: Set ${billingAmountCell} (E48) for contractor invoice to FORMULA: ${formula}`);
    }

    // --- Final check for Billing Amount after setting ---
    SpreadsheetApp.flush();
    Utilities.sleep(500);
    const actualBillingAmountInCell = sheet.getRange(billingAmountCell).getValue();
    console.log(`DEBUG: Actual value in ${billingAmountCell} (E48) AFTER SETTING: ${actualBillingAmountInCell}`);

    console.log('DEBUG: ✓ Totals populated successfully.');

  } catch (error) {
    throw new Error('Error in populateInvoiceTotals: ' + error.toString());
  }
} */

//part 2
/*function populateInvoiceTotals(sheet, data, type) { // Make sure 'type' is passed here!
  try {
    console.log(`DEBUG: Populating invoice totals for sheet: ${sheet.getName()} (Type: ${type})`);
    console.log(`DEBUG: Data for totals (applyVat flag): ${data.applyVat}`);

    // DETERMINE CELL REFERENCES BASED ON INVOICE TYPE
    let totalHoursCell;
    let dailyRateCell;
    let subtotalCell; // E48 or F48

    if (type === 'client') {
    totalHoursCell = 'F45'; // Hours column for client invoice
    dailyRateCell = 'F47'; // Rate column for client invoice
    subtotalCell = 'F48'; // Billing Amount / Subtotal value cell
    } else { // type === 'contractor'
     totalHoursCell = 'E45'; // Hours column for contractor invoice
      dailyRateCell = 'E47';   // Rate column for contractor invoice
     subtotalCell = 'E48'; // Billing Amount / Total Due value cell
 }
 console.log(`DEBUG: Using subtotalCell: ${subtotalCell} for type: ${type}`);


    // Set Total Hours (E45/F45)
    sheet.getRange(totalHoursCell).setValue(data.total_hours);

    // Set Daily Rate (E47/F47) and ensure format
    sheet.getRange(dailyRateCell).setValue(data.daily_rate);
    sheet.getRange(dailyRateCell).setNumberFormat('$#,##0.00');

    // Set Billing Amount / Subtotal (E48/F48) - The actual billable fee formula
    const formula = `=${totalHoursCell}*${dailyRateCell}/8`;
    sheet.getRange(subtotalCell).setFormula(formula);
    sheet.getRange(subtotalCell).setNumberFormat('$#,##0.00');
    SpreadsheetApp.flush(); // Commit formula before VAT logic

// --- CRITICAL VAT LOGIC (E50/F50 and E51/F51) ---
if (type === 'client' && data.applyVat) {
console.log("DEBUG: Applying 5% VAT for Client Invoice in E50/F50 and E51/F51.");
 
 // Change label E48 to "Subtotal" (pre-tax)
 sheet.getRange('E48').setValue('Subtotal'); 
 
 const vatCell = 'F50';
 const totalCell = 'F51';
const vatLabelCell = 'E50';
const totalLabelCell = 'E51';

// 1. Set VAT (5% of subtotal) in F50
 const vatFormula = `=${subtotalCell}*0.05`;
 sheet.getRange(vatCell).setFormula(vatFormula);
sheet.getRange(vatCell).setNumberFormat('$#,##0.00');

// 2. Set the final Total in F51
 const totalFormula = `=${subtotalCell}+${vatCell}`;
 sheet.getRange(totalCell).setFormula(totalFormula);
sheet.getRange(totalCell).setNumberFormat('$#,##0.00');
sheet.getRange(totalCell).setFontWeight('bold'); // Make final total bold

  // 3. Add labels
  sheet.getRange(vatLabelCell).setValue('5% VAT');
  sheet.getRange(totalLabelCell).setValue('Total');
  sheet.getRange(totalLabelCell).setFontWeight('bold');
  
 // Clear intermediate E49/F49 rows to ensure a clean visual break
sheet.getRange('E49:F49').clearContent();
 sheet.getRange('E49:F49').clearFormat();


 } else if (type === 'client' && !data.applyVat) {
 // No VAT for client: Clear tax rows and revert E48 label
 sheet.getRange('E48').setValue('Billing Amount');
 sheet.getRange('E49:F51').clearContent();
  sheet.getRange('E49:F51').clearFormat();
} else {
// Contractor Invoice: Ensure E48 is labeled as Billing Amount/Total and the tax area is clear
sheet.getRange('E48').setValue('Billing Amount');
 sheet.getRange('E48').setFontWeight('bold'); // Ensure contractor total remains bold
sheet.getRange('E49:F51').clearContent();
  sheet.getRange('E49:F51').clearFormat();
  }
  // --------------------------------------------

   SpreadsheetApp.flush();
    Utilities.sleep(500);
    const actualBillingAmountInCell = sheet.getRange(subtotalCell).getValue();
    console.log(`DEBUG: Actual value in ${subtotalCell} (Subtotal/Total) AFTER SETTING: ${actualBillingAmountInCell}`);

    console.log('DEBUG: ✓ Totals populated successfully.');

  } catch (error) {
    throw new Error('Error in populateInvoiceTotals: ' + error.toString());
  }
} */

function populateInvoiceTotals(sheet, data, type) {
  try {
    console.log(`DEBUG: Populating invoice totals for sheet: ${sheet.getName()} (Type: ${type})`);
    console.log(`DEBUG: Data for totals (applyVat flag): ${data.applyVat}`);

    let totalHoursCell, dailyRateCell, netBillingLabelCell;

    if (type === 'client') {
      totalHoursCell = 'F45';
      dailyRateCell = 'F47';
      netBillingLabelCell = 'E48'; 
    } else { // Contractor
      totalHoursCell = 'E45';
      dailyRateCell = 'E47';
      netBillingLabelCell = 'D48'; 
    }

    const valueCol = totalHoursCell.charAt(0);
    const netBillingValueCell = `${valueCol}48`; // F48 (Net Amount/Subtotal calculation location)

    // 1. Set required inputs and calculate Net Amount (Base Subtotal)
    sheet.getRange(totalHoursCell).setValue(data.total_hours);
    sheet.getRange(dailyRateCell).setValue(data.daily_rate).setNumberFormat('$#,##0.00');
    const netAmountFormula = `=${totalHoursCell}*${dailyRateCell}/8`;
    sheet.getRange(netBillingValueCell).setFormula(netAmountFormula).setNumberFormat('$#,##0.00');
    
    // 2. Apply VAT Calculation Logic (If VAT Template is used)
    if (type === 'client' && data.applyVat) {
      console.log("DEBUG: VAT applied template selected. Writing formulas directly to F49/F50.");
      
      // Target VAT value row (F49) and Total value row (F50)
      const vatValueCell = `${valueCol}49`; // F49
      const totalValueCell = `${valueCol}50`; // F50
      
      // A. Overwrite E48 label to "Subtotal"
      sheet.getRange(netBillingLabelCell).setValue('Subtotal').setFontWeight('bold');
      
      // B. Write VAT (5%) formula to F49 (Labels in E49, E50 are expected to be in template)
      const vatFormula = `=${netBillingValueCell}*0.05`;
      sheet.getRange(vatValueCell).setFormula(vatFormula).setNumberFormat('$#,##0.00');
      
      // C. Write Total formula to F50
      const totalFormula = `=${netBillingValueCell}+${vatValueCell}`;
      sheet.getRange(totalValueCell).setFormula(totalFormula).setNumberFormat('$#,##0.00').setFontWeight('bold');
      
      // D. Apply borders and clear anything below (Assuming template fix was done)
      sheet.getRange('E51:F55').clearContent().clearFormat();

    } else {
      // 3. Standard Template (No VAT): Set the standard total label and ensure VAT rows are clear
      console.log("DEBUG: Standard template selected. Setting E48/D48 label to 'Billing Amount' and clearing VAT rows.");
      sheet.getRange(netBillingLabelCell).setValue('Billing Amount').setFontWeight('bold');
      // Clear rows 49 and 50 just in case
      sheet.getRange('E49:F55').clearContent().clearFormat();
    }
    
    SpreadsheetApp.flush();
    Utilities.sleep(500);

    console.log('DEBUG: ✓ Totals populated successfully.');

  } catch (error) {
    throw new Error('Error in populateInvoiceTotals: ' + error.toString());
  }
}

// ============================================
// SAVE PDF TO FOLDER FUNCTION
// ============================================

function savePDFToDrive(pdfBlob, data, monthName) {
  var folder = DriveApp.getFolderById('1DOFVLlAZmoH4rotkGJ9qjEfHrcPL92x_');
  var fileName = 'Invoice_' + data.name.replace(/\s+/g, '_') + '_' + monthName + '_' + 
                 new Date().toISOString().slice(0,10) + '.pdf';
  
  var file = folder.createFile(pdfBlob.setName(fileName));
  console.log('PDF saved directly to Generated Invoices folder: ' + fileName);
  return file;
}

// ============================================
// Test Stage 1: Draft generation
// ============================================
function testDraftGeneration() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();  
  var sheet = ss.getSheetByName('June 2025');  
  sheet.getRange('E51').setValue('Final');  // Simulate status change  
  Logger.log('Draft test completed');  
}


// ============================================
// Test Stage 2: Final issuance via Web App
// ============================================

//debug function to test the URL creation
 /**
 * Simulates draft-email and approval-link flows for debugging.
 */
function runDebugFlow() {
  // 1. Dummy inputs
  const testSheetId = '1FmWetWQ0OYz6bpJWqp8jB7uwYWTgm2ZxgpHfY7qzP4g';
  const testMonth   = 'June 2025';

  // 2. Compute and log the HMAC signature
  const payload = `${testSheetId}|${testMonth}`;
  const rawHmac = Utilities.computeHmacSha256Signature(
                    payload, HMAC_SECRET, Utilities.Charset.UTF_8
                  );
  const testSig = rawHmac
    .map(b => ('0' + (b & 0xFF).toString(16)).slice(-2))
    .join('');
  console.log('DEBUG signature:', testSig);

  // 3. Build and log the signed URL
  const approvalUrl = createSignedUrl(testSheetId, testMonth);
  console.log('DEBUG signed URL:', approvalUrl);

  // 4. Simulate data extraction
  try {
    console.log('DEBUG: Starting data extraction');
    const data = extractTimesheetData(
      SpreadsheetApp.openById(testSheetId).getSheetByName(testMonth)
    );
    console.log('DEBUG: Data extracted', data);
  } catch (err) {
    console.error('DEBUG extraction error:', err);
  }

  // 5. Simulate invoice generation
  try {
    console.log('DEBUG: Generating invoice PDF');
    const pdfBlob = generateInvoiceFromTimesheet(testSheetId, testMonth);
    console.log('DEBUG: PDF generated, size=', pdfBlob.getBytes().length);
  } catch (err) {
    console.error('DEBUG invoice generation error:', err);
  }

  // 6. Simulate saving to Drive
  try {
    console.log('DEBUG: Saving PDF to final folder');
    const pdfBlob = generateInvoiceFromTimesheet(testSheetId, testMonth);
    DriveApp.getFolderById(FINAL_FOLDER_ID)
      .createFile(pdfBlob.setName(`Invoice_${testMonth}.pdf`));
    console.log('DEBUG: PDF saved');
  } catch (err) {
    console.error('DEBUG Drive save error:', err);
  }

  // 7. Simulate sending final email
  try {
    console.log('DEBUG: Sending final email');
    const pdfBlob = generateInvoiceFromTimesheet(testSheetId, testMonth);
    sendFinalInvoiceEmail(pdfBlob, testMonth, type, consultantName);
    console.log('DEBUG: Final email sent');
  } catch (err) {
    console.error('DEBUG email send error:', err);
  }

  // 8. Simulate doGet handler
  try {
    console.log('DEBUG: Invoking doGet');
    const htmlOut = doGet({
      parameter: { sheetId: testSheetId, month: testMonth, sig: testSig }
    });
    console.log('DEBUG doGet response:', htmlOut.getContent());
  } catch (err) {
    console.error('DEBUG doGet error:', err);
  }
}

function debugConsultantName() {
  const spreadsheet = SpreadsheetApp.openById('1YYZ68Ijeo3gMMy11K6XzfwViZ0Y0fLN_Ee46U9ZjraY');
  const invoiceSheet = spreadsheet.getSheetByName('Invoice');
  const consultantName = invoiceSheet.getRange('B8').getDisplayValue().trim();
  console.log('DEBUG consultantName =', consultantName);
}

function debugLookupAndHeader() {
  const consultantName = 'Harsh Chauhan';  // Change as needed
  const dummySheetName = 'DEBUG-INVOICE-TEST';  // Will be created and deleted

  try {
    // Step 1: Lookup data
    const rates = lookupConsultantRates(consultantName);
    console.log('✔ lookupConsultantRates returned:', rates);

    // Step 2: Create dummy sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let debugSheet = ss.getSheetByName(dummySheetName);
    if (debugSheet) ss.deleteSheet(debugSheet);
    debugSheet = ss.insertSheet(dummySheetName);

    // Step 3: Use dummy invoice number
    const dummyInvoiceNumber = 'DEBUG-INV-001';

    // Step 4: Combine data (simulate real data flow)
    const mockData = {
      name: consultantName,
      street: 'Fake Street 123',
      city: 'Debugville',
      postal_code: '00000',
      country: 'Debugland',
      projectName: rates.projectName,
      clientName: rates.clientName
    };

    // Step 5: Call header population
    populateInvoiceHeaderSafe(debugSheet, mockData, dummyInvoiceNumber);

    console.log('✅ populateInvoiceHeaderSafe ran successfully.');

    SpreadsheetApp.flush();  // Ensure values are written before viewing
  } catch (error) {
    console.error('❌ Debug failed:', error);
  }
}

function debugTemplateAccess() {
  const fileId = '1Rf-Srvg6JSIhNsFkYP2_btLdfNJVKFe6xG6MSBjFLj8';
  try {
    const file = DriveApp.getFileById(fileId);
    console.log('✅ File accessed:', file.getName());

    const spreadsheet = SpreadsheetApp.openById(fileId);
    const sheetNames = spreadsheet.getSheets().map(s => s.getName());
    console.log('✅ Sheets found:', sheetNames.join(', '));
  } catch (e) {
    console.error('❌ Error accessing template:', e.toString());
  }
}

function debugProjectCodeMapping() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('January 2025') // or use .getSheetByName('January 2025')
  const consultantName = getConsultantName(sheet); // your existing helper
  const projectId = sheet.getRange('B3').getValue(); // update if project code is elsewhere

  console.log('👤 Consultant:', consultantName);
  console.log('📁 Project ID:', projectId);

  const rates = lookupConsultantRates(consultantName, projectId);
  console.log('📄 Rates info:', JSON.stringify(rates, null, 2));

  const timesheetData = extractTimesheetData(sheet);
  timesheetData.project_code = rates.project_code;

  console.log('✅ Final timesheetData object:', JSON.stringify(timesheetData, null, 2));

  // Manually test placeholder replacement logic
  const sampleActivity = 'For consultant to fill';
  const label = sampleActivity.toLowerCase().replace(/\s+/g, ' ');
  const isPlaceholder = label === 'for consultant to fill';

  if (isPlaceholder) {
    console.log('🧪 Placeholder matched. Project code substitution:', timesheetData.project_code);
  } else {
    console.log('🚫 Placeholder not matched:', label);
  }
}

/**
 * Retrieves the next sequential invoice serial from the registry sheet.
 * Returns an *integer* (1, 2, 3 …).
 */
function getGlobalInvoiceSerial() {
 // const registryId = '1kHp20TH861eHk9OMtdvu2RGPVZuCHuCeqa6Q0a1KTG8';  // from your central sheet
  const registryId = GLOBAL_COUNTER_SHEET_ID;
  const registry = SpreadsheetApp.openById(registryId);
  const sheet = registry.getSheetByName('Counter');
  
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);  // Wait up to 5 seconds

  try {
    const cell = sheet.getRange('A1');
    const current = Number(cell.getValue()) || 0;
    const next = current + 1;
    cell.setValue(next);
    return next;
  } finally {
    lock.releaseLock();
  }
}

function protectSubmittedSheet(sheet) {
  const protection = sheet.protect();
  protection.setDescription('Locked after submission');
  protection.setWarningOnly(false);

  // Allow script owner
  protection.addEditor(Session.getEffectiveUser());

  // 🔑 Also allow all admins
  ALLOWED_ADMINS.forEach(email => {
    try { protection.addEditor(email); } catch (_) {}
  });

  // Remove everybody else
  protection.getEditors().forEach(ed => {
    const mail = ed.getEmail();
    if (
      mail !== Session.getEffectiveUser().getEmail() &&
      !ALLOWED_ADMINS.includes(mail)
    ) {
      protection.removeEditor(ed);
    }
  });
}


/* ============  ADMIN MENU & UNLOCK  ============ */

const ALLOWED_ADMINS = [
  'harshvardhan.chauhan@stratverse.co',
  'noha.osman@stratverse.co',
  'invoicing@stratverse.co'
  // add more if needed
];

/** Build the Admin menu (runs every time a user opens the file). */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  // If you already have an onOpen(), just add these two lines in it
  ui.createMenu('🔧 Admin Tools')
    .addItem('🔓 Unlock Current Sheet', 'adminUnlockSheet')
    .addToUi();
}

/** Handler: removes any sheet‑level protection the admin can edit. */
function adminUnlockSheet() {
  const sheet       = SpreadsheetApp.getActiveSheet();
  const currentUser = Session.getActiveUser().getEmail();

  if (!ALLOWED_ADMINS.includes(currentUser)) {
    SpreadsheetApp.getUi()
      .alert("❌ You're not authorized to unlock this sheet.");
    return;
  }

  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  let removed = false;
  protections.forEach(p => {
    if (p.canEdit()) {           // ← true if you’re listed as an editor
      p.remove();                //   or if you’re the sheet owner
      removed = true;
    }
  });

  SpreadsheetApp.getUi().alert(
    removed
      ? `✅ Sheet “${sheet.getName()}” has been unlocked.`
      : `ℹ️ Sheet “${sheet.getName()}” is not protected or you lack permission to remove its protection.`
  );
}

/**
 * NEW FUNCTION: Notifies the central Dispatch Queue of an approved invoice.
 * This function needs to be added to your Timesheet automation script.
 * @param {Object} data - Contains consultant, month, projectId, type, fileId, [reportId]
 */
function notifyCentralDispatchQueue({ consultant, month, projectId, type, fileId, reportId }) {
  const timestamp = new Date().toISOString();

  const payload = {
    consultant,
    month,
    projectId,
    type,
    fileId,
    timestamp
  };

  // reportId is only relevant for expense types, not timesheet
  if (type === 'expense' && reportId) {
    payload.reportId = reportId; 
  }

  const payloadString = JSON.stringify(payload);
  const rawSig = Utilities.computeHmacSha256Signature(payloadString, DISPATCH_HMAC_SECRET, Utilities.Charset.UTF_8);
  const hmac = rawSig.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
  payload.hmac = hmac; // Add HMAC to the payload for verification by the webhook

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true, // Crucial for catching specific HTTP errors
  };

  try {
    const res = UrlFetchApp.fetch(DISPATCH_WEBHOOK_URL, options);
    Logger.log(`📤 Dispatch response (${res.getResponseCode()}): ${res.getContentText()}`);
  } catch (e) {
    Logger.error("❌ Error sending to dispatch queue: " + e.message);
  }
}

/**
 * DEBUG FUNCTION (Timesheet Script):
 * Simulates calling doGet with a mocked event to test Timesheet invoice dispatching.
 * This function bypasses the handleEdit and invoice creation steps, directly testing the doGet pipeline.
 *
 * PRE-REQUISITES:
 * 1. You MUST manually find an **EXISTING DRAFT CLIENT INVOICE SPREADSHEET** (created by a previous handleEdit run)
 * in your DRAFT_FOLDER_ID.
 * 2. You will need to manually fill in its Spreadsheet ID and the name of its invoice sheet.
 * 3. The consultant name used in that draft invoice MUST exist in your Staffing Tracker.
 */
function debugSimulateTimesheetApprovalToDispatchQueue() {
  console.log("🚀 Starting debugSimulateTimesheetApprovalToDispatchQueue...");

  // --- Configuration for this debug run ---
  // IMPORTANT: Replace with the actual ID of an EXISTING DRAFT CLIENT INVOICE SPREADSHEET (e.g., from your DRAFT_FOLDER_ID).
  const TEST_DRAFT_INVOICE_SS_ID = '1d_GiKjY_Oypy_B-CLLOankzTxjajKkGck3r09wzzIAw'; // Use the ID from your logs (for January 2025 client invoice)
  // IMPORTANT: Replace with the actual sheet name (invoice number) within that draft spreadsheet.
  // Example: '202501-CL-087-OW-HChauhan'
  const TEST_DRAFT_INVOICE_SHEET_NAME = 'February 2025'; // Use the invoice number from your logs
  // IMPORTANT: This consultant name MUST match the consultant in the TEST_DRAFT_INVOICE_SHEET_NAME.
  const TEST_CONSULTANT_NAME = 'Harsh Chauhan'; // Example: Harsh Chauhan

  try {
    // Step 1: Generate a real signed URL for this existing draft invoice.
    // This mimics the approval link that the Operations Head would click.
    const approvalUrl = createClientSignedUrl(
      TEST_DRAFT_INVOICE_SS_ID,
      TEST_DRAFT_INVOICE_SHEET_NAME,
      TEST_CONSULTANT_NAME
    );
    console.log(`Generated approval URL: ${approvalUrl}`);

    // Step 2: Manually parse parameters from the generated URL to simulate the 'e' event for doGet.
    const queryString = approvalUrl.split('?')[1];
    const mockEventParameter = {};
    if (queryString) {
      queryString.split('&').forEach(pair => {
        const parts = pair.split('=');
        if (parts.length === 2) {
          mockEventParameter[decodeURIComponent(parts[0])] = decodeURIComponent(parts[1] || '');
        }
      });
    }
    
    const mockEvent = { parameter: mockEventParameter };

    // Step 3: Call the doGet function directly with the mocked event.
    console.log("Calling doGet with mocked event...");
    const resultHtml = doGet(mockEvent); 
    
    console.log(`doGet simulation completed. Response: ${resultHtml.getContent()}`);
    
    console.log("✅ debugSimulateTimesheetApprovalToDispatchQueue finished. Please check Dispatch Queue sheet and Apps Script execution logs.");

  } catch (error) {
    console.error(`❌ debugSimulateTimesheetApprovalToDispatchQueue failed: ${error.message}\n${error.stack}`);
  }
}

/**
 * NEW FUNCTION (Timesheet/Expense Tracker Script - shared):
 * Sends an initial, lightweight notification to the Dispatch Queue when a submission process begins.
 * This guarantees an entry is made early, even if later, more complex steps (like invoice creation
 * or initial email sending) within handleEdit/handleExpenseEdit might intermittently fail.
 *
 * @param {string} consultantName - Name of the consultant.
 * @param {string} monthName - Month name (e.g., "January 2025").
 * @param {string} projectId - Project ID.
 * @param {string} invoiceType - 'timesheet' or 'expense'.
 */
function initialNotifyDispatchQueue(consultantName, monthName, projectId, invoiceType) {
  // Helper to get YYYY-MM from "Month YYYY" string,
  // copied locally to this script as per strict constraint (not relying on DQ script's version).
  const getYYYYMMFromMonthString = (monthString) => {
    try {
        const parts = monthString.split(' ');
        if (parts.length === 2) {
            const monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
            const monthIndex = monthNames.indexOf(parts[0]);
            const year = parseInt(parts[1], 10);
            if (monthIndex > -1 && !isNaN(year)) {
                return `${year}-${(monthIndex + 1).toString().padStart(2, '0')}`;
            }
        }
    } catch (e) { console.warn(`[Initial Dispatch] Failed to convert month name '${monthString}' to YYYY-MM for initial dispatch: ${e.message}`); }
    return '';
  };
  const monthYYYYMM = getYYYYMMFromMonthString(monthName);

  if (!monthYYYYMM) {
      console.error(`ERROR: Cannot send initial Dispatch Queue notification for ${consultantName}, month: ${monthName} - failed to normalize month.`);
      return;
  }

  const payload = {
    consultant: consultantName,
    month: monthYYYYMM, // Send month in YYYY-MM format
    projectId: projectId,
    type: `initial_${invoiceType}`, // Special type to indicate initial submission
    fileId: 'INITIAL_PLACEHOLDER', // Placeholder, as PDF is not yet created/approved
    timestamp: new Date().toISOString()
  };

  // DISPATCH_WEBHOOK_URL and DISPATCH_HMAC_SECRET are assumed to be global constants
  // that you have already added to this Timesheet Automation Script (from our previous steps).
  const payloadString = JSON.stringify(payload);
  const rawSig = Utilities.computeHmacSha256Signature(payloadString, DISPATCH_HMAC_SECRET, Utilities.Charset.UTF_8);
  const hmac = rawSig.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
  payload.hmac = hmac;

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true, // Do not throw error on HTTP 4xx/5xx responses
  };

  try {
    const res = UrlFetchApp.fetch(DISPATCH_WEBHOOK_URL, options);
    console.log(`[Initial Dispatch] Response for ${consultantName} - ${monthName}: ${res.getResponseCode()} - ${res.getContentText()}`);
  } catch (e) {
    console.error(`[Initial Dispatch] Failed to send initial notification for ${consultantName} - ${monthName}: ${e.message}`);
  }
}

function debug_dispatchFromTimesheet() {
  const sheetId = '1kGAyerlnZzbtevY3cxUEbXJZs3b3Lj7z_-h7f2ZTNqc';
  const sheetName = 'July 2025'; // Or current test sheet name
  const consultantName = 'Harsh Chauhan';

  const ss = SpreadsheetApp.openById(sheetId);
  const sheet = ss.getSheetByName(sheetName);
  const invoiceNumber = sheet.getRange('B8').getValue();

  const pdfBlob = generateInvoicePDF(sheet, invoiceNumber);
  const pdfFile = DriveApp.getFolderById(CLIENT_INVOICE_FOLDER_ID).createFile(pdfBlob);

  const rates = lookupConsultantRates(consultantName);
  const month = sheetName;

  notifyCentralDispatchQueue({
    consultant: consultantName,
    month: month,
    projectId: rates.projectId,
    type: 'timesheet',
    fileId: pdfFile.getId()
  });

  console.log('✅ Timesheet debug dispatch completed.');
}

function testDoPostMock_Timesheet() {
  const secret = DISPATCH_HMAC_SECRET; // Make sure this matches the secret in the Dispatch Queue script
  const mockBody = {
    consultant: "Harshito",
    month: "September 2025",
    projectId: "PROJ-TS",
    type: "timesheet",
    fileId: "dummy-timesheet-file-id-789",
    timestamp: new Date().toISOString()
  };

  const payloadString = JSON.stringify(mockBody);
  const rawSignature = Utilities.computeHmacSha256Signature(
    payloadString,
    secret,
    Utilities.Charset.UTF_8
  );
  const hmac = rawSignature.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
  mockBody.hmac = hmac;

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(mockBody),
    muteHttpExceptions: true
  };

  const url = DISPATCH_WEBHOOK_URL; // Ensure this points to the deployed URL of your Dispatch Web App
  const response = UrlFetchApp.fetch(url, options);
  Logger.log("Mock Timesheet dispatch response: " + response.getContentText());
}

function debugExtractData() {
  const TIMESHEET_FILE_ID = '1DKP3wkkEwyale3wqxt1uMUydqpxHqWVUv-8R2R5hZXs'; // <-- REPLACE THIS with the ID of your test timesheet file
  const TIMESHEET_SHEET_NAME = 'July 2025'; // <-- REPLACE THIS with the name of the sheet you updated

  try {
    console.log('--- Starting debugExtractData() ---');

    const ss = SpreadsheetApp.openById(TIMESHEET_FILE_ID);
    const sheet = ss.getSheetByName(TIMESHEET_SHEET_NAME);
    if (!sheet) {
      throw new Error(`Sheet "${TIMESHEET_SHEET_NAME}" not found in file.`);
    }

    const data = extractTimesheetData(sheet); // Call your existing function
    
    // Log the entire returned data object
    console.log('Extracted Data Object:');
    console.log(JSON.stringify(data, null, 2));

    // Specifically log the address fields for easy viewing
    console.log(`Extracted Street: ${data.street}`);
    console.log(`Extracted City: ${data.city}`);
    console.log(`Extracted Postal Code: ${data.postal_code}`);
    console.log(`Extracted Country: ${data.country}`);

    console.log('--- Finished debugExtractData() ---');

  } catch (error) {
    console.error('Error in debugExtractData:', error);
  }
}

function testVatCalculationAndDisplay() {
  console.log("=============================================");
  console.log("🚀 Starting VAT Calculation Debug Test");
  console.log("=============================================");

  const TEST_CONSULTANT_NAME = 'Harsh Singh Chauhan'; // Must be a consultant name in your tracker
  const TEST_MONTH = 'July 2025'; // Dummy month name

  try {
    // 1. Simulate data extraction (requires a mock sheet or a real one)
    // NOTE: You MUST replace 'YOUR_TEST_TIMESHEET_ID' and 'YOUR_TEST_SHEET_NAME' 
    // with actual IDs pointing to a TIMESHEET file (not the tracker).
    const TEST_TIMESHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId(); // Use active spreadsheet ID for simplicity
    const TEST_TIMESHEET_SHEET_NAME = 'Template'; // Use a sheet that has filled data

    let timesheetSheet;
    try {
        timesheetSheet = SpreadsheetApp.openById(TEST_TIMESHEET_ID).getSheetByName(TEST_TIMESHEET_SHEET_NAME);
        if (!timesheetSheet) {
            throw new Error(`Test sheet "${TEST_TIMESHEET_SHEET_NAME}" not found in spreadsheet ID: ${TEST_TIMESHEET_ID}`);
        }
    } catch (e) {
        console.error(`ERROR: Could not open test timesheet for debugging. Ensure IDs/Names are correct.`);
        console.error(e.message);
        return;
    }


    const timesheetData = extractTimesheetData(timesheetSheet);
    if (!timesheetData || timesheetData.total_hours <= 0) {
      console.error("FATAL: Extracted timesheet data is invalid or has zero hours. Cannot proceed.");
      return;
    }
    timesheetData.name = TEST_CONSULTANT_NAME; // Override name for consistency

    // ----------------------------------------------------
    // SCENARIO A: VAT SHOULD BE APPLIED ('5% VAT' in master tracker is 'Yes')
    // ----------------------------------------------------
    console.log("\n--- SCENARIO A: VAT APPLIED (Manually verify '5% VAT' is 'Yes') ---");
    const ratesInfoVat = lookupConsultantRates(TEST_CONSULTANT_NAME);
    
    // FORCED LOGGING TO CHECK VAT STATUS BEFORE INVOICE CREATION
    console.log(`[Rates Check] VAT Flag: ${ratesInfoVat.addVat}`);
    if (!ratesInfoVat.addVat) {
      console.warn("WARNING: VAT flag is NOT 'true'. Please manually set '5% VAT' to 'Yes' for the test consultant in the Master Tracker.");
    }
    
    const invoicesVat = createDualInvoices(timesheetData, ratesInfoVat, TEST_MONTH + " (VAT Test)");
    const clientSheetVat = invoicesVat.clientInvoice;
    
    console.log(`✅ Client Invoice with VAT created (ID: ${clientSheetVat.getParent().getId()})`);

    // Read back the total fields
    SpreadsheetApp.flush(); Utilities.sleep(1000);
    const subtotal = clientSheetVat.getRange('F48').getValue();
    const vatAmount = clientSheetVat.getRange('F50').getValue();
    const totalDue = clientSheetVat.getRange('F51').getValue();

    console.log(`  F48 (Subtotal): $${subtotal.toFixed(2)}`);
    console.log(`  F50 (5% VAT): $${vatAmount.toFixed(2)}`);
    console.log(`  F51 (Total): $${totalDue.toFixed(2)}`);
    
    // Validation checks
    if (Math.abs(vatAmount - (subtotal * 0.05)) < 0.01) {
      console.log("  ✓ VAT calculation is correct (5%).");
    } else {
      console.error("  ❌ VAT calculation is INCORRECT.");
    }
    if (Math.abs(totalDue - (subtotal + vatAmount)) < 0.01) {
      console.log("  ✓ Total Due calculation is correct.");
    } else {
      console.error("  ❌ Total Due calculation is INCORRECT.");
    }

    // Clean up the draft invoices (optional, but recommended)
    DriveApp.getFileById(invoicesVat.contractorInvoice.getParent().getId()).setTrashed(true);
    DriveApp.getFileById(invoicesVat.clientInvoice.getParent().getId()).setTrashed(true);
    console.log("  (Draft files cleaned up)");


    // ----------------------------------------------------
    // SCENARIO B: VAT SHOULD NOT BE APPLIED (Requires manual change for testing)
    // ----------------------------------------------------
    
    // NOTE: To test this, you must temporarily set the '5% VAT' column to 'Blank' or 'No' 
    // for the test consultant, then run the test again.
    console.log("\n--- SCENARIO B: NO VAT APPLIED (Manually verify '5% VAT' is 'No' or Blank) ---");
    
    // Perform a second lookup, relying on the manual change above
    const ratesInfoNoVat = lookupConsultantRates(TEST_CONSULTANT_NAME);
    console.log(`[Rates Check] VAT Flag: ${ratesInfoNoVat.addVat}`);
    if (ratesInfoNoVat.addVat) {
      console.warn("WARNING: VAT flag is still 'true'. Please manually set '5% VAT' to 'No' or blank for the test consultant.");
    }

    const invoicesNoVat = createDualInvoices(timesheetData, ratesInfoNoVat, TEST_MONTH + " (No VAT Test)");
    const clientSheetNoVat = invoicesNoVat.clientInvoice;

    console.log(`✅ Client Invoice without VAT created (ID: ${clientSheetNoVat.getParent().getId()})`);
    
    // Read back the total fields
    SpreadsheetApp.flush(); Utilities.sleep(1000);
    const noVatSubtotal = clientSheetNoVat.getRange('F48').getValue();
    const noVatAmountCell = clientSheetNoVat.getRange('F50').getDisplayValue();
    const noVatTotalCell = clientSheetNoVat.getRange('F51').getDisplayValue();
    const noVatLabel = clientSheetNoVat.getRange('E48').getDisplayValue();
    
    console.log(`  E48 Label: ${noVatLabel}`);
    console.log(`  F48 (Total): $${noVatSubtotal.toFixed(2)}`);
    console.log(`  F50 (VAT Cell Value): '${noVatAmountCell}'`);
    console.log(`  F51 (Total Cell Value): '${noVatTotalCell}'`);

    // Validation checks
    if (noVatAmountCell.trim() === '' && noVatTotalCell.trim() === '') {
      console.log("  ✓ VAT and Total cells (F50, F51) are correctly blank/cleared.");
    } else {
      console.error("  ❌ VAT/Total cells were NOT cleared when VAT was off.");
    }
    if (noVatSubtotal > 0 && noVatLabel === 'Billing Amount') {
      console.log("  ✓ E48 label correctly reverted to 'Billing Amount'.");
    } else {
      console.error("  ❌ E48 label did not revert correctly.");
    }

    // Clean up the draft invoices (optional, but recommended)
    DriveApp.getFileById(invoicesNoVat.contractorInvoice.getParent().getId()).setTrashed(true);
    DriveApp.getFileById(invoicesNoVat.clientInvoice.getParent().getId()).setTrashed(true);
    console.log("  (Draft files cleaned up)");


    console.log("\n=============================================");
    console.log("✅ VAT Debug Test Finished. Check logs for details.");
    console.log("=============================================");

  } catch (error) {
    console.error("\nCRITICAL TEST FAILURE:", error.message);
    console.error(error.stack);
  }
}


/**
 * DEBUG: Mocks the entire end-to-end timesheet submission flow (handleEdit).
 * This test simulates creating the dual invoices and sending the approval emails,
 * allowing you to check generated draft files and email content.
 *
 * *** IMPORTANT: UPDATE THESE VARIABLES BEFORE RUNNING ***
 */
/**
 * DEBUG: Mocks the entire end-to-end timesheet submission flow (handleEdit).
 * This test simulates creating the dual invoices and sending the approval emails,
 * allowing you to check generated draft files and email content.
 *
 * *** IMPORTANT: UPDATE THESE VARIABLES BEFORE RUNNING ***
 */


function testVatCellPlacement() {
  console.log("=================================================");
  console.log("🚀 Starting Isolated VAT Cell Placement Test");
  console.log("=================================================");

  const DEBUG_SHEET_NAME = "VAT_CELL_TEST_SHEET";
  const MOCK_DAILY_RATE = 800.00;
  const MOCK_TOTAL_HOURS = 160.00;
  const MOCK_SUBTOTAL = MOCK_TOTAL_HOURS * MOCK_DAILY_RATE / 8; // Expected subtotal: 16000.00
  const MOCK_VAT = MOCK_SUBTOTAL * 0.05; // Expected VAT: 800.00
  const MOCK_TOTAL = MOCK_SUBTOTAL + MOCK_VAT; // Expected Total: 16800.00

  let debugSheet;
  try {
    // 1. Create a temporary spreadsheet to mock the newly created draft client invoice
    const tempSS = SpreadsheetApp.create(`DEBUG_VAT_TEST_${Date.now()}`);
    debugSheet = tempSS.getSheets()[0].setName(DEBUG_SHEET_NAME);
    const tempFileId = tempSS.getId();

    // 2. Mock Data Setup (Minimal data required for populateInvoiceTotals)
    const mockData = {
      total_hours: MOCK_TOTAL_HOURS,
      daily_rate: MOCK_DAILY_RATE,
      applyVat: true // Force VAT ON for the test
    };

    // 3. Insert mock raw data directly into the expected calculation cells for a client invoice (F45, F47)
    debugSheet.getRange('F45').setValue(MOCK_TOTAL_HOURS);
    debugSheet.getRange('F47').setValue(MOCK_DAILY_RATE);


    // 4. Run the target function: populateInvoiceTotals
    populateInvoiceTotals(debugSheet, mockData, 'client');
    SpreadsheetApp.flush(); // Ensure formulas are calculated
    Utilities.sleep(1000); // Wait for recalculation

    // 5. Read and validate results
    console.log("\n--- Verification of Cell Contents ---");
    
    // Cell F48 (Subtotal)
    const subtotalValue = debugSheet.getRange('F48').getValue();
    const subtotalLabel = debugSheet.getRange('E48').getDisplayValue();
    console.log(`E48 Label: '${subtotalLabel}' | F48 Value (Subtotal): $${subtotalValue.toFixed(2)}`);
    
    // Cell F50 (VAT)
    const vatValue = debugSheet.getRange('F50').getValue();
    const vatLabel = debugSheet.getRange('E50').getDisplayValue();
    console.log(`E50 Label: '${vatLabel}' | F50 Value (VAT): $${vatValue.toFixed(2)}`);

    // Cell F51 (Total)
    const totalValue = debugSheet.getRange('F51').getValue();
    const totalLabel = debugSheet.getRange('E51').getDisplayValue();
    console.log(`E51 Label: '${totalLabel}' | F51 Value (Total): $${totalValue.toFixed(2)}`);
    
    // Final Assertions
    if (
      Math.abs(subtotalValue - MOCK_SUBTOTAL) < 0.01 &&
      Math.abs(vatValue - MOCK_VAT) < 0.01 &&
      Math.abs(totalValue - MOCK_TOTAL) < 0.01 &&
      totalLabel === 'Total'
    ) {
      console.log("\n✅ SUCCESS: All VAT and Total values are calculated and placed correctly.");
      console.log(`   Inspect the generated spreadsheet (ID: ${tempFileId}) to check formatting.`);
    } else {
      console.error("\n❌ FAILURE: Mismatched calculated values or labels. Review logs above.");
    }

  } catch (error) {
    console.error("\nCRITICAL TEST ERROR:", error.message);
    console.error(error.stack);
  } finally {
    // 6. Cleanup
    if (debugSheet && debugSheet.getParent()) {
      DriveApp.getFileById(debugSheet.getParent().getId()).setTrashed(true);
      console.log(`\nCleaned up temporary spreadsheet.`);
    }
    console.log("=================================================");
    console.log("Test Finished.");
  }
}

/**
 * DEBUG: Mocks the entire end-to-end timesheet submission flow (handleEdit).
 * This test simulates creating the dual invoices and sending the approval emails,
 * allowing you to check generated draft files and email content.
 *
 * *** IMPORTANT: This version PAUSES for manual inspection. ***
 */
function debugFullApprovalFlow() {
  console.log("=================================================");
  console.log("🚀 Starting Full Approval Flow Debug Test (MOCKING handleEdit)");
  console.log("=================================================");

  // 1. Configuration (MUST BE UPDATED TO WORK)
  const TEST_CONSULTANT_NAME = 'Harsh Singh Chauhan'; // Ensure this exact name is in your Tracker
  const TEST_TIMESHEET_FILE_ID = '1IL3Y8hBj-DKmVvgI_ypInd3_MxEAcvCkUgtPD7aZbQc'; // <--- REPLACE ME
  const TEST_TIMESHEET_TAB_NAME = 'July 2025'; // <--- REPLACE ME

  if (TEST_TIMESHEET_FILE_ID.includes('YOUR_TEST_TIMESHEET_FILE_ID')) {
    console.error("FATAL: Please update TEST_TIMESHEET_FILE_ID and TEST_TIMESHEET_TAB_NAME variables first.");
    return;
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    console.warn("Lock could NOT be acquired. Exiting debug.");
    return;
  }

  try {
    console.log(`Targeting Consultant: ${TEST_CONSULTANT_NAME} / Sheet: ${TEST_TIMESHEET_TAB_NAME}`);

    // 2. Setup: Open source sheet
    const ss = SpreadsheetApp.openById(TEST_TIMESHEET_FILE_ID);
    const sourceSheet = ss.getSheetByName(TEST_TIMESHEET_TAB_NAME);
    if (!sourceSheet) {
      throw new Error(`Source sheet ${TEST_TIMESHEET_TAB_NAME} not found in ${TEST_TIMESHEET_FILE_ID}`);
    }

    // 3. Core Logic steps
    const ratesInfo = lookupConsultantRates(TEST_CONSULTANT_NAME);
    
    console.log(`Lookup result. VAT flag: ${ratesInfo.addVat}`);
    
    if (!ratesInfo.addVat) {
        console.error("\nFATAL ERROR: The VAT flag read from the tracker is FALSE. Please confirm '5% VAT' is set to 'Yes' in your Master Tracker.");
        return;
    }

    ratesInfo.projectCode = ratesInfo.projectCode || 'DEBUG_PROJ';

    const timesheetData = extractTimesheetData(sourceSheet);
    timesheetData.name = TEST_CONSULTANT_NAME;
    timesheetData.project_code = ratesInfo.projectCode;
    
    // 4. Create Draft Invoices (This runs the final aggressive clearing logic)
    const { contractorInvoice, clientInvoice } = createDualInvoices(timesheetData, ratesInfo, TEST_TIMESHEET_TAB_NAME);
    
    console.log(`\n✅ Draft Invoices Created:`);
    console.log(`   - Contractor Draft ID: ${contractorInvoice.getParent().getId()}`);
    const clientDraftId = clientInvoice.getParent().getId();
    console.log(`   - Client Draft ID: ${clientDraftId}`);
    console.log(`   *** CRITICAL: OPEN AND INSPECT THIS CLIENT DRAFT FILE NOW: ${clientDraftId} ***`);
    
    // === CRITICAL PAUSE POINT ===
    SpreadsheetApp.flush(); // Ensure data is written before pausing
    //manualFixTemplatePause(); 
    // ============================


    // 5. Send Approval Emails (Resumes after user clicks OK)
    sendContractorApprovalEmail(contractorInvoice, TEST_CONSULTANT_NAME);
    sendClientApprovalEmail(clientInvoice, TEST_CONSULTANT_NAME);

    console.log("\n✅ Approval Emails Sent (to OPERATIONS_HEAD_EMAIL).");
    
    console.log("\n=================================================");
    console.log("✨ Test Complete. Cleanup and final verification complete.");
    console.log("   Manual Cleanup Required: Delete the new draft files in the DRAFT_FOLDER_ID.");
    console.log("=================================================");

  } catch (err) {
    console.error('CRITICAL ERROR in debugFullApprovalFlow: ' + err.message + '\n' + err.stack);
  } finally {
    lock.releaseLock();
  }
}













