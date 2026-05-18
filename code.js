/**
 * Tilbrook Consulting Customer Brief Generator
 *
 * Reads prospect rows from the active Google Sheet, generates a one-page
 * Growth Opportunity Snapshot Google Doc for rows marked "Ready", and writes
 * resulting links/status updates back to the sheet.
 */

const CONFIG = {
  DATA_START_ROW: 2,
  READY_STATUS: 'Ready',
  CREATED_STATUS: 'Brief Created',
  ERROR_STATUS: 'Error',
  MENU_NAME: 'Tilbrook Briefs',
  DOCUMENT_TITLE_PREFIX: 'Growth Opportunity Snapshot',
  EXPORT_PDF_PROPERTY: 'EXPORT_PDF',
  CREATE_DRAFT_PROPERTY: 'CREATE_GMAIL_DRAFTS',
  BRIEF_FOLDER_ID_PROPERTY: 'BRIEF_FOLDER_ID'
};

const COLUMNS = Object.freeze({
  BUSINESS_NAME: 1,
  CONTACT_NAME: 2,
  EMAIL: 3,
  INDUSTRY: 4,
  LOCATION: 5,
  WEBSITE: 6,
  BUSINESS_NOTES: 7,
  GROWTH_CHALLENGES: 8,
  OFFER_ANGLE: 9,
  STATUS: 10,
  DOC_LINK: 11,
  PDF_LINK: 12,
  GMAIL_DRAFT_STATUS: 13,
  LAST_UPDATED: 14,
  ERROR_NOTES: 15
});

/**
 * Adds a custom menu when the spreadsheet opens.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(CONFIG.MENU_NAME)
    .addItem('Process ready briefs', 'processReadyBriefs')
    .addSeparator()
    .addItem('Process ready briefs and replace existing links', 'processReadyBriefsAllowOverwrite')
    .addToUi();
}

/**
 * Processes rows where Status is exactly "Ready" without overwriting existing brief links.
 */
function processReadyBriefs() {
  return processReadyBriefRows_({ allowOverwrite: false });
}

/**
 * Processes ready rows and allows existing Google Doc/PDF links to be replaced.
 * Use this intentionally; the default workflow never overwrites existing briefs.
 */
function processReadyBriefsAllowOverwrite() {
  return processReadyBriefRows_({ allowOverwrite: true });
}

/**
 * Main row processor.
 * @param {{allowOverwrite: boolean}} options
 * @returns {{processed: number, errors: number, skipped: number}}
 */
function processReadyBriefRows_(options) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow < CONFIG.DATA_START_ROW) {
    return { processed: 0, errors: 0, skipped: 0 };
  }

  const rowCount = lastRow - CONFIG.DATA_START_ROW + 1;
  const range = sheet.getRange(CONFIG.DATA_START_ROW, 1, rowCount, COLUMNS.ERROR_NOTES);
  const values = range.getValues();
  const results = { processed: 0, errors: 0, skipped: 0 };

  values.forEach((row, index) => {
    const rowNumber = CONFIG.DATA_START_ROW + index;
    const status = normalizeValue_(row[COLUMNS.STATUS - 1]);

    if (status !== CONFIG.READY_STATUS) {
      results.skipped += 1;
      return;
    }

    try {
      const prospect = buildProspectFromRow_(row, rowNumber);
      validateProspect_(prospect);

      if (!options.allowOverwrite && prospect.googleDocLink) {
        throw new Error('Existing Google Doc Link found. Skipped to avoid overwriting an existing brief. Use the explicit replace menu item if this is intentional.');
      }

      const doc = createBriefDocument_(prospect);
      const pdfLink = shouldExportPdf_() ? exportDocumentPdf_(doc.id, prospect) : prospect.pdfLink;
      const draftStatus = shouldCreateGmailDraft_() ? createGmailDraft_(prospect, doc.url, pdfLink) : prospect.gmailDraftStatus;

      sheet.getRange(rowNumber, COLUMNS.STATUS).setValue(CONFIG.CREATED_STATUS);
      sheet.getRange(rowNumber, COLUMNS.DOC_LINK).setValue(doc.url);
      sheet.getRange(rowNumber, COLUMNS.PDF_LINK).setValue(pdfLink || '');
      sheet.getRange(rowNumber, COLUMNS.GMAIL_DRAFT_STATUS).setValue(draftStatus || 'Not requested');
      sheet.getRange(rowNumber, COLUMNS.LAST_UPDATED).setValue(new Date());
      sheet.getRange(rowNumber, COLUMNS.ERROR_NOTES).clearContent();
      results.processed += 1;
    } catch (error) {
      writeRowError_(sheet, rowNumber, error);
      results.errors += 1;
    }
  });

  SpreadsheetApp.flush();
  return results;
}

/**
 * Converts a sheet row into a named prospect object.
 * @param {Array<*>} row
 * @param {number} rowNumber
 */
function buildProspectFromRow_(row, rowNumber) {
  return {
    rowNumber,
    businessName: normalizeValue_(row[COLUMNS.BUSINESS_NAME - 1]),
    contactName: normalizeValue_(row[COLUMNS.CONTACT_NAME - 1]),
    email: normalizeValue_(row[COLUMNS.EMAIL - 1]),
    industry: normalizeValue_(row[COLUMNS.INDUSTRY - 1]),
    location: normalizeValue_(row[COLUMNS.LOCATION - 1]),
    website: normalizeValue_(row[COLUMNS.WEBSITE - 1]),
    businessNotes: normalizeValue_(row[COLUMNS.BUSINESS_NOTES - 1]),
    growthChallenges: normalizeValue_(row[COLUMNS.GROWTH_CHALLENGES - 1]),
    offerAngle: normalizeValue_(row[COLUMNS.OFFER_ANGLE - 1]),
    googleDocLink: normalizeValue_(row[COLUMNS.DOC_LINK - 1]),
    pdfLink: normalizeValue_(row[COLUMNS.PDF_LINK - 1]),
    gmailDraftStatus: normalizeValue_(row[COLUMNS.GMAIL_DRAFT_STATUS - 1])
  };
}

/**
 * Ensures the minimum fields needed for a useful brief are present.
 * @param {Object} prospect
 */
function validateProspect_(prospect) {
  const missing = [];

  if (!prospect.businessName) missing.push('Business Name');
  if (!prospect.industry) missing.push('Industry');
  if (!prospect.location) missing.push('Location');

  if (missing.length > 0) {
    throw new Error(`Missing required field(s): ${missing.join(', ')}`);
  }
}

/**
 * Creates and formats the Google Doc brief.
 * @param {Object} prospect
 * @returns {{id: string, url: string}}
 */
function createBriefDocument_(prospect) {
  const title = `${CONFIG.DOCUMENT_TITLE_PREFIX} - ${prospect.businessName}`;
  const doc = DocumentApp.create(title);
  const body = doc.getBody();

  body.clear();
  appendTitle_(body, 'Growth Opportunity Snapshot');
  appendSubtitle_(body, prospect.businessName);
  appendMetadata_(body, prospect);

  appendSection_(body, 'Current Context', buildCurrentContext_(prospect));
  appendSection_(body, 'Growth Opportunities', buildGrowthOpportunities_(prospect));
  appendSection_(body, 'Practical First Steps', buildPracticalFirstSteps_(prospect));
  appendSection_(body, 'Conversation Starter', buildConversationStarter_(prospect));
  appendFooter_(body);

  doc.saveAndClose();
  moveFileToConfiguredFolder_(doc.getId());

  return {
    id: doc.getId(),
    url: doc.getUrl()
  };
}

function appendTitle_(body, text) {
  body.appendParagraph(text)
    .setHeading(DocumentApp.ParagraphHeading.TITLE)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
}

function appendSubtitle_(body, text) {
  body.appendParagraph(text)
    .setHeading(DocumentApp.ParagraphHeading.SUBTITLE)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
}

function appendMetadata_(body, prospect) {
  const details = [
    `Industry: ${valueOrFallback_(prospect.industry)}`,
    `Location: ${valueOrFallback_(prospect.location)}`,
    `Website: ${valueOrFallback_(prospect.website)}`,
    `Prepared for: ${valueOrFallback_(prospect.contactName, 'Business owner / manager')}`
  ];

  body.appendParagraph(details.join('  |  '))
    .setFontSize(9)
    .setForegroundColor('#555555')
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendHorizontalRule();
}

function appendSection_(body, heading, paragraphs) {
  body.appendParagraph(heading).setHeading(DocumentApp.ParagraphHeading.HEADING2);
  paragraphs.forEach((paragraph) => body.appendParagraph(paragraph).setFontSize(10));
}

function appendFooter_(body) {
  body.appendHorizontalRule();
  body.appendParagraph('Prepared by Tilbrook Consulting. This snapshot is intended as a practical starting point for a respectful business development conversation.')
    .setFontSize(8)
    .setForegroundColor('#666666');
}

function buildCurrentContext_(prospect) {
  const notes = prospect.businessNotes || 'The available notes suggest there is an opportunity to clarify the business story and make it easier for potential customers to understand what makes the company useful and trustworthy.';
  const challenges = prospect.growthChallenges || 'The most useful next step is to confirm the main growth constraint through a short conversation before recommending any specific tactic.';

  return [
    `${prospect.businessName} operates in the ${prospect.industry} space in ${prospect.location}. ${notes}`,
    `A practical growth conversation should focus on the challenge area currently visible from the prospect information: ${challenges}`
  ];
}

function buildGrowthOpportunities_(prospect) {
  const offerAngle = prospect.offerAngle || 'positioning the business more clearly for local customers and online discovery';

  return [
    `Improve AI and search visibility by making the business description, services, location signals, and customer proof points easier for both people and digital assistants to interpret.`,
    `Strengthen customer connection by turning common questions, objections, and service differentiators into simple website, profile, and follow-up content.`,
    `Frame the first conversation around ${offerAngle}, while keeping the recommendation specific to what the owner or manager says matters most.`
  ];
}

function buildPracticalFirstSteps_(prospect) {
  return [
    `Review the current website and local profiles for clarity: who the business serves, what problem it solves, where it operates, and how customers should take the next step.`,
    `Identify three customer questions that could become short pieces of helpful content for the website, Google Business Profile, email follow-up, or social posts.`,
    `Choose one measurable near-term goal, such as more qualified calls, improved follow-up, better repeat customer communication, or stronger local visibility.`
  ];
}

function buildConversationStarter_(prospect) {
  const contact = prospect.contactName || 'there';
  const challengePhrase = prospect.growthChallenges ? `I noticed you may be thinking about ${prospect.growthChallenges}.` : 'I was looking at a few practical ways businesses can make their value easier for customers to find and understand.';

  return [
    `Hi ${contact}, ${challengePhrase} I put together a short Growth Opportunity Snapshot for ${prospect.businessName} with a few practical ideas around visibility, customer connection, and local growth. No hard sell — just a useful starting point if you would like to compare notes.`
  ];
}

/**
 * Exports the generated Google Doc as a PDF when enabled by script property.
 * @param {string} documentId
 * @param {Object} prospect
 * @returns {string}
 */
function exportDocumentPdf_(documentId, prospect) {
  const docFile = DriveApp.getFileById(documentId);
  const pdfBlob = docFile.getBlob().getAs(MimeType.PDF);
  const pdfName = `${CONFIG.DOCUMENT_TITLE_PREFIX} - ${prospect.businessName}.pdf`;
  const folder = getConfiguredFolder_();
  const pdfFile = folder ? folder.createFile(pdfBlob).setName(pdfName) : DriveApp.createFile(pdfBlob).setName(pdfName);

  return pdfFile.getUrl();
}

/**
 * Creates a Gmail draft when enabled by script property. It never sends email.
 * @param {Object} prospect
 * @param {string} docUrl
 * @param {string} pdfUrl
 * @returns {string}
 */
function createGmailDraft_(prospect, docUrl, pdfUrl) {
  if (!prospect.email) {
    return 'Skipped - missing email';
  }

  const subject = `Growth Opportunity Snapshot for ${prospect.businessName}`;
  const greeting = prospect.contactName ? `Hi ${prospect.contactName},` : 'Hi,';
  const pdfLine = pdfUrl ? `\nPDF copy: ${pdfUrl}` : '';
  const body = `${greeting}\n\nI put together a short Growth Opportunity Snapshot for ${prospect.businessName}. It includes a few practical ideas around visibility, customer connection, and local growth.\n\nGoogle Doc: ${docUrl}${pdfLine}\n\nBest,\nTilbrook Consulting`;

  GmailApp.createDraft(prospect.email, subject, body);
  return 'Draft created';
}

function writeRowError_(sheet, rowNumber, error) {
  sheet.getRange(rowNumber, COLUMNS.STATUS).setValue(CONFIG.ERROR_STATUS);
  sheet.getRange(rowNumber, COLUMNS.ERROR_NOTES).setValue(error && error.message ? error.message : String(error));
  sheet.getRange(rowNumber, COLUMNS.LAST_UPDATED).setValue(new Date());
}

function shouldExportPdf_() {
  return getBooleanScriptProperty_(CONFIG.EXPORT_PDF_PROPERTY);
}

function shouldCreateGmailDraft_() {
  return getBooleanScriptProperty_(CONFIG.CREATE_DRAFT_PROPERTY);
}

function getBooleanScriptProperty_(key) {
  return String(PropertiesService.getScriptProperties().getProperty(key)).toLowerCase() === 'true';
}

function moveFileToConfiguredFolder_(fileId) {
  const folder = getConfiguredFolder_();

  if (!folder) {
    return;
  }

  const file = DriveApp.getFileById(fileId);
  file.moveTo(folder);
}

function getConfiguredFolder_() {
  const folderId = normalizeValue_(PropertiesService.getScriptProperties().getProperty(CONFIG.BRIEF_FOLDER_ID_PROPERTY));
  return folderId ? DriveApp.getFolderById(folderId) : null;
}

function normalizeValue_(value) {
  return value === null || value === undefined ? '' : String(value).trim();
}

function valueOrFallback_(value, fallback) {
  return value || fallback || 'Not provided';
}
