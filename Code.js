const SCHOLAR_HEADERS = [
  'Title / Citation',
  'Matched Title',
  'Abstract',
  'Status',
  'Document Link'
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Scholar Tools')
    .addItem('Setup Sheet', 'setupScholarSheet')
    .addItem('Find Abstract', 'findSelectedAbstract')
    .addItem('Find All Abstracts', 'findAllAbstracts')
    .addSeparator()
    .addItem('Generate Document', 'generateSelectedDocument')
    .addItem('Generate Combined Report', 'generateCombinedReport')
    .addToUi();
}

function setupScholarSheet() {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange(1, 1, 1, SCHOLAR_HEADERS.length).setValues([SCHOLAR_HEADERS]);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, SCHOLAR_HEADERS.length).setFontWeight('bold');
  sheet.autoResizeColumns(1, SCHOLAR_HEADERS.length);
}

function findSelectedAbstract() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const row = sheet.getActiveRange().getRow();

  if (row < 2) {
    throw new Error('Select any cell on a row below the header first.');
  }

  const rawInput = String(sheet.getRange(row, 1).getValue() || '').trim();
  if (!rawInput) {
    throw new Error('Column A does not contain a title or citation.');
  }

  sheet.getRange(row, 4).setValue('Searching...');
  SpreadsheetApp.flush();

  try {
    const result = fetchAbstractForCitation_(rawInput);
    writeAbstractResult_(sheet, row, result);
  } catch (err) {
    const status = getErrorStatus_(err);
    sheet.getRange(row, 4).setValue(status);
    throw err;
  }
}

function findAllAbstracts() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    throw new Error('No citation rows found.');
  }

  const rows = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  let found = 0;
  let skipped = 0;
  let noAbstract = 0;
  let failed = 0;

  for (let i = 0; i < rows.length; i++) {
    const sheetRow = i + 2;
    const citation = String(rows[i][0] || '').trim();
    const existingAbstract = String(rows[i][2] || '').trim();

    if (!citation) continue;

    if (existingAbstract) {
      skipped++;
      continue;
    }

    sheet.getRange(sheetRow, 4).setValue('Searching...');
    SpreadsheetApp.flush();

    try {
      const result = fetchAbstractForCitation_(citation);
      writeAbstractResult_(sheet, sheetRow, result);

      if (result && result.abstract) found++;
      else noAbstract++;
    } catch (err) {
      failed++;
      sheet.getRange(sheetRow, 4).setValue(getErrorStatus_(err));
    }

    // Gentle pacing between papers so we do not hammer the public API.
    Utilities.sleep(1200);
  }

  SpreadsheetApp.getUi().alert(
    'Batch search finished.\n' +
    'Abstracts found: ' + found + '\n' +
    'Already filled / skipped: ' + skipped + '\n' +
    'Matched but no abstract: ' + noAbstract + '\n' +
    'Failed: ' + failed
  );
}

function fetchAbstractForCitation_(rawInput) {
  const parsed = parseCitationInput_(rawInput);
  const searchTitle = parsed.title || rawInput;
  return searchSemanticScholarMatch_(searchTitle);
}

function writeAbstractResult_(sheet, row, result) {
  if (!result) {
    sheet.getRange(row, 2, 1, 3).setValues([['', '', 'No match found']]);
    return;
  }

  sheet.getRange(row, 2, 1, 3).setValues([[
    result.title || '',
    result.abstract || '',
    result.abstract ? 'Abstract found' : 'Matched - no abstract available'
  ]]);
}

function getErrorStatus_(err) {
  const message = String(err && err.message ? err.message : err);
  if (message.includes('HTTP 429')) return 'Rate limited - retry later';
  if (message.includes('HTTP 404')) return 'No match found';
  return 'Error';
}

function generateSelectedDocument() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const row = sheet.getActiveRange().getRow();

  if (row < 2) {
    throw new Error('Select the research row you want to generate.');
  }

  const citation = String(sheet.getRange(row, 1).getValue() || '').trim();
  const abstract = String(sheet.getRange(row, 3).getValue() || '').trim();

  if (!citation) {
    throw new Error('Column A does not contain a citation.');
  }

  if (!abstract) {
    throw new Error('No abstract found in Column C. Run "Find Abstract" first.');
  }

  const parsed = parseCitationInput_(citation);
  const docTitle = parsed.title
    ? 'Research - ' + parsed.title
    : 'Research Abstract';

  const doc = DocumentApp.create(docTitle);
  const body = doc.getBody();

  appendCitationAndAbstract_(body, citation, abstract);

  doc.saveAndClose();

  const url = doc.getUrl();
  sheet.getRange(row, 5).setFormula('=HYPERLINK("' + url + '","Open Document")');
  sheet.getRange(row, 4).setValue('Document created');

  SpreadsheetApp.getUi().alert('Document created successfully.');
}

function generateCombinedReport() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    throw new Error('No research rows found.');
  }

  const rows = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  const validRows = rows.filter(row => {
    const citation = String(row[0] || '').trim();
    const abstract = String(row[2] || '').trim();
    return citation && abstract;
  });

  if (!validRows.length) {
    throw new Error('No rows contain both a citation and an abstract.');
  }

  const timestamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd HHmm'
  );

  const doc = DocumentApp.create('Scholar Combined Report - ' + timestamp);
  const body = doc.getBody();

  validRows.forEach((row, index) => {
    const citation = String(row[0] || '').trim();
    const abstract = String(row[2] || '').trim();

    appendCitationAndAbstract_(body, citation, abstract);

    if (index < validRows.length - 1) {
      body.appendParagraph('');
      body.appendHorizontalRule();
      body.appendParagraph('');
    }
  });

  doc.saveAndClose();

  const url = doc.getUrl();
  sheet.getRange('E1').setFormula('=HYPERLINK("' + url + '","Open Combined Report")');

  SpreadsheetApp.getUi().alert(
    'Combined report created with ' + validRows.length + ' research entr' +
    (validRows.length === 1 ? 'y.' : 'ies.')
  );
}

function appendCitationAndAbstract_(body, citation, abstract) {
  const citationParagraph = body.appendParagraph(citation);
  citationParagraph.setSpacingAfter(12);

  const abstractParagraph = body.appendParagraph(abstract);
  abstractParagraph.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
}

function searchSemanticScholarMatch_(title) {
  const fields = 'title,abstract';
  const url = 'https://api.semanticscholar.org/graph/v1/paper/search/match' +
    '?query=' + encodeURIComponent(title) +
    '&fields=' + encodeURIComponent(fields);

  const retryDelays = [0, 2000, 5000, 10000];
  let lastResponseText = '';
  let lastCode = 0;

  for (let attempt = 0; attempt < retryDelays.length; attempt++) {
    if (retryDelays[attempt] > 0) {
      Utilities.sleep(retryDelays[attempt]);
    }

    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();
    const text = response.getContentText();
    lastCode = code;
    lastResponseText = text;

    if (code === 200) {
      const payload = JSON.parse(text);
      return payload && payload.data ? payload.data[0] || null : null;
    }

    // Retry only rate-limit and temporary server errors.
    if (code !== 429 && code < 500) {
      throw new Error('Semantic Scholar returned HTTP ' + code + ': ' + text);
    }
  }

  throw new Error('Semantic Scholar returned HTTP ' + lastCode + ': ' + lastResponseText);
}

function parseCitationInput_(input) {
  const text = String(input || '').replace(/\s+/g, ' ').trim();
  if (!text) return { title: '' };

  const yearMatch = text.match(/\((\d{4}[a-z]?)\)\s*\.??\s*/i);
  if (!yearMatch) return { title: text };

  let remainder = text.slice(yearMatch.index + yearMatch[0].length).trim();
  remainder = remainder.replace(/^[.\s]+/, '');

  const boundary = findLikelyTitleBoundary_(remainder);
  const title = boundary >= 0
    ? remainder.slice(0, boundary).trim().replace(/[.\s]+$/, '')
    : remainder.replace(/[.\s]+$/, '').trim();

  return { title: title || text };
}

function findLikelyTitleBoundary_(text) {
  if (!text) return -1;

  const sourceHints = [
    'proceedings of',
    'journal of',
    'international journal',
    'conference on',
    'conference proceedings',
    'transactions on',
    'communications of',
    'lecture notes in',
    'arxiv',
    'springer',
    'elsevier',
    'ieee',
    'acm'
  ];

  const lower = text.toLowerCase();
  for (let i = 0; i < text.length - 2; i++) {
    if (text[i] !== '.') continue;
    const after = lower.slice(i + 1).trimStart();
    if (sourceHints.some(hint => after.startsWith(hint))) return i;
  }

  const match = text.match(/^(.{20,}?)\.\s+(?=[A-Z])/);
  return match ? match[1].length : -1;
}
