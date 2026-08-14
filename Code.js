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

  const parsed = parseCitationInput_(rawInput);
  const searchTitle = parsed.title || rawInput;

  sheet.getRange(row, 4).setValue('Searching...');

  try {
    const result = searchSemanticScholarMatch_(searchTitle);

    if (!result) {
      sheet.getRange(row, 2, 1, 3).setValues([['', '', 'No match found']]);
      return;
    }

    sheet.getRange(row, 2, 1, 3).setValues([[
      result.title || '',
      result.abstract || '',
      result.abstract ? 'Abstract found' : 'Matched - no abstract available'
    ]]);
  } catch (err) {
    const message = String(err && err.message ? err.message : err);
    const status = message.includes('HTTP 429')
      ? 'Rate limited - try again shortly'
      : 'Error';

    sheet.getRange(row, 4).setValue(status);
    throw err;
  }
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

  // Put the combined report link in E1 so it is easy to find and does not
  // overwrite any individual row document links.
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

  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  if (code !== 200) {
    throw new Error('Semantic Scholar returned HTTP ' + code + ': ' + response.getContentText());
  }

  const payload = JSON.parse(response.getContentText());
  return payload && payload.data ? payload.data[0] || null : null;
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
