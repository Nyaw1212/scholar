const SCHOLAR_HEADERS = [
  'Paper Title',
  'Matched Title',
  'Authors',
  'Year',
  'Journal / Venue',
  'DOI',
  'Abstract',
  'Match Score',
  'Status',
  'Document Link'
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Scholar Tools')
    .addItem('Setup Sheet', 'setupScholarSheet')
    .addItem('Find Selected Paper', 'findSelectedPaper')
    .addItem('Generate Document', 'generateSelectedPaperDocument')
    .addToUi();
}

function setupScholarSheet() {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange(1, 1, 1, SCHOLAR_HEADERS.length).setValues([SCHOLAR_HEADERS]);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, SCHOLAR_HEADERS.length).setFontWeight('bold');
  sheet.autoResizeColumns(1, SCHOLAR_HEADERS.length);
}

function findSelectedPaper() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const row = sheet.getActiveRange().getRow();

  if (row < 2) {
    throw new Error('Select a row below the header first.');
  }

  const queryTitle = String(sheet.getRange(row, 1).getValue() || '').trim();
  if (!queryTitle) {
    throw new Error('Column A does not contain a paper title.');
  }

  sheet.getRange(row, 9).setValue('Searching...');

  const result = searchSemanticScholar_(queryTitle);

  if (!result) {
    sheet.getRange(row, 9).setValue('No match found');
    return;
  }

  const authors = (result.authors || []).map(a => a.name).join(', ');
  const doi = result.externalIds && result.externalIds.DOI
    ? result.externalIds.DOI
    : '';
  const score = titleSimilarity_(queryTitle, result.title || '');

  sheet.getRange(row, 2, 1, 8).setValues([[
    result.title || '',
    authors,
    result.year || '',
    result.venue || '',
    doi,
    result.abstract || '',
    Math.round(score * 100) + '%',
    score >= 0.90 ? 'Matched' : 'Check match'
  ]]);
}

function generateSelectedPaperDocument() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const row = sheet.getActiveRange().getRow();

  if (row < 2) {
    throw new Error('Select a paper row first.');
  }

  const values = sheet.getRange(row, 1, 1, 10).getValues()[0];
  const [queryTitle, matchedTitle, authors, year, venue, doi, abstract, matchScore, status] = values;

  if (!matchedTitle) {
    throw new Error('Run "Find Selected Paper" first.');
  }

  const doc = DocumentApp.create('Research - ' + matchedTitle);
  const body = doc.getBody();

  body.appendParagraph(matchedTitle)
    .setHeading(DocumentApp.ParagraphHeading.TITLE)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph('Authors').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(authors || 'Not available');

  body.appendParagraph('Year').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(String(year || 'Not available'));

  body.appendParagraph('Journal / Venue').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(venue || 'Not available');

  body.appendParagraph('DOI').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(doi || 'Not available');

  body.appendParagraph('Abstract').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(abstract || 'Abstract not available.')
    .setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);

  body.appendParagraph('Source Title Used').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(queryTitle || matchedTitle);

  body.appendParagraph('Match Score').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(String(matchScore || ''));

  body.appendParagraph('Basic APA-style Citation').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(buildBasicCitation_(matchedTitle, authors, year, venue, doi));

  doc.saveAndClose();

  const url = doc.getUrl();
  sheet.getRange(row, 10).setFormula('=HYPERLINK("' + url + '","Open Document")');
  sheet.getRange(row, 9).setValue(status === 'Check match' ? 'Document created - CHECK MATCH' : 'Document created');
}

function searchSemanticScholar_(title) {
  const fields = [
    'title',
    'authors',
    'year',
    'abstract',
    'venue',
    'externalIds',
    'url'
  ].join(',');

  const url = 'https://api.semanticscholar.org/graph/v1/paper/search' +
    '?query=' + encodeURIComponent(title) +
    '&limit=5' +
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
  const papers = payload.data || [];
  if (!papers.length) return null;

  papers.forEach(p => {
    p._score = titleSimilarity_(title, p.title || '');
  });

  papers.sort((a, b) => b._score - a._score);
  return papers[0];
}

function titleSimilarity_(a, b) {
  const aTokens = normalizeTitle_(a).split(' ').filter(Boolean);
  const bTokens = normalizeTitle_(b).split(' ').filter(Boolean);

  if (!aTokens.length || !bTokens.length) return 0;

  const aSet = new Set(aTokens);
  const bSet = new Set(bTokens);
  let intersection = 0;

  aSet.forEach(token => {
    if (bSet.has(token)) intersection++;
  });

  return (2 * intersection) / (aSet.size + bSet.size);
}

function normalizeTitle_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function buildBasicCitation_(title, authors, year, venue, doi) {
  let citation = '';
  citation += (authors || 'Unknown author') + '. ';
  citation += '(' + (year || 'n.d.') + '). ';
  citation += (title || 'Untitled') + '. ';
  if (venue) citation += venue + '. ';
  if (doi) citation += 'https://doi.org/' + doi;
  return citation.trim();
}
