const SCHOLAR_HEADERS = [
  'Title / Citation',
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
    throw new Error('Select any cell on a row below the header first.');
  }

  const rawInput = String(sheet.getRange(row, 1).getValue() || '').trim();
  if (!rawInput) {
    throw new Error('Column A does not contain a title or citation.');
  }

  const parsed = parseCitationInput_(rawInput);
  const searchTitle = parsed.title || rawInput;

  sheet.getRange(row, 9).setValue('Searching...');

  const result = searchSemanticScholar_(searchTitle);

  if (!result) {
    sheet.getRange(row, 9).setValue('No match found');
    return;
  }

  const authors = (result.authors || []).map(a => a.name).join(', ');
  const doi = result.externalIds && result.externalIds.DOI
    ? result.externalIds.DOI
    : '';
  const score = titleSimilarity_(searchTitle, result.title || '');

  sheet.getRange(row, 2, 1, 8).setValues([[
    result.title || '',
    authors,
    result.year || parsed.year || '',
    result.venue || parsed.venue || '',
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
  const [sourceInput, matchedTitle, authors, year, venue, doi, abstract, matchScore, status] = values;

  if (!matchedTitle) {
    throw new Error('Run "Find Selected Paper" first.');
  }

  const parsed = parseCitationInput_(sourceInput);
  const doc = DocumentApp.create('Research - ' + matchedTitle);
  const body = doc.getBody();

  body.appendParagraph(matchedTitle)
    .setHeading(DocumentApp.ParagraphHeading.TITLE)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  body.appendParagraph('Authors').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(authors || parsed.authors || 'Not available');

  body.appendParagraph('Year').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(String(year || parsed.year || 'Not available'));

  body.appendParagraph('Journal / Venue').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(venue || parsed.venue || 'Not available');

  body.appendParagraph('DOI').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(doi || 'Not available');

  body.appendParagraph('Abstract').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(abstract || 'Abstract not available.')
    .setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);

  body.appendParagraph('Original Input').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(sourceInput || matchedTitle);

  body.appendParagraph('Parsed Search Title').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(parsed.title || sourceInput || matchedTitle);

  body.appendParagraph('Match Score').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(String(matchScore || ''));

  body.appendParagraph('Basic APA-style Citation').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(buildBasicCitation_(matchedTitle, authors, year, venue, doi));

  doc.saveAndClose();

  const url = doc.getUrl();
  sheet.getRange(row, 10).setFormula('=HYPERLINK("' + url + '","Open Document")');
  sheet.getRange(row, 9).setValue(status === 'Check match' ? 'Document created - CHECK MATCH' : 'Document created');
}

/**
 * Accepts either:
 *   1) a plain paper title, or
 *   2) a citation such as:
 *      Narajala, V. S., et al. (2025). Securing Agentic AI: ... . Proceedings of ...
 *
 * Returns best-effort parsed metadata. The extracted title is used only as
 * the search query; Semantic Scholar remains the source for matched metadata.
 */
function parseCitationInput_(input) {
  const text = String(input || '').replace(/\s+/g, ' ').trim();

  const parsed = {
    raw: text,
    authors: '',
    year: '',
    title: text,
    venue: '',
    isCitation: false
  };

  if (!text) return parsed;

  // Common citation pattern: Authors (YYYY). Title. Venue...
  // Also tolerates 2025a / 2025b style years.
  const yearMatch = text.match(/\((\d{4}[a-z]?)\)\s*\.??\s*/i);
  if (!yearMatch) {
    return parsed;
  }

  parsed.isCitation = true;
  parsed.year = yearMatch[1];
  parsed.authors = text.slice(0, yearMatch.index).trim().replace(/[.,;:\s]+$/, '');

  let remainder = text.slice(yearMatch.index + yearMatch[0].length).trim();

  // Remove punctuation left immediately after the year marker.
  remainder = remainder.replace(/^[.\s]+/, '');

  // In most APA-like references, the paper title is the first sentence after
  // the year. We score sentence boundaries so abbreviations and initials are
  // less likely to terminate the title accidentally.
  const boundary = findLikelyTitleBoundary_(remainder);

  if (boundary >= 0) {
    parsed.title = remainder.slice(0, boundary).trim().replace(/[.\s]+$/, '');
    parsed.venue = remainder.slice(boundary + 1).trim();
  } else {
    parsed.title = remainder.replace(/[.\s]+$/, '').trim();
  }

  if (!parsed.title) parsed.title = text;
  return parsed;
}

function findLikelyTitleBoundary_(text) {
  if (!text) return -1;

  // Look for a period followed by whitespace and a likely venue/source phrase.
  // This works well for inputs like "Title. Proceedings of ..." and
  // "Title. Journal of ..." while preserving punctuation inside the title.
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
    if (sourceHints.some(hint => after.startsWith(hint))) {
      return i;
    }
  }

  // Fallback: first sentence-ending period followed by an uppercase letter,
  // but only after a reasonably long title segment.
  const match = text.match(/^(.{20,}?)\.\s+(?=[A-Z])/);
  return match ? match[1].length : -1;
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
