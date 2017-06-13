/**
 * @OnlyCurrentDoc
 *
 * The above comment directs Apps Script to limit the scope of file
 * access for this add-on. It specifies that this add-on will only
 * attempt to read or modify the files in which the add-on is used,
 * and not all of the user's files. The authorization request message
 * presented to users will reflect this limited scope.
 */

/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
      .addItem('Start', 'showSidebar')
      .addToUi();
}

/**
 * Runs when the add-on is installed.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE.)
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Opens a sidebar in the document containing the add-on's user interface.
 */
function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('md2gdoc');
  DocumentApp.getUi().showSidebar(ui);
}

/**
 * Generate the Google Doc based on the given markdown text
 *
 * @param {string} markdownText The GitHub-flavored markdown text
 */
function generateGdoc(markdownText) {  
  var lexer = new marked.Lexer();
  var parser = new marked.Parser();
  var tokens = lexer.lex(markdownText);
  convertMarkdownTokensToGdoc(tokens);
}

/**
 * Convert the array of tokens from the marked.js MarkDown parser
 * into a series of API calls that generate a Google Doc.
 * 
 * @param {array} tokens An array of tokens as defined in the marked.js MarkDown parser.
 */
function convertMarkdownTokensToGdoc(tokens) {
  for (var i = tokens.length-1; i >= 0; i--) {
    var token = tokens[i];
    translateTokenToGdocInsertAction(token)
  }
}

/**
 * Translate the given token into some kind of insertX() action in
 * the Google Doc.
 * 
 * @param {Object} token A token as defined in the marked.js MarkDown parser.
 */
// TODO: Handle case where token is not found.
function translateTokenToGdocInsertAction(token) {
  switch (token.type) {
    case "paragraph":
      insertParagraph(token.text);
      break;

    case "heading":
      insertHeader(token.text, token.depth);
      break;
  
    case "table":
      insertTable(token)
      break;
  }
}

/**
 * Insert H1.
 */
function insertHeader(text, depth) {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var el = body.insertParagraph(0, text)
  
  switch (depth) {
    case 1:
      el.setHeading(DocumentApp.ParagraphHeading.HEADING1);
      break;
    case 2:
      el.setHeading(DocumentApp.ParagraphHeading.HEADING2);
      break;
    case 3:
      el.setHeading(DocumentApp.ParagraphHeading.HEADING3);
      break;
    case 4:
      el.setHeading(DocumentApp.ParagraphHeading.HEADING4);
      break;
    case 5:
      el.setHeading(DocumentApp.ParagraphHeading.HEADING5);
      break;
    default:
      // Paragraph will be unformatted
  }
}

/**
 * Insert a paragraph.
 */
function insertParagraph(text) {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  body.insertParagraph(0, text)
}

/**
 * Insert a table.
 * 
 * Note that a table token is assumed to be in the following structure:
 * { type: 'table',
 *   header: [ 'foo', 'bar' ],
 *   align: [ null, null ],
 *   cells: [ [ 'baz', 'bim' ], [ 'bee', 'bop' ] ] },
 *   links: {} ] }
 */
function insertTable(token) {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var dataHeader = token.header;
  var data = token.cells;
  
  var rowsData = [];
  rowsData.push(dataHeader);

  data.forEach(function(row) {
    rowsData.push(row);
  });

  table = body.appendTable(rowsData);
  table.getRow(0).editAsText().setBold(true);
}

/**
 * Insert a paragraph in the active doc.
 */
function createDoc() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var rowsData = [['Plants', 'Animals'], ['Ficus', 'Goat'], ['Basil', 'Cat'], ['Moss', 'Frog']];
  body.insertParagraph(0, doc.getName())
      .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  table = body.appendTable(rowsData);
  table.getRow(0).editAsText().setBold(true);
}