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
function generateGdoc(markdownText) {  
  var lexer = new marked.Lexer();
  var parser = new marked.Parser();
  var tokens = lexer.lex(markdownText);
  tokens2gdoc(tokens);
}

function tokens2gdoc(tokens) {
  for (var i = tokens.length-1; i >= 0; i--) {
    var token = tokens[i];
    if (token.type == "paragraph") {
      insertParagraph(token.text);
    }
    else if (token.type == "heading") {
      insertHeader(token.text, token.depth);
    }
  }
}

/**
 * Insert H1.
 */
function insertHeader(text, depth) {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var el = body.insertParagraph(0, text)
  if (depth == 1) {
    el.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  }
  else if (depth == 2) {
    el.setHeading(DocumentApp.ParagraphHeading.HEADING2);
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

/**
 * Opens a sidebar in the document containing the add-on's user interface.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 */
function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('md2gdoc');
  DocumentApp.getUi().showSidebar(ui);
}