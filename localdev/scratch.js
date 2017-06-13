  var marked = require('marked');
  var lexer = new marked.Lexer();
  var parser = new marked.Parser();
  
  var markdownText = `## Hello

  This is a test.

| foo | bar |
| --- | --- |
| baz | bim |
  `
  var tokens = lexer.lex(markdownText);

  console.log(tokens);
  console.log(tokens[2].cells)