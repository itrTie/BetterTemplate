// Menu

function onOpen() {
  var ui = DocumentApp.getUi();
  ui.createMenu('Template Editor')
      .addSubMenu(ui.createMenu('Underscores')
          .addItem('Color: Black', 'blackUnderscore')
          .addItem('Color: Red', 'redUnderscore')
          .addItem('Custom Color', 'ccUnderscore')
          .addSeparator()
          .addItem('Style: Bold', 'boldUnderscore')
          .addItem('Style: Italics', 'italicUnderscore')
          .addItem('Style: Underline', 'underlineUnderscore'))
      .addSubMenu(ui.createMenu('Empty Lines')
          .addItem('Color: Black', 'blackLine')
          .addItem('Color: Red', 'redLine')
          .addItem('Custom Color', 'ccLine')
          .addSeparator()
          .addItem('Style: Bold', 'boldLine')
          .addItem('Style: Italics', 'italicLine')
          .addItem('Style: Underline', 'underlineLine'))
      .addToUi();
}

// Underscores - Color

function blackUnderscore() {
  var body = DocumentApp.getActiveDocument().getBody();
  var underscores = body.findText("_");
  while (underscores) {
    var elem = underscores.getElement();
    var start = underscores.getStartOffset();
    var end = underscores.getEndOffsetInclusive();
    elem.setForegroundColor(start, end, "#000000");
    underscores = body.findText("_", underscores);
  }
}

function redUnderscore() {
  var body = DocumentApp.getActiveDocument().getBody();
  var underscores = body.findText("_");
  while (underscores) {
    var elem = underscores.getElement();
    var start = underscores.getStartOffset();
    var end = underscores.getEndOffsetInclusive();
    elem.setForegroundColor(start, end, "#ff0000");
    underscores = body.findText("_", underscores);
  }
}

function ccUnderscore() {
  var ui = DocumentApp.getUi();
  var result = ui.prompt("Insert HEX Code (Include #)");
  var hex = result.getResponseText();
  var body = DocumentApp.getActiveDocument().getBody();
  var underscores = body.findText("_");
  while (underscores) {
    var elem = underscores.getElement();
    var start = underscores.getStartOffset();
    var end = underscores.getEndOffsetInclusive();
    elem.setForegroundColor(start, end, hex);
    underscores = body.findText("_", underscores);
  }
}

// Underscores - Styling

function boldUnderscore() {
  var body = DocumentApp.getActiveDocument().getBody();
  var underscores = body.findText("_");
  while (underscores) {
    var elem = underscores.getElement();
    var start = underscores.getStartOffset();
    var end = underscores.getEndOffsetInclusive();
    elem.setBold(start, end, true);
    underscores = body.findText("_", underscores);
  }
}

function italicUnderscore() {
  var body = DocumentApp.getActiveDocument().getBody();
  var underscores = body.findText("_");
  while (underscores) {
    var elem = underscores.getElement();
    var start = underscores.getStartOffset();
    var end = underscores.getEndOffsetInclusive();
    elem.setItalic(start, end, true);
    underscores = body.findText("_", underscores);
  }
}

function underlineUnderscore() {
  var body = DocumentApp.getActiveDocument().getBody();
  var underscores = body.findText("_");
  while (underscores) {
    var elem = underscores.getElement();
    var start = underscores.getStartOffset();
    var end = underscores.getEndOffsetInclusive();
    elem.setUnderline(start, end, true);
    underscores = body.findText("_", underscores);
  }
}

// Empty Line - Color

function blackLine() {
  var body = DocumentApp.getActiveDocument().getBody();
  var paragraphs = body.getParagraphs();
  var i = 0;

  for (var i = 0; i < paragraphs.length; i++) {
       if (paragraphs[i].getText() === ""){
          paragraphs[i].setForegroundColor("#000000");
       }
  }
}

function redLine() {
  var body = DocumentApp.getActiveDocument().getBody();
  var paragraphs = body.getParagraphs();
  var i = 0;

  for (var i = 0; i < paragraphs.length; i++) {
       if (paragraphs[i].getText() === ""){
          paragraphs[i].setForegroundColor("#FF0000");
       }
  }
}

function ccLine() {
  var ui = DocumentApp.getUi();
  var result = ui.prompt("Insert HEX Code (Include #)");
  var hex = result.getResponseText();
  var body = DocumentApp.getActiveDocument().getBody();
  var paragraphs = body.getParagraphs();
  var i = 0;

  for (var i = 0; i < paragraphs.length; i++) {
       if (paragraphs[i].getText() === ""){
          paragraphs[i].setForegroundColor(hex);
       }
  }
}

// Empty Line - Styling

function boldLine() {
  var body = DocumentApp.getActiveDocument().getBody();
  var paragraphs = body.getParagraphs();
  var i = 0;

  for (var i = 0; i < paragraphs.length; i++) {
       if (paragraphs[i].getText() === ""){
          paragraphs[i].setBold(true);
       }
  }
}

function italicLine() {
  var body = DocumentApp.getActiveDocument().getBody();
  var paragraphs = body.getParagraphs();
  var i = 0;

  for (var i = 0; i < paragraphs.length; i++) {
       if (paragraphs[i].getText() === ""){
          paragraphs[i].setItalic(true);
       }
  }
}

function underlineLine() {
  var body = DocumentApp.getActiveDocument().getBody();
  var paragraphs = body.getParagraphs();
  var i = 0;

  for (var i = 0; i < paragraphs.length; i++) {
       if (paragraphs[i].getText() === ""){
          paragraphs[i].setUnderline(true);
       }
  }
}
