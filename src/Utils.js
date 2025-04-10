function prettylog(object) {
  Logger.log(JSON.stringify(object, null, 2));
}

function decodeHTMLEntities(string) {
  return string.replace(/&quot;/g, '"').replace(/&amp;/g, "&");
  // return XmlService.parse(`<d>${string}</d>`).getRootElement().getText() // Doesn't work with Unreal rich text tags (<hunter>...</>)
}

function createRichTextLink(string, link) {
  if (link) {
    return SpreadsheetApp.newRichTextValue()
      .setText(string)
      .setLinkUrl(0, String(string).length, link)
      .build();
  }
}
