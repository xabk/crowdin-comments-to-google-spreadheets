function prettylog(object) {
  Logger.log(JSON.stringify(object, null, 2));
}

function decodeHTMLEntities(string) {
  return string
    .replace(/&quot;/g, '"')
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&#39;/g, "'");
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

function createRichTextWithTwoLinks(text1, link1, text2, link2) {
  const builder = SpreadsheetApp.newRichTextValue().setText(
    `${text1}\n\n${text2}`
  );
  if (text1 && link1) {
    builder.setLinkUrl(0, text1.length, link1);
  }
  if (text2 && link2) {
    builder.setLinkUrl(
      text1.length + 2,
      text1.length + 2 + text2.length,
      link2
    );
  }
  return builder.build();
}

function prepareIssueLinks(
  issue,
  addLinkUrlTemplate,
  addLinkText,
  addLinkEnabled
) {
  const stringLink = issue.link.replace("/en-XX", `/en-${issue.language}`);
  const addLink = addLinkUrlTemplate
    ? addLinkUrlTemplate
        .replace("%term%", issue.stringKey || "")
        .replace("%lang%", issue.language.toLowerCase())
    : null;

  if (
    addLinkEnabled &&
    addLinkText &&
    addLink &&
    issue.stringID &&
    stringLink
  ) {
    return createRichTextWithTwoLinks(
      addLinkText,
      addLink,
      `${issue.stringID}`,
      stringLink
    );
  } else {
    return createRichTextLink(issue.stringID, issue.link);
  }
}
