// ---  Script Parameters: see Readme.md for details   ---
var token = '' // Replace with your Crowdin API token or leave empty and you'll be prompted for it every time you use the script
// WARNING: only you save your API token here if the spreadsheet is never shared outside your trusted circle
// WARNING: API token saved here will be freely available to anyone who has access to the spreadsheet
const org = '' // Replace with your organization name or leave blank if you're using Crowdin.com
const projectID = 1 // Replace with Project ID (under Tools → API on Crowdin)
const colWidths = [40, 100, 50, 100, 110, 80, 80, 230, 450, 450]
// --- --- --- --- --- --- --- --- --- --- --- --- --- ---

const wrClip = SpreadsheetApp.WrapStrategy.CLIP
const wrWrap = SpreadsheetApp.WrapStrategy.WRAP

const header = [['iID', 'File ID / \nName', 'Str\nID', 'Date', 'User', 'Status', 'Issue Type', 'String', 'Issue/Comment', 'Context']]
const colWidths = [40, 100, 50, 100, 110, 80, 80, 230, 450, 450]
const colFontColors = ['#555', '#555', '#1155cc', '#555', '#555', '#555', '#555', '#000', '#000', '#000']
const wrapStrategies = [wrClip, wrClip, wrClip, wrClip, wrClip, wrWrap, wrWrap, wrWrap, wrWrap, wrClip]

const limit = 50
const apiBaseURL = ( org == '' ? 'https://api.crowdin.com/api/v2' : `https://${org}.api.crowdin.com/api/v2` )

var ss = SpreadsheetApp.getActiveSpreadsheet()
var ui = SpreadsheetApp.getUi()

//
// Get API token from user
//

function getTokenPrompt() {
  var ui = SpreadsheetApp.getUi()
  var result = ui.prompt("Please enter the Crowdin API token (don't share your token anywhere!)")
  
  //Get the button that the user pressed.
  var button = result.getSelectedButton()
  
  if (button === ui.Button.OK) {
    Logger.log("Got the token from the user.")
    return result.getResponseText()
  } else if (button === ui.Button.CLOSE) {
    Logger.log("Script cancelled, user haven't supplied the API token")
    return null
  }
    
}

//
// Get Project Link ID
// (project name on Crowdin.com or alphanumeric project ID used for URLs in Crowdin Enterprise)
//

function getProjectLinkID() {
  if(token == ''){
    token = getTokenPrompt()
    if(token == null){
      return null
    }
  }
  var url = `${apiBaseURL}/projects/${projectID}`
  var options = {
    'muteHttpExceptions': true,
    'headers': {
      'Authorization': `Bearer ${token}`
    }
  }
  var response = UrlFetchApp.fetch(url, options)

  // Make request to API and get response before this point.
  var json = response.getContentText();
  var data = JSON.parse(json);
  return data.data.identifier
}

//
// Adding functions to the menu
//
 
function onOpen() {
  ui.createMenu('SoC Tools')
  .addItem('Pull comments and issues from Crowdin', 'overwriteWithIssuesFromCrowdin')
//  .addSeparator()
  .addToUi()
}

//
// Downloading file list from Crowdin (to have names)
//

function downloadFileNamesFromCrowdin() {
  var url = `${apiBaseURL}/projects/${projectID}/files?limit=500`
  var options = {
    'muteHttpExceptions': true,
    'headers': {
      'Authorization': `Bearer ${token}`
    }
  }
  var response = UrlFetchApp.fetch(url, options)

  // Make request to API and get response before this point.
  var json = response.getContentText();
  var data = JSON.parse(json);

  var files = new Map()
  
  data.data.forEach(function(item){
    if (item.data.title) {
      files.set(item.data.id, item.data.title)
    } else {
      files.set(item.data.id, item.data.name)
    }
  })

  return files
}

//
// Downloading issues and comments from Crowdin
//

function downloadIssuesFromCrowdin() {
  var projectLinkID = getProjectLinkID()
  var offset = 0
  var url = `${apiBaseURL}/projects/${projectID}/comments?limit=${limit}&offset=`
  var options = {
    'muteHttpExceptions': true,
    'headers': {
      'Authorization': `Bearer ${token}`
    }
  }

  data = new Array()
  
  while (true){
    response = UrlFetchApp.fetch(url + offset, options)
    json = response.getContentText();
    responseData = JSON.parse(json);
    if (responseData.data.length == 0) { break }

    data.push(...responseData.data)

    if (responseData.data.length < limit) { break }

    offset += limit
  }

  var files = downloadFileNamesFromCrowdin()

  var issues = new Array()

  for (issueData of data) {
    issueData = issueData.data
    let issue = new Object()

    issue.id = issueData.id
    issue.date = new Date(issueData.createdAt)
    // issue.dateString = new Intl.DateTimeFormat("en" , {dateStyle: "short"}).format(issue.date)
    
    issue.text = decodeHTMLEntities(issueData.text)
    issue.context = decodeHTMLEntities(issueData.string.context)

    issue.language = issueData.languageId.charAt(0).toUpperCase() + issueData.languageId.slice(1)

    if (issueData.type == 'comment'){
      issue.status = 'Comment'
      issue.type = 'Comment'
    } else {
      issue.status = issueData.issueStatus.replace(/unresolved/,'Open').replace(/resolved/, 'Closed')
      issue.type = (issueData.issueType.charAt(0).toUpperCase() + issueData.issueType.slice(1)).replace('_',' ')
    }

    issue.stringID = issueData.stringId
    issue.stringText = decodeHTMLEntities(issueData.string.text)

    issue.userID = issueData.userId
    issue.user = issueData.user.username
    if (issueData.user.fullName & issueData.user.fullName != issueData.user.username) {
       issue.user += `\n${issueData.user.fullName}`
    }
    issue.user += `\n\nLang: ${issue.language}`

    issue.fileID = issueData.string.fileId
    issue.fileIDandName = `${issueData.string.fileId}\n\n${files.get(issueData.string.fileId)}`

    issue.link = ( org == '' ? `https://crowdin.com/translate/${projectLinkID}/all/en-XX#${issue.stringID}`
                             : `https://${org}.crowdin.com/translate/${projectLinkID}/all/en-XX#${issue.stringID}` )

    issues.push(issue)
  }

  issues.sort((firstEl, secondEl) => {
    if (firstEl.date != secondEl.date) {
      return firstEl.date - secondEl.date
    } else {
      return firstEl.id - secondEl.id
    }
  })

  var issuesMap = new Map()
  for (issue of issues) {
    if (!issuesMap.has(issue.stringID)) {
            issuesMap.set(issue.stringID, [issue])
        } else {
            issue.fileIDandName = ''
            issue.stringText = ''
            issue.context = ''
            issuesMap.get(issue.stringID).push(issue)
        }
  }

  issues = new Array()

  for ([key, value] of issuesMap) {
    for (item of value) {
      issues.push(item)
    }
  }

  return issues
}

function overwriteWithIssuesFromCrowdin() {
  var issues = downloadIssuesFromCrowdin()

  var sheet = SpreadsheetApp.getActiveSheet()
  var dataRange = sheet.getDataRange()
  var lastColumn = header[0].length
  var maxColumns = sheet.getMaxColumns()
  
  if(maxColumns < header[0].length){
    sheet.insertColumnsAfter(maxColumns, header[0].length - maxColumns)
  }
  
  sheet.getRange(1,1,1,lastColumn).setValues(header).setFontColor('#f3f3f3').setBackground('#434343').setFontWeight('bold');
  colWidths.map((width, col) => {sheet.setColumnWidth(col + 1, width)})
  sheet.setFrozenRows(1);
  
  sheet.getRange(2, 1, sheet.getMaxRows() - 1, sheet.getMaxColumns()).clearContent()
  sheet.getRange(2, 1, sheet.getMaxRows() - 1, sheet.getMaxColumns()).setBorder(false, false, false, false, false, false)
  
  issueArray = new Array()
  stringIDsAndLinks = new Array()

  for (i of issues) {
    issueArray.push([i.id, i.fileIDandName, i.stringID, i.date, i.user, i.status, i.type, i.stringText, i.text, i.context])
    stringIDsAndLinks.push([createRichTextLink(i.stringID, i.link)])

/*  let styleGrey = SpreadsheetApp.newTextStyle().setForegroundColor('#aaa').build()
    let fileIDandName = SpreadsheetApp.newRichTextValue()
                    .setText(i.fileIDandName)
                    .setTextStyle(0, String(i.fileID).length, styleGrey)
                    .build()
    var row = [i.id, fileIDandName, i.stringID, i.date, i.user, i.status, i.type, i.text, i.stringText, i.context]
    row = row.map(function(item){
      if (item.getRuns) {
        return item
      } else {
        return SpreadsheetApp.newRichTextValue().setText(item).build()
      }
    })

    issueArray.push(row) */
  }

  sheet.getRange(2, 1, dataRange.getLastRow(), lastColumn).clearContent()
  sheet.getRange(2, 1, issueArray.length, issueArray[0].length).setValues(issueArray)

  sheet.getRange(2, 3, issueArray.length).setRichTextValues(stringIDsAndLinks)

  for (let [index, val] of issues.entries()) {
    if (val.stringText != '') {
      sheet.getRange(index + 2, 1, index + 2, lastColumn)
      .setBorder(true, null, null, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
     /* .setBackground('#eeeeee') */
    }
  }

  wrapStrategies.map((strat, col) => {sheet.getRange(2, col + 1, sheet.getDataRange().getLastRow(), 1).setWrapStrategy(strat)})
  colFontColors.map((color, col) => {sheet.getRange(2, col + 1, sheet.getDataRange().getLastRow(), 1).setFontColor(color)})
  
}

//
// Service Functions
//

function prettylog(object){
  Logger.log(JSON.stringify(object,null,2))
}

function decodeHTMLEntities(string){
  return string.replace(/&quot;/g, '"').replace(/&amp;/g, '&')
  // return XmlService.parse(`<d>${string}</d>`).getRootElement().getText() // Doesn't work with Unreal rich text tags (<hunter>...</>)
}

function createRichTextLink(string, link){
  if (link) {
    return SpreadsheetApp.newRichTextValue().setText(string).setLinkUrl(0,String(string).length, link).build()
  }

/*  var styleGrey = SpreadsheetApp.newTextStyle().setForegroundColor('#aaa').build()
  var fileIDandName = SpreadsheetApp.newRichTextValue()
                      .setText(i.fileIDandName)
                      .setTextStyle(0, String(i.fileID).length, styleGrey)
                      .build()
*/

//  return SpreadsheetApp.newRichTextValue().setText(item).build()
}
