var ss = SpreadsheetApp.getActiveSpreadsheet()
var ui = SpreadsheetApp.getUi()

// ---  Script Parameters: see Readme.md for details   ---
const token = 'hexadecimaltokenthingy' // Replace with your Crowdin API token
const org = 'org_name' // Replace with your organization name or leave blank if you're using Crowdin.com
const projectID = 1 // Replace with Project ID (under Tools → API on Crowdin)
const projectLinkID = 'projectnameoralphanumericprojectlinkid' // Replace with project name (Crowdin.com) or project link ID (Enterprise)
// --- --- --- --- --- --- --- --- --- --- --- --- --- ---

const limit = 50
const apiBaseURL = ( org == '' ? 'https://api.crowdin.com/api/v2' : `https://${org}.api.crowdin.com/api/v2` )

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
    if (issueData.user.fullName) {
       issue.user += `\n${issueData.user.fullName}\n\nLang: ${issue.language}`
    }

    issue.fileID = issueData.string.fileId
    issue.fileIDandName = `${issueData.string.fileId}\n\n${files.get(issueData.string.fileId)}`

    issue.link = `https://csp.crowdin.com/translate/${projectLinkID}/all/en-XX#${issue.stringID}`

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
  var lastColumn = dataRange.getLastColumn()
  
  sheet.getRange(2, 1, dataRange.getLastRow(), lastColumn).clearContent()

  issueArray = new Array()
  stringIDsAndLinks = new Array()

  for (i of issues) {
    issueArray.push([i.id, i.fileIDandName, i.stringID, i.date, i.user, i.status, i.type, i.language, i.stringText, i.text, i.context])
    stringIDsAndLinks.push([createRichTextLink(i.stringID, i.link)])

/*  let styleGrey = SpreadsheetApp.newTextStyle().setForegroundColor('#aaa').build()
    let fileIDandName = SpreadsheetApp.newRichTextValue()
                    .setText(i.fileIDandName)
                    .setTextStyle(0, String(i.fileID).length, styleGrey)
                    .build()
    var row = [i.id, fileIDandName, i.stringID, i.date, i.user, i.status, i.type, i.language, i.text, i.stringText, i.context]
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
  sheet.getRange(2, 1, dataRange.getLastRow(), lastColumn).setBorder(false, false, false, false, false, false)
  sheet.getRange(2, 1, issueArray.length, issueArray[0].length).setValues(issueArray)

  sheet.getRange(2, 3, issueArray.length).setRichTextValues(stringIDsAndLinks)

  for (let [index, val] of issues.entries()) {
    if (val.stringText != '') {
      sheet.getRange(index + 2, 1, index + 2, lastColumn)
      .setBorder(true, null, null, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
     /* .setBackground('#eeeeee') */
    }
  }
}

//
// Service Functions
//

function prettylog(object){
  Logger.log(JSON.stringify(object,null,2))
}

function decodeHTMLEntities(string){
  return string.replace(/&(quot);/g, '"').replace(/&(amp);/g, '&')
  // return XmlService.parse(`<d>${string}</d>`).getRootElement().getText() // Doesn't work with Unreal rich text tags (<hunter>...</>)
}

function createRichTextLink(string, link){
  if (link) {
    return SpreadsheetApp.newRichTextValue().setText(string).setLinkUrl(0,String(string).length, link).build()
  }
}
