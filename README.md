# Crowdin comments to Google spreadheets
*Google Apps Script to pull comments and issues to a Google Spreadsheet*

This script exports a *static* copy of all the comments and issues from a project on Crowdin. Any edits on the resulting sheet *will not* be reflected on Crowdin, and if you rerun the script, current sheet *will be overwritten* with new data from Crowdin.

It groups all the comments and issues per string and doesn't repeat things like file name, source, and context to make the spreadsheet more readable. That makes filtering harder. Filter-friendly version coming later. It also highlights open source mistake issues with conditional formatting.

Plans:
- Less readable but filter-friendly version
- Separate list(s) of "Translation mistake" issues, split per language or filter-friendly
- Maybe smart filters via scripts (support comment threads, filter by date, issue type, language, etc.)

## Installation
- Open Google Apps Script editor via the menu in your spreadsheet: `Extensions` → `Apps Script`
- Open `Project Settings` using the cog icong on the left toolbar and check the `Show "appscript.json" manifest file in editor` checkbox
- Go back to the Editor using the angled brackets icon on the left toolbar
- Copy the contents of the `appsscript.json` file to the `appsscript.json` file in the editor
- Copy the contents of the `Code.gs` file to the `Code.gs` file in the editor
- Paste your organization name on `line 2` or leave it blank if you're using crowdin.com (see `Crowdin API parameters` below)
- Specify the project ID on `line 3` (see `Crowdin API parameters` below)
- **Don't do this unless you know what you're doing.** You can also paste you API token on `line 7` (see `Crowdin API parameters` below)

**WARNING**: Only save your API token in script source code if you're absolutely certain that your spreadsheet will only be shared within the trusted circle. Token saved in script source code will be available to anyone who has access to the spreadsheet. It is not recommended to save your token if you share the spreadsheet with outsiders.

## Crowdin API parameters
### Organization
- Crowdin.com: Leave it empty ('')
- Enterprise: It's the prefix you see in the URL bar (`https://`*organization*`.crowdin.com/...`)
### Project ID
- You can see it under `Tools` → `API`
### API token
It's best to create a separate account with limited rights and then create an token with limited scope for this.
- You can create or use any account with `Translator` access to create an API token
- Crowdin.com: You can't limit the token scope so create a separate account limited to only one project and role
- Enterprise: Limit scope to `Projects` → `Project` and `Source Files`
- Enterprise: It's also a good idea to enable `Granular Access` and limit the token to a specific project

## Usage
- After you install the script and refresh the spreadsheet, you'll see `Loc Tools` in the Spreadsheets main menu
- Use the `Pull comments and issues from Crowdin` menu item to download all comments and issues from the project and *overwrite the contents of the active sheet*
- The script will ask you for permission to access the spreadsheet it's installed in and to communicate with external services when you run it for the first time
- If you haven't saved the API token within the script, it will ask you for it when you launch it
