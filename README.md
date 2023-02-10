# Crowdin comments to Google spreadheets
*Google Apps Script to pull comments and issues to a Google Spreadsheet*

#TODO: Test both on crowdin.com and Enterprise =)

This script exports a *static* copy of all the comments and issues from a project on Crowdin. Any edits on the resulting sheet *will not* be reflected on Crowdin, and if you rerun the script, they *will be overwritten* with new data from Crowdin.

**Warning: your API token will be available to anyone you share the spreadsheet with. So it's best to create a separate account with limited rights for this, and also limit the API token scope to the absolute minimum (see below).**

## Installation
- Open Google Apps Script editor via the menu in your spreadsheet: `Extensions` → `Apps Script`
- Open `Project Settings` using the cog icong on the left toolbar and check the `Show "appscript.json" manifest file in editor` checkbox
- Go back to the Editor using the angled brackets icon on the left toolbar
- Copy the contents of the `appsscript.json` file here to the `appsscript.json` file in the editor
- Copy the contents of the `Code.gs` file here to the `Code.gs` file in the editor
- Paste your organization name on `line 6` or leave it blank if you're using crowdin.com (see `Crowdin API parameters` below)
- Specify the project ID on `line 7` (see `Crowdin API parameters` below)

- **Don't do this unless you know what you're doing.** You can also paste you API token on `line 5` (see `Crowdin API parameters` below). 

**WARNING**: Only save your API token in script source code if you're absolutely certain that your spreadsheet will only be shared within the trusted circle. Token saved in script source code will be available to anyone who has access to the spreadsheet. It is not recommended to save your token if you share the spreadsheet with outsiders.

## Crowdin API parameters
### API token
- You can use an account with `Translator` rights or above to create an API token
- Scopes needed: `Projects` → `Project` and `Source Files`
- It's also a good idea to enable `Granular Access` and limit the token to a specific project
### Project ID
- You can see it under `Tools` → `API`
### Project Link ID
- This it needed to generate direct links to strings so that you could easily open them on Crowdin.
- It's the part of the file URL that goes after `translate` or `proofread`.
- Open any file, take a look at the URL, you need the highlighted part:
- Crowdin.com: `https://crowdin.com/translate/`*projectname*`/80/en-enprt` — or you can just use the project name from Settings
- Crowdin Enterprise: `https://organization.crowdin.com/translate/`*a8a0b3b69ccb4d4460720ca9ffccc813*`/all/en-sv/49`

## Usage
- After you install the script and refresh the spreadsheet, you'll see `Loc Tools` in the Spreadsheets main menu
- Use the `Pull comments and issues from Crowdin` menu item to download all comments and issues from the project and *overwrite the contents of the active sheet*
- The script will ask you for permission to access the spreadsheet it's installed in and to communicate with external services when you run it for the first time
- If you haven't saved the API token within the script, it will ask you for it when you launch it
