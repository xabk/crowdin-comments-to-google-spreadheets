# Crowdin comments to Google spreadsheets
*Google Apps Script to pull comments and issues to a Google Spreadsheet*

This script exports a *static* copy of all the comments and issues from a project on Crowdin. Any edits on this sheet *will not* be reflected on Crowdin, and if you rerun the script, it *will overwrite* your edits with new data from Crowdin.

It groups all the comments and issues per string and hides repeating things like file name, source, and context to make the spreadsheet more readable. It does this by setting the text color white, so the text is still there and you can filter properly. It also highlights open source mistake issues using conditional formatting.

Plans:
- [x] ~~Less readable but filter-friendly version~~ Filter-friendly version
- [x] Separate list(s) of "Translation mistake" issues, split per language or filter-friendly
- [x] Separate sheets for source mistakes
- [ ] More smart filters via scripts (support comment threads, filter by date, issue type, language, etc.)
- [ ] Two-way sync between spreadsheet and Crowdin to close issues and add your answers/comments

## Installation
### By copying a template (e.g., into a new spreadsheet)
- Copy this spreadsheet: https://docs.google.com/spreadsheets/d/1X9ob_hB1r87uv2ZTu5xTJ3rfOfULA3oGfEaIPLpwCpI/edit#gid=0
- Set parameters, see below

### Manually (e.g., into an existing spreadsheet)
- Open Google Apps Script editor via the menu in your spreadsheet: `Extensions` → `Apps Script`
- Open `Project Settings` (cog icon on the left toolbar) and tick the `Show "appscript.json" manifest file in editor` checkbox
- Go back to the Editor (angled brackets icon on the left toolbar)
- Copy the contents of the `appsscript.json` file here to the `appsscript.json` file in the editor
- Copy the contents of the `Code.gs` file here to the `Code.gs` file in the editor
- Set parameters, see below

## Parameters
### Security considerations
There is no clear answer as to what is a more secure way in a real-world environment with real people.

On one hand, if each user creates their own token and it's saved per user, these tokens won't leak via this script/spreadsheet. But people tend to be lazy with tokens and scopes, especially if it requires creating a separate account, and less tech-savvy users tend to save those tokens where they can be accessed by other people.

On the other hand, hardcoding your token in the script exposes it to anyone using the spreadsheet. But if this token is limited to the current project and has a limited scope, it's not that much of a deal, as long as your spreadsheet is shared with trusted people.

### Saving parameters in properties
- Set organization via `Loc Tools` → `View or set Crowdin org (all users)`
- Set project ID via `Loc Tools` → `View or set Crowdin project ID (all users)`
- Set token via `Loc Tools` → `Set Crowdin API key (current Google user)`

The organization name and project ID will be saved in script parameters for all users of the spreadsheet. The token will be saved in user parameters for each user separately for security reasons. Each user will have to provide their own token.

### Hardcoding parameters in the script
- Go to `Extensions` → `Apps Script` → `Code.gs`
- Hardcode your organization name or set it to `""` if you're using crowdin.com
- Hardcode the project ID
- **Don't do this unless you know what you're doing.** You can also hardcode your API token

**WARNING**: Only hardcode your API token in the script source code if you're certain that your spreadsheet will only be shared within the trusted circle. Token saved in script source code will be available to anyone who has even view-only access to the spreadsheet. It is not recommended to save your token if you share the spreadsheet with outsiders.

## Crowdin API parameters
### Organization
- Crowdin.com: Leave it empty ("")
- Enterprise: It's the prefix you see in the URL bar (`https://`*organization*`.crowdin.com/...`)

### Project ID
- You can see it under `Tools` → `API`

### API token
It's best to create a separate account with limited rights and then create a token with limited scope for this.
- Create or use any account with `Translator` access to create an API token
- Crowdin.com: Create a separate translator account with access to only one project
- Enterprise: Limit scope to `Projects` → `Project` and `Source Files`
- Enterprise: It's also a good idea to enable `Granular Access` and limit the token to a specific project

## Usage
- After you install the script and refresh the spreadsheet, you'll see `Loc Tools` in the Spreadsheets main menu
- Use the `Overwrite current sheet with comments from Crowdin` menu item to download all comments and issues from the project and *overwrite the contents of the active sheet*
- The script will ask you for permission to access the spreadsheet it's installed in and to communicate with external services when you run it for the first time
- If you haven't set up the parameters one way or another, it will ask you to do it when you first launch it
