# Analyze-XsnForms

A self-contained Windows PowerShell tool that inventories InfoPath form templates (`.xsn`) so you can plan a migration off InfoPath. For each form it extracts the template, parses the logic, data, dependencies and views, then writes a multi-sheet Excel workbook, a plain-text overview, and two Markdown files sized for an LLM. It does not perform the migration. It only reads and reports.

It works on forms *downloaded* from SharePoint Server (on-premises) and SharePoint Online, and it handles both SharePoint **list** forms and form-**library** (XML) forms. This script does not talk to SharePoint at all, so it can be run by anyone with a copy of the XSN form files.

## What you need

The tool relies only on things that ship with Windows, plus one PowerShell module that it installs for you on first run.

| Requirement | Notes |
| --- | --- |
| Windows PowerShell 5.1 | Built into Windows 10/11 and Windows Server. |
| `expand.exe` | Built into Windows. Used to extract the `.xsn` (which is a CAB archive). |
| ImportExcel module | Auto-installed to the current user on first run if it is missing. Needs internet for that one-time install, or pre-install it with `Install-Module ImportExcel -Scope CurrentUser`. |

No SharePoint connection is required. The tool never writes to SharePoint.

## How to use it

1. Copy `Analyze-XsnForms.ps1` and `Analyze-XsnForms.bat` into the folder that holds your `.xsn` files.
2. Double-click `Analyze-XsnForms.bat` (or run the `.ps1` directly).
3. It processes every `.xsn` in the folder automatically. There are no prompts and no parameters.

Each form is extracted into a sibling folder named after the file, and the reports are written inside that folder.

## What it produces

For every form, in the form's extracted folder:

| Output | Purpose |
| --- | --- |
| `<form>-Analysis.xlsx` | The human report. 16 tabbed, sortable sheets (see below). |
| `<form>-Overview.txt` | A short counts-and-flags overview. |
| `LLM Context.md` | The same information as the workbook, written for an LLM, plus a Mermaid view-flow diagram and a cleaned raw-logic appendix. |
| `LLM Context - additional XSL files.md` | All of the form's view layouts (`.xsl`) with namespaces, CSS and geometry stripped, kept separate so the curated context stays small. |
| `_Unused\` | Orphaned and superseded files (old schema versions, debug symbols, sample data, the legacy upgrade transform) moved out of the way. |

At the root of the folder:

| Output | Purpose |
| --- | --- |
| `_AllForms-Summary.xlsx` | One row per form with all counts and a complexity score, sorted hardest-first. |
| `_Analysis-Log.txt` | Every extraction, parse or analyzer issue, so a partially-parsed form is never mistaken for a complete one. The script exits non-zero if any error was logged. |

## The 16 workbook sheets

| Sheet | What it covers |
| --- | --- |
| Summary | Form type, host list, complexity score, and all the counts. |
| Logic | Every rule as plain English: trigger, condition, actions (when / if / then). |
| Fields | Section, label, type, control, SharePoint-vs-XML storage, default value, usage, status. |
| Structure | The section and field hierarchy, in form order. |
| Connections | Every data connection (SharePoint list, SQL, web service, library submit, BCS, email), with SharePoint list IDs. SQL passwords are redacted. |
| Views | The form's views and their `.xsl` files. |
| Visibility | Section show/hide conditions, with a derived Power Fx `Visible` formula. |
| ReadOnly | Fields that become read-only/disabled, and under what condition, with a derived Power Fx `DisplayMode`. |
| Calculations | Calculated fields and dynamic defaults. |
| Validation | Validation rules and their messages. |
| Dropdowns | Choice options, grouped by distinct option-set so repetitive forms stay readable. |
| Navigation | Buttons and events that switch views. |
| RolesIdentity | InfoPath user roles and any identity-based logic. |
| Blockers | Features with no clean migration path (code-behind, SQL, signatures, ActiveX, BCS, UDC, repeating sections, multiple attachments). |
| Simplify | Dead or unused logic: disabled rules, empty rules, unwired rule sets, unused fields. |
| Files | Every extracted file marked active or orphaned, with the reason. |
