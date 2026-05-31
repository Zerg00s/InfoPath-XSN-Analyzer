# Analyze-XsnForms

A self-contained Windows PowerShell tool that inventories InfoPath form templates (`.xsn`) so you can plan a migration off InfoPath. For each form it extracts the template, parses the logic, data, dependencies and views, then writes a multi-sheet Excel workbook, a plain-text overview, and two Markdown files sized for an LLM. It does not perform the migration. It only reads and reports.

It works on forms downloaded from SharePoint Server (on-premises) and SharePoint Online, and it handles both SharePoint **list** forms and form-**library** (XML) forms.

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

## How a field is classified as SharePoint or XML

List forms store every field as a SharePoint list column, so they are all marked `SharePoint column`. Library (XML) forms store the submission as an XML document, and only the fields that InfoPath promotes to columns (declared in `manifest.xsf`) are queryable columns. Everything else is marked `XML only`. The count of XML-only fields feeds the complexity score, because each one is extra work to surface during a migration.

## The complexity score

The score is a weighted sum that reflects rebuild effort, not just size. It counts views, sections, fields, complex fields (lookup / choice / person), repeating sections and fields, XML-only fields, attachment controls, data connections (with extra weight for SQL and cross-list connections), rules and conditions, on-load and field-change handlers, calculations, validation, read-only rules, view navigation, and a large fixed penalty for managed code-behind. The roll-up buckets each form as Low, Medium, High or Very High.

## Safety

The tool is read-only by design. It never writes to SharePoint and never modifies the source `.xsn` files. The only files it moves are junk files inside a folder it just extracted. SQL connection-string passwords are redacted everywhere they would otherwise appear.

## Re-running

Re-running is safe and idempotent. Extraction is skipped when a form is already unpacked, the reports are regenerated in place, and orphan handling is stable across runs. If a workbook is open in Excel (or being synced by OneDrive) when you re-run, the tool waits briefly, then leaves a fresh copy beside the locked file and notes it in the log rather than failing.
"# InfoPath-XSN-Analyzer" 
