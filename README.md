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

---

# Unpack XSN / Pack XSN

A second, independent pair of scripts in the same folder. They take a form apart so you can edit it by hand and put it back together as a working `.xsn`, losing nothing along the way. Use them when you want to *change* a form; use `Analyze-XsnForms` when you only want to *understand* one.

They are separate from the analyzer on purpose. The analyzer rearranges an extracted folder (it moves orphaned files into `_Unused\` and writes reports next to them), which is exactly what you do not want in a folder you intend to repack.

## Why not just rename it to .cab

An `.xsn` is a Microsoft cabinet, so renaming it to `.cab` and extracting does get the files out. What it does not get out is everything in the cabinet that is not file content: the order the files are stored in, each file's timestamp and attribute flags, the compression algorithm, and the cabinet set id. Rebuild without those and you get a valid cabinet that is not the same cabinet.

That difference is real in practice. Of the five sample forms here, three were published from SharePoint and use MSZIP; two were saved from InfoPath and use LZX with a 21-bit window. Their file orders are completely different, and the InfoPath-saved ones carry a `manifest.xsf` stamped seven months later than every other file in the form.

So `Unpack XSN.ps1` parses the cabinet structure itself and writes what it finds to a sidecar file, `_xsn-cab.json`, next to the extracted files. `Pack XSN.ps1` reads that sidecar and rebuilds the cabinet to match.

## How to use it

1. Put both `.ps1` files and both `.bat` files in the folder that holds your `.xsn` files.
2. Double-click `Unpack XSN.bat`. Every `.xsn` in the folder is extracted into a sibling `<name>.unpacked` folder. The `.xsn` files are not modified.
3. Edit the files in that folder - `manifest.xsf` for logic and data connections, `view1.xsl` for layout, the `.xsd` files for schema.
4. Double-click `Pack XSN.bat`. Each `.unpacked` folder is rebuilt into `_Packed\<original name>.xsn`.

Nothing is overwritten: packing writes to `_Packed\`, and your original `.xsn` files stay where they are.

Re-running `Unpack XSN.bat` will not destroy your work. It compares every file against the hash recorded at unpack time and skips any folder that has been edited; delete the folder yourself if you really do want a fresh copy.

## What gets verified

`Pack XSN.ps1` never hands back a file it has not checked. For each form it:

1. re-parses the cabinet it just built and compares the file table with the sidecar - count, order, timestamps, attribute flags;
2. extracts that cabinet to a temp folder and SHA256-compares every file against the file it was built from, which is a byte-level proof that no content changed;
3. compares the finished `.xsn` against the original's SHA256.

On an unedited folder, all five sample forms come back **byte-for-byte identical to the original** - same size, same compressed bytes, same header. Round-tripping a form you have not changed is a no-op, which is the strongest guarantee available that round-tripping one you *have* changed only changed what you edited.

## Editing: what the scripts do about it

| You did this | What happens |
| --- | --- |
| Edited a file | Packed as-is; the size change is reported on screen. |
| Added a file | Appended to the end of the cabinet, with a warning. Remember to also register it in `manifest.xsf` under `<xsf:files>`, or InfoPath will ignore it. |
| Deleted a file | Packed without it, with a warning. |
| Created a subfolder | Ignored, with a warning. A cabinet is flat. |

## Output

| Output | Purpose |
| --- | --- |
| `<form>.unpacked\` | The extracted form, one file per cabinet entry, timestamps restored. |
| `<form>.unpacked\_xsn-cab.json` | The cabinet's structure: header, compression, and the order, size, timestamp, attributes and SHA256 of every file. Do not delete it - `Pack XSN.ps1` needs it. |
| `_Packed\<form>.xsn` | The rebuilt form. |
| `_Unpack-Log.txt`, `_Pack-Log.txt` | Every issue. Both scripts exit non-zero if any error was logged. |

## Requirements and limits

Windows PowerShell 5.1, plus `expand.exe` and `makecab.exe`, both built into Windows. No modules, no internet, no SharePoint connection.

Two limits worth knowing. `makecab` cannot write Quantum-compressed cabinets - nothing has produced those since the 1990s and InfoPath never did, but if one turns up it is rebuilt as LZX:21 and the log says so. And a cabinet carrying per-cabinet reserved data, or one that is part of a multi-cabinet set, is flagged at unpack time as something the rebuild cannot reproduce; neither occurs in a normal `.xsn`.
