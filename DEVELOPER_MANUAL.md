# Automated Appeal Letter Tool - Developer Manual

This document explains how `AppealLetterGenerator.py` is organized, how the main workflows run, and where to make changes safely. It is written for a programmer who needs to debug, maintain, or extend the app without having to rediscover the whole workflow from scratch.

## Table Of Contents

- Project Shape
- Quick Start For Developers
- Main Dependencies
- Source Map
- Runtime Folders And Files
- Configuration Model
- UI Tabs
- Template Placeholders
- Main Generation Workflow
- Grouped Generation
- Output Folder Structure
- Mail Output Copying
- Fax Ready.txt Logic
- Failure Handling
- CM Note Extraction
- Common Change Recipes
- Safe Testing Checklist
- Known Code Notes

## Project Shape

The application is a single-file Tkinter desktop app:

- Main source: `AppealLetterGenerator.py`
- Runtime config: `~/Desktop/Automated Appeal Letter Generation/Config/merge_config.json`
- Runtime folders:
  - `Forms`
  - `Letter Templates`
  - `Fax Coversheets`
  - `Logs`
  - `Config`

The app creates those folders under:

```text
~/Desktop/Automated Appeal Letter Generation
```

If the Desktop path is unavailable, it falls back to the executable/source folder.

There is no repository metadata in this folder at the time this manual was written, so treat local file backups carefully before large edits.

## Quick Start For Developers

From this project folder:

```powershell
python -m py_compile AppealLetterGenerator.py
python AppealLetterGenerator.py
```

If dependencies are missing, install the packages used by the app:

```powershell
pip install pandas openpyxl python-docx docxtpl pymupdf jinja2
```

The app stores user configuration and managed template files outside the project folder, under the user's Desktop runtime folder:

```text
~/Desktop/Automated Appeal Letter Generation
```

When debugging a user issue, always check both places:

- the source file in this project folder
- the generated runtime config/log/template folders on the Desktop

## Main Dependencies

The code imports and uses:

- `tkinter` for UI
- `pandas` for Excel/report data
- `openpyxl` for Excel output/table formatting
- `docxtpl` for rendering Word templates with Jinja-style placeholders
- `python-docx` for reading generated letters during CM note extraction
- `fitz` / PyMuPDF for PDF page counts
- standard library modules for filesystem, threading, JSON config, logging, etc.

## Source Map

`AppealLetterGenerator.py` is organized roughly like this:

- Top-level constants and path setup:
  - `APP_DESKTOP_FOLDER_NAME`
  - `APP_SUBDIRECTORIES`
  - `USER_PATHS`
  - log/config path constants
- Top-level utility functions:
  - column-name cleanup
  - value formatting
  - path sanitizing
  - contact/address formatting
  - page counting
- `FormattedValue`:
  - wraps Excel cell values so templates render formatted strings while comparisons can still use raw values
- `MergeApp`:
  - main Tkinter application class
  - owns UI widgets, configuration, loaded Excel data, generation, note extraction, and template management

High-value methods to know first:

- `create_widgets()`: creates all notebook tabs.
- `load_config()` / `save_config()`: persistence boundary.
- `load_excel()` / `_load_excel_data()`: report loading.
- `validate_common_inputs()`: validates required user selections before generation.
- `start_generate_appeals_thread()`: UI entry point for appeal generation.
- `run_generate_appeals()`: main generation orchestrator.
- `get_generation_row_groups()`: internal grouping by Claim User Group and Contract Code.
- `_lookup_mapping_row()`: finds the routing/method mapping for a group.
- `_resolve_required_templates()`: finds letter/fax/form templates for a group.
- `copy_mail_outputs_to_mailing_folders()`: Mail handoff copy step.
- `extract_notes_from_successes()`: note extraction and Fax `Ready.txt` handling.
- `export_failures()`: failed-row workbook export.

## Important Files And Paths

Defined near the top of `AppealLetterGenerator.py`:

- `CONFIG_FILE_PATH`: persisted app configuration JSON.
- `DEBUG_LOG_PATH`: debug log file.
- `CRASH_LOG_PATH`: crash traceback log.
- `ERROR_LOG_PATH`: UI error log persistence.
- `APP_SUBDIRECTORIES`: runtime folders the app maintains.

## Runtime Folders And Files

The runtime root is:

```text
~/Desktop/Automated Appeal Letter Generation
```

Children:

- `Forms`: managed form templates.
- `Letter Templates`: managed letter templates.
- `Fax Coversheets`: managed fax cover sheet templates.
- `Logs`: debug, crash, and merge error logs.
- `Config`: persisted JSON configuration.

Main files:

- `Config/merge_config.json`: user selections, mappings, templates, blurbs, and column formats.
- `Logs/app_debug.log`: debug output.
- `Logs/app_crash.log`: fatal exception tracebacks.
- `Logs/merge_errors.log`: persisted UI error entries.

Generated failed-row reports are not written to the runtime root. They are exported to:

```text
~/Downloads/Failed_Rows_<timestamp>.xlsx
```

## Configuration Model

`MergeApp.__init__()` creates `self.app_config`, then overlays saved JSON through `load_config()`.

Important config keys:

- `excel_path`: last selected report file.
- `output_path`: main output folder.
- `templates`: template maps by type:
  - `letter`
  - `fax`
  - `form`
- `column_formats`: per-Excel-column formatting rules.
- `blurbs`: reusable letter blurbs and subjects.
- `auto_detect_blurb`: whether rows use the `RationalName` Excel column.
- `selected_blurb`: manual dropdown selection when auto-detect is off.
- `fax_info`: sender info. This is forced to `DEFAULT_FAX_INFO`.
- `special_cols`: user-selected report columns:
  - `account_col`
  - `claim_user_group_col`
  - `contract_code_col`
  - `contract_name_col`
  - `dos_col`
- `special_col_profiles`: saved required-column mappings by workbook column layout.
- `extract_config`: CM note extraction start/end markers.
- `mapping_rows`: method routing table.
- `acronym_rows`: Claim User Group acronym and mailing-folder table.

`save_config()` writes the config back to JSON. It also syncs current UI values into `self.app_config`.

Approximate config shape:

```json
{
  "excel_path": "C:/path/to/report.xlsx",
  "output_path": "C:/path/to/output",
  "templates": {
    "letter": {
      "Claim Group Name": "C:/.../Claim Group Name - Letter template.docx"
    },
    "fax": {
      "Claim Group Name": "C:/.../Claim Group Name - Fax.docx"
    },
    "form": {
      "Claim Group Name|CONTRACT": "C:/.../Claim Group Name - Form - CONTRACT.docx"
    }
  },
  "column_formats": {
    "Column Name": "Text"
  },
  "blurbs": {
    "RationalName": {
      "subject": "Subject text",
      "body": "Body with {{ placeholders }}"
    }
  },
  "auto_detect_blurb": true,
  "selected_blurb": "",
  "special_cols": {
    "account_col": "Account Number",
    "claim_user_group_col": "Claim User Group",
    "contract_code_col": "Contract Code",
    "contract_name_col": "Contract Name",
    "dos_col": "DOS"
  },
  "special_col_profiles": {
    "account_number|claim_user_group|contract_code|contract_name|dos": {
      "account_col": "Account Number",
      "claim_user_group_col": "Claim User Group",
      "contract_code_col": "Contract Code",
      "contract_name_col": "Contract Name",
      "dos_col": "DOS"
    }
  },
  "mapping_rows": [],
  "acronym_rows": []
}
```

Use the app UI to create config data when possible. Manual JSON edits are useful for debugging, but the UI may normalize values when it loads or saves.

## UI Tabs

Tabs are created in `create_widgets()`:

- Main
- CM Note Extraction
- Column Configuration
- Client Mapping
- Rationale Editor
- Templates

### Main Tab

Created by `create_main_tab()`.

Purpose:

- Select Excel report.
- Select output folder.
- Select/manual configure rationale behavior.
- Set letter/file dates.
- Start appeal generation.

Rationale behavior:

- Checkbox: `Auto-detect rationale from RationalName Column`
- Default: enabled
- If enabled, each row must have `RationalName` matching a saved rationale.
- If disabled, `combo_blurb` is enabled and one rationale is used for all rows.

### Column Configuration Tab

Created by `create_column_config_tab()` and populated by `populate_column_config_tab()`.

Purpose:

- Select required report columns.
- Configure data formatting for detected Excel columns.

Column format controls are rendered by `_render_column_format_controls()` in a responsive multi-column grid.

### Client Mapping Tab

Created by `create_mapping_table_tab()`.

Contains two tables:

#### Method Mapping

Columns:

- `Claim User Group`
- `Contract Code`
- `Payer Name`
- `Attention To`
- `Method 1 Type`
- `Method 1 Contact`
- `Method 2 Type`
- `Method 2 Contact`
- `Method 3 Type`
- `Method 3 Contact`
- `MethodToUse`
- `Form?`

The Add/Edit UI is `_open_mapping_row_form()`.

The selected method is parsed by `_parse_selected_method()`.

Valid method types:

- `Fax`
- `Mail`

#### Claim User Group Acronyms

Columns:

- `Claim User Group`
- `Acronym`
- `Mailing Folder`

The Add/Edit UI is `_open_acronym_row_form()`.

`Mailing Folder` is used only for Mail outputs. Generated Mail files are copied there after a successful run.

### Rationale Editor

Created by `create_blurb_editor_tab()`.

Rationales are stored in `self.app_config["blurbs"]`.

Each rationale can have:

- name
- subject
- body

The body is rendered using Jinja syntax against row data in `build_letter_merge_dict()`.

### Templates Tab

Created by `create_template_manager_tab()`.

Templates are stored by type in `self.app_config["templates"]`.

Expected filenames:

- Letter: `<Claim User Group> - Letter template.docx`
- Fax: `<Claim User Group> - Fax.docx`
- Form: `<Claim User Group> - Form - <Contract Code>.docx`

Template save logic:

- Builds expected filename.
- Moves/renames selected template into the correct default folder.
- Updates `self.app_config["templates"]`.

## Template Placeholders

Templates are rendered by `create_and_save_doc()` using `DocxTemplate`.

Common non-Excel placeholders include:

- `TodayDate`
- `FileDate`
- `AddressBlock`
- `BlurbBlock`
- `SubjectBlock`
- `Patient_FirstName`
- `Patient_LastName`
- `payer_name`
- `payer_fax`
- `payer_attn`
- `total_pages` for fax templates
- sender keys from `DEFAULT_FAX_INFO`, such as `user_name`, `user_phone`, `user_fax`

`validate_template_fields()` checks template variables against:

- Excel columns
- app/system placeholders listed above

## Address And Payer Contact Formatting

Relevant helpers:

- `clean_text_value()`
- `split_contact_lines()`
- `build_contact_address_block()`

`AddressBlock` is built from:

1. Payer Name
2. `Attention To: ...`
3. For Fax: `Fax: <number>`
4. For Mail: multi-line mailing address/contact text

Mail contacts preserve line breaks from the Client Mapping.

Fax cover sheets also receive:

- `payer_name`
- `payer_fax`
- `payer_attn`

## Main Generation Workflow

Entry point:

```python
start_generate_appeals_thread()
```

It validates common inputs, disables buttons, and starts:

```python
run_generate_appeals()
```

### Validation Before Running

`validate_common_inputs()` checks:

- Excel file selected
- Output folder selected
- File date present
- Account column selected
- Claim User Group column selected
- Contract Code column selected
- Contract Name column selected
- DOS column selected
- Excel data loaded
- Required selected columns exist in the loaded DataFrame
- `Workflow_Reason` exists
- If auto-detect rationale is enabled, `RationalName` exists

## Grouped Generation

Rows are grouped by:

```text
Claim User Group + Contract Code
```

Implemented in:

```python
get_generation_row_groups()
```

Groups are sorted by Claim User Group, then Contract Code. Within each group, original Excel row order is preserved.

Before processing a group, `run_generate_appeals()` validates group-level requirements once:

- Claim User Group / Contract Code present
- mapping row exists
- selected method is valid
- required templates exist
- if method is Mail, Mailing Folder exists

If a group-level requirement fails, every row in that group is marked failed with the same reason and the app skips to the next group.

Per-row checks still happen inside the group, including:

- row-specific `RationalName`
- letter render success
- form/fax render success

## Output Folder Structure

Generated documents go under:

```text
<Output Folder>/
  <Claim User Group>/
    Faxed Letters | Mailed Letters/
      <Contract Code> - <Contract Name>/
        <Year>/
          <Month>/
            <Day>/
              <Account>/
                <FileDate>.<Acronym>.<Account> - Letter.docx
                <FileDate>.<Acronym>.<Account> - Fax.docx
                OR
                <FileDate>.<Acronym>.<Account> - Form.docx
```

The app also writes/maintains `Fax Number.txt` in the contract folder for Fax groups.

## Mail Output Copying

After successful generation, Mail outputs are copied to the Claim User Group's `Mailing Folder`.

Implemented by:

- `copy_mail_outputs_to_mailing_folders()`
- `_copy_file_flat_unique()`

Rules:

- Copies files only, not folder structures.
- Does not move the original files.
- Copies both letter and secondary document.
- If a file name already exists, the copy gets ` (2)`, ` (3)`, etc.

Important: Do not add `Ready.txt` behavior here unless the business process explicitly changes. `Ready.txt` currently belongs to the Fax flow only.

## Fax Ready.txt Logic

Fax `Ready.txt` behavior is in `extract_notes_from_successes()`.

The code groups successful rows by day folder and method. For Fax successes, it creates a `Ready.txt` marker in the day folder if it does not already exist.

This logic should remain Fax-only unless requirements change.

## Failure Handling

Failures are stored as tuples:

```python
(row_index, "Failure reason")
```

Examples:

- `Mapping not found`
- `Invalid MethodToUse value`
- `Letter template missing`
- `Fax template missing`
- `Form template missing`
- `Missing RationalName column`
- `RationalName '<name>' does not match an existing rationale`
- `Secondary document generation error`
- `Missing Mailing Folder for Claim User Group`

### Partial Output Cleanup

If the letter is created but the secondary document fails, `cleanup_failed_output()` is called.

It:

- deletes partial generated files, currently the letter
- removes the account folder if empty
- walks upward removing empty parent folders
- stops at the configured output root

This means an empty day folder created only by a failed row is removed. A month folder with other valid day folders remains.

### Failure Export

Implemented by:

```python
export_failures()
```

Exports to:

```text
~/Downloads/Failed_Rows_<timestamp>.xlsx
```

The workbook contains:

- all original Excel columns
- `Failure Reason`

If failures occurred, `run_generate_appeals()` shows a warning popup instead of a normal success-only message.

## CM Note Extraction

Implemented by:

```python
extract_notes_from_successes()
```

For each successful letter:

1. Opens the generated Word letter.
2. Extracts text between configured start/end markers.
3. Builds note rows grouped by day folder and method.
4. Writes notes workbook(s) into day folders.
5. For Fax outputs, creates `Ready.txt` in the day folder.

Extraction markers are configured on the CM Note Extraction tab and stored in:

```python
self.app_config["extract_config"]
```

## Error Log UI

The app has an error queue/log system:

- `queue_error()`
- `record_error_entry()`
- `load_persisted_error_log()`
- `display_error_entry()`
- `clear_error_log()`

These write to:

```text
Logs/merge_errors.log
```

This is separate from the failed-row workbook.

## Threading Model

Long-running actions use background threads:

- generation
- extraction
- bot upload

Typical pattern:

1. Validate input on UI thread.
2. Disable buttons with `set_buttons_state("disabled")`.
3. Start background thread.
4. Re-enable buttons in `finally`.

Be careful updating Tkinter widgets from worker threads. Existing code does some UI updates from worker paths; if making larger changes, prefer queueing UI updates back to the main thread with `after()`.

## Adding Or Changing Placeholders

If adding a new placeholder:

1. Add it to the merge context in `build_letter_merge_dict()` or the generation context in `run_generate_appeals()`.
2. Add it to `validate_template_fields()` allowed placeholders.
3. Add it to `update_placeholder_list()` so users see it.
4. Update this manual if it is user-facing.

## Adding A New Report Column Requirement

If the app needs another required report column:

1. Add it to `special_cols` if user-selectable.
2. Add a combobox in `create_column_config_tab()`.
3. Populate it in `populate_column_config_tab()`.
4. Save it in `save_config()`.
5. Validate it in `validate_common_inputs()`.
6. Use it through `self.df` rows during generation.

## Changing Generation Rules

Main places to edit:

- `get_generation_row_groups()` for group ordering.
- `_lookup_mapping_row()` for mapping lookup.
- `_parse_selected_method()` for method selection rules.
- `_resolve_required_templates()` for required template discovery.
- `run_generate_appeals()` for orchestration.
- `cleanup_failed_output()` for cleanup policy.
- `copy_mail_outputs_to_mailing_folders()` for Mail handoff behavior.
- `extract_notes_from_successes()` for notes and Fax `Ready.txt`.

Be careful not to mix Mail and Fax downstream workflows. Mail uses `Mailing Folder`; Fax uses day-folder `Ready.txt`.

## Template Naming Rules

Template file names are validated by:

- `build_template_filename()`
- `parse_template_filename()`
- `validate_template_filename_parts()`
- `move_template_to_expected_filename()`

Changing naming rules affects:

- the Templates tab
- `_resolve_required_templates()`
- existing files in users' template folders

## Data Formatting

Excel row values are wrapped by:

```python
format_row_data()
build_formatted_row()
FormattedValue
```

`FormattedValue` preserves raw values for comparisons while formatting values when rendered as strings.

Supported column formats:

- Text
- Date `(MM/DD/YYYY)`
- Monetary `($#,##.##)`
- Percentage `(##.##%)`

## Safe Testing Checklist

After changing code:

1. Run syntax check:

```powershell
python -m py_compile AppealLetterGenerator.py
```

2. Start the app and verify:

- Excel file loads.
- Column Configuration detects columns.
- Mapping Add/Edit opens full form.
- Acronym Add/Edit opens full form.
- Templates tab saves expected template names.
- Auto-detect rationale works with `RationalName`.
- Manual rationale dropdown is enabled when auto-detect is off.

3. Run a small generation test with:

- one Fax row
- one Mail row
- one intentional failure

4. Verify outputs:

- Fax creates letter/fax docs and Fax `Ready.txt` behavior remains correct.
- Mail copies generated files to Mailing Folder.
- Failure workbook goes to Downloads.
- Failed secondary generation removes partial letter/account folder.

## Known Code Notes

- The app is currently a single large file. When refactoring, extract low-risk utility modules first:
  - path/config helpers
  - template management
  - generation workflow
  - note extraction
- Several `_legacy_*` template methods remain for old UI behavior/reference. They are not the primary template flow.
- The app uses direct Tkinter widgets throughout. Avoid changing widget names casually because many methods check `hasattr()`.
- There is no automated test suite in this project today. Use the checklist above before distributing changes.
