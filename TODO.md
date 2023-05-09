# TODO List
## Complete
- [x] Run Split procedure from code.
- [x] Run Split procedure from UI.
## Halting Errors
- [ ] Prompt user when no tables are found.
- [ ] Handle cases where a table only has one row. (`.Value2` will then be a Variant and not an Array of Variants)
- [ ] Handle cases where workbook protection is enabled.
- [ ] Handle cases where worksheet protection is enabled.
- [ ] Handle attempts to split into worksheets with invalid names.
- [ ] Handle cases where sheet names are numbers. (Collection object Key property issue)
## Available Columns
- [ ] Fix checkbox behaviour of "Available Columns" ListView.
- [ ] Add warning icons to Available Columns with large amount of Unique Values.
- [ ] Option to hide hidden/zero width columns.
## Target Sheets
- [ ] Handle attempts to split into large amount of worksheets.
- [ ] Option to only show target sheet names that are already filtered. (pre-filtered table)
- [ ] Consider adding Target Name search box.
- [ ] "Add Current Selection to Filter" option for search box.
- [ ] Consider changing Select All/Select None Command Buttons to "(Select All)" List Item in List View. (See: Column Filter dropdown menu)
## Other
- [ ] Progress bar dialog and confirmation of completion.
- [ ] Undo feature that will remove the newly created worksheets. (But won't be able to restore the deleted ones)
- [ ] Option to remove the splitting column on the target sheets.
## Persistence
- [ ] Persistent workbook storage to repeat/redo the most recent Split operation.
- [ ] Persistent workbook storage for checkbox preferences.