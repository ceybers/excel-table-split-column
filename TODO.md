# TODO List
## Complete
- [x] Run Split procedure from code.
- [x] Run Split procedure from UI.
- [x] Added an easter egg About dialog.
## Halting Errors
- [x] Handle cases where a table only has one row. (`.Value2` will then be a Variant and not an Array of Variants)
- [x] Handle cases where worksheet protection is enabled.
- [x] Handle attempts to split into worksheets with invalid names.
- [x] Handle cases where sheet names are numbers. (Collection object Key property issue)
- [x] Prompt user when no tables are found.
- [x] Handle cases where workbook protection is enabled.
- [x] BUG "Show unsuitable columns" is checked. Select new table. AvailableColumns does not display unsuitable columns until unchecking and rechecking the checkbox.
- [x] BUG 'Elemnt not Found' in TargetSheets.UpdateListView when trying to update a ListView in-situ, after changing to a different AvailableColumn.
  - Only affects after changing to a different Table.
  - String comparisons not working as expected.
- [x] BUG MyDocSettings throws an error on initial run and creating the first settings file.
- [x] BUG Auto suggesting a table from the User's selection (Selection.ListObject) fails if they have a shape selected instead of a cell.
## Available Tables
- [x] Try and choose the table in Selection or on Activesheet by default, instead of first ListObject found in the workbook. 
## Available Columns
- [x] Fix checkbox behaviour of "Available Columns" ListView.
- [x] Option to hide hidden/zero width columns.
- [x] Option to hide unsuitable (non-text) columns.
- [x] Update ListView from AvailableColumns class instead of UserForm Code Behind
- [x] BUG: Changing checkboxes resets selection of column.
- [x] Handle cases where there are no suitable columns in a ListObject.
- [x] Activates/Selects the column in the worksheet when selcting it in the UserForm.
- [ ] Add warning icons to Available Columns with large amount of Unique Values.
## Target Sheets
- [x] ListView correctly clears when switching ListObjects
- [x] ListView correctly clears when switching AvailableColumns
- [x] Handle attempts to split into large amount of worksheets.
  - Hardcoded a limit of 10 Unique values for now.
- [ ] Option to only show target sheet names that are already filtered. (pre-filtered table)
- [ ] ~~Consider adding Target Name search box.~~
- [ ] ~~"Add Current Selection to Filter" option for search box.~~
- [ ] ~~Consider changing Select All/Select None Command Buttons to "(Select All)" List Item in List View. (See: Column Filter dropdown menu)~~
  - VBA ListView does not support tri-state checkboxes.
## Other
- [x] Progress bar dialog and confirmation of completion.
- [x] Fixed modal/modeless bug with Progress bar dialog.
- [x] BUG When opening the tool for the second (or more) time, it will appear in front of the worksheet it was opened on the first run. 
- [ ] Undo feature that will remove the newly created worksheets. (But won't be able to restore the deleted ones)
- [ ] Option to remove the splitting column on the target sheets.
- [ ] Decouple ProgressBar dialog from SplitTable procedure.
## Persistence
- [x] Persistent workbook storage for checkbox preferences.
- [x] BUG Persistent options being incorrectly overwritten between opening the Dialog box and the User receiving control.
- [ ] Persistent workbook storage to repeat/redo the most recent Split operation.