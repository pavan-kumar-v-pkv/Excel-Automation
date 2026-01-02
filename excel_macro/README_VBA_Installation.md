# How to Add the VBA Macro to Your Excel File

## Step 1: Open Your Excel File
- Open `Latha_test.xlsx` (or any Excel file you want to add the macro to).

## Step 2: Open the VBA Editor
- **Windows:** Press `Alt + F11`
- **Mac:** Press `Option + F11` or `Fn + Option + F11`

## Step 3: Insert a New Module
1. In the VBA Editor, click `Insert` → `Module`
2. A new module window will open.

## Step 4: Paste the Macro Code
1. Open the file `GenerateSummary.bas` in VS Code or any text editor.
2. Copy all the code (Ctrl+A / Cmd+A to select all, then Ctrl+C / Cmd+C to copy).
3. Go back to Excel's VBA Editor and paste it into the module window (Ctrl+V / Cmd+V).

## Step 5: Save as Macro-Enabled Workbook
1. Close the VBA Editor.
2. Click `File` → `Save As`
3. **IMPORTANT:** A dialog may appear saying "The following features cannot be saved in macro-free workbooks"
   - Click **"No"**
   - Then choose file type: **Excel Macro-Enabled Workbook (.xlsm)** from the dropdown
4. Save the file.

## Step 6: Add a Button to Run the Macro (Optional but Recommended)

### Option A: Quick Access Button
1. Go to `View` → `Macros` → `View Macros`
2. Select `GenerateSummary`
3. Click `Run` to test it.

### Option B: Add a Button to the Sheet
1. Go to `Developer` tab (if not visible, enable it in Excel Options → Customize Ribbon)
2. Click `Insert` → `Button (Form Control)`
3. Draw the button on your sheet
4. In the dialog, select `GenerateSummary` macro
5. Click OK
6. Right-click the button and choose `Edit Text` to rename it to "Generate Summary"

### Option C: Add to Quick Access Toolbar
1. Click the down arrow on the Quick Access Toolbar (top left)
2. Choose `More Commands`
3. Select `Macros` from the dropdown
4. Select `GenerateSummary`
5. Click `Add`
6. Click OK

## Step 7: Test the Macro
1. Make sure your Excel file has the expected sheets (Master Bathroom, Kids Bathroom, etc.)
2. Make sure each sheet has a row with "TOTAL VALUE" in column C
3. Click the button or run the macro from the Macros dialog
4. The SUMMARY sheet should be created/updated with all values!

## Troubleshooting
- **Macro security:** If macros are disabled, go to `File` → `Options` → `Trust Center` → `Trust Center Settings` → `Macro Settings` and choose "Enable all macros" (or "Disable with notification").
- **Formulas not calculated:** Go to `Formulas` → `Calculation Options` → `Automatic`
- **External links warning:** Click "Update" to allow formulas to calculate.

## Distribution to End Users
1. Save your file as `.xlsm` with the macro already embedded.
2. Add the button to the sheet so users just click it.
3. Share the `.xlsm` file with users.
4. Provide them with this instruction guide.

---

**That's it! Users can now generate the summary with a single button click.**
