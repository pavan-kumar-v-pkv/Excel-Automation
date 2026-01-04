# How to Use Kohler Image Tool

## 📷 Insert Product Images into Excel

### First Time (Windows - Fully Automatic):

1. **Double-click** `Kohler Image Tool.bat`
2. **Wait for automatic setup:**
   - Python downloads and installs (2 minutes)
   - Required packages install (1 minute)
   - GUI opens automatically
3. **You're ready to use it!**

### First Time (Mac):

1. **Double-click** `Kohler Image Tool.command`
2. **If Python not installed:**
   - Tool will open python.org for you
   - Download and install Python 3.11
   - Run the tool again
3. **Automatic Setup:**
   - Packages install automatically
   - GUI opens
4. **You're ready!**

### Every Time You Use:

1. **Double-click the launcher:**
   - **Windows:** `Kohler Image Tool.bat`
   - **Mac:** `Kohler Image Tool.command`

2. **Fill in the form**:
   - Click **Browse** → Select your Excel quotation file
   - Click **Browse** → Select your Kohler PDF catalog
   - Choose if you want a new file or update the original

3. **Click "START IMAGE PROCESSING"**

4. **Wait** for the success message (1-3 minutes)

5. **Done!** Your Excel file now has product images

---

## 📊 Create Summary Sheet (Excel Macro)

### One-Time Setup:

1. Open your Excel quotation file
2. Press `Alt+F11` (opens VBA Editor)
3. In the menu: **File → Import File...**
4. Select **GenerateSummary.bas** from the Kohler Image Tool folder
5. Save and close VBA Editor

### Every Time You Use:

1. Open your Excel file
2. Press `Alt+F8` (opens Macros)
3. Select **"GenerateSummary"**
4. Click **Run**
5. A new "Summary" sheet appears with all totals!

---

## Troubleshooting

**Tool won't open?**
- **Windows:** Make sure you're double-clicking the .bat file, not the .command file
- **Mac:** Make sure you're double-clicking the .command file, not the .bat file
- Install Python from python.org first

**Excel file not working?**
- Convert .xlsb files to .xlsx first (File → Save As)

**Macro not working?**
- Enable macros in Excel (File → Options → Trust Center)

---

That's it! Just use the tool whenever you need images in your Excel files.
