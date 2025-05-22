# 🧪 Menetelmäkoelista Automation Script

This Python script automates the process of adding new welding test methods to a structured Excel workbook (`.xlsm`). It ensures data consistency across linked sheets, automatically handles checkboxes, inserts logic validation rows, and updates Power Query M-code accordingly.

---

## ✨ Features

- ✅ Prompts the user to select the original Excel file
- ✅ Creates a timestamped copy (e.g., `muokattu_menetelma_2025-05-21_14-30.xlsm`)
- ✅ Adds a new entry to the **`WPS data`** table (`Table2`)
- ✅ Collects user input column-by-column (empty values allowed)
- ✅ Automatically increments the **"Rivi"** field
- ✅ Adds a row to the correct **material table** (e.g., `vali1`, `vali2`, ...) in the selected worksheet
- ✅ Inserts a row into the Excel sheet to prevent overlaps with other tables
- ✅ Copies modern **Excel checkbox** from the previous row into `Column1`
- ✅ Adds a new row to the **logic sheet** (`logiikkatestit`) in the correct table (`vali1logic`, `vali2logic`, ...)
- ✅ Dynamically writes an `XLOOKUP` formula linking logic to material data
- ✅ Directly modifies **Power Query M-code** to add a new condition for the newly added WPQR
- ✅ All logic is handled through a single Python executable without requiring Excel macros

---

## 📁 File Structure

```

.
├── DOC410523rev6\_Menetelmäkoelista.xlsm     # Source Excel file (must be in same folder as the .exe or script)
├── muokattu\_menetelma\_YYYY-MM-DD\_HH-MM.xlsm # Automatically generated copy with timestamp
├── add.py                                   # Main automation script (editable)
├── add.exe                                  # Compiled executable (optional for office-wide use)

````

---

## ▶️ How to Use

### 1. Install Requirements (for Python version)

Ensure Excel is installed (Windows only) and run:

```bash
pip install xlwings
````

### 2. Run the Script

```bash
python add.py
```

Alternatively, launch the compiled `add.exe`.

### 3. Follow the Prompts

* Enter the original Excel file name (`.xlsm`)
* Enter the new WPQR number
* Select the correct material sheet (e.g., Väli 1, Väli 2, ...)
* Fill in fields for WPS data and material table row
* The script will automate everything else: table updates, checkboxes, logic rows, and Power Query edits

---

## 📝 Notes

* The checkbox in `Column1` must already exist in the previous row — the script copies its formatting
* Excel formulas like `XLOOKUP` must use English locale (with `,` not `;`)
* Power Query connections must follow the format: `valiX_results` and `valiX_Parameters`
* All tables must already be present in the Excel file (e.g., `vali1`, `vali1logic`, `vali1_Parameters`, etc.)

---

## 🔒 Requirements

* Python 3.7+ (only for `.py` version)
* Microsoft Excel (Windows only)
* `xlwings`
* Excel file with:

  * Properly named sheets and tables (`valiX`, `valiXlogic`, `valiX_Parameters`)
  * Existing modern Excel checkbox in `Column1`

---

## 🧙 Author

Built with care and automation magic by **Onni** 🪄
Crafted for reliability and scalability in welding documentation workflow.
