Here's a polished `README.md` tailored to your project:

---

```markdown
# 🧪 Menetelmäkoelista Automation Script

This Python script automates the process of adding new welding test methods to a structured Excel file (`DOC410523rev6_Menetelmäkoelista.xlsm`). The system is designed to maintain data integrity across multiple linked tables and sheets, with support for logic links and checkboxes.

---

## ✨ Features

- ✅ Adds a new entry to the **`WPS data`** table (`Table2`)
- ✅ Asks user for inputs column-by-column (empty values allowed)
- ✅ Automatically increments the **"Rivi"** field
- ✅ Adds new row to the selected **material sheet** (e.g. `Väli 1`) in the correct table (`vali1`, `vali2`, ...)
- ✅ Copies modern **Excel checkbox** from previous row in `Column1`
- ✅ Adds a linked row to **logic test sheet** (`logiikkatestit`) in the correct table (`vali1logic`, `vali2logic`, ...)
- ✅ Automatically inserts a new row to avoid overlapping other tables
- ✅ Writes dynamic `XLOOKUP` formula that connects logic and material sheets

---

## 📁 File Structure

```

.
├── DOC410523rev6\_Menetelmäkoelista.xlsm     # Source Excel file (input)
├── muokattu\_menetelmä\_TIMESTAMP.xlsm        # Automatically created copy (output)
├── add.py                                   # Main automation script

````

---

## ▶️ How to Use

1. **Install requirements**

Make sure you have Excel installed (Windows only) and install the required Python package:

```bash
pip install xlwings
````

2. **Run the script**

```bash
python add.py
```

3. **Follow the prompts**
   You'll be asked to:

* Enter the new WPQR number
* Select the material sheet (e.g. Väli 1, Väli 2, ...)
* Provide row-by-row input for WPS data and material sheet
* The rest is automated

---

## 📝 Notes

* The script **copies the original Excel file** to a new one with a timestamp before making changes
* The checkbox in `Column1` must already exist in the previous row
* The new `XLOOKUP` formula dynamically points to the correct material table (e.g., `vali1`, `vali2`)
* Excel **must be in English locale** for formulas to work (`,` instead of `;` and English function names)

---

## 🔒 Requirements

* Python 3.7+
* Microsoft Excel (Windows only)
* `xlwings`

---

## 💡 Roadmap Ideas

* Add GUI for input
* Automate Power Query updates (`Close & Load`)
* Support for rollback or undo
* Integration with certificate expiration notifications

---

## 🧙 Author

Built with care and automation magic by Onni 🪄
