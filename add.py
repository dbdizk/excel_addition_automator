import os
import shutil
import xlwings as xw
from datetime import datetime

# Figure path to the original file
if getattr(sys, 'frozen', False):
    base_dir = os.path.dirname(sys.executable)
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))

# Ask original file name
original_name = input("Anna alkuperäisen Excel-tiedoston nimi (esim. DOC410523rev6_Menetelmäkoelista.xlsm): ").strip()

# Add .xlsm extension if not present
if not original_name.lower().endswith(".xlsm"):
    original_name += ".xlsm"

original_file = os.path.join(base_dir, original_name)

# Create a copy of the original file with a timestamp
timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
copy_file = os.path.join(base_dir, f"muokattu_menetelma_{timestamp}.xlsm")

# Ensure that the original file exists
if not os.path.exists(original_file):
    raise FileNotFoundError(f"Tiedostoa ei löytynyt: {original_file}")

# Make copy of the original file
shutil.copyfile(original_file, copy_file)


# Try delete old file if not locked anymore
try:
    if os.path.exists(copy_file):
        os.remove(copy_file)
except PermissionError:
    raise RuntimeError("Tiedosto 'muokattu_menetelma.xlsm' on edelleen auki. Sulje Excel ja yritä uudelleen.")

shutil.copyfile(original_file, copy_file)


# Open the copied file in background
app = xw.App(visible=False)
wb = xw.Book(copy_file)

# Read WPQR and chosen material sheet
wpqr = input("Anna WPQR-numero (esim. WPQR316): ")
materiaalit = [s.name for s in wb.sheets if "väli" in s.name.lower() or "cs" in s.name.lower()]
print("\nMateriaalivälilehdet:")
for i, nimi in enumerate(materiaalit):
    print(f"{i+1}. {nimi}")
valinta = int(input("Valitse materiaalivälilehti numerolla: ")) - 1
materiaali_sheet = wb.sheets[materiaalit[valinta]]

# Add WPS data row
wps_sheet = wb.sheets["WPS data"]
otsikot = wps_sheet.range("A1").expand("right").value
print("\nSyötä tiedot 'WPS data' -taulukkoon (Table2):")
wps_rivi = [input(f"{otsikko}: ") or None for otsikko in otsikot]
wps_sheet.range("A1").end("down").offset(1, 0).value = wps_rivi

# Add row to the right sheet table
vali_table_name = "vali" + str(valinta + 1)
vali_table = next((tbl for tbl in materiaali_sheet.api.ListObjects if tbl.Name == vali_table_name), None)
if not vali_table:
    raise ValueError(f"Taulukkoa '{vali_table_name}' ei löytynyt.")
headers = [h.Name for h in vali_table.ListColumns]
table_range = materiaali_sheet.range(vali_table.Range.Address).options(expand='table')
data = table_range.value
new_row_num = table_range.row + len(data) - 1
column1_index = headers.index("Column1")

new_row = []
for col in headers:
    if col.lower() == "rivi":
        nums = [r[headers.index(col)] for r in data[1:] if isinstance(r[headers.index(col)], (int, float))]
        seuraava_rivi = int(max(nums) + 1) if nums else 1
        print(f"{col}: {seuraava_rivi} (luotu automaattisesti)")
        new_row.append(seuraava_rivi)
    elif col.lower() == "column1":
        new_row.append(None)
    else:
        arvo = input(f"{col}: ")
        new_row.append(arvo if arvo != "" else None)

first_col = table_range.column
insert_row = table_range.last_cell.row + 1
materiaali_sheet.range(f"{insert_row}:{insert_row}").insert(shift="down")
materiaali_sheet.range((insert_row, first_col)).value = new_row


# Copy checkbox format
source = materiaali_sheet.cells(new_row_num - 1, table_range.column + column1_index)
target = materiaali_sheet.cells(new_row_num, table_range.column + column1_index)
source.copy()
target.paste(paste='formats')
print(f"✅ Rivi lisätty taulukkoon {vali_table_name} ja checkbox kopioitu.")

# Add Excel logic test row
logic_sheet = wb.sheets["logiikkatestit"]
logic_table_name = f"{vali_table_name}logic"
logic_table = next((tbl for tbl in logic_sheet.api.ListObjects if tbl.Name == logic_table_name), None)
logic_range = logic_sheet.range(logic_table.Range.Address).options(expand='table')
insert_row = logic_range.last_cell.row + 1
logic_sheet.range(f"{insert_row}:{insert_row}").insert(shift="down")

xlookup = f'=XLOOKUP(OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())), 0, -1), {vali_table_name}[Menetelmäkoenumero], {vali_table_name}[Column1], "")'
logic_sheet.range((insert_row, logic_range.column)).value = [wpqr, xlookup]
print(f"✅ Rivi lisätty logiikkatestit-taulukkoon '{logic_table_name}'.")

# Add Power Query condition through xlwings and query api
query_name = f"{vali_table_name}_result"
parameters_name = f"{vali_table_name}_Parameters"
parameters_index = seuraava_rivi - 1

target_query = None
for q in wb.api.Queries:
    if q.Name == query_name:
        target_query = q
        break

if not target_query:
    raise ValueError(f"Queryä '{query_name}' ei löytynyt Power Querysta.")

old_formula = target_query.Formula
new_clause = f'or (if {parameters_name}[Column2]{{{parameters_index}}} then ([WPQR] = "{wpqr}") else null)'

if new_clause in old_formula:
    print("⚠️ Ehto on jo olemassa Power Queryssa.")
else:
    insert_pos = old_formula.rfind("))")
    if insert_pos == -1:
        raise ValueError("Power Queryn rakenne ei ole odotetussa muodossa: ei löydy '))'.")
    new_formula = old_formula[:insert_pos] + f"\n    {new_clause}" + old_formula[insert_pos:]
    target_query.Formula = new_formula
    print("✅ Power Query -ehto lisätty.")

# Save and close the workbook
wb.save()
wb.close()
app.quit()
print(f"\n✅ Muutokset tallennettu tiedostoon: {copy_file}")
