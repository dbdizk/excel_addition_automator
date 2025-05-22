import shutil
import xlwings as xw

# 1. Tee kopio tiedostosta
original_file = "DOC410523rev6_Menetelmäkoelista.xlsm"
copy_file = "muokattu_menetelma.xlsm"
shutil.copyfile(original_file, copy_file)

# 2. Avaa Excel-työkirja (piilossa)
app = xw.App(visible=False)
wb = xw.Book(copy_file)

# 3. Pyydä käyttäjältä WPQR ja valitse materiaalivälilehti
wpqr = input("Anna WPQR-numero (esim. WPQR316): ")

materiaalit = [s.name for s in wb.sheets if "väli" in s.name.lower() or "cs" in s.name.lower()]
print("\nMateriaalivälilehdet:")
for i, nimi in enumerate(materiaalit):
    print(f"{i+1}. {nimi}")
valinta = int(input("Valitse materiaalivälilehti numerolla: ")) - 1
materiaali_sheet = wb.sheets[materiaalit[valinta]]

# 4. Lisää rivi "WPS data" -taulukkoon Table2
wps_sheet = wb.sheets["WPS data"]
header_range = wps_sheet.range("A1").expand("right")
otsikot = header_range.value

print("\nSyötä tiedot 'WPS data' -taulukkoon (Table2):")
wps_rivi = []
for otsikko in otsikot:
    arvo = input(f"{otsikko}: ")
    wps_rivi.append(arvo if arvo != "" else None)

wps_sheet.range("A1").end("down").offset(1, 0).value = wps_rivi

# 5. Lisää rivi taulukkoon valiX (esim. vali1, vali2, ...)
vali_sheet = materiaali_sheet  # valittiin jo aiemmin
vali_table_name = "vali" + str(valinta + 1)
vali_table = next((tbl for tbl in vali_sheet.api.ListObjects if tbl.Name == vali_table_name), None)
if not vali_table:
    raise ValueError(f"Taulukkoa nimeltä '{vali_table_name}' ei löytynyt välilehdeltä '{vali_sheet.name}'.")

headers = [h.Name for h in vali_table.ListColumns]

print(f"\nSyötä arvot taulukkoon '{vali_table_name}' sarakkeittain:")

# Selvitetään taulukon alue ja nykyiset rivit
table_range = vali_sheet.range(vali_table.Range.Address).options(expand='table')
data = table_range.value
row_count = len(data) - 1  # ensimmäinen rivi on header

# Määritetään uusi rivin paikka
new_row_num = table_range.row + row_count

# Etsi sarakeindeksi Column1:lle
column1_index = headers.index("Column1")

# Rakennetaan uusi rivi
new_row = []
for col in headers:
    if col.lower() == "rivi":
        current_numbers = [r[headers.index(col)] for r in data[1:] if isinstance(r[headers.index(col)], (int, float))]
        seuraava_rivi = int(max(current_numbers) + 1) if current_numbers else 1
        print(f"{col}: {seuraava_rivi} (luotu automaattisesti)")
        new_row.append(seuraava_rivi)
    elif col.lower() == "column1":
        new_row.append(None)  # checkbox lisätään erikseen
    else:
        arvo = input(f"{col}: ")
        new_row.append(arvo if arvo != "" else None)

# Etsi viimeinen rivi taulukon alueelta
last_table_row = table_range.last_cell.row
first_table_col = table_range.column

# Kirjoita uusi rivi soluun, joka on taulukon viimeisen rivin alapuolella
vali_sheet.range((last_table_row + 1, first_table_col)).value = new_row


# Kopioi edellisen rivin checkbox-solu ilman viitettä
source_cell = vali_sheet.cells(new_row_num - 1, table_range.column + column1_index)
target_cell = vali_sheet.cells(new_row_num, table_range.column + column1_index)

source_cell.copy()
target_cell.paste(paste='formats')  # Tämä säilyttää checkboxin ilman linkkiä

print(f"✅ Rivi lisätty taulukkoon {vali_table_name} ja checkbox kopioitu.")

# 5.2 Lisää rivi logiikkatestit-välilehdelle taulukkoon valiXlogic
logic_sheet = wb.sheets["logiikkatestit"]
logic_table_name = f"{vali_table_name}logic"

# Etsi logiikkataulukko nimellä
logic_table = next((tbl for tbl in logic_sheet.api.ListObjects if tbl.Name == logic_table_name), None)
if not logic_table:
    raise ValueError(f"Logiikkataulukkoa nimeltä '{logic_table_name}' ei löytynyt.")

# Selvitä viimeinen rivi logiikkataulukossa
logic_range = logic_sheet.range(logic_table.Range.Address).options(expand='table')
last_logic_row = logic_range.last_cell.row
first_logic_col = logic_range.column

# INSERT tyhjä rivi ennen lisäystä
insert_row = last_logic_row + 1
logic_sheet.range(f"{insert_row}:{insert_row}").insert(shift="down")

# Luo kaava (englanninkielinen Excel-syntaksi!)
xlookup_formula = f'=XLOOKUP(OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())), 0, -1), {vali_table_name}[Menetelmäkoenumero], {vali_table_name}[Column1], "")'

# Kirjoita WPQR ja kaava uuteen riviin
logic_sheet.range((insert_row, first_logic_col)).value = [wpqr, xlookup_formula]

print(f"✅ Rivi lisätty logiikkatestit-taulukkoon '{logic_table_name}' ja siirretty seuraavat taulukot alaspäin.")




# 6. Tallenna ja sulje
wb.save()
wb.close()
app.quit()
print(f"\n✅ Muutokset tallennettu tiedostoon: {copy_file}")
