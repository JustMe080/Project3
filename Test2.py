import pandas as pd
import win32com.client
import numpy as np
import xlwings as xw
import shutil
import datetime

current_time = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S").replace(":", "-")

data1 = pd.read_excel('Addresses.xlsx', sheet_name='Sheet1')

lot = np.array([2, 4, 5, 6, 8, 10, 26, 27, 28], dtype=int)
File = f'Receipt_{current_time}.xlsx'

original_file = 'template.xlsx'

last_number = 1000

def normalize_text(text):
    if isinstance(text, str):
        return text.strip().lower() 
    return text

def fittext(cell):
    max_width = cell.column_width * 7.5
    while cell.api.Font.Size > 8:
        text_width = len(str(cell.value)) * cell.api.Font.Size * 0.6
        if text_width <= max_width:
            break
        cell.api.Font.Size -=1

# Start a single Excel instance to prevent multiple processes
app = xw.App(visible=False)

try:
    for a in lot:
        copy_file = f"C:/project3/lot no{a}.xlsx"
        shutil.copy(original_file, copy_file)

        file_path = f"C:/project3/lot no{a}.xlsx"
        wb = app.books.open(file_path)  # Open within the same instance

        for i in range(len(data1)):  
            original_region = str(data1.loc[i, 'Region']).strip()
            original_division = str(data1.loc[i, 'Division']).strip()
            address = str(data1.loc[i, 'Address']).strip()

            region = normalize_text(original_region)
            division = normalize_text(original_division)

            data2 = pd.read_excel('quantity_bases1.xlsx', sheet_name=original_region, header=[2, 3])  
            data2.columns = ['_'.join(map(str, col)).strip() for col in data2.columns] 

            data2_normalized = data2.copy()
            data2_normalized.columns = [normalize_text(col) for col in data2.columns]

            data2_temp = pd.read_excel('quantity_bases1.xlsx', sheet_name=original_region, header=3)  
            data2_temp_header = [normalize_text(col) for col in data2_temp.columns]

            data2 = data2.drop(columns=["Unnamed: 0_level_0_Unnamed: 0_level_1"], errors="ignore") 
            data2.iloc[:, 2:6] = data2.iloc[:, 2:6].apply(pd.to_numeric, errors="coerce").fillna(0).astype(int)

            counting = ""
            for j in data2_temp_header:
                if division == j:  
                    if region == 'ncr':
                        counting = f"NATIONAL CAPITAL REGION ({original_division})"
                    elif region == 'car':
                        counting = f"CORDILLERA ADMINISTRATIVE REGION_{original_division}"
                    elif region == 'region iv-b':
                        counting = f"MIMAROPA_{original_division}"
                    else:
                        counting = f"{original_region.upper()}_{original_division}"

            if not counting:
                continue  

            counting_normalized = normalize_text(counting)
            if counting_normalized in data2_normalized.columns:
                counting = data2.columns[data2_normalized.columns.get_loc(counting_normalized)]  

            data2 = data2[~data2.apply(lambda row: row.astype(str).str.contains('total', case=False).any(), axis=1)]
            Selected_data = data2[data2[counting] > 1].reset_index(drop=True)

            for k in range(len(Selected_data)):
                lot_no = str(Selected_data.loc[k]['LOT NO._Unnamed: 1_level_1']).strip()

                if int(lot_no) == a:
                    data3 = pd.read_excel('Lot_Items.xlsx', sheet_name=lot_no)
                    data3_normalized = data3.copy()
                    data3_normalized.columns = [normalize_text(col) for col in data3.columns]

                    source_sheet = wb.sheets["Sheet1"] 
                    source_sheet.api.Copy(After=wb.sheets[-1].api)

                    sheet_name = f"{original_region}_{original_division}_{lot_no}"[:30]
                    counter = 1

                    new_sheet = wb.sheets[-1]
                    while sheet_name in [sh.name for sh in wb.sheets]:
                        truncated_base = sheet_name[:27]
                        sheet_name = f"{truncated_base}_{counter}"
                        counter += 1

                    last_number += 1
                    new_invoice = f"2025-03-{last_number:04}"

                    new_sheet.range("D7").value = new_invoice
                    new_sheet.name = sheet_name
                    new_sheet.range("B15").options(transpose=True).value = data3.iloc[:, 0].tolist() 
                    new_sheet.range("C15").options(transpose=True).value = data3.iloc[:, 1].tolist()
                    new_sheet.range("B14").value = "Lot No. " + str(Selected_data.loc[k]['LOT NO._Unnamed: 1_level_1']) + " " + Selected_data.loc[k]['LOT / QUALIFICATION_Unnamed: 2_level_1']
                    new_sheet.range("B8").value = f"TESDA {original_region} - {original_division}"
                    new_sheet.range("B9").value = address
                    new_sheet.range("C15").options(transpose=True).value = data3.iloc[:, 1].tolist()

                    start_row = 15
                    while new_sheet.range(f"B{start_row}").value:  
                        start_row += 1  

                    for idx in range(15, start_row):
                        new_sheet.range(f"A{idx}").value = Selected_data.loc[k][counting]  

                    first_value = new_sheet.range("A15").value
                    for idx in range(16, start_row):
                        if new_sheet.range(f"B{idx}").value:
                            new_sheet.range(f"A{idx}").value = first_value

                    fittext(new_sheet.range("B14"))

                    quantities = (Selected_data.loc[k][counting] * data3.iloc[:, 1]).tolist() 
                    total = sum(quantities)
                    new_sheet.range("D15").options(transpose=True).value = quantities
                    new_sheet.range("D29").value = total

                    new_sheet.range("C:C").number_format = "#,##0.00"
                    new_sheet.range("D:D").number_format = "#,##0.00"

                    wb.app.calculate()
                    print(f"✅ Sheet {sheet_name} updated successfully!")

        wb.save(file_path)
        wb.close()  # Close file properly after saving

finally:
    app.quit()  # Ensure Excel process closes

print("✅ All sheets updated successfully!")
