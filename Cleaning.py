import pandas as pd
import win32com.client
import numpy as np
import xlwings as xw


input_file = "C:/Project3/Price Schedule Form - TESDA-CO-2024-09.xlsx"
output_file = "C:/Project3/Lot_Items.xlsx"

sheet_names = pd.ExcelFile(input_file).sheet_names

lots = np.array([2, 4, 5, 6, 8, 10, 26, 27, 28])

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    for i, lot in enumerate(lots):
        df = pd.read_excel(input_file, sheet_name=sheet_names[i], header=[7])

        df = df.iloc[:, [1, -8]].dropna()

        df.to_excel(writer, sheet_name=f"{lot}", index=False)

print("✅ New Excel file 'Lot_Items.xlsx' created with sheets for each lot!")

