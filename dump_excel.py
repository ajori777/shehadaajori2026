import openpyxl

def dump_sheet(file_path):
    wb = openpyxl.load_workbook(file_path, data_only=True)
    sheet = wb.active
    print(f"Dumping {file_path}")
    for r in range(1, 15):
        row = []
        for c in range(1, 30):
            val = sheet.cell(row=r, column=c).value
            row.append(str(val) if val is not None else "")
        print("\t".join(row))

dump_sheet('12.xlsx')
