def load_data():
    workbook, _ = get_workbook_for_month()
    sheet = workbook.active
    #completer les donnees de employee  
    get_data_base()
    for row_number, (matricule, (nom, prenom, salaire_base, categories)) in enumerate(matricules.items(), start=2):
        sheet.cell(row=row_number, column=1, value=matricule)
        sheet.cell(row=row_number, column=2, value=nom)
        sheet.cell(row=row_number,column=3,value=prenom)
        sheet.cell(row=row_number,column=4, value=salaire_base)
        sheet.cell(row=row_number,column=5, value=categories)
    workbook.save(filename)
    for row in sheet.iter_rows(min_row=2, values_only=True):
        tree.insert('', END, values=row)
    workbook.close()