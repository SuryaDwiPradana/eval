from process import load_data, make_dir, load_sheet

def make_recap_sheet():
    data = load_data()
    sheet = load_sheet('format/Format_2.xlsx')
    active_sheet_idx = 0
    for cat in data:
        for _ in data[cat]:
            for nim in _:
                sheet.copy_worksheet(sheet.active)
                active_sheet_idx += 1
                sheet.active = active_sheet_idx
                sheet.active.title = nim
                sheet.active["B1"].value = data[cat][0][nim][0]['nama']
                sheet.active["B2"].value = nim
                sheet.active["B3"].value = data[cat][0][nim][0]['jabatan']
                sheet.active = 0
    return sheet

def main_process():
    import os
    import shutil
    data = load_data()
    dirname = "Process"
    path = os.getcwd()
    make_dir(dirname,path)
    print("Tolong tunggu sebentar...")
    try:
        wb_recap = make_recap_sheet()
        for cat in data:
            for y in data[cat]:
                for nim in y:
                    wb = load_sheet("Data Evaluasi/"+nim+".xlsx", 1)
                    for row in range(5,wb.active.max_row,7):
                        if wb.active["C"+str(row)].value != nim and wb.active["C"+str(row)].value != None:
                            wb_recap.active = wb_recap[wb.active["C"+str(row)].value]
                            init = 3
                            while wb_recap.active["E"+str(init)].value != None:
                                init+=1
                            wb_recap.active["E"+str(init)].value = data[cat][0][nim][0]['nama']
                            wb_recap.active["F"+str(init)].value = nim
                            wb_recap.active["G"+str(init)].value = data[cat][0][nim][0]['jabatan']
                            wb_recap.active["H"+str(init)].value = wb.active["F"+str(row+5)].value
                            wb_recap.active["I"+str(init)].value = wb.active["G"+str(row+5)].value
                            wb_recap.active["J"+str(init)].value = wb.active["H"+str(row+5)].value
                            wb_recap.active["K"+str(init)].value = wb.active["K"+str(row+5)].value
                            wb_recap.active["L"+str(init)].value = wb.active["N"+str(row+5)].value
                            wb_recap.active["M"+str(init)].value = wb.active["S"+str(row+5)].value
                            wb_recap.active["N"+str(init)].value = wb.active["W"+str(row+5)].value
                            wb_recap.active["O"+str(init)].value = wb.active["AA"+str(row+5)].value
    except Exception as e:
        print(e)
    wb_recap.active = 0
    wb_recap[wb_recap.active.title].sheet_state = "hidden"
    wb_recap.save(path+'/'+dirname+"/Rekap Evaluasi.xlsx")
    return

# def main_process():
#     data = load_data()
#     dirname = "Process"
#     path = os.getcwd()
#     make_dir(dirname,path)

#     print("Tolong tunggu sebentar...")
#     try:
#         wb_recap = load_sheet('format/Format_2.xlsx')
#         counter = 0
#         for x in data:
#             print('x ='+x)
#             for y in data[x]:
#                 for z in y:
#                     print('z ='+z)
#                     wb_recap.copy_worksheet(wb_recap.active)
#                     counter +=1
#                     wb_recap.active = counter
#                     wb_recap.active.title = z
#                     wb_recap.active["B1"].value = data[x][0][z][0]['nama']
#                     wb_recap.active["B2"].value = z
#                     wb_recap.active["B3"].value = data[x][0][z][0]['jabatan']
#                     wb_recap.active = 0
#             initTemp = 5
#             for x in data:
#                 for y in data[x]:
#                     for z in y:
#                         wb = load_sheet("Data Evaluasi/"+z+".xlsx", 1)
#                         for row in range(initTemp,wb.active.max_row,7):
#                             if wb.active["C"+str(row)].value != z and wb.active["C"+str(row)].value != None:
#                                 wb_recap.active = wb_recap[wb.active["C"+str(row)].value]
#                                 init = 3
#                                 while wb_recap.active["E"+str(init)].value != None:
#                                     init+=1
#                                 wb_recap.active["E"+str(init)].value = data[x][0][z][0]['nama']
#                                 wb_recap.active["F"+str(init)].value = z
#                                 wb_recap.active["G"+str(init)].value = data[x][0][z][0]['jabatan']
#                                 wb_recap.active["H"+str(init)].value = wb.active["F"+str(row+5)].value
#                                 wb_recap.active["I"+str(init)].value = wb.active["G"+str(row+5)].value
#                                 wb_recap.active["J"+str(init)].value = wb.active["H"+str(row+5)].value
#                                 wb_recap.active["K"+str(init)].value = wb.active["K"+str(row+5)].value
#                                 wb_recap.active["L"+str(init)].value = wb.active["N"+str(row+5)].value
#                                 wb_recap.active["M"+str(init)].value = wb.active["S"+str(row+5)].value
#                                 wb_recap.active["N"+str(init)].value = wb.active["W"+str(row+5)].value
#                                 wb_recap.active["O"+str(init)].value = wb.active["AA"+str(row+5)].value
#     except Exception as e:
#         print(e)
#     wb_recap.active = 0
#     wb_recap[wb_recap.active.title].sheet_state = "hidden"
#     wb_recap.save(path+'/'+dirname+"/Rekap Evaluasi.xlsx")
#     return

main_process()