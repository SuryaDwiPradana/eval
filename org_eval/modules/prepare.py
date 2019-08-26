from common import make_dir 

def make_dist_excel(config):
    from tqdm import tqdm 
    sheet = config.load_sheet(config.config['resources_path']+'/format/prepare.xlsx')
    for _ in tqdm(sheet.sheetnames,ascii=True,desc="Hidden"):
        if _ != "0 Indikator" and _ != "1 BPH":
            sheet[_].sheet_state = "hidden"
    sheet.active = 1
    # sheet.active = 0
    # while sheet.active.title != "1 BPH":
    #     for x in range(0,len(sheet.sheetnames)):
    #         sheet.active = x
    #         if sheet.active.title == "1 BPH":
    #             break
    init = 5
    for cat in tqdm(config.data,ascii=True,desc="Result", position=1):
        for nim in tqdm(config.data[cat],ascii=True,desc=cat, position=0):    
            if sheet.active.title == "1 BPH" and cat == "BPH" and sheet.active["B"+str(init)].value == None:
                    sheet.active["B"+str(init)].value = config.data[cat][nim]['name']
                    sheet.active["C"+str(init)].value = nim
                    sheet.active["D"+str(init)].value = config.data[cat][nim]['position']
                    init+=7
            elif sheet.active["B"+str(init)].value == None and config.data[cat][nim]['position'][:5] == 'Ketua':
                    sheet.active["B"+str(init)].value = config.data[cat][nim]['name']
                    sheet.active["C"+str(init)].value = nim
                    sheet.active["D"+str(init)].value = config.data[cat][nim]['position']
                    init+=7
    sheet.security = impl_protect(sheet)
    # sheet.security.lockStructure = True
    sheet.save(config.config['resources_path']+'/output/'+config.folder['dirname'][0]+"/Evaluation.xlsx")

def impl_protect(sheet):
    from openpyxl.workbook.protection import WorkbookProtection
    sheet.security = WorkbookProtection(workbookPassword='000', lockWindows=True, lockStructure=True)
    return sheet.security

# def excelketua(data,path,dirname,komisi):
#     wb = opensheet('format/Format_1.xlsx')
#     for sheet in wb.sheetnames:
#         if sheet != "2 Ketua Komisi" and sheet != "0 Indikator":
#             wb[sheet].sheet_state = "hidden"
#     wb.active = 0
#     while wb.active.title != "2 Ketua Komisi":
#         for x in range(0,len(wb.sheetnames)):
#             wb.active = x
#             if wb.active.title == "2 Ketua Komisi":
#                 break
#     init = 5
#     for x in data:
#         for y in data[x]:
#             for z in y:    
#                 if wb.active.title == "2 Ketua Komisi" and x == "BPH":
#                     if wb.active["B"+str(init)].value == None:
#                         wb.active["B"+str(init)].value = data[x][0][z][0]['nama']
#                         wb.active["C"+str(init)].value = z
#                         wb.active["D"+str(init)].value = data[x][0][z][0]['jabatan']
#                         init+=7
#                 elif wb.active.title == "2 Ketua Komisi" and x == komisi:
#                     if wb.active["B"+str(init)].value == None:
#                         wb.active["B"+str(init)].value = data[x][0][z][0]['nama']
#                         wb.active["C"+str(init)].value = z
#                         wb.active["D"+str(init)].value = data[x][0][z][0]['jabatan']
#                         init+=7
#                 else:
#                     if wb.active["B"+str(init)].value == None and data[x][0][z][0]['jabatan'][:5] == 'Ketua':
#                         wb.active["B"+str(init)].value = data[x][0][z][0]['nama']
#                         wb.active["C"+str(init)].value = z
#                         wb.active["D"+str(init)].value = data[x][0][z][0]['jabatan']
#                         init+=7
#     wb.save(path+'/'+dirname+"/Evaluasi Ketua Komisi "+komisi[-1:]+".xlsx")
#     return

# def excelanggota(data,path,dirname,komisi):
#     wb = opensheet('format/Format_1.xlsx')
#     for sheet in wb.sheetnames:
#         if sheet != "3 Anggota Komisi" and sheet != "0 Indikator":
#             wb[sheet].sheet_state = "hidden"
#     wb.active = 0
#     while wb.active.title != "3 Anggota Komisi":
#         for x in range(0,len(wb.sheetnames)):
#             wb.active = x
#             if wb.active.title == "3 Anggota Komisi":
#                 break;
#     init = 5
#     for x in data:
#         for y in data[x]:
#             for z in y:
#                 if wb.active.title == "3 Anggota Komisi" and x == komisi:
#                     if wb.active["B"+str(init)].value == None:
#                         wb.active["B"+str(init)].value = data[x][0][z][0]['nama']
#                         wb.active["C"+str(init)].value = z
#                         wb.active["D"+str(init)].value = data[x][0][z][0]['jabatan']
#                         init+=7
#     wb.save(path+'/'+dirname+"/Evaluasi Anggota Komisi "+komisi[-1:]+".xlsx")
#     return
    
def main_prepare(config):
    make_dist_excel(config)
    # print("Pembuatan File Evaluasi(Excel)")
    # print("1. Semua")
    # print("2. BPH")
    # print("3. Komisi A")
    # print("4. Komisi B")
    # print("5. Komisi C")
    # print("6. Komisi D")
    # print("7. Ketua Komisi")
    # print("8. Anggota Komisi")
    # print("9. Exit")
    # pilihan = int(input("Silahkan Pilih: ") or "9")
    # if pilihan == 1:
    # 	print("Tolong tunggu sebentar...")
        # excelbph(data,path,dirname)
        # excelketua(data,path,dirname,"KOM A")
        # excelanggota(data,path,dirname,"KOM A")
        # excelketua(data,path,dirname,"KOM B")
        # excelanggota(data,path,dirname,"KOM B")
        # excelketua(data,path,dirname,"KOM C")
        # excelanggota(data,path,dirname,"KOM C")
        # excelketua(data,path,dirname,"KOM D")
        # excelanggota(data,path,dirname,"KOM D")
    # elif pilihan == 2:
    # 	print("Tolong tunggu sebentar...")
        # excelbph(data,path,dirname)
    # elif pilihan == 3:
        # print("Tolong tunggu sebentar...")
        # excelketua(data,path,dirname,"KOM A")
        # excelanggota(data,path,dirname,"KOM A")
    # elif pilihan == 4:
        # print("Tolong tunggu sebentar...")
        # excelketua(data,path,dirname,"KOM B")
        # excelanggota(data,path,dirname,"KOM B")
    # elif pilihan == 5:
        # print("Tolong tunggu sebentar...")
        # excelketua(data,path,dirname,"KOM C")
        # excelanggota(data,path,dirname,"KOM C")
    # elif pilihan == 6:
        # print("Tolong tunggu sebentar...")
        # excelketua(data,path,dirname,"KOM D")
        # excelanggota(data,path,dirname,"KOM D")
    # elif pilihan == 7:
        # print("Tolong tunggu sebentar...")
        # excelketua(data,path,dirname,"KOM A")
        # excelketua(data,path,dirname,"KOM B")
        # excelketua(data,path,dirname,"KOM C")
        # excelketua(data,path,dirname,"KOM D")
    # elif pilihan == 8:
        # print("Tolong tunggu sebentar...")
        # excelanggota(data,path,dirname,"KOM A")
        # excelanggota(data,path,dirname,"KOM B")
        # excelanggota(data,path,dirname,"KOM C")
        # excelanggota(data,path,dirname,"KOM D")
    # print("Selamat Evaluasi")
    return
