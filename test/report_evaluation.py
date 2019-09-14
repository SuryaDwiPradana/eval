from process import make_dir, load_data, load_sheet

def fillidentity(work,name,nim,org,pos,th):
    work.active["B4"] = ": "+ name
    work.active["B5"] = ": "+ nim
    work.active["B6"] = ": "+ org
    work.active["B7"] = ": "+ pos
    work.active["B8"] = ": "+ th
    return

def fillscore(workset,workget):
    workset.active["F12"] = avgcell(workget,"H")
    workset.active["F13"] = avgcell(workget,"I")
    workset.active["F14"] = avgcell(workget,"J")
    workset.active["F15"] = avgcell(workget,"K")
    workset.active["F16"] = avgcell(workget,"L")
    workset.active["F17"] = avgcell(workget,"M")
    workset.active["F18"] = avgcell(workget,"N")
    workset.active["F19"] = avgcell(workget,"O")
    workset.active["C22"] = (avgcell(workget,"H")+avgcell(workget,"I")+avgcell(workget,"J")+avgcell(workget,"K")+avgcell(workget,"L")+avgcell(workget,"M")+avgcell(workget,"N")+avgcell(workget,"O"))/8
    return

def avgcell(work,letter):
    count = 0
    xavg = 0
    xsum = 0
    for x in range(3,work.max_row):
        if work[letter+str(x)].value != None:
            count +=1
    if count > 0:
        for x in range(3,count+3):
            xsum += work[letter+str(x)].value    
        xavg = xsum/count
    return xavg

def report_main():
    import os
    import shutil
    data = load_data()
    dirname = "Report"
    path = os.getcwd()
    make_dir(dirname,path)
    print("Tolong tunggu sebentar...")

    wb_recap = load_sheet('format/Format_3.xlsx')
    wb_solo = load_sheet('format/Format_3.xlsx')
    rekapfolder = "Process"
    rekapfile = "Rekap Evaluasi"
    organisasi = "BPMU"
    tahun = "2018/2019"
    wb = load_sheet(rekapfolder+'/'+rekapfile+'.xlsx', True)
    init = 0
    for cat in data:
        make_dir(cat,path+'/'+dirname)
        for y in data[cat]:
            for nim in y:
                wb.active = wb[nim]
                wb_recap.active = init
                wb_recap.copy_worksheet(wb_recap.active)
                wb_recap.active.title = nim[:9]
                # fillidentity(wb_recap,data[cat][0][nim][0]['nama'],nim,organisasi,data[cat][0][nim][0]['jabatan'],tahun)
                # fillscore(wb_recap,wb[nim])
                
                # fillidentity(wb_solo,data[cat][0][nim][0]['nama'],nim,organisasi,data[cat][0][nim][0]['jabatan'],tahun)
                # fillscore(wb_solo,wb[nim])
                wb_solo.save(path+'/'+dirname+"/"+cat+'/'+nim+".xlsx")
                init += 1       
    wb_recap.save(path+'/'+dirname+"/Recap Laporan Evaluasi.xlsx")
    return

report_main()
