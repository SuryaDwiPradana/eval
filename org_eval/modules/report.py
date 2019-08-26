from common import make_dir

def fill_identity(wb_res,name,nim,org,position,period):
    wb_res.active["B4"] = f": {name}"
    wb_res.active["B5"] = f": {nim}"
    wb_res.active["B6"] = f": {org}"
    wb_res.active["B7"] = f": {position}"
    wb_res.active["B8"] = f": {period}"
    if org.upper().startswith("BPMU") or org.upper().startswith("SMU"):
        wb_res.active["G31"] = f"Ketua Komisi Organisasi {org}"
        wb_res.active["C31"] = f"Ketua Umum {org} UKSW"
    elif org.upper().startswith("BPMF") or org.upper().startswith("SMF"):
        wb_res.active["G31"] = f"Ketua Komisi Organisasi {org}"
        wb_res.active["C31"] = f"Ketua {org} UKSW"

def fill_score(wb_res,wb_ref):
    wb_res.active["F12"] = avgcell(wb_ref,"H")
    wb_res.active["F13"] = avgcell(wb_ref,"I")
    wb_res.active["F14"] = avgcell(wb_ref,"J")
    wb_res.active["F15"] = avgcell(wb_ref,"K")
    wb_res.active["F16"] = avgcell(wb_ref,"L")
    wb_res.active["F17"] = avgcell(wb_ref,"M")
    wb_res.active["F18"] = avgcell(wb_ref,"N")
    wb_res.active["F19"] = avgcell(wb_ref,"O")
    wb_res.active["C22"] = (avgcell(wb_ref,"H")+avgcell(wb_ref,"I")+ \
        avgcell(wb_ref,"J")+avgcell(wb_ref,"K")+avgcell(wb_ref,"L")+ \
            avgcell(wb_ref,"M")+avgcell(wb_ref,"N")+avgcell(wb_ref,"O"))/8

def fill_report(wb_res,wb_ref,name,nim,org,position,period):
    fill_identity(wb_res,name,nim,org,position,period)
    fill_score(wb_res,wb_ref)
    
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

def main_report(config):
    try:
        from tqdm import tqdm
        report_idx = 0
        config.load_report()
        tqdm.write("Loading Files...")
        file_summary = "Evaluation Summaries"
        file_ref = config.load_sheet(config.config['resources_path']+'/output/'+config.folder['dirname'][1]+'/'+file_summary+'.xlsx', True)
        all_recap = config.load_sheet(config.config['resources_path']+'/format/report.xlsx')
        solo_recap = config.load_sheet(config.config['resources_path']+'/format/report.xlsx')
        tqdm.write("Make Report Files...")
        for cat in tqdm(config.data,ascii=True,desc="Result", position=1):
            make_dir(cat,config.config['resources_path']+'/output/'+config.folder['dirname'][2])
            for nim in tqdm(config.data[cat],ascii=True,desc=cat, position=0):
                file_ref.active = file_ref[nim]
                all_recap.copy_worksheet(all_recap.active)
                report_idx += 1
                all_recap.active = report_idx
                all_recap.active.title = nim[:9]

                fill_report(all_recap,file_ref[nim],config.data[cat][nim]['name'],nim,config.report['organization'],config.data[cat][nim]['position'],config.report['period'])
                fill_report(solo_recap,file_ref[nim],config.data[cat][nim]['name'],nim,config.report['organization'],config.data[cat][nim]['position'],config.report['period'])
                
                solo_recap.save(config.config['resources_path']+'/output/'+config.folder['dirname'][2]+"/"+cat+'/'+nim+".xlsx")
                all_recap.active = 0
        all_recap[all_recap.active.title].sheet_state = "hidden"
        all_recap.save(config.config['resources_path']+'/output/'+config.folder['dirname'][2]+"/Evaluation Report Recap.xlsx")
        tqdm.write("Close All Files...")
    except Exception as e:
        print(e)
    finally:
        input('\nHappy Evaluation\nBest Regard RSS\n')
    return
