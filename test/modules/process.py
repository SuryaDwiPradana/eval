# from common import make_dir

def make_sheet_recap(config):
    sheet = config.load_sheet(config.config['resources_path']+'/format/process.xlsx')
    active_sheet_idx = 0
    miss_nim = []
    import os.path
    for cat in config.data:
        for _ in config.data[cat]:
            for nim in _:
                if not os.path.exists(config.config['resources_path']+"/input/"+nim+".xlsx"):
                    miss_nim.append(nim)
                sheet.copy_worksheet(sheet.active)
                active_sheet_idx += 1
                sheet.active = active_sheet_idx
                sheet.active.title = nim
                sheet.active["B1"].value = config.data[cat][0][nim][0]['nama']
                sheet.active["B2"].value = nim
                sheet.active["B3"].value = config.data[cat][0][nim][0]['jabatan']
                sheet.active = 0
    return sheet, miss_nim

def main_process(config):
    miss_nim = []
    try:
        import os.path
        file_recap, miss_nim = make_sheet_recap(config)
        if len(miss_nim) > 0:
            print(f'There is {len(miss_nim)} missing file(s): \n{miss_nim}')
            cont = input('Continue? (Yes/No) default: No\nAnswer: ').lower() or 'No'
            if cont == 'yes':
                from tqdm import tqdm
                tqdm.write("Processing Files...")
                # for cat in tqdm(config.data,ascii=True,desc="Result", position=1):
                #     for y in config.data[cat]:
                #         for nim in tqdm(y,ascii=True,desc=cat, position=0):
                #             if os.path.exists(config.config['resources_path']+"/input/"+nim+".xlsx"):
                #                 file_ref = config.load_sheet(config.config['resources_path']+"/input/"+nim+".xlsx", 1)
                #             else:
                #                 continue
                #             for row in tqdm(range(5,file_ref.active.max_row,7),desc=nim,ascii=True,position=2):
                #                 if file_ref.active["C"+str(row)].value != nim and file_ref.active["C"+str(row)].value != None:
                #                     file_recap.active = file_recap[file_ref.active["C"+str(row)].value]
                #                     init = 3
                #                     while file_recap.active["E"+str(init)].value != None:
                #                         init+=1
                #                     file_recap.active["E"+str(init)].value = config.data[cat][0][nim][0]['nama']
                #                     file_recap.active["F"+str(init)].value = nim
                #                     file_recap.active["G"+str(init)].value = config.data[cat][0][nim][0]['jabatan']
                #                     file_recap.active["H"+str(init)].value = file_ref.active["F"+str(row+5)].value
                #                     file_recap.active["I"+str(init)].value = file_ref.active["G"+str(row+5)].value
                #                     file_recap.active["J"+str(init)].value = file_ref.active["H"+str(row+5)].value
                #                     file_recap.active["K"+str(init)].value = file_ref.active["K"+str(row+5)].value
                #                     file_recap.active["L"+str(init)].value = file_ref.active["N"+str(row+5)].value
                #                     file_recap.active["M"+str(init)].value = file_ref.active["S"+str(row+5)].value
                #                     file_recap.active["N"+str(init)].value = file_ref.active["W"+str(row+5)].value
                #                     file_recap.active["O"+str(init)].value = file_ref.active["AA"+str(row+5)].value
                file_recap.active = 0
                file_recap[file_recap.active.title].sheet_state = "hidden"
                file_recap.save(config.config['resources_path']+'/output/'+config.folder['dirname'][1]+"/Evaluation Summaries.xlsx")
    except Exception as e:
        print(e)
    finally:
        print('Happy Evaluation')
    return