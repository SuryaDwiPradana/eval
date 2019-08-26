class Configuration():
    def __init__(self):
        import os
        self.config = {}
        self.folder = {}
        self.report = {}
        self.config['root_path'] = os.getcwd()
        self.config['resources_path'] = os.path.join('org_eval','resources')

        self.folder['dirname'] = ['prepare','process','report']
        # 
        self.data = {}
        # self._load_data("datalist.txt")
        self._load_data_new()
    def _load_data_new(self):
        import openpyxl
        file = openpyxl.load_workbook(self.config['resources_path']+"\\"+'data.xlsx')
        data = {}
        file.active = 0
        if file.active.title.lower() == 'data':
            for _ in range(2, file.active.max_row+1):
                name = file.active[f"A{_}"].value
                nim = file.active[f"B{_}"].value
                position = file.active[f"C{_}"].value
                cat = file.active[f"D{_}"].value
                if None in (name,nim,position,cat):
                    continue
                else:
                    name = name.strip()
                    nim = str(nim).strip()
                    position = position.strip()
                    cat = cat.strip()
                if not cat in data:
                    data[cat] = {}
                if not nim in data[cat]:
                    data[cat][nim] = {}
                if not name in data[cat][nim]:
                    data[cat][nim] = {'name':name,'position':position}
        file.close()
        self.data = data
    def _load_data(self, name):     
        try:
            import json
            with open(self.config['resources_path']+"\\"+name) as f:
                self.data = json.load(f)
        except Exception as e:
            print(e)   
    def load_sheet(self, xlsx, data = False):
        import openpyxl
        load = openpyxl.load_workbook(xlsx, data_only = True) if data else openpyxl.load_workbook(xlsx)
        return load
    def load_report(self):
        self.report['period'] = '2018/2019'
        self.report['organization'] = 'BPMU'

def make_dir(dirname, path):
    TESTDIR = dirname
    try:
        import os
        home = os.path.expanduser(path)

        if not os.path.exists(os.path.join(home, TESTDIR)):  
            os.makedirs(os.path.join(home, TESTDIR))
    except Exception as e:
        print(e)
    return
