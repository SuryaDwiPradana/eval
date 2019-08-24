class Configuration():
    def __init__(self):
        import os
        self.config = {}
        self.folder = {}
        self.report = {}
        self.config['root_path'] = os.getcwd()
        self.config['resources_path'] = 'resources'

        self.folder['dirname'] = ['prepare','process','report']
        # 
        self.data = {}
        self._load_data("datalist.txt")
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