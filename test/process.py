def load_sheet(xlsx, data = False):
    import openpyxl
    load = openpyxl.load_workbook(xlsx, data_only = True) if data else openpyxl.load_workbook(xlsx)
    return load

def load_data():
    import json
    with open("datalist.txt") as f:
        data = json.load(f)
    return data

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