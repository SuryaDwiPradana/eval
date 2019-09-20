class Process():
    def __init__(self, excel_name='data'):
        self.__all = self.importExcel(excel_name, 0).dropna()
        self.xxx()
        self.__dictList = {}
        # self.__generateExcelProcess()

    @staticmethod
    def importExcel(filename, sheetname,usecols=None,header=0,nrows=None):
        import pandas as pd
        return pd.read_excel('resources/{}.xlsx'.format(filename), header = header, sheet_name = sheetname, usecols=usecols,nrows=nrows)

    def generateDim(self):
        from numpy import average as avg
        nim = self.importExcel('input/462015039',0,"B,G,L,Q",None,1)
        all = self.importExcel('input/462015039',0,"B,G,L,Q",1).fillna(0)
        result = [all.iloc[0][1],
        all.iloc[1][1],
        avg([avg([all.iloc[2][1],all.iloc[3][1]]),all.iloc[4][1]]),
        avg([avg([all.iloc[5][1],all.iloc[6][1]]),all.iloc[7][1]]),
        avg([avg([all.iloc[8][1],all.iloc[9][1]]),avg([all.iloc[10][1],all.iloc[11][1],all.iloc[12][1]])]),
        avg([avg([all.iloc[13][1],all.iloc[14][1],all.iloc[15][1]]),all.iloc[16][1]]),
        avg([all.iloc[17][1],avg([all.iloc[18][1],all.iloc[19][1],all.iloc[20][1]])]),
        avg([avg([all.iloc[21][1],all.iloc[22][1],all.iloc[23][1]]),avg([all.iloc[24][1],all.iloc[25][1],all.iloc[26][1]]),avg([all.iloc[27][1],all.iloc[28][1]]),avg([all.iloc[29][1],all.iloc[30][1],all.iloc[31][1]])])
        ]
        return {nim.iloc[0][6]:result}
    

    def __generateExcelProcess(self):
        from openpyxl import load_workbook
        from openpyxl.styles import Protection
        from openpyxl.workbook.protection import WorkbookProtection
        workbook = load_workbook('resources/format/process.xlsx')
        for i in range(len(self.__all)):
            workbook.active.protection.sheet = True
            workbook.active.protection.password = 'sweet'
            for x in range(3,100):
                workbook.active['H'+str(x)].protection = Protection(locked=False)
                workbook.active['I'+str(x)].protection = Protection(locked=False)
                workbook.active['J'+str(x)].protection = Protection(locked=False)
                workbook.active['K'+str(x)].protection = Protection(locked=False)
                workbook.active['L'+str(x)].protection = Protection(locked=False)
                workbook.active['M'+str(x)].protection = Protection(locked=False)
                workbook.active['N'+str(x)].protection = Protection(locked=False)
                workbook.active['O'+str(x)].protection = Protection(locked=False)
            workbook.active['B1'].value = self.__all.Nama.values[i]
            workbook.active['B2'].value = self.__all.NIM.values[i]
            workbook.active['B3'].value = self.__all.Jabatan.values[i]
            workbook.active['B4'].value = self.__all.Nama.values[i]
            workbook.active['B5'].value = self.__all.Nama.values[i]
            workbook.active.title = str(self.__all.NIM.values[i])
            workbook.copy_worksheet(workbook.active)
            workbook.active = len(workbook.sheetnames)-1
        workbook.remove(workbook.active)
        workbook.security = WorkbookProtection(workbookPassword='sweet', lockWindows=True, lockStructure=True)
        workbook.save("resources/output/process/Summaries Evaluation.xlsx")
        workbook.close()

x = Process()