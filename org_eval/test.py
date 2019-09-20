class Process():
    def __init__(self, excel_name='data'):
        self.__all = self.importExcel(excel_name, 0).dropna()
        self.__popRes, self.__popFail = self.populating()
        self.__generateExcelProcess()
        self.fillDimension()
        # print(self.__popRes)
        # print(self.__popFail)

    @staticmethod
    def importExcel(filename, sheetname, usecols=None, header=0, nrows=None):
        import pandas as pd
        return pd.read_excel('resources/{}.xlsx'.format(filename), header=header, sheet_name=sheetname, usecols=usecols,
                             nrows=nrows)

    def populating(self):
        import os
        file = (os.listdir('resources\input'))
        result = []
        resFail = []
        for i in file:
            result.append(self.generateDim(i[:9])) if self.generateDim(i[:9]) else resFail.append(i)
        return result, resFail

    def generateDim(self, file):
        profile = []
        for i in range(len(self.__all.NIM)):
            if file == str(self.__all.NIM[i])[:9]:
                profile.append(self.__all.Nama[i])
                profile.append(str(self.__all.NIM[i])[:9])
                profile.append(self.__all.Jabatan[i])
        from numpy import average as avg
        from openpyxl import load_workbook
        result = []
        try:
            workbook = load_workbook('resources/input/{file}.xlsx'.format(file=file)).sheetnames
            for sheet in workbook:
                nim = self.importExcel('input/{file}'.format(file=file), sheet, "B,G,L,Q", None, 1)
                data = self.importExcel('input/{file}'.format(file=file), sheet, "B,G,L,Q", 1).fillna(0)
                dimension = [data.iloc[0][1],  # Dimension 1
                             data.iloc[1][1],  # Dimension 2
                             avg([avg([data.iloc[2][1], data.iloc[3][1]]), data.iloc[4][1]]),  # Dimension 3
                             avg([avg([data.iloc[5][1], data.iloc[6][1]]), data.iloc[7][1]]),  # Dimension 4
                             avg([avg([data.iloc[8][1], data.iloc[9][1]]),
                                  avg([data.iloc[10][1], data.iloc[11][1], data.iloc[12][1]])]),  # Dimension 5
                             avg([avg([data.iloc[13][1], data.iloc[14][1], data.iloc[15][1]]), data.iloc[16][1]]),
                             # Dimension 6
                             avg([data.iloc[17][1], avg([data.iloc[18][1], data.iloc[19][1], data.iloc[20][1]])]),
                             # Dimension 7
                             avg([avg([data.iloc[21][1], data.iloc[22][1], data.iloc[23][1]]),
                                  avg([data.iloc[24][1], data.iloc[25][1], data.iloc[26][1]]),
                                  avg([data.iloc[27][1], data.iloc[28][1]]),
                                  avg([data.iloc[29][1], data.iloc[30][1], data.iloc[31][1]])]),  # Dimension 8
                             profile  # Add Profile
                             ]
                result.append({nim.iloc[0][6]: dimension})
            return {file: result}
        except Exception as e:
            print(e)
            return None

    def fillDimension(self):
        from openpyxl import load_workbook
        from openpyxl.styles import Protection
        from openpyxl.workbook.protection import WorkbookProtection
        workbook = load_workbook('resources/output/process/Summaries.xlsx')
        for pop in self.__popRes:
            for _ in pop:
                for index in pop[_]:
                    try:
                        for k, v in index.items():
                            init = 3
                            workbook.active = workbook[str(k)]
                            while workbook.active['E{}'.format(str(init))].value != None:
                                init += 1
                            workbook.active.protection.sheet = True
                            workbook.active.protection.password = 'sweet'
                            for x in range(3, 100):
                                workbook.active['H' + str(x)].protection = Protection(locked=False)
                                workbook.active['I' + str(x)].protection = Protection(locked=False)
                                workbook.active['J' + str(x)].protection = Protection(locked=False)
                                workbook.active['K' + str(x)].protection = Protection(locked=False)
                                workbook.active['L' + str(x)].protection = Protection(locked=False)
                                workbook.active['M' + str(x)].protection = Protection(locked=False)
                                workbook.active['N' + str(x)].protection = Protection(locked=False)
                                workbook.active['O' + str(x)].protection = Protection(locked=False)
                            workbook.active['E{}'.format(str(init))].value = v[8][0]
                            workbook.active['F{}'.format(str(init))].value = v[8][1]
                            workbook.active['G{}'.format(str(init))].value = v[8][2]
                            workbook.active['H{}'.format(str(init))].value = v[0]
                            workbook.active['I{}'.format(str(init))].value = v[1]
                            workbook.active['J{}'.format(str(init))].value = v[2]
                            workbook.active['K{}'.format(str(init))].value = v[3]
                            workbook.active['L{}'.format(str(init))].value = v[4]
                            workbook.active['M{}'.format(str(init))].value = v[5]
                            workbook.active['N{}'.format(str(init))].value = v[6]
                            workbook.active['O{}'.format(str(init))].value = v[7]
                    except Exception as e:
                        print(e)
        workbook.security = WorkbookProtection(workbookPassword='sweet', lockWindows=True, lockStructure=True)
        workbook.save("resources/output/process/Summaries.xlsx")
        workbook.close()

    def __generateExcelProcess(self):
        from openpyxl import load_workbook
        workbook = load_workbook('resources/format/process.xlsx')
        for i in range(len(self.__all)):
            workbook.active['B1'].value = self.__all.Nama.values[i]
            workbook.active['B2'].value = self.__all.NIM.values[i]
            workbook.active['B3'].value = self.__all.Jabatan.values[i]
            workbook.active['B4'].value = self.__all.Nama.values[i]
            workbook.active['B5'].value = self.__all.Nama.values[i]
            workbook.active.title = str(self.__all.NIM.values[i])[:9]
            workbook.copy_worksheet(workbook.active)
            workbook.active = len(workbook.sheetnames) - 1
        workbook.remove(workbook.active)
        workbook.save("resources/output/process/Summaries.xlsx")
        workbook.close()
