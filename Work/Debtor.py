import win32com.client as w32


class Debtor:

    def ConvertDBFtoXLS(self):
        self.Excel = w32.Dispatch("Excel.Application")
        self.wb = self.Excel.Workbooks.Open(u'D:\\DOLG.DBF')
        self.CloseEXCEL()

    def OpenEXCELFile(self):
        # метод открыыает файл Excel с должниками, перекодирует его и возвращает ссылку на объект-книгу


        self.Excel = w32.Dispatch("Excel.Application")

        try:
            self.wb = self.Excel.Workbooks.Open(u'd:\\Боржники.xlsx')
            self.ws = self.wb.ActiveSheet

        except FileNotFoundError:
            print('Файл БОРЖНИКИ не найден')

        return  # self.wb - пока не врубился с возвратом значения из методов

    def ReplaceInSheet(self):
        # wb- книга, открытая методом OpenXLS



        patterns = {u'ИТОГО по дому': u'ВСЬОГО по будинку:',
                    'Карабельная':'Корабельна',
                    'Александрийская':'Олександрійська',
                    'дом':', буд.',
                    'проспект Мира':'просп. Миру',
                    '1 Мая':'1 Травня',
                    'Парковая':'Паркова',
                    'Парусная':'Парусна',
                    'Спортивная':'Спортивна',
                    'ул.':'вул. ',
                    'Виталия':'Віталія',
                    'Данченко':'Данченка',
                    'Торговая':'Торгова',
                    'Победы':'Перемоги',
                    'пер.':'пров. ',
                    'Школьный':'Шкільний',
                    'Шевченко':'Шевченка',
                    'Лазурная':'Лазурна',
                                       }  # и т. д.

        for source_pattern in patterns:
            dest_pattern = patterns.get(source_pattern)
            try:
                self.ws.Columns('A').Replace(source_pattern, dest_pattern, 2, 2, False, True)
            #  False,
            #  False)
            except AttributeError:
               # self.wb.SaveAs(u"Боржники.xls", 51)
                self.wb.Close()
                self.Excel.Quit()
        return


    def CloseEXCEL(self):
        self.wb.SaveAs(u'd:\\Боржники.xlsx', 51)
        self.wb.Close()
        self.Excel.Quit()
        return
