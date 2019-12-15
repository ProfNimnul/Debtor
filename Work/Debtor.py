import win32com.client as w32
from os import path as p
from os import remove
from easygui import msgbox, fileopenbox


class Debtor:
    codes = {1: "площа Праці",
             10: "Сквер Перемоги",
             100: "Тимчасова вулиця",
             11: "Садова",
             12: "1 Травня",
             13: "Зелена",
             14: "Південна",
             16: "Паркова",
             17: "Спортивна",
             18: "Парусна",
             19: "Віталія Шума",
             2: "проспект Миру"
             20: "Данченко",
             3: "Олександрійська",
             31: "Лазурна",
             4: "Корабельна",
             5: "Хантадзе",
             6: "Шевченко",
             7: "Шкільний",
             8: "Торгова",
             9: "вул. Перемоги",
             99999: "Загальноміські об'єкти"
             }
    streets = {
        'Карабельная': 'Корабельна',
        'Александрийская': 'Олександрійська',
        'Олександр.': 'Олександрійська',
        'дом': ', буд.',
        'проспект Мира': 'пр-т Миру',
        'просп.': 'пр-т ',
        'Ленiна': 'пр-т. Миру',
        '1 Мая': '1 Травня',
        '1Травня': '1 Травня',
        'Парковєая': 'Паркова',
        'Парусная': 'Парусна',
        'Г.Сталiнграду ': 'Парусна',
        'Спортивная': 'Спортивна',
        'ул.': u'Вул. ',
        'Виталия': 'В.',
        'Віталия': 'В.',
        'Данченко': 'Данченка',
        'Торговая': 'Торгова',
        'Победы': 'Перемоги',
        'пер.': 'пров. ',
        'Школьный': 'Шкільний',
        'Шевченко': 'Шевченка',
        'Лазурная': 'Лазурна',
        'площадь': 'площа',
        'Труда': 'Праці',
        '\\': '/'

    }  # и т. д
    ## Блок глобвльных данных

    # *****************************************************************
    # метод открыыает файл Excel с должниками, перекодирует его и возвращает ссылку на объект-книгу
    def openExcelInstance(self, fname):
        try:

            self.Excel = w32.DispatchEx("Excel.Application")
            self.wb = self.Excel.Workbooks.Open(fname)
            self.ws = self.wb.ActiveSheet

        except FileNotFoundError as  E:
            print('Файл с долгами не найден')
            print(E)

            self.wb.Close(False)
            self.Excel.Application.Quit()
            sys.exit(-1)
        else:
            return self.Excel, self.wb, self.ws

    # *****************************************************************

    def get_file_name(self):
        """ Окно выбора файла и его возворат"""
        fname = ''
        fname = fileopenbox("Выберите файл", default="*.xlsx;*.dbf")

        filename, file_extension = p.splitext(fname)

        while len(fname) == 0:
            if (file_extension.upper() not in ('.XLSX', '.DBF')):
                msgbox('Это не XLSX и не DBF- файл', ok_button="ОК", title="Перевірте тип файла!")
                fname = ''
                # exit ( )
        return fname

    # *****************************************************************




    def ReplaceInSheet(self, Excel, wb, ws, dictionary:dict, invert:bool =  False, range:str):

        def inv_dict(d: dict) -> dict  # инвертирует словарь
            d = {v: k for k, v in d.items()}
            return d
        # wb- книга, открытая методом OpenXLS

        if invert:
            dictionary = inv_dict(dictionary)

        column = range.split(':')[0]

        for key in dictionary:
            # dest_code = codes.get(source_code)
            value = dictionary[key]
            try:
                try:
                    founded = ws.Range(range).Find(key)
                    ## print(founded)
                    if founded != None:
                        ws.Columns(column).Replace(key, value, 2, 2, False, True)
                except Exception as E:
                    print(E)
            except AttributeError:
                pass

        # wb.Close ( )
        # Excel.Quit ( )

        return

    # *****************************************************************

    def SaveAndClose(self, path, Excel, wb, ws, filetype="XLSX"):
        # ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF FileName:="sales.pdf" Quality:=xlQualityStandard DisplayFileAfterPublish:=True

        fullpath = path + u'Боржники_ред.xlsx'
        try:
            if p.exists(fullpath):
                remove(fullpath)
        except OSError as E:
            print(E)
        finally:
            if filetype == "XLSX":
                wb.SaveAs(fullpath, 51)
            elif filetype == "PDF":
                fullpath = path + u'Боржники_ред.pdf'
                ws.ExportAsFixedFormat(0, fullpath, 0, 0)

            wb.Close()
            Excel.Quit()
        return

    # *****************************************************************

    # *****************************************************

    def addHeader(self, ws):
        import datetime
        current_month = datetime.datetime.today().month

        month_name = {1: "січня",
                      2: "лютого",
                      3: "березня",
                      4: "квітня",
                      5: "травня",
                      6: "червня",
                      7: "липня",
                      8: "серпня",
                      9: "вересня",
                      10: "жовтня",
                      11: "листопада",
                      12: "грудня"
                      }[current_month]

        header_city = "м. Чорноморськ"
        header_with_month = "Перелік боржників за послуги з утримання будинків та прибудинкових теріторій" \
                            "  станом на 01 {} 2017 р. ".format(month_name)

        header_warning = "Усім боржникам потрібно терминово погасити існуючу заборгованість" \
                         " для подальшого відповідного надання послуг" \
                         " з утримання будинків та прибудинкових територій!"

        ws.Range("2:9").Insert(-4142)
        # for _ in range(5):
        #     # посмотреть и изучить количество аргументов и порядок их передачи!!!
        #     ws.Selection.Insert (Shift = -4121, CopyOrigin = 0)

        ws.Range("A1:F5").ClearContents()

        # ws.Selection.HorizontalAlignment = -4108 #xlCenter
        # ws.Selection.VerticalAlignment = -4160 #xlTop
        # ws.Selection.WrapText = -1
        # ws.Selection.Orientation = 0
        # ws.Selection.AddIndent = 0
        # ws.Selection.IndentLevel = 0
        # ws.Selection.ShrinkToFit = 0
        # ws.Selection.ReadingOrder = -5002 #xlContext
        # ws.Selection.MergeCells = -1
        # ws.Selection.Font.Size = 20
        # ws.Selection.Font.Bold = -1
        # ws.Selection.Font.Color = -16776961

        # ws.Selection.Merge()
        ws.Range("A1:F1").Merge()
        ws.Range('A1').HorizontalAlignment = -4108  # xlCenter
        ws.Range('A1').VerticalAlignment = -4160  # xlTop

        ws.Range("A1").Value = header_city
        ws.Range('A1').Font.Bold = -1

        ws.Range('A3:F4').Merge()
        ws.Range('A3').HorizontalAlignment = -4108  # xlCenter
        ws.Range('A3').VerticalAlignment = -4160  # xlTop
        ws.Range('A3').ShrinkToFit = 0
        ws.Range('A3').WrapText = -1
        ws.Range('A3').Font.Bold = -1
        ws.Range("A3").Value = header_with_month

        ws.Range('A5:F7').Merge()
        ws.Range('A5').HorizontalAlignment = -4108  # xlCenter
        ws.Range('A5').VerticalAlignment = -4160  # xlTop
        ws.Range('A5').WrapText = -1
        ws.Range('A5').Orientation = 0
        ws.Range('A5').AddIndent = 0
        ws.Range('A5').IndentLevel = 0
        ws.Range('A5').ShrinkToFit = -1
        ws.Range('A5').ReadingOrder = -5002  # xlContext
        ws.Range('A5').MergeCells = -1
        ws.Range('A5').Font.Size = 16
        ws.Range('A5').Font.Bold = -1
        ws.Range('A5').Font.Color = -16776961
        ws.Range("A5").RowHeight = 58.5

        ws.Range("A5").Value = header_warning

        ws.Range('A10').Font.Bold = -1
        ws.Range('B10').Font.Bold = -1
        ws.Range('C10').Font.Bold = -1

        ws.Range('A10').Value = "Вул., буд."
        ws.Range('B10').Value = "кв."
        ws.Range('C10').Value = "Борг"
        ## Все константі перевести в числовой вид!!!
    # *****************************************************************


if __name__ == '__main__':
    debtor = Debtor()

    fname = debtor.get_file_name()

    Excel, wb, ws = debtor.openExcelInstance(fname)

    debtor.addHeader(ws)
    debtor.ReplaceInSheet(Excel, wb, ws)

    path = p.dirname(fname)

    debtor.SaveAndClose(path, Excel, wb, ws, "PDF")
