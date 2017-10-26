import win32com.client as w32
from os import path as p
from os import remove
from easygui import msgbox, fileopenbox

class Debtor:
    ## Блок глобвльных данных


    # *****************************************************************
        # метод открыыает файл Excel с должниками, перекодирует его и возвращает ссылку на объект-книгу
    def openExcelInstance(self,fname):
        try:
            self.Excel = w32.DispatchEx("Excel.Application")
            self.wb = self.Excel.Workbooks.Open ( fname )
            self.ws = self.wb.ActiveSheet

        except FileNotFoundError as  E:
            print ( 'Файл с долгами не найден' )
            print ( E )

            self.wb.Close(False)
            self.Excel.Application.Quit()
            sys.exit(-1)
        else:
            return self.Excel,self.wb,self.ws



    #*****************************************************************


    def get_xls_file_name(self):
        """ Окно выбора файла и его возворат"""
        fname = ''
        fname = fileopenbox("Выберите файл",default = "*.xlsx")
        while len(fname) == 0:
            if not fname.endswith(".xlsx"):
                msgbox('Это не XLSx- файл', ok_button="ОК", title="Перевірте тип файла!")
                fname = ''
                # exit ( )
        return fname

    #*****************************************************************

    def ReplaceInSheet ( self,Excel, wb, ws ):
        # wb- книга, открытая методом OpenXLS
        
        patterns = {u'ИТОГО по дому': u'ВСЬОГО по будинку:',
                    'Карабельная': 'Корабельна',
                    'Александрийская': 'Олександрійська',
                    'YL_DOM': u'Вул., буд.',
                    'дом': ', буд.',
                    'проспект Мира': 'просп. Миру',
                    '1 Мая': '1 Травня',
                    'Парковєая': 'Паркова',
                    'Парусная': 'Парусна',
                    'Спортивная': 'Спортивна',
                    'ул.': u'Вул. ',
                    'Виталия': 'Віталія',
                    'Данченко': 'Данченка',
                    'Торговая': 'Торгова',
                    'Победы': 'Перемоги',
                    'пер.': 'пров. ',
                    'Школьный': 'Шкільний',
                    'Шевченко': 'Шевченка',
                    'Лазурная': 'Лазурна',
                    'N_KV': '№ кв.',
                    'SUM_D': 'Борг',
                    'ВВул': 'Вул'
                    }  # и т. д




        for source_pattern in patterns:
            # dest_pattern = patterns.get(source_pattern)
            dest_pattern = patterns[ source_pattern ]
            try:
                try:
                    founded = ws.Range('A:A').Find(source_pattern)
                    ## print(founded)
                    if founded != None:
                        print("Исх. шаблон ",source_pattern )
                        print("Рез. шаблон ",dest_pattern )

                        ws.Columns('A').Replace(source_pattern, dest_pattern, 2, 2, False, True)

                except Exception as E:
                    print(E)
            except AttributeError:
                pass

           # wb.Close ( )
           # Excel.Quit ( )

        return
    #*****************************************************************

    def SaveAndClose (self, path, Excel, wb):
        fullpath = path+u'Боржники_ред.xlsx'
        try:
            if p.exists(fullpath):
                remove(fullpath)
        except OSError as E:
            print(E)
        finally:
            wb.SaveAs(fullpath , 51 )
            wb.Close()
            Excel.Quit()
        return
    #*****************************************************************

#*****************************************************


    def addHeader(self, ws):
        import datetime
        current_month = datetime.datetime.today().month

        month_name = {1:"січня",
                      2:"лютого",
                      3:"березня",
                      4:"квітня",
                      5:"травня",
                      6:"червня",
                      7:"липня",
                      8:"серпня",
                      9:"вересня",
                      10:"жовтня",
                      11:"листопада",
                      12:"грудня"
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

        #ws.Selection.Merge()
        ws.Range("A1:F1").Merge()
        ws.Range('A1').HorizontalAlignment = -4108 #xlCenter
        ws.Range('A1').VerticalAlignment = -4160 #xlTop

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
        ws.Range('A5').HorizontalAlignment = -4108 #xlCenter
        ws.Range('A5').VerticalAlignment = -4160 #xlTop
        ws.Range('A5').WrapText = -1
        ws.Range('A5').Orientation = 0
        ws.Range('A5').AddIndent = 0
        ws.Range('A5').IndentLevel = 0
        ws.Range('A5').ShrinkToFit = -1
        ws.Range('A5').ReadingOrder = -5002 #xlContext
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
    #*****************************************************************


if __name__ == '__main__':


        debtor = Debtor ()

        fname = debtor.get_xls_file_name()


        Excel,wb,ws = debtor.openExcelInstance(fname)

        debtor.addHeader(ws)
        debtor.ReplaceInSheet (Excel,wb,ws )



        path = p.dirname(fname)

        debtor.SaveAndClose (path, Excel, wb)
