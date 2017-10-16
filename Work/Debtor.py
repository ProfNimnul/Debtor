import win32com.client as w32
import sys
from os import path as p
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

        except FileNotFoundError:
            print ( 'Файл с долгами не найден' )
            self.wb.Close(False)
            self.Excel.Application.Quit()
            sys.exit(-1)

    #*****************************************************************


    def get_xls_file_name(self):
        """ Окно выбора файла и его возворат"""
        fname = ''
        while len(fname) == 0:
            if not fname.endswith(".xls"):
                msgbox('Это не XLS- файл', ok_button="ОК", title="Перевірте тип файла!")
                fname = ''
                # exit ( )
        return fname

    #*****************************************************************

    def ReplaceInSheet ( self ):
        # wb- книга, открытая методом OpenXLS

        patterns = {u'ИТОГО по дому': u'ВСЬОГО по будинку:',
                    'Карабельная': 'Корабельна',
                    'Александрийская': 'Олександрійська',
                    'дом': ', буд.',
                    'проспект Мира': 'просп. Миру',
                    '1 Мая': '1 Травня',
                    'Парковая': 'Паркова',
                    'Парусная': 'Парусна',
                    'Спортивная': 'Спортивна',
                    'ул.': 'вул. ',
                    'Виталия': 'Віталія',
                    'Данченко': 'Данченка',
                    'Торговая': 'Торгова',
                    'Победы': 'Перемоги',
                    'пер.': 'пров. ',
                    'Школьный': 'Шкільний',
                    'Шевченко': 'Шевченка',
                    'Лазурная': 'Лазурна',
                    }  # и т. д




        for source_pattern in patterns:
            # dest_pattern = patterns.get(source_pattern)
            dest_pattern = patterns[ source_pattern ]
            try:
                self.ws.Columns ( 'A' ).Replace ( source_pattern , dest_pattern , 2 , 2 , False , True )
            # False,
            #  False)
            except AttributeError:
                # self.wb.SaveAs(u"Боржники.xls", 51)
                self.wb.Close ( )
                self.Excel.Quit ( )

        return
    #*****************************************************************
    def SaveAndClose (self, path):
        self.wb.SaveAs(path+u'Боржники.xlsx' , 51 )
        self.wb.Close ( )
        self.Excel.Quit ( )
        return
    #*****************************************************************

    def addWarning(self):
        pass
    #*****************************************************************


    def addHeader(self):
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
        header_with_month = "Перелік боржників за послуги з утримання будинків \n та прибудинкових теріторій \n" \
                            "  станом на 01 {} 2017 р. ".format(month_name)

        header_warning = "Усім боржникам потрібно терминово погасити існуючу заборгованість \n" \
                         " для подальшого відповідного надання послуг \n" \
                         " з утримання будинків та прибудинкових територій!"


        self.ws.Range("1:1").Select()
        for _ in range(5):
            # посмотреть и изучить количество аргументов и порядок их передачи!!!
            self.ws.Selection.Insert (Shift = -4121, CopyOrigin = 0)

        self.ws.Range("A1:E5").Select()
        self.ws.Selection.ClearContents()

        self.ws.Selection.HorizontalAlignment = -4108 #xlCenter
        self.ws.Selection.VerticalAlignment = -4160 #xlTop
        self.ws.Selection.WrapText = -1
        self.ws.Selection.Orientation = 0
        self.ws.Selection.AddIndent = 0
        self.ws.Selection.IndentLevel = 0
        self.ws.Selection.ShrinkToFit = 0
        self.ws.Selection.ReadingOrder = -5002 #xlContext
        self.ws.Selection.MergeCells = -1
        self.ws.Selection.Font.Size = 20
        self.ws.Selection.Font.Bold = -1
        self.ws.Selection.Font.Color = -16776961

        #self.ws.Selection.Merge()
        self.ws.Range("C1").Value = header_city
        self.ws.Range("B3").Value = header_warning
        self.ws.Range("D3").Value = header_with_month        ## Все константі перевести в числовой вид!!!
    #*****************************************************************


if __name__ == '__main__':

    fname = debtor.get_xls_file_name()
    if fname != "":

        debtor = Debtor ( fname )
        debtor.openExcelInstance(fname)
        debtor.ReplaceInSheet ( )
        debtor.addHeader()
        path = p.dirname(fname)
        debtor.SaveAndClose (path)
