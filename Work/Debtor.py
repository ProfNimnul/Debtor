import win32com.client as w32
import sys
from os import path as p

class Debtor:
    def __init__ ( self,fname):

        # метод открыыает файл Excel с должниками, перекодирует его и возвращает ссылку на объект-книгу


        self.Excel = w32.Dispatch ( "Excel.Application" )

        try:
            self.wb = self.Excel.Workbooks.Open ( fname )
            self.ws = self.wb.ActiveSheet

        except FileNotFoundError:
            print ( 'Файл с долгами не найден' )

    def ReplaceInSheet ( self ):
        # wb- книга, открытая методом OpenXLS

        patterns = {u'ИТОГО по дому': u'ВСЬОГО по будинку:' ,
                    'Карабельная': 'Корабельна' ,
                    'Александрийская': 'Олександрійська' ,
                    'дом': ', буд.' ,
                    'проспект Мира': 'просп. Миру' ,
                    '1 Мая': '1 Травня' ,
                    'Парковая': 'Паркова' ,
                    'Парусная': 'Парусна' ,
                    'Спортивная': 'Спортивна' ,
                    'ул.': 'вул. ' ,
                    'Виталия': 'Віталія' ,
                    'Данченко': 'Данченка' ,
                    'Торговая': 'Торгова' ,
                    'Победы': 'Перемоги' ,
                    'пер.': 'пров. ' ,
                    'Школьный': 'Шкільний' ,
                    'Шевченко': 'Шевченка' ,
                    'Лазурная': 'Лазурна' ,
                    }  # и т. д.

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

    def SaveAndClose (self, path):
        self.wb.SaveAs(path+u'Боржники.xlsx' , 51 )
        self.wb.Close ( )
        self.Excel.Quit ( )
        return


    def addWarning(self):
        pass


if __name__ == '__main__':
    print(sys.argv)

    if len(sys.argv) > 1:
        fname = sys.arg[ 1 ]
    else:
        fname = u'D:\\DOLG.DBF'

    path=p.dirname(fname)

    debtor = Debtor ( fname )
    assert isinstance ( debtor , object )
    debtor.ReplaceInSheet ( )
    debtor.SaveAndClose (path)
