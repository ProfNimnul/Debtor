import Debtor

debtor = Debtor.Debtor()
debtor.ConvertDBFtoXLS()
debtor.OpenEXCELFile()
debtor.ReplaceInSheet()
debtor.CloseEXCEL()