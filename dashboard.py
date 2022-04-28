import xlwings as xw
macro = "C:\\Users\sdecaluwe\Desktop\zlivfacs\macros.xlsm"
wb = xw.Book(macro)
macro1 = wb.macro("Module45.MiseEnFormeZlivFacPayview")
macro1()


