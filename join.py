from win32com.client import Dispatch
import os

xl = Dispatch("Excel.Application")
xl.Visible = True  
xl.DisplayAlerts = False
xl.ScreenUpdating = False
xl.Application.EnableEvents = False

result = xl.Workbooks.Add()

path = './excel/'
files = os.listdir(path)
for file in files:
    # if file.split('.')[-1] != 'xls':
    #     continue
    print(file)
    wb = xl.Workbooks.Open(Filename=os.path.abspath(path) + '\\' + file)
    ws = wb.Worksheets(1)
    ws.name = file.split(' ')[0]  # Номер ЛСР в имени файла ГрандСмета ставит первым.
    ws.Copy(Before=result.Worksheets(1))

result.SaveAs(Filename=os.path.abspath(path) + '\\ этап. Глава .xlsx')
result.Close(SaveChanges=True)
xl.Quit()
