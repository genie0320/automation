from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.title = 'GenieFile'
wb.save('sample.xlsx')
wb.close()
