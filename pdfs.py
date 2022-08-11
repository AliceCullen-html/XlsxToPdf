# Primeiro importamos a lib
from win32com import client

# Depois vamos definir o objeto e se queremos que o processo apareça ou não na tela
app = client.Dispatch("Excel.Application")
app.Visible = False
app.Interactive = False


# Aqui devemos criar uma var que vai receber o local do arquivo
path = input('Digite o local do arquivo')

print('Convertendo arquivo..........')


workbook = app.Workbooks.Open(path)
workbook.ActiveSheet.ExportAsFixedFormat(0, path)
workbook.close
print('Seu arquivo foi convertido!')
