import openpyxl

workbook = openpyxl.load_workbook('vendas_de_produtos.xlsx')
vendas_sheet = workbook['vendas']


with open('dados.txt', 'w', encoding='utf-8') as file:
    for linha in vendas_sheet.iter_rows(min_row=1):
        for cell in linha:
            file.write(str(cell.value) + ', ')
        file.write('\n')
print("Dados transferidos com sucesso!")
       
       
    
    
