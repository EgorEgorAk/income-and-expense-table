from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Font, Alignment, Border, Side


def draw_tables(ws , table_name: str):
    # задаём имя листа
    ws.title = "asрывлофрывлфровasdasdadsБюдже"

    # захватываем активный лист

    dwar1(ws)
    expense(ws)

def dwar1(ws):
    # ширина столбцов
    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['C'].width = 15

    # жирный шрифт
    ws['A2'].font = Font(
    name="Times New Roman",     # Название шрифта
    size=18,          # Размер шрифта
    bold=True,        # Жирный
    italic=True,      # Курсив
    underline="single",  # Подчёркивание
    )
    ws['B2'].font = Font(
    name="Times New Roman",     # Название шрифта
    size=18,          # Размер шрифта
    bold=True,        # Жирный
    italic=True,      # Курсив
    underline="single",  # Подчёркивание
    )
    ws['C2'].font = Font( 
    name="Times New Roman",     # Название шрифта
    size=18,          # Размер шрифта
    bold=True,        # Жирный
    italic=True,      # Курсив
    underline="single",  # Подчёркивание
    )
    ws['C3'].font = Font(
    name="Times New Roman",     # Название шрифта
    size=14,          # Размер шрифта
    bold=True,        # Жирный
    italic=False,      # Курсив
    
    )
    ws['C4'].font = Font(
    name="Times New Roman",     # Название шрифта
    size=14,          # Размер шрифта
    bold=True,        # Жирный
    italic=False,      # Курсив
    
    )
    ws['C5'].font = Font(
    name="Times New Roman",     # Название шрифта
    size=14,          # Размер шрифта
    bold=True,        # Жирный
    italic=False,      # Курсив
    
    )
    ws['C6'].font = Font(
        name="Times New Roman",     # Название шрифта
        size=14,          # Размер шрифта
        bold=True,        # Жирный
        italic=False,      # Курсив
        
    )
    ws['C7'].font = Font(
        name="Times New Roman",     # Название шрифта
        size=14,          # Размер шрифта
        bold=True,        # Жирный
        italic=False,      # Курсив
        
    )
    ws['C8'].font = Font(
    name="Times New Roman",     # Название шрифта
    size=18,          # Размер шрифта
    bold=True,        # Жирный
    italic=True,      # Курсив
    underline="single",  # Подчёркивание
    
    )
    # цвет ячеек
    str1 = PatternFill(fill_type="solid", fgColor="FFA200")  
    ws['A2'].fill = str1
    ws['B2'].fill = str1
    ws['C2'].fill = str1
    str2 = PatternFill(fill_type="solid", fgColor="44FF00") 
    ws['C8'].fill = str2

    # Основные данные таблицы
    main_table = PatternFill(fill_type="solid", fgColor="FF00FF") 
    ws['B3'].fill = main_table
    ws['B4'].fill = main_table
    ws['B5'].fill = main_table
    ws['B6'].fill = main_table
    ws['B7'].fill = main_table
    sum = PatternFill(fill_type="solid", fgColor="00FFA2")
    ws['C3'].fill = sum
    ws['C4'].fill = sum
    ws['C5'].fill = sum
    ws['C6'].fill = sum
    ws['C7'].fill = sum


    # Данные можно назначать непосредственно ячейкам
    ws['A2'] = '№'
    ws['A3'] = 1
    ws['A4'] = 2
    ws['A5'] = 3
    ws['A6'] = 4
    ws['A7'] = 5

    ws['B2'] = 'Категории доходов'
    ws['B3'] = 'Зарплата'
    ws['B4'] = 'Дивиденды'
    ws['B5'] = 'Переводы от родителей'
    ws['B6'] = 'Премия'
    ws['B7'] = 'Прочее'

    ws['C2'] = 'Сумма'
    ws['C3'] = 5000
    ws['C4'] = 2000
    ws['C5'] = 8000
    ws['C6'] = 10000
    ws['C7'] = 3000

    # выравнивание текста
    ws['B7'].alignment = Alignment(horizontal='right', vertical='bottom') 
    ws['C2'].alignment = Alignment(horizontal='center', vertical='bottom') 
    ws['C3'].alignment = Alignment(horizontal='center', vertical='bottom') 
    ws['C4'].alignment = Alignment(horizontal='center', vertical='bottom') 
    ws['C5'].alignment = Alignment(horizontal='center', vertical='bottom') 
    ws['C6'].alignment = Alignment(horizontal='center', vertical='bottom') 
    ws['C7'].alignment = Alignment(horizontal='center', vertical='bottom') 
    ws['C8'].alignment = Alignment(horizontal='center', vertical='bottom') 

    ws['B8'] = 'Итого'
    ws['C8'] = '=SUM(C3:C7)'

    # Создаём таблицу название таблицы
    ws.merge_cells("A1:C1")  # Объединяем ячейки
    title_cell = ws["A1"]
    title_cell.value = "Доходы"
    title_cell.font = Font(bold=True, size=14)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")

    # Создаём таблицу
    table = Table(displayName="Доходы", ref="A2:C8")

    # Добавляем таблицу на лист
    ws.add_table(table)

    # Добавляем стили
    # Чёрная тонкая граница
    thin_border = Border(
    left=Side(style='thin', color='000000'),
    right=Side(style='thin', color='000000'),
    top=Side(style='thin', color='000000'),
    bottom=Side(style='thin', color='000000')
    )
    # Пройтись по ячейкам от A2 до C8 и задать границу
    for row in ws["A2:C8"]:
        for cell in row:
            cell.border = thin_border
    # Сохраняем файл


def expense(ws):
    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['C'].width = 15

    # 🔧 Заголовки для таблицы (от A до AJ — это 36 столбцов)
    headers = [f"Колонка {i}" for i in range(1, 37)]  # Или твои реальные названия
    for col, header in enumerate(headers, start=1):
        ws.cell(row=15, column=col, value=header)

    # 👇 Убери объединение A15:AJ30 (она затрёт заголовки), перенеси его ниже или замени
    # ws.merge_cells("A15:AJ30")  # ❌ Не объединяй строку заголовков

    # Можно объединить A14:AJ14 и туда поместить заголовок таблицы
    ws.merge_cells("A14:E14")
    ws["A14"].value = "Расходы"
    ws["A14"].font = Font(bold=True, size=14)
    ws["A14"].alignment = Alignment(horizontal="center", vertical="center")

    # 📊 Создание таблицы
    table = Table(displayName="Расходы", ref="A15:AJ30")
    ws.add_table(table)

    thin_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )
    for row in ws["A15:AJ30"]:
        for cell in row:
            cell.border = thin_border