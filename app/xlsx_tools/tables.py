from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Font, Alignment, Border, Side


def draw_tables(ws):
    # задаём имя листа
    ws.title = "Бюджет"

    # захватываем активный лист

    dwar1(ws)
    expense(ws)
    all(ws)

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
    table = Table(displayName="IncomeTable", ref="A2:C8")

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
    # Цвет яцеек
    str3 = PatternFill(fill_type="solid", fgColor="FFD500")  
    ws['A15'].fill = str3
    ws['A16'].fill = str3
    ws['A17'].fill = str3
    ws['A18'].fill = str3
    ws['A19'].fill = str3
    ws['A20'].fill = str3
    ws['A21'].fill = str3
    ws['A22'].fill = str3
    ws['A23'].fill = str3
    ws['A24'].fill = str3
    ws['A25'].fill = str3
    ws['A26'].fill = str3
    ws['A27'].fill = str3
    ws['A28'].fill = str3
    ws['A29'].fill = str3
    ws['A30'].fill = str3

    ws['B15'].fill = str3
    ws['B16'].fill = str3
    ws['B17'].fill = str3
    ws['B18'].fill = str3
    ws['B19'].fill = str3
    ws['B20'].fill = str3
    ws['B21'].fill = str3
    ws['B22'].fill = str3
    ws['B23'].fill = str3
    ws['B24'].fill = str3
    ws['B25'].fill = str3
    ws['B26'].fill = str3
    ws['B27'].fill = str3
    ws['B28'].fill = str3
    ws['B29'].fill = str3
    ws['B30'].fill = str3

    str4 = PatternFill(fill_type="solid", fgColor="FF4400") 
    ws['C15'].fill = str4
    ws['C16'].fill = str4
    ws['C17'].fill = str4
    ws['C18'].fill = str4
    ws['C19'].fill = str4
    ws['C20'].fill = str4
    ws['C21'].fill = str4
    ws['C22'].fill = str4
    ws['C23'].fill = str4
    ws['C24'].fill = str4
    ws['C25'].fill = str4
    ws['C26'].fill = str4
    ws['C27'].fill = str4
    ws['C28'].fill = str4
    ws['C29'].fill = str4
    ws['C30'].fill = str4

    light_fill = PatternFill(fill_type="solid", fgColor="00E6FF")  
    dark_fill = PatternFill(fill_type="solid", fgColor="0022FF")   

    # Диапазон от D15 до AH30
    for row in range(15, 31):           # строки 15–30 включительно
        for col in range(4, 35):        # столбцы D (4) до AH (34) включительно
            cell = ws.cell(row=row, column=col)
            if (row + col) % 2 == 0:
                cell.fill = light_fill
            else:
                cell.fill = dark_fill
    # Заголовок таблицы "Расходы"
    ws.merge_cells("A14:C14")
    ws["A14"].value = "Расходы"
    ws["A14"].font = Font(
        name="Times New Roman",     # Название шрифта
        size=18,          # Размер шрифта
        bold=True,        # Жирный
        italic=True,      # Курсив
        underline="single",  # Подчёркивание
    )
    ws["A14"].alignment = Alignment(horizontal="center", vertical="center")

    # Заголовок под таблицу дней месяца
    ws.merge_cells("D14:AJ14")
    ws["D14"].value = "Дни месяца"
    ws["D14"].font = Font(
        name="Times New Roman",     # Название шрифта
        size=18,          # Размер шрифта
        bold=True,        # Жирный
        italic=True,      # Курсив
        underline="single",  # Подчёркивание
    )
    ws["D14"].alignment = Alignment(horizontal="left", vertical="center")

    # Шапка таблицы расходов
    ws["A15"].font = Font(
        name="Times New Roman",     # Название шрифта
        size=16,          # Размер шрифта
        bold=True,        # Жирный
    )
    ws["B15"].font = Font(
        name="Times New Roman",     # Название шрифта
        size=16,          # Размер шрифта
        bold=True,        # Жирный
    )
    ws["C15"].font = Font(
        name="Times New Roman",     # Название шрифта
        size=16,          # Размер шрифта
        bold=True,        # Жирный
    )

    ws["A15"] = "№"
    ws["B15"] = "Категория расходов"
    ws["C15"] = "Рас. за мес."
    ws["A15"].alignment = ws["B15"].alignment = ws["C15"].alignment = Alignment(horizontal="center", vertical="center")


    # Заголовки дней месяца (столбцы D по AH)
    for day in range(1, 32):
        cell = ws.cell(row=15, column=day + 3, value=str(day))
        cell.font = Font(
            name="Times New Roman",
            size=14,
            bold=True
        )
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Заполнение примерных категорий расходов
    categories = [
        "Продукты",
        "Транспорт",
        "Развлечения",
        "Одежда",
        "Здоровье",
        "Образование",
        "Подарки",
        "Автомобиль",           
        "Прочее"
    ]
    ws['C16'] = '=SUM(D16:AH16)'
    ws['C17'] = '=SUM(D17:AH17)'
    ws['C18'] = '=SUM(D18:AH18)'
    ws['C19'] = '=SUM(D19:AH19)'
    ws['C20'] = '=SUM(D20:AH20)'
    ws['C21'] = '=SUM(D21:AH21)'
    ws['C22'] = '=SUM(D22:AH22)'
    ws['C23'] = '=SUM(D23:AH23)'
    ws['C24'] = '=SUM(D24:AH24)'
    ws['C30'] = '=SUM(C16:C29)'

    ws['B30'] = 'Итого'
    ws["B30"].font = Font(
        name="Times New Roman",     # Название шрифта
        size=14,          # Размер шрифта
        bold=True,        # Жирный
    )
    ws["C30"].font = Font(
        name="Times New Roman",     # Название шрифта
        size=14,          # Размер шрифта
        bold=True,        # Жирный
    )

    ws['E16'] = 10000
    # Заполнение таблицы категориями расходов
    for i, category in enumerate(categories, start=16):
        # Номер строки
        ws.cell(row=i, column=1, value=i - 15).font = Font(
            name="Times New Roman",
            size=14
        )
        ws.cell(row=i, column=1).alignment = Alignment(horizontal="center")
        # Категория
        ws.cell(row=i, column=2, value=category).font = Font(
            name="Times New Roman",
            size=14
        )
        ws.cell(row=i, column=2).alignment = Alignment(horizontal="left")

  

    # Создание таблицы с названием "Расходы"
    table = Table(displayName="ExpenseTable", ref="A15:AH30")
    style = TableStyleInfo(
        name="TableStyleMedium9",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False
    )

    # Добавляем таблицу на лист
    ws.add_table(table)

    # Добавление чёрной тонкой границы ко всем ячейкам таблицы
    thin_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )
    for row in ws["A15:AH30"]:
        for cell in row:
            cell.border = thin_border


def all(ws):
    ws.merge_cells("G1:I1")  # Объединяем ячейки
    ws["G1"].value = "Общий отчет"
    ws["G1"].font = Font(
        name="Times New Roman",     # Название шрифта
        size=18,          # Размер шрифта
        bold=True,        # Жирный
        italic=True,      # Курсив
        underline="single",  # Подчёркивание
    )
    ws["G1"].alignment = Alignment(horizontal="center")
    # Настройка ширины столбцов
    ws.column_dimensions['G'].width = 25
    ws.column_dimensions['H'].width = 20
    ws.column_dimensions['I'].width = 15

    ws["G2"] = "№"
    ws["H2"] = "Описание"
    ws["I2"] = "Значение"
    ws["G2"].font = Font(
        name="Times New Roman",     # Название шрифта
        size=18,          # Размер шрифта
        bold=True,        # Жирный
        italic=True,      # Курсив
        underline="single",  # Подчёркивание
    )
    ws["H2"].font = ws["I2"].font = Font(
        name="Times New Roman",
        size=14,
        bold=True
    )
    ws["D14"].alignment = Alignment(horizontal="left", vertical="center")

    ws["G3"] = "Доходы за месяц"
    ws["G4"] = "Расходы за месяц"
    ws["G5"] = "Остаток на конец месяца"
    
    ws["H3"] = "Доходы"
    ws["H4"] = "Расходы"
    ws["H5"] = "Остаток"
    
    ws["I3"] = "=C8"
    ws["I4"] = "=C30"
    ws["I5"] = "=I3-I4"

    # Теперь создаём таблицу
    table = Table(displayName="Отчет", ref="G2:I5")
    thin_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )
    for row in ws["G2:I5"]:
        for cell in row:
            cell.border = thin_border

    ws.add_table(table)