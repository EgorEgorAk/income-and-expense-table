from openpyxl import Workbook
from xlsx_tools.tables import draw_tables


def main():
  workBook = Workbook()
  sheet = workBook.active
  tablePath = "files/table_test.xlsx"

  draw_tables(sheet, tablePath)
  workBook.save(tablePath)

if __name__ == "__main__":
  main()