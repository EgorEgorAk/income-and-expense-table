from openpyxl import Workbook
from xlsx_tools.tables import draw_tables


def main():
  wb = Workbook()
  sheet = wb.active
  tablePath = "app/files/table_test.xlsx"

  draw_tables(sheet)
  wb.save(tablePath)

if __name__ == "__main__":
  main()