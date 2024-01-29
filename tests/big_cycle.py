import sheets
import decimal

wb = sheets.Workbook()
index, name = wb.new_sheet()

for i in range(2,1000):
    wb.set_cell_contents(name, "A" + str(i), "=A" + str(i-1))

wb.set_cell_contents(name, "A1", "1.0")

assert(wb.get_cell_value(name, "A999") == decimal.Decimal(1.0))

wb.set_cell_contents(name, "A1", "=A999")

a1 = wb.get_cell_value(name, "A1")
assert(type(a1) == sheets.CellError)
assert(a1.get_type() == sheets.CellErrorType.CIRCULAR_REFERENCE)
