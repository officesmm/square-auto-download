import sys
import xlwings as xw
wbxl=xw.Book('template/SquareTemplate.xlsx')

cell_value_E76 = wbxl.sheets['確認'].range('E76').value
cell_value_E77 = wbxl.sheets['確認'].range('E77').value
cell_value_E75 = wbxl.sheets['確認'].range('E75').value
cell_value_E78 = wbxl.sheets['確認'].range('E78').value
cell_value_E79 = wbxl.sheets['確認'].range('E79').value
cell_value_E80 = wbxl.sheets['確認'].range('E80').value
cell_value_E81 = wbxl.sheets['確認'].range('E81').value

print(f"Number of Earning Result: {cell_value_E75}")
print(f"Earning Result: {cell_value_E76}")
print(f"Number of Refund Result: {cell_value_E77}")
print(f"Refund Result: {cell_value_E78}")
print(f"Number of receipts result: {cell_value_E79}")
print(f"Transaction volume result: {cell_value_E80}")
print(f"Payment amount result: {cell_value_E81}")

xw.apps.active.quit()
