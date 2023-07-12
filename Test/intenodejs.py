import sys
import json

print(f'The info')

import xlwings as xw
print(f'Before Getting')
wbxl=xw.Book('template/def.xlsx')


cell_value_E75 = wbxl.sheets['確認'].range('E75').value
print(f'Number of sales results: {cell_value_E75}')
xw.apps.active.quit()
#
# print(f'Before Getting')

# # Retrieve arguments passed from Node.js
# a1 = sys.argv[1]
# a2 = json.loads(sys.argv[2])
# b1 = int(sys.argv[3])
# b2 = json.loads(sys.argv[4])
#
# result = b2[0] + b2[1]
# # Print the arguments
# print(f'a1: {a1}')
# print(f'a1: {a2[0]}')
# print(f'a1: {a2[1]}')
# print(f'a2: {a2}')
# print(f'b1: {b1}')
# print(f'result: {result}')
