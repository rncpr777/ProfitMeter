import openpyxl

book = openpyxl.load_workbook('file.xlsx')
sheet = book.active
mm = sum(1 for row in sheet.iter_rows() if any(cell.value for cell in row))
print('Enter year:')
year = input()
print('Enter month:')
month = input()
F = True
for row in range(mm):
    if str(sheet['A' + str(row + 1)].value) == year and str(sheet['B' + str(row + 1)].value) == month:
        N = str(row + 1)
        F = False
        break
if F:
    mrow = str(mm + 1)
    sheet['A' + mrow] = str(year)
    sheet['B' + mrow] = str(month)
    sheet['C' + mrow] = '0'
    sheet['D' + mrow] = '0'
    sheet['E' + mrow] = '0'
    sheet['F' + mrow] = sheet['F' + str((int(mrow) - 1))].value
    N = mrow
print('What do you want to edit?')
print('1 - Profits')
print('2 - Expences')
num = int(input())
if num == 1:
    print('Enter the profit:')
    p = int(input())
    p_t = int(sheet['C' + N].value)
    p = p_t + p
    sheet['C' + N] = str(p)
elif num == 2:
    print('Enter the expences:')
    e = int(input())
    e_t = int(sheet['D' + N].value)
    e = e_t + e
    sheet['D' + N] = str(e)
a = int(sheet['C' + N].value)
b = int(sheet['D' + N].value)
res = a - b
sheet['E' + N] = str(res)
if int(N) > 2:
    sheet['F' + N] = str(int(sheet['F' + str(int(N) - 1)].value) + int(sheet['E' + N].value))
else:
    sheet['F' + N] = sheet['E' + N].value
for row in range(3, mm + 1):
    sheet['F' + str(row)] = str(int(sheet['F' + str(row - 1)].value) + int(sheet['E' + str(row)].value))
book.save('file.xlsx')
print('Changes applied successfully, check your file!')
