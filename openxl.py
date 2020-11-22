import openpyxl

filename = "Words.xlsx"

f = open("dict.txt")

data = []

book = openpyxl.load_workbook(filename)

sheet = book.worksheets[0]

get_cell1 = sheet['B2' : 'B500']
get_cell2 = sheet['D2' : 'D500']

for row in get_cell1:
    for cell in row:
        pre_word = (cell.value).split(" ")[1]
        word = pre_word.split(",")[0]
        print(word)

for row in get_cell2:
    for cell in row:
        pre_word = (cell.value).split(" ")[1]
        word = pre_word.split(",")[0]
        print(word)
