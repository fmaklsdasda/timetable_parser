from openpyxl import load_workbook

workbook = load_workbook(filename='schedule.xlsx')
worksheet = workbook.active

data = []
for row in worksheet.iter_rows(min_row=2, max_row=41, min_col=1, max_col=41, values_only=True):
    data.append(row)

for row in data[:5]:
    print(row)


# 1 цикл проход по всем дням, определяем диапазон трок, массив [[понедельник, [диапазон строк 4- 10]], [день n, [n, n + i]]] 
# 2 цикл по парам в один день [от n до n + 1 мы проверяем каждую ячейку и забираем из нее фамилию учителя и название пары]
# 3 описание структуры справочника  

date_column = 1
first_row = 2 