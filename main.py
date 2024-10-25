from openpyxl import load_workbook
from collections import Counter

def get_value_list(t_2d):
    tmp = [[cell.value for cell in row] for row in t_2d]
    return list([x[0] for x in tmp])

file_path = 'book.xlsx'
workbook = load_workbook(file_path)
sheet = workbook['Sheet1']

clot = get_value_list(sheet['D4:D49']) + get_value_list(sheet['D51:D145'])
text = get_value_list(sheet['F4:F49']) + get_value_list(sheet['F51:F145'])
X = max(len(x) for x in clot)
Y = max(len(y) for y in clot)
pair = [f"{clot[i]:<{X}} + {text[i]:<{Y}}" for i in range(len(text))]
print(f"count : {len(text)}")
print()


def out(name, lis):
    print("-----------------")
    print(name)
    print()
    counter = Counter(lis)
    sorted_elements = counter.most_common()
    L = max(len(x) for x in lis)
    for x in sorted_elements:
        print(f"{x[0]:<{L + 1}} : {x[1]:>{3}}")
    print()


out("cloth color", clot)
out("text color", text)
out("color pair", pair)

