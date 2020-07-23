from openpyxl import load_workbook

def get_index(line):
    line = str(line)
    return int(line.split('.')[0])

def get_len(start, stop):
    return ord(stop) - ord(start)

def get_name(start, index):
    return chr(ord(start) + index)

def by_index(filename, listname, x_start, x_stop, y_start, y_stop, fields):
    length = get_len(x_start, x_stop)
    res = [['#' for i in range(length-3)] for i in range(y_stop - y_start + 2)]
    wb = load_workbook(filename=filename)
    ws = wb[listname]

    for i in range(y_start, y_stop + 1):
        for j in range(len(fields)):
            ind = get_index(ws[x_start + str(i)].value)
            if ind == get_index(fields[j]):
                for k in range(3, length):
                    if ws[get_name(x_start, k) + str(i)].value == None:
                        res[get_index(ws[x_start + str(i)].value) - 1][k-3] = ''
                    else:
                        res[get_index(ws[x_start + str(i)].value) - 1][k-3] = ws[get_name(x_start, k) + str(i)].value
                break
    return res

#print(by_index('test.xlsx', 'Лист1', 'B', 'G', 4, 16, [str(x) + '.' + 'Stonks' for x in range(1, 15)]))