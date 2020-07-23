from openpyxl import load_workbook

def get_index(line):
    line = str(line)
    return int(line.split('.')[0])

def mnemonic(a):
    if len(a) == 1:
        res = ord(a) - ord('A') + 1
    else:
        res = (ord(a[0]) - ord('A') + 1) * 26 + ord(a[1]) - ord('A') + 1
    return res

def amnemonic(res):
    if res > 26:
        code = chr(res // 26 + ord('A') - 1)
        code += chr(res % 26 + ord('A') - 1)
    else:
        code = chr(res + ord('A') - 1)
    return code

def get_len(start, stop):
    return mnemonic(stop.upper()) - mnemonic(start.upper()) + 1

def get_name(start, index):
    return amnemonic(mnemonic(start) + index)

def by_index(filename, listname, x_start, x_stop, y_start, y_stop, fields):
    length = get_len(x_start, x_stop)
    res = [['\t' for i in range(length-3)] for i in range(y_stop - y_start + 2)]
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

#print(by_index('test.xlsx', 'Лист1', 'B', 'AB', 4, 16, [str(x) + '.' + 'Stonks' for x in range(1, 15)]))