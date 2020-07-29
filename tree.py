from openpyxl import load_workbook

def get_tree(fields, get_depth):
    res = []
    indices = [0, 0, 0, 0]
    depth = 0
    mem = 0
    for i in range(len(fields)):
        d, m = get_depth(fields[i], mem)
        mem = m
        if mem == 2:
            if depth == 3:
                depth = 1
            else:
                depth -= 1
            mem = 1
            if depth == 0:
                indices = [indices[0] + 1, 0, 0, 0]
            elif depth == 1:
                indices = [indices[0], indices[1] + 1, 0, 0]
            else:
                indices = [indices[0], indices[1], indices[2] + 1, 0]
        elif d < depth:
            if depth != 3:
                mem = 0
            depth = d
            if depth == 0:
                indices = [indices[0] + 1, 0, 0, 0]
            elif depth == 1:
                indices = [indices[0], indices[1] + 1, 0, 0]
            else:
                indices = [indices[0], indices[1], indices[2] + 1, 0]
        else:
            depth = d
            indices[depth] += 1
        r = indices.copy()
        r.append(fields[i])
        r.append(i)
        res.append(r)
    return res

def tab_depth(s, mem):
    m = 1 if mem > 0 else 0
    m0 = m
    if s.lower().find('норма дисконта') != -1:
        m0 = m + 1
    if s.startswith('        ') or s.lower().startswith('в том числе'):
        if s[7:].startswith('    ') or s.lower().startswith('в том числе'):
            return 2 + m, m0
        return 1 + m, m0
    elif s.startswith('    ') or s.lower().startswith('в том числе'):
        return 1 + m, m0
    return 0 + m, m0

def parse_xl(filename, listname, index, start, stop):
    wb = load_workbook(filename=filename)
    ws = wb[listname]
    res = []

    for i in range(start, stop + 1):
        cell = ws[index + str(i)]
        if cell.alignment.horizontal == 'right':
            s = '        ' + str(cell.value)
        else:
            s = str(cell.value)
        res.append(s)

    return res

def get_full(cell, parsed):
    if cell[1] == cell[2] == cell[3] == 0:
        return cell[4]
    if cell[2] != 0 and cell[1] == 0:
        cell[1] = cell[2]
        cell[2] = 0
    if cell[3] != 0 and cell[2] == 0:
        cell[2] = cell[3]
        cell[3] = 0
    res = cell[4]

    for i in range(3, 0, -1):
        if cell[i] != 0:
            j = cell[5]
            f = cell[i - 1]
            while parsed[j - 1][i] != 0:
                j -= 1
            res = parsed[j-1][4].strip() + ' ' + res.strip()

    return res


#res = get_tree(parse_xl("tree.xlsx", '7', 'B', 5, 100), tab_depth)
#for r in res:
    #print(get_full(r, res))