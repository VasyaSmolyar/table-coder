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
        if i == 129 or i == 128 or i == 130:
            print(m)
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
