import sort_function
import copy


def merge(d, e, edge, dir_dict):
    flag1 = False
    list_d = []
    for entry in copy.deepcopy(dir_dict[edge[0]]):
        if entry[0] == d:
            flag1 = True
            list_d.append(entry)
    if not flag1:
        dir_dict[edge[0]].append([d, edge[1], e])
    else:
        for entry in list_d:
            e_prime = entry[2]
            if e < e_prime:
                dir_dict[edge[0]].remove(entry)
                dir_dict[edge[0]].append([d, edge[1], e])
    for entry in copy.deepcopy(dir_dict[edge[0]]):
        if (entry[0] > d) and (entry[2] >= e):
            dir_dict[edge[0]].remove(entry)
    dir_dict[edge[0]].sort(key=sort_function.take_first, reverse=True)
    return dir_dict
