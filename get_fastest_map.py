import configs_pretrain
import initilization
import copy
import relax_fastest
import numpy as np


def generate_tables(goals_list=configs_pretrain.goal_list, obstacle_list=configs_pretrain.obstacle_list,
                    map_length=configs_pretrain.init_env_settings[1],
                    map_width=configs_pretrain.init_env_settings[2]):
    delay_dic = configs_pretrain.delay_dic
    # print(delay_dic[(8, 55), (8, 54)])
    inf = 2 ** 31 - 1
    goals_indice = list()
    for i in goals_list:
        goals_indice.append((i[1], i[2]))
    dir_dic = initilization.initialization([goals_indice[0]])

    q_list = [goals_indice[0]]

    while q_list:
        x, y = q_list.pop(0)
        up = x - 1, y
        if (up[0] > 0) and (list(up) not in obstacle_list):
            dir_dic_up = copy.deepcopy(dir_dic[up])
            dir_dic = relax_fastest.relax([up, (x, y)], dir_dic, inf, delay_dic)
            if not np.all(dir_dic[up] == dir_dic_up):
                if up not in q_list:
                    q_list.append(up)

        down = x + 1, y
        if (down[0] < (map_width - 1)) and (list(down) not in obstacle_list):
            # print('down', dir_dic_n[down])
            dir_dic_down = copy.deepcopy(dir_dic[down])
            dir_dic = relax_fastest.relax([down, (x, y)], dir_dic, inf, delay_dic)
            if not np.all(dir_dic[down] == dir_dic_down):
                if down not in q_list:
                    q_list.append(down)

        left = x, y - 1
        if (left[1] > 0) and (list(left) not in obstacle_list):
            # print('left', dir_dic_n[left])
            dir_dic_left = copy.deepcopy(dir_dic[left])
            dir_dic = relax_fastest.relax([left, (x, y)], dir_dic, inf, delay_dic)
            if not np.all(dir_dic[left] == dir_dic_left):
                if left not in q_list:
                    q_list.append(left)

        right = x, y + 1
        if (right[1] < map_length - 1) and (list(right) not in obstacle_list):
            # print('right', dir_dic_n[right])
            dir_dic_right = copy.deepcopy(dir_dic[right])
            dir_dic = relax_fastest.relax([right, (x, y)], dir_dic, inf, delay_dic)
            if not np.all(dir_dic[right] == dir_dic_right):
                if right not in q_list:
                    q_list.append(right)
        print(len(q_list))
    return dir_dic
