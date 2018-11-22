def concat_num_size(num, group_size):
    if num == 0:
        return 0
    return f'{int(num)}-{group_size}'
    # return str(int(num)) + '-' + group_size


def num_to_1st_2nd(num, group_cap):
    if num - group_cap == 1:
        return group_cap - 1, 2
    elif num > group_cap:
        return group_cap, num - group_cap
    else:
        return max(num, 2), 0
