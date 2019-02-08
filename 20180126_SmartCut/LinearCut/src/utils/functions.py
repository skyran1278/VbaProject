""" often use common function """


def concat_num_size(num, group_size):
    """ to output num-size format """
    if num == 0:
        return 0

    return f'{int(num)}-{group_size}'
    # return str(int(num)) + '-' + group_size


def num_to_1st_2nd(num, group_cap):
    """ all num to 1 and 2 row """
    if num - group_cap == 1:
        return group_cap - 1, 2

    if num > group_cap:
        return group_cap, num - group_cap

    return max(num, 2), 0
