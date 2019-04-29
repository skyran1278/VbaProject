"""
smart cut
"""
from itertools import combinations, tee
import numpy as np


from src.bar_functions import concat_num_size, num_to_1st_2nd
from src.rebar import rebar_area


def _calc_num_length(group, split_array):
    num = np.empty_like(split_array)
    length = np.empty_like(split_array)

    for counter, _ in enumerate(split_array):
        num[counter] = np.amax(split_array[counter])
        length[counter] = group.at[split_array[counter].index[-1], 'StnLoc'] - group.at[
            split_array[counter].index[0], 'StnLoc']
    return num, length


def _make_1st_last_diff(group_diff):
    if group_diff[0] == 0:
        group_diff[0] = 1

    if group_diff[-1] == 0:
        group_diff[-1] = -1

    return group_diff


def _get_min_cut(group_loc, group_loc_diff, loc):
    if group_loc_diff[loc] > 0:
        return group_loc.index[loc]

    return group_loc.index[loc + 1]


def cut_multiple(df, col, boundary, group_num=5):
    """
    multiple cut
    """
    if group_num <= 3:
        return cut_3(df, col, boundary)

    # initial
    min_usage = float('Inf')
    num = np.empty(group_num)
    length = np.empty(group_num)

    diff_area = (
        (df[col].diff() != 0) |
        (df[col].shift(-1).diff() != 0)
    )

    for index in combinations(df.index[diff_area], group_num - 1):

        num[0] = df.loc[:index[0], col].max()
        length[0] = df.loc[index[0], 'StnLoc'] - df['StnLoc'].min()
        num[-1] = df.loc[index[-1]:, col].max()
        length[-1] = df['StnLoc'].max() - df.loc[index[-1], 'StnLoc']

        for i, j in enumerate(range(len(index), 1)):
            num[j] = df.loc[i:j, col].max()
            length[j] = df.loc[j, 'StnLoc'] - df.loc[i, 'StnLoc']

        usage = np.sum(num * length)

        if usage < min_usage:
            min_usage = usage
            min_num = num
            min_length = length
            min_index = index

    if np.diff(min_num):
        return cut_multiple(df, col, boundary, group_num-1)
    return min_num, min_length, min_usage, min_index


def cut_3(df, col, boundary):
    """
    cut 3, depands on boundary, ex: 0.1~0.45, 0.55~0.9
    """
    # initial
    min_usage = float('Inf')

    amin = df['StnLoc'].min()
    amax = df['StnLoc'].max()

    left_boundarys = (amax - amin) * boundary['left'] + amin
    right_boundarys = (amax - amin) * boundary['right'] + amin

    left = [
        (df['StnLoc'] >= left_boundarys[0]) &
        (df['StnLoc'] <= left_boundarys[1])
    ]
    right = [
        (df['StnLoc'] >= right_boundarys[0]) &
        (df['StnLoc'] <= right_boundarys[1])
    ]

    diff_area = (
        (df[col].diff() != 0) |
        (df[col].diff().shift(-1) != 0)
    )

    left_idx = (
        df.index[left][0],
        *df.index[diff_area & left],
        df.index[left][-1]
    )

    right_idx = (
        df.index[right][0],
        *df.index[diff_area & right],
        df.index[right][-1]
    )

    group_left_diff = np.diff(left)
    group_right_diff = np.diff(right)

    group_left_diff = _make_1st_last_diff(group_left_diff)
    group_right_diff = _make_1st_last_diff(group_right_diff)

    for i in np.flatnonzero(group_left_diff):
        split_left = _get_min_cut(left, group_left_diff, i)

        for j in np.flatnonzero(group_right_diff):

            split_right = _get_min_cut(
                right, group_right_diff, j)
            split_3p_array = [
                df.loc[:split_left, col],
                df.loc[split_left: split_right, col],
                df.loc[split_right:, col]
            ]
            num, length = _calc_num_length(df, split_3p_array)

            rebar_usage = np.sum(num * length)

            if rebar_usage < min_usage:
                min_usage = rebar_usage
                min_num = num
                min_length = length

    return min_num, min_length, min_usage


def output_3(beam, etabs_design, const):
    """
    format 3 cut
    """
    rebar = const['rebar']

    output_loc = {
        'Top': {
            'START_LOC': 0,
            'TO_2nd': 1
        },
        'Bot': {
            'START_LOC': 3,
            'TO_2nd': -1
        }
    }

    for loc in rebar:
        row = output_loc[loc]['START_LOC']
        to_2nd = output_loc[loc]['TO_2nd']

        for _, group in etabs_design.groupby(['Story', 'BayID'], sort=False):
            # group capacity and size
            group_cap = group.at[group.index[0], 'Bar' + loc + 'Cap']
            group_size = group.at[group.index[0], 'Bar' + loc + 'Size']

            num, length, min_usage = cut_3(
                group, 'Bar' + loc + 'NumLd', const['boundary'])

            group_num = {
                '左': num_to_1st_2nd(num[0], group_cap),
                '中': num_to_1st_2nd(num[1], group_cap),
                '右': num_to_1st_2nd(num[2], group_cap)
            }

            group_length = {
                '左': length[0],
                # '左': length[0] if num[0] != num[1] else '',
                '中': length[1],
                # '右': length[2] if num[2] != num[1] else ''
                '右': length[2]
            }

            for bar_loc in group_num:
                loc_1st, loc_2nd = group_num[bar_loc]
                loc_length = group_length[bar_loc]
                beam.at[row, ('主筋', bar_loc)] = concat_num_size(
                    loc_1st, group_size)
                beam.at[row, ('主筋長度', bar_loc)] = round(loc_length * 100, 3)
                beam.at[row + to_2nd, ('主筋', bar_loc)
                        ] = concat_num_size(loc_2nd, group_size)

            beam.at[row, ('NOTE', '')] = min_usage * rebar_area(
                group_size) * 1000000

            row += 4

    return beam


def cut_optimization(beam, etabs_design, const):
    """
    cut 3 or 5, optimization
    """
    if ('主筋', '左') in beam:
        return output_3(beam, etabs_design, const)
    return cut_5(beam, etabs_design, const)


def main():
    """
    test
    """
    from src.execution_time import Execution
    from tests.const import const
    from src.beam import init_beam
    from src.e2k import load_e2k
    from src.etabs_design import load_etabs_design, post_e2k
    from src.stirrups import calc_stirrups
    from src.bar_size_num import calc_db
    from src.bar_ld import calc_ld, add_ld

    execution = Execution()

    e2k = load_e2k(const['e2k_path'])
    etabs_design = load_etabs_design(const['etabs_design_path'])
    etabs_design = post_e2k(etabs_design, e2k)
    beam = init_beam(etabs_design, moment=3)
    beam, etabs_design = calc_stirrups(beam, etabs_design, const)
    etabs_design = calc_db('BayID', etabs_design, const)
    etabs_design = calc_ld(etabs_design, const)
    etabs_design = add_ld(etabs_design, 'Ld', const['rebar'])

    execution.time('cut 3')
    # beam = output_3(beam, etabs_design, const)
    beam = cut_optimization(beam, etabs_design, const)
    print(beam.head())
    execution.time('cut 3')


if __name__ == '__main__':
    main()

# def cut_5(etabs_design, beam_5, const):
#     """
#     5 cut
#     """
#     rebar, iteration_gap = const['rebar'], const['iteration_gap']

#     output_loc = {
#         'Top': {
#             'START_LOC': 0,
#             'TO_2nd': 1
#         },
#         'Bot': {
#             'START_LOC': 3,
#             'TO_2nd': -1
#         }
#     }

#     for loc in rebar:
#         row = output_loc[loc]['START_LOC']
#         to_2nd = output_loc[loc]['TO_2nd']

#         bar_cap = 'Bar' + loc + 'Cap'
#         bar_size = 'Bar' + loc + 'Size'
#         bar_num_ld = 'Bar' + loc + 'NumLd'

#         for _, group in etabs_design.groupby(['Story', 'BayID'], sort=False):
#             min_usage = float('Inf')

#             group_cap = group.at[group.index[0], bar_cap]
#             group_size = group.at[group.index[0], bar_size]

#             group_max = np.amax(group['StnLoc'])
#             group_min = np.amin(group['StnLoc'])

#             span = group_max - group_min

#             # 這裡需要注意
#             # left 和 right 的意義不太一樣
#             left = span * iteration_gap['left'][0] + group_min
#             right = span * (iteration_gap['right'][1]) + group_min

#             iteration = group[bar_num_ld][(
#                 group['StnLoc'] >= left) & (group['StnLoc'] <= right)]

#             # group_left = group[bar_num_ld][(
#             #     group['StnLoc'] >= left[0]) & (group['StnLoc'] <= left[1])]
#             # group_right = group[bar_num_ld][(
#             #     group['StnLoc'] >= right[0]) & (group['StnLoc'] <= right[1])]

#             iteration_diff = np.diff(iteration)

#             # group_left_diff = np.diff(group_left)
#             # group_right_diff = np.diff(group_right)

#             iteration_diff = _make_1st_last_diff(iteration_diff)
#             # group_left_diff = _make_1st_last_diff(group_left_diff)
#             # group_right_diff = _make_1st_last_diff(group_right_diff)

#             iteration_diff_nonzero = np.flatnonzero(iteration_diff)

#             if len(iteration_diff_nonzero) == 2:
#                 # 做一個假的資料 讓他可以算
#                 split_left_1 = iteration.index[0]
#                 split_left_2 = iteration.index[1]
#                 split_right_2 = iteration.index[-2]
#                 split_right_1 = iteration.index[-1]

#                 split_5 = [
#                     group.loc[:split_left_1, bar_num_ld],
#                     group.loc[split_left_1: split_left_2,
#                               bar_num_ld],
#                     group.loc[split_left_2: split_right_2,
#                               bar_num_ld],
#                     group.loc[split_right_2: split_right_1,
#                               bar_num_ld],
#                     group.loc[split_right_1:, bar_num_ld]
#                 ]

#                 min_num, min_length = _calc_num_length(group, split_5)

#                 min_usage = np.sum(min_num * min_length)

#             elif len(iteration_diff_nonzero) == 3:
#                 split_left_1 = iteration.index[0]
#                 split_left_2 = iteration.index[iteration_diff_nonzero[1]]
#                 split_right_2 = iteration.index[-2]
#                 if split_right_2 == split_left_2:
#                     split_right_2 = iteration.index[1]
#                 split_right_1 = iteration.index[-1]

#                 split_5 = [
#                     group.loc[:split_left_1, bar_num_ld],
#                     group.loc[split_left_1: split_left_2,
#                               bar_num_ld],
#                     group.loc[split_left_2: split_right_2,
#                               bar_num_ld],
#                     group.loc[split_right_2: split_right_1,
#                               bar_num_ld],
#                     group.loc[split_right_1:, bar_num_ld]
#                 ]

#                 min_num, min_length = _calc_num_length(group, split_5)

#                 min_usage = np.sum(min_num * min_length)

#             else:
#                 iteration_diff_nonzero_range = range(
#                     len(iteration_diff_nonzero))
#                 for first in iteration_diff_nonzero_range:
#                     split_left_1 = _get_min_cut(
#                         iteration, iteration_diff, iteration_diff_nonzero[first])
#                     for second in iteration_diff_nonzero_range[(first + 1):]:
#                         split_left_2 = _get_min_cut(
#                             iteration, iteration_diff, iteration_diff_nonzero[second])
#                         for third in iteration_diff_nonzero_range[(second + 1):]:
#                             split_right_2 = _get_min_cut(
#                                 iteration, iteration_diff, iteration_diff_nonzero[third])
#                             for forth in iteration_diff_nonzero_range[(third + 1):]:
#                                 split_right_1 = _get_min_cut(
#                                     iteration, iteration_diff, iteration_diff_nonzero[forth])

#                                 split_5 = [
#                                     group.loc[:split_left_1, bar_num_ld],
#                                     group.loc[split_left_1: split_left_2,
#                                               bar_num_ld],
#                                     group.loc[split_left_2: split_right_2,
#                                               bar_num_ld],
#                                     group.loc[split_right_2: split_right_1,
#                                               bar_num_ld],
#                                     group.loc[split_right_1:, bar_num_ld]
#                                 ]

#                                 num, length = _calc_num_length(group, split_5)

#                                 rebar_usage = np.sum(num * length)

#                                 if rebar_usage < min_usage:
#                                     min_usage = rebar_usage
#                                     min_num = num
#                                     min_length = length
#             # for row in nonzero_index:
#             #     split_left = _get_min_cut(iteration, iteration_diff, row)

#             #     for j in np.flatnonzero(group_right_diff):
#             #         split_right = _get_min_cut(
#             #             group_right, group_right_diff, j)
#             #         split_3p_array = [
#             #             group.loc[:split_left, bar_num_ld], group.loc[split_left:
#             # split_right, bar_num_ld], group.loc[split_right:, bar_num_ld]]
#             #         num, length = _calc_num_length(group, split_3p_array)

#                     # rebar_usage = np.sum(num * length)
#                     # if rebar_usage < min_usage:
#                     #     min_usage = rebar_usage
#                     #     min_num = num
#                     #     min_length = length

#             group_num = {
#                 '左1': num_to_1st_2nd(min_num[0], group_cap),
#                 '左2': num_to_1st_2nd(min_num[1], group_cap),
#                 '中': num_to_1st_2nd(min_num[2], group_cap),
#                 '右2': num_to_1st_2nd(min_num[-2], group_cap),
#                 '右1': num_to_1st_2nd(min_num[-1], group_cap)
#             }

#             group_length = {
#                 '左1': min_length[0],
#                 '左2': min_length[1],
#                 '中': min_length[2],
#                 '右2': min_length[-2],
#                 '右1': min_length[-1]
#             }

#             for bar_loc in group_num:
#                 loc_1st, loc_2nd = group_num[bar_loc]
#                 loc_length = group_length[bar_loc]
#                 beam_5.at[row, ('主筋', bar_loc)] = concat_num_size(
#                     loc_1st, group_size)
#                 beam_5.at[row, ('主筋長度', bar_loc)] = loc_length * 100
#                 beam_5.at[row + to_2nd, ('主筋', bar_loc)
#                           ] = concat_num_size(loc_2nd, group_size)

#             beam_5.at[row, ('NOTE', '')] = min_usage * rebar_area(
#                 group_size) * 1000000

#             row += 4

#     return beam_5
