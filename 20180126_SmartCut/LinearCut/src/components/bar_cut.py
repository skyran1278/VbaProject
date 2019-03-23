"""
smart cut
"""
import numpy as np


from components.bar_functions import concat_num_size, num_to_1st_2nd
from data.dataset_rebar import rebar_area


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
    else:
        return group_loc.index[loc + 1]


def cut_5(etabs_design, beam_5, const):
    """
    5 cut
    """
    rebar, iteration_gap = const['rebar'], const['iteration_gap']

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

        bar_cap = 'Bar' + loc + 'Cap'
        bar_size = 'Bar' + loc + 'Size'
        bar_num_ld = 'Bar' + loc + 'NumLd'

        for _, group in etabs_design.groupby(['Story', 'BayID'], sort=False):
            min_usage = float('Inf')

            group_cap = group.at[group.index[0], bar_cap]
            group_size = group.at[group.index[0], bar_size]

            group_max = np.amax(group['StnLoc'])
            group_min = np.amin(group['StnLoc'])

            span = group_max - group_min

            # 這裡需要注意
            # left 和 right 的意義不太一樣
            left = span * iteration_gap['left'][0] + group_min
            right = span * (iteration_gap['right'][1]) + group_min

            iteration = group[bar_num_ld][(
                group['StnLoc'] >= left) & (group['StnLoc'] <= right)]

            # group_left = group[bar_num_ld][(
            #     group['StnLoc'] >= left[0]) & (group['StnLoc'] <= left[1])]
            # group_right = group[bar_num_ld][(
            #     group['StnLoc'] >= right[0]) & (group['StnLoc'] <= right[1])]

            iteration_diff = np.diff(iteration)

            # group_left_diff = np.diff(group_left)
            # group_right_diff = np.diff(group_right)

            iteration_diff = _make_1st_last_diff(iteration_diff)
            # group_left_diff = _make_1st_last_diff(group_left_diff)
            # group_right_diff = _make_1st_last_diff(group_right_diff)

            iteration_diff_nonzero = np.flatnonzero(iteration_diff)

            if len(iteration_diff_nonzero) == 2:
                # 做一個假的資料 讓他可以算
                split_left_1 = iteration.index[0]
                split_left_2 = iteration.index[1]
                split_right_2 = iteration.index[-2]
                split_right_1 = iteration.index[-1]

                split_5 = [
                    group.loc[:split_left_1, bar_num_ld],
                    group.loc[split_left_1: split_left_2,
                              bar_num_ld],
                    group.loc[split_left_2: split_right_2,
                              bar_num_ld],
                    group.loc[split_right_2: split_right_1,
                              bar_num_ld],
                    group.loc[split_right_1:, bar_num_ld]
                ]

                min_num, min_length = _calc_num_length(group, split_5)

                min_usage = np.sum(min_num * min_length)

            elif len(iteration_diff_nonzero) == 3:
                split_left_1 = iteration.index[0]
                split_left_2 = iteration.index[iteration_diff_nonzero[1]]
                split_right_2 = iteration.index[-2]
                if split_right_2 == split_left_2:
                    split_right_2 = iteration.index[1]
                split_right_1 = iteration.index[-1]

                split_5 = [
                    group.loc[:split_left_1, bar_num_ld],
                    group.loc[split_left_1: split_left_2,
                              bar_num_ld],
                    group.loc[split_left_2: split_right_2,
                              bar_num_ld],
                    group.loc[split_right_2: split_right_1,
                              bar_num_ld],
                    group.loc[split_right_1:, bar_num_ld]
                ]

                min_num, min_length = _calc_num_length(group, split_5)

                min_usage = np.sum(min_num * min_length)

            else:
                iteration_diff_nonzero_range = range(
                    len(iteration_diff_nonzero))
                for first in iteration_diff_nonzero_range:
                    split_left_1 = _get_min_cut(
                        iteration, iteration_diff, iteration_diff_nonzero[first])
                    for second in iteration_diff_nonzero_range[(first + 1):]:
                        split_left_2 = _get_min_cut(
                            iteration, iteration_diff, iteration_diff_nonzero[second])
                        for third in iteration_diff_nonzero_range[(second + 1):]:
                            split_right_2 = _get_min_cut(
                                iteration, iteration_diff, iteration_diff_nonzero[third])
                            for forth in iteration_diff_nonzero_range[(third + 1):]:
                                split_right_1 = _get_min_cut(
                                    iteration, iteration_diff, iteration_diff_nonzero[forth])

                                split_5 = [
                                    group.loc[:split_left_1, bar_num_ld],
                                    group.loc[split_left_1: split_left_2,
                                              bar_num_ld],
                                    group.loc[split_left_2: split_right_2,
                                              bar_num_ld],
                                    group.loc[split_right_2: split_right_1,
                                              bar_num_ld],
                                    group.loc[split_right_1:, bar_num_ld]
                                ]

                                num, length = _calc_num_length(group, split_5)

                                rebar_usage = np.sum(num * length)

                                if rebar_usage < min_usage:
                                    min_usage = rebar_usage
                                    min_num = num
                                    min_length = length
            # for row in nonzero_index:
            #     split_left = _get_min_cut(iteration, iteration_diff, row)

            #     for j in np.flatnonzero(group_right_diff):
            #         split_right = _get_min_cut(
            #             group_right, group_right_diff, j)
            #         split_3p_array = [
            #             group.loc[:split_left, bar_num_ld], group.loc[split_left:
            # split_right, bar_num_ld], group.loc[split_right:, bar_num_ld]]
            #         num, length = _calc_num_length(group, split_3p_array)

                    # rebar_usage = np.sum(num * length)
                    # if rebar_usage < min_usage:
                    #     min_usage = rebar_usage
                    #     min_num = num
                    #     min_length = length

            group_num = {
                '左1': num_to_1st_2nd(min_num[0], group_cap),
                '左2': num_to_1st_2nd(min_num[1], group_cap),
                '中': num_to_1st_2nd(min_num[2], group_cap),
                '右2': num_to_1st_2nd(min_num[-2], group_cap),
                '右1': num_to_1st_2nd(min_num[-1], group_cap)
            }

            group_length = {
                '左1': min_length[0],
                '左2': min_length[1],
                '中': min_length[2],
                '右2': min_length[-2],
                '右1': min_length[-1]
            }

            for bar_loc in group_num:
                loc_1st, loc_2nd = group_num[bar_loc]
                loc_length = group_length[bar_loc]
                beam_5.at[row, ('主筋', bar_loc)] = concat_num_size(
                    loc_1st, group_size)
                beam_5.at[row, ('主筋長度', bar_loc)] = loc_length * 100
                beam_5.at[row + to_2nd, ('主筋', bar_loc)
                          ] = concat_num_size(loc_2nd, group_size)

            beam_5.at[row, ('NOTE', '')] = min_usage * rebar_area(
                group_size) * 1000000

            row += 4

    return beam_5


def cut_3(group, loc, const):
    """
    cut 3, depands on iteration_gap, ex: 0.1~0.45, 0.55~0.9
    """
    iteration_gap = const['iteration_gap']

    # initial
    min_usage = float('Inf')

    bar_num_ld = 'Bar' + loc + 'NumLd'

    group_max = np.amax(group['StnLoc'])
    group_min = np.amin(group['StnLoc'])

    left = (group_max - group_min) * iteration_gap['left'] + group_min
    right = (group_max - group_min) * (
        iteration_gap['right']) + group_min

    group_left = group[bar_num_ld][(
        group['StnLoc'] >= left[0]) & (group['StnLoc'] <= left[1])]
    group_right = group[bar_num_ld][(
        group['StnLoc'] >= right[0]) & (group['StnLoc'] <= right[1])]

    group_left_diff = np.diff(group_left)
    group_right_diff = np.diff(group_right)

    group_left_diff = _make_1st_last_diff(group_left_diff)
    group_right_diff = _make_1st_last_diff(group_right_diff)

    for i in np.flatnonzero(group_left_diff):  # pylint: disable=invalid-name
        split_left = _get_min_cut(group_left, group_left_diff, i)

        for j in np.flatnonzero(group_right_diff):  # pylint: disable=invalid-name

            split_right = _get_min_cut(
                group_right, group_right_diff, j)
            split_3p_array = [
                group.loc[:split_left, bar_num_ld],
                group.loc[split_left: split_right, bar_num_ld],
                group.loc[split_right:, bar_num_ld]
            ]
            num, length = _calc_num_length(group, split_3p_array)

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

            num, length, min_usage = cut_3(group, loc, const)

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
                beam.at[row, ('主筋長度', bar_loc)] = loc_length * 100
                beam.at[row + to_2nd, ('主筋', bar_loc)
                        ] = concat_num_size(loc_2nd, group_size)

            beam.at[row, ('NOTE', '')] = min_usage * rebar_area(
                group_size) * 1000000

            row += 4

    return beam


def cut_optimization(moment, beam, etabs_design, const):
    """
    cut 3 or 5, optimization
    """
    if moment == 3:
        return output_3(beam, etabs_design, const)
    return cut_5(beam, etabs_design, const)


def main():
    """
    test
    """
    from components.init_beam import init_beam
    from const import const
    from data.dataset_etabs_design import load_beam_design
    from data.dataset_e2k import load_e2k
    from utils.execution_time import Execution
    from components.stirrups import calc_stirrups
    from components.bar_size_num import calc_db
    from components.bar_ld import calc_ld, add_ld

    e2k_path, etabs_design_path = const['e2k_path'], const['etabs_design_path']

    e2k = load_e2k(e2k_path, e2k_path + '.pkl')
    etabs_design = load_beam_design(
        etabs_design_path, etabs_design_path + '.pkl')

    beam = init_beam(etabs_design, e2k, moment=3)
    execution = Execution()
    beam, dh_design = calc_stirrups(beam, etabs_design, const)

    db_design = calc_db('BayID', dh_design, e2k, const)

    ld_design = calc_ld(db_design, e2k, const)

    ld_design = add_ld(ld_design, 'Ld', const['rebar'])

    execution.time('cut 3')
    beam = output_3(beam, ld_design, const)
    print(beam.head())
    execution.time('cut 3')


if __name__ == '__main__':
    main()
