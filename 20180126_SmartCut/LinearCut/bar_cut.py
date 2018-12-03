import os
import sys

import pandas as pd
import numpy as np

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.append(os.path.join(SCRIPT_DIR, os.path.pardir))

from utils.pkl import load_pkl
from utils.Clock import Clock
from utils.functions import concat_num_size, num_to_1st_2nd

from dataset.const import BAR, ITERATION_GAP
from dataset.dataset_e2k import load_e2k


# def calc_ld(beam_v_m):
#     # It is used for nominal concrete in case of phi_e=1.0 & phi_t=1.0.
#     # Reference:土木401-93
#     PI = 3.1415926

#     rebars, _, _, _, materials, sections = load_e2k()

#     def _ld(df, Loc):
#         # Loc = Loc.capitalize()

#         bar_size = 'Bar' + Loc + 'Size'
#         bar_1st = 'Bar' + Loc + '1st'

#         # 延伸長度比較熟悉 cm 操作
#         # m => cm
#         B = df['SecID'].apply(lambda x: sections[x, 'B']) * 100
#         material = df['SecID'].apply(lambda x: sections[x, 'MATERIAL'])
#         fc = material.apply(lambda x: materials[x, 'FC']) / 10
#         fy = material.apply(lambda x: materials[x, 'FY']) / 10
#         fyh = fy
#         cover = 0.04 * 100
#         db = df[bar_size].apply(lambda x: rebars[x, 'DIA']) * 100
#         num = df[bar_1st]
#         dh = df['VNoDuSize'].apply(lambda x: rebars[x, 'DIA']) * 100
#         spacing = df['SetSpacing'] * 100

#         # 5.2.2
#         fc[np.sqrt(fc) > 26.5] = 700

#         # R5.3.4.1.1
#         cc = dh + cover

#         # R5.3.4.1.1
#         cs = (B - db * num - dh * 2 - cover * 2) / (num - 1) / 2

#         # Vertical splitting failure / Horizontal splitting failure
#         cb = np.where(cc <= cs, cc, cs) + db / 2

#         # R5.3.4.1.2
#         ktr = np.where(cc <= cs, 1, 2 / num) * \
#             (PI * dh ** 2 / 4) * fyh / 105 / spacing

#         # if cs > cc:
#         #     # Vertical splitting failure
#         #     cb = db / 2 + cc
#         #     # R5.3.4.1.2
#         #     ktr = (PI * dh ** 2 / 4) * fyh / 105 / spacing
#         # else:
#         #     # Horizontal splitting failure
#         #     cb = db / 2 + cs
#         #     # R5.3.4.1.2
#         #     ktr = 2 * (PI * dh ** 2 / 4) * fyh / 105 / spacing / num

#         # 5.3.4.1
#         ld = 0.28 * fy / np.sqrt(fc) * db / np.minimum((cb + ktr) / db, 2.5)

#         # 5.3.4.1
#         simple_ld = 0.19 * fy / np.sqrt(fc) * db

#         # phi_s factor
#         ld[db < 2.2] = 0.8 * ld
#         simple_ld[db < 2.2] = 0.8 * simple_ld

#         # phi_t factor
#         if Loc == 'Top':
#             ld = 1.3 * ld
#             simple_ld = 1.3 * simple_ld

#         ld[ld > simple_ld] = simple_ld

#         # 5.3.1
#         ld[ld < 30] = 30

#         return {
#             # cm => m
#             Loc + 'Ld': ld / 100,
#             Loc + 'SimpleLd': simple_ld / 100
#         }

#     for Loc in BAR.keys():
#         beam_v_m = beam_v_m.assign(**_ld(beam_v_m, Loc))

#     return beam_v_m


# def add_ld(beam_v_m_ld):
#     beam_ld_added = beam_v_m_ld.copy()

#     def init_ld(df):
#         return {
#             bar_num_ld: df[bar_num],
#             # bar_1st_ld: df[bar_1st],
#             # bar_2nd_ld: df[bar_2nd]
#         }

#     for Loc in BAR.keys():
#         # Loc = Loc.capitalize()

#         bar_num = 'Bar' + Loc + 'Num'
#         ld = Loc + 'Ld'
#         bar_num_ld = bar_num + 'Ld'
#         # bar_1st_ld = bar_1st + 'Ld'
#         # bar_2nd_ld = bar_2nd + 'Ld'

#         beam_ld_added = beam_ld_added.assign(**init_ld(beam_ld_added))

#         count = 0

#         for name, group in beam_ld_added.groupby(['Story', 'BayID'], sort=False):
#             group = group.copy()
#             for i in range(len(group)):
#                 stn_loc = group.at[group.index[i], 'StnLoc']
#                 stn_ld = group.at[group.index[i], ld]
#                 stn_inter = (group['StnLoc'] >= stn_loc -
#                              stn_ld) & (group['StnLoc'] <= stn_loc + stn_ld)
#                 group.loc[stn_inter, bar_num_ld] = np.maximum(
#                     group.at[group.index[i], bar_num], group.loc[stn_inter, bar_num_ld])
#                 # group.loc[group[stn_inter].index, bar_num_ld] = np.maximum(
#                 #     group.at[group.index[i], bar_num], group.loc[group[stn_inter].index, bar_num_ld])

#             beam_ld_added.loc[group.index, bar_num_ld] = group[bar_num_ld]
#             count += 1
#             if count % 100 == 0:
#                 print(name)

#     return beam_ld_added


def _calc_num_length(group, split_array):
    num = np.empty_like(split_array)
    length = np.empty_like(split_array)

    for i in range(len(split_array)):
        num[i] = np.amax(split_array[i])
        length[i] = group.at[split_array[i].index[-1], 'StnLoc'] - group.at[
            split_array[i].index[0], 'StnLoc']
    return num, length


def _make_1st_last_diff(group_diff):
    if group_diff[0] == 0:
        group_diff[0] = 1

    if group_diff[-1] == 0:
        group_diff[-1] = -1

    return group_diff


def _get_min_cut(group_loc, group_loc_diff, i):
    if group_loc_diff[i] > 0:
        return group_loc.index[i]
    else:
        return group_loc.index[i + 1]


def cut_optimization(beam_ld_added, beam_3p):
    rebars = load_e2k()[0]

    # def _calc_num_length(group, split_array):
    #     num = np.empty_like(split_array)
    #     length = np.empty_like(split_array)

    #     for i in range(len(split_array)):
    #         num[i] = np.amax(split_array[i])
    #         length[i] = group.at[split_array[i].index[-1], 'StnLoc'] - group.at[
    #             split_array[i].index[0], 'StnLoc']
    #     return num, length

    # # def concat_num_size(num, group_size):
    # #     if num == 0:
    # #         return 0
    # #     return str(int(num)) + '-' + group_size

    # # def num_to_1st_2nd(num, group_cap):
    # #     if num - group_cap == 1:
    # #         return group_cap - 1, 2
    # #     elif num > group_cap:
    # #         return group_cap, num - group_cap
    # #     else:
    # #         return max(num, 2), 0

    # def _make_1st_last_diff(group_diff):
    #     if group_diff[0] == 0:
    #         group_diff[0] = 1

    #     if group_diff[-1] == 0:
    #         group_diff[-1] = -1

    #     return group_diff

    # def _get_min_cut(group_loc, group_loc_diff, i):
    #     # loc = group_loc_diff[i]
    #     # right = group_loc_diff[i + 1]
    #     # left = group.loc[group_loc.index[i], bar_num_ld]
    #     # right = group.loc[group_loc.index[i + 1], bar_num_ld]
    #     if group_loc_diff[i] > 0:
    #         return group_loc.index[i]
    #     else:
    #         return group_loc.index[i + 1]
    # if left == right:
    #     print(f'ERROR in get_min_cut {i}')
    # if left < right:
    #     return group_loc.index[i]
    # else:
    #     return group_loc.index[i + 1]

    # def make_first_last_diff(*args):
    #     result = []

    #     for arg in args:
    #         arg[0] = 1
    #         arg[-1] = 1
    #         result.append(arg)

    #     return tuple(result)

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

    for Loc in BAR.keys():

        k = output_loc[Loc]['START_LOC']
        to_2nd = output_loc[Loc]['TO_2nd']

        bar_cap = 'Bar' + Loc + 'Cap'
        bar_size = 'Bar' + Loc + 'Size'
        bar_num_ld = 'Bar' + Loc + 'NumLd'

        for name, group in beam_ld_added.groupby(['Story', 'BayID'], sort=False):
            min_usage = float('Inf')

            group_cap = group.at[group.index[0], bar_cap]
            group_size = group.at[group.index[0], bar_size]

            group_max = np.amax(group['StnLoc'])
            group_min = np.amin(group['StnLoc'])

            left = (group_max - group_min) * ITERATION_GAP['Left'] + group_min
            right = (group_max - group_min) * (
                ITERATION_GAP['Right']) + group_min

            group_left = group[bar_num_ld][(
                group['StnLoc'] >= left[0]) & (group['StnLoc'] <= left[1])]
            group_right = group[bar_num_ld][(
                group['StnLoc'] >= right[0]) & (group['StnLoc'] <= right[1])]

            group_left_diff = np.diff(group_left)
            group_right_diff = np.diff(group_right)

            # (group_left_diff, group_right_diff) = make_first_last_diff(
            #     group_left_diff, group_right_diff)

            # group_left_diff = np.concatenate(([1], group_left_diff, [-1]))
            # group_right_diff = np.concatenate(([1], group_right_diff, [-1]))

            group_left_diff = _make_1st_last_diff(group_left_diff)
            group_right_diff = _make_1st_last_diff(group_right_diff)

            # group_left_diff[0] = 1
            # group_left_diff[-1] = -1
            # group_right_diff[0] = 1
            # group_right_diff[-1] = -1

            for i in np.flatnonzero(group_left_diff):
                # for i in range(len(group_left_diff)):
                # if group_left_diff[i] != 0:
                # split_left = group_left.index[i + 1]
                split_left = _get_min_cut(group_left, group_left_diff, i)
                # split_left = group_left.index[i]

                for j in np.flatnonzero(group_right_diff):

                    # for j in range(len(group_right_diff)):
                    # if group_right_diff[j] != 0:
                    # split_3p_array = np.split(
                    #     group[bar_num_ld], [group_left.index[i + 1], group_right.index[j + 1]])
                    split_right = _get_min_cut(
                        group_right, group_right_diff, j)
                    # split_right = group_right.index[j]
                    split_3p_array = [
                        group.loc[:split_left, bar_num_ld], group.loc[split_left: split_right, bar_num_ld], group.loc[split_right:, bar_num_ld]]
                    num, length = _calc_num_length(group, split_3p_array)
                    # num_left = np.amax(a_left)
                    # num_mid = np.amax(a_mid)
                    # num_right = np.amax(a_right)

                    # length_left = group.at[a_left.index[-1], 'StnLoc'] - group.at[a_left.index[0], 'StnLoc']
                    # length_mid = group.at[a_mid.index[-1], 'StnLoc'] - group.at[a_mid.index[0], 'StnLoc']
                    # length_right = group.at[a_right.index[-1], 'StnLoc'] - group.at[a_right.index[0], 'StnLoc']

                    rebar_usage = np.sum(num * length)
                    # rebar_usage = num_left * len(a_left) + num_mid * len(a_mid) + num_right * len(a_right)
                    if rebar_usage < min_usage:
                        min_usage = rebar_usage
                        min_num = num
                        min_length = length
                        # min_num_mid = num_mid
                        # min_num_right = num_right
            # if min_usage == float('Inf'):
            #     min_num = np.full(3, group.at[group.index[0], bar_num_ld])
            #     min_length = np.full(3, '')

            group_num = {
                '左': num_to_1st_2nd(min_num[0], group_cap),
                '中': num_to_1st_2nd(min_num[1], group_cap),
                '右': num_to_1st_2nd(min_num[2], group_cap)
            }

            group_length = {
                '左': min_length[0],
                # '左': min_length[0] if min_num[0] != min_num[1] else '',
                '中': min_length[1],
                '右': min_length[2]
                # '右': min_length[2] if min_num[2] != min_num[1] else ''
            }

            for bar_loc in group_num.keys():
                loc_1st, loc_2nd = group_num[bar_loc]
                loc_length = group_length[bar_loc]
                beam_3p.at[k, ('主筋', bar_loc)] = concat_num_size(
                    loc_1st, group_size)
                beam_3p.at[k, ('長度', bar_loc)] = loc_length * 100
                beam_3p.at[k + to_2nd, ('主筋', bar_loc)
                           ] = concat_num_size(loc_2nd, group_size)

            beam_3p.at[k, ('NOTE', '')] = min_usage * (
                rebars[(group_size, 'AREA')]) * 1000000

            k += 4
            # # x < 1/4 => max >= Spacing => Spacing max
            # group_left_max = np.amax(SPACING[np.amin(group_left) >= SPACING])
            # group_mid_max = np.amax(SPACING[np.amin(group_mid) >= SPACING])
            # group_right_max = np.amax(
            #     SPACING[np.amin(group_right) >= SPACING])

            # beam_design_table.loc[group_left.index.tolist(),
            #                     'SetSpacing'] = group_left_max
            # beam_design_table.loc[group_mid.index.tolist(),
            #                     'SetSpacing'] = group_mid_max
            # beam_design_table.loc[group_right.index.tolist(),
            #                     'SetSpacing'] = group_right_max

            # beam_3points_table.loc[i, ('箍筋', '左')] = (
            #     group_size + str(int(group_left_max * 100)))
            # beam_3points_table.loc[i, ('箍筋', '中')] = (
            #     group_size + str(int(group_mid_max * 100)))
            # beam_3points_table.loc[i, ('箍筋', '右')] = (
            #     group_size + str(int(group_right_max * 100)))

            # i = i + 4

    return beam_3p


def cut_5(beam_ld_added, beam_5):
    rebars = load_e2k()[0]

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

    for Loc in BAR.keys():

        k = output_loc[Loc]['START_LOC']
        to_2nd = output_loc[Loc]['TO_2nd']

        bar_cap = 'Bar' + Loc + 'Cap'
        bar_size = 'Bar' + Loc + 'Size'
        bar_num_ld = 'Bar' + Loc + 'NumLd'

        for _, group in beam_ld_added.groupby(['Story', 'BayID'], sort=False):
            min_usage = float('Inf')

            group_cap = group.at[group.index[0], bar_cap]
            group_size = group.at[group.index[0], bar_size]

            group_max = np.amax(group['StnLoc'])
            group_min = np.amin(group['StnLoc'])

            left = (group_max - group_min) * ITERATION_GAP['Left'] + group_min
            right = (group_max - group_min) * (
                ITERATION_GAP['Right']) + group_min

            group_left = group[bar_num_ld][(
                group['StnLoc'] >= left[0]) & (group['StnLoc'] <= left[1])]
            group_right = group[bar_num_ld][(
                group['StnLoc'] >= right[0]) & (group['StnLoc'] <= right[1])]

            group_left_diff = np.diff(group_left)
            group_right_diff = np.diff(group_right)

            group_left_diff = _make_1st_last_diff(group_left_diff)
            group_right_diff = _make_1st_last_diff(group_right_diff)

            for i in np.flatnonzero(group_left_diff):
                split_left = _get_min_cut(group_left, group_left_diff, i)

                for j in np.flatnonzero(group_right_diff):
                    split_right = _get_min_cut(
                        group_right, group_right_diff, j)
                    split_3p_array = [
                        group.loc[:split_left, bar_num_ld], group.loc[split_left: split_right, bar_num_ld], group.loc[split_right:, bar_num_ld]]
                    num, length = _calc_num_length(group, split_3p_array)

                    rebar_usage = np.sum(num * length)
                    if rebar_usage < min_usage:
                        min_usage = rebar_usage
                        min_num = num
                        min_length = length

            group_num = {
                '左': num_to_1st_2nd(min_num[0], group_cap),
                '中': num_to_1st_2nd(min_num[1], group_cap),
                '右': num_to_1st_2nd(min_num[2], group_cap)
            }

            group_length = {
                '左': min_length[0],
                # '左': min_length[0] if min_num[0] != min_num[1] else '',
                '中': min_length[1],
                '右': min_length[2]
                # '右': min_length[2] if min_num[2] != min_num[1] else ''
            }

            for bar_loc in group_num.keys():
                loc_1st, loc_2nd = group_num[bar_loc]
                loc_length = group_length[bar_loc]
                beam_5.at[k, ('主筋', bar_loc)] = concat_num_size(
                    loc_1st, group_size)
                beam_5.at[k, ('長度', bar_loc)] = loc_length * 100
                beam_5.at[k + to_2nd, ('主筋', bar_loc)
                          ] = concat_num_size(loc_2nd, group_size)

            beam_5.at[k, ('NOTE', '')] = min_usage * (
                rebars[(group_size, 'AREA')]) * 1000000

            k += 4

    return beam_5


def main():
    clock = Clock()
    beam_3p, _ = load_pkl(SCRIPT_DIR + '/stirrups.pkl')
    # beam_v_m = load_pkl(SCRIPT_DIR + '/beam_v_m.pkl')
    beam_ld_added = load_pkl(SCRIPT_DIR + '/beam_ld_added.pkl')

    # start = time.time()
    # clock.time()
    # beam_v_m_ld = calc_ld(beam_v_m)
    # clock.time()
    # # print(time.time() - start)
    # clock.time()
    # beam_ld_added = add_ld(beam_v_m_ld)
    # clock.time()
    # beam_ld_added.to_excel(SCRIPT_DIR + '/beam_ld_added.xlsx')
    # beam_ld_added = load_pkl(SCRIPT_DIR + '/beam_ld_added.pkl', beam_ld_added)
    # beam_ld_added = load_pkl(SCRIPT_DIR + '/beam_ld_added.pkl')
    clock.time()
    beam_3p = cut_optimization(beam_ld_added, beam_3p)
    clock.time()

    # clock.time()
    # a = np.array([1, 2, 3, 4, 5])
    # for i in range(10000):
    #     np.r_[0, a, 4]
    # clock.time()
    # clock.time()
    # a = np.array([1, 2, 3, 4, 5])
    # for i in range(10000):
    #     np.insert(a, 0, 0)
    #     np.append(a, 4)
    # clock.time()
    # clock.time()
    # a = np.array([1, 2, 3, 4, 5])
    # for i in range(10000):
    #     np.concatenate(([0], a, [4]))
    # clock.time()
    beam_3p.to_excel(SCRIPT_DIR + '/beam_3p_opti.xlsx')


if __name__ == '__main__':
    main()
