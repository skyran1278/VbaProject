import os
import math
import time

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

from dataset.dataset_beam_design import load_beam_design
from dataset.dataset_e2k import load_e2k
from dataset.const import TOP_BAR, BOT_BAR, DB_SPACING
from stirrups import calc_sturrups
from output_table import init_beam_3points_table
from utils.pkl import load_pkl, create_pkl

dataset_dir = os.path.dirname(os.path.abspath(__file__))
save_file = dataset_dir + '/3pionts.xlsx'
stirrups_save_file = dataset_dir + '/stirrups.pkl'

rebars, stories, point_coordinates, lines, materials, sections = load_e2k()
# beam_with_v = load_beam_design()
# beam_3p = init_beam_3points_table()
# beam_3p, beam_design_table_with_stirrups = calc_sturrups(beam_3p)
(beam_3p, beam_with_v) = load_pkl(stirrups_save_file)

BAR = {
    'TOP': TOP_BAR,
    'BOT': BOT_BAR
}


def bar_name(Loc):
    bar_size = 'Bar' + Loc + 'Size'
    bar_num = 'Bar' + Loc + 'Num'
    bar_cap = 'Bar' + Loc + 'Cap'
    bar_1st = 'Bar' + Loc + '1st'
    bar_2nd = 'Bar' + Loc + '2nd'

    return (bar_size, bar_num, bar_cap, bar_1st, bar_2nd)


def calc_bar_size_num(LOC, i):
    Loc = LOC.capitalize()

    bar_size, bar_num, bar_cap, bar_1st, bar_2nd = bar_name(Loc)

    def calc_capacity(df):
        dh = np.array([rebars['#' + v_size.split('#')[1], 'DIA'] for v_size in df['VSize']])
        db = rebars[BAR[LOC][i], 'DIA']
        width = np.array([sections[(sec_ID, 'B')] for sec_ID in df['SecID']])
        # cover = np.array([sections[(sec_ID, 'COVER' + LOC)] for sec_ID in df['SecID']])

        return np.ceil((width - 2 * 0.04 - 2 * dh - db) / (DB_SPACING * db + db))

    def calc_1st(df):
        bar_1st = np.where(df[bar_num] > df[bar_cap], df[bar_cap], np.maximum(df[bar_num], 2))
        bar_1st[df[bar_num] - df[bar_cap] == 1] = df[bar_cap][df[bar_num] - df[bar_cap] == 1] - 1

        return bar_1st

    def calc_2nd(df):
        bar_2nd = np.where(df[bar_num] > df[bar_cap], df[bar_num] - df[bar_cap], 0)
        bar_2nd[df[bar_num] - df[bar_cap] == 1] = 2

        return bar_2nd
        # for i in range(len(df.index)):
        #     if num:
        #         pass
        # df[bar_num] - df[bar_cap] == 1
        #     beam_3p.loc[i, ('主筋', bar_loc)] = concat_size(cap_num - 1)
        #     beam_3p.loc[i + to_2nd, ('主筋', bar_loc)] = concat_size(2)
        # elif loc_num > cap_num:
        #     beam_3p.loc[i, ('主筋', bar_loc)] = concat_size(cap_num)
        #     beam_3p.loc[i + to_2nd, ('主筋', bar_loc)] = concat_size(loc_num - cap_num)
        # else:
        #     beam_3p.loc[i, ('主筋', bar_loc)] = concat_size(loc_num)
        #     beam_3p.loc[i + to_2nd, ('主筋', bar_loc)] = 0

    return {
        bar_size: BAR[LOC][i],
        bar_num: lambda x: np.ceil(x['As' + Loc] / rebars[BAR[LOC][i], 'AREA']),
        bar_cap: calc_capacity,
        bar_1st: calc_1st,
        bar_2nd: calc_2nd
    }


# def calc_capacity(width, cover, dh, db, DB_SPACING):
#     return math.ceil((width - 2 * 0.04 - 2 * dh - db) / (DB_SPACING * db + db))


# def calc_dbt(group, rebars):
#     v_size = group['VSize'].iat[0]
#     v_size_without_double = '#' + v_size.split('#')[1]
#     return rebars[v_size_without_double, 'DIA']


def calc_db_by_a_beam(beam_with_v):
    for LOC in BAR.keys():
        Loc = LOC.capitalize()

        bar_size, bar_num, bar_cap, bar_1st, bar_2nd = bar_name(Loc)

        # loc = LOC.lower()
        # LOC = LOC.upper()
        i = 0

        beam_with_v = beam_with_v.assign(**calc_bar_size_num(LOC, i))

        # beam_with_v.to_excel(save_file)

        # print(beam_with_v.head())

        for _, group in beam_with_v.groupby(['Story', 'BayID'], sort=False):
            i = 0
            # SecID = group['SecID'].iat[0]
            # dh = calc_dbt(group, rebars)
            # db = rebars[BAR[LOC][i], 'DIA']
            # width = sections[(SecID, 'B')]
            # cover = sections[(SecID, 'COVER' + LOC)]
            # capacity = calc_capacity(width, cover, dh, db, DB_SPACING)
            # group = group.assign(**calc_bar_size_num(LOC, i))
            # print(Story, BayID)

            while np.any(group[bar_num] > 2 * group[bar_cap]):
                i += 1
                group = group.assign(**calc_bar_size_num(LOC, i))
                # db = rebars[BAR[LOC][i], 'DIA']
                # capacity = calc_capacity(width, cover, dh, db, DB_SPACING)
                # print(capacity)

            beam_with_v.loc[group.index.tolist(), [bar_size, bar_num, bar_cap, bar_1st, bar_2nd]
                            ] = group[[bar_size, bar_num, bar_cap, bar_1st, bar_2nd]]
            # print(group)

    return beam_with_v


def cut_conservative(beam_with_v_m, beam_3p):
    output_loc = {
        'TOP': {
            'START_LOC': 0,
            'TO_2nd': 1
        },
        'BOT': {
            'START_LOC': 3,
            'TO_2nd': -1
        }
    }

    def get_group_num(min_loc, max_loc):
        group_loc_min = (group_max - group_min) * min_loc + group_min
        group_loc_max = (group_max - group_min) * max_loc + group_min

        max_index = group[bar_num][(group['StnLoc'] >= group_loc_min) & (group['StnLoc'] <= group_loc_max)].idxmax()

        return group.at[max_index, bar_1st], group.at[max_index, bar_2nd]

    def concat_size(num):
        if num == 0:
            return 0
        return str(int(num)) + '-' + group_size

    for LOC in BAR.keys():
        Loc = LOC.capitalize()

        i = output_loc[LOC]['START_LOC']
        to_2nd = output_loc[LOC]['TO_2nd']

        bar_size, bar_num, bar_cap, bar_1st, bar_2nd = bar_name(Loc)

        for _, group in beam_with_v_m.groupby(['Story', 'BayID'], sort=False):
            group_max = np.amax(group['StnLoc'])
            group_min = np.amin(group['StnLoc'])

            # group_left = (group_max - group_min) * 1 / 3 + group_min

            # group_mid_min = (group_max - group_min) * 1 / 4 + group_min
            # group_mid_max = (group_max - group_min) * 3 / 4 + group_min

            # group_right = (group_max - group_min) * 2 / 3 + group_min

            # cap_num = group[bar_cap].iloc[0]
            group_size = group[bar_size].iloc[0]

            group_num = {
                '左': get_group_num(0, 1/3),
                '中': get_group_num(1/4, 3/4),
                '右': get_group_num(2/3, 1)
            }

            for bar_loc in ('左', '中', '右'):
                loc_1st, loc_2nd = group_num[bar_loc]
                beam_3p.loc[i, ('主筋', bar_loc)] = concat_size(loc_1st)
                beam_3p.loc[i + to_2nd, ('主筋', bar_loc)] = concat_size(loc_2nd)
                # if loc_num - cap_num == 1:
                #     beam_3p.loc[i, ('主筋', bar_loc)] = concat_size(cap_num - 1)
                #     beam_3p.loc[i + to_2nd, ('主筋', bar_loc)] = concat_size(2)
                # elif loc_num > cap_num:
                #     beam_3p.loc[i, ('主筋', bar_loc)] = concat_size(cap_num)
                #     beam_3p.loc[i + to_2nd, ('主筋', bar_loc)] = concat_size(loc_num - cap_num)
                # else:
                #     beam_3p.loc[i, ('主筋', bar_loc)] = concat_size(loc_num)
                #     beam_3p.loc[i + to_2nd, ('主筋', bar_loc)] = 0

            i += 4

    return beam_3p


def calc_ld(beam_with_v_m):
    pass


start = time.time()

beam_with_v_m = calc_db_by_a_beam(beam_with_v)
# print(time.time() - start)
# start = time.time()
beam_3p_with_bar = cut_conservative(beam_with_v_m, beam_3p)

print(time.time() - start)
beam_3p_with_bar.to_excel(save_file)
beam_with_v_m.to_excel(dataset_dir + '/beam_v_m.xlsx')
