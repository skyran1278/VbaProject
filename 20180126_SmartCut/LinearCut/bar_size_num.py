import os
import sys
import math
import time

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.append(os.path.join(SCRIPT_DIR, os.path.pardir))

from utils.pkl import load_pkl
from utils.Clock import Clock

from dataset.const import BAR, DB_SPACING

from dataset.dataset_beam_design import load_beam_design
from dataset.dataset_e2k import load_e2k
from dataset.dataset_beam_name import load_beam_name

from stirrups import calc_sturrups

# save_file = SCRIPT_DIR + '/3pionts.xlsx'
stirrups_save_file = SCRIPT_DIR + '/stirrups.pkl'

rebars, stories, point_coordinates, lines, materials, sections = load_e2k()
# beam_v = load_beam_design()
# beam_3p = init_beam_3points_table()
# beam_3p, beam_design_table_stirrups = calc_sturrups(beam_3p)
# beam_3points_table = init_beam_3points_table()
# beam_3points_table, beam_design_table_stirrups = calc_sturrups(
#     beam_3points_table)
# (beam_3points_table, beam_design_table_stirrups) = load_pkl(
#     stirrups_save_file, (beam_3points_table, beam_design_table_stirrups))
# (beam_3p, beam_v) = load_pkl(stirrups_save_file)


def _bar_name(Loc):
    bar_size = 'Bar' + Loc + 'Size'
    bar_num = 'Bar' + Loc + 'Num'
    bar_cap = 'Bar' + Loc + 'Cap'
    bar_1st = 'Bar' + Loc + '1st'
    bar_2nd = 'Bar' + Loc + '2nd'

    return (bar_size, bar_num, bar_cap, bar_1st, bar_2nd)


def _calc_bar_size_num(Loc, i):
    # Loc = Loc.capitalize()

    bar_size, bar_num, bar_cap, bar_1st, bar_2nd = _bar_name(Loc)

    def calc_capacity(df):
        # dh = df['VSize'].apply()
        # 應該可以用 apply 來改良，晚點再來做
        # 這裡應該拿最後配的來算，但是因為號數整支梁都會相同，所以沒差
        # 後來查了一下 發現好像差不多
        dh = np.array([rebars['#' + v_size.split('#')[1], 'DIA']
                       for v_size in df['VSize']])
        db = rebars[BAR[Loc][i], 'DIA']
        width = np.array([sections[(sec_ID, 'B')] for sec_ID in df['SecID']])
        # cover = np.array([sections[(sec_ID, 'COVER' + Loc)] for sec_ID in df['SecID']])

        return np.floor((width - 2 * 0.04 - 2 * dh - db) / (DB_SPACING * db + db)) + 1
        # return np.ceil((width - 2 * 0.04 - 2 * dh - db) / (DB_SPACING * db + db))

    def calc_1st(df):
        bar_1st = np.where(df[bar_num] > df[bar_cap],
                           df[bar_cap], df[bar_num])
        bar_1st[df[bar_num] - df[bar_cap] ==
                1] = df[bar_cap][df[bar_num] - df[bar_cap] == 1] - 1

        return bar_1st

    def calc_2nd(df):
        bar_2nd = np.where(df[bar_num] > df[bar_cap],
                           df[bar_num] - df[bar_cap], 0)
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
        bar_size: BAR[Loc][i],
        bar_cap: calc_capacity,
        bar_num: lambda x: np.maximum(np.ceil(x['As' + Loc] / rebars[BAR[Loc][i], 'AREA']), 2),
        bar_1st: calc_1st,
        bar_2nd: calc_2nd
    }


# def calc_capacity(width, cover, dh, db, DB_SPACING):
#     return math.ceil((width - 2 * 0.04 - 2 * dh - db) / (DB_SPACING * db + db))


# def calc_dbt(group, rebars):
#     v_size = group['VSize'].iat[0]
#     v_sizeout_double = '#' + v_size.split('#')[1]
#     return rebars[v_sizeout_double, 'DIA']


def calc_db_by_beam(beam_v):
    beam_v_m = beam_v.copy()

    for Loc in BAR.keys():
        # Loc = Loc.capitalize()

        bar_size, bar_num, bar_cap, bar_1st, bar_2nd = _bar_name(Loc)

        # loc = Loc.lower()
        # Loc = Loc.upper()
        i = 0

        beam_v_m = beam_v_m.assign(**_calc_bar_size_num(Loc, i))

        # beam_v_m.to_excel(save_file)

        # print(beam_v_m.head())

        for _, group in beam_v_m.groupby(['Story', 'BayID'], sort=False):
            i = 0
            # SecID = group['SecID'].iat[0]
            # dh = calc_dbt(group, rebars)
            # db = rebars[BAR[Loc][i], 'DIA']
            # width = sections[(SecID, 'B')]
            # cover = sections[(SecID, 'COVER' + Loc)]
            # capacity = calc_capacity(width, cover, dh, db, DB_SPACING)
            # group = group.assign(**calc_bar_size_num(Loc, i))
            # print(Story, BayID)

            while np.any(group[bar_num] > 2 * group[bar_cap]):
                i += 1
                group = group.assign(**_calc_bar_size_num(Loc, i))
                # db = rebars[BAR[Loc][i], 'DIA']
                # capacity = calc_capacity(width, cover, dh, db, DB_SPACING)
                # print(capacity)

            beam_v_m.loc[group.index.tolist(), [bar_size, bar_num, bar_cap, bar_1st, bar_2nd]
                         ] = group[[bar_size, bar_num, bar_cap, bar_1st, bar_2nd]]
            # print(group)

    return beam_v_m


def calc_db_by_frame(beam_v):
    beam_v_m = _add_beam_name(beam_v)

    for Loc in BAR.keys():
        bar_size, bar_num, bar_cap, bar_1st, bar_2nd = _bar_name(Loc)

        i = 0

        beam_v_m = beam_v_m.assign(**_calc_bar_size_num(Loc, i))

        for _, group in beam_v_m.groupby(['Story', 'FrameID'], sort=False):
            i = 0

            while np.any(group[bar_num] > 2 * group[bar_cap]):
                i += 1
                group = group.assign(**_calc_bar_size_num(Loc, i))

            beam_v_m.loc[group.index, [bar_size, bar_num, bar_cap, bar_1st, bar_2nd]
                         ] = group[[bar_size, bar_num, bar_cap, bar_1st, bar_2nd]]

    return beam_v_m


def _add_beam_name(beam_v):
    beam_name = load_beam_name()
    beam_v = beam_v.assign(BeamID='', FrameID='')

    for (story, bayID), group in beam_v.groupby(['Story', 'BayID'], sort=False):
        beamID, frameID = beam_name.loc[(story, bayID), :]
        group = group.assign(BeamID=beamID, FrameID=frameID)
        beam_v.loc[group.index, ['BeamID', 'FrameID']
                   ] = group[['BeamID', 'FrameID']]

    return beam_v


if __name__ == "__main__":
    clock = Clock()
    (_, beam_v) = load_pkl(stirrups_save_file)
    beam_v_m = calc_db_by_beam(beam_v)
    # beam_v_m = calc_db_by_frame(beam_v)
    beam_v_m = load_pkl(SCRIPT_DIR + '/beam_v_m.pkl', beam_v_m)

# start = time.time()

# print(time.time() - start)
# start = time.time()
# beam_3p_bar = cut_conservative(beam_v_m, beam_3p)
# beam_v_m_ld = calc_ld(beam_v_m)

# beam_v_m_ld = load_pkl(SCRIPT_DIR + '/beam_v_m_ld.pkl')

# start = time.time()

# beam_v_m_add_ld = add_ld(beam_v_m_ld)

# print(time.time() - start)


# start = time.time()
# for i in range(10000):
#     b = beam_v_m_ld['StnLoc'][2]
# print(time.time() - start)

# start = time.time()
# for i in range(10000):
#     a = beam_v_m_ld.loc[beam_v_m_ld.index[2], 'StnLoc']
# print(time.time() - start)

# beam_3p_bar.to_excel(save_file)
# beam_v_m_add_ld.to_excel(SCRIPT_DIR + '/beam_v_m.xlsx')
