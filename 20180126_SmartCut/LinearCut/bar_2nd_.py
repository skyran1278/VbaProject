import os

import pandas as pd
import numpy as np

from utils.Clock import Clock
from utils.pkl import load_pkl

from dataset.const import BAR, DB_SPACING

from dataset.dataset_beam_design import load_beam_design
from dataset.dataset_e2k import load_e2k
from dataset.dataset_beam_name import load_beam_name

dataset_dir = os.path.dirname(os.path.abspath(__file__))

stirrups_save_file = dataset_dir + '/stirrups.pkl'

clock = Clock()

rebars, stories, point_coordinates, lines, materials, sections = load_e2k()


def _bar_name(Loc):
    bar_size = 'Bar' + Loc + 'Size'
    bar_num = 'Bar' + Loc + 'Num'
    bar_cap = 'Bar' + Loc + 'Cap'
    bar_1st = 'Bar' + Loc + '1st'
    bar_2nd = 'Bar' + Loc + '2nd'

    return (bar_size, bar_num, bar_cap, bar_1st, bar_2nd)


def _calc_bar_size_num(Loc, i):
    bar_size, bar_num, bar_cap, bar_1st, bar_2nd = _bar_name(Loc)

    def calc_capacity(df):
        # dh = df['VSize'].apply()
        # 應該可以用 apply 來改良，晚點再來做
        # 這裡應該拿最後配的來算，但是因為號數整支梁都會相同，所以沒差
        dh = np.array([rebars['#' + v_size.split('#')[1], 'DIA']
                       for v_size in df['VSize']])
        db = rebars[BAR[Loc][i], 'DIA']
        width = np.array([sections[(sec_ID, 'B')] for sec_ID in df['SecID']])
        # cover = np.array([sections[(sec_ID, 'COVER' + Loc)] for sec_ID in df['SecID']])

        return np.ceil((width - 2 * 0.04 - 2 * dh - db) / (DB_SPACING * db + db))

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

    return {
        bar_size: BAR[Loc][i],
        bar_cap: calc_capacity,
        bar_num: lambda x: np.maximum(np.ceil(x['As' + Loc] / rebars[BAR[Loc][i], 'AREA']), 2),
        bar_1st: calc_1st,
        bar_2nd: calc_2nd
    }


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


if __name__ == '__main__':
    (_, beam_v) = load_pkl(stirrups_save_file)
    beam_v_m = calc_db_by_frame(beam_v)
    beam_v_m = load_pkl(dataset_dir + '/beam_v_m.pkl', beam_v_m)
    beam_v_m.to_excel(dataset_dir + '/beam_v_m.xlsx')
