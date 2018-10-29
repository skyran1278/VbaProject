import os
import math

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


def calc_bar_size_num(LOC, i):
    Loc = LOC.capitalize()

    bar_size = 'Bar' + Loc + 'Size'
    bar_num = 'Bar' + Loc + 'Num'
    bar_cap = 'Bar' + Loc + 'Cap'

    def calc_capacity(df):
        dbt = np.array([rebars['#' + v_size.split('#')[1], 'DIA'] for v_size in df['VSize']])
        db = rebars[BAR[LOC][i], 'DIA']
        width = np.array([sections[(sec_ID, 'B')] for sec_ID in df['SecID']])
        # cover = np.array([sections[(sec_ID, 'COVER' + LOC)] for sec_ID in df['SecID']])

        return np.ceil((width - 2 * 0.04 - 2 * dbt - db) / (DB_SPACING * db + db))

    return {
        bar_size: BAR[LOC][i],
        bar_num: lambda x: x['As' + Loc] / rebars[BAR[LOC][i], 'AREA'],
        bar_cap: calc_capacity
    }


def calc_capacity(width, cover, dbt, db, DB_SPACING):
    return math.ceil((width - 2 * 0.04 - 2 * dbt - db) / (DB_SPACING * db + db))


def calc_dbt(group, rebars):
    v_size = group['VSize'].iat[0]
    v_size_without_double = '#' + v_size.split('#')[1]
    return rebars[v_size_without_double, 'DIA']


def main():
    for LOC in BAR.keys():
        Loc = LOC.capitalize()
        # loc = LOC.lower()
        # LOC = LOC.upper()
        i = 0
        bar_size = 'Bar' + Loc + 'Size'
        bar_num = 'Bar' + Loc + 'Num'
        bar_cap = 'Bar' + Loc + 'Cap'

        beam_with_v = beam_with_v.assign(**calc_bar_size_num(LOC, i))

        # beam_with_v.to_excel(save_file)

        # print(beam_with_v.head())

        for (Story, BayID), group in beam_with_v.groupby(['Story', 'BayID'], sort=False):
            i = 0
            # SecID = group['SecID'].iat[0]
            # dbt = calc_dbt(group, rebars)
            # db = rebars[BAR[LOC][i], 'DIA']
            # width = sections[(SecID, 'B')]
            # cover = sections[(SecID, 'COVER' + LOC)]
            # capacity = calc_capacity(width, cover, dbt, db, DB_SPACING)
            group = group.assign(**calc_bar_size_num(LOC, i))
            # print(Story, BayID)

            while np.any(group[bar_num] > 2 * group[bar_cap]):
                i += 1
                group = group.assign(**calc_bar_size_num(LOC, i))
                # db = rebars[BAR[LOC][i], 'DIA']
                # capacity = calc_capacity(width, cover, dbt, db, DB_SPACING)
                # print(capacity)

            beam_with_v.loc[group.index.tolist(), [bar_size, bar_num]] = group[[bar_size, bar_num]]
            # print(group)


main()
# beam_with_v.to_excel(save_file)
