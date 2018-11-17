import os
import sys
import time

import pandas as pd
import numpy as np

from dataset.dataset_e2k import load_e2k
from utils.pkl import load_pkl
from dataset.const import BAR

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.append(os.path.join(SCRIPT_DIR, os.path.pardir))


def cut_conservative(beam_v_m, beam_3p):
    rebars, stories, point_coordinates, lines, materials, sections = load_e2k()
    beam_3p = beam_3p.copy()

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

    def get_group_num(min_loc, max_loc):
        group_loc_min = (group_max - group_min) * min_loc + group_min
        group_loc_max = (group_max - group_min) * max_loc + group_min

        max_index = group[bar_num][(group['StnLoc'] >= group_loc_min) & (group['StnLoc'] <= group_loc_max)].idxmax()

        return (group.at[max_index, bar_1st] + group.at[max_index, bar_2nd]), group.at[max_index, bar_1st], group.at[max_index, bar_2nd]

    def concat_size(num):
        if num == 0:
            return 0
        return str(int(num)) + '-' + group_size

    for Loc in BAR.keys():

        i = output_loc[Loc]['START_LOC']
        to_2nd = output_loc[Loc]['TO_2nd']

        bar_size = 'Bar' + Loc + 'Size'
        bar_num = 'Bar' + Loc + 'Num'
        bar_1st = 'Bar' + Loc + '1st'
        bar_2nd = 'Bar' + Loc + '2nd'

        for _, group in beam_v_m.groupby(['Story', 'BayID'], sort=False):
            total_num = 0

            group_max = np.amax(group['StnLoc'])
            group_min = np.amin(group['StnLoc'])

            # group_left = (group_max - group_min) * 1 / 3 + group_min

            # group_mid_min = (group_max - group_min) * 1 / 4 + group_min
            # group_mid_max = (group_max - group_min) * 3 / 4 + group_min

            # group_right = (group_max - group_min) * 2 / 3 + group_min

            # cap_num = group[bar_cap].iloc[0]
            group_size = group.at[group.index[0], bar_size]

            group_num = {
                '左': get_group_num(0, 1/3),
                '中': get_group_num(1/4, 3/4),
                '右': get_group_num(2/3, 1)
            }

            for bar_loc in ('左', '中', '右'):
                loc_num, loc_1st, loc_2nd = group_num[bar_loc]
                beam_3p.at[i, ('主筋', bar_loc)] = concat_size(loc_1st)
                beam_3p.at[i + to_2nd, ('主筋', bar_loc)] = concat_size(loc_2nd)

                total_num = total_num + loc_num
                # if loc_num - cap_num == 1:
                #     beam_3p.at[i, ('主筋', bar_loc)] = concat_size(cap_num - 1)
                #     beam_3p.at[i + to_2nd, ('主筋', bar_loc)] = concat_size(2)
                # elif loc_num > cap_num:
                #     beam_3p.at[i, ('主筋', bar_loc)] = concat_size(cap_num)
                #     beam_3p.at[i + to_2nd, ('主筋', bar_loc)] = concat_size(loc_num - cap_num)
                # else:
                #     beam_3p.at[i, ('主筋', bar_loc)] = concat_size(loc_num)
                #     beam_3p.at[i + to_2nd, ('主筋', bar_loc)] = 0

            # 計算鋼筋體積 cm3
            beam_3p.at[i, ('NOTE', '')] = total_num * (group_max - group_min) / 3 * (
                rebars[(group_size, 'AREA')]) * 1000000

            i += 4

    return beam_3p


def main():
    start = time.time()

    (beam_3p, _) = load_pkl(SCRIPT_DIR + '/stirrups.pkl')
    beam_v_m = load_pkl(SCRIPT_DIR + '/beam_v_m.pkl')

    beam_3p_con = cut_conservative(beam_v_m, beam_3p)

    beam_3p_con.to_excel(SCRIPT_DIR + '/beam_3p_con.xlsx')

    print(time.time() - start)


if __name__ == '__main__':
    main()
