import os
import sys
import time

import pandas as pd
import numpy as np

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.append(os.path.join(SCRIPT_DIR, os.path.pardir))

from utils.pkl import load_pkl
from utils.Clock import Clock
from utils.functions import concat_num_size, num_to_1st_2nd

from dataset.const import BAR
from dataset.dataset_e2k import load_e2k
from bar_opti import calc_ld


def add_simple_ld(beam_v_m_ld):
    def init_ld(df):
        return {
            bar_num_ld: df[bar_num],
            # bar_1st_ld: df[bar_1st],
            # bar_2nd_ld: df[bar_2nd]
        }

    for Loc in BAR.keys():

        # Loc = Loc.capitalize()

        bar_num = 'Bar' + Loc + 'Num'
        ld = Loc + 'SimpleLd'
        bar_num_ld = bar_num + 'SimpleLd'
        # bar_1st_ld = bar_1st + 'Ld'
        # bar_2nd_ld = bar_2nd + 'Ld'

        if not ld in beam_v_m_ld.columns:
            beam_v_m_ld = calc_ld(beam_v_m_ld)

        beam_v_m_ld = beam_v_m_ld.assign(**init_ld(beam_v_m_ld))

        for _, group in beam_v_m_ld.groupby(['Story', 'BayID'], sort=False):
            group = group.copy()
            for i in (0, -1):
                stn_loc = group.at[group.index[i], 'StnLoc']
                stn_ld = group.at[group.index[i], ld]
                stn_inter = (group['StnLoc'] >= stn_loc -
                             stn_ld) & (group['StnLoc'] <= stn_loc + stn_ld)
                group.loc[stn_inter, bar_num_ld] = np.maximum(
                    group.at[group.index[i], bar_num], group.loc[stn_inter, bar_num_ld])
                # group.loc[group[stn_inter].index, bar_num_ld] = np.maximum(
                #     group.at[group.index[i], bar_num], group.loc[group[stn_inter].index, bar_num_ld])

            beam_v_m_ld.loc[group.index, bar_num_ld] = group[bar_num_ld]
            # print(name)

    return beam_v_m_ld


def cut_conservative(beam_v_m, beam_3p):
    rebars = load_e2k()[0]
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

    # def num_to_1st_2nd(num, group_cap):
    #     if num - group_cap == 1:
    #         return group_cap - 1, 2
    #     elif num > group_cap:
    #         return group_cap, num - group_cap
    #     else:
    #         return max(num, 2), 0

    def get_group_num(min_loc, max_loc):
        group_loc_min = (group_max - group_min) * min_loc + group_min
        group_loc_max = (group_max - group_min) * max_loc + group_min

        # max_index = group[bar_num_ld][(group['StnLoc'] >= group_loc_min) & (
        #     group['StnLoc'] <= group_loc_max)].idxmax()

        num = np.amax(group[bar_num][(group['StnLoc'] >= group_loc_min) & (
            group['StnLoc'] <= group_loc_max)])

        num_1st, num_2nd = num_to_1st_2nd(num, group_cap)

        return num, num_1st, num_2nd
        # return (group.at[max_index, bar_1st] + group.at[max_index, bar_2nd]), group.at[max_index, bar_1st], group.at[max_index, bar_2nd]

    # def concat_num_size(num):
    #     if num == 0:
    #         return 0
    #     return str(int(num)) + '-' + group_size

    # def get_num_usage(loc_num, mid_num, span):
    #     # if loc_num is None:
    #     #     return mid_num * span * 1/3
    #     if loc_num < mid_num:
    #         return loc_num * span * 1/5 + mid_num * span * 2/15
    #     else:
    #         loc_ld = group.at[group.index[0], bar_num_ld]
    #         if loc_ld > span * 1/3:
    #             return loc_num * loc_ld
    #         else:
    #             return loc_num * span * 1/3

    def get_group_length(group_num, group, ld):
        span = np.amax(group['StnLoc']) - np.amin(group['StnLoc'])

        left_num = group_num['左'][0]
        mid_num = group_num['中'][0]
        right_num = group_num['右'][0]

        left_ld = group.at[group.index[0], ld]
        right_ld = group.at[group.index[-1], ld]

        # 如果有需要，這裡或許可以加上無條件進位的函數
        left_length = _get_loc_length(left_num, left_ld, mid_num, span)
        right_length = _get_loc_length(right_num, right_ld, mid_num, span)

        mid_length = span - left_length - right_length

        return {
            '左': left_length,
            '中': mid_length,
            '右': right_length
        }

    def _get_loc_length(loc_num, loc_ld, mid_num, span):
        if loc_num > mid_num:
            if loc_ld > span * 1/3:
                # beam_3p.at[i, ('長度', bar_loc)] = loc_length * 100
                return loc_ld
            else:
                return span * 1/3
        else:
            return span * 1/5

    for Loc in BAR.keys():

        i = output_loc[Loc]['START_LOC']
        to_2nd = output_loc[Loc]['TO_2nd']

        bar_cap = 'Bar' + Loc + 'Cap'
        bar_size = 'Bar' + Loc + 'Size'
        bar_num = 'Bar' + Loc + 'Num'
        ld = Loc + 'SimpleLd'
        # bar_num_ld = bar_num + 'SimpleLd'
        # bar_1st = 'Bar' + Loc + '1st'
        # bar_2nd = 'Bar' + Loc + '2nd'

        for _, group in beam_v_m.groupby(['Story', 'BayID'], sort=False):
            num_usage = 0

            group_max = np.amax(group['StnLoc'])
            group_min = np.amin(group['StnLoc'])

            # group_left = (group_max - group_min) * 1 / 3 + group_min

            # group_mid_min = (group_max - group_min) * 1 / 4 + group_min
            # group_mid_max = (group_max - group_min) * 3 / 4 + group_min

            # group_right = (group_max - group_min) * 2 / 3 + group_min

            # cap_num = group[bar_cap].iloc[0]
            group_cap = group.at[group.index[0], bar_cap]
            group_size = group.at[group.index[0], bar_size]

            group_num = {
                '左': get_group_num(0, 1/3),
                '中': get_group_num(1/4, 3/4),
                '右': get_group_num(2/3, 1)
            }

            group_length = get_group_length(group_num, group, ld)

            # if group_length['中'] <= 0:
            #     span = np.amax(group['StnLoc']) - np.amin(group['StnLoc'])

            #     bar_max = max(group_num, key=group_num.get)
            #     loc_num, loc_1st, loc_2nd = group_num[bar_max]
            #     loc_length = span / 3

            #     for bar_loc in ('左', '中', '右'):
            #         beam_3p.at[i, ('主筋', bar_loc)] = concat_num_size(
            #             loc_1st, group_size)
            #         beam_3p.at[i + to_2nd, ('主筋', bar_loc)
            #                    ] = concat_num_size(loc_2nd, group_size)

            #         beam_3p.at[i, ('長度', bar_loc)] = loc_length * 100

            #         num_usage = num_usage + loc_num * loc_length

            # else:
            for bar_loc in ('左', '中', '右'):
                # for bar_loc in ('左', '右'):
                loc_num, loc_1st, loc_2nd = group_num[bar_loc]
                loc_length = group_length[bar_loc]
                if group_length['中'] <= 0:
                    span = np.amax(group['StnLoc']) - np.amin(group['StnLoc'])

                    bar_max = max(group_num, key=group_num.get)
                    loc_num, loc_1st, loc_2nd = group_num[bar_max]
                    loc_length = span / 3

                beam_3p.at[i, ('主筋', bar_loc)] = concat_num_size(
                    loc_1st, group_size)
                beam_3p.at[i + to_2nd, ('主筋', bar_loc)
                           ] = concat_num_size(loc_2nd, group_size)

                beam_3p.at[i, ('長度', bar_loc)] = loc_length * 100

                num_usage = num_usage + loc_num * loc_length

                # for bar_loc in ('左', '中', '右'):

                # total_num = total_num + loc_num
                # if loc_num - cap_num == 1:
                #     beam_3p.at[i, ('主筋', bar_loc)] = concat_num_size(cap_num - 1)
                #     beam_3p.at[i + to_2nd, ('主筋', bar_loc)] = concat_num_size(2)
                # elif loc_num > cap_num:
                #     beam_3p.at[i, ('主筋', bar_loc)] = concat_num_size(cap_num)
                #     beam_3p.at[i + to_2nd, ('主筋', bar_loc)] = concat_num_size(loc_num - cap_num)
                # else:
                #     beam_3p.at[i, ('主筋', bar_loc)] = concat_num_size(loc_num)
                #     beam_3p.at[i + to_2nd, ('主筋', bar_loc)] = 0

            # 沒有處理 1/7，所以比較保守

            # left_num = group_num['左'][0]

            # right_num = group_num['右'][0]

            # num_usage = num_usage + get_num_usage(left_num, mid_num, span)
            # num_usage = num_usage + get_num_usage(right_num, mid_num, span)
            # num_usage = num_usage + get_num_usage(mid_num=mid_num, span=span)

            # if left_num < mid_num:
            #     num_usage = num_usage + left_num * span * 1/5 + mid_num * span * 2/15
            # else:
            #     num_usage = num_usage + left_num * span * 1/3

            # if right_num < mid_num:
            #     num_usage = num_usage + right_num * span * 1/5 + mid_num * span * 2/15
            # else:
            #     num_usage = num_usage + right_num * span * 1/3

            # 計算鋼筋體積 cm3
            beam_3p.at[i, ('NOTE', '')] = num_usage * (
                rebars[(group_size, 'AREA')]) * 1000000

            i += 4

    return beam_3p


def main():
    start = time.time()

    (beam_3p, _) = load_pkl(SCRIPT_DIR + '/stirrups.pkl')
    # beam_v_m = load_pkl(SCRIPT_DIR + '/beam_v_m.pkl')
    beam_ld_added = load_pkl(SCRIPT_DIR + '/beam_ld_added.pkl')

    beam_ld_added = add_simple_ld(beam_v_m)
    beam_3p_con = cut_conservative(beam_ld_added, beam_3p)

    beam_3p_con.to_excel(SCRIPT_DIR + '/beam_3p_con.xlsx')

    print(time.time() - start)


if __name__ == '__main__':
    main()
