"""
traditional bar
"""
import numpy as np

from utils.execution_time import Execution
from components.bar_functions import concat_num_size, num_to_1st_2nd

from const import BAR
from data.dataset_rebar import double_area, rebar_area, rebar_db


def cut_traditional(etbas_design, beam):
    """
    traditional cut
    """
    beam = beam.copy()

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
                # beam.at[i, ('長度', bar_loc)] = loc_length * 100
                return loc_ld
            else:
                return span * 1/3
        else:
            return span * 1/5

    for loc in BAR:

        i = output_loc[loc]['START_LOC']
        to_2nd = output_loc[loc]['TO_2nd']

        bar_cap = 'Bar' + loc + 'Cap'
        bar_size = 'Bar' + loc + 'Size'
        bar_num = 'Bar' + loc + 'Num'
        ld = loc + 'SimpleLd'
        # bar_num_ld = bar_num + 'SimpleLd'
        # bar_1st = 'Bar' + loc + '1st'
        # bar_2nd = 'Bar' + loc + '2nd'

        for _, group in etbas_design.groupby(['Story', 'BayID'], sort=False):
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
            #         beam.at[i, ('主筋', bar_loc)] = concat_num_size(
            #             loc_1st, group_size)
            #         beam.at[i + to_2nd, ('主筋', bar_loc)
            #                    ] = concat_num_size(loc_2nd, group_size)

            #         beam.at[i, ('長度', bar_loc)] = loc_length * 100

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

                beam.at[i, ('主筋', bar_loc)] = concat_num_size(
                    loc_1st, group_size)
                beam.at[i + to_2nd, ('主筋', bar_loc)
                        ] = concat_num_size(loc_2nd, group_size)

                beam.at[i, ('長度', bar_loc)] = loc_length * 100

                num_usage = num_usage + loc_num * loc_length

                # for bar_loc in ('左', '中', '右'):

                # total_num = total_num + loc_num
                # if loc_num - cap_num == 1:
                #     beam.at[i, ('主筋', bar_loc)] = concat_num_size(cap_num - 1)
                #     beam.at[i + to_2nd, ('主筋', bar_loc)] = concat_num_size(2)
                # elif loc_num > cap_num:
                #     beam.at[i, ('主筋', bar_loc)] = concat_num_size(cap_num)
                #     beam.at[i + to_2nd, ('主筋', bar_loc)] = concat_num_size(loc_num - cap_num)
                # else:
                #     beam.at[i, ('主筋', bar_loc)] = concat_num_size(loc_num)
                #     beam.at[i + to_2nd, ('主筋', bar_loc)] = 0

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
            beam.at[i, ('NOTE', '')] = num_usage * (
                rebars[(group_size, 'AREA')]) * 1000000

            i += 4

    return beam


def main():
    """
    test
    """

    beam_con = cut_traditional(beam_ld_added, beam)


if __name__ == '__main__':
    main()
