"""
traditional bar
"""
import numpy as np

from src.bar_functions import concat_num_size, num_to_1st_2nd

from src.rebar import rebar_area


def _get_group_length(group_num, group, ld):  # pylint: disable=invalid-name
    span = np.amax(group['StnLoc']) - np.amin(group['StnLoc'])

    left_num = group_num['左1'][0]
    mid_num = group_num['中'][0]
    right_num = group_num['右1'][0]

    left_ld = group.at[group.index[0], ld]
    right_ld = group.at[group.index[-1], ld]

    # 如果有需要，這裡或許可以加上無條件進位的函數
    left_length = _get_loc_length(left_num, left_ld, mid_num, span)
    right_length = _get_loc_length(right_num, right_ld, mid_num, span)

    mid_length = span - left_length - right_length

    return {
        '左1': left_length,
        '中': mid_length,
        '右1': right_length
    }


def _get_loc_length(loc_num, loc_ld, mid_num, span):
    if loc_num > mid_num:
        if loc_ld > span * 1/3:
            return loc_ld
        return span * 1/3
    # because mid > end, so no need extend ld
    return span * 1/5


def cut_traditional(beam, etbas_design, rebar):
    """
    traditional cut

    algorithm:
        cut in 0~1/3, 1/4~3/4, 2/3~1 to get max bar number
        cut in 1/3, 1/5 depends on bar number, but don't have 1/7
        end length depends on simple ld and 1/3,
        if ld too long, then get max rebar and length is 1/3
    """
    beam = beam.copy()

    output_loc = {
        'Top': {
            'start_loc': 0,
            'to_2nd': 1
        },
        'Bot': {
            'start_loc': 3,
            'to_2nd': -1
        }
    }

    def _get_group_num(group, min_loc, max_loc):
        group_loc_min = (group_max - group_min) * min_loc + group_min
        group_loc_max = (group_max - group_min) * max_loc + group_min

        num = np.amax(group[bar_num][(group['StnLoc'] >= group_loc_min) & (
            group['StnLoc'] <= group_loc_max)])

        num_1st, num_2nd = num_to_1st_2nd(num, group_cap)

        return num, num_1st, num_2nd

    for loc in rebar:
        row = output_loc[loc]['start_loc']
        to_2nd = output_loc[loc]['to_2nd']

        bar_cap = 'Bar' + loc + 'Cap'
        bar_size = 'Bar' + loc + 'Size'
        bar_num = 'Bar' + loc + 'Num'
        ld = loc + 'SimpleLd'  # pylint: disable=invalid-name

        for _, group in etbas_design.groupby(['Story', 'BayID'], sort=False):
            num_usage = 0

            group_max = np.amax(group['StnLoc'])
            group_min = np.amin(group['StnLoc'])

            group_cap = group.at[group.index[0], bar_cap]
            group_size = group.at[group.index[0], bar_size]

            group_num = {
                '左1': _get_group_num(group, 0, 1/3),
                '中': _get_group_num(group, 1/4, 3/4),
                '右1': _get_group_num(group, 2/3, 1)
            }

            group_length = _get_group_length(group_num, group, ld)

            for bar_loc in ('中', '左1', '右1'):
                loc_num, loc_1st, loc_2nd = group_num[bar_loc]
                loc_length = group_length[bar_loc]

                # if mid < 0, get max num and length = 1/3
                if group_length['中'] <= 0:
                    span = np.amax(group['StnLoc']) - np.amin(group['StnLoc'])

                    bar_max = max(group_num, key=group_num.get)
                    loc_num, loc_1st, loc_2nd = group_num[bar_max]
                    loc_length = span / 3

                beam.at[row, ('主筋', bar_loc)] = concat_num_size(
                    loc_1st, group_size)
                beam.at[row + to_2nd, ('主筋', bar_loc)
                        ] = concat_num_size(loc_2nd, group_size)

                beam.at[row, ('主筋長度', bar_loc)] = round(loc_length * 100, 3)

                num_usage = num_usage + loc_num * loc_length

            # 計算鋼筋體積 cm3
            beam.at[row, ('主筋量', '')] = (
                num_usage * rebar_area(group_size) * 1000000)

            row += 4

    return beam


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

    execution.time('cut traditional')
    beam_trational = cut_traditional(beam, etabs_design, const['rebar'])
    print(beam_trational.head())
    execution.time('cut traditional')


if __name__ == '__main__':
    main()
