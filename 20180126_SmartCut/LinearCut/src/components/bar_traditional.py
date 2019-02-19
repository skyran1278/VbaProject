"""
traditional bar
"""
import numpy as np

from components.bar_functions import concat_num_size, num_to_1st_2nd

from data.dataset_rebar import rebar_area


def _get_group_length(group_num, group, ld):
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
        end length depends on simple ld and 1/3, if ld too long, then get max rebar and length is 1/3
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
        ld = loc + 'SimpleLd'

        for _, group in etbas_design.groupby(['Story', 'BayID'], sort=False):
            num_usage = 0

            group_max = np.amax(group['StnLoc'])
            group_min = np.amin(group['StnLoc'])

            group_cap = group.at[group.index[0], bar_cap]
            group_size = group.at[group.index[0], bar_size]

            group_num = {
                '左': _get_group_num(group, 0, 1/3),
                '中': _get_group_num(group, 1/4, 3/4),
                '右': _get_group_num(group, 2/3, 1)
            }

            group_length = _get_group_length(group_num, group, ld)

            for bar_loc in ('左', '中', '右'):
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

                beam.at[row, ('長度', bar_loc)] = loc_length * 100

                num_usage = num_usage + loc_num * loc_length

            # 計算鋼筋體積 cm3
            beam.at[row, ('NOTE', '')] = num_usage * rebar_area(
                group_size) * 1000000

            row += 4

    return beam


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
    from components.bar_ld import calc_ld

    e2k_path, etabs_design_path = const[
        'e2k_path'], const['etabs_design_path']

    e2k = load_e2k(e2k_path, e2k_path + '.pkl')
    etabs_design = load_beam_design(
        etabs_design_path, etabs_design_path + '.pkl')

    beam = init_beam(etabs_design, e2k, moment=3, shear=True)
    execution = Execution()
    beam, dh_design = calc_stirrups(beam, etabs_design, const)

    db_design = calc_db('BayID', dh_design, e2k, const)

    ld_design = calc_ld(db_design, e2k, const)

    execution.time('cut traditional')
    beam_trational = cut_traditional(beam, ld_design, const['rebar'])
    print(beam_trational.head())
    execution.time('cut traditional')


if __name__ == '__main__':
    main()
