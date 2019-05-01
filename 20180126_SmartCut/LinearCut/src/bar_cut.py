"""
smart cut
"""
from itertools import combinations, product

import numpy as np
from scipy.signal import argrelextrema


from src.bar_functions import concat_num_size, num_to_1st_2nd
from src.rebar import rebar_area


def cut_multiple(df, col, boundary, group_num=5):
    """
    multiple cut
    """
    if group_num <= 3:
        return cut_3(df, col, boundary)

    # initial
    min_usage = float('Inf')
    num = np.empty(group_num)
    length = np.empty(group_num)

    amin = df['StnLoc'].min()
    amax = df['StnLoc'].max()

    boundarys = (
        (amax - amin) * boundary['left'][0] + amin,
        (amax - amin) * boundary['right'][1] + amin,
    )

    combination_area = (
        (df['StnLoc'] > boundarys[0]) &
        (df['StnLoc'] < boundarys[1])
    )

    diff_area = (
        (df[col].diff() != 0) |
        (df[col].diff().shift(-1) != 0)
    )

    idxs = (
        df.index[combination_area][0],
        *df.index[diff_area & combination_area],
        df.index[combination_area][-1]
    )

    for idx in combinations(idxs, group_num - 1):
        num[0] = df.loc[:idx[0], col].max()
        length[0] = df.loc[idx[0], 'StnLoc'] - df['StnLoc'].min()
        num[-1] = df.loc[idx[-1]:, col].max()
        length[-1] = df['StnLoc'].max() - df.loc[idx[-1], 'StnLoc']

        for i, j in enumerate(range(1, len(idx))):
            id0, id1 = idx[i], idx[j]
            num[j] = df.loc[id0:id1, col].max()
            length[j] = df.loc[id1, 'StnLoc'] - df.loc[id0, 'StnLoc']

        usage = np.sum(num * length)

        if usage < min_usage:
            min_usage = usage
            min_num = num.copy()
            min_length = length.copy()

    print(df.head(1))
    # for local maxima
    # for local minima
    # input should be numpy array
    # return tuple
    incompatible = ((min_usage == float('Inf')) | (
        (argrelextrema(min_num, np.greater)[0].size > 0) &
        (argrelextrema(min_num, np.less)[0].size > 0)
    ))

    if incompatible:
        print('\n\nINCOMPATIBLE\n\n')
        return cut_multiple(df, col, boundary, group_num - 1)

    return min_num, min_length, min_usage


def cut_3(df, col, boundary):
    """
    cut 3, depands on boundary, ex: 0.1~0.45, 0.55~0.9
    """
    # initial
    idxs = {}
    min_usage = float('Inf')
    num = np.empty(3)
    length = np.empty(3)

    amin = df['StnLoc'].min()
    amax = df['StnLoc'].max()

    for index in boundary:
        boundarys = (amax - amin) * boundary[index] + amin

        boundary_area = (
            (df['StnLoc'] >= boundarys[0]) &
            (df['StnLoc'] <= boundarys[1])
        )

        diff_area = (
            (df[col].diff() != 0) |
            (df[col].diff().shift(-1) != 0)
        )

        idxs[index] = (
            df.index[boundary_area][0],
            *df.index[diff_area & boundary_area],
            df.index[boundary_area][-1]
        )

    for idx in product(idxs['left'], idxs['right']):
        num[0] = df.loc[:idx[0], col].max()
        num[-1] = df.loc[idx[-1]:, col].max()
        length[0] = df.loc[idx[0], 'StnLoc'] - df['StnLoc'].min()
        length[-1] = df['StnLoc'].max() - df.loc[idx[-1], 'StnLoc']

        for i, j in enumerate(range(1, len(idx))):
            id0, id1 = idx[i], idx[j]
            num[j] = df.loc[id0:id1, col].max()
            length[j] = df.loc[id1, 'StnLoc'] - df.loc[id0, 'StnLoc']

        usage = np.sum(num * length)

        if usage < min_usage:
            min_usage = usage
            # by reference, so deep copy
            min_num = num.copy()
            min_length = length.copy()

    return min_num, min_length, min_usage


def cut_optimization(beam, etabs_design, const, group_num=3):
    """
    cut 3 or 5, optimization
    """
    rebar = const['rebar']

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

    for loc in rebar:
        row = output_loc[loc]['START_LOC']
        to_2nd = output_loc[loc]['TO_2nd']

        for _, group in etabs_design.groupby(['Story', 'BayID'], sort=False):
            output_num = {}
            output_length = {}
            # group capacity and size
            cap = group.at[group.index[0], 'Bar' + loc + 'Cap']
            size = group.at[group.index[0], 'Bar' + loc + 'Size']

            num, length, usage = cut_multiple(
                group, 'Bar' + loc + 'NumLd', const['boundary'], group_num)

            # else:
            if len(num) % 2 == 1:
                mid = len(num) // 2
                output_num['中'] = num_to_1st_2nd(num[mid], cap)
                output_length['中'] = length[mid]

            for i in range(len(num) // 2):
                output_num[f'左{i+1}'] = num_to_1st_2nd(num[i], cap)
                output_length[f'左{i+1}'] = length[i]
                output_num[f'右{i+1}'] = num_to_1st_2nd(num[-(i+1)], cap)
                output_length[f'右{i+1}'] = length[-(i+1)]

            for col in output_num:
                loc_1st, loc_2nd = output_num[col]
                loc_length = output_length[col]
                beam.at[row, ('主筋', col)] = concat_num_size(
                    loc_1st, size)
                beam.at[row, ('主筋長度', col)] = round(loc_length * 100, 3)
                beam.at[row + to_2nd, ('主筋', col)
                        ] = concat_num_size(loc_2nd, size)

            beam.at[row, ('主筋量', '')] = usage * rebar_area(size) * 1000000

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
    beam = init_beam(etabs_design)
    beam, etabs_design = calc_stirrups(beam, etabs_design, const)
    etabs_design = calc_db('BayID', etabs_design, const)
    etabs_design = calc_ld(etabs_design, const)
    etabs_design = add_ld(etabs_design, 'Ld', const['rebar'])

    execution.time('cut 3')
    # beam = output_3(beam, etabs_design, const)
    beam = cut_optimization(beam, etabs_design, const, 3)
    print(beam.head())
    beam = cut_optimization(beam, etabs_design, const, 5)
    print(beam.head())
    execution.time('cut 3')


if __name__ == '__main__':
    main()
