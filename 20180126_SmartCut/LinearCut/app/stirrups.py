import os
import sys

import pandas as pd
import numpy as np

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.append(os.path.join(SCRIPT_DIR, os.path.pardir))

from utils.Clock import Clock

from database.dataset_beam_design import load_beam_design
from database.dataset_e2k import load_e2k
from database.const import STIRRUP_REBAR as REBAR, STIRRUP_SPACING as SPACING

from init_table import init_beam

# list to numpy
SPACING = np.array(SPACING) / 100


def first_calc_dbt_spacing(beam_design_table, rebars):
    # first calc VSize to spacing
    return beam_design_table.assign(VSize=REBAR[0], Spacing=lambda x: rebars[REBAR[0], 'AREA'] * 2 / x.VRebar)


def upgrade_size(beam_design_table, rebars):
    # tStart = time.time()
    print('Start upgrade stirrup size...')
    for _, group in beam_design_table.groupby(['Story', 'BayID'], sort=False):
        i = 1

        # if spacing < min => upgrade size and recalculate spcaing
        while np.any(group['Spacing'] < SPACING[0]):
            rebar_num, rebar_size = REBAR[i].split(sep='#')
            rebar_size = '#' + rebar_size

            if rebar_num == '2':
                spacing = rebars[rebar_size, 'AREA'] * 4 / group['VRebar']
            else:
                spacing = rebars[rebar_size, 'AREA'] * 2 / group['VRebar']

            group = group.assign(VSize=REBAR[i], Spacing=spacing)

            i += 1

        beam_design_table.loc[group.index.tolist(), ['VSize', 'Spacing']] = group[[
            'VSize', 'Spacing']]

    # tEnd = time.time()
    # print(tEnd - tStart)
    return beam_design_table


# def drop_size(beam_design_table, rebars):
#     # tStart = time.time()
#     print('Start drop double size...')

#     def get_no_du_size(df):
#         return df['VSize'].apply(lambda x: '#' + x.split(sep='#')[1])

#     beam_design_table = beam_design_table.assign(
#         RealVSize=get_no_du_size, RealSpacing=0)

#     for _, group in beam_design_table.groupby(['Story', 'BayID'], sort=False):
#         i = 1

#         # if spacing < min => upgrade size and recalculate spcaing
#         while np.any(group['Spacing'] < SPACING[0]):
#             rebar_num, rebar_size = REBAR[i].split(sep='#')
#             rebar_size = '#' + rebar_size

#             if rebar_num == '2':
#                 spacing = rebars[rebar_size, 'AREA'] * 4 / group['VRebar']
#             else:
#                 spacing = rebars[rebar_size, 'AREA'] * 2 / group['VRebar']

#             group = group.assign(VSize=REBAR[i], Spacing=spacing)

#             i += 1

#         beam_design_table.loc[group.index.tolist(), ['VSize', 'Spacing']] = group[[
#             'VSize', 'Spacing']]

#     # tEnd = time.time()
#     # print(tEnd - tStart)
#     return beam_design_table


def merge_segments(beam_design_table, beam_3points_table):
    # merge to 3 segments
    # tStart = time.time()
    print('Start merge to 3 segments...')

    # def get_no_du_size(df):
    #     return df['VSize'].apply(lambda x: '#' + x.split(sep='#')[1])

    def drop_size(spacing):
        if (np.amin(spacing) / 2) >= SPACING[0]:
            return rebar_size, spacing / 2
        return group_size, spacing

    def get_spacing(loc_min, loc_max):
        return group['Spacing'][(group['StnLoc'] >= loc_min) & (group['StnLoc'] <= loc_max)]

    beam_design_table = beam_design_table.assign(
        RealVSize='', RealSpacing=0)

    i = 0
    for _, group in beam_design_table.groupby(['Story', 'BayID'], sort=False):
        group_max = np.amax(group['StnLoc'])
        group_min = np.amin(group['StnLoc'])

        # x < 1/4
        left = (group_max - group_min) / 4 + group_min
        # x > 3/4
        right = 3 * (group_max - group_min) / 4 + group_min

        group_size = group['VSize'].iloc[0]
        rebar_num, rebar_size = group_size.split(sep='#')
        rebar_size = '#' + rebar_size

        # group_left = group['Spacing'][group['StnLoc'] <= left]
        # group_mid = group['Spacing'][(
        #     group['StnLoc'] >= left) & (group['StnLoc'] <= right)]
        # group_right = group['Spacing'][group['StnLoc'] >= right]
        # group_left = get_spacing(group_min, left)
        # group_mid = get_spacing(left, right)
        # group_right = get_spacing(right, group_max)

        group_spacing = {
            '左': get_spacing(group_min, left),
            '中': get_spacing(left, right),
            '右': get_spacing(right, group_max)
        }

        for loc in ('左', '中', '右'):
            loc_size = group_size
            loc_spacing = group_spacing[loc]

            if rebar_num == '2':
                loc_size, loc_spacing = drop_size(loc_spacing)

            loc_spacing_max = np.amax(SPACING[np.amin(loc_spacing) >= SPACING])
            beam_design_table.loc[loc_spacing.index,
                                  'RealSpacing'] = loc_spacing_max
            beam_design_table.loc[loc_spacing.index,
                                  'RealVSize'] = loc_size
            beam_3points_table.loc[i, ('箍筋', loc)
                                   ] = f'{loc_size}@{int(loc_spacing_max * 100)}'

        # group_left_size = group_size
        # group_mid_size = group_size
        # group_right_size = group_size

        # if rebar_num == '2':
        #     group_left_size, group_left = drop_size(group_left)
        #     group_mid_size, group_mid = drop_size(group_mid)
        #     group_right_size, group_right = drop_size(group_right)

        # x < 1/4 => max >= Spacing => Spacing max
        # group_left_max = np.amax(SPACING[np.amin(group_left) >= SPACING])
        # group_mid_max = np.amax(SPACING[np.amin(group_mid) >= SPACING])
        # group_right_max = np.amax(
        #     SPACING[np.amin(group_right) >= SPACING])

        # beam_design_table.loc[group_left.index, 'RealSpacing'] = group_left_max
        # beam_design_table.loc[group_mid.index, 'RealSpacing'] = group_mid_max
        # beam_design_table.loc[group_right.index, 'RealSpacing'] = group_right_max
        # beam_design_table.loc[group_left.index.tolist(),
        #                       'RealSpacing'] = group_left_max
        # beam_design_table.loc[group_mid.index.tolist(),
        #                       'RealSpacing'] = group_mid_max
        # beam_design_table.loc[group_right.index.tolist(),
        #                       'RealSpacing'] = group_right_max

        # beam_3points_table.loc[i, ('箍筋', '左')] = f'{group_left_size}@{group_left_max * 100}'
        # beam_3points_table.loc[i, ('箍筋', '中')] = f'{group_mid_size}@{group_mid_max * 100}'
        # beam_3points_table.loc[i, ('箍筋', '右')] = f'{group_right_size}@{group_right_max * 100}'

        i = i + 4

    # tEnd = time.time()
    # print(tEnd - tStart)
    return beam_3points_table, beam_design_table


def calc_sturrups(beam_3points_table):
    rebars = load_e2k()[0]
    beam_design_table = load_beam_design()

    beam_design_table = first_calc_dbt_spacing(beam_design_table, rebars)
    beam_design_table = upgrade_size(beam_design_table, rebars)
    beam_3points_table, beam_design_table = merge_segments(
        beam_design_table, beam_3points_table)

    return beam_3points_table, beam_design_table


def main():
    clock = Clock()

    beam_3points_table = init_beam()
    clock.time()
    beam_3points_table, _ = calc_sturrups(beam_3points_table)
    clock.time()
    print(beam_3points_table.loc[0, ('箍筋', '左')])
    print('Done!')


if __name__ == '__main__':
    main()
