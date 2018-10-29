import os
import re
import time
import pickle

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

from dataset.dataset_beam_design import load_beam_design
from dataset.dataset_e2k import load_e2k
from dataset.const import STIRRUP_REBAR as REBAR, STIRRUP_SPACING as SPACING
from output_table import init_beam_3points_table


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

        beam_design_table.loc[group.index.tolist(), ['VSize', 'Spacing']] = group[['VSize', 'Spacing']]

    # tEnd = time.time()
    # print(tEnd - tStart)
    return beam_design_table


def merge_segments(beam_design_table, beam_3points_table):
    # merge to 3 segments
    # tStart = time.time()
    print('Start merge to 3 segments...')
    i = 0
    for _, group in beam_design_table.groupby(['Story', 'BayID'], sort=False):
        group_max = np.amax(group['StnLoc'])
        group_min = np.amin(group['StnLoc'])

        # x < 1/4
        group_left = (group_max - group_min) / 4 + group_min
        # x > 3/4
        group_right = 3 * (group_max - group_min) / 4 + group_min

        group_size = group['VSize'].iloc[0] + '@'

        # x < 1/4 => max >= Spacing => Spacing max
        group_left_max = np.amax(SPACING[np.amax(
            group['Spacing'][group['StnLoc'] <= group_left]) >= SPACING]) * 100
        group_mid_max = np.amax(
            SPACING[np.amax(group['Spacing'][(group['StnLoc'] >= group_left) & (group['StnLoc'] <= group_right)]) >= SPACING]) * 100
        group_right_max = np.amax(SPACING[np.amax(
            group['Spacing'][group['StnLoc'] >= group_right]) >= SPACING]) * 100

        beam_3points_table.loc[i, ('箍筋', '左')] = (
            group_size + str(int(group_left_max)))
        beam_3points_table.loc[i, ('箍筋', '中')] = (
            group_size + str(int(group_mid_max)))
        beam_3points_table.loc[i, ('箍筋', '右')] = (
            group_size + str(int(group_right_max)))

        i = i + 4

    # tEnd = time.time()
    # print(tEnd - tStart)
    return beam_3points_table


def calc_sturrups(beam_3points_table):
    rebars, _, _, _, _, _ = load_e2k()
    beam_design_table = load_beam_design()

    beam_design_table = first_calc_dbt_spacing(beam_design_table, rebars)
    beam_design_table = upgrade_size(beam_design_table, rebars)
    beam_3points_table = merge_segments(beam_design_table, beam_3points_table)

    return beam_3points_table, beam_design_table


def main():
    beam_3points_table = init_beam_3points_table()
    beam_3points_table, beam_design_table = calc_sturrups(beam_3points_table)
    print(beam_3points_table.loc[0, ('箍筋', '左')])
    print('Done!')


if __name__ == '__main__':
    main()
