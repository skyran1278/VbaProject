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

# from app.init_table import init_beam

# list to numpy
SPACING = np.array(SPACING) / 100


def first_calc_dbt_spacing(beam_design_table, rebars, REBAR=REBAR):
    # first calc VSize to spacing
    return beam_design_table.assign(VSize=REBAR[0], Spacing=lambda x: rebars[REBAR[0], 'AREA'] * 2 / x.VRebar)


def upgrade_size(beam_design, rebars, REBAR=REBAR, SPACING=SPACING):
    print('Start upgrade stirrup size...')
    for _, group in beam_design.groupby(['Story', 'BayID'], sort=False):
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

        beam_design.loc[group.index.tolist(), ['VSize', 'Spacing']] = group[[
            'VSize', 'Spacing']]

    return beam_design


def merge_segments(beam_3, beam_design):
    print('Start merge to 3 segments...')

    def drop_size(spacing):
        if (np.amin(spacing) / 2) >= SPACING[0]:
            return rebar_size, spacing / 2
        return group_size, spacing

    def get_spacing(loc_min, loc_max):
        return group['Spacing'][(group['StnLoc'] >= loc_min) & (group['StnLoc'] <= loc_max)]

    beam_design = beam_design.assign(
        RealVSize='', RealSpacing=0)

    i = 0
    for _, group in beam_design.groupby(['Story', 'BayID'], sort=False):
        group_max = np.amax(group['StnLoc'])
        group_min = np.amin(group['StnLoc'])

        # x < 1/4
        left = (group_max - group_min) / 4 + group_min
        # x > 3/4
        right = 3 * (group_max - group_min) / 4 + group_min

        group_size = group['VSize'].iloc[0]
        rebar_num, rebar_size = group_size.split(sep='#')
        rebar_size = '#' + rebar_size

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
            beam_design.loc[loc_spacing.index,
                            'RealSpacing'] = loc_spacing_max
            beam_design.loc[loc_spacing.index,
                            'RealVSize'] = loc_size
            beam_3.loc[i, ('箍筋', loc)
                       ] = f'{loc_size}@{int(loc_spacing_max * 100)}'

        i = i + 4

    return beam_3, beam_design


def calc_sturrups(beam_3):
    rebars = load_e2k()[0]
    beam_design = load_beam_design()

    beam_design = first_calc_dbt_spacing(beam_design, rebars)
    beam_design = upgrade_size(beam_design, rebars)
    beam_3, beam_design = merge_segments(beam_3, beam_design)

    return beam_3, beam_design
