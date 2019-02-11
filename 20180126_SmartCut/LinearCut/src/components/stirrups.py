""" calc stirrups
"""
# import os
# import sys

# import pandas as pd
import numpy as np

# SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
# sys.path.append(os.path.join(SCRIPT_DIR, os.path.pardir))


from data.dataset_rebar import double_area
from const import STIRRUP_REBAR, STIRRUP_SPACING

# from app.init_table import init_beam

# change m to cm
SPACING = STIRRUP_SPACING / 100


def _first_calc_dbt_spacing(etabs_design):
    # first calc VSize to spacing
    return etabs_design.assign(VSize=STIRRUP_REBAR[0], Spacing=(
        lambda x: double_area(STIRRUP_REBAR[0]) / x.VRebar))


def _upgrade_size(etabs_design):
    print('Start upgrade stirrup size...')

    for _, group in etabs_design.groupby(['Story', 'BayID'], sort=False):
        i = 1

        # if spacing < min => upgrade size and recalculate spcaing
        while np.any(group['Spacing'] < SPACING[0]):
            rebar_num, rebar_size = STIRRUP_REBAR[i].split(sep='#')
            rebar_size = '#' + rebar_size

            if rebar_num == '2':
                # double stirrups so double * 2
                spacing = double_area(rebar_size) * 2 / group['VRebar']
            else:
                spacing = double_area(rebar_size) / group['VRebar']

            group = group.assign(VSize=STIRRUP_REBAR[i], Spacing=spacing)

            i += 1

        etabs_design.loc[group.index, ['VSize', 'Spacing']] = group[[
            'VSize', 'Spacing']]

    return etabs_design


def _drop_size(rebar_size, spacing):
    if (np.amin(spacing) / 2) >= SPACING[0]:
        return rebar_size[1:], spacing / 2
    return rebar_size, spacing


def _get_spacing(group, loc_min, loc_max):
    return group['Spacing'][(group['StnLoc'] >= loc_min) & (group['StnLoc'] <= loc_max)]


def _merge_segments(etabs_design, beam):
    print('Start merge to 3 segments...')

    etabs_design = etabs_design.assign(RealVSize='', RealSpacing=0)

    i = 0
    for _, group in etabs_design.groupby(['Story', 'BayID'], sort=False):
        group_max = np.amax(group['StnLoc'])
        group_min = np.amin(group['StnLoc'])

        # x < 1/4
        left = (group_max - group_min) * 1/4 + group_min
        # x > 3/4
        right = (group_max - group_min) * 3/4 + group_min

        # rebar size with double
        rebar_size = group['VSize'].iloc[0]

        # spacing depands on loc_min, loc_max
        group_spacing = {
            '左': _get_spacing(group, group_min, left),
            '中': _get_spacing(group, left, right),
            '右': _get_spacing(group, right, group_max)
        }

        for loc in ('左', '中', '右'):
            loc_size = rebar_size
            loc_spacing = group_spacing[loc]

            # if double, judge size can drop or not
            if rebar_size[0] == '2':
                loc_size, loc_spacing = _drop_size(loc_size, loc_spacing)

            # all spacing reduce to usr defined
            loc_spacing_max = np.amax(SPACING[np.amin(loc_spacing) >= SPACING])

            # for next convinience get
            etabs_design.loc[loc_spacing.index,
                             'RealSpacing'] = loc_spacing_max
            etabs_design.loc[loc_spacing.index,
                             'RealVSize'] = loc_size

            beam.loc[i, ('箍筋', loc)
                     ] = f'{loc_size}@{int(loc_spacing_max * 100)}'

        i = i + 4

    return beam, etabs_design


def calc_stirrups(etabs_design, beam):
    """ calc stirrups depands on
    """
    etabs_design = _first_calc_dbt_spacing(etabs_design)
    etabs_design = _upgrade_size(etabs_design)
    etabs_design, beam = _merge_segments(etabs_design, beam)

    return etabs_design, beam


if __name__ == "__main__":
    from utils.execution_time import Execution
