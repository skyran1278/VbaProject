""" calc stirrups
"""
import numpy as np

from src.dataset_rebar import double_area


def _calc_vc(df, e2k):
    # pylint: disable=invalid-name
    # pylint: disable=no-member

    materials, sections = e2k['materials'], e2k['sections']

    B = df['SecID'].apply(lambda x: sections[x, 'B'])
    H = df['SecID'].apply(lambda x: sections[x, 'D'])
    # D = H - 0.1  # 假設雙排 #10 箍筋 #4
    material = df['SecID'].apply(lambda x: sections[x, 'MATERIAL'])
    fc = material.apply(lambda x: materials[x, 'FC'])
    fyt = material.apply(lambda x: materials[x, 'FY'])

    amin = df.groupby(['Story', 'BayID'])['StnLoc'].transform('min')
    amax = df.groupby(['Story', 'BayID'])['StnLoc'].transform('max')

    seismic_area = np.maximum((amax - amin) / 4, 2 * H)

    new_av = np.maximum.reduce([
        df['VRebar'] - 0.53 * np.sqrt(fc) * B / fyt,
        0.2 * np.sqrt(fc) * B / fyt,
        3.5 * B / fyt
    ])

    df['VRebarConsiderVc'] = np.where(
        (
            (df['StnLoc'] > seismic_area + amin) &
            (df['StnLoc'] < amax - seismic_area)
        ),
        new_av,
        df['VRebar']
    )

    return df


def _first_calc_dbt_spacing(etabs_design, stirrup_rebar, vrebar):
    # print('Start calculate stirrup spacing and size...')
    # first calc VSize to spacing
    return etabs_design.assign(VSize=stirrup_rebar[0], Spacing=(
        lambda x: double_area(stirrup_rebar[0]) / x[vrebar]))


def _upgrade_size(etabs_design, stirrup_rebar, stirrup_spacing, vrebar):
    # print('Start upgrade stirrup size...')

    for _, group in etabs_design.groupby(['Story', 'BayID'], sort=False):
        loc = 1

        # if spacing < min => upgrade size and recalculate spcaing
        while np.any(group['Spacing'] < stirrup_spacing[0]):
            rebar_num, rebar_size = stirrup_rebar[loc].split(sep='#')
            rebar_size = '#' + rebar_size

            if rebar_num == '2':
                # double stirrups so double * 2
                spacing = double_area(rebar_size) * 2 / group[vrebar]
            else:
                spacing = double_area(rebar_size) / group[vrebar]

            group = group.assign(VSize=stirrup_rebar[loc], Spacing=spacing)

            loc += 1

        etabs_design.loc[group.index, ['VSize', 'Spacing']] = group[[
            'VSize', 'Spacing']]

    return etabs_design


def _drop_size(rebar_size, spacing, stirrup_spacing):
    if (np.amin(spacing) / 2) >= stirrup_spacing[0]:
        return rebar_size[1:], spacing / 2
    return rebar_size, spacing


def _get_spacing(group, loc_min, loc_max):
    return group['Spacing'][(group['StnLoc'] >= loc_min) & (group['StnLoc'] <= loc_max)]


def _merge_segments(beam, etabs_design, stirrup_spacing):
    # print('Start merge to 3 segments...')

    etabs_design = etabs_design.assign(RealVSize='', RealSpacing=0)

    row = 0
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

        group_length = {
            '左': (group_max - group_min) * 1/4,
            '中': (group_max - group_min) * 2/4,
            '右': (group_max - group_min) * 1/4
        }

        for loc in ('左', '中', '右'):
            loc_size = rebar_size
            loc_spacing = group_spacing[loc]

            # if double, judge size can drop or not
            if rebar_size[0] == '2':
                loc_size, loc_spacing = _drop_size(
                    loc_size, loc_spacing, stirrup_spacing)

            # all spacing reduce to usr defined
            loc_spacing_max = np.amax(
                stirrup_spacing[np.amin(loc_spacing) >= stirrup_spacing])

            # for next convinience get
            etabs_design.loc[loc_spacing.index,
                             'RealSpacing'] = loc_spacing_max

            # windows: UnicodeEncodeError so add .encode('utf-8', 'ignore').decode('utf-8')
            # remove numpy array, use default array instead
            etabs_design.loc[loc_spacing.index,
                             'RealVSize'] = loc_size

            beam.loc[row, ('箍筋', loc)] = (
                f'{loc_size}@{int(loc_spacing_max * 100)}'
            )

            beam.loc[row, ('箍筋長度', loc)] = group_length[loc] * 100

        row = row + 4

    return beam, etabs_design


def _cut_3(beam, etabs_design, stirrup_spacing):
    print('Start merge to 3 segments...')

    etabs_design = etabs_design.assign(RealVSize='', RealSpacing=0)

    row = 0
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

        group_length = {
            '左': (group_max - group_min) * 1/4,
            '中': (group_max - group_min) * 2/4,
            '右': (group_max - group_min) * 1/4
        }

        for loc in ('左', '中', '右'):
            loc_size = rebar_size
            loc_spacing = group_spacing[loc]

            # if double, judge size can drop or not
            if rebar_size[0] == '2':
                loc_size, loc_spacing = _drop_size(
                    loc_size, loc_spacing, stirrup_spacing)

            # all spacing reduce to usr defined
            loc_spacing_max = np.amax(
                stirrup_spacing[np.amin(loc_spacing) >= stirrup_spacing])

            # for next convinience get
            etabs_design.loc[loc_spacing.index,
                             'RealSpacing'] = loc_spacing_max

            # windows: UnicodeEncodeError so add .encode('utf-8', 'ignore').decode('utf-8')
            # remove numpy array, use default array instead
            etabs_design.loc[loc_spacing.index,
                             'RealVSize'] = loc_size

            beam.loc[row, ('箍筋', loc)] = (
                f'{loc_size}@{int(loc_spacing_max * 100)}'
            )

            beam.loc[row, ('箍筋長度', loc)] = group_length[loc] * 100

        row = row + 4

    return beam, etabs_design


def calc_stirrups(beam, etabs_design, e2k, const, consider_vc=False):
    """ calc stirrups
    """
    vrebar = 'VRebarConsiderVc' if consider_vc else 'VRebar'

    stirrup_rebar = const['stirrup_rebar']
    stirrup_spacing = const['stirrup_spacing']

    # change m to cm
    stirrup_spacing = stirrup_spacing / 100

    etabs_design = _calc_vc(etabs_design, e2k)

    etabs_design = _first_calc_dbt_spacing(etabs_design, stirrup_rebar, vrebar)
    etabs_design = _upgrade_size(
        etabs_design, stirrup_rebar, stirrup_spacing, vrebar)
    beam, etabs_design = _merge_segments(beam, etabs_design, stirrup_spacing)

    return beam, etabs_design


def _main():
    from src.const import const
    from src.init_beam import init_beam
    from src.dataset_e2k import load_e2k
    from src.dataset_etabs_design import load_beam_design
    from src.execution_time import Execution

    e2k_path, etabs_design_path = const['e2k_path'], const['etabs_design_path']

    e2k = load_e2k(e2k_path, e2k_path + '.pkl')
    etabs_design = load_beam_design(
        etabs_design_path, etabs_design_path + '.pkl')

    beam = init_beam(etabs_design, e2k, moment=3)
    execution = Execution()
    execution.time('Stirrup Time')
    beam, dh_design = calc_stirrups(
        beam, etabs_design, e2k, const, consider_vc=False)
    print(beam.head())
    print(dh_design.head())
    execution.time()


if __name__ == "__main__":
    _main()
