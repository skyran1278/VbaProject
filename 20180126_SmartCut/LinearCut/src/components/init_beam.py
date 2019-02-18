""" init beam output table
"""
# import os
# import sys
import math

import pandas as pd
import numpy as np

# from database.dataset_beam_design import load_beam_design
# from database.dataset_e2k import load_e2k

# SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
# sys.path.append(os.path.join(SCRIPT_DIR, os.path.pardir))

# e2k = load_e2k()
# etabs_design = load_beam_design()


def _basic_information(header, etabs_design, e2k):
    point_coordinates = e2k['point_coordinates']
    lines = e2k['lines']
    sections = e2k['sections']

    beam = pd.DataFrame(np.empty([len(etabs_design.groupby(
        ['Story', 'BayID'])) * 4, len(header)], dtype='<U16'), columns=header)

    i = 0
    for (story, bay_id), group in etabs_design.groupby(['Story', 'BayID'], sort=False):
        # print(group['StnLoc'])
        beam.at[i, '樓層'] = story
        beam.at[i, '編號'] = bay_id
        beam.at[i, 'RC 梁寬'] = sections[(group['SecID'].iloc[0], 'B')] * 100
        beam.at[i, 'RC 梁深'] = sections[(group['SecID'].iloc[0], 'D')] * 100

        point_start = lines[(bay_id, 'BEAM', 'START')]
        point_end = lines[(bay_id, 'BEAM', 'END')]
        beam_length = math.sqrt(
            sum((point_coordinates.loc[point_end] - point_coordinates.loc[point_start]) ** 2))

        beam.at[i, '梁長'] = round(beam_length, 2) * 100
        beam.at[i, ('支承寬', '左')] = round(np.amin(group['StnLoc']), 3) * 100
        beam.at[i, ('支承寬', '右')] = round(
            (beam_length - np.amax(group['StnLoc'])), 3) * 100

        beam.loc[i: i + 3, ('主筋', '')] = ['上層 第一排',
                                          '上層 第二排', '下層 第二排', '下層 第一排']

        i = i + 4

    return beam


def init_beam(etabs_design, e2k, moment=3, shear=False):
    """
    init output beam table return beam
    """
    header_info_1 = [('樓層', ''), ('編號', ''), ('RC 梁寬', ''), ('RC 梁深', '')]

    # header_rebar = [('主筋', ''), ('主筋', '左'), ('主筋', '中'), ('主筋', '右')]
    header_rebar_3 = [('主筋', ''), ('主筋', '左'), ('主筋', '中'),
                      ('主筋', '右'), ('主筋長度', '左'), ('主筋長度', '中'), ('主筋長度', '右')]
    header_rebar_5 = [('主筋', ''), ('主筋', '左1'), ('主筋', '左2'), ('主筋', '中'),
                      ('主筋', '右2'), ('主筋', '右1'), ('主筋長度', '左1'), ('主筋長度', '左2'),
                      ('主筋長度', '中'), ('主筋長度', '右2'), ('主筋長度', '右1')]

    header_sidebar = [('腰筋', '')]

    header_stirrup = [('箍筋', '左'), ('箍筋', '中'), ('箍筋', '右')]
    header_stirrup_3 = [('箍筋', '左'), ('箍筋', '中'), ('箍筋', '右'),
                        ('箍筋長度', '左'), ('箍筋長度', '中'), ('箍筋長度', '右')]

    header_info_2 = [('梁長', ''), ('支承寬', '左'), ('支承寬', '右'), ('NOTE', '')]

    if moment == 3:
        if shear:
            header = pd.MultiIndex.from_tuples(
                header_info_1 + header_rebar_3 + header_sidebar + header_stirrup_3 + header_info_2)

        else:
            header = pd.MultiIndex.from_tuples(
                header_info_1 + header_rebar_3 + header_sidebar + header_stirrup + header_info_2)

    elif moment == 5:
        if shear:
            header = pd.MultiIndex.from_tuples(
                header_info_1 + header_rebar_5 + header_sidebar + header_stirrup_3 + header_info_2)

        else:
            header = pd.MultiIndex.from_tuples(
                header_info_1 + header_rebar_5 + header_sidebar + header_stirrup + header_info_2)

    return _basic_information(header, etabs_design, e2k)


def add_and_alter_beam_id(beam, beam_name, etabs_design):
    """
    first add beam/frame name id to etabs_design
    second change bayID to usr defined beam id
    """

    etabs_design = _add_usr_beam_name(beam_name, etabs_design)
    beam = _alter_beam_id(beam, etabs_design)

    return beam, etabs_design


def _alter_beam_id(beam, etabs_design):
    """ change bayID to usr defined beam id
    """

    i = 0

    for (_, beam_id), _ in etabs_design.groupby(['Story', 'BeamID'], sort=False):
        beam.at[i, '編號'] = beam_id

        i = i + 4

    return beam


def _add_usr_beam_name(beam_name, etabs_design):
    """ add beam/frame name id to etabs_design
    """
    etabs_design = etabs_design.assign(BeamID='', FrameID='')

    for (story, bay_id), group in etabs_design.groupby(['Story', 'BayID'], sort=False):
        beam_id, frame_id = beam_name.loc[(story, bay_id), :]
        group = group.assign(BeamID=beam_id, FrameID=frame_id)
        etabs_design.loc[group.index, ['BeamID', 'FrameID']
                         ] = group[['BeamID', 'FrameID']]

    return etabs_design


def init_beam_name(etabs_design):
    """ create beam name table
    """
    (story, bay_id) = zip(*[(story, bay_id) for (story, bay_id),
                            _ in etabs_design.groupby(['Story', 'BayID'], sort=False)])

    beam_name = pd.DataFrame({
        '樓層': story,
        'ETABS 編號': bay_id,
        '施工圖編號': '',
        '一台梁': ''
    })

    return beam_name


def main():
    """
    test
    """
    from const import E2K_PATH, ETABS_DESIGN_PATH, BEAM_NAME_PATH
    from data.dataset_e2k import load_e2k
    from data.dataset_etabs_design import load_beam_design
    from data.dataset_beam_name import load_beam_name

    e2k = load_e2k(E2K_PATH, E2K_PATH + '.pkl')
    etabs_design = load_beam_design(
        ETABS_DESIGN_PATH, ETABS_DESIGN_PATH + '.pkl')
    beam_name = load_beam_name(BEAM_NAME_PATH, BEAM_NAME_PATH + '.pkl')

    beam = init_beam(etabs_design, e2k, moment=3, shear=True)
    print(beam.head())

    beam_name_empty = init_beam_name(etabs_design)
    print(beam_name_empty.head())

    beam, etabs_design = add_and_alter_beam_id(
        beam, beam_name, etabs_design)
    print(beam.head())
    print(etabs_design.head())


if __name__ == "__main__":
    main()
