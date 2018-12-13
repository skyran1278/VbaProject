import os
import sys
import math

import pandas as pd
import numpy as np

from database.dataset_beam_design import load_beam_design
from database.dataset_e2k import load_e2k

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.append(os.path.join(SCRIPT_DIR, os.path.pardir))

e2k = load_e2k()
beam_design = load_beam_design()


def _basic_information(header, beam_design, e2k):
    _, _, point_coordinates, lines, _, sections = e2k

    beam = pd.DataFrame(
        np.empty([len(beam_design.groupby(['Story', 'BayID'])) * 4, len(header)], dtype='<U16'), columns=header)

    i = 0
    for (story, bayID), group in beam_design.groupby(['Story', 'BayID'], sort=False):
        # print(group['StnLoc'])
        beam.at[i, '樓層'] = story
        beam.at[i, '編號'] = bayID
        beam.at[i, 'RC 梁寬'] = sections[(group['SecID'].iloc[0], 'B')] * 100
        beam.at[i, 'RC 梁深'] = sections[(group['SecID'].iloc[0], 'D')] * 100

        point_start = lines[(bayID, 'BEAM', 'START')]
        point_end = lines[(bayID, 'BEAM', 'END')]
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


def init_beam(multi=3, beam_design=beam_design, e2k=e2k):
    if multi == 3:
        header = pd.MultiIndex.from_tuples([('樓層', ''), ('編號', ''), ('RC 梁寬', ''), ('RC 梁深', ''), ('主筋', ''), ('主筋', '左'), ('主筋', '中'), ('主筋', '右'), (
            '長度', '左'), ('長度', '中'), ('長度', '右'), ('腰筋', ''), ('箍筋', '左'), ('箍筋', '中'), ('箍筋', '右'), ('梁長', ''), ('支承寬', '左'), ('支承寬', '右'), ('NOTE', '')])
    elif multi == 5:
        header = pd.MultiIndex.from_tuples([('樓層', ''), ('編號', ''), ('RC 梁寬', ''), ('RC 梁深', ''), ('主筋', ''), ('主筋', '左1'), ('主筋', '左2'), ('主筋', '中'), ('主筋', '右2'), ('主筋', '右1'), (
            '長度', '左1'), ('長度', '左2'), ('長度', '中'), ('長度', '右2'), ('長度', '右1'), ('腰筋', ''), ('箍筋', '左'), ('箍筋', '中'), ('箍筋', '右'), ('梁長', ''), ('支承寬', '左'), ('支承寬', '右'), ('NOTE', '')])

    return _basic_information(header, beam_design, e2k)


def change_to_beamID(beam):
    i = 0

    for (_, beamID), _ in beam_design.groupby(['Story', 'BeamID'], sort=False):
        beam.at[i, '編號'] = beamID

        i = i + 4

    return beam


def init_beam_name(beam_design):
    (story, bayID) = zip(*[(story, bayID)
                           for (story, bayID), _ in beam_design.groupby(['Story', 'BayID'], sort=False)])

    beam_name = pd.DataFrame({
        '樓層': story,
        'ETABS 編號': bayID,
        '施工圖編號': '',
        '一台梁': ''
    })

    return beam_name
