import os
import math

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

from dataset.dataset_beam_design import load_beam_design
from dataset.dataset_e2k import load_e2k

dataset_dir = os.path.dirname(os.path.abspath(__file__))
save_file = dataset_dir + '/3pionts.xlsx'


def init_beam_3points_table():
    rebars, stories, point_coordinates, lines, materials, sections = load_e2k()
    beam_design_table = load_beam_design()

    index = pd.MultiIndex.from_tuples([('樓層', ''), ('編號', ''), ('RC 梁寬', ''), ('RC 梁深', ''), ('主筋', ''), ('主筋', '左'), ('主筋', '中'), ('主筋', '右'), (
        '長度', '左'), ('長度', '中'), ('長度', '右'), ('腰筋', ''), ('箍筋', '左'), ('箍筋', '中'), ('箍筋', '右'), ('梁長', ''), ('支承寬', '左'), ('支承寬', '右'), ('NOTE', ''), ('MESSAGE', '')])

    beam_3points_table = pd.DataFrame(
        np.empty([len(beam_design_table.groupby(['Story', 'BayID'])) * 4, len(index)], dtype='<U16'), columns=index)

    i = 0
    for (story, bayID), group in beam_design_table.groupby(['Story', 'BayID'], sort=False):
        # print(group['StnLoc'])
        beam_3points_table.at[i, '樓層'] = story
        beam_3points_table.at[i, '編號'] = bayID
        beam_3points_table.at[i, 'RC 梁寬'] = sections[(group['SecID'].iloc[0], 'B')] * 100
        beam_3points_table.at[i, 'RC 梁深'] = sections[(group['SecID'].iloc[0], 'D')] * 100

        point_start = lines[(bayID, 'BEAM', 'START')]
        point_end = lines[(bayID, 'BEAM', 'END')]
        beam_length = math.sqrt(sum((point_coordinates.loc[point_end] - point_coordinates.loc[point_start]) ** 2))

        beam_3points_table.at[i, '梁長'] = round(beam_length, 2) * 100
        beam_3points_table.at[i, ('支承寬', '左')] = round(np.amin(group['StnLoc']), 3) * 100
        beam_3points_table.at[i, ('支承寬', '右')] = round((beam_length - np.amax(group['StnLoc'])), 3) * 100

        beam_3points_table.loc[i: i + 3, ('主筋', '')] = ['上層 第一排', '上層 第二排', '下層 第二排', '下層 第一排']

        i = i + 4

    return beam_3points_table


def main():
    beam_3points_table = init_beam_3points_table()
    print(beam_3points_table.head())
    beam_3points_table.to_excel(save_file)
    print('Done!')


if __name__ == '__main__':
    main()
