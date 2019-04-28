""" init beam output table
"""
import pandas as pd
import numpy as np


def init_beam(etabs_design, moment=3):
    """
    init output beam table return beam
    """
    header_info_1 = [('樓層', ''), ('編號', ''), ('RC 梁寬', ''), ('RC 梁深', '')]

    header_rebar_3 = [
        ('主筋', ''), ('主筋', '左'), ('主筋', '中'), ('主筋', '右'),
        ('主筋長度', '左'), ('主筋長度', '中'), ('主筋長度', '右')
    ]
    header_rebar_5 = [
        ('主筋', ''), ('主筋', '左1'), ('主筋', '左2'),
        ('主筋', '中'), ('主筋', '右2'), ('主筋', '右1'),
        ('主筋長度', '左1'), ('主筋長度', '左2'),
        ('主筋長度', '中'), ('主筋長度', '右2'), ('主筋長度', '右1')
    ]

    header_sidebar = [('腰筋', '')]

    header_stirrup = [
        ('箍筋', '左'), ('箍筋', '中'), ('箍筋', '右'),
        ('箍筋長度', '左'), ('箍筋長度', '中'), ('箍筋長度', '右')
    ]

    header_info_2 = [
        ('梁長', ''), ('支承寬', '左'), ('支承寬', '右'),
        ('主筋量', ''), ('箍筋量', '')
    ]

    if moment == 3:
        header = pd.MultiIndex.from_tuples(
            header_info_1 + header_rebar_3 + header_sidebar + header_stirrup + header_info_2)

    elif moment == 5:
        header = pd.MultiIndex.from_tuples(
            header_info_1 + header_rebar_5 + header_sidebar + header_stirrup + header_info_2)

    beam = pd.DataFrame(np.empty([len(etabs_design.groupby(
        ['Story', 'BayID'])) * 4, len(header)], dtype='<U16'), columns=header)

    row = 0
    for (story, bay_id), group in etabs_design.groupby(['Story', 'BayID'], sort=False):
        beam.at[row, '樓層'] = story
        beam.at[row, '編號'] = bay_id
        beam.at[row, 'RC 梁寬'] = group['B'].iloc[0] * 100
        beam.at[row, 'RC 梁深'] = group['H'].iloc[0] * 100
        beam.at[row, '梁長'] = group['Length'].iloc[0] * 100
        beam.at[row, ('支承寬', '左')] = group['LSupportWidth'].iloc[0] * 100
        beam.at[row, ('支承寬', '右')] = group['RSupportWidth'].iloc[0] * 100

        beam.loc[row: row + 3, ('主筋', '')] = [
            '上層 第一排', '上層 第二排', '下層 第二排', '下層 第一排']

        row += 4

    return beam


def put_beam_id(beam, etabs_design):
    """ change bayID to usr defined beam id
    """

    row = 0

    for (_, beam_id), _ in etabs_design.groupby(['Story', 'BeamID'], sort=False):
        beam.at[row, '編號'] = beam_id

        row += 4

    return beam


def main():
    """
    test
    """
    from tests.const import const
    from src.e2k import load_e2k
    from src.etabs_design import load_etabs_design, post_e2k, post_beam_name
    from src.beam_name import load_beam_name

    from src.execution_time import Execution

    execution = Execution()
    execution.time()

    e2k = load_e2k(const['e2k_path'])
    etabs_design = load_etabs_design(const['etabs_design_path'])
    etabs_design = post_e2k(etabs_design, e2k)

    beam = init_beam(etabs_design, moment=3)
    print(beam.head())

    beam = init_beam(etabs_design, moment=5)
    print(beam.head())

    beam_name = load_beam_name(const['beam_name_path'])
    etabs_design = post_beam_name(etabs_design, beam_name)
    beam = put_beam_id(beam, etabs_design)
    print(beam.head())

    execution.time()


if __name__ == "__main__":
    main()
