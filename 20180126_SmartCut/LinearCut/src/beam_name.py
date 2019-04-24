""" load user defined beam name
"""
import pandas as pd


def load_beam_name(read_file):
    """ load beam name
    """

    return pd.read_excel(read_file, sheet_name='梁名編號', index_col=[0, 1], usecols=[1, 2, 3, 4])


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
    from tests.const import const
    from src.etabs_design import load_etabs_design

    beam_name = load_beam_name(const['beam_name_path'])
    print(beam_name.head())

    etabs_design = load_etabs_design(const['etabs_design_path'])
    beam_name_empty = init_beam_name(etabs_design)
    print(beam_name_empty.head())


if __name__ == "__main__":
    main()
