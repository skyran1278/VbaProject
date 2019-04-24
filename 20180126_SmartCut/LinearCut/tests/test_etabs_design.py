"""
test
"""
import numpy as np

from tests.const import const
from src.etabs_design import load_etabs_design, post_beam_name, post_e2k


def test_load_etabs_design():
    """
    test load_etabs_design
    """

    etabs_design = load_etabs_design(const['etabs_design_path'])
    print(etabs_design)

    columns = ['Story', 'BayID', 'SecID']

    first_row = ['RF', 'B1', 'B50X60']

    assert all(etabs_design.loc[0, columns] == first_row)


def test_post_e2k():
    """
    test post_e2k
    """
    from src.e2k import load_e2k

    e2k = load_e2k(const['e2k_path'])
    etabs_design = load_etabs_design(const['etabs_design_path'])

    df = post_e2k(etabs_design, e2k)

    columns = [
        'B', 'H', 'Fc', 'Fy', 'Length',
        'LSupportWidth', 'RSupportWidth'
    ]

    first_row = df.loc[0, columns].values

    data = np.array(
        [0.5, 0.6, 2800.0, 42000.0, 12.0, 0.3, 0.4], dtype=object)

    np.testing.assert_array_equal(first_row, data)


def test_post_beam_name():
    """
    test post_beam_name
    """
    from src.beam_name import load_beam_name
    beam_name = load_beam_name(const['beam_name_path'])

    etabs_design = load_etabs_design(const['etabs_design_path'])

    etabs_design = post_beam_name(etabs_design, beam_name)

    assert etabs_design.at[0, 'BeamID'] == 'G1-1'
    assert etabs_design.at[0, 'FrameID'] == 'G1'
