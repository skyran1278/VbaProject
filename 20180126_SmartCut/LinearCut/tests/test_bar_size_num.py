"""
test
"""
import numpy as np


def test_calc_db():
    """
    test calc_db
    """
    from tests.const import const
    from src.beam import init_beam
    from src.e2k import load_e2k
    from src.etabs_design import load_etabs_design, post_e2k, post_beam_name
    from src.stirrups import calc_stirrups
    from src.beam_name import load_beam_name
    from src.bar_size_num import calc_db

    e2k = load_e2k(const['e2k_path'])
    etabs_design = load_etabs_design(const['etabs_design_path'])
    etabs_design = post_e2k(etabs_design, e2k)
    beam = init_beam(etabs_design)
    beam, etabs_design = calc_stirrups(beam, etabs_design, const)

    etabs_design = calc_db('BayID', etabs_design, const)
    print(etabs_design.head())

    columns = [
        'BarTopSize', 'BarTopCap', 'BarTopNum', 'BarTop1st', 'BarTop2nd',
        'BarBotSize', 'BarBotCap', 'BarBotNum', 'BarBot1st', 'BarBot2nd'
    ]

    data = np.array(
        ['#10', 5.0, 7.0, 5.0, 2.0, '#8', 6.0, 5.0, 5.0, 0.0], dtype=object)

    np.testing.assert_array_equal(
        etabs_design.loc[0, columns].values, data)

    beam_name = load_beam_name(const['beam_name_path'])
    etabs_design = post_beam_name(etabs_design, beam_name)

    etabs_design = calc_db('FrameID', etabs_design, const)
    print(etabs_design.head())

    np.testing.assert_array_equal(
        etabs_design.loc[0, columns].values, data)
