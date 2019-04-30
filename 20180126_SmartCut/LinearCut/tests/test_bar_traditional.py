"""
test
"""
import numpy as np


def test_bar_trational():
    """
    test calc_db
    """
    from tests.const import const
    from src.beam import init_beam
    from src.e2k import load_e2k
    from src.etabs_design import load_etabs_design, post_e2k
    from src.stirrups import calc_stirrups
    from src.bar_size_num import calc_db
    from src.bar_ld import calc_ld, add_ld
    from src.bar_traditional import cut_traditional

    e2k = load_e2k(const['e2k_path'])
    etabs_design = load_etabs_design(const['etabs_design_path'])
    etabs_design = post_e2k(etabs_design, e2k)
    beam = init_beam(etabs_design)
    beam, etabs_design = calc_stirrups(beam, etabs_design, const)
    etabs_design = calc_db('BayID', etabs_design, const)
    etabs_design = calc_ld(etabs_design, const)
    etabs_design = add_ld(etabs_design, 'Ld', const['rebar'])

    beam_trational = cut_traditional(beam, etabs_design, const['rebar'])
    print(beam_trational.head())

    cols = [
        ('主筋', '左1'), ('主筋', '中'), ('主筋', '右1'),
        ('主筋長度', '左1'), ('主筋長度', '中'), ('主筋長度', '右1')
    ]

    data = np.array(
        ['5-#10', '2-#10', '5-#10', 376.667, 376.667, 376.667], dtype=object)

    np.testing.assert_array_equal(
        beam_trational.loc[0, cols].values, data)
