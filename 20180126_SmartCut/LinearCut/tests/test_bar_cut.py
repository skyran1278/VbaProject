"""
test
"""
import numpy as np


def test_bar_cut():
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
    from src.bar_cut import cut_optimization

    e2k = load_e2k(const['e2k_path'])
    etabs_design = load_etabs_design(const['etabs_design_path'])
    etabs_design = post_e2k(etabs_design, e2k)
    beam = init_beam(etabs_design, moment=3)
    beam, etabs_design = calc_stirrups(beam, etabs_design, const)
    etabs_design = calc_db('BayID', etabs_design, const)
    etabs_design = calc_ld(etabs_design, const)
    etabs_design = add_ld(etabs_design, 'Ld', const['rebar'])

    beam = cut_optimization(beam, etabs_design, const)
    print(beam.head())

    cols = [
        ('主筋', '左'), ('主筋', '中'), ('主筋', '右'),
        ('主筋長度', '左'), ('主筋長度', '中'), ('主筋長度', '右')
    ]

    data = np.array(
        ['5-#10', '2-#10', '5-#10', 230, 630, 270], dtype=object)

    np.testing.assert_array_equal(
        beam.loc[0, cols].values, data)
