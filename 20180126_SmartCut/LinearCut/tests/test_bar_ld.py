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
    from src.etabs_design import load_etabs_design, post_e2k
    from src.stirrups import calc_stirrups
    from src.bar_size_num import calc_db
    from src.bar_ld import calc_ld, add_ld

    e2k = load_e2k(const['e2k_path'])
    etabs_design = load_etabs_design(const['etabs_design_path'])
    etabs_design = post_e2k(etabs_design, e2k)
    beam = init_beam(etabs_design, moment=3)
    beam, etabs_design = calc_stirrups(beam, etabs_design, const)
    etabs_design = calc_db('BayID', etabs_design, const)

    etabs_design = calc_ld(etabs_design, const)
    print(etabs_design.head())

    etabs_design = add_ld(etabs_design, 'Ld', const['rebar'])
    print(etabs_design.head())

    etabs_design = add_ld(etabs_design, 'SimpleLd', const['rebar'])
    print(etabs_design.head())

    columns = [
        'TopLd', 'TopSimpleLd', 'BotLd', 'BotSimpleLd',
        'BarTopNumLd', 'BarBotNumLd', 'BarTopNumSimpleLd', 'BarBotNumSimpleLd'
    ]

    data = [
        1.176759, 1.996288, 0.714039, 1.211316, 5.0, 6.0, 7.0, 3.0
    ]

    # 因為 16 才有差異
    np.testing.assert_allclose(
        etabs_design.loc[16, columns].values.astype(float), data, atol=0.001)
