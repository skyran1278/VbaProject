"""
test
"""
import numpy as np
from tests.const import const
from src.beam import init_beam, put_beam_id, put_column_order


def test_init_beam():
    """
    test init_beam
    """
    from src.e2k import load_e2k
    from src.etabs_design import load_etabs_design, post_e2k

    e2k = load_e2k(const['e2k_path'])
    etabs_design = load_etabs_design(const['etabs_design_path'])
    etabs_design = post_e2k(etabs_design, e2k)

    beam = init_beam(etabs_design)
    print(beam.head())

    columns = [
        ('樓層', ''), ('編號', ''), ('RC 梁寬', ''), ('RC 梁深', ''),
        ('主筋', ''), ('梁長', ''), ('支承寬', '左'), ('支承寬', '右')
    ]

    data = np.array(
        ['RF', 'B1', 50.0, 60.0, '上層 第一排', 1200.0, 30.0, 40.0], dtype=object)

    np.testing.assert_equal(beam.loc[0, columns].values, data)


def test_put_beam_id():
    """
    test init_beam
    """
    from src.e2k import load_e2k
    from src.etabs_design import load_etabs_design, post_e2k, post_beam_name
    from src.beam_name import load_beam_name

    e2k = load_e2k(const['e2k_path'])
    etabs_design = load_etabs_design(const['etabs_design_path'])
    etabs_design = post_e2k(etabs_design, e2k)

    beam = init_beam(etabs_design)

    beam_name = load_beam_name(const['beam_name_path'])
    etabs_design = post_beam_name(etabs_design, beam_name)
    beam = put_beam_id(beam, etabs_design)
    print(beam.head())

    assert beam.at[0, ('編號', '')] == 'G1-1'


def test_put_column_order():
    """
    test put_column_order
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
    beam = init_beam(etabs_design)
    beam, etabs_design = calc_stirrups(beam, etabs_design, const)
    etabs_design = calc_db('BayID', etabs_design, const)
    etabs_design = calc_ld(etabs_design, const)
    etabs_design = add_ld(etabs_design, 'Ld', const['rebar'])

    beam = cut_optimization(beam, etabs_design, const)

    beam = put_column_order(beam)

    assert beam.columns[5] == ('主筋', '左1')
