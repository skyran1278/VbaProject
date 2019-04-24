"""
test
"""
import numpy as np
from tests.const import const
from src.beam import init_beam, put_beam_id


def test_init_beam():
    """
    test init_beam
    """
    from src.e2k import load_e2k
    from src.etabs_design import load_etabs_design, post_e2k

    e2k = load_e2k(const['e2k_path'])
    etabs_design = load_etabs_design(const['etabs_design_path'])
    etabs_design = post_e2k(etabs_design, e2k)

    beam = init_beam(etabs_design, moment=3)
    print(beam.head())

    assert ('主筋', '左') in beam

    beam = init_beam(etabs_design, moment=5)
    print(beam.head())

    assert ('主筋', '左1') in beam

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

    beam = init_beam(etabs_design, moment=3)

    beam_name = load_beam_name(const['beam_name_path'])
    etabs_design = post_beam_name(etabs_design, beam_name)
    beam = put_beam_id(beam, etabs_design)
    print(beam.head())

    assert beam.at[0, ('編號', '')] == 'G1-1'
