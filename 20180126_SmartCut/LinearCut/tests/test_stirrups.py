"""
test
"""
import numpy as np
from tests.const import const
from src.stirrups import calc_stirrups, calc_stirrups_3


def test_stirrups():
    """
    test stirrups
    """
    from src.beam import init_beam
    from src.e2k import load_e2k
    from src.etabs_design import load_etabs_design, post_e2k

    e2k = load_e2k(const['e2k_path'])
    etabs_design = load_etabs_design(const['etabs_design_path'])
    etabs_design = post_e2k(etabs_design, e2k)
    beam = init_beam(etabs_design)

    beam, dh_design = calc_stirrups(beam, etabs_design, const, False)
    print(beam.head())
    print(dh_design.head())

    beam_cols = [
        ('箍筋', '左'), ('箍筋', '中'), ('箍筋', '右'),
        ('箍筋長度', '左'), ('箍筋長度', '中'), ('箍筋長度', '右')
    ]

    dh_design_cols = [
        ('VRebarConsiderVc'), ('VSize'), ('Spacing'),
        ('RealVSize'), ('RealSpacing')
    ]

    beam_data = np.array(
        ['#4@12', '#4@15', '#4@10', 282.5, 565, 282.5], dtype=object)

    # 原始數據沒有四捨五入
    dh_design_data = np.array(
        [0.00210399995557964, '#4', 0.12043726490012673, '#4', 0.12], dtype=object)

    np.testing.assert_array_equal(
        beam.loc[0, beam_cols].values, beam_data)

    np.testing.assert_array_equal(
        dh_design.loc[0, dh_design_cols].values, dh_design_data)

    beam, _ = calc_stirrups(beam, etabs_design, const, True)
    print(beam.head())

    beam_data = np.array(
        ['#4@12', '#4@20', '#4@10', 282.5, 565, 282.5], dtype=object)

    np.testing.assert_array_equal(
        beam.loc[0, beam_cols].values, beam_data)


def test_stirrups_3():
    """
    test stirrups
    """
    from src.beam import init_beam
    from src.e2k import load_e2k
    from src.etabs_design import load_etabs_design, post_e2k

    e2k = load_e2k(const['e2k_path'])
    etabs_design = load_etabs_design(const['etabs_design_path'])
    etabs_design = post_e2k(etabs_design, e2k)
    beam = init_beam(etabs_design)

    beam, dh_design = calc_stirrups_3(beam, etabs_design, const, False)
    print(beam.head())
    print(dh_design.head())

    beam_cols = [
        ('箍筋', '左'), ('箍筋', '中'), ('箍筋', '右'),
        ('箍筋長度', '左'), ('箍筋長度', '中'), ('箍筋長度', '右')
    ]

    dh_design_cols = [
        ('VRebarConsiderVc'), ('VSize'), ('Spacing'),
        ('UsrSpacing'), ('RealVSize'), ('RealSpacing')
    ]

    beam_data = np.array(
        ['#4@12', '#4@15', '#4@10', 150, 790, 190], dtype=object)

    # 原始數據沒有四捨五入
    dh_design_data = np.array(
        [0.00210399995557964, '#4', 0.12043726490012673, 0.12, '#4', 0.12], dtype=object)

    np.testing.assert_array_equal(
        beam.loc[0, beam_cols].values, beam_data)

    np.testing.assert_array_equal(
        dh_design.loc[0, dh_design_cols].values, dh_design_data)

    beam, _ = calc_stirrups_3(beam, etabs_design, const, True)
    print(beam.head())

    beam_data = np.array(
        ['#4@12', '#4@22', '#4@10', 290, 530, 310], dtype=object)

    np.testing.assert_array_equal(
        beam.loc[0, beam_cols].values, beam_data)
