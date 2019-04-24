"""
test
"""
import numpy as np
from src.init_beam import *


def test_init_beam():
    """
    test init_beam
    """
    from tests.const import const
    from src.dataset_e2k import load_e2k
    from src.dataset_etabs_design import load_beam_design, merge_e2k_to_etbas_design
    from src.dataset_beam_name import load_beam_name

    e2k_path, etabs_design_path, beam_name_path = const[
        'e2k_path'], const['etabs_design_path'], const['beam_name_path']

    e2k = load_e2k(e2k_path)
    etabs_design = load_beam_design(etabs_design_path)
    etabs_design = merge_e2k_to_etbas_design(etabs_design, e2k)
    beam_name = load_beam_name(beam_name_path)

    beam = init_beam(etabs_design, e2k, moment=3)

    columns = [
        ('樓層', ''), ('編號', ''), ('RC 梁寬', ''), ('RC 梁深', ''),
        ('主筋', ''), ('梁長', ''), ('支承寬', '左'), ('支承寬', '右')
    ]

    data = np.array(
        ['RF', 'B1', 50.0, 60.0, '上層 第一排', 1200.0, 30.0, 40.0], dtype=object)

    np.testing.assert_equal(beam.loc[0, columns].values, data)
