"""
test
"""
import numpy as np

from tests.const import const
from src.dataset_etabs_design import load_beam_design, merge_e2k_to_etbas_design


def test_load_beam_design():
    """
    test load_beam_design
    """

    dataset = load_beam_design(const['etabs_design_path'])
    print(dataset)

    columns = ['Story', 'BayID', 'SecID']

    first_row = ['RF', 'B1', 'B50X60']

    assert all(dataset.loc[0, columns] == first_row)


def test_merge_etabs():
    """
    test merge_etabs
    """
    from src.dataset_e2k import load_e2k

    e2k = load_e2k(const['e2k_path'])
    dataset = load_beam_design(const['etabs_design_path'])

    df = merge_e2k_to_etbas_design(dataset, e2k)

    columns = [
        'B', 'H', 'Fc', 'Fy', 'Length',
        'LSupportWidth', 'RSupportWidth'
    ]

    first_row = df.loc[0, columns].values.astype(float)

    data = [0.5, 0.6, 2800.0, 42000.0, 12.0, 0.3, 0.4]

    np.testing.assert_allclose(first_row, data)
