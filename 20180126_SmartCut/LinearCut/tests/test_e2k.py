"""
test
"""
import numpy as np

from src.e2k import load_e2k
from tests.const import const


def test_load_e2k():
    """
    test load_e2k
    """

    dataset = load_e2k(const['e2k_path'])

    print(dataset['point_coordinates'])
    print(dataset['lines'])
    print(dataset['materials'])
    print(dataset['sections'])

    point_coordinates = {
        '1': np.array([0., 0.]),
        '2': np.array([12., 0.]),
        '3': np.array([24., 0.]),
        '4': np.array([36., 0.]),
    }

    lines = {
        ('C1', 'COLUMN'): ('1', '1'),
        ('C2', 'COLUMN'): ('2', '2'),
        ('C3', 'COLUMN'): ('3', '3'),
        ('C4', 'COLUMN'): ('4', '4'),
        ('B1', 'BEAM'): ('1', '2'),
        ('B2', 'BEAM'): ('2', '3'),
        ('B3', 'BEAM'): ('3', '4')
    }

    materials = {
        ('CONC', 'FY'): 42184.18, ('CONC', 'FC'): 2812.279,
        ('C280', 'FY'): 42000.0, ('C280', 'FC'): 2800.0
    }

    sections = {
        ('C80X80', 'MATERIAL'): 'C280', ('C80X80', 'H'): 0.8,
        ('C80X80', 'B'): 0.8, ('B50X60', 'MATERIAL'): 'C280',
        ('B50X60', 'H'): 0.6, ('B50X60', 'B'): 0.5,
        ('C60X60', 'MATERIAL'): 'C280', ('C60X60', 'H'): 0.6, ('C60X60', 'B'): 0.6
    }

    np.testing.assert_equal(dataset['point_coordinates'], point_coordinates)

    assert dataset['lines'] == lines
    assert dataset['materials'] == materials
    assert dataset['sections'] == sections
