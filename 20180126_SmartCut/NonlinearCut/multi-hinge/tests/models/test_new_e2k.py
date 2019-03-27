"""
test
"""
# pylint: disable=redefined-outer-name
import numpy as np

import pytest


@pytest.fixture(scope='module')
def new_e2k():
    """
    new_e2k config
    """
    from src.models.new_e2k import NewE2k
    from tests.config import config

    return NewE2k(config['e2k_path'])


def test_post_point_coordinates(new_e2k):
    """
    test post coor correct
    """
    coordinates = [
        [0., 0.],
        [0.67445007, 0.],
        [0.87367754, 0.],
        [7.12632229, 0.],
        [7.32554951, 0.],
        [8., 0.]
    ]

    new_e2k.post_point_coordinates(coordinates)

    point_coordinates = {
        '1': np.array([0., 0.]),
        '2': np.array([8., 0.]),
        '3': np.array([0.67445007, 0.]),
        '4': np.array([0.87367754, 0.]),
        '5': np.array([7.12632229, 0.]),
        '6': np.array([7.32554951, 0.])
    }

    point_rebars = [
        (0.0046452, 0.0027097),
        (0.0046452, 0.0027097),
        (0.0046452, 0.0027097),
        (0.0046452, 0.0027097),
        (0.0046452, 0.0027097),
        (0.0046452, 0.0027097)
    ]

    np.testing.assert_equal(new_e2k.point_coordinates.get(), point_coordinates)


def test_post_lines(new_e2k):
    """
    test
    """
    coordinates = [
        [0., 0.],
        [0.67445007, 0.],
        [0.87367754, 0.],
        [7.12632229, 0.],
        [7.32554951, 0.],
        [8., 0.]
    ]

    new_e2k.post_lines(coordinates)

    lines = {
        'B1': ['1', '2'],
        'B2': ['1', '3'],
        'B3': ['3', '4'],
        'B4': ['4', '5'],
        'B5': ['5', '6'],
        'B6': ['6', '2']
    }

    assert new_e2k.lines.get() == lines
