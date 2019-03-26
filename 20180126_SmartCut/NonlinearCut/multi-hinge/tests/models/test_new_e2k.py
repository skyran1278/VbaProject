"""
test
"""
import numpy as np

from src.models.new_e2k import NewE2k
from tests.config import config


def test_post_point_coordinates():
    """
    test post coor correct
    """
    new_e2k = NewE2k(config['e2k_path'])

    coordinates = [
        [0., 0.],
        [0.67445007, 0.],
        [0.87367754, 0.],
        [7.12632229, 0.],
        [7.32554951, 0.],
        [8., 0.]
    ]

    new_e2k.post_point_coordinates(coordinates)

    point_coordinates = {'1': np.array([0., 0.]), '2': np.array([8., 0.]), '3': np.array([0.67445007, 0.]), '4': np.array([
        0.87367754, 0.]), '5': np.array([7.12632229, 0.]), '6': np.array([7.32554951, 0.])}

    np.testing.assert_equal(new_e2k.point_coordinates.get(), point_coordinates)
