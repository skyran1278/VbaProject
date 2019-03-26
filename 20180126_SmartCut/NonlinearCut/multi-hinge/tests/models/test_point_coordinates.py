"""
test
"""
import numpy as np

from src.models.point_coordinates import PointCoordinates


def test_point_coordinates():
    """
    test
    """
    point_coordinates = PointCoordinates()

    point_coordinates.post(key='1', value=np.array([0, 0]))
    point_coordinates.post(value=[1 / 3, 1])

    np.testing.assert_equal(point_coordinates.get('1'), np.array([0, 0]))
    assert point_coordinates.get(value=np.array([1 / 3, 1])) == '2'
