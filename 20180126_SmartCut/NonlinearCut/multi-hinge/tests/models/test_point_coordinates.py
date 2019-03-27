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
    point_coordinates.post(value=[0, 1])
    point_coordinates.post(value=[1 / 3, 1])
    point_coordinates.post(value=[1 / 3, 1])

    point_coordinates_has = {
        '1': np.array([0, 0]),
        '2': np.array([0, 1]),
        '3': np.array([1 / 3, 1.])
    }

    np.testing.assert_equal(
        point_coordinates.get('1'), np.array([0, 0])
    )

    np.testing.assert_equal(
        point_coordinates.get(), point_coordinates_has
    )

    assert point_coordinates.get(value=np.array([1 / 3, 1])) == '3'
