"""
point coordinates
"""
import numpy as np


class PointCoordinates:
    def __init__(self):
        self.data = {}
        self._keys = []
        self._values = []

    def get(self, parameter_list):
        pass

    def post(self, key=None, value=None):
        """
        I will use this method in e2k and external and new e2k
        """
        if type(value) is list:
            array = value
            np_array = np.array(value)
        elif type(value) is numpy.ndarray:
            array = value.tolist()
            np_array = value

        if key is not None:
            self.data[key] = np_array
            self._keys.append(int(key))
            self._values.append(array)
            return None


def main():
    """
    test
    """
    point_coordinates = PointCoordinates()

    point_coordinates.post(key='1', value=np.array([0, 0]))
    print(type(point_coordinates))


if __name__ == "__main__":
    main()
