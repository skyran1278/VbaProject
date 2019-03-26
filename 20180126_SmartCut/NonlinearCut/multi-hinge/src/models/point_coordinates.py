"""
point coordinates
"""
import numpy as np


class PointCoordinates:
    """
    use in e2k
    """

    def __init__(self):
        self.__data = {}

        # key is str(int), so bulid __keys to store pure int.
        # convenient to use int plus
        self.__keys = []

        # because numpy is difficult to check
        # __values easy to check if exist
        self.__values = []

    def get(self, key=None):
        """
        get by str key, haven't support int key
        """
        if key is None:
            return self.__data

        return self.__data[key]

    def post(self, key=None, value=None):
        """
        I will use this method in e2k and new e2k
        """
        if isinstance(value, list):
            array = value
            np_array = np.array(value)
        elif isinstance(value, np.ndarray):
            array = value.tolist()
            np_array = value
        else:
            raise Exception('no give value')

        if array in self.__values:
            return

        if key is None:
            int_key = 1
            while int_key in self.__keys:
                int_key += 1

            key = str(int_key)

        if not isinstance(key, str):
            raise Exception('key error')

        self.__data[key] = np_array
        self.__keys.append(int(key))
        self.__values.append(array)


def main():
    """
    test
    """
    point_coordinates = PointCoordinates()

    point_coordinates.post(key='1', value=np.array([0, 0]))
    point_coordinates.post(value=[0, 1])
    point_coordinates.post(value=[0, 1])
    print(point_coordinates.get())
    print(point_coordinates.get('1'))


if __name__ == "__main__":
    main()
