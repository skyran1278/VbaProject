"""
point coordinates
"""
import numpy as np


def _get_tuple_nparray(value):
    if isinstance(value, np.ndarray):
        array = tuple(value.tolist())
        np_array = value
    elif isinstance(value, (tuple, list)):
        array = tuple(value)
        np_array = np.array(value)
    else:
        raise Exception('no give value')

    return array, np_array


class PointCoordinates:
    """
    use in e2k
    """

    def __init__(self):
        self.__data = {}

        # key is str(int), so bulid __keys to store pure int.
        # convenient to use int plus
        self.__keys = []

        # remember key to do next loop
        self.key = 1

        # because numpy is difficult to check
        # __reverse_data easy to check if value exist
        self.__reverse_data = {}

    def get(self, key=None, value=None):
        """
        get by str key, haven't support int key
        """
        if key is None and value is None:
            return self.__data

        if key is None:
            array, _ = _get_tuple_nparray(value)
            return self.__reverse_data[array]

        return self.__data[key]

    def post(self, key=None, value=None):
        """
        I will use this method in e2k and new e2k
        """
        array, np_array = _get_tuple_nparray(value)

        if array in self.__reverse_data:
            return self.__reverse_data[array]

        if key is None:
            while self.key in self.__keys:
                self.key += 1

            key = str(self.key)

        if not isinstance(key, str):
            raise Exception('key error')

        self.__data[key] = np_array
        self.__reverse_data[array] = key
        self.__keys.append(int(key))

        return key


def main():
    """
    test
    """
    point_coordinates = PointCoordinates()

    point_coordinates.post(key='1', value=np.array([0, 0]))
    point_coordinates.post(value=[0, 1])
    point_coordinates.post(value=[1 / 3, 1])
    point_coordinates.post(value=[1 / 3, 1])
    print(point_coordinates.get())
    print(point_coordinates.get('1'))
    print(point_coordinates.get(value=np.array([0, 1])))


if __name__ == "__main__":
    main()
