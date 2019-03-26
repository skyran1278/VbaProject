"""
line connectivities
"""


class Lines:
    """
    lines use in e2k and post
    """

    def __init__(self):
        self.__data = {}
        self.__keys = []

    def get(self, key=None):
        """
        if key is None, return all
        """
        if key is None:
            return self.__data

        return self.__data[key]

    def post(self, key=None, value=None):
        """
        I will use this method in e2k and new e2k
        """
        if not isinstance(value, list):
            raise Exception("value isn't list")

        if value in self.__data.values():
            return

        if key is None:
            int_key = 1
            while int_key in self.__keys:
                int_key += 1

            key = f'B{int_key}'

        if not isinstance(key, str):
            raise Exception('key error')

        self.__data[key] = value
        self.__keys.append(int(key[1:]))


def main():
    """
    test
    """
    lines = Lines()

    lines.post(key='B1', value=['1', '2'])
    lines.post(value=['2', '3'])
    lines.post(value=['1', '2'])
    print(lines.get())
    print(lines.get('B1'))


if __name__ == "__main__":
    main()
