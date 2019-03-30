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

        # remember key to do next loop
        self.key = 1

        # to get key by value
        self.__reverse_data = {}

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
        value = tuple(value)

        # low performance
        # for line, points in self.__data.items():
        #     if value == points:
        #         return line
        #
        # better performance in large data
        if value in self.__reverse_data:
            return self.__reverse_data[value]

        if key is None:
            while self.key in self.__keys:
                self.key += 1

            key = f'B{self.key}'

        if not isinstance(key, str):
            raise Exception('key error')

        self.__data[key] = value
        self.__reverse_data[value] = key
        self.__keys.append(int(key[1:]))

        return key


def main():
    """
    test
    """
    from random import randint

    lines = Lines()

    for _ in range(10000):
        strat = randint(1, 30)
        end = randint(1, 30)

        value = [str(strat), str(end)]
        lines.post(value=value)

        lines.get('B1')

    lines.post(key='B1', value=['1', '2'])
    lines.post(value=['2', '3'])
    lines.post(value=['1', '2'])
    print(lines.get())
    print(lines.get('B1'))


if __name__ == "__main__":
    import cProfile
    cProfile.run('main()')
