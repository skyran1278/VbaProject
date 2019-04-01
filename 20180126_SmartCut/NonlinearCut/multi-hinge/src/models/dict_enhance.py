"""
dict enhance
"""


class DictEnhance:
    """
    for
        point assigns
        line assigns
        line_loads
    """

    def __init__(self):
        self.__data = {}

    def post(self, key, value=None, copy_from=None):
        """
        post
        """
        if value is not None:
            self.__data[key] = value

        elif copy_from is not None:
            self.__data[key] = self.get(copy_from)

    def get(self, key=None):
        """
        get
        """
        if key is None:
            return self.__data

        return self.__data[key]
