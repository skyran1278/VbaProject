"""
line assigns
"""
from collections import defaultdict


class LineAssigns:
    """
    line assigns
    """

    def __init__(self):
        self.__data = defaultdict(dict)

    def post(self, key, data=None, copy_from=None):
        """
        post
        """
        if copy_from is not None:
            self.__data[key] = self.get(key=copy_from, return_all=True)

        if data is not None:
            self.__data[key] = {
                **self.__data[key], **data
            }

    def get(self, key=None, return_all=False):
        """
        get
        """
        if key is None:
            return self.__data

        if return_all:
            return self.__data[key]

        return self.__data[key]['SECTION']
