"""
sections post get put delete
"""
from collections import defaultdict


class DefaultdictEnhance:
    """
    for
        sections
        line assigns
        point assigns
    method
        get
        post
        delete
    feature
        can copy from exist data
    """

    def __init__(self):
        self.__data = defaultdict(dict)

    def post(self, key, value=None, copy_from=None):
        """
        post
        """
        if copy_from is not None:
            self.__data[key] = self.get(copy_from)

        # if value is str, then just post value
        if isinstance(value, str):
            self.__data[key] = (*self.__data[key], value)

        elif isinstance(value, dict):
            self.__data[key] = {
                **self.__data[key], **value
            }

    def get(self, key=None, key2=None):
        """
        get
        """
        if key is None:
            return self.__data

        if key2 is None:
            return self.__data[key]

        return self.__data[key][key2]

    def delete(self, key):
        """
        delete
        """
        del self.__data[key]


def main():
    """
    test
    """
    sections = DefaultdictEnhance()

    data = {
        'FY': 42000,
        'FYH': 42000,
        'FC': 2800
    }

    sections.post('B60', {'FYH': 28000})
    sections.post('B60', value=data)
    sections.post('B601', copy_from='B60', value={'FY': 42000})
    print(sections.get())


if __name__ == "__main__":
    main()
