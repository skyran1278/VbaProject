"""
sections post get put delete
"""
import warnings
from collections import defaultdict


def nest_defaultdict():
    """
    nest defaultdict
    because defaultdict have side effect
        when get key will create key = {}
        it is difficult to debug
        so maybe it is right to go back to simple {}
    """
    return defaultdict(nest_defaultdict)


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
        self.__data = {}

    def post(self, key, value=None, copy_from=None):
        """
        post
        """
        if key in self.__data:
            warnings.warn(f'{key} in data')
        if copy_from is not None:
            self.__data[key] = self.get(copy_from)

        # if value is str, then just post value
        if isinstance(value, (str, list, tuple)):
            self.__data[key] = (*self.__data.setdefault(key, ()), value)

        elif isinstance(value, dict):
            self.__data[key] = {
                **self.__data.setdefault(key, {}), **value
            }

    def get(self, key=None, key2=None):
        """
        get
        """
        if key2 is not None:
            # if key not in self.__data:
            #     return None
            # if key2 not in self.__data[key]:
            #     return None
            return self.__data[key][key2]

        if key is not None:
            # if key not in self.__data:
            #     return None
            return self.__data.get(key, None)

        return self.__data

        # if key is None:
        #     return self.__data

        # if key2 is None:
        #     if key not in self.__data:
        #         # warnings.warn(f'{key} is not in dict')
        #         return None
        #     return self.__data[key]

        # if key not in self.__data:
        #     warnings.warn(f'{key} is not in dict')

        # if key2 not in self.__data[key]:
        #     warnings.warn(f'{key, key2} is not in dict')

        # # return self.__data[key].get(key2, {})
        # return self.__data[key][key2]

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
    sections.post('B59', copy_from='B60', value={'FY': 42000})
    sections.post('B1', value=(1, 2))
    print(sections.get('B59', 'FY'))
    print(sections.get())


if __name__ == "__main__":
    main()
