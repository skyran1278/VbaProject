"""
sections post get put delete
"""
from collections import defaultdict


class DefaultdictEnhance:
    """
    for
        line assigns
        sections
        point assigns
    method
        get
        post
    feature
        can copy from exist data
    """

    def __init__(self):
        self.__data = defaultdict(dict)

    def post(self, key, data=None, copy_from=None):
        """
        post
        """
        if copy_from is not None:
            self.__data[key] = self.get(copy_from)

        if data is not None:
            self.__data[key] = {
                **self.__data[key], **data
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
    sections.post('B60', data=data)
    sections.post('B601', copy_from='B60', data={'FY': 42000})
    print(sections.get())

    # sections.get('B60', 'FQ')

    # sections.post(section_name, words[count], float(words[count + 1]))


if __name__ == "__main__":
    main()
