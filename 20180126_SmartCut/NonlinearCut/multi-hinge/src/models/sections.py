"""
sections post get put delete
"""
from collections import defaultdict


class Sections:
    """
    sections
    """

    def __init__(self):
        self.__data = defaultdict(dict)

    def post(self, section, data=None, copy_from=None):
        """
        post
        """
        if copy_from is not None:
            self.__data[section] = self.get(section=copy_from)

        if data is not None:
            self.__data[section] = {
                **self.__data[section], **data
            }

    def get(self, section=None, key=None):
        """
        get
        """
        if section is None:
            return self.__data

        if key is None:
            return self.__data[section]

        return self.__data[section][key]


def main():
    """
    test
    """
    sections = Sections()

    data = {
        'FY': 42000,
        'FYH': 42000,
        'FC': 2800
    }

    sections.post('B60', {'FYH': 28000})
    sections.post(section='B60', data=data)
    sections.post(section='B601', copy_from='B60', data={'FY': 42000})
    print(sections.get())

    # sections.get('B60', 'FQ')

    # sections.post(section_name, words[count], float(words[count + 1]))


if __name__ == "__main__":
    main()
