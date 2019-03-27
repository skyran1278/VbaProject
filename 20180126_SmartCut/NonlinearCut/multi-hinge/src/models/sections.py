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

    def post(self, section=None, key=None, value=None):
        """
        post
        """
        self.__data[section][key] = value

    def get(self, section=None, key=None):
        """
        get
        """
        if section is None:
            return self.__data
        return self.__data[section][key]


def main():
    """
    test
    """
    sections = Sections()

    # sections.post(section_name, words[count], float(words[count + 1]))


if __name__ == "__main__":
    main()
