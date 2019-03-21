"""
sections post get put delete
"""
from collections import defaultdict


class Sections:
    def __init__(self, init_data=None):
        self.data = defaultdict(dict)

        if init_data is not None:
            self.data = init_data

    def post(self, section, key, value):
        self.data[section][key] = value

    def get(self, section, key):
        return self.data[section][key]


def main():
    """
    test
    """
    sections = Sections()

    # sections.post(section_name, words[count], float(words[count + 1]))


if __name__ == "__main__":
    main()
