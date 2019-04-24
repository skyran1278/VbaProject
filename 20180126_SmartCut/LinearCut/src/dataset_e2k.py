""" load e2k
"""
import os
import pickle
import re

import numpy as np


def _load_e2k(read_file):
    with open(read_file, encoding='big5') as path:
        content = path.readlines()
        content = [x.strip() for x in content]

    return content


def _is_right_version(checking, words):
    if checking == '$ PROGRAM INFORMATION' and words[0] == 'PROGRAM':
        if words[1] != '"ETABS"':
            print('PROGRAM should be "ETABS"')
        if words[3] != '"9.7.3"':
            print('VERSION should be "9.7.3"')


def _is_right_unit(checking, words):
    if checking == '$ CONTROLS' and words[0] == 'UNITS':
        if words[1] != '"TON"' and words[2] != '"M"':
            print('UNITS should be "TON"  "M"')


def load_e2k(read_file):
    """ load e2k file
    """
    content = _load_e2k(read_file)

    point_coordinates = {}
    lines = {}
    materials = {}
    sections = {}

    for line in content:
        # 正規表達式，轉換多格成一格，因為 ETABS 自己好像也不管
        line = re.sub(' +', ' ', line)
        words = np.array(line.split(' '))

        # checking 是不容易變的
        if words[0] == '$':
            checking = line

        _is_right_version(checking, words)
        _is_right_unit(checking, words)

        if checking == '$ MATERIAL PROPERTIES' and (
                words[0] == 'MATERIAL' and words[3] == '"CONCRETE"'):
            material_name = words[1].strip('"')
            materials[(material_name, 'FY')] = float(words[5])
            materials[(material_name, 'FC')] = float(words[7])

        if checking == '$ FRAME SECTIONS' and (
                words[0] == 'FRAMESECTION' and words[5] == '"Rectangular"'):
            section_name = words[1].strip('"')
            sections[(section_name, 'MATERIAL')] = words[3].strip('"')
            sections[(section_name, 'H')] = float(words[7])
            sections[(section_name, 'B')] = float(words[9])

        if checking == '$ POINT COORDINATES' and words[0] == 'POINT':
            point_name = words[1].strip('"')
            point_coordinates[point_name] = np.array(
                [words[2], words[3]]).astype(np.float)

        if checking == '$ LINE CONNECTIVITIES' and words[0] == 'LINE':
            line_name = words[1].strip('"')
            line_type = words[2]
            lines[(line_name, line_type)] = (
                words[3].strip('"'), words[4].strip('"')
            )

    return {
        'point_coordinates': point_coordinates,
        'lines': lines,
        'materials': materials,
        'sections': sections
    }


def main():
    """
    test
    """
    from tests.const import const

    dataset = load_e2k(const['e2k_path'])

    print(dataset['point_coordinates'])
    print(dataset['lines'])
    print(dataset['materials'])
    print(dataset['sections'])


if __name__ == "__main__":
    main()
