""" load e2k
"""
import os
import pickle
import re

import pandas as pd
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


def _init_e2k(read_file):
    content = _load_e2k(read_file)

    # rebars = {}

    # stories = {}
    point_coordinates = {}
    # point_coordinates = []
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

        # if checking == '$ STORIES - IN SEQUENCE FROM TOP' and words[0] == 'STORY':
        #     story_name = words[1].strip('"')
        #     height = float(words[3])
        #     stories[story_name] = height

        if checking == '$ MATERIAL PROPERTIES' and (
                words[0] == 'MATERIAL' and words[3] == '"CONCRETE"'):
            material_name = words[1].strip('"')
            materials[(material_name, 'FY')] = float(words[5])
            materials[(material_name, 'FC')] = float(words[7])

        if checking == '$ FRAME SECTIONS' and (
                words[0] == 'FRAMESECTION' and words[5] == '"Rectangular"'):
            section_name = words[1].strip('"')
            sections[(section_name, 'MATERIAL')] = words[3].strip('"')
            sections[(section_name, 'D')] = float(words[7])
            sections[(section_name, 'B')] = float(words[9])

        # if checking == '$ REBAR DEFINITIONS' and words[0] == 'REBARDEFINITION':
        #     rebar_name = words[1].strip('"')
        #     rebars[(rebar_name, 'AREA')] = float(words[3])
        #     rebars[(rebar_name, 'DIA')] = float(words[5])

        if checking == '$ CONCRETE SECTIONS' and (
                words[0] == 'CONCRETESECTION' and words[3] == '"BEAM"'):
            section_name = words[1].strip('"')
            sections[(section_name, 'COVERTOP')] = float(words[5])
            sections[(section_name, 'COVERBOT')] = float(words[7])

        if checking == '$ POINT COORDINATES' and words[0] == 'POINT':
            # point_coordinates.append(
            #     (words[1].strip('"'), float(words[2]), float(words[3])))
            point_name = words[1].strip('"')
            point_coordinates[point_name] = [float(words[2]), float(words[3])]
            # point_coordinates[(point_name, 'X')] = float(words[2])
            # point_coordinates[(point_name, 'Y')] = float(words[3])

        if checking == '$ LINE CONNECTIVITIES' and words[0] == 'LINE':
            line_name = words[1].strip('"')
            line_type = words[2]
            lines[(line_name, line_type, 'START')] = words[3].strip('"')
            lines[(line_name, line_type, 'END')] = words[4].strip('"')

        # if checking == '$ LINE ASSIGNS' and words[0] == 'LINEASSIGN' and words[3] == 'SECTION':
        #     # ANG 沒處理
        #     # CARDINALPT 沒處理
        #     line_name = words[1].strip('"')
        #     story = words[2].strip('"')
        #     lines[(line_name, story, 'SECTION')] = words[4].strip('"')

    # point_coordinates = np.array(point_coordinates)
    # point_coordinates = np.array(
    #     point_coordinates, [('name', '<U16'), ('X', '<f8'), ('Y', '<f8')])
    point_coordinates = pd.DataFrame.from_dict(
        point_coordinates, orient='index', columns=['X', 'Y'])

    return {
        'point_coordinates': point_coordinates,
        'lines': lines,
        'materials': materials,
        'sections': sections
    }


def _init_pkl(read_file, save_file):
    dataset = _init_e2k(read_file)

    print("Creating pickle file ...")
    with open(save_file, 'wb') as f:
        pickle.dump(dataset, f, True)
    print("Done!")


def load_e2k(read_file, save_file):
    """ load e2k file
    """
    if not os.path.exists(save_file):
        _init_pkl(read_file, save_file)

    with open(save_file, 'rb') as f:
        e2k = pickle.load(f)

    return e2k


if __name__ == '__main__':
    from const import const
    E2K_PATH = const['e2k_path']
    READ_FILE = E2K_PATH
    SAVE_FILE = f'{E2K_PATH}.pkl'

    _init_pkl(READ_FILE, SAVE_FILE)
    E2K = load_e2k(READ_FILE, SAVE_FILE)
    print(E2K['point_coordinates'])
