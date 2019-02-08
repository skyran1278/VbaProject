""" load e2k
"""
import os
import sys
import pickle
import re

import pandas as pd
import numpy as np

# SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
# sys.path.append(os.path.join(SCRIPT_DIR, os.path.pardir))

from const import E2K

read_file = f'{SCRIPT_DIR}/{E2K}'
save_file = f'{SCRIPT_DIR}/../temp/{E2K}.pkl'


def _load_e2k(read_file):
    # with open(read_file, encoding='utf-8') as f:
    with open(read_file) as f:
        content = f.readlines()
        content = [x.strip() for x in content]
    return content


def _init_e2k(read_file):
    content = _load_e2k(read_file)

    rebars = {}

    stories = {}
    point_coordinates = {}
    # point_coordinates = []
    lines = {}

    materials = {}

    sections = {}

    for line in content:
        # 轉換多格成一格，因為 ETABS 自己好像也不管
        line = re.sub(' +', ' ', line)
        words = np.array(line.split(' '))

        if words[0] == '':
            continue

        if words[0] == '$':
            checking = line

        if checking == '$ PROGRAM INFORMATION' and words[0] == 'PROGRAM':
            if words[1] != '"ETABS"':
                print('PROGRAM should be "ETABS"')
            if words[3] != '"9.7.3"':
                print('VERSION should be "9.7.3"')
            continue

        if checking == '$ CONTROLS' and words[0] == 'UNITS':
            if words[1] != '"TON"' and words[2] != '"M"':
                print('UNITS should be "TON"  "M"')
            continue

        if checking == '$ STORIES - IN SEQUENCE FROM TOP' and words[0] == 'STORY':
            story_name = words[1].strip('"')
            height = float(words[3])
            stories[story_name] = height

        if checking == '$ MATERIAL PROPERTIES' and words[0] == 'MATERIAL' and words[3] == '"CONCRETE"':
            material_name = words[1].strip('"')
            materials[(material_name, 'FY')] = float(words[5])
            materials[(material_name, 'FC')] = float(words[7])

        if checking == '$ FRAME SECTIONS' and words[0] == 'FRAMESECTION' and words[5] == '"Rectangular"':
            section_name = words[1].strip('"')
            sections[(section_name, 'MATERIAL')] = words[3].strip('"')
            sections[(section_name, 'D')] = float(words[7])
            sections[(section_name, 'B')] = float(words[9])

        if checking == '$ REBAR DEFINITIONS' and words[0] == 'REBARDEFINITION':
            rebar_name = words[1].strip('"')
            rebars[(rebar_name, 'AREA')] = float(words[3])
            rebars[(rebar_name, 'DIA')] = float(words[5])

        if checking == '$ CONCRETE SECTIONS' and words[0] == 'CONCRETESECTION' and words[3] == '"BEAM"':
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

    return rebars, stories, point_coordinates, lines, materials, sections


def _init_pkl(read_file, save_file):
    dataset = _init_e2k(read_file)

    print("Creating pickle file ...")
    with open(save_file, 'wb') as f:
        pickle.dump(dataset, f, True)
    print("Done!")


def load_e2k(read_file=read_file, save_file=save_file):
    if not os.path.exists(save_file):
        _init_pkl(read_file, save_file)

    with open(save_file, 'rb') as f:
        rebars, stories, point_coordinates, lines, materials, sections = pickle.load(
            f)

    return rebars, stories, point_coordinates, lines, materials, sections


if __name__ == '__main__':
    _init_pkl(read_file, save_file)
    rebars, stories, point_coordinates, lines, materials, sections = load_e2k(
        read_file, save_file)
    print(stories)
