"""
e2k model
"""
import re

import numpy as np
import pandas as pd

from utils.load_file import load_file


class E2k:
    """
    e2k model
    """

    def __init__(self, path):
        self.content = load_file(path)
        self.title = ''
        self.words = np.array()

        self.stories, self.point_coordinates, self.lines, self.materials, self.sections = self._init_e2k()

    def _check_version(self):
        words = self.words
        if self.title == '$ PROGRAM INFORMATION' and words[0] == 'PROGRAM':
            if (words[1] + words[2]) != '"ETABS2016"':
                print('PROGRAM should be "ETABS 2016"')

    def _check_unit(self):
        words = self.words
        if self.title == '$ CONTROLS' and words[0] == 'UNITS':
            if words[1] != '"TON"' and words[2] != '"M"' and words[3] != '"C"':
                print('UNITS should be "TON"  "M"  "C"')

    def _set_story(self, stories):
        if self.title == '$ STORIES - IN SEQUENCE FROM TOP' and self.words[0] == 'STORY':
            story_name = self.words[1].strip('"')
            height = float(self.words[3])
            stories[story_name] = height

    def _init_e2k(self):
        stories = {}
        point_coordinates = {}
        lines = {}
        materials = {}
        sections = {}

        for line in self.content:
            # 正規表達式，轉換多個空格變成一個空格，因為 ETABS 自己好像也不管
            line = re.sub(' +', ' ', line)
            self.words = np.array(line.split(' '))

            # title 是不容易變的
            if self.words[0] == '$':
                self.title = line
                continue

            self._check_version()
            self._check_unit()
            self._set_story(stories)

            if title == '$ MATERIAL PROPERTIES' and (
                    words[0] == 'MATERIAL' and words[3] == '"CONCRETE"'):
                material_name = words[1].strip('"')
                materials[(material_name, 'FY')] = float(words[5])
                materials[(material_name, 'FC')] = float(words[7])

            if title == '$ FRAME SECTIONS' and (
                    words[0] == 'FRAMESECTION' and words[5] == '"Rectangular"'):
                section_name = words[1].strip('"')
                sections[(section_name, 'MATERIAL')] = words[3].strip('"')
                sections[(section_name, 'D')] = float(words[7])
                sections[(section_name, 'B')] = float(words[9])

            if title == '$ CONCRETE SECTIONS' and (
                    words[0] == 'CONCRETESECTION' and words[3] == '"BEAM"'):
                section_name = words[1].strip('"')
                sections[(section_name, 'COVERTOP')] = float(words[5])
                sections[(section_name, 'COVERBOT')] = float(words[7])

            if title == '$ POINT COORDINATES' and words[0] == 'POINT':
                # point_coordinates.append(
                #     (words[1].strip('"'), float(words[2]), float(words[3])))
                point_name = words[1].strip('"')
                point_coordinates[point_name] = [
                    float(words[2]), float(words[3])]
                # point_coordinates[(point_name, 'X')] = float(words[2])
                # point_coordinates[(point_name, 'Y')] = float(words[3])

            if title == '$ LINE CONNECTIVITIES' and words[0] == 'LINE':
                line_name = words[1].strip('"')
                line_type = words[2]
                lines[(line_name, line_type, 'START')] = words[3].strip('"')
                lines[(line_name, line_type, 'END')] = words[4].strip('"')

            # if title == '$ LINE ASSIGNS' and words[0] == 'LINEASSIGN' and words[3] == 'SECTION':
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
