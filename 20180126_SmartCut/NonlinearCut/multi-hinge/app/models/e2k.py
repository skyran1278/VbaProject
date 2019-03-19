"""
e2k model
"""
import re

import numpy as np
import pandas as pd

from app.utils.load_file import load_file


class E2k:
    """
    e2k model
    """

    def __init__(self, path):
        self.content = load_file(path)

        self.title = ''
        self.words = ''

        self.stories = {}
        self.point_coordinates = {}
        self.lines = {}
        self.materials = {}
        self.sections = {}

        self._init_e2k()

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

    def _set_story(self):
        if self.title == '$ STORIES - IN SEQUENCE FROM TOP' and self.words[0] == 'STORY':
            story_name = self.words[1].strip('"')
            height = float(self.words[3])
            self.stories[story_name] = height

    def _set_material(self):
        words = self.words
        if self.title == '$ MATERIAL PROPERTIES' and (words[2] == 'FC' or words[2] == 'FY'):
            material_name = words[1].strip('"')
            self.materials[(material_name, words[2])] = float(words[3])

    def _set_section(self):
        words = self.words
        if self.title == '$ FRAME SECTIONS' and (
                words[0] == 'FRAMESECTION' and words[5] == '"Rectangular"'):
            section_name = words[1].strip('"')
            self.sections[(section_name, 'MATERIAL')] = words[3].strip('"')
            self.sections[(section_name, 'D')] = float(words[7])
            self.sections[(section_name, 'B')] = float(words[9])

        if self.title == '$ CONCRETE SECTIONS' and (
                words[0] == 'CONCRETESECTION' and words[3] == '"BEAM"'):
            section_name = words[1].strip('"')
            self.sections[(section_name, 'COVERTOP')] = float(words[5])
            self.sections[(section_name, 'COVERBOT')] = float(words[7])

    def _set_point_coordinate(self):
        words = self.words
        if self.title == '$ POINT COORDINATES' and words[0] == 'POINT':
            # point_coordinates.append(
            #     (words[1].strip('"'), float(words[2]), float(words[3])))
            point_name = words[1].strip('"')
            self.point_coordinates[point_name] = [
                float(words[2]), float(words[3])]
            # point_coordinates[(point_name, 'X')] = float(words[2])
            # point_coordinates[(point_name, 'Y')] = float(words[3])

    def _set_line(self):
        words = self.words
        if self.title == '$ LINE CONNECTIVITIES' and words[0] == 'LINE':
            line_name = words[1].strip('"')
            line_type = words[2]
            self.lines[(line_name, line_type, 'START')] = words[3].strip('"')
            self.lines[(line_name, line_type, 'END')] = words[4].strip('"')

    def _init_e2k(self):
        for line in self.content:
            # 正規表達式，轉換多個空格變成一個空格，因為 ETABS 自己好像也不管
            line = re.sub(' +', ' ', line)
            if line == '':
                continue
            self.words = np.array(line.split(' '))

            # title 是不容易變的
            if self.words[0] == '$':
                self.title = line
                continue

            self._check_version()
            self._check_unit()
            self._set_story()
            self._set_material()
            # self._set_section()
            # self._set_point_coordinate()
            # self._set_line()

            # if title == '$ LINE ASSIGNS' and words[0] == 'LINEASSIGN' and words[3] == 'SECTION':
            #     # ANG 沒處理
            #     # CARDINALPT 沒處理
            #     line_name = words[1].strip('"')
            #     story = words[2].strip('"')
            #     lines[(line_name, story, 'SECTION')] = words[4].strip('"')

        # point_coordinates = np.array(point_coordinates)
        # point_coordinates = np.array(
        #     point_coordinates, [('name', '<U16'), ('X', '<f8'), ('Y', '<f8')])
        # self.point_coordinates = pd.DataFrame.from_dict(
        #     self.point_coordinates, orient='index', columns=['X', 'Y'])


def main():
    """
    test
    """
    # pylint: disable=line-too-long
    path = 'D:/GitHub/VbaProject/20180126_SmartCut/NonlinearCut/multi-hinge/tests/20190103 v3.0 3floor v16.e2k'

    e2k = E2k(path)
    print(e2k.materials)


if __name__ == "__main__":
    main()
