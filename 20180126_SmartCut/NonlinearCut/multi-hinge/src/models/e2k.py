"""
e2k model
"""
import shlex

from collections import defaultdict

import numpy as np

from src.utils.load_file import load_file
from src.models.point_coordinates import PointCoordinates


class E2k:
    """
    e2k model
    """

    def __init__(self, path):
        self.path = path
        self.content = load_file(path)

        self.title = ''
        self.words = ''

        self.stories = {}
        self.materials = {}
        self.sections = defaultdict(dict)
        self.point_coordinates = PointCoordinates()
        self.lines = {}
        # self.point_assigns = {}
        self.line_assigns = {}

        self._init_e2k()

    def _check_version(self):
        if self.title == '$ PROGRAM INFORMATION':
            if self.words[1] != 'ETABS 2016':
                print('PROGRAM should be "ETABS 2016"')

    def _check_unit(self):
        words = self.words
        if self.title == '$ CONTROLS' and words[0] == 'UNITS':
            if words[1] != 'TON' and words[2] != 'M' and words[3] != 'C':
                print('UNITS should be "TON"  "M"  "C"')

    def _post_story(self):
        if self.title == '$ STORIES - IN SEQUENCE FROM TOP':
            self.stories[self.words[1]] = float(self.words[3])

    def _post_material(self):
        words = self.words
        if self.title == '$ MATERIAL PROPERTIES' and (words[2] == 'FC' or words[2] == 'FY'):
            if words[1] in self.materials:
                raise Exception('Material name duplicate!', words[1])
            self.materials[words[1]] = float(words[3])

    def _post_section(self):
        words = self.words
        if self.title == '$ FRAME SECTIONS' and words[5] == 'Concrete Rectangular':
            section_name = words[1]
            self.sections[section_name]['FC'] = words[3]
            self.sections[section_name]['D'] = float(words[7])
            self.sections[section_name]['B'] = float(words[9])

        if self.title == '$ FRAME SECTIONS' and words[-2] == 'I3MOD':
            section_name = words[1]
            count = 2
            while count < len(words):
                self.sections[section_name][words[count]
                                            ] = float(words[count + 1])
                count += 2

        if self.title == '$ CONCRETE SECTIONS' and words[7] == 'Beam':
            section_name = words[1]
            self.sections[section_name]['FY'] = words[3]
            self.sections[section_name]['FYH'] = words[5]

    def _post_point_coordinate(self):
        words = self.words
        if self.title == '$ POINT COORDINATES':
            self.point_coordinates.post(
                words[1], [float(words[2]), float(words[3])])

    # def _post_point_assign(self):
    #     words = self.words
    #     if self.title == '$ POINT ASSIGNS':
    #         self.point_assigns[(words[1], words[2])] = [
    #             float(words[2]), float(words[3])]

    def _post_line(self):
        words = self.words
        if self.title == '$ LINE CONNECTIVITIES' and words[2] == 'BEAM':
            self.lines[words[1]] = [words[3], words[4]]

    def _post_line_assign(self):
        words = self.words
        if self.title == '$ LINE ASSIGNS':
            self.line_assigns[(words[2], words[1])] = words[4]

    def _init_e2k(self):
        for line in self.content:
            # skip space line
            if line == '':
                continue

            # split by space, but ignore space in quotes
            # also adress too many space
            # convenience method
            self.words = shlex.split(line)

            if self.words[0] == '$':
                # post title
                self.title = line
                continue

            self._check_version()
            self._check_unit()
            self._post_story()
            self._post_material()
            self._post_section()
            self._post_point_coordinate()
            self._post_line()
            self._post_line_assign()

        self.sections = dict(self.sections)

    def get_section(self, story, bay_id):
        return self.line_assigns[(story, bay_id)]

    def get_fc(self, story, bay_id):
        """
        get fc
        """
        section = self.line_assigns[(story, bay_id)]
        material = self.sections[section]['FC']

        return self.materials[material]

    def get_fy(self, story, bay_id):
        """
        get fy
        """
        section = self.line_assigns[(story, bay_id)]
        material = self.sections[section]['FY']

        return self.materials[material]

    def get_fyh(self, story, bay_id):
        """
        get fyh
        """
        section = self.line_assigns[(story, bay_id)]
        material = self.sections[section]['FYH']

        return self.materials[material]

    def get_width(self, story, bay_id):
        """
        get width
        """
        section = self.line_assigns[(story, bay_id)]

        return self.sections[section]['B']

    def get_coordinate(self, bay_id=None, point_id=None):
        """
        get coordinate
        """
        if bay_id is not None:
            point_id_start, point_id_end = self.lines[bay_id]
            return self.point_coordinates.get(
                point_id_start), self.point_coordinates.get(point_id_end)

        return self.point_coordinates.get(point_id)


def main():
    """
    test
    """
    from tests.config import config

    e2k = E2k(config['e2k_path'])

    print(e2k.stories)
    print(e2k.materials)
    print(e2k.sections)
    print(e2k.point_coordinates.get())
    print(e2k.lines)
    print(e2k.line_assigns)


if __name__ == "__main__":
    main()
