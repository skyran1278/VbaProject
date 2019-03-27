"""
e2k model
"""
import shlex

# from collections import defaultdict

from src.utils.load_file import load_file
from src.models.point_coordinates import PointCoordinates
from src.models.lines import Lines
from src.models.sections import Sections


class E2k:
    """
    e2k model
    """

    def __init__(self, path):
        self.path = path
        self.content = load_file(path)

        self.stories = {}
        self.materials = {}
        self.sections = Sections()
        self.point_coordinates = PointCoordinates()
        self.lines = Lines()
        # self.point_assigns = {}
        self.line_assigns = {}

        self._init_e2k()

    def _init_e2k(self):  # pylint: disable=too-many-branches
        for line in self.content:
            # skip space line
            if line == '':
                continue

            # split by space, but ignore space in quotes
            # also adress too many space
            # convenience method
            words = shlex.split(line)

            if words[0] == '$':
                # post title
                title = line
                continue

            if title == '$ PROGRAM INFORMATION':
                if words[1] != 'ETABS 2016':
                    print('PROGRAM should be "ETABS 2016"')

            elif title == '$ CONTROLS' and words[0] == 'UNITS':
                if words[1] != 'TON' and words[2] != 'M' and words[3] != 'C':
                    print('UNITS should be "TON"  "M"  "C"')

            elif title == '$ STORIES - IN SEQUENCE FROM TOP':
                self.stories[words[1]] = float(words[3])

            elif title == '$ MATERIAL PROPERTIES' and (words[2] == 'FC' or words[2] == 'FY'):
                if words[1] in self.materials:
                    raise Exception('Material name duplicate!', words[1])
                self.materials[words[1]] = float(words[3])

            elif title == '$ FRAME SECTIONS' and words[5] == 'Concrete Rectangular':
                section = words[1]
                self.sections.post(section, 'FC', words[3])
                self.sections.post(section, 'D', float(words[7]))
                self.sections.post(section, 'B', float(words[9]))

            elif title == '$ FRAME SECTIONS' and words[2] != 'MATERIAL':
                section = words[1]
                count = 2
                while count < len(words):
                    self.sections.post(
                        section, words[count], float(words[count + 1]))
                    count += 2

            elif title == '$ CONCRETE SECTIONS' and words[7] == 'Beam':
                section = words[1]
                self.sections.post(section, 'FY', words[3])
                self.sections.post(section, 'FYH', words[5])
                self.sections.post(section, 'COVERTOP', float(words[9]))
                self.sections.post(section, 'COVERBOTTOM', float(words[11]))
                self.sections.post(section, 'ATI', float(words[13]))
                self.sections.post(section, 'ABI', float(words[15]))
                self.sections.post(section, 'ATJ', float(words[17]))
                self.sections.post(section, 'ABJ', float(words[19]))

            elif title == '$ POINT COORDINATES':
                self.point_coordinates.post(
                    words[1], (float(words[2]), float(words[3])))

            elif title == '$ LINE CONNECTIVITIES' and words[2] == 'BEAM':
                self.lines.post(words[1], [words[3], words[4]])

            # elif title == '$ POINT ASSIGNS':
            #     self.point_assigns[(words[1], words[2])] = [
            #         float(words[2]), float(words[3])]

            elif title == '$ LINE ASSIGNS':
                self.line_assigns[(words[2], words[1])] = words[4]

    def get_section(self, story, bay_id):
        """
        sections
        """
        return self.line_assigns[(story, bay_id)]

    def get_fc(self, story, bay_id):
        """
        get fc
        """
        section = self.line_assigns[(story, bay_id)]
        material = self.sections.get(section, 'FC')

        return self.materials[material]

    def get_fy(self, story, bay_id):
        """
        get fy
        """
        section = self.line_assigns[(story, bay_id)]
        material = self.sections.get(section, 'FY')

        return self.materials[material]

    def get_fyh(self, story, bay_id):
        """
        get fyh
        """
        section = self.line_assigns[(story, bay_id)]
        material = self.sections.get(section, 'FYH')

        return self.materials[material]

    def get_width(self, story, bay_id):
        """
        get width
        """
        section = self.line_assigns[(story, bay_id)]

        return self.sections.get(section, 'B')

    def get_coordinate(self, bay_id=None, point_id=None):
        """
        get coordinate
        """
        if bay_id is not None:
            point_id_start, point_id_end = self.lines.get(bay_id)
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
    print(e2k.sections.get())
    print(e2k.point_coordinates.get())
    print(e2k.lines.get())
    print(e2k.line_assigns)


if __name__ == "__main__":
    main()
