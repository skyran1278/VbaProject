"""
e2k model
"""
import shlex

# from collections import defaultdict

from src.utils.load_file import load_file
from src.models.point_coordinates import PointCoordinates
from src.models.lines import Lines
from src.models.defaultdict_enhance import DefaultdictEnhance


class E2k:
    """
    e2k model
    """

    def __init__(self, path):
        self.path = path
        self.content = load_file(path)

        self.stories = {}
        self.materials = {}
        self.sections = DefaultdictEnhance()
        self.point_coordinates = PointCoordinates()
        self.lines = Lines()
        self.columns = Lines()
        self.point_assigns = DefaultdictEnhance()
        self.line_assigns = DefaultdictEnhance()
        self.line_hinges = []
        self.dead_load_name = None
        self.line_loads = DefaultdictEnhance()

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

                self.sections.post(
                    section,
                    {
                        'FC': words[3],
                        'D': float(words[7]),
                        'B': float(words[9])
                    }
                )

            elif title == '$ FRAME SECTIONS' and words[2] != 'MATERIAL':
                section = words[1]
                count = 2
                while count < len(words):
                    self.sections.post(
                        section, {
                            'Property Modifiers': ' '.join(words[2:])
                        })
                    count += 2

            elif title == '$ CONCRETE SECTIONS' and words[7] == 'Beam':
                section = words[1]
                self.sections.post(
                    section,
                    {
                        'FY': words[3],
                        'FYH': words[5],
                        'COVERTOP': float(words[9]),
                        'COVERBOTTOM': float(words[11]),
                        'ATI': float(words[13]),
                        'ABI': float(words[15]),
                        'ATJ': float(words[17]),
                        'ABJ': float(words[19])
                    }
                )

            elif title == '$ CONCRETE SECTIONS' and words[7] == 'Column':
                section = words[1]
                self.sections.delete(section)

            elif title == '$ POINT COORDINATES':
                self.point_coordinates.post(
                    words[1], (float(words[2]), float(words[3])))

            elif title == '$ LINE CONNECTIVITIES' and words[2] == 'BEAM':
                self.lines.post(words[1], [words[3], words[4]])

            elif title == '$ LINE CONNECTIVITIES' and words[2] == 'COLUMN':
                self.columns.post(words[1], [words[3], words[4]])

            elif title == '$ POINT ASSIGNS':
                self.point_assigns.post(
                    (words[2], words[1]), ' '.join(words[3:])
                )

            elif title == '$ LINE ASSIGNS':
                count = 3
                while count < len(words):
                    self.line_assigns.post(
                        (words[2], words[1]), {
                            words[count]: words[count + 1]
                        })
                    count += 2

            elif title == '$ LOAD PATTERNS' and words[2] == 'TYPE':
                if words[3] == 'Dead':
                    self.dead_load_name = words[1]

            elif title == '$ FRAME OBJECT LOADS':
                self.line_loads.post((words[2], words[1]), {
                    words[8]: {
                        'TYPE': words[4],
                        'DIR': words[6],
                        'FVAL': words[10],
                    }
                })

    def get_section(self, story, bay_id):
        """
        sections
        """
        return self.line_assigns.get((story, bay_id), 'SECTION')

    def get_fc(self, story, bay_id):
        """
        get fc
        """
        section = self.get_section(story, bay_id)
        material = self.sections.get(section, 'FC')

        return self.materials[material]

    def get_fy(self, story, bay_id):
        """
        get fy
        """
        section = self.get_section(story, bay_id)
        material = self.sections.get(section, 'FY')

        return self.materials[material]

    def get_fyh(self, story, bay_id):
        """
        get fyh
        """
        section = self.get_section(story, bay_id)
        material = self.sections.get(section, 'FYH')

        return self.materials[material]

    def get_width(self, story, bay_id):
        """
        get width
        """
        section = self.get_section(story, bay_id)

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

    e2k = E2k(config['e2k_path_test_v1'])

    print(e2k.stories)
    print(e2k.materials)
    print(e2k.sections.get())
    print(e2k.point_coordinates.get())
    print(e2k.lines.get())
    print(e2k.point_assigns.get())
    print(e2k.line_assigns.get())
    print(e2k.dead_load_name)
    print(e2k.line_loads.get())

    print(e2k.get_section('3F', 'B1'))
    print(e2k.get_fc('3F', 'B1'))
    print(e2k.get_fy('3F', 'B1'))
    print(e2k.get_fyh('3F', 'B1'))
    print(e2k.get_width('3F', 'B1'))
    print(e2k.get_coordinate(bay_id='B1'))
    print(e2k.get_coordinate(point_id='1'))

    print(e2k.sections.get())
    print(e2k.point_assigns.get())


if __name__ == "__main__":
    main()
