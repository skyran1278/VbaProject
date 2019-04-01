"""
write to new e2k
"""
from src.models.e2k import E2k
from src.models.lines import Lines


class NewE2k(E2k):
    """
    use to write new e2k
    """

    def __init__(self, *args, **kwargs):
        self.f = None
        self.line_hinges = []
        self.new_lines = Lines()
        super(NewE2k, self).__init__(*args, **kwargs)

    def post_point_coordinates(self, coordinates):
        """
        post list of coordinates to point_coordinates
        """
        point_keys = []
        for coor in coordinates:
            point_keys.append(self.point_coordinates.post(value=coor))

        return point_keys

    def post_lines(self, point_keys):
        """
        post list of coordinates to lines
        """
        line_keys = []
        # coor_id = []
        # for coor in coordinates:
        #     coor_id.append(self.point_coordinates.get(value=coor))

        length = len(point_keys) - 1
        index = 0
        while index < length:
            line_keys.append(self.new_lines.post(
                value=[point_keys[index], point_keys[index + 1]]
            ))
            index += 1

        return line_keys

    def post_sections(self, section, rebars):
        """
        post list of coordinates to lines
        """
        new_sections = []

        length = len(rebars) - 1
        index = 0
        while index < length:
            ati = rebars[index][0]
            abi = rebars[index][1]
            atj = rebars[index + 1][0]
            abj = rebars[index + 1][1]

            new_section = f'{section} {ati} {abi} {atj} {abj}'

            data = {
                'ATI': ati,
                'ABI': abi,
                'ATJ': atj,
                'ABJ': abj
            }

            self.sections.post(new_section, data, copy_from=section)

            index += 1

            new_sections.append(new_section)

        return new_sections

    def post_point_assigns(self, points, story):
        """
        combine line and section
        """
        start = self.point_assigns.get(key=(story, points[0]))
        end = self.point_assigns.get(key=(story, points[-1]))

        if start != end:
            print('Warning start key != end key')

        for point in points:
            self.point_assigns.post(
                key=(story, point), copy_from=(story, points[0])
            )

    def post_line_assigns(self, lines, sections, copy_from):
        """
        combine line and section
        """
        story, _ = copy_from

        for line, section in zip(lines, sections):
            self.line_assigns.post(
                key=(story, line),
                value={
                    'SECTION': section
                },
                copy_from=copy_from
            )

    def post_line_hinges(self, lines, story):
        """
        post hinge
        """
        # for line in lines:
        #     self.line_hinges.post((story, line, 0), {
        #         'AUTOHINGETYPE': 'ASCE41-13',
        #         'TABLEITEM': 'Concrete Beams',
        #         'DOF': 'M3',
        #         'CASECOMBO': self.dead_load_name,
        #         'AUTOSUBDIVIDERELLENGTH': '0.02',
        #     })

        # self.line_hinges.post((story, lines[-1], 1), {
        #     'AUTOHINGETYPE': 'ASCE41-13',
        #     'TABLEITEM': 'Concrete Beams',
        #     'DOF': 'M3',
        #     'CASECOMBO': self.dead_load_name,
        #     'AUTOSUBDIVIDERELLENGTH': '0.02',
        # })

        for line in lines:
            self.line_hinges.append((story, line, 'M3', 0))

        self.line_hinges.append((story, lines[-1], 'M3', 1))

    def post_line_loads(self, lines, copy_from):
        """
        post and delete line loads
        """
        story, _ = copy_from
        for line in lines:
            self.line_loads.post((story, line), copy_from=copy_from)

        self.line_loads.delete(copy_from)

    def __frame_sections(self):
        # pylint: disable=invalid-name
        sections = self.sections.get()
        for section in sections:
            fc = sections[section]['FC']
            D = sections[section]['D']
            B = sections[section]['B']
            propertys = sections[section]['PROPERTIES']
            self.f.write(
                f'FRAMESECTION  "{section}"  MATERIAL "{fc}"  '
                f'SHAPE "Concrete Rectangular"  D {D} B {B} '
                f'INCLUDEAUTORIGIDZONEAREA "No"\n'
            )
            self.f.write(
                f'FRAMESECTION  "{section}"  {propertys}\n')

    def __concrete_sections(self):
        # pylint: disable=invalid-name
        sections = self.sections.get()
        for section in sections:
            fy = sections[section]['FY']
            fyh = sections[section]['FYH']
            cover_top = sections[section]['COVERTOP']
            cover_bot = sections[section]['COVERBOTTOM']
            ati = sections[section]['ATI']
            abi = sections[section]['ABI']
            atj = sections[section]['ATJ']
            abj = sections[section]['ABJ']
            self.f.write(
                f'CONCRETESECTION  "{section}"  LONGBARMATERIAL "{fy}"  '
                f'CONFINEBARMATERIAL "{fyh}"  TYPE "Beam"  COVERTOP {cover_top} '
                f'COVERBOTTOM {cover_bot} ATI {ati} ABI {abi} ATJ {atj} ABJ {abj}\n'
            )

    def __point_coordinates(self):
        coor = self.point_coordinates.get()
        for point in coor:
            start = coor[point][0]
            end = coor[point][1]
            self.f.write(f'POINT "{point}"  {start} {end}\n')

    def __line_connectivities(self):
        columns = self.columns.get()
        beams = self.new_lines.get()
        for column in columns:
            start, end = columns[column]
            self.f.write(f'LINE  "{column}"  COLUMN  "{start}"  "{end}"  1\n')
        for beam in beams:
            start, end = beams[beam]
            self.f.write(f'LINE  "{beam}"  BEAM  "{start}"  "{end}"  0\n')

    def __point_assigns(self):
        point_assigns = self.point_assigns.get()
        for story, key in point_assigns:
            point_property = point_assigns[(story, key)]
            self.f.write(
                f'POINTASSIGN  "{key}"  "{story}"  {point_property}\n'
            )

    def __line_assigns(self):
        line_assigns = self.line_assigns.get()
        for story, key in line_assigns:
            section = line_assigns[(story, key)]['SECTION']
            properties = line_assigns[(story, key)]['PROPERTIES']
            self.f.write(
                f'LINEASSIGN  "{key}"  "{story}"  SECTION "{section}"  {properties}\n'
            )

    def __frame_hinge_assignments(self):
        self.f.write(f'\n$ FRAME HINGE ASSIGNMENTS\n')
        load = self.dead_load_name
        for hinge in self.line_hinges:
            story, line, dof, rdistance = hinge
            self.f.write(
                f'HINGEASSIGN "{line}"  "{story}"  AUTOHINGETYPE "ASCE41-13"  '
                f'TABLEITEM "Concrete Beams"  DOF "{dof}"  '
                f'CASECOMBO "{load}"  RDISTANCE {rdistance}\n'
            )

    def __frame_hinge_overwrites(self):
        self.f.write(f'\n$ FRAME HINGE OVERWRITES\n')
        for hinge in self.line_hinges:
            story, line, _, _ = hinge
            self.f.write(
                f'HINGEOVERWRITE "{line}"  "{story}"  AUTOSUBDIVIDERELLENGTH 0.02\n'
            )

    def to_e2k(self):
        """
        only call once, write to e2k
        """
        with open(self.path + ' new.e2k', mode='w', encoding='big5') as self.f:
            for line in self.content:
                # skip space line
                if line == '':
                    self.f.write('\n')
                    continue

                if line[0] == '$':
                    # write permission
                    write = True

                if write:
                    self.f.write(line)
                    self.f.write('\n')

                if line == '$ FRAME SECTIONS':
                    self.__frame_sections()

                elif line == '$ CONCRETE SECTIONS':
                    self.__concrete_sections()

                elif line == '$ POINT COORDINATES':
                    write = False
                    self.__point_coordinates()

                elif line == '$ LINE CONNECTIVITIES':
                    write = False
                    self.__line_connectivities()

                elif line == '$ POINT ASSIGNS':
                    write = False
                    self.__point_assigns()

                elif line == '$ LINE ASSIGNS':
                    write = False
                    self.__line_assigns()
                    self.__frame_hinge_assignments()
                    self.__frame_hinge_overwrites()


def main():
    """
    test
    """
    from tests.config import config

    new_e2k = NewE2k(config['e2k_path_test_v1'])

    coordinates = [
        [0., 0.],
        [0.67445007, 0.],
        [0.87367754, 0.],
        [7.12632229, 0.],
        [7.32554951, 0.],
        [8., 0.]
    ]

    point_rebars = [
        (0.0046452, 0.0027097),
        (0.0046452, 0.0027097),
        (0.0046452, 0.0027097),
        (0.0046452, 0.0027097),
        (0.0046452, 0.0027097),
        (0.0046452, 0.0027097)
    ]

    point_keys = new_e2k.post_point_coordinates(coordinates)
    print(point_keys)

    line_keys = new_e2k.post_lines(point_keys)
    print(line_keys)

    section_keys = new_e2k.post_sections('B60X80C28', point_rebars)
    print(section_keys)

    new_e2k.post_point_assigns(point_keys, story='RF')

    new_e2k.post_line_assigns(
        line_keys, section_keys, copy_from=('RF', 'B1'))

    new_e2k.post_line_hinges(line_keys, story='RF')

    new_e2k.post_line_loads(line_keys, ('RF', 'B1'))

    print(new_e2k.point_coordinates.get())
    print(new_e2k.lines.get())
    print(new_e2k.sections.get())
    print(new_e2k.point_assigns.get())
    print(new_e2k.line_assigns.get())
    print(new_e2k.line_hinges)
    print(new_e2k.line_loads.get())

    new_e2k.to_e2k()
    print(new_e2k.point_assigns.get())

    print(new_e2k.line_assigns.get())


if __name__ == "__main__":
    main()
