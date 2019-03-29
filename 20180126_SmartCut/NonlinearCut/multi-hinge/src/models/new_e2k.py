"""
write to new e2k
"""
from src.models.e2k import E2k


class NewE2k(E2k):
    """
    use to write new e2k
    """

    def post_point_coordinates(self, coordinates):
        """
        post list of coordinates to point_coordinates
        """
        for coor in coordinates:
            self.point_coordinates.post(value=coor)

    def post_lines(self, coordinates):
        """
        post list of coordinates to lines
        """
        line_keys = []
        coor_id = []
        for coor in coordinates:
            coor_id.append(self.point_coordinates.get(value=coor))

        length = len(coor_id) - 1
        index = 0
        while index < length:
            line_keys.append(
                self.lines.post(value=[coor_id[index], coor_id[index + 1]])
            )
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

    def post_line_assigns(self, lines, sections, copy_from):
        """
        combine line and section
        """
        story, _ = copy_from

        for line, section in zip(lines, sections):
            self.line_assigns.post(
                key=(story, line),
                data={
                    'SECTION': section
                },
                copy_from=copy_from
            )

    def __point_coordinates_to_e2k(self, f):
        f.write('$ POINT COORDINATES')
        f.write('\n')
        coor = self.point_coordinates.get()
        for point in coor:
            start = coor[point][0]
            end = coor[point][1]
            f.write(f'POINT "{point}"  {start} {end}')
            f.write('\n')

    def to_e2k(self):
        """
        only call once, write to e2k
        """
        with open(self.path + ' new.e2k', mode='w', encoding='big5') as f:
            for line in self.content:
                # skip space line
                if line == '':
                    f.write('\n')
                    continue

                if line[0] == '$':
                    # post title
                    title = line

                # write once
                if line == '$ POINT COORDINATES':
                    self.__point_coordinates_to_e2k(f)

                if not title == '$ POINT COORDINATES':
                    f.write(line)
                    f.write('\n')


def main():
    """
    test
    """
    from tests.config import config

    new_e2k = NewE2k(config['e2k_path'])

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

    new_e2k.post_point_coordinates(coordinates)
    new_e2k.post_lines(coordinates)
    new_e2k.post_sections('B60X80C28', point_rebars)
    print(new_e2k.point_coordinates.get())
    print(new_e2k.lines.get())
    print(new_e2k.sections.get())
    new_e2k.to_e2k()


if __name__ == "__main__":
    main()
