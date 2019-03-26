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
        coor_id = []
        for coor in coordinates:
            coor_id.append(self.point_coordinates.get(value=coor))
            # self.lines.post(value=coor)

        length = len(coor_id) - 1
        index = 0
        while index < length:
            self.lines.post(value=[coor_id[index], coor_id[index + 1]])
            index += 1

    def _point_coordinates_to_e2k(self, f):
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
                    self._point_coordinates_to_e2k(f)

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

    # a = {
    #     '1': [1, 2],
    #     '2': 2
    # }

    # print(a.values())

    # if [1, 2] in a.values():
    #     print('ok')

    new_e2k.post_point_coordinates(coordinates)
    print(new_e2k.point_coordinates.get())
    new_e2k.to_e2k()
    new_e2k.post_lines(coordinates)


if __name__ == "__main__":
    main()
