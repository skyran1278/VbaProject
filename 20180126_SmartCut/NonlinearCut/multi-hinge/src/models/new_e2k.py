"""
write to new e2k
"""
import shlex
from src.models.e2k import E2k


class NewE2k(E2k):
    """
    use to write new e2k
    """

    def post_point_coordinates(self, coordinates):
        """
        post list of coordinates
        """
        for coor in coordinates:
            self.point_coordinates.post(value=coor)

    def to_e2k(self):
        """
        only call once, write to e2k
        """
        with open(self.path + ' new.e2k', mode='w', encoding='big5') as f:
            for line in self.content:
                # skip space line
                if line == '':
                    f.write(line)
                    f.write('\n')
                    continue

                if line[0] == '$':
                    # post title
                    title = line

                if line == '$ POINT COORDINATES':
                    f.write('$ POINT COORDINATES')
                    f.write('\n')
                    a = self.point_coordinates.get()
                    for point in a:
                        f.write(
                            f'POINT "{point}"  {a[point][0]} {a[point][1]}')
                        f.write('\n')

                        print(f'POINT "{point}"  {a[point][0]} {a[point][1]}')

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


if __name__ == "__main__":
    main()
