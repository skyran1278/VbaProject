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
        post list of coordinates
        """
        for coor in coordinates:
            self.point_coordinates.post(value=coor)
            # print(np.isin(coor, self.point_coordinates.values()))
            # if not np.any(np.isin(coor, self.point_coordinates.values())):
            #     i = 1
            #     x = str(i)
            #     while x in self.point_coordinates:
            #         i += 1
            #         x = str(i)
            #     self.point_coordinates[x] = coor


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


if __name__ == "__main__":
    main()
