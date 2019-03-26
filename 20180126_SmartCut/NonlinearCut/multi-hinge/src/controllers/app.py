"""
entry
"""
import numpy as np

from src.models.e2k import E2k
from src.models.design import Design
from src.models.new_e2k import NewE2k
from src.controllers.get_points import get_points


def get_points_coordinates(bay_id, rel_points, e2k):
    coor_start, coor_end = e2k.get_coordinate(bay_id=bay_id)
    return rel_points.reshape(-1, 1) * (coor_end - coor_start) + coor_start


def main():
    """
    test
    """

    from tests.config import config

    design = Design(config['design_path'])

    e2k = E2k(config['e2k_path'])

    new_e2k = NewE2k(config['e2k_path'])

    for index in range(0, design.get_len(), 4):
        abs_points, rel_points = get_points(index, design, e2k)

        section = design.get(index)
        # story = section[('樓層', '')]
        bay_id = section[('編號', '')]

        coordinates = get_points_coordinates(bay_id, rel_points, e2k)

        new_e2k.post_point_coordinates(coordinates)

        print(new_e2k.point_coordinates.get())

    new_e2k.to_e2k()


if __name__ == "__main__":
    main()
