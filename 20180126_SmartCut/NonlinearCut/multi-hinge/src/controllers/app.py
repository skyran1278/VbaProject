"""
entry
"""
import numpy as np

from src.models.e2k import E2k
from src.models.design import Design
from src.models.new_e2k import NewE2k
from src.controllers.get_points import get_points


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

        coordinate_start, coordinate_end = e2k.get_coordinate(bay_id=bay_id)
        dif = rel_points.reshape(-1, 1) * (
            coordinate_end - coordinate_start) + coordinate_start

        new_e2k.post_point_coordinates(dif)
        print(dif)


if __name__ == "__main__":
    main()
