"""
entry
"""
from src.models.e2k import E2k
from src.models.design import Design
from src.models.new_e2k import NewE2k
from src.controllers.get_points import get_points


def get_points_coordinates(bay_id, rel_points, e2k):
    """
    get global coordinates
    """
    coor_start, coor_end = e2k.get_coordinate(bay_id=bay_id)
    return rel_points.reshape(-1, 1) * (coor_end - coor_start) + coor_start


def get_point_rebar_area(index, abs_points, design):
    """
    get length point rebar area
    """
    rebar_points = []

    for abs_length in abs_points:
        left_boundary = (design.get(index, ('主筋長度', '左')) +
                         design.get(index, ('支承寬', '左'))) / 100
        right_boundary = (design.get(index, ('梁長', '')) - (
            design.get(index, ('主筋長度', '右')) + design.get(index, ('支承寬', '右')))) / 100

        if abs_length < left_boundary:
            col = ('主筋', '左')
        elif abs_length > right_boundary:
            col = ('主筋', '右')
        else:
            col = ('主筋', '中')

        top = design.get_total_area(index, col)
        bot = design.get_total_area(index + 2, col)

        rebar_points.append((top, bot))

    return rebar_points


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
        new_e2k.post_lines(coordinates)

        rebar_points = get_point_rebar_area(index, abs_points, design)

        print(rebar_points)

        print(new_e2k.point_coordinates.get())
        print(new_e2k.lines.get())

    new_e2k.to_e2k()


if __name__ == "__main__":
    main()
