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


def get_points_rebar_area(index, abs_points, design):
    """
    get length point rebar area
    """
    point_rebars = []

    for abs_length in abs_points:
        top, bot = design.get_length_area(index, abs_length)

        point_rebars.append((top, bot))

    return point_rebars


def main():
    """
    test
    """

    from tests.config import config

    design = Design(config['design_path_test_v2'])

    e2k = E2k(config['e2k_path_test_v2'])

    new_e2k = NewE2k(config['e2k_path_test_v2'])

    for index in range(0, design.get_len(), 4):
        abs_points, rel_points = get_points(index, design, e2k)

        beam = design.get(index)
        story = beam[('樓層', '')]
        bay_id = beam[('編號', '')]

        coordinates = get_points_coordinates(bay_id, rel_points, e2k)

        new_e2k.post_point_coordinates(coordinates)
        bay_ids = new_e2k.post_lines(coordinates)

        section = e2k.get_section(story, bay_id)

        point_rebars = get_points_rebar_area(index, abs_points, design)

        sections = new_e2k.post_sections(section, point_rebars)

        new_e2k.post_line_assigns(bay_ids, sections, copy_from=(story, bay_id))

        # print(abs_points)
        # print(point_rebars)

    print(new_e2k.point_coordinates.get())
    print(new_e2k.lines.get())
    # print(new_e2k.sections.get())
    print(new_e2k.line_assigns.get())

    new_e2k.to_e2k()


if __name__ == "__main__":
    main()
