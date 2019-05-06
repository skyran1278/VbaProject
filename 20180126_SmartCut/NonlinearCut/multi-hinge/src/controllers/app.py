"""
entry
"""
from src.models.e2k import E2k
from src.models.design import Design
from src.models.new_e2k import NewE2k
from src.controllers.get_points import get_points


def get_global_coordinates(bay_id, rel_points, e2k):
    """
    get global coordinates
    """
    coor_start, coor_end = e2k.get_coordinate(bay_id=bay_id)
    return rel_points.reshape(-1, 1) * (coor_end - coor_start) + coor_start


def get_points_rebar_area(index, abs_coors, design):
    """
    get length point rebar area
    """
    point_rebars = []

    for abs_length in abs_coors:
        top, bot = design.get_length_area(index, abs_length)

        point_rebars.append((top, bot))

    return point_rebars


def multi():
    """
    多點斷筋
    """
    from tests.config import config

    design = Design(config['design_path_test_v1'])

    e2k = E2k(config['e2k_path_test_v1'])

    new_e2k = NewE2k(config['e2k_path_test_v1'])

    for index in range(0, design.get_len(), 4):
        abs_coors, rel_coors = get_points(index, design, e2k)

        story = design.get(index, ('樓層', ''))
        line_key = design.get(index, ('編號', ''))

        # get point keys
        point_keys = new_e2k.post_point_coordinates(
            get_global_coordinates(line_key, rel_coors, e2k)
        )

        # after post point keys, then post point assigns
        new_e2k.post_point_assigns(point_keys, story)

        # get points rebar
        point_rebars = get_points_rebar_area(index, abs_coors, design)

        # get line keys, by use point keys
        line_keys = new_e2k.post_lines(point_keys)

        # get section
        section_keys = new_e2k.post_sections(
            point_rebars, copy_from=e2k.get_section(story, line_key))

        # then post line assigns
        new_e2k.post_line_assigns(
            line_keys, section_keys, copy_from=(story, line_key))

        # then post hinges
        new_e2k.post_line_hinges(line_keys, story)

        # then post line load
        new_e2k.post_line_loads(line_keys, (story, line_key))

    new_e2k.to_e2k()


def normal():
    """
    傳統斷筋
    """
    from tests.config import config

    design = Design(config['design_path_test_v1'], '傳統斷筋')

    e2k = E2k(config['e2k_path_test_v1'])

    new_e2k = NewE2k(config['e2k_path_test_v1'])

    for index in range(0, design.get_len(), 4):
        story = design.get(index, ('樓層', ''))
        line_key = design.get(index, ('編號', ''))
        line_length = design.get(index, ('梁長', '')) / 100

        # get points rebar
        point_rebars = get_points_rebar_area(
            index, [0, line_length], design)

        # get section
        section_keys = new_e2k.post_sections(
            point_rebars, copy_from=e2k.get_section(story, line_key))

        # then post line assigns
        new_e2k.post_line_assigns(
            [line_key], section_keys, copy_from=(story, line_key))

        # then post hinges
        new_e2k.post_line_hinges([line_key], story)

    new_e2k.to_e2k()


def main():
    """
    test
    """
    # import cProfile
    # import pstats

    # cProfile.run('multi()', 'restats')
    # p = pstats.Stats('restats')
    # p.strip_dirs().sort_stats('cumtime').print_stats(100)
    multi()
    normal()


if __name__ == "__main__":
    main()
