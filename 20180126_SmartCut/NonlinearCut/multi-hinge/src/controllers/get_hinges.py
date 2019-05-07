"""
get section points
"""
from itertools import product
import numpy as np

from src.utils.get_ld import get_ld


def post_mid_hinges(hinges, section_index, design, e2k):
    """
    3 multi hinge, side hinge, consider ld
    """
    # pylint: disable=invalid-name
    cover = 0.04

    group_num = design.group_num

    section = design.get(section_index)

    for loc, col in product(('top', 'bot'), range(5, group_num + 5)):
        length_col = col + group_num
        if loc == 'top':
            top = True
            index = section_index
        else:
            top = False
            index = section_index + 3

        story = section[('樓層', '')]
        bay_id = section[('編號', '')]
        num = design.get_num(index, col)
        db = design.get_diameter(index, col)

        fc = e2k.get_fc(story, bay_id)
        fy = e2k.get_fy(story, bay_id)
        fyh = e2k.get_fyh(story, bay_id)
        B = e2k.get_width(story, bay_id)

        # 避免邊界取錯，所以微調 1 公分
        # 有其必要，因為真的會取錯
        start_length, end_length = design.get_abs_length(index, length_col)
        start_length += 0.01
        end_length -= 0.01

        mid_area = design.get_total_area(index, col)

        if col == 5:
            post_hinge(hinges, 0)

        elif mid_area > design.get_total_area(index, col - 1):
            stirrup_col = design.get_colname_by_length(index, start_length)[1]
            dh = design.get_diameter(index, ('箍筋', stirrup_col))
            ah = design.get_area(index, ('箍筋', stirrup_col))
            spacing = design.get_spacing(index, ('箍筋', stirrup_col))
            ld = get_ld(B, num, db, dh, ah, spacing, top, fc, fy, fyh, cover)
            post_hinge(hinges, start_length + ld)
        elif mid_area < design.get_total_area(index, col - 1):
            post_hinge(hinges, start_length)

        if col == group_num + 5 - 1:
            post_hinge(hinges, section[('梁長', '')] / 100)

        elif mid_area > design.get_total_area(index, col + 1):
            stirrup_col = design.get_colname_by_length(index, end_length)[1]
            dh = design.get_diameter(index, ('箍筋', stirrup_col))
            ah = design.get_area(index, ('箍筋', stirrup_col))
            spacing = design.get_spacing(index, ('箍筋', stirrup_col))
            ld = get_ld(B, num, db, dh, ah, spacing, top, fc, fy, fyh, cover)
            post_hinge(hinges, end_length - ld)
        elif mid_area < design.get_total_area(index, col + 1):
            post_hinge(hinges, end_length)


def post_hinge(hinges, hinge):
    """
    post hinge
    """
    if not any(np.isclose(hinge, hinges, atol=0.1)):
        hinges.append(round(hinge, 7))


def get_hinges(section_index, design, e2k):
    """
    get section points
    """
    # pylint: disable=invalid-name
    # need initial value
    hinges = [0]

    section_index = section_index // 4 * 4

    post_mid_hinges(hinges, section_index, design, e2k)

    hinges = np.sort(hinges)

    rel_points = hinges / hinges[-1]

    return hinges, rel_points


def main():
    """
    test
    """
    from tests.config import config
    from src.models.e2k import E2k
    from src.models.design import Design

    design = Design(config['design_path'])

    e2k = E2k(config['e2k_path'])

    points = get_hinges(0, design, e2k)
    print(points)


if __name__ == "__main__":
    main()
