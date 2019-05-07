"""
get section points
"""
from itertools import product
import numpy as np

from src.utils.get_ld import get_ld


def _get_conservative_stirrup(index, side, design):
    # pylint: disable=invalid-name
    control_position = side

    if design.get(index, ('主筋長度', f'{side}1')) > design.get(index, ('箍筋長度', side)):
        shear_side = design.get_shear(index, ('箍筋', side))
        shear_mid = design.get_shear(index, ('箍筋', '中'))

        if shear_mid > shear_side:
            control_position = '中'

    dh = design.get_diameter(index, ('箍筋', control_position))
    ah = design.get_area(index, ('箍筋', control_position))
    spacing = design.get_spacing(index, ('箍筋', control_position))

    return dh, ah, spacing


def post_side_hinges(points, section_index, design, e2k):
    """
    3 multi hinge, side hinge, consider ld
    """
    # pylint: disable=invalid-name
    positions = [
        ('左1', True, section_index), ('左1', False, section_index + 3),
        ('右1', True, section_index), ('右1', False, section_index + 3)
    ]

    section = design.get(section_index)

    for (left_or_right, top, index) in positions:
        cover = 0.04

        story = section[('樓層', '')]
        bay_id = section[('編號', '')]
        num = design.get_num(index, ('主筋', left_or_right))
        db = design.get_diameter(index, ('主筋', left_or_right))

        fc = e2k.get_fc(story, bay_id)
        fy = e2k.get_fy(story, bay_id)
        fyh = e2k.get_fyh(story, bay_id)
        B = e2k.get_width(story, bay_id)

        dh, ah, spacing = _get_conservative_stirrup(
            index, left_or_right[0], design)

        ld = get_ld(B, num, db, dh, ah, spacing, top, fc, fy, fyh, cover)

        # exclude '支承寬'
        length = design.get(index, ('主筋長度', left_or_right)) / 100 - ld

        # smaller than 10cm, represent too close to boundary
        # then without append to list, direcly continue
        if length < 0.1:
            continue

        # plus '支承寬'
        local_coordinate = length + section[('支承寬', left_or_right[0])] / 100

        # covert to real local coordinate
        # if '右1' should minus '梁長'
        if left_or_right == '右1':
            local_coordinate = section[('梁長', '')] / 100 - local_coordinate

        if not np.allclose(local_coordinate, points, atol=0.1):
            points.append(local_coordinate)


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


def post_left_hinges(hinges, section_index, design, e2k):
    """
    3 multi hinge, side hinge, consider ld
    """
    # pylint: disable=invalid-name
    cover = 0.04

    section = design.get(section_index)

    # ('主筋', '左1')
    col = 5
    length_col = col + design.group_num

    # 最左端
    post_hinge(hinges, 0)

    for loc in ('top', 'bot'):
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
        _, end_length = design.get_abs_length(index, length_col)
        end_length -= 0.01

        rebar_col, end_stirrup_col = design.get_colname_by_length(
            index, end_length)

        mid_area = design.get_total_area(index, col)
        right_area = design.get_total_area(index, col + 1)

        if mid_area > right_area:
            dh = design.get_diameter(index, ('箍筋', end_stirrup_col))
            ah = design.get_area(index, ('箍筋', end_stirrup_col))
            spacing = design.get_spacing(index, ('箍筋', end_stirrup_col))
            ld = get_ld(B, num, db, dh, ah, spacing, top, fc, fy, fyh, cover)
            post_hinge(hinges, end_length - ld)
        elif mid_area < right_area:
            post_hinge(hinges, end_length)


def post_right_hinges(hinges, section_index, design, e2k):
    """
    3 multi hinge, side hinge, consider ld
    """
    # pylint: disable=invalid-name
    cover = 0.04

    group_num = design.get_group_num()

    section = design.get(section_index)

    # ('主筋', '右1')
    col = 5 + group_num - 1
    length_col = col + group_num

    # 最右端
    post_hinge(hinges, section[('梁長', '')] / 100)

    for loc in ('top', 'bot'):
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
        start_length, _ = design.get_abs_length(index, length_col)
        start_length += 0.01

        rebar_col, start_stirrup_col = design.get_colname_by_length(
            index, start_length)

        left_area = design.get_total_area(index, col - 1)
        mid_area = design.get_total_area(index, col)

        if mid_area > left_area:
            dh = design.get_diameter(index, ('箍筋', start_stirrup_col))
            ah = design.get_area(index, ('箍筋', start_stirrup_col))
            spacing = design.get_spacing(index, ('箍筋', start_stirrup_col))
            ld = get_ld(B, num, db, dh, ah, spacing, top, fc, fy, fyh, cover)
            post_hinge(hinges, start_length + ld)
        elif mid_area < left_area:
            post_hinge(hinges, start_length)


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

    # post_left_hinges(hinges, section_index, design, e2k)
    post_mid_hinges(hinges, section_index, design, e2k)
    # post_right_hinges(hinges, section_index, design, e2k)

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
