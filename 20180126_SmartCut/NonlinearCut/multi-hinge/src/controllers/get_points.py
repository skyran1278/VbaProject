"""
get section points
"""
import numpy as np

from src.utils.get_ld import get_ld


def _get_conservative_stirrup(index, side, design):
    # pylint: disable=invalid-name
    control_position = side

    if design.get(index, ('主筋長度', side)) > design.get(index, ('箍筋長度', side)):
        shear_side = design.get_shear(index, ('箍筋', side))
        shear_mid = design.get_shear(index, ('箍筋', '中'))

        if shear_mid > shear_side:
            control_position = '中'

    dh = design.get_diameter(index, ('箍筋', control_position))
    ah = design.get_area(index, ('箍筋', control_position))
    spacing = design.get_spacing(index, ('箍筋', control_position))

    return dh, ah, spacing


def get_points(section_index, design, e2k):
    """
    get section points
    """
    # pylint: disable=invalid-name
    positions = [
        ('左1', True, section_index), ('左1', False, section_index + 3),
        ('右1', True, section_index), ('右1', False, section_index + 3)
    ]

    section = design.get(section_index)

    # left and right side hinge
    points = [0, section[('梁長', '')] / 100]

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
            index, left_or_right, design)

        ld = get_ld(B, num, db, dh, ah, spacing, top, fc, fy, fyh, cover)

        # exclude '支承寬'
        length = design.get(index, ('主筋長度', left_or_right)) / 100 - ld

        # smaller than 10cm, represent too close to boundary
        # then without append to list, direcly continue
        if length < 0.1:
            continue

        # plus '支承寬'
        local_coordinate = length + section[('支承寬', left_or_right)] / 100

        # covert to real local coordinate
        # if '右1' should minus '梁長'
        if left_or_right == '右1':
            local_coordinate = section[('梁長', '')] / 100 - local_coordinate

        if not np.allclose(local_coordinate, points, atol=0.1):
            points.append(local_coordinate)

    points = np.sort(points)

    rel_points = points / points[-1]

    return points, rel_points


def main():
    """
    test
    """
    from tests.config import config
    from src.models.e2k import E2k
    from src.models.design import Design

    design = Design(config['design_path'])

    e2k = E2k(config['e2k_path'])

    points = get_points(0, design, e2k)
    print(points)


if __name__ == "__main__":
    main()
