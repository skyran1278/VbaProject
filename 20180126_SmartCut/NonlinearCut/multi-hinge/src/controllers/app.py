"""
entry
"""
from src.models.e2k import E2k
from src.models.design import Design
from src.controllers.get_ld import get_ld


# path = 'D:/GitHub/VbaProject/20180126_SmartCut/NonlinearCut/multi-hinge/tests/20190103 v3.0 3floor v16.e2k'
# path = '/Users/skyran/Documents/GitHub/VbaProject/20180126_SmartCut/NonlinearCut/multi-hinge/tests/20190103 v3.0 3floor v16.e2k'

design = Design(
    'D:/GitHub/VbaProject/20180126_SmartCut/NonlinearCut/multi-hinge/tests/20190323 203316 SmartCut.xlsx')

e2k = E2k('D:/GitHub/VbaProject/20180126_SmartCut/NonlinearCut/multi-hinge/tests/20190103 v3.0 3floor v16.e2k')

positions = [
    {
        'i': ('主筋', '左')
    }
]

for index in range(0, design.get_len(), 4):
    story = design.get(index, ('樓層', ''))
    bay_id = design.get(index, ('編號', ''))
    num = design.get_num(index, ('主筋', '左'))
    db = design.get_diameter(index, ('主筋', '左'))
    spacing = design.get_spacing(index, ('箍筋', '左'))

    fc = e2k.get_fc(story, bay_id)
    fy = e2k.get_fy(story, bay_id)
    fyh = e2k.get_fyh(story, bay_id)
    B = e2k.get_width(story, bay_id)

    top = True

    cover = 0.04

    if design.get(index, ('主筋長度', '左')) <= design.get(index, ('箍筋長度', '左')):
        dh = design.get_diameter(index, ('箍筋', '左'))
        ah = design.get_area(index, ('箍筋', '左'))
        spacing = design.get_spacing(index, ('箍筋', '左'))

    else:
        shear = design.get_shear(index, ('箍筋', '左'))
        shear_mid = design.get_shear(index, ('箍筋', '中'))

        if shear < shear_mid:
            dh = design.get_diameter(index, ('箍筋', '左'))
            ah = design.get_area(index, ('箍筋', '左'))
            spacing = design.get_spacing(index, ('箍筋', '左'))

        else:
            dh = design.get_diameter(index, ('箍筋', '中'))
            ah = design.get_area(index, ('箍筋', '中'))
            spacing = design.get_spacing(index, ('箍筋', '中'))

    a = get_ld(B, num, db, dh, ah, spacing, top, fc, fy, fyh, cover)
    print(a)
