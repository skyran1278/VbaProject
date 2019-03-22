"""
entry
"""
from src.models.e2k import E2k
from src.models.design import Design
from src.controllers.get_ld import get_ld


# path = 'D:/GitHub/VbaProject/20180126_SmartCut/NonlinearCut/multi-hinge/tests/20190103 v3.0 3floor v16.e2k'
# path = '/Users/skyran/Documents/GitHub/VbaProject/20180126_SmartCut/NonlinearCut/multi-hinge/tests/20190103 v3.0 3floor v16.e2k'

design = Design(
    'D:/GitHub/VbaProject/20180126_SmartCut/NonlinearCut/multi-hinge/tests/20190312 235751 SmartCut.xlsx')

e2k = E2k('D:/GitHub/VbaProject/20180126_SmartCut/NonlinearCut/multi-hinge/tests/20190103 v3.0 3floor v16.e2k')

positions = [
    {
        'i': ('主筋', '左')
    }
]

for index in range(0, design.get_len(), 4):
    story = design.get_story(index)
    bay_id = design.get_id(index)

    fc = e2k.get_fc(story, bay_id)
    fy = e2k.get_fy(story, bay_id)
    fyh = e2k.get_fyh(story, bay_id)
    B = e2k.get_width(story, bay_id)

    num = design.get_num(index, ('主筋', '左'))
    db = design.get_diameter(index, ('主筋', '左'))

    span = design.get_span(index)

    if span / 4 > :
        pass

    dh = design.get_diameter(index, ('箍筋', '左'))
    ah = design.get_diameter(index, ('箍筋', '左'))

    get_ld(B, num, db, dh, ah, spacing, top, fc, fy, fyh, cover)
