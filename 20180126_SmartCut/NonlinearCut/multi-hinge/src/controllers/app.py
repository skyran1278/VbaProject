"""
entry
"""
from src.models.e2k import E2k
from src.models.design import Design


# path = 'D:/GitHub/VbaProject/20180126_SmartCut/NonlinearCut/multi-hinge/tests/20190103 v3.0 3floor v16.e2k'
# path = '/Users/skyran/Documents/GitHub/VbaProject/20180126_SmartCut/NonlinearCut/multi-hinge/tests/20190103 v3.0 3floor v16.e2k'

design = Design(
    'D:/GitHub/VbaProject/20180126_SmartCut/NonlinearCut/multi-hinge/tests/20190312 235751 SmartCut.xlsx')

e2k = E2k('D:/GitHub/VbaProject/20180126_SmartCut/NonlinearCut/multi-hinge/tests/20190103 v3.0 3floor v16.e2k')

print(e2k.get('2F', 'B1', fy=True))
print(e2k.get('2F', 'B1', fyh=True))
print(e2k.get('2F', 'B1', fc=True))

for index in range(0, design.get_len(), 4):
    design[]
