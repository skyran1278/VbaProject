"""
entry
"""
from src.models.e2k import E2k

path = 'D:/GitHub/VbaProject/20180126_SmartCut/NonlinearCut/multi-hinge/tests/20190103 v3.0 3floor v16.e2k'

e2k = E2k(path)

print(e2k.get('2F', 'B1', fy=True))
