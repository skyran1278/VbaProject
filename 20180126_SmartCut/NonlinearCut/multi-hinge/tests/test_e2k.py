"""
test
"""
from app.models.e2k import E2k


def test_e2k():
    """
    material
    """
    # pylint: disable=line-too-long
    path = 'D:/GitHub/VbaProject/20180126_SmartCut/NonlinearCut/multi-hinge/tests/20190103 v3.0 3floor v16.e2k'

    e2k = E2k(path)
    # print(e2k.materials)
    materials = {('STEEL', 'FY'): 35153.48, ('CONC', 'FC'): 2800.0, ('C350', 'FC'): 3500.0, ('C280', 'FC'): 2800.0, ('C245', 'FC'): 2450.0,
                 ('C210', 'FC'): 2100.0, ('C420', 'FC'): 4200.0, ('C490', 'FC'): 4900.0, ('A615Gr60', 'FY'): 42184.18, ('RMAT', 'FY'): 42000.0}

    sections = {'B60X80C28': {'MATERIAL': 'C280', 'D': 0.8, 'B': 0.6, 'JMOD': 0.0001, 'I2MOD': 0.7, 'I3MOD': 0.7, 'FY': 'RMAT',
                              'FYT': 'RMAT'}, 'C90X90C28': {'MATERIAL': 'C280', 'D': 0.9, 'B': 0.9, 'JMOD': 0.0001, 'I2MOD': 0.7, 'I3MOD': 0.7}}

    assert e2k.materials == materials
    assert e2k.sections == sections

    # snapshot.assert_match(e2k.materials)
