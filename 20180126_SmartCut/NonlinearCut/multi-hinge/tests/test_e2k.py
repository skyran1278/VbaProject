"""
test
"""
from src.models.e2k import E2k


def test_e2k():
    """
    material
    """
    # pylint: disable=line-too-long
    # path = 'D:/GitHub/VbaProject/20180126_SmartCut/NonlinearCut/multi-hinge/tests/20190103 v3.0 3floor v16.e2k'
    path = '/Users/skyran/Documents/GitHub/VbaProject/20180126_SmartCut/NonlinearCut/multi-hinge/tests/20190103 v3.0 3floor v16.e2k'

    e2k = E2k(path)

    stories = {'RF': 3.0, '3F': 3.0, '2F': 3.0, '1F': 3.0, 'BASE': 0.0}
    materials = {('STEEL', 'FY'): 35153.48, ('CONC', 'FC'): 2800.0, ('C350', 'FC'): 3500.0, ('C280', 'FC'): 2800.0, ('C245', 'FC'): 2450.0,
                 ('C210', 'FC'): 2100.0, ('C420', 'FC'): 4200.0, ('C490', 'FC'): 4900.0, ('A615Gr60', 'FY'): 42184.18, ('RMAT', 'FY'): 42000.0}

    sections = {'B60X80C28': {'MATERIAL': 'C280', 'D': 0.8, 'B': 0.6, 'JMOD': 0.0001, 'I2MOD': 0.7, 'I3MOD': 0.7, 'FY': 'RMAT', 'FYT': 'RMAT'},
                'C90X90C28': {'MATERIAL': 'C280', 'D': 0.9, 'B': 0.9, 'JMOD': 0.0001, 'I2MOD': 0.7, 'I3MOD': 0.7}}

    point_coordinates = {'1': [0.0, 0.0], '2': [8.0, 0.0]}
    lines = {'B1': ['1', '2']}
    line_assigns = {('2F', 'B1'): 'B60X80C28', ('2F', 'C1'): 'C90X90C28', ('2F', 'C2'): 'C90X90C28', ('RF', 'B1'): 'B60X80C28', ('RF', 'C1'): 'C90X90C28', ('RF', 'C2'): 'C90X90C28', ('3F', 'B1'): 'B60X80C28', ('3F', 'C1'): 'C90X90C28', ('3F', 'C2'): 'C90X90C28'}

    assert e2k.stories == stories
    assert e2k.materials == materials
    assert e2k.sections == sections
    assert e2k.point_coordinates == point_coordinates
    assert e2k.lines == lines
    assert e2k.line_assigns == line_assigns
