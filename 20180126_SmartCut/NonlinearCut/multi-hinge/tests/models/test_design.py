"""
test
"""
from math import isclose

from src.models.design import Design
from tests.config import config


def test_design():
    """
    design
    """
    # pylint: disable=line-too-long

    design = Design(config['design_path'])

    section = {('樓層', ''): 'RF', ('編號', ''): 'B1', ('RC 梁寬', ''): 60.0, ('RC 梁深', ''): 80.0, ('箍筋', '左'): '2#4@15', ('箍筋', '中'): '#4@12', ('箍筋', '右'): '2#4@15', ('箍筋長度', '左'): 177.5000050663947, ('箍筋長度', '中'): 355.0000101327894, ('箍筋長度', '右'): 177.5000050663947, ('梁長', ''): 800.0, ('支承寬', '左'): 45.0, ('支承寬', '右'): 45.0, ('NOTE', ''): 22142.12094137667}

    assert design.get(1) == section
    assert design.get(3, ('主筋', '左')) == '7-#7'
    assert isclose(design.get(2, ('主筋長度', '左')), 109.9999964237209)
