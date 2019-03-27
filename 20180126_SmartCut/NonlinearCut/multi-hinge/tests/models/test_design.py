"""
test
"""
from math import isclose

import pytest

# pylint: disable=redefined-outer-name


@pytest.fixture(scope='module')
def design():
    """
    design config
    """
    from tests.config import config
    from src.models.design import Design

    return Design(config['design_path'])


def test_get(design):
    """
    design
    """
    # pylint: disable=line-too-long

    section = {('樓層', ''): 'RF', ('編號', ''): 'B1', ('RC 梁寬', ''): 60.0, ('RC 梁深', ''): 80.0, ('箍筋', '左'): '2#4@15', ('箍筋', '中'): '#4@12', ('箍筋', '右'): '2#4@15', (
        '箍筋長度', '左'): 177.5000050663947, ('箍筋長度', '中'): 355.0000101327894, ('箍筋長度', '右'): 177.5000050663947, ('梁長', ''): 800.0, ('支承寬', '左'): 45.0, ('支承寬', '右'): 45.0, ('NOTE', ''): 22142.12094137667}

    # col is None
    assert design.get(1) == section

    # col is '主筋'
    assert design.get(3, ('主筋', '左')) == '7-#7'

    # col is '主筋長度'
    assert isclose(design.get(9, ('主筋長度', '中')), 330.000019073486)
    assert isclose(design.get(2, ('主筋長度', '左')), 109.9999964237209)

    # col is others
    assert design.get(3, ('RC 梁寬', '')) == 60


def test_len(design):
    """
    test len
    """
    assert design.get_len() == 12
