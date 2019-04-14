"""
test
"""
# pylint: disable=redefined-outer-name
from math import isclose

import pytest


@pytest.fixture(scope='module')
def design():
    """
    design config
    """
    from config import config
    from evaluation.design import Design

    return Design(config['design_path_test_v1'])


def test_get(design):
    """
    design
    """
    # pylint: disable=line-too-long

    section = {('樓層', ''): 'PR', ('編號', ''): 'B131', ('RC 梁寬', ''): 50.0, ('RC 梁深', ''): 70.0, ('箍筋', '左'): '#4@30', ('箍筋', '中'): '#4@30', ('箍筋', '右'): '#4@25', ('箍筋長度', '左'): 116.2499949336053, ('箍筋長度', '中'): 232.4999898672106, ('箍筋長度', '右'): 116.2499949336053, ('梁長', ''): 525.0, ('支承寬', '左'): 30.0, ('支承寬', '右'): 30.0, ('NOTE', ''): 4021.194573700431}

    # col is None
    print(design.get(1))
    assert design.get(1) == section

    # col is '主筋'
    assert design.get(3, ('主筋', '左')) == '2-#7'

    # col is '主筋長度'
    assert isclose(design.get(9, ('主筋長度', '中')), 160.000014305115)
    assert isclose(design.get(2, ('主筋長度', '左')), 49.5000004768372)

    # col is others
    assert design.get(3, ('RC 梁寬', '')) == 50


def test_len(design):
    """
    test len
    """
    assert design.get_len() == 1748


def test_total_area(design):
    """
    test total_area
    """
    print(design.get_total_area(11, ('主筋', '中')))
    assert design.get_total_area(11, ('主筋', '中')) == 0.0011613


def test_length_area(design):
    """
    test length_area
    """
    print(design.get_length_area(11, 0.24))
    assert design.get_length_area(11, 0.24) == (0.0019355, 0.0011613)


def test_num(design):
    """
    test num
    """
    assert design.get_num(6, ('主筋', '右')) == 0
    assert design.get_num(8, ('主筋', '右')) == 2


def test_dia(design):
    """
    test dia
    """

    assert design.get_diameter(6, ('箍筋', '右')) == 0.0127
    assert design.get_diameter(6, ('主筋', '右')) == 0


def test_area(design):
    """
    test area
    """

    assert design.get_area(6, ('主筋', '右')) == 0
    assert design.get_area(6, ('箍筋', '右')) == 0.0001267


def test_spacing(design):
    """
    test spacing
    """

    assert design.get_spacing(10, ('箍筋', '右')) == 0.22


def test_shear(design):
    """
    test shear
    """
    print(design.get_shear(10, ('箍筋', '右')))
    assert isclose(design.get_shear(10, ('箍筋', '右')), 0.0005759090909090909)
    print(design.get_shear(10, ('箍筋', '中')))
    assert design.get_shear(10, ('箍筋', '中')) == 0.0008446666666666666
