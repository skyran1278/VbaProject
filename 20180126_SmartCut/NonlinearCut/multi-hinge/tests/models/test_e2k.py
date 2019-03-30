"""
test
"""
# pylint: disable=redefined-outer-name
import numpy as np

import pytest


@pytest.fixture(scope='module')
def e2k():
    """
    design config
    """
    from tests.config import config
    from src.models.e2k import E2k

    return E2k(config['e2k_path_test_v1'])


def test_properties(e2k):
    """
    material
    """
    # pylint: disable=line-too-long

    stories = {
        'RF': 3.0, '3F': 3.0, '2F': 3.0, '1F': 3.0, 'BASE': 0.0
    }

    materials = {
        'STEEL': 35153.48,
        'CONC': 2800.0,
        'C350': 3500.0,
        'C280': 2800.0,
        'C245': 2450.0,
        'C210': 2100.0,
        'C420': 4200.0,
        'C490': 4900.0,
        'A615Gr60': 42184.18,
        'RMAT': 42000.0
    }

    sections = {
        'B60X80C28': {
            'FC': 'C280', 'D': 0.8, 'B': 0.6, 'JMOD': 0.0001, 'I2MOD': 0.7, 'I3MOD': 0.7, 'FY': 'RMAT', 'FYH': 'RMAT', 'COVERTOP': 0.08, 'COVERBOTTOM': 0.08, 'ATI': 0, 'ABI': 0, 'ATJ': 0, 'ABJ': 0
        },
        'C90X90C28': {
            'FC': 'C280', 'D': 0.9, 'B': 0.9, 'JMOD': 0.0001, 'I2MOD': 0.7, 'I3MOD': 0.7
        }
    }

    point_coordinates = {
        '1': np.array([0.0, 0.0]),
        '2': np.array([8.0, 0.0])
    }

    lines = {'B1': ('1', '2')}

    line_assigns = {
        ('2F', 'B1'): {'SECTION': 'B60X80C28', 'RIGIDZONE': '0.75', 'CARDINALPT': '8', 'MAXSTASPC': '0.1', 'MESH': 'POINTSANDLINES'},
        ('2F', 'C1'): {'SECTION': 'C90X90C28', 'RIGIDZONE': '0.75', 'MINNUMSTA': '3', 'MESH': 'POINTSANDLINES'},
        ('2F', 'C2'): {'SECTION': 'C90X90C28', 'RIGIDZONE': '0.75', 'MINNUMSTA': '3', 'MESH': 'POINTSANDLINES'},
        ('RF', 'B1'): {'SECTION': 'B60X80C28', 'RIGIDZONE': '0.75', 'CARDINALPT': '8', 'MAXSTASPC': '0.1', 'MESH': 'POINTSANDLINES'},
        ('RF', 'C1'): {'SECTION': 'C90X90C28', 'RIGIDZONE': '0.75', 'MINNUMSTA': '3', 'MESH': 'POINTSANDLINES'},
        ('RF', 'C2'): {'SECTION': 'C90X90C28', 'RIGIDZONE': '0.75', 'MINNUMSTA': '3', 'MESH': 'POINTSANDLINES'},
        ('3F', 'B1'): {'SECTION': 'B60X80C28', 'RIGIDZONE': '0.75', 'CARDINALPT': '8', 'MAXSTASPC': '0.1', 'MESH': 'POINTSANDLINES'},
        ('3F', 'C1'): {'SECTION': 'C90X90C28', 'RIGIDZONE': '0.75', 'MINNUMSTA': '3', 'MESH': 'POINTSANDLINES'},
        ('3F', 'C2'): {'SECTION': 'C90X90C28', 'RIGIDZONE': '0.75', 'MINNUMSTA': '3', 'MESH': 'POINTSANDLINES'}
    }

    assert e2k.stories == stories
    assert e2k.materials == materials
    assert e2k.sections.get() == sections
    np.testing.assert_equal(
        e2k.point_coordinates.get(), point_coordinates
    )
    assert e2k.lines.get() == lines
    assert e2k.line_assigns.get() == line_assigns


def test_method(e2k):
    """
    test method
    """

    assert e2k.get_section('3F', 'B1') == 'B60X80C28'
    assert e2k.get_fc('3F', 'B1') == 2800
    assert e2k.get_fy('3F', 'B1') == 42000
    assert e2k.get_fyh('3F', 'B1') == 42000
    assert e2k.get_width('3F', 'B1') == 0.6
    np.testing.assert_equal(
        e2k.get_coordinate(bay_id='B1'),
        (np.array([0., 0.]), np.array([8., 0.]))
    )
    np.testing.assert_equal(
        e2k.get_coordinate(point_id='1'),
        np.array([0., 0.])
    )
