"""
test
"""
# pylint: disable=redefined-outer-name
import numpy as np

import pytest


@pytest.fixture(scope='module')
def new_e2k():
    """
    new_e2k config
    """
    from src.models.new_e2k import NewE2k
    from tests.config import config

    return NewE2k(config['e2k_path_test_v1'])


def test_post_point_coordinates(new_e2k):
    """
    test post coor correct
    """
    coordinates = [
        [0., 0.],
        [0.67445007, 0.],
        [0.87367754, 0.],
        [7.12632229, 0.],
        [7.32554951, 0.],
        [8., 0.]
    ]

    assert new_e2k.post_point_coordinates(
        coordinates) == ['1', '3', '4', '5', '6', '2']

    point_coordinates = {
        '1': np.array([0., 0.]),
        '2': np.array([8., 0.]),
        '3': np.array([0.67445007, 0.]),
        '4': np.array([0.87367754, 0.]),
        '5': np.array([7.12632229, 0.]),
        '6': np.array([7.32554951, 0.])
    }

    np.testing.assert_equal(new_e2k.point_coordinates.get(), point_coordinates)


def test_post_lines(new_e2k):
    """
    test
    """
    point_keys = ['1', '3', '4', '5', '6', '2']

    assert new_e2k.post_lines(point_keys) == ['B2', 'B3', 'B4', 'B5', 'B6']

    lines = {
        'B1': ('1', '2'),
        'B2': ('1', '3'),
        'B3': ('3', '4'),
        'B4': ('4', '5'),
        'B5': ('5', '6'),
        'B6': ('6', '2')
    }

    assert new_e2k.lines.get() == lines


def test_post_sections(new_e2k):
    """
    test post_sections
    """

    point_rebars = [
        (0.0046452, 0.0027097),
        (0.0046452, 0.0027097),
        (0.0046452, 0.0027097),
        (0.0046452, 0.0027097),
        (0.0046452, 0.0027097),
        (0.0046452, 0.0027097)
    ]

    data = {
        'B60X80C28': {
            'FC': 'C280', 'D': 0.8, 'B': 0.6, 'JMOD': 0.0001, 'I2MOD': 0.7, 'I3MOD': 0.7,
            'FY': 'RMAT', 'FYH': 'RMAT', 'COVERTOP': 0.08, 'COVERBOTTOM': 0.08,
            'ATI': 0.0, 'ABI': 0.0, 'ATJ': 0.0, 'ABJ': 0.0
        },
        'C90X90C28': {'FC': 'C280', 'D': 0.9, 'B': 0.9, 'JMOD': 0.0001, 'I2MOD': 0.7, 'I3MOD': 0.7},
        'B60X80C28 0.0046452 0.0027097 0.0046452 0.0027097': {
            'FC': 'C280', 'D': 0.8, 'B': 0.6, 'JMOD': 0.0001, 'I2MOD': 0.7, 'I3MOD': 0.7,
            'FY': 'RMAT', 'FYH': 'RMAT', 'COVERTOP': 0.08, 'COVERBOTTOM': 0.08,
            'ATI': 0.0046452, 'ABI': 0.0027097, 'ATJ': 0.0046452, 'ABJ': 0.0027097
        }
    }

    section_keys = [
        'B60X80C28 0.0046452 0.0027097 0.0046452 0.0027097',
        'B60X80C28 0.0046452 0.0027097 0.0046452 0.0027097',
        'B60X80C28 0.0046452 0.0027097 0.0046452 0.0027097',
        'B60X80C28 0.0046452 0.0027097 0.0046452 0.0027097',
        'B60X80C28 0.0046452 0.0027097 0.0046452 0.0027097'
    ]

    assert new_e2k.post_sections('B60X80C28', point_rebars) == section_keys

    assert new_e2k.sections.get() == data


def test_point_assigns(new_e2k):
    """
    test point_assigns
    """
    point_keys = ['1', '3', '4', '5', '6', '2']

    data = {
        ('2F', '1'): {'DIAPH': 'D1'},
        ('2F', '2'): {'DIAPH': 'D1'},
        ('1F', '1'): {'RESTRAINT': 'UX UY UZ RX RY RZ', 'DIAPH': 'D1'},
        ('1F', '2'): {'RESTRAINT': 'UX UY UZ RX RY RZ', 'DIAPH': 'D1'},
        ('RF', '1'): {'DIAPH': 'D1', 'USERJOINT': 'Yes'},
        ('RF', '2'): {'DIAPH': 'D1', 'USERJOINT': 'Yes'},
        ('3F', '1'): {'DIAPH': 'D1', 'USERJOINT': 'Yes'},
        ('3F', '2'): {'DIAPH': 'D1', 'USERJOINT': 'Yes'},
        ('RF', '3'): {'DIAPH': 'D1', 'USERJOINT': 'Yes'},
        ('RF', '4'): {'DIAPH': 'D1', 'USERJOINT': 'Yes'},
        ('RF', '5'): {'DIAPH': 'D1', 'USERJOINT': 'Yes'},
        ('RF', '6'): {'DIAPH': 'D1', 'USERJOINT': 'Yes'}
    }

    new_e2k.post_point_assigns(point_keys, story='RF')

    assert new_e2k.point_assigns.get() == data


def test_line_assigns(new_e2k):
    """
    test line_assigns
    """
    line_keys = ['B2', 'B3', 'B4', 'B5', 'B6']
    section_keys = [
        'B60X80C28 0.0046452 0.0027097 0.0046452 0.0027097',
        'B60X80C28 0.0046452 0.0027097 0.0046452 0.0027097',
        'B60X80C28 0.0046452 0.0027097 0.0046452 0.0027097',
        'B60X80C28 0.0046452 0.0027097 0.0046452 0.0027097',
        'B60X80C28 0.0046452 0.0027097 0.0046452 0.0027097'
    ]
    data = {
        ('2F', 'B1'): {
            'SECTION': 'B60X80C28',
            'RIGIDZONE': '0.75', 'CARDINALPT': '8', 'MAXSTASPC': '0.1', 'MESH': 'POINTSANDLINES'},
        ('2F', 'C1'): {
            'SECTION': 'C90X90C28',
            'RIGIDZONE': '0.75', 'MINNUMSTA': '3', 'MESH': 'POINTSANDLINES'
        },
        ('2F', 'C2'): {
            'SECTION': 'C90X90C28',
            'RIGIDZONE': '0.75', 'MINNUMSTA': '3', 'MESH': 'POINTSANDLINES'
        },
        ('RF', 'B1'): {
            'SECTION': 'B60X80C28',
            'RIGIDZONE': '0.75', 'CARDINALPT': '8', 'MAXSTASPC': '0.1', 'MESH': 'POINTSANDLINES'
        },
        ('RF', 'C1'): {
            'SECTION': 'C90X90C28',
            'RIGIDZONE': '0.75', 'MINNUMSTA': '3', 'MESH': 'POINTSANDLINES'
        },
        ('RF', 'C2'): {
            'SECTION': 'C90X90C28',
            'RIGIDZONE': '0.75', 'MINNUMSTA': '3', 'MESH': 'POINTSANDLINES'
        },
        ('3F', 'B1'): {
            'SECTION': 'B60X80C28',
            'RIGIDZONE': '0.75', 'CARDINALPT': '8', 'MAXSTASPC': '0.1', 'MESH': 'POINTSANDLINES'
        },
        ('3F', 'C1'): {
            'SECTION': 'C90X90C28',
            'RIGIDZONE': '0.75', 'MINNUMSTA': '3', 'MESH': 'POINTSANDLINES'
        },
        ('3F', 'C2'): {
            'SECTION': 'C90X90C28',
            'RIGIDZONE': '0.75', 'MINNUMSTA': '3', 'MESH': 'POINTSANDLINES'
        },
        ('RF', 'B2'): {
            'SECTION': 'B60X80C28 0.0046452 0.0027097 0.0046452 0.0027097',
            'RIGIDZONE': '0.75', 'CARDINALPT': '8', 'MAXSTASPC': '0.1', 'MESH': 'POINTSANDLINES'
        },
        ('RF', 'B3'): {
            'SECTION': 'B60X80C28 0.0046452 0.0027097 0.0046452 0.0027097',
            'RIGIDZONE': '0.75', 'CARDINALPT': '8', 'MAXSTASPC': '0.1', 'MESH': 'POINTSANDLINES'
        },
        ('RF', 'B4'): {
            'SECTION': 'B60X80C28 0.0046452 0.0027097 0.0046452 0.0027097',
            'RIGIDZONE': '0.75', 'CARDINALPT': '8', 'MAXSTASPC': '0.1', 'MESH': 'POINTSANDLINES'
        },
        ('RF', 'B5'): {
            'SECTION': 'B60X80C28 0.0046452 0.0027097 0.0046452 0.0027097',
            'RIGIDZONE': '0.75', 'CARDINALPT': '8', 'MAXSTASPC': '0.1', 'MESH': 'POINTSANDLINES'
        },
        ('RF', 'B6'): {
            'SECTION': 'B60X80C28 0.0046452 0.0027097 0.0046452 0.0027097',
            'RIGIDZONE': '0.75', 'CARDINALPT': '8', 'MAXSTASPC': '0.1', 'MESH': 'POINTSANDLINES'
        }
    }

    new_e2k.post_line_assigns(
        line_keys, section_keys, copy_from=('RF', 'B1'))
