"""
test
"""
import numpy as np

from src.models.e2k import E2k
from tests.config import config


def test_e2k():
    """
    material
    """
    # pylint: disable=line-too-long

    e2k = E2k(config['e2k_path'])

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
            'FC': 'C280', 'D': 0.8, 'B': 0.6, 'JMOD': 0.0001, 'I2MOD': 0.7, 'I3MOD': 0.7, 'FY': 'RMAT', 'FYH': 'RMAT'
        },
        'C90X90C28': {
            'FC': 'C280', 'D': 0.9, 'B': 0.9, 'JMOD': 0.0001, 'I2MOD': 0.7, 'I3MOD': 0.7
        }
    }

    point_coordinates = {
        '1': np.array([0.0, 0.0]),
        '2': np.array([8.0, 0.0])
    }

    lines = {'B1': ['1', '2']}

    line_assigns = {
        ('2F', 'B1'): 'B60X80C28',
        ('2F', 'C1'): 'C90X90C28',
        ('2F', 'C2'): 'C90X90C28',
        ('RF', 'B1'): 'B60X80C28',
        ('RF', 'C1'): 'C90X90C28',
        ('RF', 'C2'): 'C90X90C28',
        ('3F', 'B1'): 'B60X80C28',
        ('3F', 'C1'): 'C90X90C28',
        ('3F', 'C2'): 'C90X90C28'
    }

    assert e2k.stories == stories
    assert e2k.materials == materials
    assert e2k.sections.get() == sections
    np.testing.assert_equal(
        e2k.point_coordinates.get(), point_coordinates)
    assert e2k.lines.get() == lines
    assert e2k.line_assigns == line_assigns
