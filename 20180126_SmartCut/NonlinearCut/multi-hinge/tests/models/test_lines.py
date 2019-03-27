"""
test
"""
from src.models.lines import Lines


def test_lines():
    """
    test
    """
    lines = Lines()

    lines.post(key='B1', value=['1', '2'])
    lines.post(value=['2', '3'])
    lines.post(value=['1', '2'])

    assert lines.get() == {'B1': ['1', '2'], 'B2': ['2', '3']}
    assert lines.get('B1') == ['1', '2']
