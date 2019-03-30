"""
test
"""
from src.models.lines import Lines


def test_lines():
    """
    test
    """
    lines = Lines()

    assert lines.post(key='B1', value=['1', '2']) == 'B1'
    assert lines.post(value=['2', '3']) == 'B2'
    assert lines.post(value=['1', '2']) == 'B1'

    assert lines.get() == {'B1': ('1', '2'), 'B2': ('2', '3')}
    assert lines.get('B1') == ('1', '2')
