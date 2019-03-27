"""
test
"""
from src.models.sections import Sections


def test_sections():
    """
    test post
    """
    sections = Sections()

    data = {
        'FY': 42000,
        'FYH': 42000,
        'FC': 2800
    }

    sections.post(section='B60', keys='FYH', values=28000)

    assert sections.get('B60', 'FYH') == 28000

    sections.post(
        section='B60',
        keys=('FY', 'FC', 'B'),
        values=(42000, 2800, 0.6)
    )

    assert sections.get('B60') == {
        'FY': 42000,
        'FC': 2800,
        'B': 0.6,
        'FYH': 28000
    }

    sections.post(section='B60', data=data)

    assert sections.get('B60', 'FYH') == 42000

    sections.copy(new_section='B601', copy_from='B60', data={'D': 0.8})

    assert sections.get('B601', 'FYH') == 42000
    assert sections.get('B601', 'D') == 0.8
