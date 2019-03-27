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

    sections.post('B60', {'FYH': 28000})

    assert sections.get('B60', 'FYH') == 28000

    sections.post(section='B60', data=data)

    assert sections.get('B60', 'FYH') == 42000

    sections.post(section='B601', data={'D': 0.8}, copy_from='B60')

    assert sections.get('B601', 'FYH') == 42000
    assert sections.get('B601', 'D') == 0.8
