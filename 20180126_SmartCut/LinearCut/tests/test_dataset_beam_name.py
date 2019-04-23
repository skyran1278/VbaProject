"""
test
"""
from tests.const import const
from src.dataset_beam_name import load_beam_name


def test_load_beam_name():
    """
    test load_beam_name
    """

    dataset = load_beam_name(const['beam_name_path'])
    print(dataset)

    assert all(dataset.loc[('RF', 'B1'), ['施工圖編號', '一台梁']] == ['G1-1', 'G1'])
