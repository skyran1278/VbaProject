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

    # assert False
