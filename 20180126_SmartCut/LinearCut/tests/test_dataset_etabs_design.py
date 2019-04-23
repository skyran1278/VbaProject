"""
test
"""
from tests.const import const
from src.dataset_etabs_design import load_beam_design


def test_load_beam_design():
    """
    test load_beam_design
    """

    dataset = load_beam_design(const['etabs_design_path'])
    print(dataset)

    assert all(
        dataset.loc[0, ['Story', 'BayID', 'SecID']] == ['RF', 'B1', 'B50X60'])
