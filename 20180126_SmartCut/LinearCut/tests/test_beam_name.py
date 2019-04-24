"""
test
"""
from tests.const import const
from src.beam_name import load_beam_name, init_beam_name


def test_load_beam_name():
    """
    test
    """

    dataset = load_beam_name(const['beam_name_path'])
    print(dataset)

    assert all(dataset.loc[('RF', 'B1'), ['施工圖編號', '一台梁']] == ['G1-1', 'G1'])


def test_init_beam_name():
    """
    test
    """
    from src.etabs_design import load_etabs_design

    etabs_design = load_etabs_design(const['etabs_design_path'])
    beam_name_empty = init_beam_name(etabs_design)
    print(beam_name_empty.head())

    assert all(beam_name_empty.loc[0, :] == ['RF', 'B1', '', ''])
