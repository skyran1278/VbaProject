import os

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

from dataset_beam_design import load_beam_design


beam_design_table = load_beam_design()

index = pd.MultiIndex.from_tuples([('樓層', ''), ('編號', ''), ('RC 梁寬', ''), ('RC 梁深', ''), ('主筋', ''), ('主筋', '左'), ('主筋', '中'), ('主筋', '右'), (
    '長度', '左'), ('長度', '中'), ('長度', '右'), ('腰筋', ''), ('箍筋', '左'), ('箍筋', '中'), ('箍筋', '右'), ('梁長', ''), ('支承寬', ''), ('NOTE', ''), ('MESSAGE', '')])

beam_3points_table = pd.DataFrame(
    np.empty([len(beam_design_table.groupby(['Story', 'BayID'])) * 4, 19], dtype='<U16'), columns=index)


def init_beam_3points_table():
    return beam_3points_table
# beam_3points_table.to_excel(dataset_dir + '/3pionts.xlsx')
