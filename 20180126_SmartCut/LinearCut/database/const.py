import numpy as np

BEAM_DESIGN = '20190103 v3.0 3floor for v9.xlsx'
E2K = '20190103 v3.0 3floor for v9.e2k'

STIRRUP_REBAR = ['#4', '2#4', '2#5', '2#6']
STIRRUP_SPACING = [10, 12, 15, 18, 20, 22, 25, 30]

BAR = {
    'Top': ['#7', '#8', '#10', '#11', '#14'],
    'Bot': ['#7', '#8', '#10', '#11', '#14']
}

DB_SPACING = 1.5

ITERATION_GAP = {
    'Left': np.array([0.15, 0.45]),
    'Right': np.array([0.55, 0.85])
}
