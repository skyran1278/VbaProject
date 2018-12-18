import numpy as np

BEAM_DESIGN = '20181215 nonlinear v3 hinge pushover 0.1m.xlsx'
E2K = '20181215 nonlinear v3 hinge pushover 0.1m.e2k'

STIRRUP_REBAR = ['#4', '2#4', '2#5', '2#6']
STIRRUP_SPACING = [10, 12, 15, 18, 20, 22, 25, 30]

BAR = {
    'Top': ['#8', '#10', '#11', '#14'],
    'Bot': ['#8', '#10', '#11', '#14']
}

DB_SPACING = 1.5

ITERATION_GAP = {
    'Left': np.array([0.15, 0.45]),
    'Right': np.array([0.55, 0.85])
}
