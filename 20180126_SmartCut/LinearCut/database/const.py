import numpy as np

BEAM_DESIGN = '20181215 nonlinear v3 hinge DLLL.xlsx'
E2K = '20181215 nonlinear v3 hinge DLLL.e2k'

STIRRUP_REBAR = ['#4', '2#4', '2#5', '2#6']
STIRRUP_SPACING = [10, 12, 15, 18, 20, 22, 25, 30]

BAR = {
    'Top': ['#7', '#8', '#10', '#11', '#14'],
    'Bot': ['#7', '#8', '#10', '#11', '#14']
}

DB_SPACING = 1.5

ITERATION_GAP = {
    'Left': np.array([0.2, 0.4]),
    'Right': np.array([0.6, 0.8])
}
