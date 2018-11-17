import numpy as np

E2K = '2018-0214'
BEAM_DESIGN = 'Concrete Design 2 Beam Summary Data ACI 318-05 IBC 2003'

STIRRUP_REBAR = ['#4', '2#4', '2#5', '2#6']
STIRRUP_SPACING = [10, 12, 15, 18, 20, 25, 30]

BAR = {
    'Top': ['#7', '#8', '#10', '#11'],
    'Bot': ['#7', '#8', '#10', '#11']
}
DB_SPACING = 1.5

ITERATION_GAP = {
    'Left': np.array([0.15, 0.4]),
    'Right': np.array([0.6, 0.85])
}
