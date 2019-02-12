""" const from user. """
import numpy as np

# from gui import FRAME

ETABS_DESIGN_PATH = '/Users/skyran/Documents/GitHub/VbaProject/20180126_SmartCut/LinearCut/data/20190103 v3.0 3floor for v9.xlsx'
E2K_PATH = '/Users/skyran/Documents/GitHub/VbaProject/20180126_SmartCut/LinearCut/data/20190103 v3.0 3floor for v9.e2k'

BEAM_NAME_PATH = '/Users/skyran/Documents/GitHub/VbaProject/20180126_SmartCut/LinearCut/out/first_run IDA #8 v1.0.xlsx'

STIRRUP_REBAR = ['#4', '2#4', '2#5', '2#6']
STIRRUP_SPACING = np.array([10, 12, 15, 18, 20, 22, 25, 30])

BAR = {
    'Top': ['#7', '#8', '#10', '#11', '#14'],
    'Bot': ['#7', '#8', '#10', '#11', '#14']
}

DB_SPACING = 1.5

ITERATION_GAP = {
    'Left': np.array([0.15, 0.45]),
    'Right': np.array([0.55, 0.85])
}

COVER = 0.04
