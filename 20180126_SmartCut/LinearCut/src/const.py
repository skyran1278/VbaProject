""" const from user. """
import os

import numpy as np

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

const = {
    'etabs_design_path': SCRIPT_DIR + '/tests/20190103 v3.0 3floor for v9.xlsx',
    'e2k_path': SCRIPT_DIR + '/tests/20190103 v3.0 3floor for v9.e2k',

    'beam_name_path': SCRIPT_DIR + '/tests/first_run ida #8 v1.0.xlsx',

    'output_dir': SCRIPT_DIR + '/tests',

    'stirrup_rebar': np.array(['#4', '2#4', '2#5', '2#6']),
    'stirrup_spacing': np.array([10, 12, 15, 18, 20, 22, 25, 30]),

    'rebar': {
        'Top': np.array(['#7', '#8', '#10', '#11', '#14']),
        'Bot': np.array(['#7', '#8', '#10', '#11', '#14'])
    },

    'db_spacing': 1.5,

    'iteration_gap': {
        'left': np.array([0.15, 0.45]),
        'right': np.array([0.55, 0.85])
    },

    'cover': 0.04,
}

# ETABS_DESIGN_PATH = SCRIPT_DIR + '/tests/20190103 v3.0 3floor for v9.xlsx'

# # ETABS_DESIGN_PATH = FRAME.PANEL.get_etabs_design_path()

# # E2K_PATH = FRAME.PANEL.e2k_path

# # BEAM_NAME_PATH = FRAME.PANEL.beam_name_path

# # OUTPUT_DIR = FRAME.PANEL.output_dir

# E2K_PATH = SCRIPT_DIR + '/tests/20190103 v3.0 3floor for v9.e2k'

# BEAM_NAME_PATH = SCRIPT_DIR + '/tests/first_run IDA #8 v1.0.xlsx'

# OUTPUT_DIR = SCRIPT_DIR + '/tests'

# STIRRUP_REBAR = ['#4', '2#4', '2#5', '2#6']
# STIRRUP_SPACING = np.array([10, 12, 15, 18, 20, 22, 25, 30])

# BAR = {
#     'Top': ['#7', '#8', '#10', '#11', '#14'],
#     'Bot': ['#7', '#8', '#10', '#11', '#14']
# }

# DB_SPACING = 1.5

# ITERATION_GAP = {
#     'Left': np.array([0.15, 0.45]),
#     'Right': np.array([0.55, 0.85])
# }

# COVER = 0.04
