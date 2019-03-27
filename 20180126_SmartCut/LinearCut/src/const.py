""" const from user. """
import os

import numpy as np

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

const = {  # pylint: disable=invalid-name
    'etabs_design_path': SCRIPT_DIR + '/tests/20190103 v3.0 3floor for v9.xlsx',
    'e2k_path': SCRIPT_DIR + '/tests/20190103 v3.0 3floor for v9.e2k',

    'beam_name_path': SCRIPT_DIR + '/tests/first_run ida #8 v1.0.xlsx',

    'output_dir': SCRIPT_DIR + '/tests',

    'stirrup_rebar': ['#4', '2#4', '2#5', '2#6'],
    'stirrup_spacing': np.array([10, 12, 15, 18, 20, 22, 25, 30]),

    'rebar': {
        'Top': ['#8', '#10', '#11', '#14'],
        'Bot': ['#8', '#10', '#11', '#14']
    },

    'db_spacing': 1.5,

    'iteration_gap': {
        'left': np.array([0.15, 0.45]),
        'right': np.array([0.55, 0.85])
    },

    'cover': 0.04,
}
