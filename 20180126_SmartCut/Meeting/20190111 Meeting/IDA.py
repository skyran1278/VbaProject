import os
import sys
import pickle

import pandas as pd
import numpy as np

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.append(os.path.join(SCRIPT_DIR, os.path.pardir))

filename = '20181229 story drift'

read_file = f'{SCRIPT_DIR}/{filename}'

dataset = pd.read_excel(read_file, sheet_name='Story Drifts')
