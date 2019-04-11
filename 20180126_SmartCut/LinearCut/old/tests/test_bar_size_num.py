import os
import sys
import pytest

import numpy as np

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.append(os.path.join(SCRIPT_DIR, os.path.pardir))

from tests.const import BAR, DB_SPACING
from app.bar_size_num import calc_db_by_beam, calc_db_by_frame
