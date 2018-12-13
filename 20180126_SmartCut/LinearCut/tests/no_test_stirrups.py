import os
import sys
import pytest

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.append(os.path.join(SCRIPT_DIR, os.path.pardir))

from app.stirrups import calc_sturrups
from database.dataset_beam_design import load_beam_design
from tests.const import BEAM_DESIGN

read_file = f'{SCRIPT_DIR}/{BEAM_DESIGN}.xlsx'
save_file = f'{SCRIPT_DIR}/../temp/{BEAM_DESIGN}.xlsx.pkl'


def test_(parameter_list):
    pass
