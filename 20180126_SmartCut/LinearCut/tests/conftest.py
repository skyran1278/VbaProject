import os
import sys
import pytest

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.append(os.path.join(SCRIPT_DIR, os.path.pardir))

from database.dataset_beam_design import load_beam_design
from database.dataset_e2k import load_e2k
from tests.const import BEAM_DESIGN, E2K


# @pytest.fixture(scope='session')
@pytest.fixture(scope='session')
def beam_design():
    beam_design = load_beam_design(
        f'{SCRIPT_DIR}/{BEAM_DESIGN}', f'{SCRIPT_DIR}/../temp/{BEAM_DESIGN}.pkl')

    return beam_design


@pytest.fixture(scope='session')
def e2k():
    e2k = load_e2k(
        f'{SCRIPT_DIR}/{E2K}', f'{SCRIPT_DIR}/../temp/{E2K}.pkl')

    return e2k
