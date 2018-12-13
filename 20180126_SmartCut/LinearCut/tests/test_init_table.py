import os
import sys
import unittest

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.append(os.path.join(SCRIPT_DIR, os.path.pardir))

from app.init_table import init_beam, init_beam_name
from database.dataset_beam_design import load_beam_design
from tests.const import BEAM_DESIGN

read_file = f'{SCRIPT_DIR}/{BEAM_DESIGN}.xlsx'
save_file = f'{SCRIPT_DIR}/../temp/{BEAM_DESIGN}.xlsx.pkl'


class Test_init_table(unittest.TestCase):
    def setUp(self):
        self.beam_design = load_beam_design(read_file, save_file)

    def test_init_beam(self):
        beam_name = init_beam_name(self.beam_design)

        self.assertEqual(beam_name.at[2, '樓層'], '12F')


if __name__ == '__main__':
    unittest.main()
