import os
import sys
import unittest

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.append(os.path.join(SCRIPT_DIR, os.path.pardir))

from app.init_table import init_beam, init_beam_name
from database.dataset_beam_design import load_beam_design
from database.dataset_e2k import load_e2k
from tests.const import BEAM_DESIGN, E2K


class Test_init_table(unittest.TestCase):
    def setUp(self):
        self.beam_design = load_beam_design(
            f'{SCRIPT_DIR}/{BEAM_DESIGN}', f'{SCRIPT_DIR}/../temp/{BEAM_DESIGN}.pkl')
        self.e2k = load_e2k(
            f'{SCRIPT_DIR}/{E2K}', f'{SCRIPT_DIR}/../temp/{E2K}.pkl')

    def test_init_beam_name(self):
        beam_name = init_beam_name(self.beam_design)

        self.assertEqual(beam_name.at[2, '樓層'], '12F')

    def test_beam_con(self):
        beam_con = init_beam(
            multi=3, beam_design=self.beam_design, e2k=self.e2k)

        self.assertEqual(beam_con.at[0, ('梁長', '')], 525)
        self.assertEqual(beam_con.at[0, ('主筋', '左')], '')

    def test_beam_cut(self):
        beam_cut = init_beam(
            multi=5, beam_design=self.beam_design, e2k=self.e2k)

        self.assertEqual(beam_cut.at[0, ('主筋', '左1')], '')


if __name__ == '__main__':
    unittest.main()
