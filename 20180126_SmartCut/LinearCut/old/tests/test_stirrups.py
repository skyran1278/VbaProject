import os
import sys
import pytest

import numpy as np

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.append(os.path.join(SCRIPT_DIR, os.path.pardir))

from tests.const import STIRRUP_REBAR as REBAR, STIRRUP_SPACING as SPACING
from app.init_table import init_beam
from app.stirrups import calc_sturrups, first_calc_dbt_spacing, upgrade_size, merge_segments

# list to numpy
SPACING = np.array(SPACING) / 100


class Test_stirrups(object):
    def test_first_calc_dbt_spacing(self, beam_design, e2k):
        rebars = e2k[0]

        beam_design = first_calc_dbt_spacing(beam_design, rebars, REBAR)

        assert beam_design.at[0, 'VSize'] == '#4'
        assert beam_design.at[0, 'Spacing'] == (
            1.267 * 2 / 10000 / 0.000723999983165413)

    def test_upgrade_size(self, beam_design, e2k):
        rebars = e2k[0]
        beam_design = first_calc_dbt_spacing(beam_design, rebars, REBAR)

        beam_design = upgrade_size(beam_design, rebars, REBAR, SPACING)

        assert beam_design.at[0, 'VSize'] == '#4'
        assert beam_design.at[0, 'Spacing'] == (
            1.267 * 2 / 10000 / 0.000723999983165413)

        # 12F B131
        assert beam_design.at[153, 'VSize'] == '2#4'
        assert beam_design.at[153, 'Spacing'] == (
            1.267 * 4 / 10000 / 0.00298399990424514)

    def test_merge_segments(self, beam_design, e2k):
        rebars = e2k[0]
        beam_con = init_beam(multi=3, beam_design=beam_design, e2k=e2k)

        beam_design = first_calc_dbt_spacing(beam_design, rebars, REBAR)
        beam_design = upgrade_size(beam_design, rebars, REBAR, SPACING)
        beam_con, beam_v = merge_segments(beam_con, beam_design)

        # 10F B170
        assert beam_con.at[12, ('箍筋', '左')] == '2#4@18'
        assert beam_con.at[12, ('箍筋', '中')] == '#4@10'

        assert beam_v.at[154, 'VSize'] == '2#4'
        assert beam_v.at[154, 'Spacing'] == 0.19153439937503527
        assert beam_v.at[154, 'RealVSize'] == '2#4'
        assert beam_v.at[154, 'RealSpacing'] == 0.18

        assert beam_v.at[170, 'VSize'] == '2#4'
        assert beam_v.at[170, 'Spacing'] == 0.21565957757910206
        assert beam_v.at[170, 'RealVSize'] == '#4'
        assert beam_v.at[170, 'RealSpacing'] == 0.1
