import os
import sys
import pytest

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.append(os.path.join(SCRIPT_DIR, os.path.pardir))

from tests.const import STIRRUP_REBAR as REBAR, STIRRUP_SPACING as SPACING
from app.init_table import init_beam
from app.stirrups import calc_sturrups, first_calc_dbt_spacing, upgrade_size, merge_segments


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
        assert False

    def test_merge_segments(self, beam_design, e2k):
        rebars = e2k[0]
        beam_con = init_beam(multi=3, beam_design=beam_design, e2k=e2k)
        beam_design = first_calc_dbt_spacing(beam_design, rebars, REBAR)
        beam_design = upgrade_size(beam_design, rebars, REBAR, SPACING)
        beam_con, beam_v = merge_segments(beam_con, beam_design)
        assert False
