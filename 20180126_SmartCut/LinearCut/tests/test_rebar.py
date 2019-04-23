"""
test
"""
from src.rebar import *


def test_get_area():
    """
    test get_area
    """

    assert get_area('2#4') == 2 * 0.0001267
    assert get_area('#4') == 0.0001267


def test_get_diameter():
    """
    test get_diameter
    """
    assert get_diameter('2#4') == 0.0127
    assert get_diameter('#4') == 0.0127


def test_double_area():
    """
    test double_area
    """
    assert double_area('2#4') == 2 * 0.0001267
    assert double_area('#4') == 2 * 0.0001267


def test_rebar_db():
    """
    test rebar_db
    """
    assert rebar_db('2#4') == 0.0127
    assert rebar_db('#4') == 0.0127


def test_rebar_area():
    """
    test rebar_area
    """
    assert rebar_area('2#4') == 0.0001267
    assert rebar_area('#4') == 0.0001267
