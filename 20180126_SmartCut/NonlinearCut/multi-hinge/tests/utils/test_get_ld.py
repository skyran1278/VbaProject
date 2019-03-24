from math import isclose
from src.utils.get_ld import get_ld


def test_get_ld():
    ld = get_ld(B=0.6, num=9, db=0.0222, dh=0.0127, ah=0.0001267 * 2,
                spacing=0.15, top=True, fc=2800, fy=42000, fyh=42000, cover=0.04)

    assert isclose(ld, 1.01091526616028)

    ld = get_ld(B=0.6, num=9, db=0.0222, dh=0.0127, ah=0.0001267,
                spacing=0.12, top=False, fc=2800, fy=42000, fyh=42000, cover=0.04)

    assert isclose(ld, 0.890165907543924)
