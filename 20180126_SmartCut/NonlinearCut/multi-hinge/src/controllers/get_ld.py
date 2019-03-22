"""
It is used for nominal concrete in case of phi_e=1.0 & phi_t=1.0.
Reference:土木401-93
PI = 3.1415926
"""
from math import sqrt


def get_ld(B, num, db, dh, ah, spacing, top, fc, fy, fyh, cover):
    """
    It is used for nominal concrete in case of phi_e=1.0 & phi_t=1.0.
    Reference:土木401-93
    PI = 3.1415926

    prams:
        'B: Girder/Beam Width (m)
        'num: numbers of flexural rebar at lap location
        'db: diameter of flexural rebar (m)
        'dh: stirrup diameter (m)
        'ah: stirrup area (m2)
        'spacing: space of stirrup at lap location of flexural rebar (m)
        'top: is top flexural rebar (boolean)
        'fc: 28-days concrete compressive strength (tonf/cm2)
        'fy: Flexural rebar Nominal Yielding Streng (tonf/cm2)
        'fyh: Stirrup Nominal Yielding Streng (tonf/cm2)
        'cover: Clear Cover (m)

    return:
        ld: general development length (m)

    example:
        #
    """
    # pylint: disable=invalid-name
    PI = 3.1415926

    # change unit m to cm, familiar adress in cm
    # m => cm
    B = B * 100
    fc = fc / 10
    fy = fy / 10
    fyh = fyh / 10
    cover = cover * 100
    db = db * 100
    dh = dh * 100
    ah = ah * 10000
    spacing = spacing * 100

    # 5.2.2
    if sqrt(fc) > 26.5:
        fc = 700

    # R5.3.4.1.1
    cc = dh + cover

    # R5.3.4.1.1
    cs = (B - db * num - dh * 2 - cover * 2) / (num - 1) / 2

    if cc < cs:
        # Vertical splitting failure
        cb = db / 2 + cc
        # R5.3.4.1.2
        ktr = (PI * dh ** 2 / 4) * fyh / 105 / spacing
    else:
        # Horizontal splitting failure
        cb = db / 2 + cs
        # R5.3.4.1.2
        ktr = 2 * (PI * dh ** 2 / 4) * fyh / 105 / spacing / num

    # 5.3.4.1
    ld = 0.28 * fy / sqrt(fc) * db / min((cb + ktr) / db, 2.5)

    # 5.3.4.1
    simple_ld = 0.19 * fy / sqrt(fc) * db

    if simple_ld < ld:
        ld = simple_ld

    # phi_s factor
    if db < 2.2:
        ld = 0.8 * ld

    # phi_t factor
    if top:
        ld = 1.3 * ld

    # 5.3.1
    if ld < 30:
        ld = 30

    return ld
