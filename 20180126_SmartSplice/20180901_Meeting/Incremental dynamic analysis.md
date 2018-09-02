# Incremental dynamic analysis 2720

## SUMMARY

- 總結了各篇會講甚麼，不重要

1. INTRODUCTION

    - 隨著電腦越來越快，分析的方法就越加複雜。
    - 線性靜力、動力 => 非線性靜力 => 非線性動力
    - SINGLE STATIC ANALYSIS
    - INCREMENTAL SPO
    - SINGLE TIME-HISTORY ANALYSIS
    - INCREMENTAL ONE
    - the range of response or ‘demands’ versus the range of potential levels of a ground motion record
    - the structural implications of rarer/more severe ground motion levels
    - the changes in the nature of the structural response as the intensity of ground motion increases (e.g. changes in peak deformation patterns with height, onset of sti4ness and strength degradation and their patterns and magnitudes)
    - estimates of the dynamic capacity of the global structural system
    - a multi-record IDA study, understanding how stable (or variable) all these items are from one ground motion record to another

2. FUNDAMENTALS OF SINGLE-RECORD IDAs

    - Scale Factor: 放大係數
    - IM
    - DM
    - Single-Record IDA
    - IDA Curve

3. LOOKING AT AN IDA CURVE: SOME GENERAL PROPERTIES

    - yield = 0.2g
    - a b: SOFTENS
    - c d: 等位移
    - a b: infinity
    - c d: SOFTENING HARDENING, IDA 可能會有強度越大損害卻變小的狀況
    - EARLIER YIELDING IN THE STRONGER GROUND MOTION LEADS TO A LOWER ABSOLUTE PEAK RESPONSE
    - FUSE
    - EXTREME CASE, STRUCTURAL RESURRECTION 結構可能會復活

4. CAPACITY AND LIMIT-STATES ON SINGLE IDA CURVES

    - DM-based rule
    - IM-based rule
        - difficulty in prescribing a CIM value
        - FEMA 20% TANGENT SLOPE APPROACH
    - composite

5. MULTI-RECORD IDAS AND THEIR SUMMARY

    - Multi-Record IDA
    - IDA Curve Set
        - PARAMETRIC METHODS
            - median，只 fit 一條線，失去靈活性但是簡單
        - NON-PARAMETRIC METHODS
            - scatterplot smoothers
            - 無限大無法平均 => 中位數
            - MEAN MEDIAN
        - CAPACITY
            - Capacity and Demand Correlation

6. THE IDA IN A PBEE FRAMEWORK

    - 感覺是一個機率統計的框架

7. SCALING LEGITIMACY AND IM SELECTION

    - 探討放大的合法性
        - 真的可以代表未縮放的地震嗎
        - 兩個回歸線非常相近
        - ROOF DUCTILITY RESPONSE OF A T1= 1 S
        - MDOF STEEL FRAME SUBJECTED TO 20 RECORDS
        - SCALED TO 5 LEVELS OF SA(T1;5%).
    - 看 DM IM 都會影響
        - 如何挑選合適的成為一個議題

8. THE IDA VERSUS THE R-FACTOR

    - It can be said that the difference between the R-factor and IDA perspectives is one of design versus assessment.

9. THE IDA VERSUS THE NON-LINEAR SPO

    - 彈性區域很吻合 ELASTIC
    - 非負 => 等位移 NON-NEGATIVE => EQUAL DISPLACEMENT
    - 負 => 軟化 NEGATIVE => SOFTENING
    - 隨後接上非負 => 新的等位移

10. IDA ALGORITHMS

    - HUNTING PHASE
    - FILL PHASE

11. CONCLUSIONS
