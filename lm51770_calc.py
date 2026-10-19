#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LM51770 4-switch Buck-Boost power-train calculator
===================================================

Python 版的 `LM51770_Buck-Boost_Quickstart_Tool_V1_0_1(3).xlsm`。
計算步驟與 xlsm 的 Design Converter (Step 1 ~ Step 8) 相同，公式來源見
`LM51770_Design_Converter_分析.md` / `LM51770_xlsm_全分頁分析.md`。

Duty cycle 有兩種模型 (`duty_model`)：

* ``"lm5177"``（預設）：依 LM5177 的四開關 buck-boost 時序計算。
    - 先用 xlsm `Timings` 分頁的 LM5177 時序參數算出 Buck↔Buck-Boost↔Boost
      的轉換電壓（VIN_buck_to_BB、VIN_BB_to_buck、VIN_boost_to_BB、VIN_BB_to_boost）。
    - 依 VIN 落在哪一區決定模式：BUCK / BOOST / BUCK-BOOST。
    - BUCK、BOOST：含 RDS(on)、DCR、RS 壓降的 duty
      （= xlsm `Efficiency vs VIN Calculations` 的 t1 欄）。
    - BUCK-BOOST：三相位 (t1: Q1+Q3 直通、t2: Q1+Q4 升壓充磁、t3: Q2+Q3 降壓放磁)，
      t1 = T − (N4 + 2·N5) 固定，t2 由伏秒平衡求得，t3 = T − t1 − t2。
    - 電感電流以分段線性波形計算峰值、漣波與 RMS，電容 / MOSFET RMS 亦由波形積分。
* ``"xlsm"``：完全照 xlsm `Variable Management` 的公式
  （只以 VIN > VOUT 分 buck/boost，duty 含 DCR 與 RS 壓降），用來和 xlsm 數值比對。

Bode 圖：電壓迴路（Stability Calculations，峰值電流模式 Ridley 型模型）與
平均電流限制迴路（IL Stability Calculations，複數阻抗模型）。

用法::

    python3 lm51770_calc.py                 # 以 xlsm 存檔值計算並列出報表
    python3 lm51770_calc.py --model xlsm    # 使用 xlsm 原始 duty 算法
    python3 lm51770_calc.py --json out.json # 輸出全部結果 (含 Bode / 效率曲線)
    python3 lm51770_calc.py --plot out_dir  # 需要 matplotlib：輸出 Bode / 效率圖
    python3 lm51770_calc.py --set vin_nom=20 --set iout=6

只依賴 Python 標準函式庫（numpy / matplotlib 為選用）。
"""

from __future__ import annotations

import argparse
import cmath
import dataclasses
import json
import math
from dataclasses import dataclass, field
from typing import Dict, List, Optional

PI = math.pi

# ---------------------------------------------------------------------------
# Constants copied from the xlsm (Variable Management / DeviceParameter)
# ---------------------------------------------------------------------------

DEVICES = {
    # gm_imon, gm_il, gm_csa, CS threshold min/typ/max
    "LM51770": dict(gm_imon=1e-3, gm_il=200e-6, gm_csa=600e-6,
                    vcs_min=0.0425, vcs_typ=0.050, vcs_max=0.0575),
    "LM517701": dict(gm_imon=1e-3, gm_il=200e-6, gm_csa=600e-6,
                     vcs_min=0.065625, vcs_typ=0.075, vcs_max=0.084375),
}

# LM5177 / LM51772 timing constants (Timings!G1:H5)
TIMING_DEVICES = {
    "LM5177": dict(I1=115e-9, I2=2, I3=1, I4=2, I5=1),
    "LM51772": dict(I1=80e-9, I2=4, I3=1, I4=2, I5=1),
}

# DeviceParameter!P7:Q107 : boost tON(min) [ns] vs fSW [MHz]
TON_MIN_TABLE = [
    (0.1, 435.3921), (0.1212, 404.8119), (0.1424, 379.0609), (0.1636, 355.5805),
    (0.1848, 334.3524), (0.2061, 315.6626), (0.2273, 299.4314), (0.2485, 283.9873),
    (0.2697, 270.1134), (0.2909, 257.7974), (0.3121, 246.5349), (0.3333, 236.0069),
    (0.3545, 226.2988), (0.3758, 217.5111), (0.397, 209.2128), (0.4182, 201.4597),
    (0.4394, 193.9455), (0.4606, 187.2546), (0.4818, 180.9858), (0.503, 175.1552),
    (0.5242, 169.5948), (0.5455, 164.3228), (0.5667, 159.3714), (0.5879, 154.7458),
    (0.6, 151.8441), (0.6091, 150.1909), (0.6303, 146.0382), (0.6515, 141.9745),
    (0.6727, 138.1506), (0.6939, 134.4471), (0.7152, 130.9936), (0.7364, 127.8038),
    (0.7576, 124.6821), (0.7788, 121.6017), (0.8, 118.6832), (0.8212, 115.9532),
    (0.8424, 113.2763), (0.8636, 110.4466), (0.8848, 108.4077), (0.9061, 105.8044),
    (0.9273, 103.5754), (0.9485, 101.3655), (0.9697, 99.1579), (0.9909, 97.0561),
    (1.0121, 95.115), (1.0333, 93.2483), (1.0545, 91.3776), (1.0758, 91.2921),
    (1.097, 91.2118), (1.1182, 91.1045), (1.1394, 91.0418), (1.1606, 90.7792),
    (1.1818, 90.677), (1.203, 90.6271), (1.2242, 90.4862), (1.2455, 90.2243),
    (1.2667, 90.2613), (1.2879, 90.1815), (1.3091, 89.9721), (1.3303, 89.7389),
    (1.3515, 89.6257), (1.3727, 89.5104), (1.3939, 89.3317), (1.4152, 89.3375),
    (1.4364, 89.2183), (1.4576, 89.145), (1.4788, 88.8905), (1.5, 88.917),
    (1.5212, 88.7541), (1.5424, 88.5991), (1.5636, 88.5257), (1.5848, 88.7184),
    (1.6061, 88.3713), (1.6273, 88.4202), (1.6485, 88.2529), (1.6697, 88.2802),
    (1.6909, 88.052), (1.7121, 88.0665), (1.7333, 88.0375), (1.7545, 87.7745),
    (1.7758, 87.7724), (1.797, 87.6184), (1.8182, 87.6316), (1.8394, 87.3767),
    (1.8606, 87.3151), (1.8818, 87.4674), (1.903, 87.2847), (1.9242, 87.1411),
    (1.9455, 87.1667), (1.9667, 86.794), (1.9879, 86.9755), (2.0091, 86.9146),
    (2.0303, 86.9184), (2.0515, 86.6816), (2.0727, 86.4207), (2.0939, 86.6138),
    (2.1152, 86.5314), (2.1364, 86.2811), (2.1576, 86.297), (2.1788, 86.4205),
    (2.2, 86.2333),
]

# CFG Setting!B4:B19 : R_CFG [kΩ] for CFG #1..#16
RCFG_TABLE = [0, 0.511, 1.15, 1.9, 2.7, 3.8, 5.1, 6.5, 8.3, 10.5, 13.3, 16.2,
              20.5, 24.9, 30.1, 36.5]

# Stability Calculations!B13:B613 : xlsm Bode 掃頻點 (Hz)
XLSM_FREQS = [
    1, 1.0965, 1.2023, 1.3183, 1.4454, 1.5849, 1.7378, 1.9055, 2.0893, 2.2909, 2.5119,
    2.7542, 3.02, 3.3113, 3.6308, 3.9811, 4.3652, 4.7863, 5.2481, 5.7544, 6.3096, 6.9183,
    7.5858, 8.3176, 9.1201, 10, 10.965, 12.023, 13.183, 14.454, 15.849, 17.378, 19.055,
    20.893, 22.909000000000002, 25.119, 27.542, 30.2, 33.113, 36.308, 39.811, 43.652,
    47.863, 52.481, 57.544000000000004, 63.096, 69.183, 75.858, 83.176, 91.20100000000001,
    100, 109.65, 120.23, 131.83, 144.54, 158.49, 173.78, 190.55, 208.93, 229.09, 251.19,
    275.42, 302, 331.13, 363.08, 398.11, 436.52, 478.63, 524.81, 575.44, 630.96, 691.83,
    758.58, 831.76, 912.01, 1000, 1096.5, 1202.3, 1318.3, 1445.4, 1584.9, 1737.8, 1905.5,
    2089.3, 2290.9, 2511.9, 2754.2, 3020, 3311.3, 3630.8, 3981.1, 4365.2, 4786.3, 5248.1,
    5370.3, 5432.85, 5495.4, 5559.4, 5623.4, 5688.9, 5754.4, 5821.4, 5888.4, 5957, 6025.6,
    6095.8, 6166, 6237.8, 6309.6, 6383.05, 6456.5, 6531.7, 6606.9, 6683.85, 6760.8,
    6839.55, 6918.3, 6998.9, 7079.5, 7161.95, 7244.4, 7328.75, 7413.1, 7499.45, 7585.8,
    7674.15, 7762.5, 7852.9, 7943.3, 8035.8, 8128.3, 8222.95, 8270.275000000001, 8317.6,
    8366.05, 8414.5, 8462.95, 8511.4, 8560.95, 8610.5, 8660.05, 8709.6, 8760.325, 8811.05,
    8861.775, 8912.5, 8964.4, 9016.3, 9068.2, 9120.1, 9173.2, 9226.3, 9279.4, 9332.5,
    9386.85, 9441.2, 9495.55, 9549.9, 9605.525, 9661.15, 9716.775, 9772.4, 9829.3, 9886.2,
    9943.1, 10000, 10058.25, 10116.5, 10174.75, 10233, 10292.5, 10352, 10411.5, 10471,
    10532, 10593, 10654, 10715, 10777.5, 10840, 10902.5, 10965, 11028.75, 11092.5,
    11156.25, 11220, 11285.5, 11351, 11416.5, 11482, 11548.75, 11615.5, 11682.25, 11749,
    11817.5, 11886, 11954.5, 12023, 12093, 12163, 12233, 12303, 12374.5, 12446, 12517.5,
    12589, 12662.25, 12735.5, 12808.75, 12882, 12957.25, 13032.5, 13107.75, 13183,
    13259.75, 13336.5, 13413.25, 13490, 13568.5, 13647, 13725.5, 13804, 13884.25, 13964.5,
    14044.75, 14125, 14207.25, 14289.5, 14371.75, 14454, 14622.5, 14706.75, 14791,
    14877.25, 14963.5, 15049.75, 15136, 15224, 15312, 15400, 15488, 15578.25, 15668.5,
    15758.75, 15849, 15941.25, 16033.5, 16125.75, 16218, 16312.5, 16407, 16501.5, 16596,
    16692.5, 16789, 16885.5, 16982, 17081, 17180, 17279, 17378, 17479.25, 17580.5,
    17681.75, 17783, 17886.5, 17990, 18093.5, 18197, 18303, 18409, 18515, 18621, 18729.5,
    18838, 18946.5, 19055, 19165.75, 19276.5, 19387.25, 19498, 19611.75, 19725.5, 19839.25,
    19953, 20069, 20185, 20301, 20417, 20536, 20655, 20774, 20893, 21014.75, 21136.5,
    21258.25, 21380, 21504.5, 21629, 21753.5, 21878, 22005.25, 22132.5, 22259.75, 22387,
    22517.5, 22648, 22778.5, 22909, 23042.25, 23175.5, 23308.75, 23442, 23578.5, 23715,
    23851.5, 23988, 24127.75, 24267.5, 24407.25, 24547, 24690, 24833, 24976, 25119,
    25265.25, 25411.5, 25557.75, 25704, 25853.75, 26003.5, 26153.25, 26303, 26456, 26609,
    26762, 26915, 27071.75, 27228.5, 27385.25, 27542, 27702.5, 27863, 28023.5, 28184,
    28348, 28512, 28676, 28840, 29008, 29176, 29344, 29512, 29684, 29856, 30028, 30200,
    30375.75, 30551.5, 30727.25, 30903, 31083, 31263, 31443, 31623, 31807, 31991, 32175,
    32359, 32547.5, 32736, 32924.5, 33113, 33305.75, 33498.5, 33691.25, 33884, 34081.5,
    34279, 34476.5, 34674, 34875.75, 35077.5, 35279.25, 35481, 35687.75, 35894.5, 36101.25,
    36308, 36519.5, 36731, 36942.5, 37154, 37370.25, 37586.5, 37802.75, 38019, 38240.5,
    38462, 38683.5, 38905, 39131.5, 39358, 39584.5, 39811, 40042.75, 40274.5, 40506.25,
    40738, 40975.25, 41212.5, 41449.75, 41687, 41929.75, 42172.5, 42415.25, 42658, 42906.5,
    43155, 43403.5, 43652, 43906, 44160, 44414, 44668, 44928.25, 45188.5, 45448.75, 45709,
    45975.25, 46241.5, 46507.75, 46774, 47046.25, 47318.5, 47590.75, 47863, 48141.75,
    48420.5, 48699.25, 48978, 49263.25, 49548.5, 49833.75, 50119, 50410.75, 50702.5,
    50994.25, 51286, 51584.75, 51883.5, 52182.25, 52481, 52786.5, 53092, 53397.5, 53703,
    54015.75, 54328.5, 54641.25, 54954, 55274, 55594, 55914, 56234, 56561.5, 56889,
    57216.5, 57544, 57879, 58214, 58549, 58884, 59227, 59570, 59913, 60256, 60607, 60958,
    61309, 61660, 62019, 62378, 62737, 63096, 63463.25, 63830.5, 64197.75, 64565, 64941,
    65317, 65693, 66069, 66453.75, 66838.5, 67223.25, 67608, 68001.75, 68395.5, 68789.25,
    69183, 69586, 69989, 70392, 70795, 71207.25, 71619.5, 72031.75, 72444, 72865.75,
    73287.5, 73709.25, 74131, 74562.75, 74994.5, 75426.25, 75858, 76299.75, 76741.5,
    77183.25, 77625, 78077, 78529, 78981, 79433, 79895.5, 80358, 80820.5, 81283, 81756.25,
    82229.5, 82702.75, 83176, 83660.5, 84145, 84629.5, 85114, 85609.5, 86105, 86600.5,
    87096, 87603.25, 88110.5, 88617.75, 89125, 89644, 90163, 90682, 91201, 91732, 92263,
    92794, 93325, 93868.5, 94412, 94955.5, 95499, 96611.5, 97724, 98862, 100000, 101165,
    102330, 103520, 104710, 105930, 107150, 108400, 109650, 110925, 112200, 113510, 114820,
    116155, 117490, 118860, 120230, 126030, 131830, 144540, 158490, 173780, 190550, 208930,
    229090, 251190, 275420, 302000, 331130, 363080, 398110, 436520, 478630, 524810, 575440,
    630960, 691830, 758580, 831760, 912010, 1000000,
]

E12_BASE = [1.0, 1.2, 1.5, 1.8, 2.2, 2.7, 3.3, 3.9, 4.7, 5.6, 6.8, 8.2]


# ---------------------------------------------------------------------------
# Inputs (defaults = values stored in the xlsm)
# ---------------------------------------------------------------------------

@dataclass
class Mosfet:
    rdson: float = 3.5e-3     # Ω
    qg: float = 67e-9         # C
    qgd: float = 10.31e-9     # C
    qgs: float = 11.25e-9     # C
    rg: float = 0.63          # Ω
    gfs: float = 117.0        # S
    vth: float = 2.0          # V
    vsd: float = 0.75         # V (body diode)
    qrr: float = 0.0          # C
    theta_ja: float = 65.0    # °C/W


@dataclass
class Design:
    device: str = "LM51770"
    # Step 1
    vin_min: float = 6.0
    vin_nom: float = 12.0
    vin_max: float = 36.0
    vout: float = 16.0
    iout: float = 8.0
    fsw: float = 398e3
    efficiency: float = 0.95          # 估計效率 (VM B16)
    mode_fpwm: bool = True            # VM B30 Mode: True = FPWM, False = PSM
    # Step 2
    lf: float = 1.8e-6
    rdcr: float = 3.2e-3
    # Step 3
    ilimit_margin: float = 20.0       # %
    rs: float = 1.3e-3
    rslope: Optional[float] = None    # None = 使用建議值 (E96)
    rs_calc_mode: int = 1             # VM L14 Calc_mode: 1 = "Joerg", 0 = datasheet 式 23
    # Step 4
    vout_ripple: float = 0.150
    cout: float = 130e-6
    cout_esr: float = 2e-3
    # Step 5
    vin_ripple: float = 0.180
    cin: float = 100e-6
    cin_esr: float = 2e-3
    # Step 6
    tss: float = 1.8e-3
    vin_uvlo_on: float = 4.6
    vin_uvlo_off: float = 4.025
    cfg_drss: int = 1                 # 1 = off, 2 = on
    cfg_hiccup: int = 2               # 1 = off, 2 = on
    cfg_psm: int = 1                  # 1 = 10 %, 2 = 15 %
    cfg_imon_lim: int = 3             # 1 = off, 2 = IMON, 3 = ILIM +, 4 = ILIM −
    current_setpoint: float = 5.0     # A (IMON 滿刻度 / 平均電流限制)
    # Step 7
    fcross: float = 4e3
    rfbt: float = 71.5e3
    rcomp: Optional[float] = None     # None = 使用 E96/E12 標準值
    ccomp: Optional[float] = None
    chf: Optional[float] = None
    # Step 7.2 (選用的 IMON 補償，xlsm DC!M91/M92)
    cimon: float = 2.2e-9
    rimon: float = 700.0
    # Step 8
    q1: Mosfet = field(default_factory=Mosfet)
    q2: Mosfet = field(default_factory=lambda: Mosfet(qrr=45e-9))
    q3: Mosfet = field(default_factory=lambda: Mosfet(qrr=56e-9))
    q4: Mosfet = field(default_factory=Mosfet)
    vf_schottky_buck: float = 0.0     # 外接蕭特基 (Q2 並聯)
    qrr_schottky_buck: float = 0.0
    vf_schottky_boost: float = 0.0    # 外接蕭特基 (Q3 並聯)
    qrr_schottky_boost: float = 0.0
    # LM5177 duty 模型選項
    timing_device: str = "LM5177"     # Timings!B1 (1 = LM5177)
    vin_direction: str = "rising"     # 轉換區遲滯：rising (= xlsm 效率頁) / falling


# Fixed model constants (VM)
CONST = dict(
    vref=1.0, ifb=60e-9, iss=10e-6,
    vuvlo_on=1.28, vuvlo_off=1.17, iuvlo=6e-6,
    iout_threshold=0.050,           # ΔV(ISNS) 50 mV
    csamp_gain=10.0, gm_ea=600e-6,
    cslope=2.5e-12, gm_slopecomp=20e-9 * 125000 * 0.75, isc_const=0.0,
    gbw=13e6, roea=10e6, gbw_il=2e6, roea_il=20e6, cc2il=1e-15,
    r_cs=10.0, r_v2i_ipeak=25000.0,
    ta=25.0, ltc=0.004, theta_ca=20.0, vdd=7.5,
    rdrv_up=1.5, rdrv_dn=1.0, deadtime=50e-9, iq=3e-3,
    rdson_tc_top=0.004, rdson_tc_bot=0.004,
)


# ---------------------------------------------------------------------------
# Standard values (Standard Value Calculator)
# ---------------------------------------------------------------------------

def e96(x: float) -> float:
    """E96 rounding, same formula as SVC column O."""
    if x is None or x <= 0:
        return x
    lg = math.log10(x)
    dec = math.floor(lg)
    frac = lg - dec
    n = round(96 * frac)
    if 96 * frac - n < 0:
        n -= 1
    scale = 10 ** (dec - 2)
    lo = int(0.5 + 100 * 10 ** (n / 96)) * scale
    hi = int(0.5 + 100 * 10 ** ((n + 1) / 96)) * scale
    return hi if x > (lo + hi) / 2 else lo


def e12(x: float) -> float:
    """Nearest E12 value (linear distance), same as SVC columns B~I."""
    if x is None or x <= 0:
        return x
    dec = math.floor(math.log10(x))
    cands = [b * 10 ** d for d in (dec - 1, dec, dec + 1) for b in E12_BASE]
    return min(cands, key=lambda c: abs(c - x))


def ton_min_boost(fsw: float) -> float:
    """VLOOKUP(Fsw/1e6, P7:Q107, 2) (approximate match) → seconds."""
    f = fsw / 1e6
    val = TON_MIN_TABLE[0][1]
    for fm, t in TON_MIN_TABLE:
        if fm <= f:
            val = t
        else:
            break
    return val * 1e-9


# ---------------------------------------------------------------------------
# LM5177 duty / timing model (Timings sheet)
# ---------------------------------------------------------------------------

def lm5177_transitions(fsw: float, vout: float, timing_device: str = "LM5177",
                       il: float = 0.0) -> Dict[str, float]:
    """Buck↔Buck-Boost↔Boost 轉換電壓與 BB 模式固定時間 (Timings!F2:F5, T2:T6)."""
    p = TIMING_DEVICES[timing_device]
    T = 1.0 / fsw
    J1 = J2 = 0.0
    J3 = J4 = 1e-6
    J5 = 1 if il > 0 else 0
    J6 = 1 if il < 0 else 0
    J7 = J5 - J6
    L1 = p["I1"]
    L3 = 0.225 - p["I2"] * 0.025
    L4 = 0.75 + p["I3"] * 0.25
    L5 = 1 - p["I4"] * 0.0416666
    L6 = 0.75 + p["I5"] * 0.25
    B3 = fsw / 1000.0                       # kHz
    N1 = 0.75 / L1 * J3 / 2e6 * 0.2
    N2 = min(J3 / 2e6 * B3 * 1000 * L5 + J4 * (1 - L5), N1)
    N4 = L3 / B3 * 0.001                    # = L3 · T
    N5 = L4 * 75e-9 / N2 * 1e-6             # minimum on/off pulse
    N6 = N4 + 2 * N5
    P3 = N5 + J7 * J1
    P4 = T - P3                             # max buck on-time
    R2 = N5 - J7 * J1                       # min boost on-time
    T2 = max(max(0.5 * N5 / L4 * L6, J6 * J2) - J7 * J1, 0.0)
    T3 = max(0.5 * N5 / L4 * L6, J5 * J2)
    T5 = N4 + 2 * N5
    T4 = max(T5 - T3 - J7 * J1, 0.0)
    T6 = T - T5                             # BB 模式 Phase t1 (Q1+Q3) 固定時間
    return dict(
        vin_buck_to_bb=vout / (P4 * fsw),               # P9
        vin_bb_to_buck=vout * ((T - T2) / (T6 + T2)),   # T9
        vin_boost_to_bb=vout * (1 - R2 * fsw),          # R8
        vin_bb_to_boost=vout * ((T - T4) / (T6 + T4)),  # T8
        t_min_pulse=N5, t_bb_window=T5, t1_bb=T6, N4=N4, N5=N5, N6=N6,
        T2=T2, T4=T4,
    )


def _mode_lm5177(vin: float, vout: float, tr: Dict[str, float], fpwm: bool,
                 direction: str) -> str:
    if not fpwm:
        return "BUCK" if vin > vout else "BOOST"
    if direction == "falling":
        # 從高壓往下：buck 直到 buck→BB，BB 直到 BB→boost
        if vin > tr["vin_buck_to_bb"]:
            return "BUCK"
        if vin <= tr["vin_bb_to_boost"]:
            return "BOOST"
        return "BUCK-BOOST"
    # rising (與 xlsm 效率頁相同)：boost 直到 boost→BB，BB 直到 BB→buck
    if vin > tr["vin_bb_to_buck"]:
        return "BUCK"
    if vin <= tr["vin_boost_to_bb"]:
        return "BOOST"
    return "BUCK-BOOST"


def _seg_rms2(dt: float, ia: float, ib: float) -> float:
    """∫ i² dt over a linear segment."""
    return dt * (ia * ib + (ia - ib) ** 2 / 3.0)


def _seg_int(dt: float, ia: float, ib: float) -> float:
    return dt * (ia + ib) / 2.0


def operating_point(d: Design, vin: float, iout: Optional[float] = None,
                    model: str = "lm5177") -> Dict:
    """單一 VIN 的功率級工作點：模式、duty、三相時間、電感電流波形與各元件 RMS 電流。"""
    iout = d.iout if iout is None else iout
    vout, fsw, L = d.vout, d.fsw, d.lf
    T = 1.0 / fsw
    eff = d.efficiency
    iin = vout * iout / (vin * eff)
    tr = lm5177_transitions(fsw, vout, d.timing_device)

    if model == "xlsm":
        # VM B81:B83  Don_Vinxxx_2
        if vin > vout:
            mode = "BUCK"
            D = (vout + iout * (d.rdcr + d.rs)) / (vin + iout * d.rs)
        else:
            mode = "BOOST"
            D = (vout - vin + iin * d.rdcr) / (vout - iin * d.rs)
    else:
        mode = _mode_lm5177(vin, vout, tr, d.mode_fpwm, d.vin_direction)
        r1, r2, r3, r4 = d.q1.rdson, d.q2.rdson, d.q3.rdson, d.q4.rdson
        if mode == "BUCK":   # Efficiency vs VIN!M (含 RDS(on))
            D = (vout + (r2 + d.rs + d.rdcr) * iout) / (vin + (d.rs + r2 - r1) * iout)
        elif mode == "BOOST":
            D = (vout + (r3 + d.rdcr) * iin - vin) / (vout + (r3 - r4 - d.rs) * iin)
        else:
            D = None
    if D is not None:
        D = min(max(D, 0.0), 1.0)

    # segments: (dt, slope [A/s], switches on)
    if mode == "BUCK":
        t1 = D * T
        t2 = T - t1
        t3 = 0.0
        a1, a2 = (vin - vout) / L, -vout / L
        i0 = iout - 0.5 * a1 * t1
        i1 = i0 + a1 * t1
        segs = [(t1, i0, i1, ("Q1", "Q3")), (t2, i1, i0, ("Q2", "Q3"))]
        d_buck, d_boost = D, 0.0
    elif mode == "BOOST":
        t1 = D * T
        t2 = T - t1
        t3 = 0.0
        a1 = vin / L
        i0 = iin - 0.5 * a1 * t1
        i1 = i0 + a1 * t1
        segs = [(t1, i0, i1, ("Q1", "Q4")), (t2, i1, i0, ("Q1", "Q3"))]
        d_buck, d_boost = 1.0, D
    else:  # BUCK-BOOST (LM5177 三相位)
        t1 = min(tr["t1_bb"], T)                                # Q1+Q3 : (Vin−Vout)/L
        t2 = max((vout * T - vin * t1) / (vin + vout), 0.0)     # Q1+Q4 : Vin/L
        t3 = max(T - t1 - t2, 0.0)                              # Q2+Q3 : −Vout/L
        a1, a2 = (vin - vout) / L, vin / L
        # 輸入電荷 Iin·T = ∫(t1+t2) iL dt  (Efficiency vs VIN!P)
        i0 = (iin * T - (0.5 * a1 * t1 ** 2 + a1 * t1 * t2 + 0.5 * a2 * t2 ** 2)) / (t1 + t2)
        i1 = i0 + a1 * t1
        i2 = i1 + a2 * t2
        segs = [(t1, i0, i1, ("Q1", "Q3")), (t2, i1, i2, ("Q1", "Q4")),
                (t3, i2, i0, ("Q2", "Q3"))]
        d_buck, d_boost = (t1 + t2) / T, t2 / T
        D = d_boost

    pts = [s[1] for s in segs] + [s[2] for s in segs]
    il_max, il_min = max(pts), min(pts)
    il_avg = sum(_seg_int(*s[:3]) for s in segs) / T
    il_rms = math.sqrt(sum(_seg_rms2(*s[:3]) for s in segs) / T)

    def rms_of(sw):
        return math.sqrt(sum(_seg_rms2(*s[:3]) for s in segs if sw in s[3]) / T)

    def avg_of(sw):
        return sum(_seg_int(*s[:3]) for s in segs if sw in s[3]) / T

    irms = {q: rms_of(q) for q in ("Q1", "Q2", "Q3", "Q4")}
    i_in_avg, i_out_avg = avg_of("Q1"), avg_of("Q3")
    icin_rms = math.sqrt(max(irms["Q1"] ** 2 - i_in_avg ** 2, 0.0))
    icout_rms = math.sqrt(max(irms["Q3"] ** 2 - i_out_avg ** 2, 0.0))
    irs_rms = math.sqrt(irms["Q2"] ** 2 + irms["Q4"] ** 2)   # RS 在低側 (Q2/Q4 共源)

    return dict(
        vin=vin, iout=iout, iin=iin, mode=mode, model=model, duty=D,
        d_buck_leg=d_buck, d_boost_leg=d_boost, t1=t1, t2=t2, t3=t3,
        segments=[dict(dt=s[0], i_start=s[1], i_end=s[2], on=list(s[3])) for s in segs],
        ripple=il_max - il_min, il_peak=il_max, il_valley=il_min,
        il_avg=il_avg, il_rms=il_rms, irms=irms, irs_rms=irs_rms,
        i_in_avg=i_in_avg, i_out_avg=i_out_avg,
        icin_rms=icin_rms, icout_rms=icout_rms,
    )


# ---------------------------------------------------------------------------
# Bode plots
# ---------------------------------------------------------------------------

def freq_grid(n_per_decade: int = 200, fmin: float = 1.0, fmax: float = 1e6) -> List[float]:
    n = int(round(math.log10(fmax / fmin) * n_per_decade))
    return [fmin * 10 ** (k / n_per_decade) for k in range(n + 1)]


def _crossover(freqs, gain_db, phase_deg, interpolate=True):
    """第一個由正轉負的 0 dB 交越點；回傳 (fc, phase_at_fc).

    interpolate=False 時與 xlsm 相同：取增益仍 > 0 的最後一個掃頻點。
    """
    for k in range(len(freqs) - 1):
        if gain_db[k] > 0 >= gain_db[k + 1]:
            if not interpolate:
                return freqs[k], phase_deg[k]
            g0, g1 = gain_db[k], gain_db[k + 1]
            x = g0 / (g0 - g1)
            lf = math.log10(freqs[k]) + x * (math.log10(freqs[k + 1]) - math.log10(freqs[k]))
            ph = phase_deg[k] + x * (phase_deg[k + 1] - phase_deg[k])
            return 10 ** lf, ph
    return None, None


def voltage_loop_bode(p: Dict, freqs: List[float], interpolate: bool = True) -> Dict:
    """Stability Calculations：漸近 (實數) 模型，相位以各項 atan 相加。"""
    out = dict(f=freqs, gain=[], phase=[], ea_gain=[], ea_phase=[],
               plant_gain=[], plant_phase=[])
    fn = p["f_sample"]
    pterm1 = p["Q"] * fn
    pterm2 = fn ** 2 + 0.01
    for f in freqs:
        C, Dp = math.hypot(1, f / p["f_load"]), math.atan(f / p["f_load"])
        E, F = math.hypot(1, f / p["f_esr"]), math.atan(f / p["f_esr"])
        if p["rhp"]:
            G, H = math.hypot(1, f / p["f_rhp"]), -math.atan(f / p["f_rhp"])
        else:
            G, H = 1.0, 0.0
        I, J = math.hypot(1, f / p["fp_ea"]), math.atan(f / p["fp_ea"])
        K, Lz = math.hypot(1, f / p["fz_comp"]), math.atan(f / p["fz_comp"])
        M, N = math.hypot(1, f / p["fp_comp"]), math.atan(f / p["fp_comp"])
        re = 1 - f ** 2 / pterm2
        im = f / pterm1
        U = math.hypot(re, im)
        V = math.atan2(im, re)          # 0 ~ π (與 xlsm 象限修正相同)
        gain = 20 * math.log10(E * K * G / (C * I * U * M) * p["adc"])
        ph = ((F + Lz + H) - (Dp + J + V + N)) * 180 / PI
        out["gain"].append(gain)
        out["phase"].append(ph)
        out["ea_gain"].append(20 * math.log10(K / (I * M) * p["aea"]))
        out["ea_phase"].append((Lz - (J + N)) * 180 / PI)
        out["plant_gain"].append(20 * math.log10(E * G / (C * U) * p["acs"] * p["am"]))
        out["plant_phase"].append((F + H - (Dp + V)) * 180 / PI)
    fc, ph = _crossover(freqs, out["gain"], out["phase"], interpolate)
    out["fc"] = fc
    out["pm"] = None if ph is None else 180.0 + ph
    fce, phe = _crossover(freqs, out["ea_gain"], out["ea_phase"], interpolate)
    out["fc_ea"] = fce
    return out


def current_limit_bode(d: Design, p: Dict, freqs: List[float],
                       interpolate: bool = True) -> Dict:
    """IL Stability Calculations：平均電流限制迴路 (複數模型，直接移植 xlsm 公式)."""
    vin, vout, fsw, L = p["vin"], d.vout, d.fsw, d.lf
    iload = d.current_setpoint
    rload = d.vout / d.iout
    co2 = d.cout / 2
    co1 = d.cout - co2
    esr = d.cout_esr
    duty = p["duty_il"]
    rsns = p["rshunt"]
    gm_csa = DEVICES[d.device]["gm_csa"]
    i_sns_gain = d.rs * gm_csa * CONST["r_v2i_ipeak"]
    if vin >= vout:
        vse = max(vout, vin / 2)
        sn = i_sns_gain * (vin - vout) / L
        se = i_sns_gain * vse / L
    else:
        vse = max(vout - vin, 5)
        sn = i_sns_gain * vin / L
        se = i_sns_gain * (vout - vin) / L
    fm = fsw / (sn + se)
    ail = DEVICES[d.device]["gm_il"] * CONST["roea_il"]
    il_z1 = 1 / (2 * PI * d.rimon * d.cimon)
    il_p1 = 1 / (2 * PI * (CONST["roea_il"] + d.rimon) * d.cimon)
    cbw_il = DEVICES[d.device]["gm_il"] / (2 * PI * CONST["gbw_il"])
    il_p2 = 1 / (2 * PI * d.rimon * (CONST["cc2il"] + cbw_il))

    out = dict(f=freqs, gain=[], phase=[], amp_gain=[], amp_phase=[],
               plant_gain=[], plant_phase=[])
    for f in freqs:
        w = 2 * PI * f
        zl = 1j * w * L
        zco1 = esr + 1 / (1j * w * co1)
        zco2 = esr + 1 / (1j * w * co2)
        zout = 1 / (1 / rload + 1 / zco2)
        zout_sns = zout + rsns
        zall = 1 / (1 / zco1 + 1 / (rsns + zout))
        he = 1 + 1j * w / (fsw * PI * (-2) / PI) + (1j * w / (fsw * PI)) ** 2
        if vin >= vout:
            g_d_il = vin / (zl + zall)
            g_x = g_d_il
        else:
            g_d_il = (vout + zall * iload) / (zl + zall * (1 - duty) ** 2)
            g_x = g_d_il * (1 - duty) - iload / (1 - duty)
        g_d_isns = g_x / (zout_sns * (1 / zco1 + 1 / zout_sns))
        ti = g_d_il * i_sns_gain * he * fm
        g_vc_isns = fm * g_d_isns / (1 + ti)
        amp = ail * (1 + 1j * f / il_z1) / ((1 + 1j * f / il_p1) * (1 + 1j * f / il_p2))
        loop = g_vc_isns * amp * rsns
        out["gain"].append(20 * math.log10(abs(loop)))
        out["phase"].append(math.degrees(cmath.phase(loop)))
        out["amp_gain"].append(20 * math.log10(abs(rsns * amp)))
        out["amp_phase"].append(math.degrees(cmath.phase(rsns * amp)))
        out["plant_gain"].append(20 * math.log10(abs(g_vc_isns)))
        out["plant_phase"].append(math.degrees(cmath.phase(g_vc_isns)))
    fc, ph = _crossover(freqs, out["gain"], out["phase"], interpolate)
    out["fc"] = fc
    out["pm"] = None if ph is None else (ph - 180 if ph > 0 else ph + 180)
    out.update(il_zero=il_z1, il_pole1=il_p1, il_pole2=il_p2, fm=fm,
               i_sns_gain=i_sns_gain, sn=sn, se=se)
    return out


def plant_gain_1hz_il(d: Design, p: Dict) -> float:
    return current_limit_bode(d, p, [1.0])["plant_gain"][0]


# ---------------------------------------------------------------------------
# Efficiency (Efficiency vs VIN / vs Load Calculations)
# ---------------------------------------------------------------------------

def _self_heat(p_cond: float, tc: float, theta: float, extra: float = 0.0) -> float:
    ta = CONST["ta"]
    return (extra + p_cond * (1 + tc * (ta - 25))) / (1 - p_cond * tc * theta)


def losses(d: Design, vin: float, iout: Optional[float] = None,
           xlsm_compat: bool = False) -> Dict:
    """單點損耗分析 (永遠使用 LM5177 duty 模型，與 xlsm 效率頁一致).

    xlsm_compat=True 時連同 xlsm 效率頁的近似 / 錯誤一起重現 (驗證用)：
    電容 RMS 用封閉式 (BB 模式 duty 欄缺括號)、BB 模式反向恢復損耗重複計算、
    IMON/ILIM 關閉時仍計入 RSHUNT 損耗。
    """
    c = CONST
    op = operating_point(d, vin, iout, model="lm5177")
    iout = op["iout"]
    iin = op["iin"]
    vout, fsw = d.vout, d.fsw
    mode = op["mode"]
    dI = op["ripple"]
    q1, q2, q3, q4 = d.q1, d.q2, d.q3, d.q4
    dt = c["deadtime"]

    # 驅動時間 (VM B471:D472)
    iin_nom = vout * d.iout / (d.vin_nom * d.efficiency)
    vsp1 = q1.vth + d.iout / q1.gfs
    vsp4 = q4.vth + iin_nom / q4.gfs
    tr_buck = (q1.qgd + q1.qgs / 2) * (c["rdrv_up"] + q1.rg) / (c["vdd"] - vsp1)
    tf_buck = (q1.qgd + q1.qgs / 2) * (c["rdrv_dn"] + q1.rg) / vsp1
    tr_boost = (q4.qgd + q4.qgs / 2) * (c["rdrv_up"] + q4.rg) / (c["vdd"] - vsp4)
    tf_boost = (q4.qgd + q4.qgs / 2) * (c["rdrv_dn"] + q4.rg) / vsp4

    # 電感
    irms_l = op["il_rms"]
    p_dcr = irms_l ** 2 * d.rdcr * (1 + c["ltc"] * (c["ta"] - 25)) / \
        (1 - irms_l ** 2 * d.rdcr * c["ltc"] * c["theta_ca"])
    p_core = abs(7.759e-3 * fsw ** 0.9458 * (0.00787 * dI) ** 2.304)
    # 電容
    icout_rms, icin_rms = op["icout_rms"], op["icin_rms"]
    if xlsm_compat:
        T = 1 / fsw
        if mode == "BUCK":
            dT, dU = op["duty"], 0.0
        elif mode == "BOOST":
            dT, dU = 1.0, op["duty"]
        else:   # Efficiency vs VIN!T/U 的原始寫法
            dT = min(1.0, fsw * op["t2"] + op["t1"])
            dU = min(1.0, fsw * op["t1"] + op["t3"])
        ac = math.sqrt(iout ** 2 + dI ** 2 / 12)
        icout_rms = dI / math.sqrt(12) if vin > vout else math.sqrt(dU / (1 - dU)) * ac
        icin_rms = math.sqrt(dT * (1 - dT)) * ac if vin > vout else dI / math.sqrt(12)
    p_cout = d.cout_esr * icout_rms ** 2
    p_cin = d.cin_esr * icin_rms ** 2
    # MOSFET
    ir = op["irms"]
    # Q1 buck control
    p_q1_cond = q1.rdson * ir["Q1"] ** 2
    if mode == "BUCK":
        p_q1_sw = vin * iout / 2 * fsw * (tr_buck + tf_buck)
    elif mode == "BOOST":
        p_q1_sw = 0.0
    else:
        il_avg = op["il_avg"]
        if xlsm_compat:   # Efficiency vs VIN!S：以各段起點電流取樣平均
            sg = op["segments"]
            il_avg = fsw * (sg[0]["i_start"] * op["t1"] + sg[1]["i_start"] * op["t2"]
                            + sg[2]["i_start"] * op["t3"])
        p_q1_sw = vin * il_avg / 2 * fsw * (tr_buck + tf_buck)
    p_q1_dt = 0.0
    if vin > vout and iout <= dI / 2:
        p_q1_dt = abs(iout - dI / 2) * q1.vsd * dt * fsw
    p_q1 = _self_heat(p_q1_cond, c["rdson_tc_top"], q1.theta_ja, p_q1_dt + p_q1_sw)
    # Q2 buck sync
    vsd2 = 0.0 if d.vf_schottky_buck > 0 else q2.vsd
    p_q2_cond = q2.rdson * ir["Q2"] ** 2
    if mode == "BOOST":
        p_q2_dt = 0.0
    elif iout < dI / 2:
        p_q2_dt = dt * fsw * vsd2 * (iout + dI / 2)
    else:
        p_q2_dt = 2 * dt * fsw * vsd2 * iout
    p_q2 = _self_heat(p_q2_cond, c["rdson_tc_bot"], q2.theta_ja, p_q2_dt)
    # Q4 boost control
    p_q4_cond = q4.rdson * ir["Q4"] ** 2
    if mode == "BUCK":
        p_q4_sw = 0.0
    elif mode == "BOOST":
        p_q4_sw = vout * iin / 2 * fsw * (tr_boost + tf_boost)
    else:
        p_q4_sw = vout * iout / 2 * fsw * (tr_boost + tf_boost)
    p_q4_dt = 0.0
    if vin <= vout and iin <= dI / 2:
        p_q4_dt = abs(iin - dI / 2) * q4.vsd * dt * fsw
    p_q4 = _self_heat(p_q4_cond, c["rdson_tc_top"], q4.theta_ja, p_q4_dt + p_q4_sw)
    # Q3 boost sync
    vsd3 = 0.0 if d.vf_schottky_boost > 0 else q3.vsd
    p_q3_cond = q3.rdson * ir["Q3"] ** 2
    if mode == "BUCK":
        p_q3_dt = 0.0
    elif iin < dI / 2:
        p_q3_dt = dt * fsw * vsd3 * (iin + dI / 2)
    else:
        p_q3_dt = 2 * dt * fsw * vsd3 * iin
    p_q3 = _self_heat(p_q3_cond, c["rdson_tc_bot"], q3.theta_ja, p_q3_dt)
    # 外接蕭特基死區導通
    p_ext_d = 0.0
    if mode in ("BUCK", "BUCK-BOOST"):
        p_ext_d += 2 * dt * fsw * d.vf_schottky_buck * iout
    if mode in ("BOOST", "BUCK-BOOST"):
        p_ext_d += 2 * dt * fsw * d.vf_schottky_boost * iin
    # 反向恢復 (xlsm 在 BB 模式把兩項各算兩次，這裡各算一次)
    qrr2 = d.qrr_schottky_buck if d.qrr_schottky_buck > 0 else q2.qrr
    qrr3 = d.qrr_schottky_boost if d.qrr_schottky_boost > 0 else q3.qrr
    p_rr_buck = qrr2 * fsw * (vin - iout * q1.rdson) if mode != "BOOST" else 0.0
    p_rr_boost = qrr3 * fsw * vout if mode != "BUCK" else 0.0
    if xlsm_compat and mode == "BUCK-BOOST":
        p_rr_buck = p_rr_boost = p_rr_buck + p_rr_boost
    # 感測電阻
    p_rs = d.rs * op["irs_rms"] ** 2
    rshunt = _rshunt(d)
    p_shunt = rshunt * iout ** 2 if (d.cfg_imon_lim > 1 or xlsm_compat) else 0.0
    # IC
    vsup = vout if vout > 8 else vin
    p_iq = c["iq"] * vsup
    if mode == "BUCK":
        qg = q1.qg + q2.qg
    elif mode == "BOOST":
        qg = q3.qg + q4.qg
    else:
        qg = q1.qg + q2.qg + q3.qg + q4.qg
    p_gate = qg * fsw * vsup
    p_ic = p_iq + p_gate

    items = dict(
        inductor=p_dcr + p_core, inductor_dcr=p_dcr, inductor_core=p_core,
        cout=p_cout, cin=p_cin, q1=p_q1, q2=p_q2, q3=p_q3, q4=p_q4,
        ext_diode=p_ext_d, rr_buck=p_rr_buck, rr_boost=p_rr_boost,
        rs=p_rs, shunt=p_shunt, ic=p_ic,
    )
    total = sum(v for k, v in items.items() if k not in ("inductor_dcr", "inductor_core"))
    pout = vout * iout
    eta = pout / (pout + total) if pout > 0 else 0.0
    return dict(op=op, losses=items, total_loss=total, pout=pout, efficiency=eta,
                tr_buck=tr_buck, tf_buck=tf_buck, tr_boost=tr_boost, tf_boost=tf_boost)


def efficiency_vs_vin(d: Design, steps: int = 200, xlsm_compat: bool = False) -> Dict:
    vins = [d.vin_min + (d.vin_max - d.vin_min) * k / steps for k in range(steps + 1)]
    rows = [losses(d, v, xlsm_compat=xlsm_compat) for v in vins]
    return _eff_table(vins, rows, "vin")


def efficiency_vs_load(d: Design, steps: int = 100, vin: Optional[float] = None) -> Dict:
    vin = d.vin_nom if vin is None else vin
    ious = [d.iout * k / steps for k in range(1, steps + 1)]
    rows = [losses(d, vin, i) for i in ious]
    return _eff_table(ious, rows, "iout")


def _eff_table(xs, rows, xname):
    return {
        xname: xs,
        "mode": [r["op"]["mode"] for r in rows],
        "efficiency": [r["efficiency"] * 100 for r in rows],
        "total_loss": [r["total_loss"] for r in rows],
        "inductor": [r["losses"]["inductor"] for r in rows],
        "buck_ctrl": [r["losses"]["q1"] + 0.66 * r["losses"]["rr_buck"] for r in rows],
        "buck_sync": [r["losses"]["q2"] + 0.33 * r["losses"]["rr_buck"] for r in rows],
        "boost_ctrl": [r["losses"]["q4"] + 0.66 * r["losses"]["rr_boost"] for r in rows],
        "boost_sync": [r["losses"]["q3"] + 0.33 * r["losses"]["rr_boost"] for r in rows],
        "rs": [r["losses"]["rs"] for r in rows],
        "duty_buck_leg": [r["op"]["d_buck_leg"] for r in rows],
        "duty_boost_leg": [r["op"]["d_boost_leg"] for r in rows],
        "ripple": [r["op"]["ripple"] for r in rows],
    }


# ---------------------------------------------------------------------------
# Main design calculation (Design Converter Step 1 ~ 7.2)
# ---------------------------------------------------------------------------

def _rshunt(d: Design) -> float:
    """VM B270 (Ω)."""
    return CONST["iout_threshold"] / d.current_setpoint if d.current_setpoint else 0.0


def calculate(d: Design, model: str = "lm5177", bode_points_per_decade: int = 200,
              with_curves: bool = True) -> Dict:
    c = CONST
    dev = DEVICES[d.device]
    vout, iout, fsw, L, rs = d.vout, d.iout, d.fsw, d.lf, d.rs
    res: Dict = {"model": model}

    # ---------------- Step 1 ----------------
    rout = vout / iout
    pout = vout * iout
    pin = pout / d.efficiency
    rt_ideal = (1 / fsw - 20e-9) * 30.3e9
    rt = e96(rt_ideal)
    tr = lm5177_transitions(fsw, vout, d.timing_device)
    if not d.mode_fpwm:
        tr_disp = {k: vout for k in ("vin_buck_to_bb", "vin_bb_to_buck",
                                     "vin_boost_to_bb", "vin_bb_to_boost")}
        tr = {**tr, **tr_disp}
    res["step1"] = dict(rout=rout, pout=pout, pin=pin,
                        iin_min=pin / d.vin_min, iin_nom=pin / d.vin_nom, iin_max=pin / d.vin_max,
                        rt_ideal=rt_ideal, rt=None if rt / 1000 > 288 else rt,
                        transitions=tr)

    ops = {k: operating_point(d, v, model=model)
           for k, v in (("min", d.vin_min), ("nom", d.vin_nom), ("max", d.vin_max))}
    res["operating_points"] = ops

    # ---------------- Step 2 : inductor ----------------
    d_ideal_min = (vout - d.vin_min) / vout if d.vin_min <= vout else vout / d.vin_min
    lopt = d.vin_min ** 2 * (vout - d.vin_min) / (0.2 * iout * fsw * vout ** 2) \
        if d.vin_min < vout else (d.vin_max - vout) * vout / (0.2 * iout * fsw * d.vin_max)
    lmin = max(10 * vout * rs / fsw, rs / 8000)
    lmax = min(rs * 35 * vout / fsw, rs / 100)
    lopt_limit = max(min(lopt, lmax), lmin)
    ripple_pct = {}
    for k, op in ops.items():
        base = iout if op["mode"] == "BUCK" else op["iin"]
        ripple_pct[k] = op["ripple"] / base * 100
    ton_min = ton_min_boost(fsw)
    c_cs = ton_min / (2 * PI * (2 * c["r_cs"]) * 10)
    res["step2"] = dict(
        lopt=lopt, lmin=lmin, lmax=lmax, l_recommended=lopt_limit,
        ripple={k: op["ripple"] for k, op in ops.items()}, ripple_pct=ripple_pct,
        duty={k: op["duty"] for k, op in ops.items()},
        mode={k: op["mode"] for k, op in ops.items()},
        rdiff=c["r_cs"], ton_min=ton_min, cdiff_ideal=c_cs, cdiff=e12(c_cs),
    )

    # ---------------- Step 3 : sense resistor, OCP, slope ----------------
    vth_min, vth, vth_max = dev["vcs_min"], dev["vcs_typ"], dev["vcs_max"]
    margin = d.ilimit_margin / 100
    if d.rs_calc_mode == 0:
        ipp = d_ideal_min * d.vin_min / (L * fsw)
        iin_avg_max = vout * iout / (d.efficiency * d.vin_min)
        rs_ideal = vth_min / ((iin_avg_max + 0.5 * ipp) * (1 + margin))
    else:
        rs_ideal = vth / (max(iout, vout / d.vin_min * iout) * 1.43)
    ipeak = {}
    for k, op in ops.items():
        if model == "xlsm":
            ipeak[k] = (iout / d.efficiency if op["mode"] == "BUCK" else op["iin"]) + op["ripple"] / 2
        else:
            ipeak[k] = op["il_peak"]
    ilpk_limit = vth_max / rs
    iout_ocp = {}
    for k, op in ops.items():
        iout_ocp[k] = (vth / rs - op["ripple"] / 2) * (1 - op["d_boost_leg"])
    p_rs_oc = (vth_max / rs) ** 2 * rs * (vout / d.vin_max)
    p_rs_oc_ds = (vth_max / rs) ** 2 * rs * (1 - vout / d.vin_max) if d.vin_max > vout else None
    rslope_ideal = L / rs * 50e6
    rslope_rec = e96(rslope_ideal)
    rslope = d.rslope if d.rslope else rslope_rec
    rs_l = rs / L
    res["step3"] = dict(
        rs_ideal=rs_ideal, ilpk_limit=ilpk_limit, ipeak=ipeak, ipeak_max=max(ipeak.values()),
        p_rs_oc=p_rs_oc, p_rs_oc_datasheet=p_rs_oc_ds, rslope_ideal=rslope_ideal,
        rslope_recommended=rslope_rec, rslope=rslope, iout_ocp=iout_ocp,
        rs_over_l=rs_l, rs_in_range=(100 < rs_l < 8000 and rs_l < fsw / (vout * 10)),
        rs_rms_full_load=max(op["irs_rms"] for op in ops.values()),
    )

    # ---------------- Step 4 : output capacitor (以 VIN(min) 為準) ----------------
    o = ops["min"]
    ox = ops["max"]
    vr = d.vout_ripple
    if o["mode"] == "BUCK":
        cout_min = o["ripple"] / (fsw * 8 * vr)
        try:
            esr_max = math.sqrt((vr / o["ripple"]) ** 2 - (1 / (fsw * d.cout * 8)) ** 2)
        except ValueError:
            esr_max = None
        vripple = ox["ripple"] * math.sqrt(d.cout_esr ** 2 + (1 / (8 * fsw * d.cout)) ** 2)
    else:
        db = o["d_boost_leg"]
        cout_min = iout * db / (vr * fsw)
        esr_max = (vr - iout * db / (fsw * d.cout)) / iout
        vripple = d.cout_esr * iout + iout * db / (fsw * d.cout)
    if esr_max is not None and (esr_max > 0.3 or esr_max < 0):
        esr_max = None
    icout = {}
    for k, op in ops.items():
        if model == "xlsm":
            icout[k] = op["ripple"] / math.sqrt(12) if op["mode"] == "BUCK" else \
                iout * math.sqrt((vout - op["vin"]) / op["vin"])
        else:
            icout[k] = op["icout_rms"]
    res["step4"] = dict(cout_min=max(cout_min, 0.1e-6), esr_max=esr_max, vripple=vripple,
                        icout_rms=icout, icout_rms_max=max(icout.values()))

    # ---------------- Step 5 : input capacitor (以 VIN(max) 為準) ----------------
    vir = d.vin_ripple
    use_half = d.vin_max >= 2 * vout
    if ox["mode"] != "BOOST":
        Dx = 0.5 if use_half else ox["d_buck_leg"]
        q = iout * Dx * (1 - Dx)
        cin_min = q / (vir * fsw)
        cin_esr_max = -(q / (d.cin * fsw) - vir) / (iout + ox["ripple"] / 2)
        vinripple = q / (d.cin * fsw) + (iout + ox["ripple"] / 2) * d.cin_esr
    else:
        cin_min = ox["ripple"] / (fsw * 8 * vir)
        try:
            cin_esr_max = math.sqrt((vir / ox["ripple"]) ** 2 - (1 / (fsw * d.cin * 8)) ** 2)
        except ValueError:
            cin_esr_max = None
        vinripple = ox["ripple"] * math.sqrt(d.cin_esr ** 2 + (1 / (8 * fsw * d.cin)) ** 2)
    if cin_esr_max is not None and (cin_esr_max > 0.3 or cin_esr_max < 0):
        cin_esr_max = None
    icin = {}
    for k, op in ops.items():
        if model == "xlsm":
            Dk = op["duty"]
            if op["mode"] == "BUCK":
                if k == "min":   # xlsm B207 混用 Don_Vinmin_2 與 Don_Vinmax_2
                    Dk2 = ops["max"]["duty"]
                    icin[k] = iout * math.sqrt(Dk * (1 - Dk2))
                else:
                    icin[k] = iout * math.sqrt(Dk * (1 - Dk))
            else:
                icin[k] = op["ripple"] / math.sqrt(12)
        else:
            icin[k] = op["icin_rms"]
    res["step5"] = dict(cin_min=max(cin_min, 10e-6), esr_max=cin_esr_max, vripple=vinripple,
                        icin_rms=icin, icin_rms_max=max(icin.values()), d_half_used=use_half)

    # ---------------- Step 6 : SS, UVLO, CFG, IMON/ILIM ----------------
    tss_rec = d.cout * vout / iout
    css_ideal = c["iss"] * d.tss / c["vref"]
    css = e12(css_ideal)
    cfg = 1 + (d.cfg_drss - 1) * 1 + (d.cfg_hiccup - 1) * 2 + \
        (1 if d.cfg_imon_lim > 2 else 0) * 4 + (d.cfg_psm - 1) * 8
    iuv = c["iuvlo"]
    ruv1 = (d.vin_uvlo_on - c["vuvlo_on"] * d.vin_uvlo_off / c["vuvlo_off"]) / iuv
    ruv2 = ruv1 / (d.vin_uvlo_off / c["vuvlo_off"] - 1) if ruv1 > 0 else None
    ruv1_std = e96(ruv1) if ruv1 > 0 else None
    ruv2_std = e96(ruv2) if ruv2 else None
    uv_on_act = uv_off_act = None
    if ruv1_std and ruv2_std:
        uv_on_act = (ruv1_std + ruv2_std) / ruv2_std * c["vuvlo_on"] + iuv * ruv1_std
        uv_off_act = (ruv1_std + ruv2_std) / ruv2_std * c["vuvlo_off"]
    imonlim = d.cfg_imon_lim - 1
    rshunt = _rshunt(d)
    vimon_max = {0: None, 1: 1.5, 2: 5.0, 3: 0.0}[imonlim]
    rimon_out = (1.5 / (d.current_setpoint * rshunt * dev["gm_imon"])) if imonlim == 1 else None
    res["step6"] = dict(
        tss_recommended=tss_rec, css_ideal=css_ideal, css=css, cfg=cfg,
        rcfg=RCFG_TABLE[cfg - 1] * 1e3, ruv1_ideal=ruv1, ruv2_ideal=ruv2,
        ruv1=ruv1_std, ruv2=ruv2_std, uvlo_on_actual=uv_on_act, uvlo_off_actual=uv_off_act,
        imonlim_select=imonlim, rshunt=rshunt, vimon_max=vimon_max, rimonout=rimon_out,
        sync_at_startup={2: "High", 3: "Low"}.get(imonlim),
        p_shunt=rshunt * iout ** 2,
    )

    # ---------------- Step 7 : compensation ----------------
    on = ops["nom"]
    vin = d.vin_nom
    ri = rs * c["csamp_gain"]
    afb = c["vref"] / vout
    # 小訊號模型使用的 duty
    if on["mode"] == "BUCK-BOOST":
        d1, d2 = on["d_buck_leg"], on["d_boost_leg"]
        boost_like = vin < vout
        d_ss = 1 - (1 - d2) / d1 if boost_like else d1 / (1 - d2)
        d_ss = min(max(d_ss, 0.0), 0.99)
        f_rhp_nom = (1 - d2) ** 2 * rout / (2 * PI * L * d1)
        rhp = True
    else:
        boost_like = on["mode"] == "BOOST"
        d_ss = on["duty"]
        f_rhp_nom = (1 - d_ss) ** 2 * rout / (2 * PI * L)
        rhp = boost_like
    om = ops["min"]
    if om["mode"] == "BUCK-BOOST":
        f_rhp_min = (1 - om["d_boost_leg"]) ** 2 * rout / (2 * PI * L * om["d_buck_leg"])
    else:
        f_rhp_min = (1 - om["duty"]) ** 2 * rout / (2 * PI * L)
    gm, cs_, gmsc = c["gm_ea"], c["cslope"], c["gm_slopecomp"]
    se_ext = (abs(vin - vout) * gmsc + c["isc_const"]) / (cs_ * rslope * ri)
    if vin > vout:
        sn, sf = (vin - vout) / L, vout / L
    else:
        sn, sf = vin / L, (vout - vin) / L
    mc_buck = 1 + se_ext / sf
    mc_boost = 1 + se_ext / sn
    k_buck = mc_buck * d_ss
    k_boost = mc_boost * (1 - d_ss)
    mc = mc_boost if boost_like else mc_buck
    k = k_boost if boost_like else k_buck
    Q = 1 / PI / (k - 0.5)
    cbw = gm / (2 * PI * c["gbw"])
    aea = gm * c["roea"]
    acs = 1 / ri
    am_buck = rout / ri / (1 + rout * (k - 0.5) / (fsw * L))
    am_boost = rout * (1 - d_ss) / (2 * ri)
    am = am_boost if boost_like else am_buck
    f_load = (2 if boost_like else 1) / (2 * PI * rout * d.cout)
    f_esr = 1 / (2 * PI * d.cout * d.cout_esr)
    fc = d.fcross
    if boost_like:
        rcomp_ideal = fc * d.cout * 2 * PI * ri / (gm * afb * (1 - d_ss) *
                                                    math.sqrt(1 + (fc / f_rhp_nom) ** 2))
    else:
        rcomp_ideal = fc * d.cout * 2 * PI * ri / (gm * afb)
    ccomp_ideal = 1 / (rcomp_ideal * (1.5 * f_load) * 2 * PI)
    chf_ideal = max(1 / (rcomp_ideal * 2 * PI * max(fc * 7, fc * 10)) - cbw, 0.0)
    rcomp = d.rcomp if d.rcomp else e96(rcomp_ideal)
    ccomp = d.ccomp if d.ccomp else e12(ccomp_ideal)
    chf = d.chf if d.chf else e12(chf_ideal)
    fp_ea = 1 / (2 * PI * (c["roea"] + rcomp) * ccomp)
    fz_comp = 1 / (2 * PI * rcomp * ccomp)
    fp_comp = 1 / (2 * PI * rcomp * (chf + cbw))
    fc_sug_parts = [fsw * 0.05]
    if d.vin_min < vout:
        fc_sug_parts.append(f_rhp_min / 4)
    if rhp:
        fc_sug_parts.append(f_rhp_nom / 4)
    rfbb_ideal = None if vout == c["vref"] else c["vref"] * d.rfbt / (vout - c["vref"])
    rfbb = None if rfbb_ideal is None else e96(rfbb_ideal)
    vout_actual = c["vref"] * (d.rfbt + rfbb) / rfbb if rfbb else c["vref"]
    bode_p = dict(adc=am * afb * aea, aea=aea, acs=acs, am=am, Q=Q, f_load=f_load,
                  f_esr=f_esr, f_rhp=f_rhp_nom, rhp=rhp, fp_ea=fp_ea, fz_comp=fz_comp,
                  fp_comp=fp_comp, f_sample=fsw / 2)
    res["step7"] = dict(
        small_signal_mode="BOOST" if boost_like else "BUCK", duty_small_signal=d_ss,
        f_load=f_load, f_esr=f_esr,
        f_rhp_min=f_rhp_min if d.vin_min < vout or om["mode"] != "BUCK" else None,
        f_rhp_nom=f_rhp_nom if rhp else None,
        fc_suggested=min(fc_sug_parts), fcross=fc, fp_ea=fp_ea,
        rfbt=d.rfbt, rfbb_ideal=rfbb_ideal, rfbb=rfbb, vout_actual=vout_actual,
        rcomp_ideal=rcomp_ideal, ccomp_ideal=ccomp_ideal, chf_ideal=chf_ideal,
        rcomp=rcomp, ccomp=ccomp, chf=chf, fz_comp=fz_comp, fp_comp=fp_comp,
        ri=ri, afb=afb, aea=aea, am=am, mc=mc, k=k, Q=Q, se_ext=se_ext, sn=sn, sf=sf,
        cbw=cbw, adc=bode_p["adc"],
    )

    xl = model == "xlsm"
    freqs = list(XLSM_FREQS) if xl else freq_grid(bode_points_per_decade)
    bode_v = voltage_loop_bode(bode_p, freqs, interpolate=not xl)
    res["step7"]["fc_bode"] = bode_v["fc"]
    res["step7"]["pm_bode"] = bode_v["pm"]

    # ---------------- Step 7.2 : current-limit loop ----------------
    if model == "xlsm":
        duty_il = vout / vin if vin >= vout else (vout - vin) / vout   # ILSTAB!B11 (理想)
    else:
        duty_il = d_ss
    il_p = dict(vin=vin, duty_il=duty_il, rshunt=rshunt)
    g1 = plant_gain_1hz_il(d, il_p)
    gm_il = dev["gm_il"]
    fcb = bode_v["fc"] or fc
    c_il = gm_il / (3 * fcb * 10 ** (g1 / 20))
    r_il = 1 / (2 * PI * c_il * 20 * 3 * fcb)
    bode_il = current_limit_bode(d, il_p, freqs, interpolate=not xl)
    res["step7_2"] = dict(cimon_recommended=c_il, cimon_std=e12(c_il),
                          rimon_recommended=r_il, rimon_std=e96(r_il),
                          cimon=d.cimon, rimon=d.rimon, plant_gain_1hz_db=g1,
                          fc_bode=bode_il["fc"], pm_bode=bode_il["pm"],
                          il_zero=bode_il["il_zero"], il_pole1=bode_il["il_pole1"],
                          il_pole2=bode_il["il_pole2"])

    # ---------------- Step 8 : efficiency ----------------
    res["step8"] = {k: dict(vin=op["vin"], **{kk: vv for kk, vv in losses(d, op["vin"]).items()
                                                   if kk != "op"})
                    for k, op in ops.items()}

    if with_curves:
        res["bode_voltage"] = bode_v
        res["bode_current_limit"] = bode_il
        res["efficiency_vs_vin"] = efficiency_vs_vin(d)
        res["efficiency_vs_load"] = efficiency_vs_load(d)
    return res


# ---------------------------------------------------------------------------
# Report / CLI
# ---------------------------------------------------------------------------

def _f(x, scale=1.0, nd=3, unit=""):
    if x is None:
        return "N/A"
    if isinstance(x, str):
        return x
    return f"{x * scale:.{nd}f} {unit}".rstrip()


def report(d: Design, r: Dict) -> str:
    s1, s2, s3, s4, s5, s6, s7, s72 = (r[k] for k in (
        "step1", "step2", "step3", "step4", "step5", "step6", "step7", "step7_2"))
    ops = r["operating_points"]
    tr = s1["transitions"]
    L = []
    a = L.append
    a(f"LM51770 Buck-Boost 計算結果  (duty model = {r['model']})")
    a("=" * 64)
    a("[Step 1] Operating specifications")
    a(f"  VIN min/nom/max = {d.vin_min} / {d.vin_nom} / {d.vin_max} V, VOUT = {d.vout} V, "
      f"IOUT = {d.iout} A, fSW = {d.fsw / 1e3:g} kHz")
    a(f"  RT = {_f(s1['rt'], 1e-3, 2, 'kΩ')}  (ideal {_f(s1['rt_ideal'], 1e-3, 2, 'kΩ')})")
    a(f"  LM5177 轉換電壓: Boost→BB {tr['vin_boost_to_bb']:.2f} V, BB→Boost "
      f"{tr['vin_bb_to_boost']:.2f} V, Buck→BB {tr['vin_buck_to_bb']:.2f} V, "
      f"BB→Buck {tr['vin_bb_to_buck']:.2f} V")
    a("  工作點 (duty):")
    for k, op in ops.items():
        extra = ""
        if op["mode"] == "BUCK-BOOST":
            extra = (f"  t1(Q1+Q3)={op['t1'] * 1e9:.0f} ns, t2(Q1+Q4)={op['t2'] * 1e9:.0f} ns, "
                     f"t3(Q2+Q3)={op['t3'] * 1e9:.0f} ns")
        a(f"    VIN {k:3s} {op['vin']:6.2f} V  {op['mode']:10s} D_buck={op['d_buck_leg']:.4f} "
          f"D_boost={op['d_boost_leg']:.4f}{extra}")
    a("[Step 2] Inductor")
    a(f"  建議 L = {s2['l_recommended'] * 1e6:.3f} µH (Lopt {s2['lopt'] * 1e6:.3f}, "
      f"Lmin {s2['lmin'] * 1e6:.3f}, Lmax {s2['lmax'] * 1e6:.3f})  選用 {d.lf * 1e6:g} µH")
    for k in ("min", "nom", "max"):
        a(f"  ΔIL @VIN {k}: {s2['ripple'][k]:.3f} A ({s2['ripple_pct'][k]:.1f} %)")
    a(f"  RDIFF = {s2['rdiff']:g} Ω, CDIFF = {s2['cdiff'] * 1e12:.0f} pF "
      f"(tON,min {s2['ton_min'] * 1e9:.1f} ns)")
    a("[Step 3] Sense resistor / OCP / slope")
    a(f"  建議 RS = {s3['rs_ideal'] * 1e3:.3f} mΩ  選用 {d.rs * 1e3:g} mΩ, "
      f"RS/L = {s3['rs_over_l']:.0f} ({'OK' if s3['rs_in_range'] else 'NG'})")
    a(f"  ILPK,limit = {s3['ilpk_limit']:.2f} A, 最大電感峰值 = {s3['ipeak_max']:.2f} A")
    a(f"  RS 過流功耗 = {s3['p_rs_oc']:.3f} W (datasheet 式24: {_f(s3['p_rs_oc_datasheet'], 1, 3, 'W')})")
    a(f"  RSLOPE = {s3['rslope'] / 1e3:.1f} kΩ (ideal {s3['rslope_ideal'] / 1e3:.2f} kΩ)")
    a("  IOUT at OCP: " + ", ".join(f"{k} {v:.2f} A" for k, v in s3["iout_ocp"].items()))
    a("[Step 4] Output capacitor")
    a(f"  COUT,min = {s4['cout_min'] * 1e6:.2f} µF, ESR max = {_f(s4['esr_max'], 1e3, 2, 'mΩ')}, "
      f"ΔVOUT = {s4['vripple'] * 1e3:.1f} mVpp, ICOUT,rms = {s4['icout_rms_max']:.2f} A")
    a("[Step 5] Input capacitor")
    a(f"  CIN,min = {s5['cin_min'] * 1e6:.2f} µF, ESR max = {_f(s5['esr_max'], 1e3, 2, 'mΩ')}, "
      f"ΔVIN = {s5['vripple'] * 1e3:.1f} mVpp, ICIN,rms = {s5['icin_rms_max']:.2f} A")
    a("[Step 6] Soft-start / UVLO / CFG")
    a(f"  tSS,rec = {s6['tss_recommended'] * 1e3:.2f} ms, CSS = {s6['css'] * 1e9:.1f} nF, "
      f"CFG #{s6['cfg']} → RCFG = {s6['rcfg'] / 1e3:g} kΩ")
    a(f"  RUV1 = {_f(s6['ruv1'], 1e-3, 1, 'kΩ')}, RUV2 = {_f(s6['ruv2'], 1e-3, 1, 'kΩ')} "
      f"(實際 ON {_f(s6['uvlo_on_actual'], 1, 3, 'V')} / OFF {_f(s6['uvlo_off_actual'], 1, 3, 'V')})")
    a(f"  RSHUNT = {s6['rshunt'] * 1e3:.2f} mΩ")
    a("[Step 7] Voltage-loop compensation")
    a(f"  小訊號模型 {s7['small_signal_mode']} (D = {s7['duty_small_signal']:.4f}), "
      f"fp,load = {s7['f_load']:.0f} Hz, fESR = {s7['f_esr'] / 1e3:.1f} kHz")
    a(f"  fRHP @VIN min = {_f(s7['f_rhp_min'], 1e-3, 2, 'kHz')}, @VIN nom = "
      f"{_f(s7['f_rhp_nom'], 1e-3, 2, 'kHz')}")
    a(f"  建議 fc = {s7['fc_suggested'] / 1e3:.2f} kHz, 期望 fc = {s7['fcross'] / 1e3:g} kHz")
    a(f"  RFBB = {_f(s7['rfbb'], 1e-3, 2, 'kΩ')}, VOUT 實際 = {s7['vout_actual']:.3f} V")
    a(f"  RCOMP = {s7['rcomp'] / 1e3:.3f} kΩ (ideal {s7['rcomp_ideal'] / 1e3:.3f}), "
      f"CCOMP = {s7['ccomp'] * 1e9:.1f} nF (ideal {s7['ccomp_ideal'] * 1e9:.2f}), "
      f"CHF = {s7['chf'] * 1e12:.0f} pF (ideal {s7['chf_ideal'] * 1e12:.0f})")
    a(f"  fz = {s7['fz_comp']:.0f} Hz, fp,HF = {s7['fp_comp'] / 1e3:.1f} kHz, fp,EA = {s7['fp_ea']:.3f} Hz")
    a(f"  Bode: fc = {_f(s7['fc_bode'], 1e-3, 3, 'kHz')}, PM = {_f(s7['pm_bode'], 1, 1, '°')}")
    a("[Step 7.2] Current-limit loop")
    a(f"  建議 CIMON = {s72['cimon_recommended'] * 1e9:.3f} nF (E12 {s72['cimon_std'] * 1e9:.2f}), "
      f"RIMON = {s72['rimon_recommended']:.0f} Ω (E96 {s72['rimon_std']:.0f})")
    a(f"  選用 CIMON = {d.cimon * 1e9:g} nF, RIMON = {d.rimon:g} Ω → fc = "
      f"{_f(s72['fc_bode'], 1e-3, 3, 'kHz')}, PM = {_f(s72['pm_bode'], 1, 1, '°')}")
    a("[Step 8] Efficiency (LM5177 duty model)")
    for k, e in r["step8"].items():
        a(f"  VIN {k:3s} {e['vin']:6.2f} V: η = {e['efficiency'] * 100:.2f} %, "
          f"Ploss = {e['total_loss']:.2f} W")
    return "\n".join(L)


def vin_report(d: Design, vin: float, model: str = "lm5177") -> str:
    """單一 VIN 的模式判斷與 duty 計算過程 (對應 HTML 的 VIN 滑桿)."""
    op = operating_point(d, vin, model=model)
    tr = lm5177_transitions(d.fsw, d.vout, d.timing_device)
    T = 1 / d.fsw
    L = [f"VIN = {vin:.3f} V → {op['mode']}  (duty model = {model})"]
    if model == "xlsm":
        L.append(f"  判斷：xlsm 原始算法，VIN {'>' if vin > d.vout else '≤'} VOUT ({d.vout:g} V)")
    elif not d.mode_fpwm:
        L.append(f"  判斷：PSM，轉換點 = VOUT ({d.vout:g} V)")
    else:
        L.append(f"  轉換電壓：上升 Boost→BB {tr['vin_boost_to_bb']:.2f} V、BB→Buck "
                 f"{tr['vin_bb_to_buck']:.2f} V；下降 Buck→BB {tr['vin_buck_to_bb']:.2f} V、"
                 f"BB→Boost {tr['vin_bb_to_boost']:.2f} V（目前依 VIN {'下降' if d.vin_direction == 'falling' else '上升'}）")
    if op["mode"] == "BUCK-BOOST":
        L.append(f"  t1 (Q1+Q3) = T − (0.175·T + 2·tmin) = {op['t1'] * 1e9:.1f} ns  (tmin {tr['t_min_pulse'] * 1e9:.1f} ns)")
        L.append(f"  t2 (Q1+Q4) = (VOUT·T − VIN·t1)/(VIN+VOUT) = {op['t2'] * 1e9:.1f} ns")
        L.append(f"  t3 (Q2+Q3) = T − t1 − t2 = {op['t3'] * 1e9:.1f} ns")
    L.append(f"  D_buck (Q1) = {op['d_buck_leg']:.4f}, D_boost (Q4) = {op['d_boost_leg']:.4f}, "
             f"VOUT/VIN = {d.vout / vin:.4f}")
    L.append(f"  IIN = {op['iin']:.3f} A, IL avg = {op['il_avg']:.3f} A, IL peak = {op['il_peak']:.3f} A, "
             f"ΔIL = {op['ripple']:.3f} A, IL rms = {op['il_rms']:.3f} A")
    L.append(f"  η (估算) = {losses(d, vin)['efficiency'] * 100:.2f} %")
    for i, sg in enumerate(op["segments"], 1):
        L.append(f"    t{i}: {sg['dt'] * 1e9:7.1f} ns ({sg['dt'] / T * 100:5.1f} %)  ON: {'+'.join(sg['on']):6s} "
                 f"iL {sg['i_start']:.2f} → {sg['i_end']:.2f} A")
    return "\n".join(L)


def _apply_set(d: Design, items: List[str]) -> None:
    for it in items:
        key, val = it.split("=", 1)
        obj = d
        parts = key.split(".")
        for p in parts[:-1]:
            obj = getattr(obj, p)
        cur = getattr(obj, parts[-1])
        if isinstance(cur, bool):
            v = val.lower() in ("1", "true", "yes")
        elif isinstance(cur, (int, float)) or cur is None:
            v = float(val)
            if isinstance(cur, int) and not isinstance(cur, bool):
                v = int(v)
        else:
            v = val
        setattr(obj, parts[-1], v)


def _plot(r: Dict, outdir: str) -> None:
    import os
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    os.makedirs(outdir, exist_ok=True)
    for key, title in (("bode_voltage", "Voltage loop"), ("bode_current_limit", "Current-limit loop")):
        b = r[key]
        fig, ax = plt.subplots(figsize=(8, 5))
        ax.semilogx(b["f"], b["gain"], label="Gain (dB)")
        ax2 = ax.twinx()
        ax2.semilogx(b["f"], b["phase"], "C1", label="Phase (°)")
        ax.axhline(0, color="gray", lw=0.6)
        ax.set_xlabel("Frequency (Hz)")
        ax.set_ylabel("Gain (dB)")
        ax2.set_ylabel("Phase (°)")
        ax.set_title(f"{title}: fc = {b['fc'] or 0:.0f} Hz, PM = {b['pm'] or 0:.1f}°")
        ax.grid(True, which="both", alpha=0.3)
        fig.savefig(os.path.join(outdir, f"{key}.png"), dpi=120, bbox_inches="tight")
        plt.close(fig)
    e = r["efficiency_vs_vin"]
    fig, ax = plt.subplots(figsize=(8, 5))
    ax.plot(e["vin"], e["efficiency"])
    ax.set_xlabel("VIN (V)")
    ax.set_ylabel("Efficiency (%)")
    ax.grid(True, alpha=0.3)
    fig.savefig(os.path.join(outdir, "efficiency_vs_vin.png"), dpi=120, bbox_inches="tight")
    plt.close(fig)


def main(argv=None):
    ap = argparse.ArgumentParser(description="LM51770 buck-boost power-train calculator")
    ap.add_argument("--model", choices=("lm5177", "xlsm"), default="lm5177",
                    help="duty 模型 (預設 lm5177)")
    ap.add_argument("--set", action="append", default=[], metavar="NAME=VALUE",
                    help="修改輸入，如 vin_nom=20、q1.rdson=0.004 (SI 單位)")
    ap.add_argument("--vin", type=float, action="append", metavar="V",
                    help="只列出這個 VIN 的模式 (Buck / Buck-Boost / Boost) 與 duty，可重複")
    ap.add_argument("--json", metavar="FILE", help="輸出全部結果為 JSON")
    ap.add_argument("--plot", metavar="DIR", help="輸出 PNG 圖 (需要 matplotlib)")
    args = ap.parse_args(argv)
    d = Design()
    _apply_set(d, args.set)
    if args.vin:
        print("\n\n".join(vin_report(d, v, args.model) for v in args.vin))
        return
    r = calculate(d, model=args.model)
    print(report(d, r))
    if args.json:
        with open(args.json, "w", encoding="utf-8") as fh:
            json.dump({"inputs": dataclasses.asdict(d), "results": r}, fh,
                      ensure_ascii=False, indent=1)
    if args.plot:
        _plot(r, args.plot)


if __name__ == "__main__":
    main()
