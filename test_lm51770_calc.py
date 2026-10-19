#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""以 xlsm 存檔快取值 (LM51770_xlsm_全儲存格清單.csv) 驗證 lm51770_calc.py.

    python3 -m unittest test_lm51770_calc -v
"""

import csv
import math
import os
import unittest

import lm51770_calc as lc

HERE = os.path.dirname(os.path.abspath(__file__))
CSV = os.path.join(HERE, "LM51770_xlsm_全儲存格清單.csv")


def load_cells():
    cells = {}
    with open(CSV, encoding="utf-8-sig") as fh:
        for row in csv.reader(fh):
            cells[(row[0], row[1])] = row[6]
    return cells


CELLS = load_cells() if os.path.exists(CSV) else {}


def xl(sheet, cell):
    return float(CELLS[(sheet, cell)])


@unittest.skipUnless(CELLS, "xlsm cell list not found")
class TestAgainstXlsm(unittest.TestCase):
    """duty_model='xlsm' 必須重現 Design Converter 的數值."""

    @classmethod
    def setUpClass(cls):
        cls.d = lc.Design()
        cls.r = lc.calculate(cls.d, model="xlsm")

    def check(self, cell, value, rel=2e-3, sheet="Design Converter"):
        ref = xl(sheet, cell)
        self.assertTrue(math.isclose(value, ref, rel_tol=rel, abs_tol=1e-9),
                        f"{sheet}!{cell}: python {value!r} != xlsm {ref!r}")

    def test_design_converter(self):
        r = self.r
        s1, s2, s3, s4, s5, s6, s7, s72 = (r[k] for k in (
            "step1", "step2", "step3", "step4", "step5", "step6", "step7", "step7_2"))
        cases = {
            "E12": s1["rt"] / 1e3,
            "E15": s2["l_recommended"] * 1e6,
            "E18": s2["ripple"]["min"], "E19": s2["ripple"]["nom"], "E20": s2["ripple"]["max"],
            "E21": s2["ripple_pct"]["min"], "E22": s2["ripple_pct"]["nom"],
            "E23": s2["ripple_pct"]["max"],
            "E24": s2["rdiff"], "E25": s2["cdiff"] * 1e12,
            "E29": s3["rs_ideal"] * 1e3, "E31": s3["ilpk_limit"], "E32": s3["ipeak_max"],
            "E33": s3["p_rs_oc"], "E34": s3["rslope_recommended"] / 1e3,
            "E36": s3["iout_ocp"]["min"], "E37": s3["iout_ocp"]["nom"],
            "E38": s3["iout_ocp"]["max"],
            "E42": s4["cout_min"] * 1e6, "E44": s4["esr_max"] * 1e3,
            "E46": s4["vripple"] * 1e3, "E47": s4["icout_rms_max"],
            "E51": s5["cin_min"] * 1e6, "E53": s5["esr_max"] * 1e3,
            "E55": s5["vripple"] * 1e3, "E56": s5["icin_rms_max"],
            "E59": s6["tss_recommended"] * 1e3, "E61": s6["css"] * 1e9,
            "E63": s6["cfg"], "E64": s6["rcfg"] / 1e3,
            "E67": s6["ruv1"] / 1e3, "E68": s6["ruv2"] / 1e3, "E72": s6["rshunt"] * 1e3,
            "E78": s7["f_load"], "E79": s7["f_esr"] / 1e3,
            "E80": s7["f_rhp_min"] / 1e3, "E81": s7["f_rhp_nom"] / 1e3,
            "E82": s7["fc_suggested"] / 1e3, "E84": s7["fp_ea"],
            "E86": s7["rfbb"] / 1e3, "E87": s7["vout_actual"],
            "B91": s7["rcomp_ideal"] / 1e3, "C91": s7["rcomp"] / 1e3,
            "B92": s7["ccomp_ideal"] * 1e9, "C92": s7["ccomp"] * 1e9,
            "B93": s7["chf_ideal"] * 1e12, "C93": s7["chf"] * 1e12,
            "G92": s7["fz_comp"] / 1e3, "G93": s7["fp_comp"] / 1e3,
            "K91": s72["cimon_recommended"] * 1e9, "L91": s72["cimon_std"] * 1e9,
            "K92": s72["rimon_recommended"] / 1e3, "L92": s72["rimon_std"] / 1e3,
        }
        for cell, val in cases.items():
            with self.subTest(cell=cell):
                self.check(cell, val)

    def test_transitions(self):
        tr = self.r["step1"]["transitions"]
        for cell, key in (("F2", "vin_buck_to_bb"), ("F3", "vin_bb_to_buck"),
                          ("F4", "vin_boost_to_bb"), ("F5", "vin_bb_to_boost"),
                          ("W15", "t1_bb")):
            with self.subTest(cell=cell):
                self.check(cell, tr[key], rel=1e-9, sheet="Timings")

    def test_bode(self):
        s = "Stability Calculations"
        self.check("Q4", self.r["step7"]["fc_bode"] / 1e3, sheet=s)
        self.check("Q5", self.r["step7"]["pm_bode"], sheet=s)
        self.check("U4", self.r["step7_2"]["fc_bode"] / 1e3, sheet="IL Stability Calculations")
        self.check("U5", self.r["step7_2"]["pm_bode"], sheet="IL Stability Calculations")
        self.check("BE13", self.r["step7_2"]["plant_gain_1hz_db"], rel=1e-6,
                   sheet="IL Stability Calculations")
        bv = self.r["bode_voltage"]
        bi = self.r["bode_current_limit"]
        for k in (0, 50, 100, 200, 300, 400, 500, 600):
            row = 13 + k
            with self.subTest(row=row):
                self.check(f"W{row}", bv["gain"][k], rel=1e-6, sheet=s)
                self.check(f"X{row}", bv["phase"][k] if bv["phase"][k] >= -180
                           else bv["phase"][k] + 360, rel=1e-6, sheet=s)
                self.check(f"AW{row}", bi["gain"][k], rel=1e-6, sheet="IL Stability Calculations")

    def test_efficiency_vs_vin(self):
        """xlsm_compat=True 須逐列重現效率頁；預設 (修正版) 與其差異應很小."""
        s = "Efficiency vs VIN Calculations"
        e = lc.efficiency_vs_vin(self.d, xlsm_compat=True)
        e2 = lc.efficiency_vs_vin(self.d)
        for row in range(9, 210, 5):
            k = row - 9
            ref_mode = CELLS[(s, f"I{row}")]
            with self.subTest(row=row, mode=ref_mode):
                self.assertEqual(e["mode"][k].replace("BUCK-BOOST", "BUCK-BST"), ref_mode)
                self.check(f"V{row}", e["ripple"][k], rel=1e-9, sheet=s)
                self.check(f"BO{row}", e["efficiency"][k], rel=1e-9, sheet=s)
                self.check(f"BP{row}", e["inductor"][k], rel=1e-9, sheet=s)
                self.check(f"BR{row}", e["buck_ctrl"][k], rel=1e-9, sheet=s)
                self.check(f"BU{row}", e["boost_sync"][k], rel=1e-9, sheet=s)
                # 預設模型修正了 BB 模式 Qrr 重複計算，效率約高 0.5 %
                self.check(f"BO{row}", e2["efficiency"][k], rel=1e-2, sheet=s)


class TestLM5177Duty(unittest.TestCase):
    def test_bb_volt_second_balance(self):
        d = lc.Design()
        for vin in (14.5, 16.0, 18.0, 22.0):
            op = lc.operating_point(d, vin, model="lm5177")
            self.assertEqual(op["mode"], "BUCK-BOOST")
            T = 1 / d.fsw
            self.assertAlmostEqual(op["t1"] + op["t2"] + op["t3"], T, places=15)
            # 電感伏秒平衡 → 電流回到起點
            seg = op["segments"]
            self.assertAlmostEqual(seg[-1]["i_end"], seg[0]["i_start"], places=9)
            # 理想轉換比 VOUT/VIN = D_buck / (1 − D_boost)
            m = op["d_buck_leg"] / (1 - op["d_boost_leg"])
            self.assertAlmostEqual(m, d.vout / vin, places=9)
            # 輸入電荷守恆：Q1 導通期間的平均電流 = Iin
            self.assertAlmostEqual(op["i_in_avg"], op["iin"], places=9)

    def test_hysteresis(self):
        d = lc.Design()
        self.assertEqual(lc.operating_point(d, 12.0)["mode"], "BOOST")
        d.vin_direction = "falling"
        self.assertEqual(lc.operating_point(d, 12.0)["mode"], "BUCK-BOOST")
        self.assertEqual(lc.operating_point(d, 20.0)["mode"], "BUCK")

    def test_psm_mode(self):
        d = lc.Design(mode_fpwm=False)
        self.assertEqual(lc.operating_point(d, 15.0)["mode"], "BOOST")
        self.assertEqual(lc.operating_point(d, 17.0)["mode"], "BUCK")

    def test_rms_waveform_matches_closed_form(self):
        d = lc.Design()
        op = lc.operating_point(d, 36.0)
        self.assertEqual(op["mode"], "BUCK")
        self.assertAlmostEqual(op["icout_rms"], op["ripple"] / math.sqrt(12), places=9)


if __name__ == "__main__":
    unittest.main()
