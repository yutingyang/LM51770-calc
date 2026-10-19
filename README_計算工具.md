# LM51770 Buck-Boost 計算工具（Python / HTML）

由 `LM51770_Buck-Boost_Quickstart_Tool_V1_0_1(3).xlsm` 轉換。公式依據 `LM51770_Design_Converter_分析.md` 與 `LM51770_xlsm_全分頁分析.md`。

| 檔案 | 內容 |
| --- | --- |
| `lm51770_calc.py` | Python 計算核心與 CLI（只用標準函式庫；`--plot` 需要 matplotlib） |
| `LM51770_calc.html` | 單一 HTML 檔，可離線開啟。內含與 Python 相同的計算核心（JS），提供輸入表單、結果、Bode 圖、效率圖與電感電流波形 |
| `test_lm51770_calc.py` | 讀 `LM51770_xlsm_全儲存格清單.csv` 的 xlsm 快取值，逐格驗證 |

## 使用

```bash
python3 lm51770_calc.py                          # 預設輸入 = xlsm 存檔值
python3 lm51770_calc.py --model xlsm             # 用 xlsm 原始 duty 算法（與 Design Converter 完全相同）
python3 lm51770_calc.py --set vin_nom=18 --set q1.rdson=0.004   # 改輸入（SI 單位）
python3 lm51770_calc.py --json result.json       # 全部結果，含 Bode / 效率曲線
python3 -m unittest test_lm51770_calc -v         # 對 xlsm 驗證
```

### VIN 滑桿（單點模式與 duty）

HTML 最上方「基本參數與 VIN 工作點」：左邊是基本參數（與左側 Step 輸入同步），右邊拉 VIN 滑桿（範圍 VIN(min)~VIN(max)，也可直接輸入或按 min / nom / max），即時顯示：

* 該 VIN 落在 **BUCK / BUCK-BOOST / BOOST** 哪一區，以及判斷依據（與哪個轉換電壓比較）
* D Buck 臂（Q1 導通比）、D Boost 臂（Q4 導通比）、代入數值的 duty 公式
* 模式區間圖（VIN 上升 / 下降兩條，標出四個轉換電壓與目前 VIN）
* 三相位 t1 / t2 / t3 的時間、Q1~Q4 開關狀態、電感電流斜率與波形、估算效率

Python 對應：`python3 lm51770_calc.py --vin 12 --vin 16 --vin 30`

HTML 直接用瀏覽器開啟 `LM51770_calc.html`；輸入會存在瀏覽器 localStorage，「回復 xlsm 預設值」可重設，「匯出 JSON」可下載結果。

## 計算流程（同 Design Converter）

1. **操作規格**：ROUT、PIN（η = 95 %）、IIN、RT = (1/fSW − 20 ns)·30.3 GΩ/s → E96。
2. **工作點 / duty**（見下節）：VIN(min/nom/max) 的模式、duty、電感電流波形。
3. **電感**：Lopt（式 20）夾在 Lmin/Lmax（式 13、14）之間；ΔIL、漣波百分比；CDIFF = tON,min/(2π·20 Ω·10)。
4. **RS / OCP / Slope**：建議 RS、ILPK,Limit = VCS,max/RS、最大電感峰值、過流 RS 功耗（同時列 datasheet 式 24）、RSLOPE = L/RS·50 MV/(A·s)、OCP 時 IOUT。
5. **輸出電容**（以 VIN(min) 為準）、**輸入電容**（以 VIN(max) 為準，VIN(max) ≥ 2·VOUT 時取 D = 0.5）。
6. **軟啟動 / UVLO / CFG / IMON**：CSS、RUV1/RUV2（E96）與實際門檻、CFG 編號與 RCFG、RSHUNT。
7. **電壓迴路補償**：負載極點、ESR 零點、RHP 零點、建議 fc、RCOMP/CCOMP/CHF（標準值），Bode 圖求交越頻率與相位裕度。
8. **平均電流限制迴路**：建議 CIMON / RIMON（3×fc 經驗式），複數阻抗模型 Bode 圖。
9. **效率**：各 MOSFET（導通、切換、死區、反向恢復，含 RDS(on) 溫升修正）、電感 DCR + 鐵損、電容 ESR、RS、RSHUNT、IC 損耗；效率對 VIN / 對負載曲線。

## Duty：LM5177 計算方式（預設 `--model lm5177`）

LM51770 沿用 LM5177 的四開關時序（xlsm `Timings` 分頁，`Device = 1: LM5177`）：

* **轉換電壓**：以最小脈寬 tmin（N5）、BB 視窗 0.175·T + 2·tmin 算出
  Boost→BB、BB→Boost、Buck→BB、BB→Buck 四個電壓（範例 14.20 / 11.13 / 18.02 / 23.00 V）。
  有遲滯，預設依 VIN 上升方向判斷（與 xlsm 效率頁相同），可改 `vin_direction=falling`。PSM 模式下全部簡化為 VOUT。
* **BUCK**：D = (VOUT + (RDS2+RS+DCR)·IOUT) / (VIN + (RS+RDS2−RDS1)·IOUT)
* **BOOST**：D = (VOUT + (RDS3+DCR)·IIN − VIN) / (VOUT + (RDS3−RDS4−RS)·IIN)
* **BUCK-BOOST**（三相位）：
  * t1：Q1+Q3 直通，固定 = T − (0.175·T + 2·tmin)，斜率 (VIN−VOUT)/L
  * t2：Q1+Q4 充磁，t2 = (VOUT·T − VIN·t1)/(VIN+VOUT)（伏秒平衡），斜率 VIN/L
  * t3：Q2+Q3 放磁，t3 = T − t1 − t2，斜率 −VOUT/L
  * D_buck 臂 = (t1+t2)/T、D_boost 臂 = t2/T；VOUT/VIN = D_buck/(1 − D_boost)
  * 電感谷值由輸入電荷 IIN·T = ∫(t1+t2) iL dt 求得。
* 峰值、漣波、各 MOSFET / CIN / COUT / RS 的 RMS 電流都由分段線性電感電流波形積分。
* 小訊號模型：VIN(nom) 落在 BB 區時，以等效 duty（VIN<VOUT：1 − (1−D_boost)/D_buck）套用 boost 模型，
  RHP 零點 = (1−D_boost)²·ROUT/(2π·L·D_buck)；VIN ≥ VOUT 時沿用 xlsm 的 buck 模型，另加上該 RHP 零點。這是近似。

`--model xlsm` 則完全照 `Variable Management`：只以 VIN > VOUT 分 buck/boost，duty 只含 DCR 與 RS 壓降。

## 驗證

* `--model xlsm` 重現 Design Converter 全部顯示值（E12 ~ E87、B91 ~ G93、K91 ~ L92）、Timings 轉換電壓、
  Stability / IL Stability 的增益相位、交越頻率與相位裕度。
* 效率 `xlsm_compat=True` 逐列重現 `Efficiency vs VIN Calculations`（含 buck-boost 列）。
* HTML 內的 JS 核心與 Python 對 12 組設計 / 模型逐項比對，相對誤差 < 1e-14。

## 與 xlsm 的差異（預設 lm5177 模型）

| 項目 | xlsm | 本工具（lm5177） |
| --- | --- | --- |
| 模式判斷 | Step 2~7 只看 VIN > VOUT | 依 LM5177 轉換電壓分 BUCK / BUCK-BOOST / BOOST |
| Duty | 只含 DCR、RS | 另含 RDS(on)；BB 為三相位 |
| 電容 / 開關 RMS | 封閉式近似 | 由實際波形積分 |
| Bode 交越頻率 | 取掃頻點（4 kHz 附近間隔 0.04 decade） | 每 decade 200 點並內插 |
| 效率頁 BB 欄 | `Fsw*t2+t1` 少括號、Qrr 重複計算 | 修正 |
| RSHUNT 損耗 | IMON/ILIM 關閉也計入 | 關閉時不計 |
