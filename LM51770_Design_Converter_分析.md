# LM51770 Buck-Boost Quickstart Tool — Design Converter 計算流程分析

> 分析對象：`LM51770_Buck-Boost_Quickstart_Tool_V1_0_1(3).xlsm`（Version 1.0.1）
> 對照資料：`lm51770x/lm51770.md`（LM51770x datasheet, Rev. A, 2025/08）與 `lm51770x/zh-TW/*.md`
> 範例數值：取自 xlsm 目前存檔時的值（VIN = 6 / 12 / 36 V、VOUT = 16 V、IOUT = 8 A、fSW = 398 kHz），與 datasheet §9.2 典型應用相同。

---

## 0. 總覽

### 0.1 工作表之間的關係

```
             ┌──────────────────────────┐
  使用者輸入 → │   Design Converter (DC)  │ ← 顯示結果
             └────────────┬─────────────┘
                          │ 幾乎所有輸入先被搬到 Variable Management（並換成 SI 單位）
                          ▼
             ┌──────────────────────────┐        ┌────────────────────────────┐
             │   Variable Management    │ ◀────▶ │ Standard Value Calculator  │
             │  (VM，核心計算/具名範圍)    │        │ (E96 電阻 / E12 電容取標準值)  │
             └─┬──────────┬──────────┬──┘        └────────────────────────────┘
               │          │          │
               ▼          ▼          ▼
       DeviceParameter  CFG Setting  Timings
       (IC 參數/門檻)   (CFG 編碼表)  (Buck-Boost 轉換區邊界)
               │
               ▼
   Stability Calculations / IL Stability Calculations（Bode 圖、交越頻率、相位裕度）
   Efficiency vs VIN / vs Load Calculations（效率與損耗）
```

- **Design Converter**：只有使用者輸入格（白色 Input Box）與顯示格；顯示格大多直接取 `Variable Management` 某格，或取一個「具名範圍 (Defined Name)」。
- **Variable Management**：所有公式的核心。具名範圍（如 `VIN_min`、`Lf`、`Rs`）幾乎都指向此表 B 欄。
- **Standard Value Calculator**：把計算出的理想值四捨五入到 **E96 電阻**（`O` 欄）或 **E12 電容**（`B~I` 欄）。
- **DeviceParameter**：依 `Device`（LM51770 / LM517701 …）選 IC 參數（CS 門檻、gm 等），並含 tON,min vs fSW 查表。
- **CFG Setting**：CFG 腳位 16 種組合的對照表（= datasheet Table 8-1）。
- **Timings**：計算 Buck↔Buck-Boost↔Boost 轉換電壓（`VIN_edgebuckboost` 等）。

### 0.2 工具內寫死的常數（與 datasheet 對照）

| 名稱 (VM 位置) | 工具值 | Datasheet 值 (§6.5) | 備註 |
| --- | --- | --- | --- |
| `Efficiency` (B16) | 0.95 | 式 22 估計 95% | 所有輸入電流都用它估算 |
| `Vref` (B24) | 1 V | V_REF 0.99/1/1.01 V | |
| `Ifb` (B27) | 60 nA | FB leakage max 60 nA | 只用於 `Rfb2_max` |
| `Mode` (B30) | 1 | — | 0: PSM / 1: FPWM |
| `Calc_mode` (L14) | 1 | — | **1 = "Joerg" 算法 / 0 = LM5177(datasheet 式 23)**，影響 Rs 建議值 |
| CS 門檻 min/typ/max (DeviceParameter B16:B18) | 42.5 / 50 / 57.5 mV | V_th+(CSB−CSA) 42.5 / 50 / 57.5 mV | LM517701 為 65.6 / 75 / 84.4 mV |
| `Iss` (B233) | 10 µA | I_SS 8.9/10/11 µA | |
| `Vuvlo_on` (B247) | 1.28 V | V_T+(UVLO) 1.22/**1.25**/1.28 V | 工具取 **max** |
| `Vuvlo_off` (B248) | 1.17 V | V_T−(UVLO) 1.17/**1.2**/1.23 V | 工具取 **min** |
| `Iuvlo1` (B250) | 0.006 mA = 6 µA | I_UVLO 4/**5**/6 µA | 工具取 **max** |
| `Iout_threshold` (B269) | 50 mV | ΔV(ISNS) 49/50/51.7 mV | 平均電流限制門檻 |
| `g_m`(IMON) (DeviceParameter B11) | 1 mS | gm(IMON) 0.9/1/1.1 mS | |
| `gm_IL` (DeviceParameter B12) | 200 µS | gm(ILIM) 170/200/220 µS | |
| `CSAmp_gain` (B297) | 10 | 式 3 中的「10 × R(CS)」 | `Ri = 10·Rs` |
| `gm` EA (B299) | 600 µS | EA Transconductance 600 µS | |
| `GBW` (B347) / `RoEA` (B349) | 13 MHz / 10 MΩ | 未列 | 誤差放大器模型參數 |
| `GBW_IL` / `RoEA_IL` | 2 MHz / 20 MΩ | 未列 | 電流限制放大器模型 |
| `Cslope`/`gm_slopecomp` | 2.5 pF / 1.875 mS | 未列 | 只用於 COMP 電壓範圍與 mc 估算 |
| `R_CS`（CS 濾波電阻, DeviceParameter N8） | 10 Ω | 建議 R_DIFF1/2 = 10 Ω | |

---

## Step 1：Operating Specifications（操作規格）

| DC 儲存格 | 項目 | 類型 | 引用來源 / 公式 | 範例值 | Datasheet |
| --- | --- | --- | --- | --- | --- |
| E6 / E7 / E8 | VIN(min/nom/max) | 輸入 | → VM `B7/B8/B9` (`VIN_min`、`VIN_nom`=`Vin`、`VIN_max`) | 6 / 12 / 36 V | Table 9-1 |
| E9 | VOUT | 輸入 | → VM `B12` (`Vout`) | 16 V | |
| E10 | IOUT(max) | 輸入 | → VM `B13` (`Iout`) | 8 A | |
| E11 | 切換頻率 | 輸入 (kHz) | → VM `B22` `Fsw = E11×1000` | 398 kHz | |
| E12 | RT 電阻 | 計算 | `=IF('Standard Value Calculator'!O4/1000>288,"N/A", O4/1000)`；O4 = E96(`Rfadj`) | 75 kΩ | 式 2 / 式 18 |

**衍生變數（VM，後面所有步驟都會用到）**

$$R_{OUT}=\frac{V_{OUT}}{I_{OUT}},\quad P_{OUT}=V_{OUT}I_{OUT},\quad P_{IN}=\frac{P_{OUT}}{\eta}\ (\eta=0.95)$$

$$I_{IN}(V)=\frac{P_{IN}}{V}\quad\Rightarrow\quad I_{IN,min}=22.46\,A,\ I_{IN,nom}=11.23\,A,\ I_{IN,max}=3.74\,A$$

**RT 公式**（VM B26 `Rfadj`）：

$$R_T=\left(\frac{1}{f_{SW}}-20\,ns\right)\times 30.3\,\frac{G\Omega}{s}=75.52\,k\Omega\ \xrightarrow{E96}\ 75.0\,k\Omega$$

與 datasheet 式 2 完全相同。（datasheet §9.2.2.2 文字寫 400 kHz 選 78.7 kΩ，但用式 18 算出約 75.1 kΩ，Table 9-3 也列 75 kΩ；工具結果與 75 kΩ 一致。）RT 大於 288 kΩ（約 fSW < 105 kHz）時工具顯示 N/A。

**操作模式判斷**：工具在每個 VIN 點以 `IF(VIN>Vout,"BUCK","BOOST")` 判斷模式；VIN_nom 另外會跟 `Timings` 表算出的 `VIN_edgebuckboost`（= `VIN_BB_to_boost` = 11.13 V）比較，用以標示「BUCK-BOOST」區。

---

## Step 2：Filter Inductor（電感）

| DC 儲存格 | 項目 | 類型 | 引用來源 | 範例值 |
| --- | --- | --- | --- | --- |
| E15 | 建議電感值 | 計算 | `=Lopt_Limit×1e6` → VM `B74` | 1.829 µH |
| E16 | 電感值 LF | 輸入 | → VM `B75` `Lf` | 1.8 µH |
| E17 | 電感 DCR | 輸入 | → VM `B76` `Rdcr` | 3.2 mΩ |
| E18 / E19 / E20 | ΔIL @ VIN(min/nom/max) | 計算 | VM `B87/B88/B89` | 5.28 / 4.23 / 12.43 A |
| E21 / E22 / E23 | ΔIL 百分比 | 計算 | VM `B99/B100/B101 ×100` | 23.5 / 37.7 / 155.4 % |
| E24 | 建議 RDIFF | 常數 | `=R_CS` → DeviceParameter `N8` | 10 Ω |
| E25 | 建議 CDIFF | 計算 | `=C_CS_ideal×1e12` → DeviceParameter `N10` → Std Value `O14` | 180 pF |

### 2.1 建議電感值 E15

1. **Datasheet 公式**（VM B71 `Lopt`，= datasheet 式 20，ΔIL = 20 %·I_L at VIN(min)）：

$$L_{opt}=\frac{V_{IN(min)}^2\,(V_{OUT}-V_{IN(min)})}{0.2\,I_{OUT}\,f_{SW}\,V_{OUT}^2}=2.208\,\mu H$$

2. **上下限**（VM O65/O66，來自 datasheet 斜率補償準則式 13、14）：

$$L_{min}=\max\!\left(\frac{10\,V_{OUT}R_S}{f_{SW}},\ \frac{R_S}{8000}\right)=0.523\,\mu H$$

$$L_{max}=\min\!\left(\frac{35\,V_{OUT}R_S}{f_{SW}},\ \frac{R_S}{100}\right)=1.829\,\mu H$$

   - $R_S/L<\frac{f_{SW}}{10V_O}$ ⇔ $L>\frac{10V_OR_S}{f_{SW}}$（式 13）
   - $100\,Hz<R_S/L<8000\,Hz$ ⇔ $R_S/8000<L<R_S/100$（式 14）
   - 「35·VOUT·RS/fSW」是工具自訂的上限（datasheet 沒有）。

3. **最終**：`Lopt_Limit = MAX(MIN(Lopt, Lmax), Lmin)` = 1.829 µH（本例被 Lmax 截斷）。

> ⚠️ 注意循環引用：Step 2 的 Lmin/Lmax 使用的是 **Step 3 使用者選定的 RS（E30）**，所以改 RS 會反過來改變建議電感值。

VM B61~B67 另有「30 % / 40 % / 60 % 漣波」的舊版算法（`L_Vinmin/nom/max`），目前未顯示在 DC 上。

### 2.2 含損耗的 duty cycle（VM B81~B83，Step 2~5 全部使用）

$$D_{buck}=\frac{V_{OUT}+I_{OUT}(R_{DCR}+R_S)}{V_{IN}+I_{OUT}R_S}\qquad(V_{IN}>V_{OUT})$$

$$D_{boost}=\frac{V_{OUT}-V_{IN}+I_{IN}R_{DCR}}{V_{OUT}-I_{IN}R_S}\qquad(V_{IN}\le V_{OUT})$$

範例：D(6 V)=0.6306（boost）、D(12 V)=0.2525（boost）、D(36 V)=0.4453（buck）。
（理想值為 0.625 / 0.25 / 0.444；datasheet 式 21 用理想值。）

### 2.3 漣波電流 E18~E20

$$\Delta I_L=\begin{cases}\dfrac{(V_{IN}-V_{OUT})\,D_{buck}}{L\,f_{SW}} & V_{IN}>V_{OUT}\\[2ex]\dfrac{V_{IN}\,D_{boost}}{L\,f_{SW}} & V_{IN}\le V_{OUT}\end{cases}$$

Boost 式 = datasheet 式 21 $I_{L(PP)}=(1-\frac{V_{IN}}{V_{OUT}})\frac{V_{IN}}{Lf_{SW}}$（datasheet 5.23 A，工具含損耗 5.28 A）。

### 2.4 漣波百分比 E21~E23

$$\%\Delta I_L=\begin{cases}\Delta I_L/I_{OUT} & \text{buck}\\ \Delta I_L/I_{IN} & \text{boost}\end{cases}$$

### 2.5 電流感測差動濾波 E24 / E25（datasheet §8.3.12 式 16）

$$C_{DIFF}=\frac{t_{on,min}}{2\pi\,(R_{DIFF1}+R_{DIFF2})\cdot 10}=\frac{209.2\,ns}{2\pi\cdot 20\,\Omega\cdot 10}=166.5\,pF\ \xrightarrow{E12}\ 180\,pF$$

- `t_on,min` 來自 **DeviceParameter N6**：`VLOOKUP(Fsw/1e6, P7:Q107, 2)`，是 Boost 最小 on-time 對 fSW 的查表（0.1 MHz→435 ns … 2.2 MHz→86 ns）。datasheet §6.5 僅給出 RT=16.2 kΩ 時 88 ns。
- 與 datasheet §9.2.2.4「In this configuration 180pF is used」一致。

---

## Step 3：OCP, Sense Resistors, Slope Comp（過電流、感測電阻、斜率補償）

| DC 儲存格 | 項目 | 類型 | 引用來源 | 範例值 |
| --- | --- | --- | --- | --- |
| E28 | 峰值電流限制裕度 | 輸入 | → VM `B109` `Ilimit_Margin` | 20 % |
| E29 | 建議感測電阻 | 計算 | `=Rs_ideal×1000` → VM `B145` | 1.639 mΩ |
| E30 | 感測電阻 RS | 輸入 | → VM `B146` `Rs` | 1.3 mΩ |
| E31 | 電感峰值過流限制 ILPK,Limit | 計算 | VM `B159` | 44.23 A |
| E32 | 最大電感峰值電流 | 計算 | `=Ipeak_inductor_max` → VM `B115` | 25.10 A |
| E33 | 過流時 RS 功耗 | 計算 | VM `B162` | 1.13 W |
| E34 | 建議 SLOPE 電阻 | 計算 | VM `B165/1000` → Std Value `O15` | 69.8 kΩ |
| E35 | SLOPE 電阻 RSLOPE | 計算(同 E34) | VM `B165/1000`（→ VM `B166` `Rslope`） | 69.8 kΩ |
| E36 / E37 / E38 | OCP 起始時的 IOUT | 計算 | VM `B151/B152/B153` | 13.23 / 27.17 / 32.25 A |

使用的 IC 參數（DeviceParameter → VM）：`Buck/Boost_CS_threshold_min/typ/max` = 42.5 / 50 / 57.5 mV（VM B118~B124 → DeviceParameter B16~B18）。

### 3.1 建議感測電阻 E29（VM B145，依 `Calc_mode` 切換）

**Calc_mode = 1（目前設定，"Joerg" 算法）**：

$$R_{S,ideal}=\frac{V_{th,typ}}{\max\!\left(I_{OUT},\ \frac{V_{OUT}}{V_{IN(min)}}I_{OUT}\right)\times 1.43}=\frac{50\,mV}{21.33\,A\times1.43}=1.639\,m\Omega$$

**Calc_mode = 0（= datasheet 式 21~23）**：

$$I_{IN,AVG(MAX)}=\frac{V_{OUT}I_{OUT}}{\eta\,V_{IN(min)}},\quad I_{L(PP)}=\frac{D\,V_{IN(min)}}{L f_{SW}}$$

$$R_{S,ideal}=\frac{V_{th,min}}{\left(I_{IN,AVG(MAX)}+\tfrac12 I_{L(PP)}\right)(1+\text{Margin})}=1.41\,m\Omega\ (\text{datasheet 值})$$

> ⚠️ 在目前的 Calc_mode = 1 下，**E28 的裕度（20 %）不會影響 E29**，裕度被寫死為 1.43；且使用的是 typ 門檻 50 mV 與不含效率的理想輸入電流。若要與 datasheet 式 23 一致需把 VM `L14` 改為 0。

### 3.2 E31 電感峰值過流限制（VM B156~B159）

$$I_{LPK,Limit}=\frac{V_{th,max}}{R_S}=\frac{57.5\,mV}{1.3\,m\Omega}=44.23\,A$$

（buck、boost 門檻相同，所以三個 VIN 點結果一樣，取 MAX。）這是元件（電感飽和、MOSFET）需承受的最壞峰值。

### 3.3 E32 正常滿載下最大電感峰值電流（VM B111~B115）

$$I_{L,pk}=\begin{cases}\dfrac{I_{OUT}}{\eta}+\dfrac{\Delta I_L}{2} & \text{buck}\\[1.5ex] I_{IN}+\dfrac{\Delta I_L}{2} & \text{boost}\end{cases}\qquad E32=\max(V_{IN,min},V_{IN,nom},V_{IN,max})$$

範例：25.10 A（VIN=6 V）/ 13.34 A / 14.64 A。
（VM B116 另有 ×(1+Margin) 的版本 30.1 A，未顯示。）

> 小瑕疵：B113（VIN_max 為 boost 時）用了 `Iin_Vinmin`，只有在 VIN(max) < VOUT 的設計才會出現影響。

### 3.4 E33 過流時 RS 功耗（VM B162）

$$P_{RS,OC}=\left(\frac{V_{th,max}}{R_S}\right)^2 R_S\times\frac{V_{OUT}}{V_{IN(max)}}=1.13\,W$$

> ⚠️ 與 datasheet 式 24 不同：datasheet 為 $\times\left(1-\frac{V_{OUT}}{V_{IN(max)}}\right)$（以 RCS=1 mΩ 算得 1.84 W）。工具乘的是 buck duty $D=V_{OUT}/V_{IN}$。兩者在 VIN(max)=2·VOUT 時才相等，選 RS 額定功率時建議兩者都算、取大者。
> VM 另有以 RMS 電流計算的滿載損耗 B95（0.415 W）與 OCP 時損耗 M156~M158（最大 0.76 W），未顯示在 DC。

### 3.5 E34/E35 斜率補償電阻（VM B164/B165；datasheet 式 12 / 式 25）

$$R_{SLOPE}=\frac{L_F}{R_S}\times 50\,\frac{MV}{A\cdot s}=\frac{1.8\,\mu H}{1.3\,m\Omega}\times 50\times10^6=69.23\,k\Omega\ \xrightarrow{E96}\ 69.8\,k\Omega$$

- E35 雖標示為「SLOPE Resistor」，在此版本中是**公式**（= E34），不是可輸入格；後續 `Rslope`（VM B166/B173）= E35×1000。
- VM B147 `Rs_in_range` 檢查 datasheet 式 13、14：`100 < Rs/Lf < 8000` 且 `Rs/Lf < fSW/(10·VOUT)`（本例 Rs/L = 722 Hz → 1 = OK）。

### 3.6 E36~E38 OCP 起始時的輸出電流（VM B151~B153）

$$I_{OUT,OCP}=\begin{cases}\dfrac{V_{th,typ}}{R_S}-\dfrac{\Delta I_L}{2} & \text{buck}\\[1.5ex]\left(\dfrac{V_{th,typ}}{R_S}-\dfrac{\Delta I_L}{2}\right)(1-D_{boost}) & \text{boost}\end{cases}$$

（使用 typ 門檻 50 mV，代表典型觸發點。）

---

## Step 4：Output Capacitor（輸出電容）

| DC 儲存格 | 項目 | 類型 | 引用來源 | 範例值 |
| --- | --- | --- | --- | --- |
| E41 | 輸出漣波規格 | 輸入 | → VM `B85` `Vripple1` (V) | 150 mVpp |
| E42 | 最小輸出電容 | 計算 | VM `B197`（← `Coutmin` B196） | 84.5 µF |
| E43 | COUT | 輸入 | → VM `B198` `Cout` | 130 µF |
| E44 | 最大允許 ESR | 計算 | VM `B201`（← `CoutEsrMax` B199） | 6.56 mΩ |
| E45 | 輸出電容 ESR | 輸入 | → VM `B202` `CoutEsr` | 2 mΩ |
| E46 | 實際輸出漣波 | 計算 | `=Vripple2×1000` → VM `B203` | 113.5 mVpp |
| E47 | 輸出電容 RMS 電流 | 計算 | VM `B192` = MAX(B189:B191) | 10.33 A |

所有 Step 4 公式都以 **VIN(min)** 判斷模式（`IF(VIN_min>Vout, buck, boost)`），因為 boost 在 VIN(min) 時輸出電容最嚴苛。

### 4.1 E42 最小輸出電容（VM B196/B197）

$$C_{OUT,min}=\begin{cases}\dfrac{\Delta I_{L,VIN(min)}}{8\,f_{SW}\,\Delta V_{OUT}} & \text{buck}\\[2ex]\dfrac{I_{OUT}\,D_{boost,VIN(min)}}{\Delta V_{OUT}\,f_{SW}} & \text{boost}\end{cases}=\frac{8\times0.6306}{0.15\times398k}=84.5\,\mu F$$

boost 式對應 datasheet 式 28（ΔV_RIPPLE(COUT)）反解。顯示下限為 0.1 µF（VM H197）。

### 4.2 E44 最大允許 ESR（VM B199/B201）

$$ESR_{max}=\begin{cases}\sqrt{\left(\dfrac{\Delta V_{OUT}}{\Delta I_L}\right)^2-\left(\dfrac{1}{8f_{SW}C_{OUT}}\right)^2} & \text{buck}\\[2ex]\dfrac{\Delta V_{OUT}-\dfrac{I_{OUT}D}{f_{SW}C_{OUT}}}{I_{OUT}} & \text{boost}\end{cases}=6.56\,m\Omega$$

結果 > 300 mΩ、< 0 或錯誤時顯示 "N/A"。

### 4.3 E46 實際輸出漣波（VM B203）

$$\Delta V_{OUT}=\begin{cases}\Delta I_{L,VIN(max)}\sqrt{ESR^2+\left(\dfrac{1}{8f_{SW}C_{OUT}}\right)^2} & \text{buck}\\[2ex] ESR\cdot I_{OUT}+\dfrac{I_{OUT}D}{f_{SW}C_{OUT}} & \text{boost}\end{cases}=16+97.5=113.5\,mV$$

> 與 datasheet 對照：式 28 電容漣波 = 96 mV（理想 D），工具 97.5 mV（含損耗 D）一致；
> 但 **ESR 項不同**：datasheet 式 27 用 $\frac{I_{OUT}V_{OUT}}{V_{IN(min)}}\times ESR$（= 輸入平均電流 21.3 A → 42.6 mV），工具只用 $I_{OUT}\times ESR$（16 mV）。工具的 ESR 漣波較樂觀。

### 4.4 E47 輸出電容 RMS 電流（VM B189~B192；datasheet 式 26）

$$I_{COUT,RMS}=\begin{cases}\Delta I_L/\sqrt{12} & \text{buck}\\ I_{OUT}\sqrt{\dfrac{V_{OUT}-V_{IN}}{V_{IN}}}=I_{OUT}\sqrt{\dfrac{V_{OUT}}{V_{IN}}-1} & \text{boost}\end{cases}$$

取 VIN(min/nom/max) 三點最大值 = 10.33 A（datasheet 10.3 A）。

---

## Step 5：Input Capacitor（輸入電容）

| DC 儲存格 | 項目 | 類型 | 引用來源 | 範例值 |
| --- | --- | --- | --- | --- |
| E50 | 輸入漣波規格 | 輸入 | → VM `B212` `Vinripple1` | 180 mVpp |
| E51 | 最小輸入電容 | 計算 | VM `B217`（← `Cinmin` B216） | 27.9 µF |
| E52 | CIN | 輸入 | → VM `B218` `Cin` | 100 µF |
| E53 | 最大允許 ESR | 計算 | VM `B222`（← `CinEsrMax` B221） | 9.13 mΩ |
| E54 | 輸入電容 ESR | 輸入 | → VM `B223` `RCinEsr` | 2 mΩ |
| E55 | 實際輸入漣波 | 計算 | `=Vinripple2×1000` → VM `B227` | 78.7 mVpp |
| E56 | 輸入電容 RMS 電流 | 計算 | VM `B210` = MAX(B207:B209) | 3.98 A |

Step 5 以 **VIN(max)** 為主（buck 時輸入電容最嚴苛）。若 VIN(max) ≥ 2·VOUT，buck 會經過 D = 0.5 的最壞點，工具改用 D = 0.5（VM B215/B220/B226）；本例 36 V ≥ 32 V，所以採用 D = 0.5 分支。

### 5.1 E51 最小輸入電容（VM B214~B217）

$$C_{IN,min}=\begin{cases}\dfrac{I_{OUT}\,D(1-D)}{\Delta V_{IN}\,f_{SW}} & \text{buck}\ (D=D_{VIN(max)}\ \text{或}\ 0.5)\\[2ex]\dfrac{\Delta I_{L,VIN(max)}}{8f_{SW}\Delta V_{IN}} & \text{boost}\end{cases}=\frac{8\times0.25}{0.18\times398k}=27.9\,\mu F$$

顯示下限 10 µF（B217）。datasheet 無對應公式（§9.2.2.7 只給 RMS）。

### 5.2 E53 最大允許 ESR（VM B219~B222）

$$ESR_{IN,max}=\frac{\Delta V_{IN}-\dfrac{I_{OUT}D(1-D)}{C_{IN}f_{SW}}}{I_{OUT}+\Delta I_L/2}\quad(\text{buck})$$

boost 分支同 Step 4 的 buck 型式（√ 形式）。

### 5.3 E55 實際輸入漣波（VM B225~B227）

$$\Delta V_{IN}=\frac{I_{OUT}D(1-D)}{C_{IN}f_{SW}}+\left(I_{OUT}+\frac{\Delta I_{L,VIN(max)}}{2}\right)ESR_{IN}=50.3+28.4=78.7\,mV$$

### 5.4 E56 輸入電容 RMS 電流（VM B207~B210；datasheet 式 29）

$$I_{CIN,RMS}=\begin{cases}I_{OUT}\sqrt{D(1-D)} & \text{buck}\\ \Delta I_L/\sqrt{12} & \text{boost}\end{cases}$$

三點取最大 = 3.98 A（VIN=36 V、D=0.445）。datasheet 以 D=0.5 給出 IOUT/2 = 4.0 A。

> 小瑕疵：B207（VIN(min) 為 buck 時）寫成 `Don_Vinmin_2*(1-Don_Vinmax_2)`，混用了兩個 duty；只在 VIN(min) > VOUT 的純 buck 設計才有影響。

---

## Step 6：Soft-start, UVLO, CFG（軟啟動、UVLO、設定腳）

| DC 儲存格 | 項目 | 類型 | 引用來源 | 範例值 |
| --- | --- | --- | --- | --- |
| E59 | 建議軟啟動時間 | 計算 | `=E43×1e-6×E9/E10×1000` | 0.26 ms |
| E60 | 軟啟動時間 tSS | 輸入 | → VM `B232` `Tss` | 1.8 ms |
| E61 | CSS | 計算 | `=Css_u×1e9` → VM `B235` → Std Value `B4` | 18 nF |
| E63 | CFG 設定編號 | 計算 | `=CFG` → VM `O10` → `'CFG Setting'!G28` | 7 |
| E64 | RCFG | 計算 | `=INDEX('CFG Setting'!B4:B19, CFG)` | 5.1 kΩ |
| E65 | UVLO On 電壓 | 輸入 | → VM `B239` | 4.6 V |
| E66 | UVLO Off 電壓 | 輸入 | → VM `B240` | 4.025 V |
| E67 | RUV1（上） | 計算 | VM `G241` ← Std Value `O16` ← `Ruvlo1` (B241) | 32.4 kΩ |
| E68 | RUV2（下） | 計算 | VM `G242` ← Std Value `O17` ← `Ruvlo2` (B242) | 13.3 kΩ |
| D70 / E70 | 監測最大電流 / 平均電流限制 | 標籤由 `IMONLIM_select` 決定；E70 輸入 → VM `B268` | 5 A |
| E71 | IMONOUT 最大電壓 | `CHOOSE(IMONLIM_select+1,"","1.5","5","0")` → VM `B273` | "5" |
| E72 | 電流感測分流電阻 RCS(out/in) | `=Rshunt` → VM `B270` | 10 mΩ |
| D73 / E73 | RIMON 或 SYNC 腳啟動電位 | `CHOOSE(IMONLIM_select+1,"", B274/1000, "High", "Low")` | "High" |

### 6.1 E59 建議軟啟動時間

$$t_{SS,rec}=\frac{C_{OUT}V_{OUT}}{I_{OUT}}$$

即「以滿載電流對 COUT 充電到 VOUT」所需時間；軟啟動若比它快，充電電流會超過 IOUT。此為工具的經驗式，datasheet 無對應公式（§8.3.11 只說明軟啟動可限制大輸出電容的湧浪電流）。

### 6.2 E61 CSS（VM B234；datasheet 式 15 / 式 31）

$$C_{SS}=\frac{I_{SS}\,t_{SS}}{V_{REF}}=\frac{10\,\mu A\times1.8\,ms}{1\,V}=18\,nF\ \xrightarrow{E12}\ 18\,nF$$

### 6.3 E63/E64 CFG 腳位設定（CFG Setting 表；datasheet §8.3.16 Table 8-1）

使用者在 `CFG Setting!C27:F27` 選擇（DC 上的下拉/按鈕會寫入此處）：

| 功能 | 儲存格 | 選項 | 位元權重 |
| --- | --- | --- | --- |
| DRSS 展頻 | C27 | 1=off / 2=on | 2⁰ = 1 |
| SCP Hiccup | D27 | 1=off / 2=on | 2¹ = 2 |
| PSM 進入門檻 | E27 | 1=10 % / 2=15 % | 2³ = 8 |
| IMON / ILIMIT | F27 | 1=off / 2=IMON / 3=ILIMIT 正向 / 4=ILIMIT 負向 | 2² = 4（F27>2 才算 Current Limit 致能） |

$$CFG=1+(C27-1)\cdot1+(D27-1)\cdot2+(F27>2)\cdot4+(E27-1)\cdot8$$

本例：DRSS off、Hiccup on、PSM 10 %、ILIMIT 正向 → CFG = 1+0+2+4+0 = **7** → R_CFG = **5.1 kΩ**，與 Table 8-1 #7（DRSS 停用、SCP 啟用、10 %、Current Limit 啟用）一致。

另外 `IMONLIM_select = F27 − 1`（0: off / 1: IMON / 2: ILIM+ / 3: ILIM−），`With_RCS = F27`。

> 注意：datasheet §9.2.2.13 範例用 R_CFG = 13.3 kΩ（#11：SCP 啟用、PSM 15 %、無 Current Limit），與工具目前的設定不同，屬設計選擇差異。

### 6.4 E67/E68 UVLO 分壓（VM B241/B242；datasheet 式 1 / 式 30）

Datasheet 式 1：

$$V_{IN,ON}=V_{T+(UVLO)}\left(1+\frac{R_{UV1}}{R_{UV2}}\right)+R_{UV1}\,I_{UVLO},\qquad V_{IN,OFF}=V_{T-(UVLO)}\left(1+\frac{R_{UV1}}{R_{UV2}}\right)$$

工具聯立解出：

$$R_{UV1}=\frac{1}{I_{UVLO}}\left(V_{IN,ON}-V_{T+}\frac{V_{IN,OFF}}{V_{T-}}\right)=32.76\,k\Omega\ \xrightarrow{E96}\ 32.4\,k\Omega$$

$$R_{UV2}=\frac{R_{UV1}}{V_{IN,OFF}/V_{T-}-1}=13.43\,k\Omega\ \xrightarrow{E96}\ 13.3\,k\Omega$$

（I_UVLO 以 mA 代入，結果單位即為 kΩ。）若 RUV1 < 0（ON/OFF 遲滯設得太小）顯示 "N/A"。
標準值回推的實際門檻（VM B244/B245，未顯示）：ON = 4.59 V、OFF = 4.02 V。

> ⚠️ 工具用的門檻為 V_T+ = 1.28 V（max）、V_T− = 1.17 V（min）、I_UVLO = 6 µA（max），而非 datasheet typ 值 1.25 V / 1.2 V / 5 µA；以 typ 值代入同一組電阻，ON/OFF 會略有偏移。

### 6.5 E70~E73 電流監測 / 平均電流限制（VM B266~B274；datasheet §8.3.13 式 17）

- **分流電阻**（VM B270）：

$$R_{SHUNT}=\frac{\Delta V_{ISNS}}{I_{set}}=\frac{50\,mV}{5\,A}=10\,m\Omega$$

  ΔV_ISNS = 50 mV 為 datasheet 的電流限制門檻（§6.5 ΔV(ISNS) 49/50/51.7 mV）。IMON 模式也用同一式（把 E70 當作滿刻度 50 mV 的電流）。

- **IMON 模式（IMONLIM_select = 1）**：E71 固定 1.5 V，E73 為

$$R_{IMON}=\frac{V_{IMONOUT,max}}{I_{max}\,R_{SHUNT}\,gm_{IMON}}$$

  即 datasheet 式 17 $V_{IMONOUT}=(V_{ISNSP}-V_{ISNSN})\,gm_{IMON}\,R_{IMONOUT}$ 的反解。

- **ILIM 模式（IMONLIM_select = 2/3）**：E73 顯示「SYNC pin during startup = High / Low」，對應 datasheet：EN/UVLO 上升到軟啟動開始期間 SYNC 為 High → 正向電流限制、Low → 負向電流限制。
- 分流電阻損耗 $R_{SHUNT}I_{OUT}^2$（0.64 W）計入效率頁 `Iout Shunt Loss`。

---

## Step 7：Compensation Design（電壓迴路補償）

| DC 儲存格 | 項目 | 類型 | 引用來源 | 範例值 |
| --- | --- | --- | --- | --- |
| E78 | 負載極點頻率 | 計算 | `=Fp_dominant` → VM `B368` | 1224 Hz |
| E79 | COUT ESR 零點 | 計算 | `=Fesr/1000` → VM `B372` | 612 kHz |
| E80 | Boost RHP 零點 @ VIN(min) | 計算 | `=IF(E6<E9, Frhp_Vinmin/1000, "---")` → VM `B374` | 24.1 kHz |
| E81 | Boost RHP 零點 @ VIN(nom) | 計算 | `=IF(E7<E9, Frhp/1000, "---")` → VM `B373` | 98.8 kHz |
| E82 | 建議交越頻率 | 計算 | `=MIN(Fsw/1000×0.05, E80/4, IF(E81<>"---", E81/4, E80/4))` | 6.03 kHz |
| E83 | 期望交越頻率 | 輸入 | → VM `B380` `Fcross` | 4 kHz |
| E84 | 誤差放大器極點 | 計算 | `=Fp1_u` → VM `B370` | 0.284 Hz |
| E85 | RFBT | 輸入 | → VM `B396` `Rfb1_` | 71.5 kΩ |
| E86 | RFBB | 計算 | `=IF(Vout=Vref,"OPEN", Rfb2_u/1000)` → VM `B398` → Std Value `O10` | 4.75 kΩ |
| E87 | 實際 VOUT | 計算 | `=1/E86×(E85+E86)`（隱含 ×Vref=1 V） | 16.05 V |

模式判斷：Step 7 以 **VIN(nom)** 判斷 buck/boost（`IF(VIN_nom>Vout, buck, boost)`），本例 12 V < 16 V → boost。

### 7.1 功率級極零點

$$f_{p,load}=\begin{cases}\dfrac{1}{2\pi R_{OUT}C_{OUT}}=612\,Hz & \text{buck（datasheet 式 43）}\\[2ex]\dfrac{2}{2\pi R_{OUT}C_{OUT}}=1224\,Hz & \text{boost（datasheet 式 40 / 式 11）}\end{cases}$$

$$f_{ESR}=\frac{1}{2\pi C_{OUT}ESR}=612\,kHz\quad(\text{datasheet 式 41})$$

> datasheet 式 41 寫「= 61.2 kHz」，但以其 2 mΩ、130 µF 代入實為 612 kHz，datasheet 數值疑似差 10 倍；工具算法正確。

$$f_{RHP}=\frac{(1-D)^2R_{OUT}}{2\pi L_F}\quad(\text{datasheet 式 42 / 式 7})$$

- E80 用 D(VIN(min)) = 0.6306 → 24.1 kHz（datasheet 用理想 D=0.625 → 24.87 kHz）
- E81 用 D(VIN(nom)) = 0.2525 → 98.8 kHz

### 7.2 E82 建議交越頻率

$$f_{c,sug}=\min\left(\frac{f_{SW}}{20},\ \frac{f_{RHP,VIN(min)}}{4},\ \frac{f_{RHP,VIN(nom)}}{4}\right)$$

> Datasheet：式 8 $f_{BW}<f_{RHPZ}/3$、式 9 $f_{BW}<\frac{1}{10}(1-D_{max})f_{SW}$，§9.2.2.12 則說取 $f_{RHP}/3$ 與 $f_{SW}/20$ 的較小者。工具用 **/4**（較保守），且沒有實作式 9。

### 7.3 E86/E87 回授分壓（VM B397/B398；datasheet 式 19）

$$R_{FBB}=\frac{V_{REF}\,R_{FBT}}{V_{OUT}-V_{REF}}=\frac{71.5k}{15}=4.767\,k\Omega\ \xrightarrow{E96}\ 4.75\,k\Omega$$

$$V_{OUT,eff}=V_{REF}\frac{R_{FBT}+R_{FBB}}{R_{FBB}}=16.05\,V$$

（datasheet 文字選 4.7 kΩ，Table 9-2 E48 欄為 4.87 kΩ；工具採 E96 → 4.75 kΩ。）VOUT = VREF 時 RFBB = OPEN。

### 7.4 E84 誤差放大器極點（VM B370）

$$f_{p,EA}=\frac{1}{2\pi(R_{O,EA}+R_{COMP})C_{COMP}}=\frac{1}{2\pi(10M+1.5k)\cdot56n}=0.284\,Hz$$

（使用 Step 7.1 選定後的標準值。）

---

## Step 7.1：Compensation Components（補償元件）

| DC 儲存格 | 項目 | 計算值（B 欄） | 標準值（C 欄 ← Std Value） | 選用（E 欄 = C 欄） | 頻率（G 欄） |
| --- | --- | --- | --- | --- | --- |
| 第 91 列 | RCOMP | `=Rcea1/1000` (VM B381) → 1.514 kΩ | `O27` (= E96(`Rcea1`)/1000) → 1.5 kΩ | 1.5 kΩ → VM `B386` `Rcea1_u` | G91 `=Fp1_u` 0.284 Hz (FPEA) |
| 第 92 列 | CCOMP | `=Cc1ea×1e9` (VM B382) → 57.24 nF | `O28` (E12) → 56 nF | 56 nF → VM `B387` `Cc1ea_u` | G92 `=Fz1_u/1000` 1.89 kHz (FZCOMP) |
| 第 93 列 | CHF | `=MAX(0, Cc2ea×1e12)` (VM B383) → 2621 pF | `O29` (E12) → 2700 pF | 2700 pF → VM `B388` `Cc2ea_u` | G93 `=Fp2_u/1000` 39.2 kHz (FPCOMP) |

### 7.1.1 RCOMP（VM B381；datasheet 式 3 / 式 46）

定義 $R_i=A_{CS}R_S=10\times R_S=13\,m\Omega$、$A_{fb}=V_{REF}/V_{OUT}=0.0625$：

$$R_{COMP}=\begin{cases}\dfrac{2\pi f_c\,C_{OUT}\,R_i}{gm_{EA}\,A_{fb}} & \text{buck}\\[2.5ex]\dfrac{2\pi f_c\,C_{OUT}\,R_i}{gm_{EA}\,A_{fb}\,(1-D_{nom})\sqrt{1+(f_c/f_{RHP,nom})^2}} & \text{boost}\end{cases}=1.514\,k\Omega$$

這與 datasheet 式 46 形式完全相同（$\frac{1}{A_{fb}}=\frac{R_{FB1}+R_{FB2}}{R_{FB2}}$）。

> ⚠️ 差異：datasheet 式 46 使用 **D_MAX（VIN(min) 時的 duty）與 f_RHP @ VIN(min)**；工具使用 **VIN(nom) 的 duty 與 f_RHP**（`Don_Vinnom_2`、`PS_RHPZ = Frhp`）。本例若用 D_MAX=0.6306、f_RHP=24.1 kHz，RCOMP 會變成約 3.0 kΩ。即工具是以 VIN(nom) 為補償設計點；datasheet 則以最嚴苛的 VIN(min) 設計（例題算出 1.9 kΩ 是以 5 kHz 頻寬）。

### 7.1.2 CCOMP（VM B382；datasheet 式 5、10、47）

$$C_{COMP}=\frac{1}{2\pi R_{COMP}\cdot(1.5\,f_{p,dominant})}=57.2\,nF$$

即補償零點放在主極點的 1.5 倍（datasheet：$f_{CZ}=1.5\times f_{pole,boost}$）。

### 7.1.3 CHF（VM B383/B384；datasheet 式 6、48）

$$C_{HF}=\frac{1}{2\pi R_{COMP}\cdot\max(7f_c,10f_c)}-C_{bw},\qquad C_{bw}=\frac{gm_{EA}}{2\pi\,GBW}=7.35\,pF$$

高頻極點放在 10·fc（datasheet：7~10 倍 f_bw，式 6 用 10 倍），並扣除誤差放大器頻寬等效電容 Cbw（datasheet 無此項）。

### 7.1.4 實際極零點（G 欄，使用標準值）

$$f_{Z,COMP}=\frac{1}{2\pi R_{COMP}C_{COMP}}=1.89\,kHz,\qquad f_{P,COMP}=\frac{1}{2\pi R_{COMP}(C_{HF}+C_{bw})}=39.2\,kHz$$

### 7.1.5 Bode 圖（Stability Calculations 表）

- 迴路增益：$A_{dc}=A_{EA}\cdot A_{fb}\cdot A_m$，$A_{EA}=gm\cdot R_{O,EA}=6000$。
- 調變器增益（VM B355~B357）：

$$A_{m,boost}=\frac{R_{OUT}(1-D)}{2R_i},\qquad A_{m,buck}=\frac{R_O/R_i}{1+R_O(k-0.5)/(f_{SW}L)}$$

- 另含取樣雙極點 $f_{SW}/2$（Q = $1/[\pi(k-0.5)]$，k 由斜率補償 `mc` 計算）、ESR 零點、RHP 零點、EA 極零點。
- 掃頻 1 Hz 起共約 600 點，以「增益穿越 0 dB」找交越頻率 `f_cross_Bode`（= `Stability Calculations!Q4`，本例 3.98 kHz）及相位裕度（Q5，本例 71.7°）。這兩個結果會被 Step 7.2 使用。

---

## Step 7.2：Current Limiter Compensation Design（平均電流限制迴路補償）

| DC 儲存格 | 項目 | 計算值（K 欄） | 標準值（L 欄） | 選用（M 欄，輸入） |
| --- | --- | --- | --- | --- |
| 第 91 列 | CIMON | `=C_IL_recommend×1e9` (VM B292) → 1.40 nF | `'Variable Management'!F292` ← Std Value `O23` → 1.5 nF | 2.2 nF → VM `B289` `Cc1IL` |
| 第 92 列 | RIMON | `=R_IL_recommend/1000` (VM B293) → 0.477 kΩ | `F293` ← Std Value `O24` (E96) → 0.475 kΩ | 0.7 kΩ → VM `B288` `RcIL` |

### 7.2.1 公式

$$C_{IL}=\frac{gm_{ILIM}}{10^{\left[20\log_{10}(3f_{c,Bode})+G_{plant}\right]/20}}=\frac{gm_{ILIM}}{3f_{c,Bode}\cdot 10^{G_{plant}/20}}$$

- $gm_{ILIM}$ = DeviceParameter B12 = 200 µS（datasheet gm(ILIM)）
- $f_{c,Bode}$ = `Stability Calculations!Q4`（電壓迴路的實際交越頻率，本例 3.98 kHz）
- $G_{plant}$ = `'IL Stability Calculations'!BE13`：電流迴路功率級（控制→分流電阻電流）在最低掃頻點的增益 dB（本例 21.58 dB）

$$R_{IL}=\frac{1}{2\pi\,C_{IL}\cdot 20\cdot 3f_{c,Bode}}$$

意義：讓電流限制迴路的頻寬約為電壓迴路交越頻率的 **3 倍**，補償零點再放在其 20 倍處。對應 datasheet §8.3.13：「For most applications, a compensation bandwidth with a factor of 3× to 5× faster than the compensation of the output voltage loop has shown good results」。Datasheet 未提供 RIMON/CIMON 的計算式，此處為工具的經驗公式。

M 欄（使用者選用值）會回饋到 `IL Stability Calculations`，用來畫電流限制迴路 Bode 圖（本例交越 2.29 kHz、相位裕度 101.9°）。

---

## Step 8：Efficiency（效率 / 功率損耗分析）

此步驟在 DC 上只有**輸入格**，數值被搬到 VM B429~B465，再由 `Efficiency vs VIN Calculations`（VIN 從 VIN(min) 到 VIN(max) 分 200 步）與 `Efficiency vs Load Calculations` 計算損耗與效率曲線。

### 8.1 輸入對應

| DC 儲存格 | 項目 | Q1 / Q3（E 欄，高側） | Q2 / Q4（F 欄，低側） | 單位換算 |
| --- | --- | --- | --- | --- |
| 100 / 113 | RDS(on) | `Rdson1` B430 / `Rdson3` B447 | `Rdson2` D430 / `Rdson4` D447 | mΩ→Ω |
| 101 / 114 | QG | `Qg_Q1` / `Qg_Q3` | `Qg_Q2` / `Qg_Q4` | nC→C |
| 102 / 115 | QGD | `Qgd_Q1` / `Qgd_Q3` | `Qgd_Q2` / `Qgd_Q4` | nC→C |
| 103 / 116 | QGS | `Qgs_Q1` / `Qgs_Q3` | `Qgs_Q2` / `Qgs_Q4` | nC→C |
| 104 / 117 | RG | `Rgate1` / `Rgate3` | `Rgate2` / `Rgate4` | Ω |
| 105 / 118 | gFS | `Gfs_Q1` / `Gfs_Q3` | `Gfs_Q2` / `Gfs_Q4` | S |
| 106 / 119 | VTH | `Vth_Q1` / `Vth_Q3` | `Vth_Q2` / `Vth_Q4` | V |
| 107 / 120 | VBD | `Vsd_Q1` / `Vsd_Q3` | `Vsd_Q2` / `Vsd_Q4` | V |
| 108 / 121 | QRR | E108 空 → `Qrr_Q1`=0；**E121=56 → `Qrr_Q3`** | **F108=45 → `Qrr_Q2`**；F121 空 → `Qrr_Q4`=0 | nC→C |
| 109 / 122 | θJA | `ThetaJA_Q1` / `ThetaJA_Q3` | `ThetaJA_Q2` / `ThetaJA_Q4` | °C/W |
| 125 / 126 | 外部蕭特基 VF / QRR | `VfdSky_Q2` / `QrrSky_Q2`（Buck leg） | `VfdSky_Q3` / `QrrSky_Q3`（Boost leg） | |

VM 中其他寫死的效率參數：Ta = 25 °C、電感/RDS(on) 溫度係數 4000 ppm、θCA = 20 °C/W、VCC = 7.5 V、閘極驅動 RDS(on) 上拉 1.5 Ω / 下拉 1 Ω、死區 50 ns、IQ = 3 mA、trr = 40 ns。

### 8.2 主要損耗公式（Efficiency vs VIN Calculations）

- **時序**：依 `Timings` 表的轉換電壓決定每個 VIN 點為 BUCK / BOOST / BUCK-BST，計算三相位時間 t1(Q1-Q4)、t2(Q1-Q3)、t3(Q2-Q4) 與各相位電感電流 I0/I1/I2，再求各元件 RMS 電流。
- **MOSFET 平台電壓與切換時間**（VM B438、B471~B472）：

$$V_{sp}=V_{TH}+\frac{I}{g_{FS}},\quad t_r=\frac{(Q_{GD}+Q_{GS}/2)(R_{DRV,up}+R_G)}{V_{CC}-V_{sp}},\quad t_f=\frac{(Q_{GD}+Q_{GS}/2)(R_{DRV,dn}+R_G)}{V_{sp}}$$

- **導通損耗**：$P=R_{DS(on)}I_{rms}^2$，並以 $(1+TC(T_a-25))/(1-P\cdot TC\cdot\theta_{JA})$ 作溫升修正。對應 datasheet 式 32、33、35~37、39。
- **切換損耗**：buck 控制管 $\frac12 V_{IN}I_{OUT}(t_r+t_f)f_{SW}$（datasheet 式 34）；boost 控制管 $\frac12 V_{OUT}I_{IN}(t_r+t_f)f_{SW}$（datasheet 式 38）。
- **死區（體二極體）損耗**：$2\,t_{dead}f_{SW}V_{BD}I$。
- **反向恢復損耗**：buck 用 $Q_{RR,Q2}f_{SW}V_{IN}$、boost 用 $Q_{RR,Q3}f_{SW}V_{OUT}$。
- **電感**：DCR 損耗（含溫升）+ 核心損耗經驗式 $7.759\times10^{-3}f_{SW}^{0.9458}(0.00787\,\Delta I_L)^{2.304}$。
- **電容**：$ESR\cdot I_{rms}^2$（Cin、Cout）。
- **感測電阻**：RS（依模式的 RMS）＋ RSHUNT·IOUT²。
- **IC 損耗**：$I_Q V$ + $\sum Q_G\,f_{SW}\,V$（VOUT > 8 V 時以 VOUT 供電）。
- **效率**：$\eta=P_{OUT}/(P_{OUT}+\sum P_{loss})$；本例 VIN=6 V 為 90.3 %。

---

## 附錄 A：Design Converter 引用的其他工作表一覽

| 被引用工作表 | 儲存格 / 具名範圍 | 被 DC 哪一項使用 |
| --- | --- | --- |
| Standard Value Calculator | `O4`（RT E96） | E12 |
| | `O14`（C_CS E12，經 DeviceParameter N10） | E25 |
| | `O15`（RSLOPE E96，經 VM B165） | E34、E35 |
| | `B4`（CSS E12，經 VM B235） | E61 |
| | `O16`/`O17`（RUV1/RUV2 E96，經 VM E241/E242→G241/G242） | E67、E68 |
| | `O10`（RFBB E96，經 VM B398） | E86 |
| | `O27`/`O28`/`O29`（RCOMP/CCOMP/CHF） | C91~C93 |
| | `O23`/`O24`（CIMON/RIMON，經 VM F292/F293） | L91、L92 |
| | `O21`（`Rshunt_std`） | （僅定義，DC 未用） |
| DeviceParameter | `N8` `R_CS` | E24 |
| | `N9`/`N10` `C_CS`/`C_CS_ideal`，`N6` tON,min 查表 | E25 |
| | `B16~B18` CS 門檻 | E29~E38 |
| | `B11` gm_IMON、`B12` gm_IL | E73、K91/K92 |
| CFG Setting | `G28` CFG 編號、`B4:B19` R_CFG 表、`F27` IMON/ILIM 選擇 | E63、E64、D70~E73 |
| Timings | `F2~F5` 模式轉換電壓（`VIN_edgebuckboost` 等） | Step 2/3 模式標籤、效率頁 |
| Stability Calculations | `Q4` `f_cross_Bode` | K91、K92 |
| IL Stability Calculations | `BE13` 電流迴路 plant 增益 | K91、K92 |
| Licenses | `B5` 版本號 | J4 |

## 附錄 B：標準值取整方法（Standard Value Calculator）

- **電阻（O 欄）**：E96 系列，公式為
  `INT(0.5 + 100·10^(ROUND(96·frac(log10 R))/96)) × 10^(INT(log10 R) − 2)`，並比較上下兩個 E96 值取最接近者。
- **電容（B~I 欄）**：在 A 欄（0.47、0.56、0.68、0.82、1、1.2、1.5 … pF，E12 系列）逐列比較，取最接近值後加總得到結果（B4、C4、D4、G4、I4）。

## 附錄 C：工具與 Datasheet 差異重點整理

| # | 項目 | 工具 | Datasheet | 影響 |
| --- | --- | --- | --- | --- |
| 1 | 建議 RS（E29） | Calc_mode=1：50 mV / (IOUT·VOUT/VIN(min) × 1.43)，**E28 裕度無作用** | 式 23：42.5 mV / ((I_IN,avg + ½ΔI)·1.2) | 1.64 vs 1.41 mΩ |
| 2 | 過流 RS 功耗（E33） | × VOUT/VIN(max) | 式 24：× (1 − VOUT/VIN(max)) | 本例工具偏低 |
| 3 | 輸出漣波 ESR 項（E44/E46） | IOUT × ESR | 式 27：IOUT·VOUT/VIN(min) × ESR | 工具偏樂觀（16 vs 42.6 mV） |
| 4 | RCOMP（B91） | 用 VIN(nom) 的 D 與 f_RHP | 式 46：用 D_MAX、f_RHP @ VIN(min) | 本例 1.51 vs ~3.0 kΩ（同 4 kHz） |
| 5 | 建議交越（E82） | min(fSW/20, f_RHP/4) | min(fSW/20, f_RHP/3)；式 9 另有 (1−D)·fSW/10 | 工具較保守 |
| 6 | CHF | 扣除 Cbw = gm/(2π·GBW) | 無 | 小幅差異 |
| 7 | UVLO 門檻 | 1.28 V / 1.17 V / 6 µA（極限值） | typ 1.25 V / 1.2 V / 5 µA | 實際門檻略偏移 |
| 8 | 建議 L（E15） | Lopt 再以 Lmin/Lmax（式 13、14 + 35·VOUT·RS/fSW）截斷 | 式 20 只給 Lopt = 2.21 µH | 本例 1.83 µH |
| 9 | ΔIL、D | 含 DCR 與 RS 損耗的 duty | 理想 duty | 約 1 % 差 |
| 10 | 推薦 tSS（E59） | COUT·VOUT/IOUT | 無 | — |
| 11 | RIMON/CIMON（Step 7.2） | 3×fc 經驗式 | 只說 3~5 倍電壓迴路頻寬 | — |
| 12 | Datasheet 自身 | — | 式 41 f_ESR 寫 61.2 kHz（應為 612 kHz）；§9.2.2.2 寫 78.7 kΩ（式 18 算出約 75 kΩ） | 工具算法正確 |
