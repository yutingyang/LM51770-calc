# LM51770 Buck-Boost Quickstart Tool：16 個分頁逐頁分析

> 分析對象：`LM51770_Buck-Boost_Quickstart_Tool_V1_0_1(3).xlsm`（Version 1.0.1）
> 方法：用 openpyxl 讀出每一格公式與快取值，用 oletools 讀 VBA，直接解析 xlsx 的 XML 取得表單控制項、圖表資料來源與條件格式。
> 範例數值都是檔案存檔時的快取值：VIN = 6 / 12 / 36 V、VOUT = 16 V、IOUT = 8 A、fSW = 398 kHz、L = 1.8 µH、RS = 1.3 mΩ。
> Design Converter 的設計步驟與 datasheet 公式對照，見同目錄的 `LM51770_Design_Converter_分析.md`。本文重點是**每個分頁本身**：它做什麼、引用了哪些格、用了哪些變數。
>
> **配套檔案**
> * `LM51770_xlsm_全儲存格清單.xlsx` / `.csv`：全部 79 400 個非空格，每格一列（分頁、位置、公式或常數、快取值、A 欄標籤、具名範圍、R1C1 公式樣式、儲存格註解）。本文對大分頁只寫公式樣式，要查任一格的原始內容請用這份清單篩選。
> * `xlsm_images/`：從 xlsm 抽出的 26 張圖（EMF 已轉成 PNG），分析見 §18。

---

## 0. 總覽

### 0.1 分頁清單

| # | 分頁 | 範圍 | 非空格 | 公式數 | 圖表 | 角色 |
|---|---|---|---|---|---|---|
| 1 | Design Converter (DC) | A1:AI475 | 360 | 73 | 0（有連結圖片） | **使用者介面**：輸入與結果顯示 |
| 2 | Licenses | A1:AM76 | 11 | 0 | 0 | 版本號、版本歷史、TI 免責聲明 |
| 3 | Variable Management (VM) | A1:R559 | 1517 | 488 | 0 | **核心計算**，大部分具名範圍都在這裡 |
| 4 | DeviceParameter (DP) | A1:Q107 | 273 | 12 | 0 | 依 IC 型號選參數、tON(min) 對 fSW 查表 |
| 5 | CFG Setting | A1:R151 | 296 | 155 | 0 | CFG 腳位 16 種組合表，算 RCFG |
| 6 | Stability Calculations (STAB) | A1:AN687 | 19 636 | 18 912 | 3 | 電壓迴路 Bode 圖（純量／漸近模型） |
| 7 | IL Stability Calculations (ILSTAB) | A1:BO687 | 33 758 | 32 997 | 3 | 平均電流限制迴路 Bode 圖（複數模型） |
| 8 | Schematic Management | A1:B9 | 1 | 1 | 0 | 選擇要顯示哪一張電路圖 |
| 9 | Efficiency Plot Details | A1:C7 | 0 | 0 | 14（全部 #REF!） | 舊版遺留圖表 |
| 10 | Efficiency Plot Management | A1:A9 | 2 | 2 | 4 | 選擇 DC 上要顯示哪張效率圖／Bode 圖 |
| 11 | Timings | A1:AG96 | 296 | 136 | 1 | Buck↔Buck-Boost↔Boost 轉換電壓、三相時間 |
| 12 | Efficiency vs VIN Calculations | A1:CN224 | 14 122 | 13 410 | 4 | 效率／損耗對 VIN 掃描（201 點） |
| 13 | Efficiency vs Load Calculations | A1:CN209 | 7 424 | 6 811 | 3 | 效率／損耗對負載掃描（0–100 %，101 點） |
| 14 | Standard Value Calculator (SVC) | A1:P159 | 1 641 | 1 565 | 0 | 理想值取 E96 電阻／E12 電容標準值 |
| 15 | Snubber | A1:E7 | 27 | 8 | 0 | 獨立小工具：RC snubber |
| 16 | Parallel Output Caps | A1:E16 | 36 | 10 | 0 | 獨立小工具：兩組並聯電容的漣波分流 |

全部 16 個分頁都是「可見」狀態。活頁簿另有 **326 個具名範圍**（見附錄 A）與一個 VBA 專案（見 §1.4）。

### 0.2 資料流

```
                 ┌────────────────────────────── 使用者 ──────────────────────────────┐
                 │  輸入格 / 下拉選單 / 微調鈕 / 核取方塊 / FIX/RESET 按鈕                  │
                 ▼                                                                    │
        ┌──────────────────┐   ='Design Converter'!Exx (換成 SI 單位)    ┌───────────────┐
        │ Design Converter │ ─────────────────────────────────────────▶ │ Variable Mgmt │
        │      (DC)        │ ◀───── 顯示：='Variable Management'!Bxx /   │     (VM)      │
        └──────────────────┘          具名範圍 (Lopt_Limit, Rs_ideal…)    └──────┬────────┘
            ▲   ▲   ▲                                                        │ ▲
            │   │   │ INDEX('CFG Setting'!B4:B19,CFG)                         │ │
            │   │   └──────────── CFG Setting ◀── 下拉選單寫入 C27:F27 ─────────┤ │
            │   │                                                            │ │
            │   └── 'Standard Value Calculator'!O4/O27/O28/O29 ◀── 理想值 ─────┤ │ 標準值
            │                                                                ▼ │
            │   DeviceParameter (Device→CS 門檻、gm、tON,min→C_CS)        Timings (VIN 轉換點)
            │                                                                │
            │   Stability / IL Stability (Bode) ◀── 極零點、增益 ───────────────┤
            │   Efficiency vs VIN / vs Load ◀──── 元件參數、Timings 三相時間 ───┘
            │                │
            └─ 連結圖片 ◀── Efficiency Plot Management / Schematic Management (INDIRECT 切換)
```

**原則**：DC 只放輸入與顯示，所有計算都在 VM。VM 的 B 欄值一律用 SI 單位（V、A、H、F、Ω、Hz），由 DC 的 µH、mΩ、kHz 等顯示單位換算過來。

### 0.3 目前存檔的使用者選項

| 控制項 | 寫入格 | 目前值 | 意義 |
|---|---|---|---|
| Device 下拉（L29:L30） | `VM!C29` | 1 | `Device` = "LM51770"（2 = LM517701） |
| 效率圖類型下拉（K8:K9） | `VM!L10` = `PLOT_TYPE` | 1 | 1 = Efficiency vs IOUT，2 = vs VIN |
| Show I_Limit Bode Plot 核取方塊 | `VM!L11` = `BODEPLOT_TYPE` | FALSE | 顯示電壓迴路 Bode 圖 |
| DRSS 下拉 | `CFG!C27` | 1 | off |
| HICCUP 下拉 | `CFG!D27` | 2 | on |
| PSM Entry 下拉 | `CFG!E27` | 1 | 10 % |
| IMON/ILIMIT 下拉 | `CFG!F27` | 3 | ILIMIT positive |
| VIN(nom) 微調鈕 | `DC!E7` | 12 | 0–58 V，步進 1 |
| 交越頻率微調鈕 | `DC!E83` | 4 | 1–100 kHz |
| `Calc_mode`（無控制項，直接寫值） | `VM!L14` | 1 | 1 = "Joerg" 算法，0 = LM5177 算法 |
| `Mode`（直接寫值） | `VM!B30` | 1 | 0 = PSM，1 = FPWM |

---

## 1. Design Converter（DC）：使用者介面

### 1.1 功能

使用者在這一頁依 Step 1 到 Step 8 填入規格與元件值，同一頁會顯示建議值、檢查結果、Bode 圖、效率圖與電路圖。白色格是輸入格（F3 的圖例寫「= Input Box」）。

### 1.2 輸入格（常數）

| 格 | 項目 | 值 | 單位 | VM 中接收的格 / 名稱 |
|---|---|---|---|---|
| E6 | VIN(min) | 6 | V | B7 `VIN_min` |
| E7 | VIN(nom)（有微調鈕） | 12 | V | B8 `VIN_nom` / `Vin` |
| E8 | VIN(max) | 36 | V | B9 `VIN_max` |
| E9 | VOUT | 16 | V | B12 `Vout` |
| E10 | IOUT(max) | 8 | A | B13 `Iout` |
| E11 | fSW | 398 | kHz | B22 `Fsw`（×1000） |
| E16 | LF | 1.8 | µH | B75 `Lf`（÷1e6） |
| E17 | DCR | 3.2 | mΩ | B76 `Rdcr`（÷1000） |
| E28 | 峰值限流裕度 | 20 | % | B109 `Ilimit_Margin` |
| E30 | RS | 1.3 | mΩ | B146 `Rs` |
| E41 | 輸出漣波規格 | 150 | mVpp | B85 `Vripple1` |
| E43 | COUT | 130 | µF | B198 `Cout` |
| E45 | COUT ESR | 2 | mΩ | B202 `CoutEsr` |
| E50 | 輸入漣波規格 | 180 | mVpp | B212 `Vinripple1` |
| E52 | CIN | 100 | µF | B218 `Cin` |
| E54 | CIN ESR | 2 | mΩ | B223 `RCinEsr` |
| E60 | tSS | 1.8 | ms | B232 `Tss` |
| E65 | UVLO On | 4.6 | V | B239 `VINuvlo_on` |
| E66 | UVLO Off | 4.025 | V | B240 `VINuvlo_off` |
| E70 | 定電流迴路設定點 | 5 | A | B268 `Constant_Current_Loop_Setpoint` |
| E83 | 期望交越頻率（有微調鈕） | 4 | kHz | B380 `Fcross` |
| E85 | RFBT | 71.5 | kΩ | B396 `Rfb1_` |
| M91 / M92 | 選用的 CIMON / RIMON | 2.2 nF / 0.7 kΩ | | B289 `Cc1IL` / B288 `RcIL` |
| E98, E100:F109 | 降壓臂 MOSFET Q1/Q2（SQJ422EP）參數 | RDS(on) 3.5 mΩ、QG 67 nC、QGD 10.31、QGS 11.25、RG 0.63 Ω、gFS 117 S、VTH 2 V、VBD 0.75 V、θJA 65 | | VM B430:D443 |
| F108 | Q2 QRR | `=45` | nC | VM D441 `Qrr_Q2` |
| E111, E113:F122 | 升壓臂 MOSFET Q3/Q4 參數（同上） | | | VM B447:D460 |
| E121 | Q3 QRR | 56 | nC | VM B458 `Qrr_Q3` |
| E125:F126 | 外接 Schottky 的 VF、QRR | 0 | | VM B463:D464 |

### 1.3 公式格（全部 73 格）

**標題 / 連結**

| 格 | 公式 | 值 | 說明 |
|---|---|---|---|
| A1 | `=CONCATENATE(Device," DC/DC Controller Design Tool")` | LM51770 DC/DC Controller Design Tool | 標題隨 `Device` 變動 |
| F3 | `= Input Box` | 圖例文字 | |
| J3 | `=HYPERLINK(CONCATENATE("http://www.ti.com/product/",LOWER(Device)))` | ti.com/product/lm51770 | 產品頁連結 |
| J4 | `=CONCATENATE("Version: ",Licenses!B5)` | Version: 1.0.1 | |

**Step 1：操作規格**

| 格 | 公式 | 值 | 說明 |
|---|---|---|---|
| E12 | `=IF('Standard Value Calculator'!O4/1000>288,"N/A",'Standard Value Calculator'!O4/1000)` | 75 kΩ | RT 取 E96 標準值；理想值 VM B26 `Rfadj = (1/Fsw − 20 ns)·30.3e9` |

**Step 2：電感**

| 格 | 公式 | 值 | 說明 |
|---|---|---|---|
| E15 | `=Lopt_Limit*1000000` | 1.829 µH | 建議電感（VM B74） |
| E18/E19/E20 | `='Variable Management'!B87` / `B88` / `B89` | 5.28 / 4.23 / 12.43 Apk-pk | VIN(min/nom/max) 的電感漣波 |
| E21/E22/E23 | `='Variable Management'!B99*100` / `B100` / `B101` | 23.5 / 37.7 / 155.4 % | 漣波百分比（Boost 除以 Iin，Buck 除以 Iout） |
| E24 | `=R_CS` | 10 Ω | CS 差動濾波電阻（DP N8 常數） |
| E25 | `=C_CS_ideal*1000000000000` | 180 pF | CS 濾波電容（DP N10 → SVC O14） |

**Step 3：OCP、感測電阻、斜率補償**

| 格 | 公式 | 值 | 說明 |
|---|---|---|---|
| E29 | `=Rs_ideal*1000` | 1.639 mΩ | 建議 RS（VM B145） |
| E31 | `='Variable Management'!B159` | 44.23 A | 最壞情況峰值限流 = CS門檻max / RS |
| E32 | `=Ipeak_inductor_max` | 25.10 A | 滿載下最大電感峰值電流 |
| E33 | `='Variable Management'!B162` | 1.13 W | 過流時 RS 的功耗 |
| E34 | `='Variable Management'!B165/1000` | 69.8 kΩ | 建議 RSLOPE（E96） |
| E35 | `='Variable Management'!B165/1000` | 69.8 kΩ | 「選用的」RSLOPE。這格是公式不是輸入格，要自訂就得直接覆寫 |
| E36/E37/E38 | `='Variable Management'!B151` / `B152` / `B153` | 13.2 / 27.2 / 32.2 A | OCP 起始時的輸出電流 |

**Step 4 / 5：輸出 / 輸入電容**

| 格 | 公式 | 值 | 說明 |
|---|---|---|---|
| E42 | `='Variable Management'!B197` | 84.5 µF | 最小 COUT |
| E44 | `='Variable Management'!B201` | 6.56 mΩ | 允許的最大 ESR（不合理時顯示 "N/A"） |
| E46 | `=Vripple2*1000` | 113.5 mVpp | 實際輸出漣波（VM B203） |
| E47 | `='Variable Management'!B192` | 10.33 Arms | COUT RMS 電流最大值 |
| E51 | `='Variable Management'!B217` | 27.9 µF | 最小 CIN（下限 10 µF） |
| E53 | `='Variable Management'!B222` | 9.13 mΩ | CIN 最大 ESR |
| E55 | `=Vinripple2*1000` | 78.7 mVpp | 實際輸入漣波 |
| E56 | `='Variable Management'!B210` | 3.98 Arms | CIN RMS 電流最大值 |

**Step 6：軟啟動、UVLO、CFG、IMON/ILIM**

| 格 | 公式 | 值 | 說明 |
|---|---|---|---|
| E59 | `=E43*0.000001*E9/(E10)*1000` | 0.26 ms | 建議最短 tSS = COUT·VOUT/IOUT |
| N59 / N97 | `=VIN_nom&"V"` | 12V | 圖表標註 |
| E61 | `=Css_u*1000000000` | 18 nF | CSS 標準值（VM B235 ← SVC B4） |
| E63 | `=CFG` | 7 | CFG 組合編號（VM O10 ← CFG G28） |
| E64 | `=INDEX('CFG Setting'!B4:B19,CFG,0)` | 5.1 kΩ | RCFG |
| E67 / E68 | `='Variable Management'!G241` / `G242` | 32.4 / 13.3 kΩ | RUV1 / RUV2（E96） |
| D70 | `=CHOOSE(IMONLIM_select+1,"","Maximum Current to Monitor","Average Current Limit","Average Current Limit")` | Average Current Limit | 標籤依 IMON/ILIM 模式切換 |
| D71 | `=CHOOSE(IMONLIM_select+1,"","Maximum Voltage at IMONOUT","","")` | (空) | |
| E71 | `=CHOOSE(IMONLIM_select+1,"","1.5","5","0")` | "5" | IMONOUT 最大電壓（**文字型**） |
| E72 | `=Rshunt` | 10 mΩ | 輸出 / 輸入電流感測分流電阻 |
| D73 / E73 | `=CHOOSE(IMONLIM_select+1,"","IMONOUT Resistance, RIMON","SYNC pin during startup",…)` / `=CHOOSE(…,'Variable Management'!B274/1000,"High","Low")` | SYNC pin during startup / High | IMON 模式顯示 RIMON，ILIM 模式顯示 SYNC 腳狀態 |

**Step 7：補償**

| 格 | 公式 | 值 | 說明 |
|---|---|---|---|
| E78 | `=Fp_dominant` | 1224 Hz | 負載極點 |
| E79 | `=Fesr/1000` | 612 kHz | ESR 零點 |
| E80 | `=IF(E6<E9,Frhp_Vinmin/1000,"---")` | 24.1 kHz | VIN(min) 的 Boost RHP 零點 |
| E81 | `=IF(E7<E9,Frhp/1000,"---")` | 98.8 kHz | VIN(nom) 的 RHP 零點 |
| E82 | `=MIN(Fsw/1000*0.05,E80/4,IF(E81<>"---",E81/4,E80/4))` | 6.03 kHz | 建議交越頻率 = min(5 % fSW, fRHP/4) |
| F83 | `="     kHz "` | | 單位文字 |
| E84 | `=Fp1_u` | 0.284 Hz | EA 低頻極點 |
| E86 / F86 | `=IF(Vout=Vref,"OPEN",Rfb2_u/1000)` / `=IF(Vout=Vref,"","kΩ")` | 4.75 kΩ | RFBB 標準值 |
| E87 | `=1/E86*(E85+E86)` | 16.05 V | 用標準值算出的實際 VOUT（Vref = 1 V） |

**Step 7.1 / 7.2：補償元件**

| 格 | 公式 | 值 | 說明 |
|---|---|---|---|
| B91 / C91 / **E91** | `=Rcea1/1000` / `='Standard Value Calculator'!O27` / `=C91` | 1.514 / 1.5 / 1.5 kΩ | RCOMP：計算值 / 標準值 / **選用值**（E91 = 具名範圍 `DC_RC`，FIX/RESET 巨集操作的就是這格） |
| G91 | `=Fp1_u` | 0.284 Hz | |
| B92 / C92 / E92 | `=Cc1ea*1e9` / `=SVC!O28` / `=C92` | 57.2 / 56 / 56 nF | CCOMP |
| G92 | `=Fz1_u/1000` | 1.895 kHz | 補償零點 |
| B93 / C93 / E93 | `=IF(Cc2ea*1e12<0,0,Cc2ea*1e12)` / `=SVC!O29` / `=C93` | 2621 / 2700 / 2700 pF | CHF |
| G93 | `=Fp2_u/1000` | 39.2 kHz | 補償高頻極點 |
| K91 / L91 | `=C_IL_recommend*1e9` / `='Variable Management'!F292*1e9` | 1.396 / 1.5 nF | 建議 CIMON / 標準值 |
| K92 / L92 | `=R_IL_recommend/1000` / `='Variable Management'!F293/1000` | 0.477 / 0.475 kΩ | 建議 RIMON / 標準值 |
| M90 | `="    Selected"` | | 標題 |

### 1.4 控制項、巨集與圖片

| 物件 | 類型 | 連結 / 巨集 | 作用 |
|---|---|---|---|
| Button 492「RESET VALUES」 | 按鈕 | `RESETValues` | 把 E91:E93 改回公式 `=RC[-2]`（也就是 `=C91`…`=C93`），補償值重新跟著標準值走 |
| Button 520「FIX VALUES」 | 按鈕 | `FIXValues` | 把 E91:E93 複製後「貼上值」，凍結補償元件。之後掃 VIN(nom) 時補償就不會跟著重算 |
| Spinner ×4 | 微調鈕 | `$E$7`（0–58）、`$E$83`（1–100） | 微調 VIN(nom) 與交越頻率 |
| Drop Down 4707 | 下拉 | `VM!C29`，清單 `VM!L29:L30` | Device |
| Drop Down 896 | 下拉 | `VM!L10`，清單 `VM!K8:K9` | 效率圖類型 |
| Drop Down 1749/1752/1753/1754 | 下拉 | `CFG!C27`/`D27`/`E27`/`F27` | DRSS / HICCUP / PSM / IMON-ILIMIT |
| Check Box 3516 | 核取方塊 | `VM!L11` | 顯示 I-Limit Bode 圖 |
| Picture 30 / 28 / 68 / 21 | 連結圖片（照相機） | 名稱 `PICTURE`、`Picture_Bode`、`PICTURE2` | 用 `INDIRECT` 切換顯示哪張效率圖、Bode 圖、電路圖 |
| 巨集 `OpenLM27403ProductFolder` | 圖片的 macro 屬性 | VBA 中**不存在** | 從 LM27403 工具沿用下來的殘留，點了會出錯 |

VBA 只有 `Module5`（上述兩個巨集）與 `Sheet1` 的四個空事件程序（`ListBox1_Click` 等）。

### 1.5 條件格式（輸入檢查）

| 格 | 條件 | 意義 |
|---|---|---|
| E7 | 不在 E6~E8 之間 | VIN(nom) 超出範圍 |
| E9 | 不在 3.3~78 之間 | VOUT 超範圍 |
| E30 | `ROUND(E30,1)>ROUND(E29,1)` | 選用的 RS 大於建議值 |
| E41 / E50 | ≤ 0 / = 0 | 漣波規格無效 |
| E43 / E52 | < E42 / < E51 | 電容小於最小值 |
| E45 / E54 | > E44 / > E53（E54 另有 = 0） | ESR 過大 |
| E60 | ≤ E59 | tSS 太短 |
| E65 / E66 | 不在 2.8~78 之間 | UVLO 超範圍 |

另有 x14 擴充格式的條件規則（openpyxl 不支援），例如儲存格註解提到的 E32、E34 標紅，本文沒有列出。

---

## 2. Licenses

| 格 | 內容 |
|---|---|
| B5 | 版本號 `1.0.1`（DC!J4 引用） |
| A36:B40 | 版本歷史：1.0.0 Initial Release；1.0.1 改善 Desired Crossover Frequency 計算、加入 LM517701、LM51770 最大峰值限流門檻調為 57.5 mV |
| 文字方塊 ×4 | "About this tool…"、© 2026 Texas Instruments、IMPORTANT NOTICE AND DISCLAIMER |

沒有公式。

---

## 3. Variable Management（VM）：核心計算

欄位慣例：A = 變數名稱，B = 值（SI），C = 單位，D/E = 註解。K~R 欄是對照或替代算法。表頭 B2:B4 說明顏色：Input from user（從 DC 搬來）、Output、Constant。

### 3.1 Step 1：操作規格（列 7–47）

| 格 | 名稱 | 公式 | 值 | 說明 |
|---|---|---|---|---|
| B7 | `VIN_min` | `='Design Converter'!E6` | 6 | |
| B8 | `VIN_nom`, `Vin` | `='Design Converter'!E7` | 12 | **注意 `Vin` 與 `VIN_nom` 指同一格**；補償、Bode、IL 模型都用 VIN(nom) 計算 |
| B9 | `VIN_max` | `='Design Converter'!E8` | 36 | |
| B10 | `VIN_edgebuckboost` | `=B44` | 11.13 | Buck-Boost 區下緣（= BB→Boost 轉換點） |
| B12 | `Vout` | `='Design Converter'!E9` | 16 | |
| B13 | `Iout` | `='Design Converter'!E10` | 8 | |
| B14 | `Rout` | `=Vout/Iout` | 2 Ω | 滿載電阻 |
| B15 | `Pout` | `=Vout*Iout` | 128 W | |
| B16 | `Efficiency` | 常數 0.95 | | 註解寫 "from 96% to 90%"，但實際值是 95 % |
| B17 | `Pin` | `=Pout/Efficiency` | 134.7 W | |
| B18/B19/B20 | `Iin_Vinmin/nom/max` | `=Pin/VIN_min` 等 | 22.46 / 11.23 / 3.74 A | |
| B22 | `Fsw` | `='Design Converter'!E11*1000` | 398 000 Hz | |
| B23 | `Ts` | `=1/Fsw` | 2.513 µs | |
| B24 | `Vref` | 常數 1 V | | |
| B25 | Vref_eff | `=Vref` | 1 | 沒有被引用 |
| B26 | `Rfadj` (Rt) | `=(1/Fsw-0.00000002)*30300000000` | 75 525 Ω | RT 擬合式（註解寫 "edited based on LM5177 curve fit"） |
| B27 | `Ifb` | 6e-8 A | | 只用於 `Rfb2_max` |
| B29 | `Device` | `=CHOOSE(C29,L29,L30)` | LM51770 | C29 由下拉選單寫入 |
| B30 | `Mode` | 常數 1 | | 0 = PSM，1 = FPWM |
| L8:L9, K10:L11 | `PLOT_TYPE`(L10)、`BODEPLOT_TYPE`(L11) | 由控制項寫入 | 1 / FALSE | |
| L14 | `Calc_mode` | 常數 1 | | 決定 `Rs_ideal` 用哪個公式 |
| O10 | `CFG` | `='CFG Setting'!G28` | 7 | |
| B44 | `VIN_BB_to_boost` | `=IF($B$30=0,Vout,Timings!F5)` | 11.13 V | PSM 模式下所有轉換點都簡化成 Vout |
| B45 | `VIN_boost_to_BB` | `=IF($B$30=0,Vout,Timings!F4)` | 14.20 V | |
| B46 | VIN_buck_to_BB | `=IF($B$30=0,Vout,Timings!F2)` | 18.02 V | |
| B47 | `VIN_BB_to_buck` | `=IF($B$30=0,Vout,Timings!F3)` | 23.00 V | |

> **列 31–43 是空的**，但仍有具名範圍指向那裡：`Toff_min_buck`(B32)、`Ton_min_boost`(B33)、`D_min_boost`(B36)、`D_max_buck_BB`(B38)、`D_min_boost_exitBB_to_buck`(B40)、`D_max_exitBB_to_buck`(B41)、`D_max_buck_enterBB`(B42)。另外 B58 `Don_buckboost = B37` 也是空格，等於 0，所以 `Iripple_edgebuckboost` 恆為 0、`Ipeak_inductor_buckboost` = Iout/η。

### 3.2 Step 2：電感（列 51–104）

| 格 | 名稱 | 公式 | 值 | 說明 |
|---|---|---|---|---|
| B52 / C52 | `_Don1` | `=Vout/VIN_nom` / `=(Vout-VIN_nom)/Vout` | 1.333 / 0.25 | 理想 duty（Buck 式 / Boost 式） |
| B53 | `Doff1` | `=1-_Don1` | −0.333 | 沒有被下游使用 |
| B55:B57 | `Don_Vinmin/nom/max` | `=IF(VIN_x>Vout, Vout/VIN_x, (Vout-VIN_x)/Vout)` | 0.625 / 0.25 / 0.444 | 理想 duty；C 欄標 BUCK/BOOST |
| B58 | `Don_buckboost` | `=B37` | 0 | 指向空格（見上方說明） |
| B59 | Rs_Init | `=Boost_CS_Threshold/(Iin_Vinmin*1.4)` | 1.59 mΩ | 沒有被引用 |
| B61:C63 | Lf_Vinmin/nom/max | B 欄是 Buck 式 `(VIN−Vout)·D/(Fsw·Iout·k)`，C 欄是 Boost 式 `VIN/(Iin·k·Fsw)·(1−VIN/Vout)`；k = 0.3 / 0.4 / 0.6 | | 依 30/40/60 % 漣波目標算的電感 |
| B65 | L_Vinmin | `=MAX(IF(VIN_min>Vout,B61,C61),O65)` | 1.40 µH | |
| B66 | L_Vinnom | `=IF(VIN_nom>Vout,B62,C62)` | 1.68 µH | |
| B67 | L_Vinmax | `=MIN(IF(VIN_max>Vout,B63,C63),O66)` | 1.83 µH | |
| L64 / O64 | `Labsmin` / `Labsmax` | `=Rs/8000` / `=Rs/100` | 0.1625 / 13 µH | Rs/L 必須在 100~8000 之間 |
| L65 / L66 | Lbuck / Lboost | `=(VIN_max-Vout)*Vout/(0.4*Iout*Fsw*VIN_max)` / `=(Vout-VIN_min)*VIN_min^2/(0.3*Iout*Fsw*Vout^2)` | 6.98 / 1.47 µH | "Joerg's Calculation" 對照區 |
| O65 | Lmin | `=MAX(10*Vout*Rs/Fsw,Labsmin)` | 0.523 µH | |
| O66 | Lmax | `=MIN(Rs*35*Vout/(Fsw*1),Labsmax)` | 1.829 µH | 註解："updated factor of 25 to 35 to get EVM into allowed range" |
| L70 / L71 | L_avg / L_min | `=MAX(MIN(AVERAGE(L65:L66),O66),O65)` / `=MAX(MIN(MIN(L65:L66),O66),O65)` | 1.83 / 1.47 µH | 只是參考，沒有被引用 |
| L73 | L(peak limit) | `=VIN_min*(Vout-VIN_min)/(0.15*Fsw*Vout)*Rs/0.05` | 1.63 µH | 註解："used datasheet value of 15%" |
| B71 | `Lopt` | `=VIN_min^2*(Vout-VIN_min)/(0.2*Iout*Fsw*Vout^2)` | 2.21 µH | 以 VIN(min) Boost、20 % 漣波為準 |
| B72 / B73 | Lmin / Lmax | `=O65` / `=O66` | | |
| B74 | `Lopt_Limit` | `=MAX(MIN(B71,B73),B72)` | 1.829 µH | **DC E15 的建議值**（Lopt 夾在 Lmin~Lmax 之間） |
| B75 | `Lf` | `='Design Converter'!E16/1000000` | 1.8 µH | |
| B76 | `Rdcr` | `='Design Converter'!E17/1000` | 3.2 mΩ | |
| B78 / B79 | `_Don2` / `Doff2` | `=(Vout+Iout*(Rdcr+Rs))/(Vin+Iout*Rs)` / `=1-_Don2` | | 含壓降的 Buck duty（只給殘留公式 B402/B404 用） |
| B81:B83 | `Don_Vinmin_2/nom/max` | Buck：`(Vout+Iout·(Rdcr+Rs))/(VIN+Iout·Rs)`；Boost：`(Vout−VIN+Iin·Rdcr)/(Vout−Iin·Rs)` | 0.631 / 0.252 / 0.445 | **含 DCR 與 RS 壓降的 duty，Step 2~5 都用這組** |
| L81 | Vijay's simplification | `=(1-VIN_min*90%/Vout)` | 0.6625 | 參考 |
| B85 | `Vripple1` | `='Design Converter'!E41/1000` | 0.15 V | |
| B87:B89 | `Iripple_Vinmin/nom/max` | Buck：`(VIN−Vout)·D2/(Fsw·Lf)`；Boost：`VIN·D2/(Lf·Fsw)` | 5.28 / 4.23 / 12.43 A | C 欄一律用 Buck 式（Boost 時是負值）；H87 用 Lopt 對照 |
| B90 | `Iripple_edgebuckboost` | `=VIN_edgebuckboost*Don_buckboost/(Lf*Fsw)` | 0 | 因 Don_buckboost = 0 |
| B92:B94 | Ishunt_rms_Vinmin/nom/max | Buck：`√(1−D)·√(Iout²+ΔI²/12)`；Boost：`√D·√(Iin²+ΔI²/12)` | 17.9 / 5.68 / 6.53 A | E/F 欄是兩種公式都算 |
| B95 | Max shunt power loss at full load | `=MAX(B92,B94)^2*Rs` | 0.415 W | |
| B96 | | `=(Buck_CS_threshold_max/Rs)^2*Rs*Don_Vinmin` | 1.59 W | 沒有被引用 |
| B97 | Max inductor ripple current | `=MAX(B87:B89)` | 12.43 A | |
| B99:B101 | % Ripple | Buck：`ΔI/Iout`；Boost：`ΔI/Iin` | 0.235 / 0.377 / 1.554 | DC E21~E23 |
| B103 / B104 | Temp / DCR at hot | 100 / `=Rdcr*(1+0.0039*(B103-25))` | 4.14 mΩ | 沒有被引用 |

### 3.3 Step 3：感測電阻、斜率補償、限流（列 107–180）

| 格 | 名稱 | 公式 | 值 | 說明 |
|---|---|---|---|---|
| B109 | `Ilimit_Margin` | `='Design Converter'!E28` | 20 % | |
| B111:B113 | `Ipeak_inductor_Vinmin/nom/max` | Buck：`Iout/η+ΔI/2`；Boost：`Iin+ΔI/2` | 25.10 / 13.34 / 14.64 A | ⚠ B113 的 Boost 分支用了 `Iin_Vinmin`（應為 `Iin_Vinmax`）；只有 VIN_max < Vout 時才會影響結果 |
| B114 | `Ipeak_inductor_buckboost` | `=Iout/Efficiency+Iripple_edgebuckboost/2` | 8.42 A | |
| B115 | `Ipeak_inductor_max` | `=MAX(B111:B113)` | 25.10 A | DC E32 |
| B116 | Ipeak_inductor_max_margin | `=Ipeak_inductor_max*(1+Ilimit_Margin/100)` | 30.1 A | |
| B118:B120 | `Buck_CS_threshold_min/typ/max` | `=DeviceParameter!B16/B17/B18` | 42.5 / 50 / 57.5 mV | L/M 欄是兩種 IC 的對照表（LM517701 = 65.6/75/84.4 mV） |
| B122:B124 | `Boost_CS_Threshold_min/typ/max` | 同上 | 同上 | Buck 與 Boost 門檻相同 |
| B127:B130 | Shunt required for OCP @ VIN(min/nom/max/edge) | `=門檻/Ipeak/(1+Margin/100)` | 1.66 / 3.12 / 2.85 / 4.95 mΩ | 參考 |
| B133 | Ipeak_peak_inductor | `=Don_Vinmin*VIN_min/(Lf*Fsw)` | 5.23 A | 給 Calc_mode = 0 用 |
| B134 | IIN_avg_max | `=Vout*Iout/(Efficiency*VIN_min)` | 22.46 A | |
| B145 | `Rs_ideal` | `=IF(Calc_mode=0, Boost_CS_Threshold_min/((B134+0.5*B133)*(1+Ilimit_Margin/100)), Buck_CS_threshold/(MAX(Iout,Vout/VIN_min*Iout)*1.43))` | 1.639 mΩ | **DC E29**。Calc_mode = 1：50 mV/(1.43·max(Iout, Iout·Vout/VIN_min)) |
| B146 | `Rs` | `='Design Converter'!E30/1000` | 1.3 mΩ | |
| B147 | `Rs_in_range` | `=IF(AND(100<Rs/Lf,Rs/Lf<8000,Rs/Lf<(1*Fsw/(Vout*10))),1,0)` | 1 | Rs/L 範圍檢查 |
| B148 / B149 | BOOST / BUCK current limit - peak | `=Boost_CS_Threshold_max/Rs` / `=Buck_CS_threshold_max/Rs` | 44.2 A | |
| B151:B153 | `Iout_OCP_Vinmin/nom/max` | Buck：`門檻/Rs−ΔI/2`；Boost：`(門檻/Rs−ΔI/2)·(1−D2)` | 13.2 / 27.2 / 32.2 A | DC E36~E38 |
| B154 | Iout_edgebuckboost | `=(Buck_CS_threshold/Rs-Iripple_edgebuckboost/2)*(1-Don_buckboost)` | 38.5 A | |
| B156:B158 | `ILpeak_OCP_Vinmin/nom/max` | `=門檻max/Rs`（nom 依 BUCK/BOOST/BUCK-BOOST 選擇） | 44.2 A | H/K/M 欄是過流時 RS 的 RMS 電流與功耗 |
| B159 / B161 | Worst case peak / `IL_peak_OCP` | `=MAX(B156:B158)` | 44.2 A | DC E31 |
| B162 | Max shunt power loss at OC | `=(Boost_CS_Threshold_max/Rs)^2*Rs*(Vout/VIN_max)` | 1.13 W | DC E33 |
| B164 | `Rslope_ideal` | `=Lf/Rs*50*10^6` | 69 231 Ω | "using 50MV/As from Datasheet" |
| B165 | Rslope_recommended | `='Standard Value Calculator'!O15` | 69 800 Ω | E96 |
| B166 / B173 | Rslope_selected / `Rslope` | `='Design Converter'!E35*1000` / `=B166` | 69 800 Ω | |
| B172 | `Islope_fixed` | `=Isc_const` | 0 | |
| B174:B176 | SLOPE comp dv/dt @ VIN(min/nom/max) | `=(ABS(VIN_x-Vout)*gm_slopecomp/Cslope+Islope_fixed)/Rslope` | 1.07e5 / 4.30e4 / 2.15e5 V/s | K/L 欄乘上 (1−D)/Fsw 或 D/Fsw，得到每週期的斜坡電壓 |
| B178:B180 | DV_SLOPEcomp @ VIN(min/nom/max) | `=IF(D="BUCK",K,L)` | 0.170 / 0.027 / 0.300 V | |

### 3.4 Step 4：輸出電容（列 188–204）

| 格 | 名稱 | 公式 | 值 |
|---|---|---|---|
| B189:B191 | `Icoutrms`(B189)… | Buck：`ΔI/√12`；Boost：`Iout·√((Vout−VIN)/VIN)` | 10.33 / 4.62 / 3.59 A |
| B192 | Max Cout rms | `=MAX(B189:B191)` | 10.33 A |
| B194 | | `=Vripple1*1000` | 150 mV |
| B196 | `Coutmin` | Buck：`ΔI/(8·Fsw·Vripple1)`；Boost：`Iout·D2/(Vripple1·Fsw)`（判斷基準一律是 VIN_min） | 84.5 µF |
| B197 | Reported Coutmin | `=IF(Coutmin*1e6<H197,H197,Coutmin*1e6)`（H197 = 0.1） | 84.5 µF |
| B198 | `Cout` | `='Design Converter'!E43/1000000` | 130 µF |
| B199 | `CoutEsrMax` | Buck：`√((Vr/ΔI)²−(1/(8FswC))²)`；Boost：`(Vr−Iout·D2/(Fsw·C))/Iout` | 6.56 mΩ |
| B201 | Reported ESR max | `=IF(OR(>0.3,<0,ISERROR),"N/A",×1000)` | 6.56 mΩ |
| B202 | `CoutEsr` / `RCoutEsr` | `='Design Converter'!E45/1000` | 2 mΩ |
| B203 | `Vripple2` | Buck：`ΔI_max·√(ESR²+(1/(8FswC))²)`；Boost：`ESR·Iout+Iout·D2/(Fsw·C)` | 113.5 mV |
| K202:L203, K204 | 偵錯用的 ISERROR/OR 檢查 | | |

### 3.5 Step 5：輸入電容（列 205–229）

| 格 | 名稱 | 公式 | 值 | 說明 |
|---|---|---|---|---|
| B207:B209 | Icinrms | Buck：`Iout·√(D·(1−D))`；Boost：`ΔI/√12` | 1.52 / 1.22 / 3.98 A | ⚠ B207 的 Buck 分支寫成 `Don_Vinmin_2*(1-Don_Vinmax_2)`，min 與 max 混用 |
| B210 | Max Cin rms | `=MAX(B207:B209)` | 3.98 A | DC E56 |
| B212 | `Vinripple1` | `='Design Converter'!E50/1000` | 0.18 V | |
| B214 | Cinmin_Vinmax | Buck：`Iout·D(1−D)/(Vr·Fsw)`；Boost：`ΔI/(8·Fsw·Vr)` | 27.6 µF | |
| B215 | Cinmin_50%D | D = 0.5 的最壞情況 | 27.9 µF | |
| B216 | `Cinmin` | `=IF(VIN_max<2*Vout,B214,B215)` | 27.9 µF | VIN_max ≥ 2·Vout 時才會經過 D = 50 % |
| B217 | Reported Cinmin | 下限 10 µF | 27.9 µF | |
| B218 | `Cin` | `='Design Converter'!E52/1000000` | 100 µF | |
| B219:B221 | `CinEsrMax` | 同樣分 Vinmax / 50 %D 兩種算法 | 9.13 mΩ | |
| B222 | Reported | `=IF(OR(>0.3,<0),"N/A",×1000)` | 9.13 | |
| B223 | `RCinEsr` | `='Design Converter'!E54/1000` | 2 mΩ | |
| B225:B227 | `Vinripple2` | 電容充放電 + ESR 壓降 | 78.7 mV | |
| B229 | | `=Iout*Don_Vinmax_2*(1-Don_Vinmax_2)/(Cin*Fsw)` | 49.7 mV | 參考 |

### 3.6 Step 6：軟啟動、UVLO、IMON / ILIM（列 231–293）

| 格 | 名稱 | 公式 / 常數 | 值 | 說明 |
|---|---|---|---|---|
| B232 | `Tss` | `='Design Converter'!E60/1000` | 1.8 ms | |
| B233 | `Iss` | 1e-5 A | | |
| B234 | `Css` | `=Iss*Tss/Vref` | 18 nF | |
| B235 | `Css_u` | `='Standard Value Calculator'!B4` | 18 nF | E12 標準值 |
| B239 / B240 | `VINuvlo_on/off` | DC E65 / E66 | 4.6 / 4.025 V | |
| B241 | `Ruvlo1` | `=1/Iuvlo1*(VINuvlo_on-Vuvlo_on*VINuvlo_off/Vuvlo_off)` | 32.76 kΩ | 單位 kΩ（因 Iuvlo1 用 mA） |
| B242 | `Ruvlo2` | `=Ruvlo1/(VINuvlo_off/Vuvlo_off-1)` | 13.43 kΩ | |
| E241 / E242 | | `=SVC!O16` / `=SVC!O17` | 32.4 / 13.3 | E96 |
| G241 / G242 | | `=IF(Ruvlo1<0,"N/A",E24x)` | 32.4 / 13.3 | DC E67/E68 |
| H241 / H242 | | 加上 "kΩ" 的文字 | | 電路圖標籤 |
| B244 / B245 | VINUVLO_ON/OFF_ACTUAL | `=(E241+E242)/E242*Vuvlo_on+Iuvlo1*E241` / `=(E241+E242)/E242*Vuvlo_off` | 4.593 / 4.020 V | 用標準值反算的實際門檻 |
| B247 / B248 / B249 | `Vuvlo_on`/`Vuvlo_off`/`Vuvlo_hys` | 1.28 / 1.17 / `=B247-B248` | | |
| B250 | `Iuvlo1` | 0.006 mA | | B251/B252（`Iuvlo2`、`Iuvlo_hys`）是空格 |
| B254 / B256 | `Rpg` / Rh | 20 kΩ / 0 | | B255 `Rsb`、B257 Rmode 是空格（TBD） |
| B266 | `IMONLIM_select` | `='CFG Setting'!F27-1` | 2 | 0 = off，1 = IMON，2 = ILIM(+)，3 = ILIM(−) |
| B268 | `Constant_Current_Loop_Setpoint` | `='Design Converter'!E70` | 5 A | |
| B269 | `Iout_threshold` | 50 mV | | |
| B270 | `Rshunt` | `=IF(ISNUMBER(setpoint), Iout_threshold/setpoint, 0)` | 10 mΩ | |
| B271 | `With_RCS` | `='CFG Setting'!F27` | 3 | 用來選電路圖 |
| B272 | g_m | `=DeviceParameter!B11` | 1 mS | IMON 放大器 gm |
| B273 | Vimonitor,out,max | `='Design Converter'!E71` | "5" | 文字 |
| B274 | RImonitorOut | `=B273/(setpoint*Rshunt/1000*B272)` | 100 kΩ | RIMON |
| B280 | `gm_IL` | `=DeviceParameter!B12` | 200 µS | |
| B281 | `GBW_IL` | 2 MHz | | |
| B282 | `Cbw_IL` | `=gm_IL/(2*Pi*GBW_IL*1e6)` | 15.9 pF | |
| B283 | `RoEA_IL` | 20 MΩ | | |
| B284 | `A_OL_IL` | `=gm_IL*RoEA_IL` | 4000 | |
| B285 | `Fp_IL` | `=1/(2*Pi*(RoEA+B288)*B289)` | 7.23 Hz | ⚠ 用了電壓迴路的 `RoEA`（10 MΩ），應為 `RoEA_IL` |
| B286 | `Fp_ILHF` | `=1/(2*Pi*B288*(B289+Cbw))` | 103 kHz | ⚠ 用了 `Cbw`，應為 `Cbw_IL` |
| B287 / F287 | `Co2_` / `Co2_u` | `=Cout/2` / `=Co2_` | 65 µF | IL 模型把 COUT 拆成兩半 |
| B288 / B289 | `RcIL` / `Cc1IL` | `=DC!M92*1000` / `=DC!M91*1e-9` | 700 Ω / 2.2 nF | 使用者選的 RIMON / CIMON |
| B290 | `Cc2IL` | 1e-15 F | | |
| B292 | `C_IL_recommend` | `=DeviceParameter!B12/10^((20*LOG(3*f_cross_Bode*1000)+'IL Stability Calculations'!BE13)/20)` | 1.396 nF | gm_IL ÷（3 倍電壓迴路交越頻率 × 1 Hz 的 plant 增益） |
| F292 | | `=SVC!O23` | 1.5 nF | |
| B293 | `R_IL_recommend` | `=1/(2*PI()*C_IL_recommend*20*3*f_cross_Bode*1000)` | 477 Ω | |
| F293 | | `=SVC!O24` | 475 Ω | |

### 3.7 Step 7：補償（列 295–399）

**放大器與斜率補償常數**

| 格 | 名稱 | 值 / 公式 | 說明 |
|---|---|---|---|
| B296 | `Pi` | `=PI()` | |
| B297 | `CSAmp_gain` | 10 | |
| B298 | `Ri` | `=Rs*CSAmp_gain` = 13 mΩ | 等效電流感測電阻 |
| B299 | `gm` | 600 µS | EA 跨導 |
| B300 | `Cslope` | 2.5 pF | 內部斜率補償電容 |
| B301 | `gm_slopecomp` | `=0.00000002*125000*0.75` = 1.875 mS | 自適應斜率補償 gm |
| B302 | `Isc_const` | 0 | 固定斜率電流 |
| B304:B306 | `Se_fixed` / `Se_proportional` / `Se` | `=Isc_const/Cslope`、`=ABS(Vin-Vout)*gm_slopecomp/(Cslope*Rslope)`、和 | 4.30e4 V/s |
| B308:B310 | `V_sc_fixed`/`_proportional`/`_total` | `=Se.../Fsw` | 0.108 V |
| B311 | `K_SL` | `=gm_slopecomp/(Cslope*Fsw)` | 1884 |
| B314 | `Se_ext` | `=(ABS(Vin-Vout)*gm_slopecomp+Isc_const)/(Cslope*Rslope*Ri)` | 3.31e6 A/s | 換算到電感電流的斜率 |
| B315 / B316 | `Sn` / `Sf` | Buck：`(Vin−Vout)/Lf`、`Vout/Lf`；Boost：`Vin/Lf`、`(Vout−Vin)/Lf` | 6.67e6 / 2.22e6 A/s |
| B318:B320 | Se_int / Sn_int / Sf_int | mV/µs 換算 | |
| B322 / B324 | `Toff_Vinmax` / `Ton_Vinmin` | `=(1-Don_Vinmax_2)/Fsw` / Boost 用 `D/Fsw` | 1.39 / 1.58 µs |
| B326:F327, B329:B330 | COMP min/max voltage | `1.6+5·Rs·(…)±斜坡` | 1.26 / 1.93 V | COMP 擺幅估算，沒有被引用 |

**功率級模型（Ridley 型電流模式模型）**

| 格 | 名稱 | 公式 | 值 |
|---|---|---|---|
| B333 / B334 | `mc_buck` / `k_buck` | `=1+Se_ext/Sf` / `=mc_buck*Don_Vinnom_2` | 2.488 / 0.628 |
| B335 / B336 | `mc_boost` / `k_boost` | `=1+Se_ext/Sn` / `=mc_boost*(1-Don_Vinnom_2)` | 1.496 / 1.118 |
| B338 / B339 | `mc` / `k` | `=IF(VIN_nom>Vout, buck, boost)` | 1.496 / 1.118 |
| B340 | `Q` | `=1/Pi/(k-0.5)` | 0.515（註解：Q = 2/π 為理想阻尼） |
| B342:B344 | `KM` / `K_1` / `KD` | 較完整的 Boost 調變器增益參數 | 只出現在 J356 的文字公式 |
| B347:B349 | `GBW` / `Cbw` / `RoEA` | 13 MHz / `=gm/(2*Pi*GBW*1e6)` = 7.35 pF / 10 MΩ | |
| B350:B352 | EA Gain / dB / `AEA` | `=gm*RoEA` = 6000（75.6 dB） | |
| B353 | `Afb` | `=Vref/Vout` = 0.0625 | |
| B354 | `Acs` | `=1/Ri` = 76.9 | |
| B355 | `Am_buck` | `=Ro/Ri*1/((1+Ro*(k-0.5)/(Fsw*Lf)))` = 56.4 | |
| B356 | `Am_boost` | `=Rout*(1-Don_Vinnom_2)/(2*Ri)` = 57.5 | 註解："need to adjust for adaptive slope comp" |
| B357 | `Am` | `=IF(VIN_nom>Vout,Am_buck,Am_boost)` | 57.5 |
| B359 / B360 | Rcomp scaling | `=2*Pi*Ri/gm` | 136 |
| B363 / B364 | `Ro` / Rout | `=Rout` = 2 Ω | |
| B366 / B367 | `Fload_buck` / `Fload_boost` | `=1/(2*Pi*Ro*Cout)` / `=2/(2*Pi*Ro*Cout)` | 612 / 1224 Hz |
| K366 / L366 | | 含斜率補償的 Buck 極點 / LC 諧振 | 1669 / 7777 Hz（參考） |
| B368 | `Fp_dominant` | `=IF(VIN_nom>Vout,Fload_buck,Fload_boost)` | 1224 Hz |
| B370 | `Fp1_u` (Fp_EA) | `=1/(2*Pi*(RoEA+Rcea1_u)*Cc1ea_u)` | 0.284 Hz |
| B371 | `Fp2_u` (Fp_HF) | `=1/(2*Pi*Rcea1_u*(Cc2ea_u+Cbw))` | 39.2 kHz |
| B372 | `Fesr` | `=1/(2*Pi*Cout*CoutEsr)` | 612 kHz |
| B373 | `Frhp` | `=(1-Don_Vinnom_2)^2*Ro/(2*Pi*Lf)` | 98.8 kHz（用 VIN(nom)） |
| B374 | `Frhp_Vinmin` | `=(1-Don_Vinmin_2)^2*Ro/(2*Pi*Lf)` | 24.1 kHz |
| B375 | F_mod | `=Fsw/2` | 199 kHz |

**補償元件計算**

| 格 | 名稱 | 公式 | 值 |
|---|---|---|---|
| B380 | `Fcross` | `='Design Converter'!E83*1000` | 4000 Hz |
| B381 | `Rcea1` | Buck：`Fc·Cout·2π·Ri/(gm·Afb)`；Boost：`Fc·Cout·2π·Ri/(gm·Afb·(1−D)·√(1+(Fc/fRHP)²))` | 1514 Ω |
| K381 | | 不含 RHP 修正的版本 | 1515 Ω |
| B382 | `Cc1ea` | `=1/(Rcea1*(1.5*Fp_dominant)*2*Pi)` | 57.2 nF（零點放在 1.5×負載極點） |
| B384 | | `=MAX(Fcross*7,Fcross*10)` | 40 kHz |
| B383 | `Cc2ea` | `=1/((Rcea1*2*Pi*B384))-Cbw` | 2.62 nF（高頻極點放在 10×fc） |
| B386:B388 | `Rcea1_u` / `Cc1ea_u` / `Cc2ea_u` | `=DC!E91*1000` / `=DC!E92*1e-9` / `=DC!E93*1e-12` | 1.5 kΩ / 56 nF / 2.7 nF（選用值） |
| B390 | Cslope_u | `='Design Converter'!E35*1*10^-12` | 6.98e-11（單位用錯，沒有被引用） |
| B393 | `Fz1_u` | `=1/(2*Pi*Rcea1_u*Cc1ea_u)` | 1895 Hz |
| B394 | Fp1 | 串並聯等效 | 41.1 kHz |
| B396 | `Rfb1_` | `='Design Converter'!E85*1000` | 71.5 kΩ |
| B397 | `Rfb2_` | `=IF(Vout=Vref,1e8,Vref*B396/(Vout-Vref))` | 4767 Ω |
| B398 | `Rfb2_u` | `=IF(Vout=Vref,"N/A",SVC!O10)` | 4750 Ω |
| B399 | Rfb2_max | `=Vref/(100*Ifb)` | 166.7 kΩ |

**殘留 / 錯誤公式（列 402–406）**：H402、B403、B404、B405 引用了不存在的名稱 `RdsonTop`、`RdsonBot`、`Iripple1`，結果是 `#NAME?`，下游沒有使用。

### 3.8 Step 8：效率參數（列 408–484）

把 DC Step 8 的 MOSFET / Schottky 參數換成 SI 單位，並放幾個常數：

| 格 | 名稱 | 值 / 公式 |
|---|---|---|
| B411 | `Ta` | 25 °C |
| B413 | `Ltc` | 0.004（4000 ppm/°C） |
| B414 | `ThetaCa` | 20 °C/W |
| B415 | RCinEsr | `=RCinEsr` |
| B417 | `Vdd` | `=7.5` V |
| B419 / B420 | `RdsonDRT` / `RdsonDRT2` | 1.5 / 1 Ω（高側驅動上拉 / 下拉） |
| B422 / B423 | `RdsonDRB` / `RdsonDRB2` | 1.5 / 1 Ω |
| B425 | `Deadtime` | 50 ns |
| B427 | `Iq` | 3 mA |
| B430 / D430 | `Rdson1` / `Rdson2` | `=DC!E100/1000`、`=DC!F100/1000` |
| B431 | `RdsonTopTc` | 0.004 |
| B432:D437 | `Qg_Q1/Q2`、`Qgd_Q1/Q2`、`Qgs_Q1/Q2`、`Rgate1/2`、`Gfs_Q1/Q2`、`Vth_Q1/Q2` | 從 DC E101:F106 換算 |
| B438 / D438 | `Vsp_Q1` / `Vsp_Q2` | `=Vth_Q1+Iout/Gfs_Q1`（米勒平台電壓） |
| B440:D443 | `Vsd_Q1/Q2`、`Qrr_Q1/Q2`、`Trr_Q1/Q2`（40 ns 常數）、`ThetaJA_Q1/Q2` | |
| B447:D460 | Q3/Q4 同上（`Rdson3/4`、`RdsonBotTc`、`Qg_Q3/4`…`ThetaJA_Q3/4`） | `Vsp_Q3 = Vth_Q3+Iin_Vinnom/Gfs_Q3` |
| B463:D465 | `VfdSky_Q2/Q3`、`QrrSky_Q2/Q3`、`Ireverse` | |
| B468 | `ThetaJaCtrl` | 30.5 °C/W |
| B471 / D471 | `Tr_buck` / `Tr_boost` | `=(Qgd+Qgs/2)*(RdsonDRT+Rgate)/(Vdd-Vsp)` = 6.25 / 6.28 ns |
| B472 / D472 | `Tf_buck` / `Tf_boost` | `=(Qgd+Qgs/2)*(RdsonDRT2+Rgate)/Vsp` = 12.6 / 12.4 ns |
| B475 / B476 | IC Power Dis / Temp rise | `=#REF!` / `=B475*ThetaJaCtrl` | 失效 |
| B477:B480 | 開關電流、Qrr 電流、二極體電流、總和 | B479/C479 引用了不存在的 `Ipp`，結果是 `#NAME?` |

### 3.9 電路圖標籤（列 486–558）

這些是組好的字串（例如 `="RC1 = "&'Design Converter'!E91&"kΩ"`），給 VM 上的電路圖文字方塊與 DC 的標註用。

| 格 | 內容 | 值 |
|---|---|---|
| A/B487–497 | RC1、RUV1、RUV2、LF、Rdcr、Rs、Rshunt、VOUT、COUT | "1.5kΩ"、"32.4kΩ"… |
| B499 / B503 | `=ROUNDUP(Cout*1e6/22,0)&" x"` / `=ROUNDUP(Cin*1e6/10,0)&" x"` | 6 × 22 µF、10 × 10 µF |
| A/B505–516 | CC1、CC2、gm、RoEA、GBW、Cslope、RFB1、RFB2 | B513 `ROUND(Cslope,0)` 顯示 "0pF"、A513 顯示 "0.0000000000025pF"（單位錯）；A/B/C514 引用了不存在的 `DITHER`，結果是 `#NAME?` |
| D517 / B517 | `=IF(BODEPLOT_TYPE,'IL Stability Calculations'!U4,'Stability Calculations'!Q4)` | 3.98 → "4 kHz" |
| D518 / B518 | `=IF(BODEPLOT_TYPE,ILSTAB!U5,STAB!Q5)` | 71.7 → "72°" |
| B520 | Low line efficiency | `=IF(PLOT_TYPE=1,"",ROUND(EFF_VIN!BO9,1)&"%")` |
| B521 | Full-load/high line efficiency | `=IF(PLOT_TYPE=1, ROUND('Efficiency vs Load Calculations'!BP109,1)&"%", ROUND(EFF_VIN!BO109,1)&"%")` → **"1%"** ⚠ |
| B522 | Half-load/mid-line | `=IF(PLOT_TYPE=1, ROUND(EFF_LOAD!BP59,1)&"%", ROUND(EFF_VIN!BO209,1)&"%")` → **"0.7%"** ⚠ |
| B525:B528, B533:B539 | Rs、Rset、Rshunt、Rt、Css、Cslope、輸出電流、VIN、VIN range | `'Standard Value Calculator'!O4/O6/O13`… |
| A541:B558 | 腳位編號對照（VIN = 12、CBOOT = 18、VDD = 14、SYNC = 6、EN = 7、PGOOD = 11、SS/TRACK = 1、FADJ = 5、GND = 13、COMP = 4、FB = 3、CS− = 24、CS+ = 23、LG = 15、SW = 16、HG = 17、RS = 2、OTP = 8） | 常數 |

> ⚠ **B521 / B522 的錯誤**：在 PLOT_TYPE = 1（vs IOUT）時，這兩格取的是 EFF_LOAD 的 **BP 欄（電感損耗 W）**，不是 **BO 欄（效率 %）**。所以 DC 效率圖上顯示的 "1%"、"0.7%" 其實是電感損耗 1.0 W / 0.7 W。正確的效率在 BO109 = 95.2 %、BO59 = 94.4 %。

---

## 4. DeviceParameter（DP）

**用途**：依 `Device` 選出 IC 參數；用 fSW 查 tON(min)，再算 CS 差動濾波電容。

| 格 | 項目 | 公式 | 值 | 各型號（D = LM5177、E = LM51771、F = LM51770、G = LM517701） |
|---|---|---|---|---|
| B7 | MaxVin | 巢狀 IF：`IF(Device=D6,D7,IF(Device=E6,E7,IF(Device=F6,F7,IF(Device=G6,G7,""))))` | 78 V | 60 / – / 78 / 78 |
| B8 | MinVin | 同上 | 2.9 V | 2.8 / – / 2.9 / 2.9 |
| B9 | MaxfSW | 同上 | 1800 kHz | 600 / – / 1800 / 1800 |
| B11 | g_m（IMON） | 同上 | 1 mS | 全部 1 mS |
| B12 | gm_IL | 同上 | 200 µS | |
| B13 | `gm_CDC`（gm_csa） | 同上 | 600 µS | |
| B16:B18 | CS threshold min/typ/max | 同上 | 42.5 / 50 / 57.5 mV | LM517701：65.625 / 75 / 84.375 mV |
| N6 | t_on_min_Boost | `=VLOOKUP((Fsw/1000000),P7:Q107,2)` | 209.2 ns | 近似比對 |
| N8 | `R_CS` | 常數 10 Ω | | |
| N9 | `C_CS` | `=(N6*1e-9)/(2*PI()*(2*N8)*10)` | 166.5 pF | 濾波時間常數 ≈ tON,min/(2π·10) |
| N10 | `C_CS_ideal` | `='Standard Value Calculator'!O14` | 180 pF | E12 |
| P7:Q107 | tON,min 對 fSW 查表 | 常數 | 0.1 MHz → 435 ns … 2.2 MHz → 86 ns | 101 點 |

B7、B8、B9（MaxVin、MinVin、MaxfSW）在活頁簿裡**沒有被任何公式引用**。輸入範圍檢查是寫死在 DC 的條件格式裡。

---

## 5. CFG Setting

**用途**：LM51770 的 CFG 腳位用一顆電阻設定 4 個功能位元，共 16 種組合（datasheet Table 8-1）。這一頁由使用者選擇算出組合編號與 RCFG。

| 區域 | 內容 |
|---|---|
| A4:A19 | 組合編號 1–16 |
| B4:B19 | RCFG（kΩ）：0、0.511、1.15、1.9、2.7、3.8、5.1、6.5、8.3、10.5、13.3、16.2、20.5、24.9、30.1、36.5 |
| C:F 4–19 | 位元：DRSS（bit 0）、SCP-HICCUP（bit 1）、Current LIMIT（bit 2）、PSM Entry（bit 3） |
| G4:J19 | `=IF(C4=1,"Enable","Disable")`；I 欄是常數 "10%" / "15%" |
| L4:R19 | `=CONCATENATE(G$3,": ",C4)` 等，組成說明字串。P 欄引用空格 `K$3`，結果是 ": Disable" |
| C21:F21 | 每個選項對應的位元權重：C = 0、D = 1、E = 3、F = 2 |
| C22:F25 | 下拉選單清單（DRSS off/on、HICCUP off/on、PSM Entry 10 %/15 %、IMON_LIMIT off / IMON on / ILIMIT positive / ILIMIT negative） |
| C27:F27 | **下拉選單寫入值**：1 / 2 / 1 / 3 |
| C28:E28 | `=(C27-1)*2^C21`… | 0 / 2 / 0 |
| F28 | `=(IF(F27>2,1,0)*2^F21)` | 4：只有 ILIMIT（正或負）才設 bit 2；IMON 不設 |
| **G28** | `=SUM(C28:F28)+1` | **7** = `CFG` |
| H27 / H28 | `=CONCATENATE("R_",A2)` / `=INDEX(B4:B19,G28)` | R_CFG / 5.1 kΩ |

其他分頁的引用：VM O10 `CFG` ← G28；VM B266 `IMONLIM_select` = F27−1；VM B271 `With_RCS` = F27；DC E64 = `INDEX('CFG Setting'!B4:B19,CFG,0)`。

---

## 6. Stability Calculations（STAB）：電壓迴路 Bode 圖

**用途**：用漸近模型（每個極點 / 零點分別算 |·| 與 ∠·）算出電壓迴路的迴路增益、相位、交越頻率與相位裕度，並畫 3 張圖。

### 6.1 參數區（列 2–7）

| 格 | 名稱 | 公式 | 值 |
|---|---|---|---|
| C3 | Afb1 | `=Afb` | 0.0625 |
| C4 | `EA_Zero1` | `=Fz1_u` | 1895 Hz |
| C5 | `EA_Pole1` | `=Fp1_u` | 0.284 Hz |
| C6 | Aea | `=AEA` | 6000 |
| C7 | `EA_Pole2` | `=Fp2_u` | 39.2 kHz |
| F3:F7 | Acs、Am、Q、mc、k | `=Acs` 等 | |
| I3 | gm | `=gm` | |
| K3 | `PS_Zero1` | `=Fesr` | 612 kHz |
| K4 | `PS_Pole1` | `=Fp_dominant` | 1224 Hz |
| K5 | `PS_RHPZ` | `=IF(VIN_nom>Vout,10000000,Frhp)` | 98.8 kHz |
| N3 | `PS_Pole2` | `=Fsw/2` | 199 kHz（取樣雙極點） |
| N4 | `pterm1` | `=Q*PS_Pole2` | |
| N5 | `pterm2` | `=PS_Pole2^2+0.01` | |
| Q3 | `Adc` | `=Am*Afb*AEA` | 21 563（86.7 dB） |
| Q4 | `f_cross_Bode` | `=AJ8/1000` | **3.98 kHz** |
| Q5 | Phase Margin | `=AJ9` | **71.7°** |
| Q6 | `radconv` | `=180/Pi` | |
| S3 | | `=20*LOG(Adc)` | |
| V4:V7 | IL zero / IL pole1 / `AIL` / IL pole2 | `=1/(2*Pi*RcIL*Cc1IL)`… | 給 ILSTAB 用（`AIL` 名稱指向 V6） |
| AJ8 / AJ9 | 系統交越 / PM | `=SUM(AI12:AI613)` / `=SUM(AJ13:AJ613)` | |
| AL8 / AL9 | EA 交越 / PM | `=SUM(AK12:AK613)` / `=SUM(AL13:AL613)` | 3631 Hz / −32.8° |

### 6.2 頻率掃描表（列 13–613，B13 = 1 Hz 到 B613 = 1 MHz，共 601 點）

| 欄 | 公式（第 13 列） | 意義 |
|---|---|---|
| C / D | `=SQRT(1+(B13/PS_Pole1)^2)` / `=ATAN(B13/PS_Pole1)` | 負載極點 |
| E / F | 對 `PS_Zero1` | ESR 零點 |
| G / H | `=IF(VIN_nom>Vout,1,SQRT(1+(B13/PS_RHPZ)^2))` / `=IF(…,0,-ATAN(B13/PS_RHPZ))` | RHP 零點（增益像零點、相位像極點） |
| I / J | 對 `EA_Pole1` | EA 低頻極點 |
| K / L | 對 `EA_Zero1` | 補償零點 |
| M / N | 對 `EA_Pole2` | 補償高頻極點 |
| O~T | 對 `$V$5`、`$V$4`、`$V$7` | IL 放大器極零點（只算不用） |
| U / V | `=SQRT((1-(B^2/pterm2))^2+(B/pterm1)^2)` / 象限修正的 ATAN | Fsw/2 取樣雙極點 |
| **W** | `=20*LOG10(((E*K*G)/(C*I*U*M))*Adc)` | **迴路增益 (dB)** |
| X | `=((F+L+H)-(D+J+V+N))*radconv`（< −180 時加 360） | 迴路相位 |
| Y | `=IF(X>0,X-180,X+180)` | 相位裕度曲線 |
| AB / AC | `=20*LOG10((K/(I*M))*AEA)` / `=(L-(J+N))*radconv` | EA 單獨的 Bode |
| AF / AG | `=20*LOG10((E*G/(C*U))*Acs*Am)` / `=(F+H-(D+V))*radconv` | Plant 單獨的 Bode |
| AI | `=SUM((W14<0)*(W13>0))*B13` | 交越偵測：增益從正變負的那一點填入頻率 |
| AJ | 13–89 列：`=IF(AI13>0,Y11,0)`；90 列之後：`=IF(AI90>0,Y90,0)` | 交越點的相位裕度。⚠ 前 77 列的參考位移了 2 列（交越頻率 < ~6 Hz 時才會用到） |
| AK / AL | 同樣方法套用在 EA 曲線 | |

### 6.3 圖表用的輔助列（列 616–640）

* B617 = `Q4*1000`（交越頻率）。B622:B629 是極零點頻率（`PS_Pole1`、`PS_Zero1`、`EA_Pole1`、`EA_Zero1`、`PS_Pole2`、`Fp2_u`、`PS_RHPZ`），同樣代入公式，給 Bode 圖標記點用。
* C632:E637 整理成「頻率 / 增益」配對，給 Efficiency Plot Management 的 Bode 圖畫標記。
* A639、A640 = `Q4*1000`；AJ687 = `SUM(AJ689:AJ1306)`（空範圍，殘留）。

### 6.4 圖表

| 圖表 | X 軸 | Y 軸 |
|---|---|---|
| "LM5177 Bode Plot" | B13:B613 | W（增益 dB），標題沿用 LM5177 |
| "Error Amp Bode Plot" | B | AB、AC |
| "Plant Bode Plot" | B | AF、AG |

---

## 7. IL Stability Calculations（ILSTAB）：平均電流限制迴路 Bode 圖

**用途**：IMON/ILIM 開啟時，平均電流限制迴路由 gm_IL 放大器、RIMON / CIMON 補償，經過輸出（或輸入）分流電阻 Rshunt 回授。這一頁用 Excel 複數函數（`COMPLEX`、`IMDIV`、`IMPRODUCT`、`IMSUM`、`IMABS`、`IMARGUMENT`）建完整的阻抗模型。

### 7.1 參數區（A2:B34、D2:V7）

| 格 | 名稱 | 公式 | 值 |
|---|---|---|---|
| B2 / B3 | Vout / Vin | `=Vout` / `=Vin` | 16 / 12 |
| B4 | Iload | `=Constant_Current_Loop_Setpoint` | 5 A |
| B5 | Rload | `=Rout` | 2 Ω |
| B6 / B7 | Co1 / Co2 | `=Cout-B7` / `=Co2_u` | 65 / 65 µF |
| B8 / B9 | Co1_esr / Co2_esr | `=CoutEsr` | 2 mΩ |
| B10 | L_dcr | `=Rdcr` | |
| B11 | Duty | `=IF(B3>=Vout,Vout/B3,(Vout-Vin)/Vout)` | 0.25 |
| B12 | | `=IF(Vin<Vout,"Boost","Buck")` | Boost |
| B13 | L | `=Lf` | |
| B15 | Rsns | `=Rshunt/1000` | 10 mΩ |
| B18 | | `=IF(Vin>=Vout,P13,Q13)` | 1 Hz 的複數值（偵錯用） |
| B21:B26 | | 1 Hz 下各阻抗 / 傳遞函數的 dB 值 | 偵錯用 |
| B27 | Gm_csa | `=gm_CDC` | 600 µS |
| B28 | r_v2i_ipeak | 25 000 | 常數 |
| B29 | i_sns_gain | `=Rs*B27*B28` | 0.0195 |
| B30 | Vse | `=IF(Vin>=Vout,MAX(Vout,Vin/2),MAX(Vout-Vin,5))` | 5 |
| B31 / B32 | Sn / Se | `=IF(Vin>=Vout,B29*(Vin-Vout)/Lf,B29*Vin/Lf)` / `=IF(…,B29*B30/Lf,B29*(Vout-Vin)/Lf)` | 1.3e5 / 4.33e4 |
| B33 | mc | `=1+B32/B31` | 1.333 |
| B34 | Fm | `=Fsw/(B31+B32)` | 2.30（調變器增益） |
| E4 / E5 / E6 / E7 | `IL_Zero1` / `IL_Pole1` / AIL / `IL_Pole2` | `=1/(2*Pi*RcIL*Cc1IL)` = 103 kHz / `=1/(2*Pi*(RoEA_IL+RcIL)*Cc1IL)` = 3.6 Hz / `=DP!B12*RoEA_IL` = 4000 / `=1/(2*Pi*RcIL*(Cc2IL+Cbw_IL))` = 14 MHz | |
| H3:H7, K3, O3:O5, R3:R5, U3 | 複製 STAB 的參數 | `Am_IL` = H4、`PS_IL_*`、`pterm*_IL`、`Adc_IL` = U3 | |
| U4 / U5 | 交越 / PM | `=BM8/1000` / `=BM9` | **2.29 kHz / 101.9°** |
| U6 / U7 | `radconv_IL` / Adc dB | | |
| BB6 / BC6 | 1 Hz 的 IL 放大器增益與相位 | | |
| BM8 / BM9 | | `=SUM(BL12:BL613)` / `=SUM(BM13:BM613)` | |
| BO8 / BO9 | EA 交越 / PM | | 144.5 Hz |

### 7.2 頻率掃描表（D13:D613 = 1 Hz~1 MHz）

| 欄 | 標題 | 公式（第 13 列） | 意義 |
|---|---|---|---|
| E | ZL | `=COMPLEX(0,2*PI()*D13*Lf)` | 電感阻抗 |
| F / G | ZCo1 / ZCo2 | `=IMSUM(COMPLEX($B$8,0),IMDIV(1,COMPLEX(0,2πf·$B$6)))` | 兩組輸出電容（含 ESR） |
| H | Zout | `1/(1/Rload + 1/ZCo2)` | |
| I | Zout_sns | `=IMSUM(H13,COMPLEX($B$15,0))` | 加上分流電阻 |
| J | Zout_all | `1/(1/ZCo1 + 1/(Rsns+Zout))` | |
| K | He | `1 + s/(−Fsw/2·…) + s²/(π·Fsw)²` | 取樣保持模型 |
| M | G_d_il | Buck：`Vin/(ZL+Zall)`；Boost：`(Vout+Zall·Iload)/(ZL+Zall·(1−D)²)` | duty→電感電流 |
| N | G_d_iout | Boost：`G_d_il·(1−D) − Iload/(1−D)` | |
| O | G_d_isns | `G_d_iout / (Zout_sns·(1/ZCo1+1/Zout_sns))` | |
| P / Q | G_d_mid / G_d_vo | | duty→中間點 / 輸出電壓 |
| R | G_d_il_smpl | `=IMPRODUCT(M13,$B$29,K13,$B$34)` | 內電流迴路增益 Ti |
| S / T / U | G_vc_Il / G_vc_isns / G_vc_vo | `=Fm·G/(1+Ti)` | 控制電壓→各量（閉內迴路） |
| V~AH | ADC_EA、W/Z/AC（p1/z1/p2 的複數、|·|、∠）、AF = 總 EA | | 電壓 EA（只算不用） |
| AI~AU | ADC_IL = `COMPLEX(AIL,0)`、AJ/AM/AP（IL 的 p1/z1/p2）、AS = 總 IL 放大器 | | |
| **AV** | `=IMPRODUCT(T13,AS13,COMPLEX($B$15,0))` | | **迴路 = G_vc_isns × IL 放大器 × Rsns** |
| AW / AX / AY | 增益 dB、相位、相位裕度 | | |
| BB / BC | `20log|Rsns·AS|`、相位 | | IL 放大器 Bode |
| BE / BF | `20log|T|`、相位 | | Plant Bode。**BE13 被 VM B292 `C_IL_recommend` 引用** |
| BL / BM | 交越偵測 / PM（與 STAB 相同手法，BM13:BM89 同樣有 2 列位移） | | |
| BN / BO | IL 放大器的交越 / PM | | |

列 617–640 是圖表輔助列（同 STAB）。

### 7.3 圖表

"LM5177 Bode Plot"（D vs AW）、"I Limit Amp Bode Plot"（BB、BC）、"Plant Bode Plot"（BE、BF）。

---

## 8. Schematic Management

| 格 / 物件 | 公式 / 內容 |
|---|---|
| A2 | `=CHOOSE(With_RCS,"SCH_NO_RCS","SCH_RCS_MON","SCH_RCS_LIM","SCH_RCS_LIM_NEG")` → "SCH_RCS_LIM" |
| Object 49–52 | 4 張內嵌電路圖，分別位在具名範圍 `SCH_NO_RCS`(B5)、`SCH_RCS_MON`(D5)、`SCH_RCS_LIM`(F5)、`SCH_RCS_LIM_NEG`(H5) |
| `PICTURE2` | `=INDIRECT('Schematic Management'!$A$2)`，DC 上的連結圖片用它動態顯示對應電路圖 |
| `SCH772_*` | 指向第 11 列（空），應是 LM51772 版本留下來的 |

---

## 9. Efficiency Plot Details

沒有儲存格內容，只有 14 張折線圖（效率 vs IOUT、vs VIN，以及四顆 MOSFET 的損耗）。資料來源**全部是 `#REF!`**，是舊版遺留物。另有工作表層級的同名名稱（`EFF_IOUT`、`EFF_VIN`、`BodePlot`…）指向本頁，但實際使用的是活頁簿層級、指向 Efficiency Plot Management 的版本。

---

## 10. Efficiency Plot Management

| 格 / 物件 | 公式 / 內容 |
|---|---|
| A2 | `=IF(PLOT_TYPE=1,"EFF_IOUT","EFF_VIN")` → `PICTURE` 名稱用 INDIRECT 指向這格 |
| A9 | `=IF(OR(BODEPLOT_TYPE=FALSE,IMONLIM_select=0),"BodePlot","BodePlot_ILim")` → `Picture_Bode` |
| 名稱 `EFF_IOUT`(B5)、`EFF_VIN`(B7)、`BodePlot`(B10)、`BodePlot_ILIM`(B12) | 圖表所在的儲存格區域（照相機來源） |
| chart21 | 效率 vs IOUT：`EFF_LOAD!K9:K109` 對 BO（效率）、BP（電感損耗）、BV（分流損耗）、BR/BS（MOSFET 損耗） |
| chart22 | 效率 vs VIN：`EFF_VIN!H9:H209` 對 BO、BP、BV、BR、BS… |
| chart23 | 電壓迴路 Bode：`STAB!B13:B613` vs `W13:W613`，加上 C/D632–637、B627/W627 等極零點標記 |
| chart24 | IL 迴路 Bode：`STAB!B13:B613` vs `ILSTAB!AW13:AW613`，加上標記 |
| 文字方塊 "92.4%"、"92.5%"、核取方塊 "Show I_Limit Gain" | 核取方塊連到 `'[1]Variable Management'!$L$11`。**`[1]` 是外部活頁簿參照**（從別的檔案複製過來的殘留），這顆核取方塊實際上沒有作用；DC 上的那顆才有效 |

---

## 11. Timings

**用途**：依照 LM51770 的 buck-boost 時序（Phase I：Q1+Q3、Phase II：Q1+Q4、Phase III：Q2+Q4）算出 Buck↔BB↔Boost 的切換電壓，以及 VIN(nom) 下三個相位的時間與電感電流波形。

### 11.1 輸入（A1:B10）

| 格 | 名稱 | 公式 / 值 |
|---|---|---|
| B1 | Device | 1（"1:LM5177 / 2:LM51772"，寫死） |
| B3 | Fsw [kHz] | `=Fsw/1000` |
| B4 | IL [A] | 0 |
| B5 | Vin | `=VIN_nom` |
| B6 | Vout | `=Vout` |
| B7 | Iout | `=Iout` |
| B8 | `Iin` | `=Iin_Vinnom` |
| B9 | `eff` | `=Efficiency` |
| B10 | `L` | `=Lf` |

### 11.2 IC 時序參數（G1:L6、I 欄依 B1 選 G 或 H）

| 格 | 公式 | 值 | 推測意義 |
|---|---|---|---|
| I1 / L1 | `=IF($B$1=2,H1,G1)` | 115 ns | 轉換時間常數 |
| I2:I5 | 同上 | 2 / 1 / 2 / 1 | 型號相關係數 |
| J3, J4 | 1e-6 | | |
| J1, J2 | （空白 = 0） | | |
| J5 / J6 / J7 | `=IF(B4>0,1,0)` / `=IF(B4<0,1,0)` / `=J5-J6` | 0 | 電流方向旗標 |
| L3 | `=0.225-I2*0.025` | 0.175 | |
| L4 / L5 / L6 | `=0.75+I3*0.25` / `=1-I4*0.0416666` / `=0.75+I5*0.25` | 1 / 0.917 / 1 | |
| N1 | `=0.75/$L$1*$J$3/2000000*0.2` | 652 ns | |
| N2 | `=MIN($J$3/2000000*$B$3*1000*$L$5+$J$4*(1-$L$5),$N$1)` | 266 ns | |
| N4 | `=$L$3/$B$3*0.001` | 440 ns | 0.175·T |
| N5 | `=$L$4*0.000000075/$N$2*0.000001` | 282 ns | |
| N6 | `=N4+2*N5` | 1004 ns | |

### 11.3 轉換點（P/R/T 欄）

| 格 | 公式 | 值 | 意義 |
|---|---|---|---|
| P2 / P3 / P4 | `=$N$5-$J$7*$J$1` / `=$N$5+$J$7*$J$1` / `=0.001/$B$3-$P$3` | | Buck→BB 的最小 off-time |
| **P9** | `=$B$6/($P$4*1000*$B$3)` | **18.02 V** | Buck → BuBo（`VIN_buck_to_BB`） |
| R2:R4, **R8** | `=$B$6*(1-$R$2*1000*$B$3)` | **14.20 V** | Boost → BuBo（`VIN_boost_to_BB`） |
| T2 / T3 / T4 / T5 / T6 | `=MAX(MAX(0.5*N5/L4*L6, J6*J2)-J7*J1,0)`… / `=N4+2*N5` / `=0.001/B3-T5` | | |
| **T8** | `=$B$6*((0.001/$B$3-$T$4)/($T$6+$T$4))` | **11.13 V** | BuBo → Boost（`VIN_BB_to_boost`） |
| **T9** | `=$B$6*((0.001/$B$3-$T$2)/($T$6+$T$2))` | **23.00 V** | BuBo → Buck（`VIN_BB_to_buck`） |
| F2:F5 | `=$P$9`、`=$T$9`、`=$R$8`、`=$T$8` | | 整理後給 VM B44:B47 |

### 11.4 相位時間與 duty（E8:H17）

| 列 | Phase II (Q1-Q4) | Phase I (Q1-Q3) | Phase III (Q2-Q4) |
|---|---|---|---|
| 9 BuBo→Buck [ns] | `=T6*1e9` = 1508 | `=T2*1e9` = 141 | `=(T5-T2)*1e9` = 863 |
| 10 BuBo→Boost | `=T6*1e9` | `=T4*1e9` = 863 | `=(T5-T4)*1e9` = 141 |
| 11 BuBo center | `=T6*1e9` | `=N6*1e9*L11` = 660 | `=N6*1e9*(1-L11)` = 344 |
| 15–17 duty [%] | `=F9*$B$3*1e-6*100`… | | |

L11 = 0.657 是 "Simplis Align factor"（K12 註記 default 0.5）。

### 11.5 三相電流波形（V14:Y18、X1:AC11）

| 名稱 | 格 | 公式 | 值 |
|---|---|---|---|
| `t1_Ph_II` | W15 | `=$F$11*1e-9` | 1.508 µs |
| `t2_Ph_I` | W16 | `=(Vout/Fsw-Vin*W15)/(Vin+Vout)` | 0.789 µs（伏秒平衡） |
| `t3_Ph_III` | W17 | `=1/Fsw-(W15+W16)` | 0.215 µs |
| | W18 | `=1/Fsw` | |
| | Y1 / Y2 / Y3 | `=1/2*Fsw/L*((W16)^2*(Vin-Vout)+(W15)^2*Vin+(W17)^2*Vout)` 等 | 平均漣波修正 |
| `iavg` | Y5 | `=IF(B5>B6,Iin+Y1,Iout+Y1)` | 10.82 A |
| `dI1_` / `dI2_` / `dI3_` | Y6 / Y7 / Y8 | `=Vin/L*W16`、`=(Vin-Vout)/L*W15`、`=Vout/L*W17` | 5.26 / −3.35 / 1.91 A |
| `I_0` / `I_1` / `I_2` | Y11 / Y12 / Y13 | `=IF(Vin>B6,0,Y6)+Iout-Y2`… / `=Y11+Y8` | 13.46 / 8.19 / 15.37 A |
| | AA11 | `=1/(W15+W16)*(Iin/Fsw-…)` | 9.32（另一種算法） |
| `I_PP` | Y14 | `=ABS(MAX(Y11:Y13))-ABS(MIN(Y11:Y13))` | 7.17 A |
| `i_rms` | Y16 | `=SQRT(Fsw*(M75*(I_0*I_1+((I_1-I_0)^2)/3)+M76*(…)+M77*(…)))` | 11.60 A |

### 11.6 其他區域

* J29:R59：三組 SIMPLIS 模擬的 t0~t3 時間戳（K 欄常數）與 L/M 欄計算值的比較（"Simplis / Ref - Calc"），用來驗證模型。D45:D46 記錄 EVM 用的電感 XGL1060-332MEC。
* F69:F96：伏秒平衡與 RMS 公式的推導文字。M75:M78 是 t1/t2/t3/T（Y16 引用）。
* AA24:AG30：「Operation mode ranges」堆疊長條圖的資料（VIN_min → BB→Boost → Boost→BB → Buck→BB → BB→Buck → VIN_max 的各段寬度）。

---

## 12. Efficiency vs VIN Calculations（Excel 表格 `Table16`）

**用途**：在 IOUT = Iout（滿載）下，把 VIN 從 VIN_min 到 VIN_max 分成 200 等分（201 點），逐點算出各元件損耗與效率。

### 12.1 左側輸入區（A9:D80）

全部是把 VM 的值換成顯示單位，例如 `B11 =Vin`、`B13 =Fsw/1000`、`B18 =Lf*1000000`、`B37 =Rdson1*1000`、`B39 =Qg_Q1*1e9`…（共 47 格，涵蓋 Vout、Iout、Ta、Vdd、L、DCR、Ltc、θCA、Cin、ESR、驅動電阻、Deadtime、Iq、Q1/Q2 全部參數、Schottky）。B31（RdsonDRB2 = 1）、B62（θJA 控制器 30）、B63（2）是常數。

輸出區：B68 `=#REF!`（Dmax 失效）、B69 `=1-B68`、B70 `=Iripple_Vinnom`、B72/B73 = Tr/Tf（ns）、B75 `=BH9`（IC 損耗）、B76 `=B75*B62`、B77:B80（B79/C79 引用了不存在的 `VsdTop`、`VsdBot`，結果是 `#NAME?`）。

H2 / H3 / H4：`=BK9` / `=BK109` / `=BK209` = 90.3 % / 93.8 % / 91.6 %，標籤寫 Low / Nom / High VIN。⚠ 第 109 列是 step 100，VIN = (6+36)/2 = **21 V**，不是 VIN(nom) = 12 V。

### 12.2 掃描表欄位（第 9–209 列，以結構化參照寫成）

| 欄 | 欄名 | 公式重點 |
|---|---|---|
| G | Step | 0…200（常數） |
| H | Vin | `=VIN_min+(VIN_max-VIN_min)*(Step/200)` |
| I | Mode | `=IF(Mode=2, IF(Vin>Vout,"BUCK","BOOST"), IF(Vin>VIN_BB_to_buck,"BUCK", IF(Vin<=VIN_boost_to_BB,"BOOST","BUCK-BST")))`。⚠ `Mode` 只會是 0 或 1，所以 `Mode=2` 分支永遠走不到 |
| J | Mode Digit | BUCK = 1、BOOST = 2、BUCK-BST = 3 |
| K | Iout | `=Iout` |
| L | Iin | `=Vout*Iout/(Vin*Efficiency)` |
| M | t1_Ph_II | `MIN(1/Fsw, CHOOSE(digit, Buck 含壓降 duty/Fsw, Boost 含壓降 duty/Fsw, t1_Ph_II))`。BB 模式用的是 **Timings 在 VIN(nom) 算出的固定 t1** |
| N | t2_Ph_I | Buck/Boost：`1/Fsw − t1`；BB：`(Vout/Fsw − Vin·t1_Ph_II)/(Vin+Vout)` |
| O | t3_Ph_III | `1/Fsw − (t1+t2)` |
| P / Q / R | I_0 / I_1 / I_2 | 三相電流端點（Buck：`Iout ∓ ΔI/2`；Boost：`Iin ∓ ΔI/2`；BB：由平均電流與三段斜率反推） |
| S | Avg IL | `Fsw·(I0·t1 + I1·t2 + I2·t3)` |
| T / U | Duty BUCK / BOOST Leg | |
| V | Ripple pk-pk | `MAX(I0:I2) − MIN(I0:I1)` |
| W | Irms Lf | 梯形 RMS：`√(Fsw·Σ tk·(Ia·Ib + (Ia−Ib)²/3))` |
| X | DCR Loss | `I²·Rdcr·(1+Ltc·(Ta−25)) / (1 − I²·Rdcr·Ltc·θCA)`（含自熱的封閉解） |
| Y | Core Loss | `ABS(7.759e-3·Fsw^0.9458·(0.00787·ΔI)^2.304)`（固定的 Steinmetz 係數） |
| Z | Total Inductor Loss | X + Y |
| AA / AB / AC | Irms Cout / ESR Loss / Total | Buck：`ΔI/√12`；Boost：`√(D/(1−D))·√(Iout²+ΔI²/12)`；`RCoutEsr·I²` |
| AD / AE / AF | Irms Cin / ESR Loss / Total | Buck：`√(D(1−D))·√(Iout²+ΔI²/12)`；Boost：`ΔI/√12` |
| AG~AK | Q1（Buck 控制管）：Irms、`Rdson1·I²`、轉換損耗 `Vin·I/2·Fsw·(Tr_buck+Tf_buck)`、死區損耗、Total（含溫度係數自熱修正，用 `RdsonTopTc`、`ThetaJA_Q1`） | |
| AL~AO | Q2（Buck 同步管）：Irms、`Rdson2·I²`、死區體二極體損耗（有 Schottky 時為 0）、Total | |
| AP~AT | Q4（Boost 控制管）：Irms、`Rdson4·I²`、`Vout·Iin/2·Fsw·(Tr_boost+Tf_boost)`、死區、Total | |
| AU~AX | Q3（Boost 同步管）：Irms、`Rdson3·I²`、死區、Total | |
| AY | Q1/Q3 RDS(on) Loss @ D = 100 % | 只算不加總 |
| AZ | Deadtime Loss ext D | 外接 Schottky 的死區導通損耗 |
| BA / BB | Reverse Recovery Loss BUCK / BOOST | `Qrr·Fsw·V` |
| BC | Isense Shunt Loss | `Rs·(I_BuSync² 或 I_BoCtrl²)` |
| BD | Iout Shunt Loss | `Rshunt/1000·Iout²` |
| BE | IQ Loss | `Iq·(Vout if Vout>8 else Vin)` |
| BF | Switching Current Loss | `ΣQg·Fsw·(Vout or Vin)` |
| BG | Switching Current Diss | 驅動器內部損耗（不加總） |
| BH | Total IC Loss | BE + BF |
| BI | Total Loss | Z + AC + AF + AK + AO + AT + AX + AZ + BA + BB + BH + BC + BD |
| BJ | Pout | `Vout·Iout` |
| **BK** | **EFF, η** | `Pout/(Pout+Total Loss)` |
| BL / BM | Control / Sync FET Temp Rise | `θJA·Loss` |
| BO | η (%) | `BK·100` |
| BP | Inductor Cu & Core Loss | `=Z` |
| BQ | Total MOSFET Loss | |
| BR / BS / BT / BU | 各 MOSFET 損耗（反向恢復損耗分 0.66 / 0.33 給控制管 / 同步管） | 給圖表用 |
| BV | Shunt Loss | `=BC` |
| CC:CN 6–17 | 驅動器在不同 θJA 下的溫升（CE6:CN6 = 80…50） | CC 欄是空的，結果無意義，殘留 |

### 12.3 圖表

"Overall Efficiency"（H vs BK）、"MOSFET & Inductor Power Dissipation"（AK、AO、Z）、"MOSFET Temp Rise"（BL、BM）、未命名圖（T、U 兩臂 duty）。

---

## 13. Efficiency vs Load Calculations（`Table25`）

和第 12 頁同一套模型，差別如下：

| 項目 | vs VIN | vs Load |
|---|---|---|
| 掃描變數 | H = Vin（201 點） | G = 0…100（%）；H = `Vin`（固定 VIN(nom)）；K10:K109 = `Iout*G/100`（K9 = 0） |
| M（t1）| 有 `MIN(1/Fsw,…)` 上限 | 沒有上限 |
| T / U | 有 `MIN(1,…)` | 沒有 |
| AG / AL / AP / AU | 每一列都用完整的三相 RMS | **只有第 9 列用完整公式**，第 10–109 列直接 `=[Irms Lf]`（簡化，Buck/Boost 時會高估開關管電流） |
| BA / BB | | Qrr 按負載比例縮放 `Qrr_Q2·Iout_row/Iout` |
| BF | `×(Vout or Vin)` | `×MIN(Vin,Vout)` |
| BL6 / BM6 | | 動態標題 `=IF(Vin>Vout,"Buck ''CONTROL''…","Boost ''CONTROL''…")` |
| H2 / H3 / H4 | | `=BK19` / `=BK59` / `=BK109` = 10 % / 50 % / 100 % 負載效率 81.9 / 94.4 / 95.2 % |

圖表："Overall Efficiency"（K vs BK）、"MOSFET & Inductor Power Dissipation"（BR、BS、Z）、"MOSFET Temp Rise"（BL、BM）。

---

## 14. Standard Value Calculator（SVC）

### 14.1 電容：E12 系列（A6:A159、B~J 欄）

* A6:A33 = E12 基本值 0.47、0.56、0.68、0.82、1、1.2、1.5…82 pF；A34:A159 = `=A22*10`…，每個十進位遞推到 2.7e12 pF。
* 每欄第 3 列是輸入（理想值，F），第 6–159 列用同一套公式找最接近的 E12 值：

  `=IF( IF((X$3·1e12−$A7)·(X$3·1e12−$A6)·−1<0,0,1) · IF(ABS(X$3·1e12−$A7)>(X$3·1e12−$A6),$A6,$A7) = X7, 0, 同上 )`

  只有理想值落在 [A_n, A_n+1] 區間的那一列，會填入比較近的那個標準值，其他列為 0。第 4 列再用 `SUM(…)/1e12` 換回 F。

| 欄 | 輸入（第 3 列） | 結果（第 4 列） | 用途 |
|---|---|---|---|
| B | `=Css` | 18 nF | `Css_u` |
| C | `=Cc1ea` | `SUM/1e15`（O28 再 ×1e12 = 56，單位 nF） | CCOMP |
| D | `=Cc2ea` | 2700 pF | CHF |
| E / F | `=Cout` / `=Cin` | | 沒有被引用 |
| G | `=N14`（C_CS） | 180 pF | `C_CS_ideal` |
| H | H2 `=M19`、H3 `=N19`（Csby） | 0 | 沒有使用 |
| I | `=C_IL_recommend` | 1.5 nF | CIMON |
| J | `=Csby` | 0 | 沒有使用 |

### 14.2 電阻：E96 系列（M/N/O 欄）

O 欄公式（以 O4 為例）是一條很長的 E96 四捨五入式：

```
n   = ROUND(96·frac(log10 N))，若 96·frac − n < 0 則 n−1   → 下方 E96 索引
Rlo = INT(0.5+100·10^(n/96))·10^(INT(log10 N)−2)
Rhi = INT(0.5+100·10^((n+1)/96))·10^(INT(log10 N)−2)
O   = IF(N > (Rlo+Rhi)/2, Rhi, Rlo)
```

| 列 | M 名稱 | N 輸入 | O 標準值 | 被誰使用 |
|---|---|---|---|---|
| 4 | Rt | `=Rfadj` 75 525 | 75 000 Ω | DC E12、VM B533 |
| 6 | Rshunt | `=Rshunt` 10 | 10 | VM B525/B527 |
| 7 | Rc1ea | `=Rcea1` 1514 | 1500 | O27 → DC C91 |
| 9 | Rfb1 | `=Rfb1_` | 71 500 | |
| 10 | Rfb2 | `=Rfb2_` 4767 | 4750（允許 "OPEN"） | VM B398 |
| 13 | Rs | `=Rs_ideal*1000` | 1.65 mΩ | VM B526 |
| 14 | Ccsa_ideal | `=C_CS` | `=G4` 180 pF | DP N10 |
| 15 | Rslope_ideal | `=Rslope_ideal` 69 231 | 69 800 | VM B165 |
| 16 / 17 | Ruv1 / Ruv2 | `=Ruvlo1` / `=Ruvlo2` | 32.4 / 13.3 kΩ | VM E241/E242 |
| 19 | free | `=Csby` | `=H4` | |
| 21 | Rshunt | `=Rshunt` | 10 mΩ | 名稱 `Rshunt_std`（沒有被引用） |
| 23 | C_IL_recommend | | `=I4` 1.5 nF | VM F292 |
| 24 | R_IL_recommend | 477 | 475 Ω | VM F293 |
| 27 | Rc1 | `=O7/1000&" kΩ"` | `=O7/1000` = 1.5 | DC C91 |
| 28 | Cc1 | `=C4*1e9&" nF"`（顯示 "0.056 nF"，標籤錯） | `=C4*1e12` = 56 | DC C92 |
| 29 | Cc2 | `=D4*1e12&" pF"` | `=D4*1e12` = 2700 | DC C93 |

---

## 15. Snubber（獨立小工具）

| 格 | 項目 | 公式 | 值 |
|---|---|---|---|
| B2 / D2 | 振鈴頻率 | 85 MHz / `=B2*1e6` | |
| B3 / D3 | 開關節點電容 | 3000 pF / `=B3*1e-12` | |
| B4 / D4 | Lpar | `=1e9/((D2·2·3.14159)²·D3)` | 1.17 nH |
| B5 / D5 | Csnub | `=4*B3` | 12 nF |
| B6 / D6 | Rsnub | `=SQRT(D4/D3)` | 0.624 Ω |

不引用其他分頁，也不被其他分頁引用。

---

## 16. Parallel Output Caps（獨立小工具）

兩組並聯電容（R1/C1 = 0.5 mΩ / 20 µF，R2/C2 = 20 mΩ / 68 µF）的漣波電流分流比例：

| 格 | 公式 | 值 |
|---|---|---|
| B3 | `=Iripple1` | `#NAME?`（名稱不存在） |
| B4 | `=Fsw` | 398 kHz |
| D6/D7/D9/D10 | 單位換算 | |
| B12 | `=(D7/(D7+D10))*SQRT((D9*D10*B4)^2+1)/SQRT((((D6+D9)*D7*D10/(D10+D7))*B4)^2+1)` | I1 佔 25.6 % |
| B13 | `=1-B12` | I2 佔 74.4 % |
| B15 / B16 | `=B12*(Iripple1/(2*SQRT(3)))` | `#NAME?` |

---

## 17. 發現的問題與注意事項

| # | 位置 | 問題 | 影響 |
|---|---|---|---|
| 1 | VM B521、B522 | PLOT_TYPE = 1 時取的是 EFF_LOAD 的 **BP**（電感損耗 W），不是 **BO**（效率 %） | DC 效率圖顯示 "1%"、"0.7%"，應為 95.2 %、94.4 % |
| 2 | EFF_VIN H3 | 標籤寫 "Nom VIN"，但取的是 step 100 = 21 V | 不是 VIN(nom) 的效率 |
| 3 | VM 列 31–43 | 7 個具名範圍與 `Don_buckboost`(=B37) 指向空格 | Buck-Boost 邊緣的漣波與峰值電流都算成 0 漣波 |
| 4 | VM B113 | Boost 分支用了 `Iin_Vinmin` | 只有 VIN_max < Vout 時才有影響 |
| 5 | VM B207 | Buck 分支用 `Don_Vinmin_2*(1-Don_Vinmax_2)` | 只有 VIN_min > Vout 時才有影響 |
| 6 | VM B285 / B286 | IL 迴路極點用了電壓 EA 的 `RoEA` / `Cbw` | `Fp_IL` 與 `Fp_ILHF` 數值不正確（但下游沒有使用） |
| 7 | STAB / ILSTAB AJ13:AJ89、BM13:BM89 | 相位裕度取了往上 2 列的值 | 只有交越頻率 < 約 6 Hz 時才會出錯 |
| 8 | EFF_VIN / EFF_LOAD I 欄 | `Mode = 2` 的條件永遠不成立 | 一律用 Timings 的 BB 轉換點判斷模式 |
| 9 | EFF 兩頁 BB 模式 | t1 用 Timings 在 **VIN(nom)** 算出的固定值 | 掃 VIN 時，BB 區的損耗只是近似 |
| 10 | EFF_LOAD AG/AL/AP/AU 第 10 列之後 | 開關管 RMS 一律等於電感 RMS | Buck/Boost 模式會高估開關管損耗 |
| 11 | 失效公式 | `#REF!`：VM B475、EFF 兩頁 B68、Efficiency Plot Details 的 14 張圖。`#NAME?`：`Iripple1`、`RdsonTop`、`RdsonBot`、`DITHER`、`Ipp`、`VsdTop`、`VsdBot` | 都在沒有被使用的支線上，不影響主要結果 |
| 12 | 外部連結 | Efficiency Plot Management 的核取方塊連到 `'[1]Variable Management'!$L$11` | 這顆核取方塊無效 |
| 13 | 巨集 | 圖片指定了 `OpenLM27403ProductFolder`，但 VBA 中沒有這個巨集 | 點了會出錯 |
| 14 | DC E35 | 「選用 RSLOPE」是公式，不是輸入格 | 想自訂得直接覆寫公式 |
| 15 | 名稱共用 | `Vin` 與 `VIN_nom` 是同一格 | 補償與 Bode 只在 VIN(nom) 計算，所以 DC 的註解提醒要掃 VIN(nom) 並先按 FIX VALUES |
| 16 | 雜項 | VM B16 註解寫 90 %，實際 95 %；SVC N28 顯示 "0.056 nF"；VM A513/B513 Cslope 單位錯；圖表標題寫 "LM5177 Bode Plot"；Timings B1 寫死 LM5177 | 只是顯示問題 |
| 17 | 圖檔 | 電路圖 IC 標成 LM5177、L 上限係數圖寫 25 但公式用 35、VM 有過時的公式截圖、Timings 有 NDA 浮水印的 datasheet 截圖 | 詳見 §18.5 |

---

## 18. 圖檔分析

xlsm 裡共有 **29 個媒體檔**（`xl/media/`）與 **4 個內嵌 Visio 檔**（`xl/embeddings/*.vsdx`）。去掉 3 組重複（image1 = image5、image2 = image6、image4 = image7，是同一張圖的 DrawingML 版與 VML 版），實際上是 **26 張不同的圖**。EMF 向量圖已經轉成 PNG，全部放在專案的 `xlsm_images/` 資料夾。

下表的「位置」是圖片左上角錨點所在的儲存格（由 drawing XML 的 `from` / `to` 換算）。

### 18.1 總表

| 檔案（xlsm_images/） | 原始檔 | 分頁 / 物件 / 位置 | 類型 | 內容 | 和公式 / 變數的關係 |
|---|---|---|---|---|---|
| DC_01_schematic_snapshot_PICTURE2.png | image4/7.emf | DC Picture 21，G8:Q35 | 連結圖片（照相機）的快取畫面 | 完整應用電路圖（ILIM 正版本） | 公式 `=PICTURE2` → `INDIRECT(Schematic Management!A2)` → 依 `With_RCS` 從 18.3 的 4 張圖選一張 |
| DC_02_bode_snapshot_Picture_Bode.png | image2/6.emf | DC Picture 28，G58:Q86 | 同上 | 電壓迴路 Bode 圖 | `=Picture_Bode` → EPM!A9 → `BodePlot` / `BodePlot_ILim` |
| DC_03_efficiency_snapshot_PICTURE.png | image1/5.emf | DC Picture 30，H96:Q124 | 同上 | 效率 vs 負載電流圖 | `=PICTURE` → EPM!A2 → `EFF_IOUT` / `EFF_VIN` |
| DC_04_WideVin_logo.jpg | image3.jpg | DC Picture 68，A1:D1 | 靜態圖片 | "WIDEVIN DC/DC Power Solutions" 標誌（National Semiconductor 時期的品牌） | 無 |
| DC_05_print_header_National_logo.jpeg | image8.jpeg | DC 頁首（VML `RH` = 右側頁首） | 列印頁首圖片 | National Semiconductor 標誌（模糊浮水印風格） | 只在列印時出現 |
| LIC_01_TI_logo.gif | image9.gif | Licenses Picture 1，A1:B2 | 靜態圖片 | Texas Instruments 標誌（descr = "ti logo"） | 無 |
| VM_01_error_amp_compensation.png | image10.png | VM Picture 11（Group 14 的一部分），O25:Y61 | 靜態圖片（descr = "Error amp - no ref des.tif"） | 跨導 EA + Type II 補償（RCOMP 串 CCOMP，並聯 CHF）+ FB 分壓電路 | 和同一群組的 8 個文字方塊（CC1 = 56nF、RFB1 = 71.5kΩ、RFB2 = 4.75kΩ、RC1 = 1.5kΩ、VOUT = 16V、gm = 0.6mS、RoEA = 10MΩ、GBW = 13MHz）疊在一起。文字方塊內容來自 VM A505:B516，所以這張圖會「顯示目前的補償元件值」 |
| VM_02_eq_9p5V.png | image11.png | VM Picture 12，C40:F44 | 公式截圖 | `VOUT / ((2Tsw−TOFFmin)/2Tsw · 2Tsw/(2Tsw−TON)) = 9.5 V` | 見 18.2 |
| VM_03_eq_10p56V.png | image12.png | VM Picture 14，G40:H44 | 公式截圖 | `VOUT / (Tsw/(Tsw−TONmin)) = 10.56 V` | 見 18.2 |
| VM_04_eq_13p043V.png | image13.png | VM Picture 15，C44:E48 | 公式截圖 | `VOUT / ((Tsw−TOFFmin)/Tsw) = 13.043 V` | 見 18.2 |
| VM_05_eq_14p842V.png | image14.png | VM Picture 16，E44:H47 | 公式截圖 | `VOUT / ((2Tsw−TOFF)/2Tsw · 2Tsw/(2Tsw−TONmin)) = 14.842 V` | 見 18.2 |
| VM_06_Lmin_Lmax_limits.png | image15.png | VM Picture 2，R64:X82（"Joerg's Calculation" 區旁邊） | 簡報截圖 | "Minimum and maximum L" 四條限制 | 見 18.2，和 L64、O64、O65、O66 一一對應 |
| SCH_B5_SCH_NO_RCS.png | image19.emf + Microsoft_Visio_Drawing3.vsdx | Schematic Management，B5 | 內嵌 Visio 物件 | 沒有電流感測電阻的電路 | 名稱 `SCH_NO_RCS`；`With_RCS` = 1 |
| SCH_D5_SCH_RCS_MON.png | image16.emf + Microsoft_Visio_Drawing.vsdx | Schematic Management，D5 | 內嵌 Visio 物件 | IMON 監測 | `SCH_RCS_MON`；`With_RCS` = 2 |
| SCH_F5_SCH_RCS_LIM.png | image17.emf + Microsoft_Visio_Drawing1.vsdx | Schematic Management，F5 | 內嵌 Visio 物件 | 平均電流限制，正（SYNC 上拉） | `SCH_RCS_LIM`；`With_RCS` = 3（**目前選用**） |
| SCH_H5_SCH_RCS_LIM_NEG.png | image18.emf + Microsoft_Visio_Drawing2.vsdx | Schematic Management，H5 | 內嵌 Visio 物件 | 平均電流限制，負（SYNC 下拉） | `SCH_RCS_LIM_NEG`；`With_RCS` = 4 |
| TIM_01_buckboost_phases_fig8-4.png | image24.png | Timings Picture 5，C22:I33 | datasheet 截圖（有 "NDA Restricted" 浮水印） | 四開關 buck-boost 的三種導通路徑：II = Q1+Q4（綠）、I = Q1+Q3（紅）、III = Q2+Q4（藍） | 定義了 `t1_Ph_II`、`t2_Ph_I`、`t3_Ph_III`（W15:W17）以及 EFF 表的 M/N/O 欄 |
| TIM_02_inductor_current_fig8-5.png | image25.png | Timings Picture 6，C34:I42 | datasheet 截圖（NDA 浮水印），"Figure 8-5 Inductor Current in Continuous Buck-Boost Operation" | 時脈 → II（平）→ I（上升到 IL(peak)）→ III（下降）的電感電流波形 | 對應 `I_0`、`I_1`、`I_2`（Y11:Y13）與 `dI1_`~`dI3_` |
| TIM_03_XGL1060-332MEC_L_vs_I.png | image29.png | Timings Picture 10，D46:I63（D45:D46 文字寫 "EVM / XGL1060-332MEC"） | 電感 datasheet 截圖 | Coilcraft XGL1060-332MEC 的電感值對電流曲線：0 A 約 3.3 µH，34 A 約 2.0 µH | 說明 EVM 用的電感與飽和特性；沒有被公式引用（工具假設 L 固定） |
| TIM_04_simplis_probe_stats.png | image20.png | Timings Picture 1，T23:X27 | SIMPLIS 量測截圖 | 電感電流探針：Max 18.907 A、Mean 13.588 A、RMS 13.698 A、Min −10.45 nA | 用來比對 `iavg`(Y5)、`i_rms`(Y16) 的計算值 |
| TIM_05_simplis_probe_stats_compare.png | image23.png | Timings Picture 4，T62:Y69 | SIMPLIS 量測截圖 | 兩次模擬比較（多一組 simplis_tran4：Mean 11.97 A、RMS 12.04 A） | 同上 |
| TIM_06_simplis_wave_boost.png | image21.png | Timings Picture 2，T29:X37（"Simplis / Boost / Vin<vout 16→18" 表旁） | 波形截圖 | 一個週期的電感電流（平→升→降） | 對應 J31:Q35 的 SIMPLIS 時間戳（K 欄）與計算值（M 欄） |
| TIM_07_simplis_wave_buck.png | image26.png | Timings Picture 7，T39:W47（"Buck / 20→18" 表旁） | 波形截圖 | 同上，橫軸 1.083~1.084 ms | 對應 J39:Q43 |
| TIM_08_simplis_wave_case3.png | image22.png | Timings Picture 3，T53:X60（"Buck / 16→14" 表旁） | 波形截圖 | 兩條電感電流曲線疊在一起比較 | 對應 J55:Q59 |
| TIM_09_rms_formula.png | image27.png | Timings Picture 8，F89:K91 | 公式截圖 | `rms = √((t1+t2)/Tswitch · (min·max + (max−min)²/3))` | 梯形 RMS 公式，即 Y16 `i_rms` 與 EFF 表 W/AG/AL/AP/AU 欄用的形式 |
| TIM_10_rms_waveform.png | image28.png | Timings Picture 9，L89:N97 | 示意圖 | 梯形電流波形，標註 min、max、t1、t3 | 同上 |

### 18.2 圖片內容和公式的對照

**VM_06（Minimum and maximum L）與 VM 的 L 限制公式**

| 圖上的條件 | 圖上的式子 | VM 對應格 | 一致嗎 |
|---|---|---|---|
| Max (max RSL) | L < RSNS / 100 Hz | O64 `Labsmax = Rs/100` | 一致 |
| Min (min RSL) | L > RSNS / 8000 Hz | L64 `Labsmin = Rs/8000` | 一致 |
| Headroom limitation | L > RSNS·10·Vout / (fs·1V) | O65 `=MAX(10*Vout*Rs/Fsw, Labsmin)` | 一致 |
| Ripple current limitation | L ≤ RSNS·**25**·Vout / (fs·1V) | O66 `=MIN(Rs*`**35**`*Vout/Fsw, Labsmax)` | **不一致**：公式把 25 改成 35，O66 的註解寫 "updated factor of 25 to 35 to get EVM into allowed range"。這會讓建議電感上限 `Lopt_Limit`（DC E15）放寬 40 % |

**VM_02~05（四條 VIN 轉換點公式）**

* 這 4 張圖剛好放在 VM **第 40–47 列**，也就是 §3.1 提到「空白但仍有具名範圍指向」的區域（`D_min_boost`、`D_max_buck_BB`、`D_max_exitBB_to_buck`…），旁邊就是 B44:B47 的 `VIN_BB_to_boost`、`VIN_boost_to_BB`、`VIN_buck_to_BB`、`VIN_BB_to_buck`。
* 圖中的值 9.5 / 10.56 / 13.043 / 14.842 V 是**舊範例的靜態截圖**，不會隨輸入更新。和目前 Timings 算出的 11.13 / 14.20 / 18.02 / 23.00 V 不同。
* 推論：早期版本在第 31–43 列用 TON,min / TOFF,min 直接算這四個轉換點，後來改由 Timings 分頁的時序模型計算（B44:B47 → `Timings!F2:F5`）。公式刪掉了，但名稱和說明圖片沒有清掉。這也解釋了為什麼 `Don_buckboost` 指向空的 B37。

**TIM_01 / TIM_02 與 Timings 的相位定義**

| 相位 | 導通 MOSFET | 電感電壓 | Timings 名稱 | dI |
|---|---|---|---|---|
| II | Q1 + Q4 | VIN − VOUT | `t1_Ph_II` = W15 | `dI2_` = (Vin−Vout)/L·t1 |
| I | Q1 + Q3 | VIN | `t2_Ph_I` = W16 | `dI1_` = Vin/L·t2 |
| III | Q2 + Q4 | −VOUT | `t3_Ph_III` = W17 | `dI3_` = Vout/L·t3 |

伏秒平衡（Timings F74:F80 的推導文字）：(Vi−Vo)·t1 + Vi·t2 − Vo·t3 = 0 → t2 = (Vo·T − Vi·t1)/(Vi+Vo)，即 W16 的公式。

### 18.3 四張電路圖的差異（Schematic Management）

四張圖的主體相同：IC、Q1~Q4、RS（電感電流感測，接 CSA/CSB，經 RDIFF1/2 + CDIFF 濾波）、LF、CIN/COUT、RUV1/RUV2、RT、CSS、RCOMP/CCOMP/CHF、RFBT/RFBB、RCFG、RSLOPE、CBOOT1/2、CVCC、CBIAS、RnFLT。差異只在電流監測 / 限制相關的部分：

| 圖 | RCS(OUT) | IMONOUT 腳 | CIMON | RSYNC | ISNSP / ISNSN | 對應 CFG F27 選項 |
|---|---|---|---|---|---|---|
| SCH_NO_RCS（B5） | 無 | 接 VCC | 無 | VCC 上拉 | 一起接到 AGND | 1：IMON_LIMIT off |
| SCH_RCS_MON（D5） | 有（輸出端） | RIMON 接地 | 無 | VCC 上拉 | 接 RCS(OUT) 兩端 | 2：IMON on |
| SCH_RCS_LIM（F5） | 有 | RIMON + CIMON 接地 | 有 | **VCC 上拉**（SYNC = High） | 接 RCS(OUT) 兩端 | 3：ILIMIT positive |
| SCH_RCS_LIM_NEG（H5） | 有 | RIMON + CIMON 接地 | 有 | **接地**（SYNC = Low） | 接 RCS(OUT) 兩端 | 4：ILIMIT negative |

這和 DC E73 的公式 `CHOOSE(IMONLIM_select+1,"",RIMON,"High","Low")`（ILIM 模式時顯示「SYNC pin during startup = High / Low」）一致。

電路圖元件標註和公式變數的對應：

| 電路圖標註 | DC 格 | VM 名稱 |
|---|---|---|
| RUV1 / RUV2 | E67 / E68 | `Ruvlo1` / `Ruvlo2` |
| RIMON / CIMON | M92 / M91（或 E73） | `RcIL` / `Cc1IL`（ILIM）、B274（IMON） |
| RDIFF1 / RDIFF2 / CDIFF | E24 / E25 | `R_CS` / `C_CS_ideal` |
| RS | E30 | `Rs` |
| LF | E16 | `Lf` |
| CIN / COUT | E52 / E43 | `Cin` / `Cout` |
| RCS(OUT) | E72 | `Rshunt` |
| RT | E12 | `Rfadj` → SVC O4 |
| CSS | E61 | `Css_u` |
| RCOMP / CCOMP / CHF | E91 / E92 / E93 | `Rcea1_u` / `Cc1ea_u` / `Cc2ea_u` |
| RFBT / RFBB | E85 / E86 | `Rfb1_` / `Rfb2_u` |
| RCFG | E64 | `CFG` → CFG!B4:B19 |
| RSLOPE | E35 | `Rslope` |
| Q1~Q4 | E98:F122 | `Rdson1`~`4`、`Qg_Q1`~`Q4`… |

### 18.4 DC 上三張連結圖片的快取畫面

這三張是 Excel「照相機」連結圖片在存檔時的畫面，打開檔案後會依目前資料重新繪製：

* **DC_01 電路圖**：等同 SCH_RCS_LIM（`With_RCS` = 3）。
* **DC_02 Bode 圖**："Bode Plot, VIN = VIN(nom)"。紅線是迴路增益（W 欄），藍虛線是相位裕度曲線（Y 欄）。標記點有 fPLOAD（×，1224 Hz）、fZCOMP（○，1895 Hz）、交越點（+，約 4 kHz，PM 標註 **71.69°**）、fPCOMP（×，39 kHz）、fZRHP（粉紅 ○，98.8 kHz）、FZESR（○，612 kHz）。和 STAB!Q4 = 3.98 kHz、Q5 = 71.69° 一致。接近 1 MHz 時相位曲線從 −180° 跳到 +180°，是 Y 欄 `IF(X>0,X-180,X+180)` 繞回造成的繪圖假象，不是真的相位變化。
* **DC_03 效率圖**："η, VIN = VIN(nom)"（對應 PLOT_TYPE = 1）。滿載效率約 95.2 %，左軸效率、右軸損耗（W）。8 A 時：Boost "Control" MOSFET 損耗最大（約 1.3 W 以上，超出圖框），電感損耗約 **1.04 W**，Boost "SYNC" 約 1.0 W，Buck Control 約 0.74 W，Buck SYNC 約 0.62 W，Shunt 約 0.17 W。

  → 這也印證了 §17 第 1 項：效率圖上的文字方塊 "1%" 其實是這裡的**電感損耗 1.0 W**（`EFF_LOAD!BP109`），被誤當成效率顯示。

### 18.5 從圖檔發現的其他問題

| # | 發現 | 說明 |
|---|---|---|
| 1 | 電路圖上的 IC 標示為 **LM5177**，不是 LM51770 | 4 張 Visio 圖與 DC 快取都是 LM5177。腳位名稱（BIAS、CSA/CSB、nFLT、DTRK、SS/ATRK、ISNSP/ISNSN、IMONOUT）是 LM5177 系列的命名 |
| 2 | VM A541:B558 的腳位表和電路圖不一致 | 表中的 FADJ、OTP、CS+/CS−、RS、SS/TRACK、PGOOD 等腳名不在電路圖上，看起來是更早的 IC（例如 LM5175）沿用下來的，而且沒有被任何公式使用 |
| 3 | VM_06 的 25 和公式的 35 不一致 | 見 18.2 |
| 4 | VM_02~05 是過時的靜態截圖 | 數值和目前設計不同，容易誤導 |
| 5 | **TIM_01、TIM_02 有 "NDA Restricted" 浮水印** | 截自尚未公開的 datasheet 草稿。這個 xlsm 和本專案的 `xlsm_images/` 對外分享前要注意保密 |
| 6 | 品牌遺留 | DC 左上的 WIDEVIN 標誌和列印頁首的 National Semiconductor 標誌，是 TI 收購 National 前的範本留下來的 |
| 7 | Bode 圖在 1 MHz 附近的相位跳變 | 只是繪圖繞回，見 18.4 |

---

## 附錄 A：具名範圍總表（326 個）

「被哪些分頁的公式使用」是掃描所有公式中出現的名稱 token 得出的。縮寫：DC = Design Converter、VM = Variable Management、DP = DeviceParameter、CFG = CFG Setting、STAB / ILSTAB = (IL) Stability Calculations、SCH = Schematic Management、EPM = Efficiency Plot Management、TIM = Timings、EFF_VIN / EFF_LOAD = 兩張效率表、SVC = Standard Value Calculator、POC = Parallel Output Caps。標示「未被引用」的名稱，可能只被圖表、控制項或 INDIRECT 間接使用（例如 `EFF_IOUT`、`SCH_*`、`BodePlot`），也可能是殘留。

| 名稱 | 位置 | 目前值 | 被哪些分頁的公式使用 |
|---|---|---|---|
| `_Don1` | `'Variable Management'!$B$52` | 1.33333 | VM |
| `_Don2` | `'Variable Management'!$B$78` | 1.33518 | VM |
| `A_OL_IL` | `'Variable Management'!$B$284` | 4000 | —（未被引用） |
| `Acs` | `'Variable Management'!$B$354` | 76.9231 | ILSTAB, STAB |
| `Adc` | `'Stability Calculations'!$Q$3` | 21563.2 | ILSTAB, STAB |
| `Adc_IL` | `'IL Stability Calculations'!$U$3` | 21563.2 | —（未被引用） |
| `AEA` | `'Variable Management'!$B$352` | 6000 | ILSTAB, STAB |
| `Afb` | `'Variable Management'!$B$353` | 0.0625 | ILSTAB, STAB, VM |
| `AIL` | `'Stability Calculations'!$V$6` | 4000 | ILSTAB |
| `Am` | `'Variable Management'!$B$357` | 57.5019 | ILSTAB, STAB |
| `Am_boost` | `'Variable Management'!$B$356` | 57.5019 | VM |
| `Am_buck` | `'Variable Management'!$B$355` | 56.4375 | VM |
| `Am_IL` | `'IL Stability Calculations'!$H$4` | 57.5019 | —（未被引用） |
| `BodePlot` | `'Efficiency Plot Management'!$B$10` | (空白) | —（未被引用） |
| `BodePlot_ILIM` | `'Efficiency Plot Management'!$B$12` | (空白) | —（未被引用） |
| `BODEPLOT_TYPE` | `'Variable Management'!$L$11` | False | EPM, VM |
| `Boost_CS_Threshold` | `'Variable Management'!$B$123` | 0.05 | VM |
| `Boost_CS_Threshold_max` | `'Variable Management'!$B$124` | 0.0575 | VM |
| `Boost_CS_Threshold_min` | `'Variable Management'!$B$122` | 0.0425 | VM |
| `Buck_CS_threshold` | `'Variable Management'!$B$119` | 0.05 | VM |
| `Buck_CS_threshold_max` | `'Variable Management'!$B$120` | 0.0575 | VM |
| `Buck_CS_threshold_min` | `'Variable Management'!$B$118` | 0.0425 | —（未被引用） |
| `C_CS` | `DeviceParameter!$N$9` | 1.66486e-10 | SVC |
| `C_CS_ideal` | `DeviceParameter!$N$10` | 1.8e-10 | DC |
| `C_IL_recommend` | `'Variable Management'!$B$292` | 1.396e-09 | DC, SVC, VM |
| `Calc_mode` | `'Variable Management'!$L$14` | 1 | VM |
| `Cb` | `'Variable Management'!$B$262` | (空白) | —（未被引用） |
| `Cbw` | `'Variable Management'!$B$348` | 7.34561e-12 | VM |
| `Cbw_IL` | `'Variable Management'!$B$282` | 1.59155e-11 | ILSTAB, STAB |
| `Cby` | `'Variable Management'!$B$261` | (空白) | —（未被引用） |
| `Cc1ea` | `'Variable Management'!$B$382` | 5.7245e-08 | DC, SVC |
| `Cc1ea_u` | `'Variable Management'!$B$387` | 5.6e-08 | VM |
| `Cc1IL` | `'Variable Management'!$B$289` | 2.2e-09 | ILSTAB, STAB |
| `Cc2ea` | `'Variable Management'!$B$383` | 2.62078e-09 | DC, SVC |
| `Cc2ea_u` | `'Variable Management'!$B$388` | 2.7e-09 | VM |
| `Cc2IL` | `'Variable Management'!$B$290` | 1e-15 | ILSTAB, STAB |
| `Cdither` | `'Variable Management'!$B$261` | (空白) | VM |
| `CFG` | `'Variable Management'!$O$10` | 7 | DC |
| `Cin` | `'Variable Management'!$B$218` | 0.0001 | EFF_LOAD, EFF_VIN, SVC, VM |
| `CinEsrMax` | `'Variable Management'!$B$221` | 0.00912694 | VM |
| `Cinmin` | `'Variable Management'!$B$216` | 2.79174e-05 | VM |
| `Co2_` | `'Variable Management'!$B$287` | 6.5e-05 | VM |
| `Co2_u` | `'Variable Management'!$F$287` | 6.5e-05 | ILSTAB |
| `Constant_Current_Loop_Setpoint` | `'Variable Management'!$B$268` | 5 | ILSTAB, VM |
| `Cout` | `'Variable Management'!$B$198` | 0.00013 | EFF_LOAD, EFF_VIN, ILSTAB, SVC, VM |
| `CoutEsr` | `'Variable Management'!$B$202` | 0.002 | EFF_LOAD, EFF_VIN, ILSTAB, VM |
| `CoutEsrMax` | `'Variable Management'!$B$199` | 0.00656133 | VM |
| `Coutmin` | `'Variable Management'!$B$196` | 8.45081e-05 | VM |
| `CSAmp_gain` | `'Variable Management'!$B$297` | 10 | VM |
| `Csb` | `'Variable Management'!$B$259` | (空白) | —（未被引用） |
| `Csby` | `'Variable Management'!$B$260` | (空白) | SVC |
| `Cslope` | `'Variable Management'!$B$300` | 2.5e-12 | VM |
| `Css` | `'Variable Management'!$B$234` | 1.8e-08 | SVC |
| `Css_u` | `'Variable Management'!$B$235` | 1.8e-08 | DC |
| `D_max_buck_BB` | `'Variable Management'!$B$38` | (空白) | —（未被引用） |
| `D_max_buck_enterBB` | `'Variable Management'!$B$42` | (空白) | —（未被引用） |
| `D_max_exitBB_to_buck` | `'Variable Management'!$B$41` | (空白) | —（未被引用） |
| `D_min_boost` | `'Variable Management'!$B$36` | (空白) | —（未被引用） |
| `D_min_boost_exitBB_to_buck` | `'Variable Management'!$B$40` | (空白) | —（未被引用） |
| `DC_RC` | `'Design Converter'!$E$91` | 1.5 | —（未被引用） |
| `Deadtime` | `'Variable Management'!$B$425` | 5e-08 | EFF_LOAD, EFF_VIN, VM |
| `Device` | `'Variable Management'!$B$29` | LM51770 | DC, DP |
| `dI1_` | `Timings!$Y$6` | 5.26189 | TIM |
| `dI2_` | `Timings!$Y$7` | -3.35205 | TIM |
| `dI3_` | `Timings!$Y$8` | 1.90983 | —（未被引用） |
| `Doff1` | `'Variable Management'!$B$53` | -0.333333 | —（未被引用） |
| `Doff2` | `'Variable Management'!$B$79` | -0.335176 | VM |
| `Don_buckboost` | `'Variable Management'!$B$58` | 0 | VM |
| `Don_Vinmax` | `'Variable Management'!$B$57` | 0.444444 | VM |
| `Don_Vinmax_2` | `'Variable Management'!$B$83` | 0.445316 | VM |
| `Don_Vinmin` | `'Variable Management'!$B$55` | 0.625 | VM |
| `Don_Vinmin_2` | `'Variable Management'!$B$81` | 0.630642 | VM |
| `Don_Vinnom` | `'Variable Management'!$B$56` | 0.25 | VM |
| `Don_Vinnom_2` | `'Variable Management'!$B$82` | 0.252476 | VM |
| `EA_Pole1` | `'Stability Calculations'!$C$5` | 0.284163 | ILSTAB, STAB |
| `EA_Pole2` | `'Stability Calculations'!$C$7` | 39190.9 | ILSTAB, STAB |
| `EA_Zero1` | `'Stability Calculations'!$C$4` | 1894.7 | ILSTAB, STAB |
| `eff` | `Timings!$B$9` | 0.95 | EFF_LOAD, EFF_VIN |
| `EFF_IOUT` | `'Efficiency Plot Management'!$B$5` | (空白) | —（未被引用） |
| `EFF_VIN` | `'Efficiency Plot Management'!$B$7` | (空白) | —（未被引用） |
| `Efficiency` | `'Variable Management'!$B$16` | 0.95 | EFF_LOAD, EFF_VIN, TIM, VM |
| `f_cross_Bode` | `'Stability Calculations'!$Q$4` | 3.9811 | VM |
| `Fcross` | `'Variable Management'!$B$380` | 4000 | VM |
| `Fdither` | `'Variable Management'!$B$259` | (空白) | —（未被引用） |
| `Fesr` | `'Variable Management'!$B$372` | 612134 | DC, ILSTAB, STAB |
| `Fload_boost` | `'Variable Management'!$B$367` | 1224.27 | VM |
| `Fload_buck` | `'Variable Management'!$B$366` | 612.134 | VM |
| `Fload_u` | `'Variable Management'!$B$366` | 612.134 | —（未被引用） |
| `Fp1_u` | `'Variable Management'!$B$370` | 0.284163 | DC, STAB |
| `Fp2_u` | `'Variable Management'!$B$371` | 39190.9 | DC, STAB |
| `Fp_dominant` | `'Variable Management'!$B$368` | 1224.27 | DC, ILSTAB, STAB, VM |
| `Fp_IL` | `'Variable Management'!$B$285` | 7.23381 | —（未被引用） |
| `Fp_ILHF` | `'Variable Management'!$B$286` | 103003 | —（未被引用） |
| `Frhp` | `'Variable Management'!$B$373` | 98816.2 | DC, ILSTAB, STAB |
| `Frhp_Vinmin` | `'Variable Management'!$B$374` | 24125.3 | DC |
| `Fsw` | `'Variable Management'!$B$22` | 398000 | DC, DP, EFF_LOAD, EFF_VIN, ILSTAB, POC, STAB, TIM, VM |
| `Fz1_u` | `'Variable Management'!$B$393` | 1894.7 | DC, STAB |
| `GBW` | `'Variable Management'!$B$347` | 13 | VM |
| `GBW_IL` | `'Variable Management'!$B$281` | 2 | VM |
| `Gfs_Q1` | `'Variable Management'!$B$436` | 117 | EFF_LOAD, EFF_VIN, VM |
| `Gfs_Q2` | `'Variable Management'!$D$436` | 117 | VM |
| `Gfs_Q3` | `'Variable Management'!$B$453` | 117 | VM |
| `Gfs_Q4` | `'Variable Management'!$D$453` | 117 | VM |
| `gm` | `'Variable Management'!$B$299` | 0.0006 | ILSTAB, STAB, VM |
| `gm_CDC` | `DeviceParameter!$B$13` | 0.0006 | ILSTAB |
| `gm_IL` | `'Variable Management'!$B$280` | 0.0002 | STAB, VM |
| `gm_slopecomp` | `'Variable Management'!$B$301` | 0.001875 | VM |
| `I_0` | `Timings!$Y$11` | 13.4557 | EFF_LOAD, EFF_VIN, TIM |
| `I_1` | `Timings!$Y$12` | 8.19383 | EFF_LOAD, EFF_VIN, TIM |
| `I_2` | `Timings!$Y$13` | 15.3656 | EFF_LOAD, EFF_VIN, TIM |
| `I_PP` | `Timings!$Y$14` | 7.17172 | —（未被引用） |
| `i_rms` | `Timings!$Y$16` | 11.5954 | —（未被引用） |
| `iavg` | `Timings!$Y$5` | 10.8248 | —（未被引用） |
| `Icoutrms` | `'Variable Management'!$B$189` | 10.328 | —（未被引用） |
| `Ifb` | `'Variable Management'!$B$27` | 6e-08 | VM |
| `Iin` | `Timings!$B$8` | 11.2281 | EFF_LOAD, EFF_VIN, TIM |
| `Iin_Vinmax` | `'Variable Management'!$B$20` | 3.74269 | VM |
| `Iin_Vinmin` | `'Variable Management'!$B$18` | 22.4561 | VM |
| `Iin_Vinnom` | `'Variable Management'!$B$19` | 11.2281 | TIM, VM |
| `IL_peak_OCP` | `'Variable Management'!$B$161` | 44.2308 | —（未被引用） |
| `Ilimit_Margin` | `'Variable Management'!$B$109` | 20 | VM |
| `ILpeak_OCP_Vinmax` | `'Variable Management'!$B$158` | 44.2308 | —（未被引用） |
| `ILpeak_OCP_Vinmin` | `'Variable Management'!$B$156` | 44.2308 | —（未被引用） |
| `ILpeak_OCP_Vinnom` | `'Variable Management'!$B$157` | 44.2308 | —（未被引用） |
| `IMONLIM_select` | `'Variable Management'!$B$266` | 2 | DC, EPM |
| `Iout` | `'Variable Management'!$B$13` | 8 | EFF_LOAD, EFF_VIN, TIM, VM |
| `Iout_OCP_Vinmax` | `'Variable Management'!$B$153` | 32.2455 | VM |
| `Iout_OCP_Vinmin` | `'Variable Management'!$B$151` | 13.2307 | VM |
| `Iout_OCP_Vinnom` | `'Variable Management'!$B$152` | 27.1703 | VM |
| `Iout_threshold` | `'Variable Management'!$B$269` | 50 | VM |
| `Ipeak_inductor_buckboost` | `'Variable Management'!$B$114` | 8.42105 | VM |
| `Ipeak_inductor_max` | `'Variable Management'!$B$115` | 25.097 | DC, VM |
| `Ipeak_inductor_Vinmax` | `'Variable Management'!$B$113` | 14.6371 | VM |
| `Ipeak_inductor_Vinmin` | `'Variable Management'!$B$111` | 25.097 | VM |
| `Ipeak_inductor_Vinnom` | `'Variable Management'!$B$112` | 13.3426 | VM |
| `Iq` | `'Variable Management'!$B$427` | 0.003 | EFF_LOAD, EFF_VIN, VM |
| `Ireverse` | `'Variable Management'!$B$465` | 0 | EFF_LOAD, EFF_VIN |
| `Iripple_edgebuckboost` | `'Variable Management'!$B$90` | 0 | VM |
| `Iripple_Vinmax` | `'Variable Management'!$B$89` | 12.432 | VM |
| `Iripple_Vinmin` | `'Variable Management'!$B$87` | 5.28176 | VM |
| `Iripple_Vinnom` | `'Variable Management'!$B$88` | 4.22908 | EFF_LOAD, EFF_VIN, VM |
| `Isc_const` | `'Variable Management'!$B$302` | 0 | VM |
| `Islope_fixed` | `'Variable Management'!$B$172` | 0 | VM |
| `Iss` | `'Variable Management'!$B$233` | 1e-05 | VM |
| `Iuvlo1` | `'Variable Management'!$B$250` | 0.006 | VM |
| `Iuvlo2` | `'Variable Management'!$B$251` | (空白) | —（未被引用） |
| `Iuvlo_hys` | `'Variable Management'!$B$252` | (空白) | —（未被引用） |
| `k` | `'Variable Management'!$B$339` | 1.11824 | CFG, ILSTAB, STAB, VM |
| `K_1` | `'Variable Management'!$B$343` | 120.123 | VM |
| `k_boost` | `'Variable Management'!$B$336` | 1.11824 | VM |
| `k_buck` | `'Variable Management'!$B$334` | 0.628101 | VM |
| `K_SL` | `'Variable Management'!$B$311` | 1884.42 | VM |
| `KD` | `'Variable Management'!$B$344` | 54718 | VM |
| `KM` | `'Variable Management'!$B$342` | 0.00210183 | VM |
| `L` | `Timings!$B$10` | 1.8e-06 | EFF_LOAD, EFF_VIN, TIM |
| `Labsmax` | `'Variable Management'!$O$64` | 1.3e-05 | VM |
| `Labsmin` | `'Variable Management'!$L$64` | 1.625e-07 | VM |
| `Lf` | `'Variable Management'!$B$75` | 1.8e-06 | EFF_LOAD, EFF_VIN, ILSTAB, TIM, VM |
| `Lopt` | `'Variable Management'!$B$71` | 2.20831e-06 | VM |
| `Lopt_Limit` | `'Variable Management'!$B$74` | 1.82915e-06 | DC |
| `Ltc` | `'Variable Management'!$B$413` | 0.004 | EFF_LOAD, EFF_VIN |
| `mc` | `'Variable Management'!$B$338` | 1.49592 | ILSTAB, STAB |
| `mc_boost` | `'Variable Management'!$B$335` | 1.49592 | VM |
| `mc_buck` | `'Variable Management'!$B$333` | 2.48777 | VM |
| `Mode` | `'Variable Management'!$B$30` | 1 | EFF_LOAD, EFF_VIN |
| `Pi` | `'Variable Management'!$B$296` | 3.14159 | DP, ILSTAB, STAB, VM |
| `PICTURE` | `INDIRECT('Efficiency Plot Management'!$A$2)` |  | —（未被引用） |
| `PICTURE2` | `INDIRECT('Schematic Management'!$A$2)` |  | —（未被引用） |
| `Picture_Bode` | `INDIRECT('Efficiency Plot Management'!$A$9)` |  | —（未被引用） |
| `Pin` | `'Variable Management'!$B$17` | 134.737 | VM |
| `PLOT_TYPE` | `'Variable Management'!$L$10` | 1 | EPM, VM |
| `Pout` | `'Variable Management'!$B$15` | 128 | EFF_LOAD, EFF_VIN, VM |
| `PS_IL_Pole1` | `'IL Stability Calculations'!$O$4` | 1224.27 | —（未被引用） |
| `PS_IL_Pole2` | `'IL Stability Calculations'!$R$3` | 199000 | —（未被引用） |
| `PS_IL_RHPZ` | `'IL Stability Calculations'!$O$5` | 98816.2 | —（未被引用） |
| `PS_IL_Zero1` | `'IL Stability Calculations'!$O$3` | 612134 | —（未被引用） |
| `PS_Pole1` | `'Stability Calculations'!$K$4` | 1224.27 | ILSTAB, STAB |
| `PS_Pole2` | `'Stability Calculations'!$N$3` | 199000 | ILSTAB, STAB |
| `PS_RHPZ` | `'Stability Calculations'!$K$5` | 98816.2 | ILSTAB, STAB, VM |
| `PS_Zero1` | `'Stability Calculations'!$K$3` | 612134 | ILSTAB, STAB |
| `pterm1` | `'Stability Calculations'!$N$4` | 102458 | STAB |
| `pterm1_IL` | `'IL Stability Calculations'!$R$4` | 102458 | —（未被引用） |
| `pterm2` | `'Stability Calculations'!$N$5` | 3.9601e+10 | STAB |
| `pterm2_IL` | `'IL Stability Calculations'!$R$5` | 3.9601e+10 | —（未被引用） |
| `Q` | `'Variable Management'!$B$340` | 0.514866 | ILSTAB, STAB |
| `Qg_Q1` | `'Variable Management'!$B$432` | 6.7e-08 | EFF_LOAD, EFF_VIN, VM |
| `Qg_Q2` | `'Variable Management'!$D$432` | 6.7e-08 | EFF_LOAD, EFF_VIN, VM |
| `Qg_Q3` | `'Variable Management'!$B$449` | 6.7e-08 | EFF_LOAD, EFF_VIN |
| `Qg_Q4` | `'Variable Management'!$D$449` | 6.7e-08 | EFF_LOAD, EFF_VIN |
| `Qgd_Q1` | `'Variable Management'!$B$433` | 1.031e-08 | EFF_LOAD, EFF_VIN, VM |
| `Qgd_Q2` | `'Variable Management'!$D$433` | 1.031e-08 | —（未被引用） |
| `Qgd_Q3` | `'Variable Management'!$B$450` | 1.031e-08 | —（未被引用） |
| `Qgd_Q4` | `'Variable Management'!$D$450` | 1.031e-08 | VM |
| `Qgs_Q1` | `'Variable Management'!$B$434` | 1.125e-08 | EFF_LOAD, EFF_VIN, VM |
| `Qgs_Q2` | `'Variable Management'!$D$434` | 1.125e-08 | —（未被引用） |
| `Qgs_Q3` | `'Variable Management'!$B$451` | 1.125e-08 | —（未被引用） |
| `Qgs_Q4` | `'Variable Management'!$D$451` | 1.125e-08 | VM |
| `Qrr_Q1` | `'Variable Management'!$B$441` | 0 | —（未被引用） |
| `Qrr_Q2` | `'Variable Management'!$D$441` | 4.5e-08 | EFF_LOAD, EFF_VIN, VM |
| `Qrr_Q3` | `'Variable Management'!$B$458` | 5.6e-08 | EFF_LOAD, EFF_VIN |
| `Qrr_Q4` | `'Variable Management'!$D$458` | 0 | —（未被引用） |
| `QrrSky_Q2` | `'Variable Management'!$B$464` | 0 | EFF_LOAD, EFF_VIN |
| `QrrSky_Q3` | `'Variable Management'!$D$464` | 0 | EFF_LOAD, EFF_VIN |
| `R_CS` | `DeviceParameter!$N$8` | 10 | DC |
| `R_IL_recommend` | `'Variable Management'!$B$293` | 477.286 | DC, SVC |
| `radconv` | `'Stability Calculations'!$Q$6` | 57.2958 | STAB |
| `radconv_IL` | `'IL Stability Calculations'!$U$6` | 57.2958 | —（未被引用） |
| `Rcea1` | `'Variable Management'!$B$381` | 1513.96 | DC, SVC, VM |
| `Rcea1_u` | `'Variable Management'!$B$386` | 1500 | VM |
| `RcIL` | `'Variable Management'!$B$288` | 700 | ILSTAB, STAB |
| `RCinEsr` | `'Variable Management'!$B$223` | 0.002 | EFF_LOAD, EFF_VIN, VM |
| `RCoutEsr` | `'Variable Management'!$B$202` | 0.002 | EFF_LOAD, EFF_VIN |
| `Rdcr` | `'Variable Management'!$B$76` | 0.0032 | EFF_LOAD, EFF_VIN, ILSTAB, VM |
| `Rdson1` | `'Variable Management'!$B$430` | 0.0035 | EFF_LOAD, EFF_VIN |
| `Rdson2` | `'Variable Management'!$D$430` | 0.0035 | EFF_LOAD, EFF_VIN |
| `Rdson3` | `'Variable Management'!$B$447` | 0.0035 | EFF_LOAD, EFF_VIN |
| `Rdson4` | `'Variable Management'!$D$447` | 0.0035 | EFF_LOAD, EFF_VIN |
| `RdsonBotTc` | `'Variable Management'!$B$448` | 0.004 | EFF_LOAD, EFF_VIN |
| `RdsonDRB` | `'Variable Management'!$B$422` | 1.5 | EFF_LOAD, EFF_VIN |
| `RdsonDRB2` | `'Variable Management'!$B$423` | 1 | EFF_LOAD, EFF_VIN |
| `RdsonDRT` | `'Variable Management'!$B$419` | 1.5 | EFF_LOAD, EFF_VIN, VM |
| `RdsonDRT2` | `'Variable Management'!$B$420` | 1 | EFF_LOAD, EFF_VIN, VM |
| `RdsonTopTc` | `'Variable Management'!$B$431` | 0.004 | EFF_LOAD, EFF_VIN |
| `Rfadj` | `'Variable Management'!$B$26` | 75524.7 | SVC |
| `Rfb1_` | `'Variable Management'!$B$396` | 71500 | SVC |
| `Rfb2_` | `'Variable Management'!$B$397` | 4766.67 | SVC |
| `Rfb2_u` | `'Variable Management'!$B$398` | 4750 | DC, VM |
| `Rfets` | `'Variable Management'!$B$402` | 0 | —（未被引用） |
| `Rgate1` | `'Variable Management'!$B$435` | 0.63 | EFF_LOAD, EFF_VIN, VM |
| `Rgate2` | `'Variable Management'!$D$435` | 0.63 | EFF_LOAD, EFF_VIN |
| `Rgate3` | `'Variable Management'!$B$452` | 0.63 | —（未被引用） |
| `Rgate4` | `'Variable Management'!$D$452` | 0.63 | VM |
| `Ri` | `'Variable Management'!$B$298` | 0.013 | VM |
| `Ro` | `'Variable Management'!$B$363` | 2 | VM |
| `RoEA` | `'Variable Management'!$B$349` | 10000000 | VM |
| `RoEA_IL` | `'Variable Management'!$B$283` | 20000000 | ILSTAB, STAB, VM |
| `Rout` | `'Variable Management'!$B$14` | 2 | ILSTAB, VM |
| `Rpg` | `'Variable Management'!$B$254` | 20 | —（未被引用） |
| `Rs` | `'Variable Management'!$B$146` | 0.0013 | EFF_LOAD, EFF_VIN, ILSTAB, VM |
| `Rs_ideal` | `'Variable Management'!$B$145` | 0.00163899 | DC, SVC |
| `Rs_in_range` | `'Variable Management'!$B$147` | 1 | —（未被引用） |
| `Rsb` | `'Variable Management'!$B$255` | (空白) | —（未被引用） |
| `Rshunt` | `'Variable Management'!$B$270` | 10 | DC, EFF_LOAD, EFF_VIN, ILSTAB, SVC, VM |
| `Rshunt_std` | `'Standard Value Calculator'!$O$21` | 10 | —（未被引用） |
| `Rslope` | `'Variable Management'!$B$173` | 69800 | VM |
| `Rslope_ideal` | `'Variable Management'!$B$164` | 69230.8 | SVC |
| `Ruvlo1` | `'Variable Management'!$B$241` | 32.7635 | SVC, VM |
| `Ruvlo2` | `'Variable Management'!$B$242` | 13.4267 | SVC |
| `SCH772_NO_RCS` | `'Schematic Management'!$B$11` | (空白) | —（未被引用） |
| `SCH772_RCS_LIM` | `'Schematic Management'!$F$11` | (空白) | —（未被引用） |
| `SCH772_RCS_LIM_NEG` | `'Schematic Management'!$H$11` | (空白) | —（未被引用） |
| `SCH772_RCS_MON` | `'Schematic Management'!$D$11` | (空白) | —（未被引用） |
| `SCH_NO_RCS` | `'Schematic Management'!$B$5` | (空白) | —（未被引用） |
| `SCH_RCS_LIM` | `'Schematic Management'!$F$5` | (空白) | —（未被引用） |
| `SCH_RCS_LIM_NEG` | `'Schematic Management'!$H$5` | (空白) | —（未被引用） |
| `SCH_RCS_MON` | `'Schematic Management'!$D$5` | (空白) | —（未被引用） |
| `Se` | `'Variable Management'!$B$306` | 42979.9 | VM |
| `Se_ext` | `'Variable Management'!$B$314` | 3.30615e+06 | VM |
| `Se_fixed` | `'Variable Management'!$B$304` | 0 | VM |
| `Se_proportional` | `'Variable Management'!$B$305` | 42979.9 | VM |
| `Sf` | `'Variable Management'!$B$316` | 2.22222e+06 | VM |
| `Sn` | `'Variable Management'!$B$315` | 6.66667e+06 | VM |
| `t1_Ph_II` | `Timings!$W$15` | 1.50842e-06 | EFF_LOAD, EFF_VIN |
| `t2_Ph_I` | `Timings!$W$16` | 7.89283e-07 | EFF_LOAD, EFF_VIN |
| `t3_Ph_III` | `Timings!$W$17` | 2.14856e-07 | EFF_LOAD, EFF_VIN |
| `Ta` | `'Variable Management'!$B$411` | 25 | EFF_LOAD, EFF_VIN |
| `Tf_boost` | `'Variable Management'!$D$472` | 1.23924e-08 | EFF_LOAD, EFF_VIN |
| `Tf_buck` | `'Variable Management'!$B$472` | 1.25577e-08 | EFF_LOAD, EFF_VIN |
| `ThetaCa` | `'Variable Management'!$B$414` | 20 | EFF_LOAD, EFF_VIN |
| `ThetaJA_Q1` | `'Variable Management'!$B$443` | 65 | EFF_LOAD, EFF_VIN |
| `ThetaJA_Q2` | `'Variable Management'!$D$443` | 65 | EFF_LOAD, EFF_VIN |
| `ThetaJA_Q3` | `'Variable Management'!$B$460` | 65 | EFF_LOAD, EFF_VIN |
| `ThetaJA_Q4` | `'Variable Management'!$D$460` | 65 | EFF_LOAD, EFF_VIN |
| `ThetaJaCtrl` | `'Variable Management'!$B$468` | 30.5 | VM |
| `TL` | `'Variable Management'!$B$105` | (空白) | —（未被引用） |
| `Toff_min_buck` | `'Variable Management'!$B$32` | (空白) | —（未被引用） |
| `Toff_Vinmax` | `'Variable Management'!$B$322` | 1.39368e-06 | VM |
| `Ton_min_boost` | `'Variable Management'!$B$33` | (空白) | —（未被引用） |
| `Ton_Vinmin` | `'Variable Management'!$B$324` | 1.58453e-06 | VM |
| `Tr_boost` | `'Variable Management'!$D$471` | 6.28078e-09 | EFF_LOAD, EFF_VIN |
| `Tr_buck` | `'Variable Management'!$B$471` | 6.24888e-09 | EFF_LOAD, EFF_VIN |
| `Trr_Q1` | `'Variable Management'!$B$442` | 4e-08 | —（未被引用） |
| `Trr_Q2` | `'Variable Management'!$D$442` | 4e-08 | EFF_LOAD, EFF_VIN |
| `Trr_Q3` | `'Variable Management'!$B$459` | 4e-08 | —（未被引用） |
| `Trr_Q4` | `'Variable Management'!$D$459` | 4e-08 | —（未被引用） |
| `Ts` | `'Variable Management'!$B$23` | 2.51256e-06 | —（未被引用） |
| `Tss` | `'Variable Management'!$B$232` | 0.0018 | VM |
| `V_sc_fixed` | `'Variable Management'!$B$308` | 0 | VM |
| `V_sc_proportional` | `'Variable Management'!$B$309` | 0.10799 | —（未被引用） |
| `V_sc_total` | `'Variable Management'!$B$310` | 0.10799 | —（未被引用） |
| `Vdd` | `'Variable Management'!$B$417` | 7.5 | EFF_LOAD, EFF_VIN, VM |
| `VfdSky_Q2` | `'Variable Management'!$B$463` | 0 | EFF_LOAD, EFF_VIN |
| `VfdSky_Q3` | `'Variable Management'!$D$463` | 0 | EFF_LOAD, EFF_VIN |
| `Vin` | `'Variable Management'!$B$8` | 12 | EFF_LOAD, EFF_VIN, ILSTAB, TIM, VM |
| `VIN_BB_to_boost` | `'Variable Management'!$B$44` | 11.1293 | —（未被引用） |
| `VIN_BB_to_buck` | `'Variable Management'!$B$47` | 23.0024 | EFF_LOAD, EFF_VIN |
| `VIN_boost_to_BB` | `'Variable Management'!$B$45` | 14.2028 | EFF_LOAD, EFF_VIN |
| `VIN_edgebuckboost` | `'Variable Management'!$B$10` | 11.1293 | VM |
| `VIN_max` | `'Variable Management'!$B$9` | 36 | EFF_VIN, TIM, VM |
| `VIN_min` | `'Variable Management'!$B$7` | 6 | EFF_VIN, TIM, VM |
| `VIN_nom` | `'Variable Management'!$B$8` | 12 | DC, ILSTAB, STAB, TIM, VM |
| `Vinripple1` | `'Variable Management'!$B$212` | 0.18 | VM |
| `Vinripple2` | `'Variable Management'!$B$227` | 0.0786833 | DC |
| `VINuvlo_off` | `'Variable Management'!$B$240` | 4.025 | VM |
| `VINuvlo_on` | `'Variable Management'!$B$239` | 4.6 | VM |
| `Vout` | `'Variable Management'!$B$12` | 16 | DC, EFF_LOAD, EFF_VIN, ILSTAB, STAB, TIM, VM |
| `Vref` | `'Variable Management'!$B$24` | 1 | DC, VM |
| `Vripple1` | `'Variable Management'!$B$85` | 0.15 | VM |
| `Vripple2` | `'Variable Management'!$B$203` | 0.113509 | DC, VM |
| `Vsd_Q1` | `'Variable Management'!$B$440` | 0.75 | EFF_LOAD, EFF_VIN, VM |
| `Vsd_Q2` | `'Variable Management'!$D$440` | 0.75 | EFF_LOAD, EFF_VIN, VM |
| `Vsd_Q3` | `'Variable Management'!$B$457` | 0.75 | EFF_LOAD, EFF_VIN |
| `Vsd_Q4` | `'Variable Management'!$D$457` | 0.75 | EFF_LOAD, EFF_VIN |
| `Vsp_Q1` | `'Variable Management'!$B$438` | 2.06838 | EFF_LOAD, EFF_VIN, VM |
| `Vsp_Q2` | `'Variable Management'!$D$438` | 2.06838 | —（未被引用） |
| `Vsp_Q3` | `'Variable Management'!$B$455` | 2.09597 | —（未被引用） |
| `Vsp_Q4` | `'Variable Management'!$D$455` | 2.09597 | VM |
| `Vth_Q1` | `'Variable Management'!$B$437` | 2 | EFF_LOAD, EFF_VIN, VM |
| `Vth_Q2` | `'Variable Management'!$D$437` | 2 | VM |
| `Vth_Q3` | `'Variable Management'!$B$454` | 2 | VM |
| `Vth_Q4` | `'Variable Management'!$D$454` | 2 | VM |
| `Vuvlo_hys` | `'Variable Management'!$B$249` | 0.11 | —（未被引用） |
| `Vuvlo_off` | `'Variable Management'!$B$248` | 1.17 | VM |
| `Vuvlo_on` | `'Variable Management'!$B$247` | 1.28 | VM |
| `With_RCS` | `'Variable Management'!$B$271` | 3 | SCH |
