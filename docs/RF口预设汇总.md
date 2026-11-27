# RF 口预设汇总（CMW500 LTE 测试工具）

本文汇总工具中六套 RF 口预设的端口映射、三种接收模式的路由命令与 FDC 校准口。
所有内容由源码中的 `RF_PORT_COMBOS` 与实际生成函数导出，与界面悬停提示、运行日志一致。

下拉顺序：**RF1+RF2 / RF1+RF3 / RF1+RF4 / RF3+RF1 / RF3+RF2 / RF3+RF4**，最后一项为「自定义」。

## 1. 六套预设总览

| 预设 | 主集口 | 分集口 | 上行（连接器/转换器） | 主集下行 | 分集下行（双流第二路） | DRX 单流下行 | FDC 校准 TX | FDC 校准 RX |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| **RF1+RF2**（默认） | RF1COM | RF2COM | RF1COM / RFRX1 | RFTX1 | RFTX3 | RFTX1（主集转换器打到分集口） | RF1C / RF2C | RF1C |
| **RF1+RF3** | RF1COM | RF3COM | RF1COM / RFRX1 | RFTX1 | RFTX2 | **RFTX2**（分集口自己的转换器） | RF1C / RF3C | RF1C |
| **RF1+RF4** | RF1COM | RF4COM | RF1COM / RFRX1 | RFTX1 | RFTX2 | **RFTX2**（分集口自己的转换器） | RF1C / RF4C | RF1C |
| **RF3+RF1** | RF3COM | RF1COM | RF3COM / RFRX2 | RFTX2 | RFTX1 | **RFTX1**（分集口自己的转换器） | RF3C / RF1C | RF3C |
| **RF3+RF2** | RF3COM | RF2COM | RF3COM / RFRX2 | RFTX2 | RFTX1 | **RFTX1**（分集口自己的转换器） | RF3C / RF2C | RF3C |
| **RF3+RF4** | RF3COM | RF4COM | RF3COM / RFRX2 | RFTX2 | RFTX4 | RFTX2（主集转换器打到分集口） | RF3C / RF4C | RF3C |

> 关键差异：RF1+RF2 与 RF3+RF4 的 DRX 单流用**主集**下行转换器；RF1+RF3、RF1+RF4 用分集口的
> TX2，RF3+RF1、RF3+RF2 用分集口的 TX1，这四套组合定义里多了 `drx_tx` 字段做覆盖。
> PRX+DRX 双流的第二路始终用分集口转换器（div.tx）。

## 2. 各预设的三种接收模式

三种模式的含义：PRX = 主集单流（SCEL）、DRX = 分集单流（SCEL）、PRX+DRX = 双流 2x2（MIMO22）。

### 2.1 RF1+RF2（默认，RF1COM 主集 + RF2COM 分集）

```
MODE     SCEN    DOWNLINK (RF Out)            UPLINK (RF In)
PRX      SCEL    RF1COM/RFTX1                 RF1COM/RFRX1
DRX      SCEL    RF2COM/RFTX1                 RF1COM/RFRX1
PRX+DRX  MIMO22  RF1COM/RFTX1 + RF2COM/RFTX3  RF1COM/RFRX1
```

| 模式 | SCPI 路由命令 |
| --- | --- |
| PRX | `ROUT:LTE:SIGN:SCEN:SCEL RF1C,RX1,RF1C,TX1` |
| DRX | `ROUT:LTE:SIGN:SCEN:SCEL RF1C,RX1,RF2C,TX1` |
| PRX+DRX | `ROUT:LTE:SIGN:SCEN:MIMO22 RF1C,RX1,RF1C,TX1,RF2C,TX3` |

FDC 校准口：TX = RF1C / RF2C，RX = RF1C。

### 2.2 RF1+RF3（RF1COM 主集 + RF3COM 分集）

```
MODE     SCEN    DOWNLINK (RF Out)            UPLINK (RF In)
PRX      SCEL    RF1COM/RFTX1                 RF1COM/RFRX1
DRX      SCEL    RF3COM/RFTX2                 RF1COM/RFRX1
PRX+DRX  MIMO22  RF1COM/RFTX1 + RF3COM/RFTX2  RF1COM/RFRX1
```

| 模式 | SCPI 路由命令 |
| --- | --- |
| PRX | `ROUT:LTE:SIGN:SCEN:SCEL RF1C,RX1,RF1C,TX1` |
| DRX | `ROUT:LTE:SIGN:SCEN:SCEL RF1C,RX1,RF3C,TX2` |
| PRX+DRX | `ROUT:LTE:SIGN:SCEN:MIMO22 RF1C,RX1,RF1C,TX1,RF3C,TX2` |

FDC 校准口：TX = RF1C / RF3C，RX = RF1C。

### 2.3 RF1+RF4（RF1COM 主集 + RF4COM 分集）

```
MODE     SCEN    DOWNLINK (RF Out)            UPLINK (RF In)
PRX      SCEL    RF1COM/RFTX1                 RF1COM/RFRX1
DRX      SCEL    RF4COM/RFTX2                 RF1COM/RFRX1
PRX+DRX  MIMO22  RF1COM/RFTX1 + RF4COM/RFTX2  RF1COM/RFRX1
```

| 模式 | SCPI 路由命令 |
| --- | --- |
| PRX | `ROUT:LTE:SIGN:SCEN:SCEL RF1C,RX1,RF1C,TX1` |
| DRX | `ROUT:LTE:SIGN:SCEN:SCEL RF1C,RX1,RF4C,TX2` |
| PRX+DRX | `ROUT:LTE:SIGN:SCEN:MIMO22 RF1C,RX1,RF1C,TX1,RF4C,TX2` |

FDC 校准口：TX = RF1C / RF4C，RX = RF1C。

### 2.4 RF3+RF1（RF3COM 主集 + RF1COM 分集）

```
MODE     SCEN    DOWNLINK (RF Out)            UPLINK (RF In)
PRX      SCEL    RF3COM/RFTX2                 RF3COM/RFRX2
DRX      SCEL    RF1COM/RFTX1                 RF3COM/RFRX2
PRX+DRX  MIMO22  RF3COM/RFTX2 + RF1COM/RFTX1  RF3COM/RFRX2
```

| 模式 | SCPI 路由命令 |
| --- | --- |
| PRX | `ROUT:LTE:SIGN:SCEN:SCEL RF3C,RX2,RF3C,TX2` |
| DRX | `ROUT:LTE:SIGN:SCEN:SCEL RF3C,RX2,RF1C,TX1` |
| PRX+DRX | `ROUT:LTE:SIGN:SCEN:MIMO22 RF3C,RX2,RF3C,TX2,RF1C,TX1` |

FDC 校准口：TX = RF3C / RF1C，RX = RF3C。

### 2.5 RF3+RF2（RF3COM 主集 + RF2COM 分集）

```
MODE     SCEN    DOWNLINK (RF Out)            UPLINK (RF In)
PRX      SCEL    RF3COM/RFTX2                 RF3COM/RFRX2
DRX      SCEL    RF2COM/RFTX1                 RF3COM/RFRX2
PRX+DRX  MIMO22  RF3COM/RFTX2 + RF2COM/RFTX1  RF3COM/RFRX2
```

| 模式 | SCPI 路由命令 |
| --- | --- |
| PRX | `ROUT:LTE:SIGN:SCEN:SCEL RF3C,RX2,RF3C,TX2` |
| DRX | `ROUT:LTE:SIGN:SCEN:SCEL RF3C,RX2,RF2C,TX1` |
| PRX+DRX | `ROUT:LTE:SIGN:SCEN:MIMO22 RF3C,RX2,RF3C,TX2,RF2C,TX1` |

FDC 校准口：TX = RF3C / RF2C，RX = RF3C。

### 2.6 RF3+RF4（RF3COM 主集 + RF4COM 分集）

```
MODE     SCEN    DOWNLINK (RF Out)            UPLINK (RF In)
PRX      SCEL    RF3COM/RFTX2                 RF3COM/RFRX2
DRX      SCEL    RF4COM/RFTX2                 RF3COM/RFRX2
PRX+DRX  MIMO22  RF3COM/RFTX2 + RF4COM/RFTX4  RF3COM/RFRX2
```

| 模式 | SCPI 路由命令 |
| --- | --- |
| PRX | `ROUT:LTE:SIGN:SCEN:SCEL RF3C,RX2,RF3C,TX2` |
| DRX | `ROUT:LTE:SIGN:SCEN:SCEL RF3C,RX2,RF4C,TX2` |
| PRX+DRX | `ROUT:LTE:SIGN:SCEN:MIMO22 RF3C,RX2,RF3C,TX2,RF4C,TX4` |

FDC 校准口：TX = RF3C / RF4C，RX = RF3C。

## 3. 实测灵敏度对比（2026-11-27）

数据来自 14:22–16:18 的 18 份测试报告（6 套组合 × PRX/DRX/PRX+DRX 三种模式 × 低/中/高信道），
测试条件：B1 / 10MHz，CH50 / CH300 / CH550，Sens limit = -96.3 dBm。下表为三信道平均值；
「优于限值」= Sens limit − Sens@Pmin，数值越大余量越充足。

| 预设 | PRX 均值 (dBm) | 优于限值 (dB) | DRX 均值 (dBm) | 优于限值 (dB) | PRX+DRX 均值 (dBm) | 优于限值 (dB) |
| --- | --- | --- | --- | --- | --- | --- |
| **RF1+RF2** | -97.67 | 1.37 | -97.93 | 1.63 | -101.03 | 4.73 |
| **RF1+RF3** | -97.67 | 1.37 | -98.37 | 2.07 | -101.00 | 4.70 |
| **RF1+RF4** | -97.70 | 1.40 | -98.07 | 1.77 | -100.90 | 4.60 |
| **RF3+RF1** | -97.80 | 1.50 | -98.23 | 1.93 | -100.97 | 4.67 |
| **RF3+RF2** | -97.73 | 1.43 | -97.93 | 1.63 | -100.87 | 4.57 |
| **RF3+RF4** | -97.77 | 1.47 | -98.03 | 1.73 | -101.03 | 4.73 |

- 六套组合全部达标，灵敏度余量 **1.37~4.73 dB**；双流模式约 -101 dBm，比单流好约 3 dB。
- 单流 PRX 均在 -97.7 dBm 左右；DRX 略差 0.2~0.7 dB。六套组合之间单流最大差 0.7 dB、
  双流最大差 0.2 dB，没有哪套接线口径表现出明显优势。
- 完整逐信道数值、路径、起止时间与来源报告见
  `result/RF六组合灵敏度汇总_20261127.xlsx`（「灵敏度明细」54 行 /「组合对比」18 行 /「端口映射」18 行）。
  该汇总由 18 份 `result/LTE_20261127_*.txt` 报告生成，早上 10 点的旧口径报告未纳入。

## 4. 配置与切换方法

- **界面**：终测仪设置行右侧【RF口】下拉，六套预设 + 「自定义」。选中后 PRX/DRX/PRX+DRX 三种
  模式的路由命令、FDC 校准口自动展开，并写回 `config.json`；鼠标悬停下拉可看完整映射表。
- **配置文件**：`config.json` 只需改一行，例如 `"rf_port_combo": "RF1+RF4"`；工具会按「主集/分集」
  两行自动覆盖 `rx_path_ports`，不必手改 6 处端口。
- **自定义接法**：把 `rf_port_combo` 留空，改用 `rx_path_ports` 手写端口表
  （`rx = [上行连接器, 转换器]`，`tx = [[下行连接器, 转换器], ...]`）。
- **热重载**：参数配置文件保存后界面自动同步；测量进行中不刷新，结束后或下次初始化生效。
- 手册依据：CMW500 手册 23.5.4 Signal Routing（p.1741~1743，SCEL / MIMO22）。

## 5. 注意事项

1. **不要混淆 DRX 转换器**：RF1+RF2 的 DRX 是 RF2COM/RFTX1（主集 TX1 打到分集口）、
   RF3+RF4 的 DRX 是 RF4COM/RFTX2（主集 TX2 打到分集口）；RF1+RF3 / RF1+RF4 用分集口的 TX2，
   RF3+RF1 / RF3+RF2 用分集口的 TX1。接错转换器会导致分集灵敏度数据无效。
2. **上行口分组**：RF1+RF2 / RF1+RF3 / RF1+RF4 上行走 RF1COM/RFRX1；
   RF3+RF1 / RF3+RF2 / RF3+RF4 上行走 RF3COM/RFRX2。上行口不同不可混用数据。
3. **断点续测受组合影响**：断点指纹包含 `rf_port_combo`。切换组合后，之前按另一套口径测出的
   数据不可续测（工具会拒绝自动续测），需要从头测或改回原组合。
4. 六套预设的展开逻辑与自检断言均在 `cmw500_lte_tx_tool.py` 中，可用 `--selftest` 验证路由命令。
