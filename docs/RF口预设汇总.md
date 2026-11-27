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

## 3. 实测灵敏度数据（2026-11-27）

数据来自 14:22–16:18 的 18 份测试报告（6 套组合 × PRX/DRX/PRX+DRX 三种模式 × 低/中/高信道），
测试条件：B1 / 10MHz，CH50 / CH300 / CH550，Sens limit = -96.3 dBm。逐条列出，不做平均。
「优于限值」= Sens limit − Sens@Pmin，正值代表达标余量。

| 预设 | 模式 | 信道 | Sens limit (dBm) | Sens@Pmin (dBm) | 优于限值 (dB) | 来源报告 |
| --- | --- | --- | --- | --- | --- | --- |
| RF1+RF2 | PRX | 低 CH50 | -96.3 | -97.6 | 1.3 | LTE_20261127_142539.txt |
| RF1+RF2 | PRX | 中 CH300 | -96.3 | -97.7 | 1.4 | LTE_20261127_142539.txt |
| RF1+RF2 | PRX | 高 CH550 | -96.3 | -97.7 | 1.4 | LTE_20261127_142539.txt |
| RF1+RF2 | DRX | 低 CH50 | -96.3 | -98.1 | 1.8 | LTE_20261127_142641.txt |
| RF1+RF2 | DRX | 中 CH300 | -96.3 | -97.8 | 1.5 | LTE_20261127_142641.txt |
| RF1+RF2 | DRX | 高 CH550 | -96.3 | -97.9 | 1.6 | LTE_20261127_142641.txt |
| RF1+RF2 | PRX+DRX | 低 CH50 | -96.3 | -101.1 | 4.8 | LTE_20261127_142741.txt |
| RF1+RF2 | PRX+DRX | 中 CH300 | -96.3 | -101.0 | 4.7 | LTE_20261127_142741.txt |
| RF1+RF2 | PRX+DRX | 高 CH550 | -96.3 | -101.0 | 4.7 | LTE_20261127_142741.txt |
| RF1+RF3 | PRX | 低 CH50 | -96.3 | -97.6 | 1.3 | LTE_20261127_142438.txt |
| RF1+RF3 | PRX | 中 CH300 | -96.3 | -97.7 | 1.4 | LTE_20261127_142438.txt |
| RF1+RF3 | PRX | 高 CH550 | -96.3 | -97.7 | 1.4 | LTE_20261127_142438.txt |
| RF1+RF3 | DRX | 低 CH50 | -96.3 | -98.5 | 2.2 | LTE_20261127_142324.txt |
| RF1+RF3 | DRX | 中 CH300 | -96.3 | -98.2 | 1.9 | LTE_20261127_142324.txt |
| RF1+RF3 | DRX | 高 CH550 | -96.3 | -98.4 | 2.1 | LTE_20261127_142324.txt |
| RF1+RF3 | PRX+DRX | 低 CH50 | -96.3 | -101.0 | 4.7 | LTE_20261127_142225.txt |
| RF1+RF3 | PRX+DRX | 中 CH300 | -96.3 | -101.0 | 4.7 | LTE_20261127_142225.txt |
| RF1+RF3 | PRX+DRX | 高 CH550 | -96.3 | -101.0 | 4.7 | LTE_20261127_142225.txt |
| RF1+RF4 | PRX | 低 CH50 | -96.3 | -97.6 | 1.3 | LTE_20261127_161356.txt |
| RF1+RF4 | PRX | 中 CH300 | -96.3 | -97.7 | 1.4 | LTE_20261127_161356.txt |
| RF1+RF4 | PRX | 高 CH550 | -96.3 | -97.8 | 1.5 | LTE_20261127_161356.txt |
| RF1+RF4 | DRX | 低 CH50 | -96.3 | -98.3 | 2.0 | LTE_20261127_161250.txt |
| RF1+RF4 | DRX | 中 CH300 | -96.3 | -97.9 | 1.6 | LTE_20261127_161250.txt |
| RF1+RF4 | DRX | 高 CH550 | -96.3 | -98.0 | 1.7 | LTE_20261127_161250.txt |
| RF1+RF4 | PRX+DRX | 低 CH50 | -96.3 | -101.0 | 4.7 | LTE_20261127_161145.txt |
| RF1+RF4 | PRX+DRX | 中 CH300 | -96.3 | -100.8 | 4.5 | LTE_20261127_161145.txt |
| RF1+RF4 | PRX+DRX | 高 CH550 | -96.3 | -100.9 | 4.6 | LTE_20261127_161145.txt |
| RF3+RF1 | PRX | 低 CH50 | -96.3 | -97.7 | 1.4 | LTE_20261127_161552.txt |
| RF3+RF1 | PRX | 中 CH300 | -96.3 | -97.8 | 1.5 | LTE_20261127_161552.txt |
| RF3+RF1 | PRX | 高 CH550 | -96.3 | -97.9 | 1.6 | LTE_20261127_161552.txt |
| RF3+RF1 | DRX | 低 CH50 | -96.3 | -98.4 | 2.1 | LTE_20261127_161650.txt |
| RF3+RF1 | DRX | 中 CH300 | -96.3 | -98.1 | 1.8 | LTE_20261127_161650.txt |
| RF3+RF1 | DRX | 高 CH550 | -96.3 | -98.2 | 1.9 | LTE_20261127_161650.txt |
| RF3+RF1 | PRX+DRX | 低 CH50 | -96.3 | -101.0 | 4.7 | LTE_20261127_161752.txt |
| RF3+RF1 | PRX+DRX | 中 CH300 | -96.3 | -100.9 | 4.6 | LTE_20261127_161752.txt |
| RF3+RF1 | PRX+DRX | 高 CH550 | -96.3 | -101.0 | 4.7 | LTE_20261127_161752.txt |
| RF3+RF2 | PRX | 低 CH50 | -96.3 | -97.6 | 1.3 | LTE_20261127_144241.txt |
| RF3+RF2 | PRX | 中 CH300 | -96.3 | -97.8 | 1.5 | LTE_20261127_144241.txt |
| RF3+RF2 | PRX | 高 CH550 | -96.3 | -97.8 | 1.5 | LTE_20261127_144241.txt |
| RF3+RF2 | DRX | 低 CH50 | -96.3 | -98.1 | 1.8 | LTE_20261127_144139.txt |
| RF3+RF2 | DRX | 中 CH300 | -96.3 | -97.8 | 1.5 | LTE_20261127_144139.txt |
| RF3+RF2 | DRX | 高 CH550 | -96.3 | -97.9 | 1.6 | LTE_20261127_144139.txt |
| RF3+RF2 | PRX+DRX | 低 CH50 | -96.3 | -100.9 | 4.6 | LTE_20261127_144040.txt |
| RF3+RF2 | PRX+DRX | 中 CH300 | -96.3 | -100.8 | 4.5 | LTE_20261127_144040.txt |
| RF3+RF2 | PRX+DRX | 高 CH550 | -96.3 | -100.9 | 4.6 | LTE_20261127_144040.txt |
| RF3+RF4 | PRX | 低 CH50 | -96.3 | -97.7 | 1.4 | LTE_20261127_143058.txt |
| RF3+RF4 | PRX | 中 CH300 | -96.3 | -97.8 | 1.5 | LTE_20261127_143058.txt |
| RF3+RF4 | PRX | 高 CH550 | -96.3 | -97.8 | 1.5 | LTE_20261127_143058.txt |
| RF3+RF4 | DRX | 低 CH50 | -96.3 | -98.2 | 1.9 | LTE_20261127_142959.txt |
| RF3+RF4 | DRX | 中 CH300 | -96.3 | -97.9 | 1.6 | LTE_20261127_142959.txt |
| RF3+RF4 | DRX | 高 CH550 | -96.3 | -98.0 | 1.7 | LTE_20261127_142959.txt |
| RF3+RF4 | PRX+DRX | 低 CH50 | -96.3 | -101.1 | 4.8 | LTE_20261127_142859.txt |
| RF3+RF4 | PRX+DRX | 中 CH300 | -96.3 | -100.9 | 4.6 | LTE_20261127_142859.txt |
| RF3+RF4 | PRX+DRX | 高 CH550 | -96.3 | -101.1 | 4.8 | LTE_20261127_142859.txt |

六套组合全部达标，逐条余量范围 1.3~4.8 dB。完整的下行/上行路径、起止时间与耗时见
`result/RF六组合灵敏度汇总_20261127.xlsx` 的「灵敏度明细」工作表（54 行）；
早上 10 点的旧口径报告未纳入。
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
