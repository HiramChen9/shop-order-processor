---
name: process-shop-orders
description: >
  处理Shopee、Lazada、TikTok Shop导出的订单Excel文件，
  包含筛选有效订单、印尼数值格式转换、SKU匹配成本和运费、
  成本合计/运费合计计算、Shopee黑底优惠券清零、每日统计实际销售金额，
  并优先在两分钟内生成结果，在对话中直接报告每日订单数、件数、净销售额、成本和运费。
  当用户提供电商平台导出的订单Excel文件时触发，文件名通常包含
  Shopee/Lazada/TikTok关键字。
---

# 电商订单处理

用户提供三种电商平台导出的订单 Excel 文件，
运行 scripts/process_orders.py 自动处理并生成筛选后的文件。

当用户同时上传多个平台文件时，逐个运行脚本，不要合并原始文件。

## 两分钟快速路径

对于可正常读取、表头符合 Shopee、Lazada 或 TikTok Shop 导出格式的 1–3 个订单文件，以从取得可访问文件路径开始 120 秒内返回结果为目标：

- 立即运行本 Skill 的 `scripts/process_orders.py`；标准处理不再加载或执行通用电子表格创建、重排、渲染或视觉设计流程。
- 多个文件彼此独立时并行运行；每个文件单独生成结果，不合并原始订单。
- 只做必要校验：进程成功退出、输出文件存在且非空、包含“筛选后”和“每日统计”两个工作表、成功取得 `DAILY_SUMMARY`。
- 不预先渲染源文件，也不逐页渲染结果；仅当脚本报错、平台无法识别、必要列缺失，或用户明确要求视觉检查时，才检查对应文件。
- 处理成功后立即返回下载文件和对话框汇总，不为可选美化或额外分析延迟交付。
- 若个别文件无法在时限内处理，先交付其他成功文件，并在 120 秒内指出失败文件和具体错误；不要让单个异常文件阻塞全部结果。

这套专用脚本已经定义筛选、计算、格式和校验逻辑。除非用户要求自定义工作簿结构，否则不需要另行重建工作簿。

## 使用方式

python <skill_path>/scripts/process_orders.py <platform> <input_file> [output_file]

| 参数 | 说明 |
|---|---|
| platform | shopee / lazada / tiktok / auto (默认auto, 自动识别) |
| input_file | 用户给的文件路径 |
| output_file | (可选) 输出路径，默认同目录下 原文件名_筛选后.xlsx |

示例:
  # 自动识别平台 (auto为默认, 可省略)
  python scripts/process_orders.py "用户给的文件路径/Shopee7.25.xlsx"
  # 或明确指定
  python scripts/process_orders.py shopee "用户给的文件路径/Shopee7.25.xlsx"
  python scripts/process_orders.py lazada "用户给的文件路径/Lazada.xlsx"
  python scripts/process_orders.py tiktok "用户给的文件路径/TikTok.xlsx"

## Shopee 处理规则
- 筛选: 剔除 Batal / Belum Bayar / Pembatalan diajukan
- 保留列: 订单号, 订单创建时间(仅日期), 参考sku编码, 件数, 订单小计, 卖家承担的优惠券
- 新增列: 成本, 运费(件数左侧), 成本合计, 运费合计(件数右侧)
- 数值: 印尼 . 为千分位, 自动去除 (448.500 -> 448500)
- 黑底行 (fill.rgb == 00000000): 优惠券清零
- SKU 去除尾部后缀(如 -ST、-STS 等)后匹配

## Lazada 处理规则
- 筛选: status 剔除 canceled/cancelled/unpaid/batal/belum bayar
- 保留列: sellerSku, createTime(仅日期), orderNumber, paidPrice, status
- 新增列: 成本, 运费, 件数(同一订单号的行数), 成本合计, 运费合计
- 件数 = 1 (每行代表一件)
- 成本合计/运费合计 = 同一单各行的sku成本/运费相加
- paidPrice 去除印尼千分位

## TikTok Shop 处理规则
- 筛选: Order Status 剔除取消/未支付; Normal or Pre-order 筛除空白行
- 保留列: Order ID, Order Status, Normal or Pre-order, Seller SKU, Quantity, SKU Platform Discount, SKU Subtotal After Discount, Payment platform discount, Created Time(仅日期)
- 新增列: 成本, 运费(Quantity左侧), 成本合计, 运费合计(Quantity右侧)
- 金额列去除印尼千分位

## 日期处理
所有时间字段只保留 YYYY-MM-DD, 去除时分秒。

## 每日统计
第二个工作表 "每日统计", 按日期汇总:
- 日期, 订单总数, 件数总数, 金额小计, 优惠券, 成本总额, 运费总额, 实际销售金额
- 成本总额/运费总额 = 对应sku编码的成本和运费 x 对应订单件数
- 实际销售金额:
  * Shopee: 金额小计 - 优惠券(白底优惠券总额)
  * Lazada: paidPrice 合计
  * TikTok: SKU Platform Discount + SKU Subtotal After Discount + Payment platform discount
- 末行: 合计

## 对话框汇总

脚本会为文件中的每个日期输出一行以 `DAILY_SUMMARY ` 开头的 JSON。处理完成后，无需等待用户追问，必须把这些结果直接显示在最终回复中：

| 文件 | 平台 | 日期 | 订单数 | 件数 | 净销售额 | 成本 | 运费 |
|---|---|---|---:|---:|---:|---:|---:|

- “净销售额”等同于“每日统计”中的“实际销售金额”，沿用各平台现有计算规则。
- 一个文件含多个日期时，每个日期单独一行；不要只显示“合计”行。
- 同时上传多个平台文件时，按文件分别显示，不要擅自跨平台合并。
- 净销售额使用千分位；成本和运费使用千分位并保留两位小数。
- 即使已提供处理后的 Excel 下载链接，也不能省略此对话框汇总。

## SKU 成本运费匹配规则
(去除 `-ST`、`-STS` 等尾部字母后缀后按顺序匹配，匹配到即返回；表中 `xxx` 代表任意非空 SKU 内容。)

不匹配的成本和运费留空。

| 模式 | 成本 | 运费 |
|---|---|---|
| JGSxxx-2.8-ST | 13.2924 | 6.300285 |
| ANxxx-3 | 2.8 | 0.57843072 |
| CZTJZH-xxx-60-2.8-ST | 9.78 | 5.3960088 |
| CZT-157-3060*30-1.0-ST | 24.72 | 5.4553056 |
| CZTJZH-xxx-90-2.8-ST | 12.224 | 7.69902 |
| KCxxx-3060-*5-ST | 6.795 | 2.5 |
| KCxxx-3060-*10-ST | 13.59 | 5.0599936 |
| KCxxx-60-2.8-ST | 13.356 | 6.300285 |
| JGSxxx-60-2.8-ST | 13.3428 | 6.300285 |
| CZT-xxx-90-2.8-ST | 12.224 | 7.69902 |
| CZT-xxx-120-2.8-IXPE-1-ST | 15.9132 | 5.9069655 |
| HZT-xxx-2020*36-ST | 12.72 | 0.58997925 |
| CZTWCxxx-3060*20 | 26.40 | 0 |
| CZTWCxxx-3060*10 | 13.20 | 0 |
| BKCZT-xxx-3060*20 | 13.60 | 1.07524864 |
| BKCZT-xxx-3060*10 | 6.80 | 0.53762432 |
| CZT-xxx-4090-1.0*5 | 8.25 | 2.0017452 |
| CZT-xxx-4090*5 | 9.20 | 4.0034904 |
| CZT-xxx-4090 | 1.84 | 0.80069808 |
| CZT-xxx-3060-1.0*20 | 16.48 | 3.6368704 |
| CZT-xxx-3060-1.0*30 | 24.72 | 5.4553056 |
| CZT-xxx-3060*20 | 19.00 | 8.0643648 |
| CZT-xxx-3060*30 | 28.50 | 12.0965472 |
| CZT-xxx-3060-1.0*10 | 8.24 | 1.8184352 |
| CZT-xxx-3060*10 | 9.50 | 4.0321824 |
| CZTZHZ-xxx-xxx | 27.52 | 12.3683242 |
| XZ-xxx-3 | 3.06 | 0.8301552 |
| Nxxx-3 / DExxx-3 / MAxxx-3 | 2.80 | 0.57843072 |
| PPDxxx-8 | 27.21 | 0 |
| PPDxxx-10 | 31.60 | 0 |
| YXX-xxx-45-5 | 2.84 | 0.51131535 |
| ANxxx-5 / DExxx-5 / MAxxx-5 | 4.45 | 0.7230384 |
| CZTJZH-xxx-120-2.8 | 19.28 | 10.549889 |
| MBTxxx-10-5 | 2.83 | 1.42751555555556 |

## 输出格式
- 两个工作表: 筛选后 + 每日统计
- 蓝色表头 (#4472C4), 细边框
- 金额列千分位, 成本/运费保留两位小数
- 自动调整列宽

## 自定义 SKU 规则
编辑 scripts/process_orders.py 中的 SKU_RULES 列表，按顺序匹配。
