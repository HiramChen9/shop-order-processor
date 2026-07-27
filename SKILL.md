---
name: process-shop-orders
description: >
  处理Shopee、Lazada、TikTok Shop导出的订单Excel文件，
  包含筛选有效订单、印尼数值格式转换、SKU匹配成本和运费、
  成本合计/运费合计计算、Shopee黑底优惠券清零、每日统计实际销售金额等。
  当用户提供电商平台导出的订单Excel文件时触发，文件名通常包含
  Shopee/Lazada/TikTok关键字。
---

# 电商订单处理

用户提供三种电商平台导出的订单 Excel 文件，
运行 scripts/process_orders.py 自动处理并生成筛选后的文件。

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

## SKU 成本运费匹配规则
(去除 -ST 后缀后按顺序匹配, 匹配到即返回)

| 模式 | 成本 | 运费 |
|---|---|---|
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

不匹配的留空。

## 输出格式
- 两个工作表: 筛选后 + 每日统计
- 蓝色表头 (#4472C4), 细边框
- 金额列千分位, 成本/运费保留两位小数
- 自动调整列宽

## 自定义 SKU 规则
编辑 scripts/process_orders.py 中的 SKU_RULES 列表，按顺序匹配。
 
## SKU 成本运费匹配规则
(去除 -ST 后缀后按顺序匹配, 匹配到即返回)

| 模式 | 成本 | 运费 |
|---|---|---|
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
