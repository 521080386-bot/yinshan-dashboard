---
name: feishu-yinshan-daily-report
description: Fill the Yinshan Youmai Feishu daily data report from local Douyin Excel exports. Use when the user asks to 填写阴山优麦数据日报, 数据日报复盘, 前一日/昨日直播明细, 抖音电商罗盘, 千川总消耗, GMV表, 月度GMV大表, or asks Codex to calculate and write Yinshan Youmai daily report values into Feishu sheets.
---

# Yinshan Youmai Feishu Daily Report

## Quick Start

Use `scripts/yinshan_daily_report.py` to read local Excel files, calculate GMV/cost/ROI, and write to Feishu sheets:

```bash
python3 skills/feishu-yinshan-daily-report/scripts/yinshan_daily_report.py --date YYYY-MM-DD --mode write
```

If the script cannot run due to sandboxing, escalate network permissions for `feishu.cn`.

Default target date is yesterday. The script expects these files in `cwd`:

- `直播明细_全部账号_YYYYMMDD_YYYYMMDD.xlsx`
- `抖音电商罗盘-成交分析-YYYYMMDD-YYYYMMDD.xlsx`

## Sheet Tokens

Defined at the top of `yinshan_daily_report.py`:

| Token | Sheet | Purpose |
|-------|-------|---------|
| `Jos6sfYSRh4eWXtZalBcoFetnCe` | `UUAtO2` | 仪表盘数据源 — daily per-row data |
| `shtcnzv0loZQg1qk25PcCVonplW` | `6j5jOY` | 月度GMV大表 |
| `shtcniNUd2HqPd8ZvZOt8btOScf` | `bea4e1` | 阴山优麦冲饮旗舰店直播间明细 |

Lark CLI at `/Users/apple/Documents/Codex/2026-06-03/cli/lark-cli`.

## Write Targets

### UUAtO2 — Dashboard Data Source

Row = `day_of_month + 6` (row 6 is header, row 7 = 7月1日, row 26 = 7月20日). Columns:

| Column | Field | Source |
|--------|-------|--------|
| B | 抖音GMV | `report_total_gmv` |
| C | 千川消耗 | **人工填写，脚本不再写入**（2026-08-19 起） |
| D | 直播GMV | carriers["直播"].成交金额 |
| E | 短视频GMV | carriers["短视频"].成交金额 |
| F | 商品卡 | carriers["商品卡"].成交金额 |
| G | 其他 | carriers["其他"].成交金额 |
| H | 图文 | carriers["图文"].成交金额 |
| J | 邮政GMV | postal account GMV |
| K | 邮政千川消耗 | postal account cost |
| L | 邮政ROI | GMV / cost |

Columns C (千川消耗, 人工填写), I (当日ROI), M (视频号GMV), N (视频号消耗), O (视频号ROI) are manual input — the script does not write them.

### 6j5jOY — Monthly GMV Table

Row = `day_of_month + 4`. Columns:

- `D`: full GMV (total minus oat accounts)
- `O/P/Q`: postal GMV / cost / ROI
- `CM/CN/CO`: oat_rice GMV / cost / ROI (阴山优麦有机燕麦米)
- `CP/CQ/CR`: oat_flake GMV / cost / ROI (阴山有机-宿州)

### bea4e1 — Live Stream Detail

Write to the row matching the target date's Excel serial number (column A). Fields:

| Column | Field |
|--------|-------|
| B | 直播间观看次数 |
| C | 直播间成交金额 |
| E | 直播时长（小时） |
| F | 投放消耗 |
| G | 投放GMV (= 成交金额) |
| H | 最高在线 |
| I | 平均在线 |
| J | 平均停留（秒） |
| K | 成交件数 |
| L | 成交人数 |
| M | 观看-成交率(人数) |
| N | 商品点击-成交率(人数) |

Volume metrics (场观/成交金额/时长/消耗/件数/人数) → **sum** across multiple records.
Rate/peak metrics → **max**.

## Data Sources & Business Logic

### Account List

```python
ACCOUNTS = [
    "阴山优麦冲饮旗舰店",
    "阴山优麦有机燕麦片",
    "阴山有机-宿州",
    "阴山优麦有机燕麦米",
    "蒙邮优选",
    "中国邮政乌兰察布分公司",
    "中国邮政集团有限公司乌兰察布市分公司",
]
```

- **oat_flake** = `阴山有机-宿州` or `阴山优麦有机燕麦片`
- **oat_rice** = `阴山优麦有机燕麦米`
- **postal** = `蒙邮优选` or `中国邮政乌兰察布分公司` or `中国邮政集团有限公司乌兰察布市分公司`
- **total_cost** = sum of all ACCOUNTS' cost

### Live Data (`load_live`)

Reads `直播明细_全部账号_{ymd}_{ymd}.xlsx`, sheet `直播间明细`. Sums GMV (col 25) and cost (col 43) per account name (col 1).

### Compass Data (`load_compass`)

Reads `抖音电商罗盘-成交分析-{ymd}-{ymd}.xlsx`, sheet `自营成交`. For the target date, extracts `成交金额` by carrier (`直播`, `短视频`, `商品卡`, `图文`, `其他`).

### Calculations

- `report_total_gmv` = sum(carrier `成交金额`)
- `full_gmv` = `report_total_gmv` − oat_flake GMV − oat_rice GMV
- ROI = GMV / cost (rounded to 2 decimals; 0 if cost is 0)
- If an account has no live data, write 0 for GMV, cost, and ROI

## Dashboard Data Refresh

Since 2026-08-19, `--mode write` **automatically refreshes the dashboard** by
running `dashboard_data_updater.py` after writing (it reads all rows from UUAtO2
and writes `dashboard_data.json` consumed by `index.html`). No separate refresh
step is needed after filling a daily report.

If you need to refresh the dashboard manually (e.g. after a manual cell edit):

```bash
python3 dashboard_data_updater.py
```

## Validation

Read back the written cells:

```bash
lark-cli sheets +read --spreadsheet-token "Jos6sfYSRh4eWXtZalBcoFetnCe" --range "UUAtO2!A26:O26" --format json
lark-cli sheets +read --spreadsheet-token "shtcniNUd2HqPd8ZvZOt8btOScf" --range "bea4e1!A:N" --format json
lark-cli sheets +read --spreadsheet-token "shtcnzv0loZQg1qk25PcCVonplW" --range "6j5jOY!D7:CR7" --format json
```

## Troubleshooting

### Network / Auth failures

The lark-cli needs network access to `feishu.cn`. Use `--mode write` with escalated permissions (`require_escalated`). Auth auto-refreshes on first call.

### Missing Excel files

If 直播明细 or 罗盘 file is missing for the target date, the script prints a warning and returns empty data (GMV=0, cost=0).

### Verification

Check that:
- `UUAtO2` row matches the expected day number
- `bea4e1` row matches the Excel serial number
- `6j5jOY` row matches day_of_month + 4
