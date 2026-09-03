---
name: feishu-yinshan-daily-report
description: Fill the Yinshan Youmai Feishu daily data report from local Douyin Excel exports and Qianchuan MCP data. Use when the user asks to 填写阴山优麦数据日报, 填日报, 日报复盘, 前一日/昨日直播明细, 抖音电商罗盘, 千川总消耗, GMV表, 月度GMV大表, or asks Codex to calculate and write Yinshan Youmai daily report values into Feishu sheets.
---

# Yinshan Youmai Feishu Daily Report

## 触发方式（重要）

- 只在用户明确下达命令时执行，等待这些关键词：`填写日报`、`填日报`、`日报复盘`、`前一日/昨日直播明细`、`千川总消耗`、`GMV表`、`月度GMV大表`、`数据日报`。
- 不主动运行，也不按固定时间自动填表；用户没有下命令就不要动飞书表格。
- 日期以用户说法为准，默认取昨天；如用户说“8月27日”，日期就是 `2026-08-27`。

## 标准流程（避免绕弯）

1. 确认本地明细文件已就位：
   - 冲饮店：`阴山冲饮数据明细/直播明细_全部账号_YYYYMMDD_YYYYMMDD.xlsx`、`阴山冲饮数据明细/抖音电商罗盘-成交分析-YYYYMMDD-YYYYMMDD.xlsx`
   - 食品店：`阴山食品店数据明细/` 下同名文件，文件名可能带 `(1)` 后缀
   - 脚本会自动在这些子目录里找文件，不需要把文件移到根目录。

2. 用巨量引擎 MCP 查询千川消耗：
   - 工具：`qianchuan_report_uni_promotion_get_v1`
   - 日期必须带时间：`start_date="YYYY-MM-DD 00:00:00"`、`end_date="YYYY-MM-DD 23:59:59"`
   - 字段：`["stat_cost", "total_pay_order_gmv_include_coupon_for_roi2"]`
   - 参数：`marketing_goal="ALL"`、`order_platform="ALL"`
   - 取返回值里的 `stat_cost` 作为当日千川消耗。
   - 关键广告主 ID（2026-08-28 确认）：冲饮店 `1842329157537987`，食品店 `1867784441852039`。
   - 如果接口提示账户角色不对，先用 `oauth2_advertiser_get` 拿授权账户，冲饮再用 `customer_center_account_list_v3`，食品再用 `qianchuan_shop_advertiser_list_v1` 拿真正的千川广告主 ID。

3. 运行脚本写表：
   ```bash
   python3 skills/feishu-yinshan-daily-report/scripts/yinshan_daily_report.py \
     --date YYYY-MM-DD --shop drink --qianchuan-cost <冲饮千川消耗> --mode write

   python3 skills/feishu-yinshan-daily-report/scripts/yinshan_daily_report.py \
     --date YYYY-MM-DD --shop food --qianchuan-cost <食品千川消耗> --mode write
   ```
   - `--shop drink` 写冲饮店；`--shop food` 写食品店。
   - 脚本会自动写：冲饮 `UUAtO2`、`YKWQjE`、`bea4e1`、`cC79qR`；食品 `9VVamg`、`YKWQjE`、`weIxvN`。
   - 直播明细表已存在当日数据时会自动跳过，不会重复追加。
   - 脚本联网写飞书需要授权；沙箱内失败就升级网络权限后重跑。

4. 回读验证写好的单元格（见下方验证命令）。

5. 同步公网驾驶舱：
   - 脚本 `--mode write` 会自动刷新本地 `dashboard_data.json`。
   - 提交并推送：
     ```bash
     git add dashboard_data.json dashboard_data_updater.py
     git commit -m "chore: refresh dashboard data [$(date -u +'%Y-%m-%d %H:%M UTC')]"
     git push origin main
     ```
   - GitHub 直连失败时走本机代理：
     ```bash
     HTTPS_PROXY=http://127.0.0.1:7897 git push origin main
     ```
   - GitHub Pages 发布通常需要 1-2 分钟；验证地址：
     `https://521080386-bot.github.io/yinshan-dashboard/dashboard_data.json`
   - 只提交本次改动的数据/解析文件，不要顺手提交用户的其他本地改动。

## Sheet Tokens

| Token | Sheet | Purpose |
|-------|-------|---------|
| `Jos6sfYSRh4eWXtZalBcoFetnCe` | `UUAtO2` | 冲饮店店铺数据（仪表盘数据源） |
| `Jos6sfYSRh4eWXtZalBcoFetnCe` | `9VVamg` | 食品店店铺数据 |
| `Jos6sfYSRh4eWXtZalBcoFetnCe` | `cC79qR` | 冲饮店直播间明细 |
| `Jos6sfYSRh4eWXtZalBcoFetnCe` | `weIxvN` | 食品店直播数据表 |
| `shtcnzv0loZQg1qk25PcCVonplW` | `YKWQjE` | 2026-9 月度大表 |
| `shtcniNUd2HqPd8ZvZOt8btOScf` | `bea4e1` | 冲饮店直播间汇总表 |
| `CD2psfSlghdutdt9wHmcOJaDntf` | `cdzpqi` | 食品店交班数据（主播班次） |

## Write Targets

### 冲饮店

- `UUAtO2`：按 Excel 日期序列定位行；9 月 1 日从第 33 行开始。
  - `B` 抖音GMV、`C` 千川消耗（传入 `--qianchuan-cost` 才写）、`D-H` 直播/短视频/商品卡/其他/图文
  - `L` 低GI视频号 ROI；`J/K`（低GI视频号GMV/消耗）由用户手动维护，脚本不写
  - `I` 当日ROI 由公式计算，不写
- `YKWQjE`：行 = `day + 4`
  - `D` 抖音GMV、`F` 千川/全店消耗、`Q/R/S` 邮政 GMV / 消耗 / ROI
- `bea4e1`：按 Excel 日期序列定位行
  - `B` 场观、`C` 直播间成交金额、`E` 直播时长小时、`F` 投放消耗、`G` 投放GMV
  - `H/I/J` 最高在线 / 平均在线 / 平均停留秒
  - `K/L/M/N` 成交件数 / 成交人数 / 观看-成交率 / 商品点击-成交率
- `cC79qR`：自动追加当日原始直播明细行

### 食品店

- `9VVamg`：按 Excel 日期序列定位行
  - `B` 抖音GMV、`C` 千川消耗、`D-H` 直播/短视频/商品卡/其他/图文、`I` 当日ROI
- `YKWQjE`：行 = `day + 4`
  - `K` 抖音GMV、`M` 千川/全店消耗；`J` 总GMV含智能券由 `K+L` 公式自动计算
- `weIxvN`：自动追加当日原始直播明细行

## 主播数据与算式

- 食品店交班表 `cdzpqi` 的 GMV/消耗可能写成算式，例如 `15874-9088`、`8182-4992`。
- `dashboard_data_updater.py` 已支持解析这类简单算式；如果某主播在驾驶舱缺失，先检查交班表对应格子是不是算式，再刷新驾驶舱。

## 数据来源与业务逻辑

- 直播明细：`直播间明细` 表，账号名在 B 列，GMV 在第 26 列，投放消耗在第 44 列。
- 罗盘：`自营成交` 表，取目标日期、`不限` 投放时段、`成交金额` 列（脚本按表头自动识别列位置，冲饮/食品罗盘列序不同也能兼容）。
- 冲饮全店 GMV = 各载体成交金额之和；YKWQjE `D` 再扣除燕麦片/燕麦米账号 GMV。
- ROI = GMV / 千川消耗，保留 2 位小数；消耗为 0 时写 0。

## 验证

```bash
lark-cli sheets +read --as user --spreadsheet-token "Jos6sfYSRh4eWXtZalBcoFetnCe" --range "UUAtO2!A28:O28" --value-render-option FormattedValue --format json
lark-cli sheets +read --as user --spreadsheet-token "Jos6sfYSRh4eWXtZalBcoFetnCe" --range "9VVamg!A240:I240" --value-render-option FormattedValue --format json
lark-cli sheets +read --as user --spreadsheet-token "shtcnzv0loZQg1qk25PcCVonplW" --range "YKWQjE!A5:CR5" --value-render-option FormattedValue --format json
lark-cli sheets +read --as user --spreadsheet-token "shtcniNUd2HqPd8ZvZOt8btOScf" --range "bea4e1!B625:N625" --value-render-option FormattedValue --format json
```

## 注意事项

- 飞书读写需要网络和用户身份授权；沙箱报错时用升级权限重跑。
- 本地明细文件缺失时脚本会告警并返回空数据，不要用空值覆盖已填日期。
- 推送公网前确认 `dashboard_data.json` 最新日期正确，再 commit/push。
