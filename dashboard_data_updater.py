#!/usr/bin/env python3
"""
从 UUAtO2 读取全部每日数据 + 从 KFTIUP 读取主播班次数据，供前端仪表盘使用

敏感配置通过环境变量传入（供 GitHub Actions 使用）：
  LARK_CLI                lark-cli 可执行文件路径（默认本机路径）
  DASHBOARD_SHEET_TOKEN   仪表盘数据源表 token（UUAtO2）
  HANDOFF_SHEET_TOKEN     主播班次表 token（KFTIUP）
身份固定使用 bot（--as bot），便于在 CI 中运行。

漏斗数据（funnel）从本地最新「直播明细_全部账号_*.xlsx」提取；CI 云端无
本地 Excel 时保留上一版值（不覆盖）。
"""
import json, subprocess, sys, os, re, ast, glob
from datetime import datetime, date

LARK_CLI = os.environ.get("LARK_CLI", "/Users/apple/Documents/Codex/2026-06-03/cli/lark-cli")
TOKEN = os.environ.get("DASHBOARD_SHEET_TOKEN", "Jos6sfYSRh4eWXtZalBcoFetnCe")
HANDOFF = os.environ.get("HANDOFF_SHEET_TOKEN", "shtcnwOdgFZCQAf4ZjiR5egkoTc")
HANDOFF_SHEET = "KFTIUP"
OUTPUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dashboard_data.json")

def lark_read(token, range_expr, value_render=None):
    cmd = [LARK_CLI, "sheets", "+read", "--as", "bot",
           "--spreadsheet-token", token,
           "--range", range_expr,
           "--format", "json"]
    if value_render:
        cmd += ["--value-render-option", value_render]
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=15)
        data = json.loads(r.stdout)
        if data.get("ok") and "data" in data and "valueRange" in data["data"]:
            return data["data"]["valueRange"]["values"]
        else:
            print(f"[warn] lark read failed: {data.get('error', r.stderr[:200])}", file=sys.stderr)
            return None
    except Exception as e:
        print(f"[error] {e}", file=sys.stderr)
        return None

def pn(v):
    if v is None or v == "" or v == " ":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "")
    try:
        return float(s)
    except:
        return None

def col_num(index):
    """0-based column index to letter(s)."""
    result = ""
    n = index + 1
    while n:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result

def eval_formula(value, cells, visiting=None):
    """Evaluate formula strings like '30001-C3' or 'C3/C4' using resolved cell values."""
    if not isinstance(value, str):
        return pn(value) or 0.0
    expr = value.strip()
    if not expr:
        return 0.0
    if visiting is None:
        visiting = set()

    def replace_ref(match):
        ref = match.group(0)
        # Skip cross-sheet refs like '阴山支援主播业绩表格'!V29
        if "!" in ref:
            return "0"
        if ref in visiting:
            return "0"
        visiting.add(ref)
        resolved = eval_formula(cells.get(ref), cells, visiting)
        visiting.remove(ref)
        return str(resolved)

    expr = re.sub(r"\b[A-Z]{1,3}[0-9]+\b", replace_ref, expr)
    if not re.fullmatch(r"[0-9+\-*/(). ]+", expr):
        return 0.0
    try:
        tree = ast.parse(expr, mode="eval")
        for node in ast.walk(tree):
            if not isinstance(node, (
                ast.Expression, ast.BinOp, ast.UnaryOp,
                ast.Add, ast.Sub, ast.Mult, ast.Div,
                ast.USub, ast.UAdd, ast.Constant, ast.Load,
            )):
                return 0.0
        return float(eval(compile(tree, "<formula>", "eval"), {"__builtins__": {}}, {}))
    except:
        return 0.0

def excel_serial(day):
    """date -> Excel serial number."""
    from datetime import date
    return (day - date(1899, 12, 30)).days


def serial_to_date(serial):
    """Excel serial number -> date."""
    from datetime import date
    return date(1899, 12, 30) + __import__("datetime").timedelta(days=int(serial))


def parse_anchor_block(values, start_idx):
    """Parse one 5-row anchor block starting at start_idx. Returns list of anchors."""
    anchors = []
    for col_offset in (1, 3, 5, 7):  # B, D, F, H
        name_row = values[start_idx]
        if len(name_row) <= col_offset + 1:
            continue
        name = name_row[col_offset + 1]
        if not name or not str(name).strip():
            continue
        name_str = str(name).strip().replace("冲饮主播：", "").strip()
        if not name_str or name_str.startswith("主播"):
            continue

        hrs = 0
        if start_idx + 1 < len(values):
            hrs_row = values[start_idx + 1]
            if len(hrs_row) > col_offset + 1:
                hrs = pn(hrs_row[col_offset + 1]) or 0

        gmv = 0
        if start_idx + 2 < len(values):
            gmv_row = values[start_idx + 2]
            if len(gmv_row) > col_offset + 1:
                gmv = rounded(pn(gmv_row[col_offset + 1]) or 0, 2)

        cost = 0
        if start_idx + 3 < len(values):
            cost_row = values[start_idx + 3]
            if len(cost_row) > col_offset + 1:
                cost = rounded(pn(cost_row[col_offset + 1]) or 0, 2)

        if gmv == 0 and cost == 0:
            continue  # 无数据的空主播位跳过

        anchors.append({
            "name": name_str,
            "gmv": gmv,
            "spend": cost,
            "roi": round(gmv / cost, 2) if cost > 0 else 0,
            "hours": hrs,
        })
    return anchors


def fetch_anchor_blocks():
    """读取 KFTIUP 全部日期块（FormattedValue 直接取公式计算值）。
    返回 {serial: anchors_list}，按日期升序。
    A 列在 FormattedValue 模式下可能是日期字符串（如 2026/03/15）或 serial 数字，均需兼容。"""
    values = lark_read(HANDOFF, f"{HANDOFF_SHEET}!A:I", value_render="FormattedValue")
    if not values:
        return {}

    def to_serial(v):
        """把 A 列值转成 Excel serial。支持数字或日期字符串。"""
        n = pn(v)
        if n is not None and n > 0:
            return int(n)
        if isinstance(v, str) and "/" in v:
            try:
                import datetime as _dt
                d = _dt.datetime.strptime(v.strip(), "%Y/%m/%d").date()
                return excel_serial(d)
            except Exception:
                return None
        return None

    blocks = {}
    i = 0
    while i < len(values):
        row = values[i]
        serial = to_serial(row[0]) if row else None
        if serial is not None and serial > 0 and i + 4 < len(values):
            anchors = parse_anchor_block(values, i)
            if anchors:
                blocks[serial] = anchors
            i += 5
        else:
            i += 1
    return blocks


def fetch_anchors(target_serial=None):
    """读取指定日期（Excel serial）的主播数据。默认取 blocks 中 <= 最新一天的最近块。
    注意：KFTIUP 中可能存在"当日"（数据未填完）块，因此默认按调用方传入的
    target_serial（通常为最新完整数据日 = 前一日）取数，避免取到当日不完整数据。"""
    blocks = fetch_anchor_blocks()
    if not blocks:
        return []
    if target_serial is not None and int(target_serial) in blocks:
        return blocks[int(target_serial)]
    # 回退：取所有块中 serial 最大的
    last = max(blocks.keys())
    return blocks[last]


def fetch_anchor_history():
    """全部日期块的主播历史数据，转成 [{serial, date, anchors:[...]}] 升序。
    date 为 YYYY-MM-DD 字符串，供前端历史查询按日期范围筛选。"""
    blocks = fetch_anchor_blocks()
    if not blocks:
        return []
    history = []
    for serial in sorted(blocks.keys()):
        day = serial_to_date(serial)
        history.append({
            "serial": serial,
            "date": day.isoformat(),
            "anchors": blocks[serial],
        })
    return history


def fetch_funnel_from_excel():
    """从本地最新「直播明细_全部账号_*.xlsx」提取阴山优麦冲饮旗舰店直播漏斗数据。
    返回 dict（曝光/观看/商品曝光/商品点击/成交人数 + 各转化率），无文件时返回 None。
    列：曝光人数=6, 观看人数=8, 商品曝光人数=20, 商品点击人数=21, 成交人数=29。
    """
    try:
        from openpyxl import load_workbook
        base_dir = os.path.dirname(os.path.abspath(__file__))
        files = sorted(glob.glob(os.path.join(base_dir, "直播明细_全部账号_*.xlsx")))
        if not files:
            return None
        wb = load_workbook(files[-1], data_only=True, read_only=True)
        ws = wb["直播间明细"]
        f = {"exposure": 0.0, "views": 0.0, "prod_exposure": 0.0, "prod_click": 0.0, "buyers": 0.0}
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row and row[1] == "阴山优麦冲饮旗舰店":
                f["exposure"] += float(row[6] or 0)
                f["views"] += float(row[8] or 0)
                f["prod_exposure"] += float(row[20] or 0)
                f["prod_click"] += float(row[21] or 0)
                f["buyers"] += float(row[29] or 0)
        wb.close()
        if f["exposure"] <= 0:
            return None
        def rate(a, b):
            return round(a / b * 100, 2) if b else 0
        # 文件名形如 直播明细_全部账号_20260820_20260820.xlsx，取中间 8 位日期
        import re as _re
        m = _re.search(r"(\d{8})_", files[-1])
        return {
            **f,
            "rate_view": rate(f["views"], f["exposure"]),          # 曝光-观看率
            "rate_prod_exposure": rate(f["prod_exposure"], f["views"]),  # 观看-商品曝光率
            "rate_click": rate(f["prod_click"], f["prod_exposure"]),      # 商品曝光-点击率
            "rate_buy": rate(f["buyers"], f["prod_click"]),              # 商品点击-成交转化率
            "rate_total": rate(f["buyers"], f["exposure"]),              # 曝光-成交转化率
            "date": m.group(1) if m else "",
        }
    except Exception as e:
        print(f"[warn] funnel excel: {e}", file=sys.stderr)
        return None


def rounded(v, d=2):
    return round(float(v) + 1e-9, d)

def fetch_data():
    rows = lark_read(TOKEN, "UUAtO2!A7:O38")
    if rows is None:
        return None

    out = {}

    # 所有已知日期数据
    known = []
    postal_hist = []
    latest_idx = None
    for i, row in enumerate(rows):
        dv = pn(row[1]) if len(row) > 1 else None
        sp = pn(row[2]) if len(row) > 2 else None
        if dv is not None:
            day_num = i + 1
            known.append({
                "day": day_num,
                "dv": dv,
                "sp": sp or 0,
                "live": pn(row[3]) if len(row) > 3 else None,
                "short_video": pn(row[4]) if len(row) > 4 else None,
                "card": pn(row[5]) if len(row) > 5 else None,
                "other": pn(row[6]) if len(row) > 6 else None,
                "graphic": pn(row[7]) if len(row) > 7 else None,
            })
            latest_idx = i

            pg = pn(row[9]) if len(row) > 9 else None
            ps = pn(row[10]) if len(row) > 10 else None
            if pg is not None and pg > 0:
                postal_hist.append({"day": day_num, "pg": pg, "ps": ps or 0})

    out["knownDays"] = known
    out["postalHistory"] = postal_hist

    # 最新一天的明细
    if latest_idx is not None:
        row = rows[latest_idx]
        out["latest"] = {
            "day": latest_idx + 1,
            "b": pn(row[1]) if len(row) > 1 else None,
            "c": pn(row[2]) if len(row) > 2 else None,
            "d": pn(row[3]) if len(row) > 3 else None,
            "e": pn(row[4]) if len(row) > 4 else None,
            "f": pn(row[5]) if len(row) > 5 else None,
            "g": pn(row[6]) if len(row) > 6 else None,
            "h": pn(row[7]) if len(row) > 7 else None,
        }
        pg = pn(row[9]) if len(row) > 9 else None
        ps = pn(row[10]) if len(row) > 10 else None
        out["postal"] = {
            "gmv": pg if pg else 0,
            "spend": ps if ps else 0,
            "roi": round(pg / ps, 2) if (pg is not None and ps and ps > 0) else 0
        }
        vg = pn(row[12]) if len(row) > 12 else None
        vs = pn(row[13]) if len(row) > 13 else None
        out["video"] = {
            "gmv": vg if vg else 0,
            "spend": vs if vs else 0,
            "roi": round(vg / vs, 2) if (vg is not None and vs and vs > 0) else 0
        }
    else:
        out["latest"] = {"day": 0}
        out["postal"] = {"gmv": 0, "spend": 0, "roi": 0}
        out["video"] = {"gmv": 0, "spend": 0, "roi": 0}

    # 主播班次数据（从 KFTIUP 读取）
    # target_serial = 最新完整数据日的 Excel serial（取 knownDays 最新一天对应日期）
    # 需求：主播栏维持前一日数据，不取 KFTIUP 中"当日"（可能未填完）的块
    if latest_idx is not None:
        latest_date_serial = pn(rows[latest_idx][0])
    else:
        latest_date_serial = None
    out["anchors"] = fetch_anchors(latest_date_serial)

    # 历史主播数据（全部日期块，供历史查询页使用）
    out["anchorHistory"] = fetch_anchor_history()

    # 读取月度汇总数据 from specific cells
    cmd2 = [LARK_CLI, "sheets", "+read", "--as", "bot",
           "--spreadsheet-token", TOKEN,
           "--range", "UUAtO2!J38:O38",
           "--value-render-option", "UnformattedValue",
           "--format", "json"]
    try:
        r2 = subprocess.run(cmd2, capture_output=True, text=True, timeout=15)
        d2 = json.loads(r2.stdout)
        if d2.get("ok") and "data" in d2 and "valueRange" in d2["data"]:
            sr = d2["data"]["valueRange"]["values"][0]
            out["postalMonthly"] = {
                "gmv": pn(sr[0]) if len(sr) > 0 and sr[0] is not None else 0,
                "spend": pn(sr[1]) if len(sr) > 1 and sr[1] is not None else 0,
                "roi": pn(sr[2]) if len(sr) > 2 and sr[2] is not None else 0,
            }
            out["videoMonthly"] = {
                "gmv": pn(sr[3]) if len(sr) > 3 and sr[3] is not None else 0,
                "spend": pn(sr[4]) if len(sr) > 4 and sr[4] is not None else 0,
                "roi": pn(sr[5]) if len(sr) > 5 and sr[5] is not None else 0,
            }
        else:
            out["videoMonthly"] = {"gmv": 0, "spend": 0, "roi": 0}
            out["postalMonthly"] = {"gmv": 0, "spend": 0, "roi": 0}
    except Exception as e:
        print(f"[warn] summary cells: {e}", file=sys.stderr)
        out["videoMonthly"] = {"gmv": 0, "spend": 0, "roi": 0}
        out["postalMonthly"] = {"gmv": 0, "spend": 0, "roi": 0}
    out["targetGmv"] = 2800000
    out["targetRoi"] = 3
    # 漏斗数据：优先本地 Excel（最新日期）；CI 云端无 Excel 时保留上一版值
    funnel = fetch_funnel_from_excel()
    if funnel is not None:
        out["funnel"] = funnel
    elif os.path.exists(OUTPUT):
        try:
            prev = json.load(open(OUTPUT, encoding="utf-8"))
            if prev.get("funnel"):
                out["funnel"] = prev["funnel"]
        except Exception:
            pass
    if "funnel" not in out:
        out["funnel"] = None
    out["fetchedAt"] = subprocess.run(
        ["date", "+%Y-%m-%d %H:%M:%S"], capture_output=True, text=True
    ).stdout.strip()

    return out

def main():
    print("[dashboard_data_updater] Fetching data from Feishu (UUAtO2 + KFTIUP)...")
    data = fetch_data()
    if data is None:
        print("[dashboard_data_updater] FAILED to fetch data", file=sys.stderr)
        sys.exit(1)
    with open(OUTPUT, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"[dashboard_data_updater] Written {OUTPUT}")
    ld = data.get("latest", {})
    anc = data.get("anchors", [])
    print(f"[dashboard_data_updater] Latest day: {ld.get('day')}, GMV: {ld.get('b')}")
    print(f"[dashboard_data_updater] Anchors: {len(anc)} hosts")
    for a in anc:
        print(f"  {a['name']}: GMV {a['gmv']}, Cost {a['spend']}, ROI {a['roi']}")

if __name__ == "__main__":
    main()
