#!/usr/bin/env python3
"""
从 UUAtO2 读取全部每日数据 + 从 KFTIUP 读取主播班次数据，供前端仪表盘使用

敏感配置通过环境变量传入（供 GitHub Actions 使用）：
  LARK_CLI                lark-cli 可执行文件路径（默认本机路径）
  DASHBOARD_SHEET_TOKEN   仪表盘数据源表 token（UUAtO2）
  HANDOFF_SHEET_TOKEN     主播班次表 token（KFTIUP）
身份固定使用 bot（--as bot），便于在 CI 中运行。
"""
import json, subprocess, sys, os, re, ast
from datetime import datetime

LARK_CLI = os.environ.get("LARK_CLI", "/Users/apple/Documents/Codex/2026-06-03/cli/lark-cli")
TOKEN = os.environ.get("DASHBOARD_SHEET_TOKEN", "Jos6sfYSRh4eWXtZalBcoFetnCe")
HANDOFF = os.environ.get("HANDOFF_SHEET_TOKEN", "shtcnwOdgFZCQAf4ZjiR5egkoTc")
HANDOFF_SHEET = "KFTIUP"
OUTPUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dashboard_data.json")

def lark_read(token, range_expr):
    cmd = [LARK_CLI, "sheets", "+read", "--as", "bot",
           "--spreadsheet-token", token,
           "--range", range_expr,
           "--format", "json"]
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

def fetch_anchors():
    """Read latest anchor data from KFTIUP handoff sheet."""
    values = lark_read(HANDOFF, f"{HANDOFF_SHEET}!A:I")
    if not values:
        return []

    # Build cell map for formula resolution
    cells = {}
    for r_idx, row in enumerate(values, start=1):
        for c_idx, val in enumerate(row):
            cells[f"{col_num(c_idx)}{r_idx}"] = val

    # Find latest block with at least 5 rows (name/hours/GMV/cost/ROI)
   # Iterate backwards to find the last complete block
    i = len(values) - 1
    while i >= 0:
        row = values[i]
        if row and pn(row[0]) is not None and pn(row[0]) > 0:
            if i + 4 < len(values):
                # Check if any anchor names exist (B, D, F, H columns)
                has_name = False
                for co in (1, 3, 5, 7):
                    if len(row) > co + 1 and row[co + 1] is not None and str(row[co + 1]).strip() not in ('', 'None'):
                        has_name = True
                        break
                if has_name:
                    # Also check GMV row (i+2) has actual data
                    gmv_row = values[i + 2] if i + 2 < len(values) else []
                    has_gmv = False
                    for co in (2, 4, 6, 8):
                        if len(gmv_row) > co and gmv_row[co] is not None and str(gmv_row[co]).strip() not in ('', 'None'):
                            has_gmv = True
                            break
                    if has_gmv:
                        break
        i -= 1
    else:
        return []  # no complete block found

    latest_start = i
    anchors = []
    for col_offset in (1, 3, 5, 7):  # B, D, F, H
        name_row = values[latest_start]
        name = name_row[col_offset + 1] if len(name_row) > col_offset + 1 else None
        if not name or not name.strip():
            continue
        name_str = str(name).strip().replace("冲饮主播：", "")
        if not name_str:
            continue
        if name_str.startswith("主播"):
            continue

        # Hours
        hrs = 0
        if latest_start + 1 < len(values):
            hrs_row = values[latest_start + 1]
            if len(hrs_row) > col_offset + 1:
                hrs = pn(hrs_row[col_offset + 1]) or 0

        # GMV
        gmv = 0
        if latest_start + 2 < len(values):
            gmv_row = values[latest_start + 2]
            if len(gmv_row) > col_offset + 1:
                gmv_raw = gmv_row[col_offset + 1]
                gmv = rounded(eval_formula(gmv_raw, cells), 2)

        # Cost
        cost = 0
        if latest_start + 3 < len(values):
            cost_row = values[latest_start + 3]
            if len(cost_row) > col_offset + 1:
                cost_raw = cost_row[col_offset + 1]
                cost = rounded(eval_formula(cost_raw, cells), 2)

        anchors.append({
            "name": name_str,
            "gmv": gmv,
            "spend": cost,
            "roi": round(gmv / cost, 2) if cost > 0 else 0,
            "hours": hrs,
        })

    return anchors

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
            known.append({"day": day_num, "dv": dv, "sp": sp or 0})
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
    out["anchors"] = fetch_anchors()

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
