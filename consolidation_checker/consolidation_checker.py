"""
Order Consolidation Checker
============================
Reads a Google Sheet order tracker and a NetSuite Item Fulfillment export,
identifies orders that can be consolidated into a single shipment, and writes
the recommendations back to the Sheet's Note column.

Usage:
  1. Export today's IF table from NetSuite as Excel XML (.xls)
  2. Place it in this directory (or set IF_XLS_PATH in .env)
  3. Run:  python consolidation_checker.py [--write | --dry-run]
"""

import pandas as pd
import xml.etree.ElementTree as ET
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta
from itertools import combinations
import sys, os, re

from config import CONFIG


# ══════════════════════════════════════════
# 工具函数
# ══════════════════════════════════════════

def _parse_date(val):
    if not val or str(val).strip() == "":
        return None
    val = str(val).strip()
    for fmt in ("%Y-%m-%dT%H:%M:%S", "%m/%d/%Y", "%Y-%m-%d", "%m/%d/%y", "%m/%d"):
        try:
            d = datetime.strptime(val, fmt)
            # 补年份（%m/%d 格式）
            if d.year == 1900:
                d = d.replace(year=datetime.today().year)
            return d
        except ValueError:
            continue
    return None


def _normalize_so(val):
    if pd.isna(val) or str(val).strip() == "":
        return ""
    return re.sub(r"[^\d]", "", str(val))


def _extract_so_from_created(val):
    if not val:
        return ""
    m = re.search(r"SO(\d+)", str(val), re.IGNORECASE)
    return m.group(1) if m else ""


def _date_diff_days(d1, d2):
    """两个 datetime 相差天数，任一为 None 返回 None"""
    if d1 is None or d2 is None:
        return None
    return abs((d1 - d2).days)


def _fmt_md(d):
    """跨平台日期格式化：月/日，不带前导零（Windows 不支持 %-m/%-d）"""
    if d is None:
        return ""
    return f"{d.month}/{d.day}"


# ══════════════════════════════════════════
# 读取 NetSuite .xls（SpreadsheetML 格式）
# ══════════════════════════════════════════

def load_if_xls(path):
    if not os.path.exists(path):
        print(f"❌ 找不到 IF 表文件：{path}")
        sys.exit(1)
    print(f"📥 读取 IF 表：{path}")
    try:
        tree = ET.parse(path)
    except ET.ParseError as e:
        print(f"❌ IF 表解析失败：{e}")
        print("   请确认从 NetSuite 导出的是 Excel XML 格式（.xls），不是 CSV 或二进制 xlsx。")
        sys.exit(1)

    root = tree.getroot()
    ns = {"ss": "urn:schemas-microsoft-com:office:spreadsheet"}

    worksheets = root.findall(".//ss:Worksheet", ns)
    if not worksheets:
        print("❌ IF 表中找不到任何 Worksheet，请检查导出格式。")
        sys.exit(1)
    ws    = worksheets[0]
    table = ws.find(".//ss:Table", ns)
    if table is None:
        print("❌ IF 表 Worksheet 中找不到 Table，文件可能为空。")
        sys.exit(1)
    rows  = table.findall("ss:Row", ns)
    if len(rows) < 2:
        print("❌ IF 表只有表头或为空，请检查 NetSuite 导出范围。")
        sys.exit(1)

    def get_vals(row):
        cells = row.findall("ss:Cell", ns)
        vals = []
        for cell in cells:
            idx = cell.get("{urn:schemas-microsoft-com:office:spreadsheet}Index")
            if idx:
                while len(vals) < int(idx) - 1:
                    vals.append("")
            data = cell.find("ss:Data", ns)
            vals.append(data.text if data is not None else "")
        return vals

    headers = get_vals(rows[0])
    data = []
    for row in rows[1:]:
        vals = get_vals(row)
        while len(vals) < len(headers):
            vals.append("")
        data.append(vals[:len(headers)])

    df = pd.DataFrame(data, columns=headers)

    # 缺列检查
    required_if_cols = ["Created From", "Date", "Document Number", "Status", "Customer Name", "Shipping Zip"]
    missing = [c for c in required_if_cols if c not in df.columns]
    if missing:
        print(f"❌ IF 表缺少以下列：{missing}")
        print(f"   实际列名：{list(df.columns)}")
        sys.exit(1)

    df["so_norm"]   = df["Created From"].apply(_extract_so_from_created)
    df["pick_date"] = df["Date"].apply(_parse_date)
    df["if_doc"]    = df["Document Number"].astype(str).str.strip()
    df["if_status"] = df["Status"].astype(str).str.strip()

    # 保留关键列
    out = df[["so_norm", "pick_date", "if_doc", "if_status",
              "Customer Name", "Shipping Zip"]].copy()
    out = out[out["so_norm"] != ""]
    print(f"   解析到 {len(out)} 条有效 IF 记录")
    return out


# ══════════════════════════════════════════
# IF_history tab 管理
# ══════════════════════════════════════════

IF_HISTORY_HEADERS = ["so_norm", "pick_date", "if_doc", "if_status",
                      "customer_name", "shipping_zip", "updated_at"]


def _ensure_if_history_tab(sh):
    """确保 IF_history tab 存在，不存在则创建"""
    try:
        ws = sh.worksheet(CONFIG["if_history_tab"])
        return ws
    except gspread.exceptions.WorksheetNotFound:
        print(f"   IF_history tab 不存在，自动创建...")
        ws = sh.add_worksheet(title=CONFIG["if_history_tab"], rows=2000, cols=10)
        ws.append_row(IF_HISTORY_HEADERS)
        return ws


def load_if_history(sh):
    """从 IF_history tab 读取历史数据（兼容表头缺失/重复/空列）"""
    ws = _ensure_if_history_tab(sh)
    all_vals = ws.get_all_values()

    # 用固定表头，不依赖 Sheet 第一行内容（防止空列头报错）
    headers = IF_HISTORY_HEADERS

    # 没有数据行（空表或只有表头）→ 返回空 DataFrame，不清除任何内容
    data_rows = [r for r in all_vals[1:] if any(v.strip() for v in r)]
    if not data_rows:
        df = pd.DataFrame(columns=IF_HISTORY_HEADERS)
        df["pick_date_dt"] = None
        return df, ws

    records = []
    for row in data_rows:
        while len(row) < len(headers):
            row.append("")
        records.append({h: row[i] for i, h in enumerate(headers)})

    df = pd.DataFrame(records)
    df = df[df["so_norm"].str.strip() != ""].copy()
    df["pick_date_dt"] = df["pick_date"].apply(_parse_date)
    return df, ws


def upsert_if_history(ws, history_df, new_if_df, active_so_set):
    """
    将当日 IF 数据 upsert 进 IF_history：
    - 只处理追踪表里存在的 SO（active_so_set）
    - 唯一 key = so_norm + if_doc（支持同一 SO 多次抓货/分批 fulfill）
    - key 已存在 → 更新；不存在 → 追加
    """
    print("🔄 更新 IF_history...")
    now_str = datetime.today().strftime("%Y-%m-%d")

    # 只保留追踪表里有的 SO
    filtered_if = new_if_df[new_if_df["so_norm"].isin(active_so_set)].copy()
    print(f"   IF 表共 {len(new_if_df)} 条 → 追踪表匹配 {len(filtered_if)} 条，其余忽略")

    if len(filtered_if) == 0:
        print("   无需更新")
        return history_df, ws

    existing = history_df.copy()
    # 唯一 key = so_norm::if_doc
    existing["_ukey"] = existing["so_norm"].astype(str) + "::" + existing["if_doc"].astype(str)
    ukey_to_row = {row["_ukey"]: i + 2 for i, row in existing.iterrows()}

    updates = []
    new_rows = []

    for _, row in filtered_if.iterrows():
        so       = str(row["so_norm"])
        if_doc   = str(row["if_doc"])
        ukey     = f"{so}::{if_doc}"
        pick_str = row["pick_date"].strftime("%Y-%m-%d") if row["pick_date"] else ""
        new_vals = [
            so, pick_str, if_doc,
            str(row["if_status"]),
            str(row.get("Customer Name", "")),
            str(row.get("Shipping Zip", "")),
            now_str,
        ]
        if ukey in ukey_to_row:
            sheet_row = ukey_to_row[ukey]
            updates.append({"range": f"A{sheet_row}:G{sheet_row}", "values": [new_vals]})
        else:
            new_rows.append(new_vals)

    if updates:
        import time
        for i in range(0, len(updates), 50):
            ws.batch_update(updates[i:i+50])
            if i + 50 < len(updates):
                time.sleep(1.2)  # 避免触发每分钟写入限速
        print(f"   更新已有记录：{len(updates)} 条")
    if new_rows:
        ws.append_rows(new_rows)
        print(f"   新增记录：{len(new_rows)} 条")

    # 重新读取（用固定表头，避免 get_all_records 因空列头报错）
    all_vals = ws.get_all_values()
    headers = IF_HISTORY_HEADERS
    data_rows = all_vals[1:] if len(all_vals) > 1 else []
    records = []
    for row in data_rows:
        while len(row) < len(headers):
            row.append("")
        records.append({h: row[i] for i, h in enumerate(headers)})
    df = pd.DataFrame(records) if records else pd.DataFrame(columns=IF_HISTORY_HEADERS)
    df = df[df["so_norm"].str.strip() != ""].copy()
    df["pick_date_dt"] = df["pick_date"].apply(_parse_date)
    return df, ws


def cleanup_if_history(ws, history_df, active_so_set):
    """
    删除 IF_history 中已不在追踪表的行。
    IF_history 数量通常很小（只存追踪表匹配的 SO），直接删除安全且快速。
    从后往前删，避免行号偏移。
    """
    import time

    all_vals = ws.get_all_values()
    if len(all_vals) <= 1:
        print("   IF_history 无需清理")
        return

    headers = all_vals[0]
    so_col_idx = headers.index("so_norm") if "so_norm" in headers else 0

    rows_to_delete = []
    for i, row in enumerate(all_vals[1:], start=2):
        so = _normalize_so(row[so_col_idx]) if so_col_idx < len(row) else ""
        if so and so not in active_so_set:
            rows_to_delete.append(i)

    if not rows_to_delete:
        print("   IF_history 无需清理")
        return

    print(f"   清理 {len(rows_to_delete)} 条过期记录...")
    # 从后往前删，每10行暂停1秒避免触发 API 限速
    for idx, sheet_row in enumerate(reversed(rows_to_delete)):
        ws.delete_rows(sheet_row)
        if (idx + 1) % 10 == 0:
            time.sleep(1)
    print(f"   已清理 {len(rows_to_delete)} 条过期 IF 记录")


# ══════════════════════════════════════════
# 读取追踪表
# ══════════════════════════════════════════

def load_tracking_sheet(client):
    print("📥 读取 Google Sheet 追踪表...")
    sh = client.open_by_key(CONFIG["sheet_id"])
    ws = sh.worksheet(CONFIG["tracking_tab"])

    # 表头在第3行，数据从第4行开始
    all_values = ws.get_all_values()
    headers = all_values[2]   # 第3行（index 2）是列名

    # 去掉空列头，避免重复列名报错
    data_rows = all_values[3:]  # 第4行开始是数据
    records = []
    for row in data_rows:
        # 补齐行长度
        while len(row) < len(headers):
            row.append("")
        record = {}
        seen = {}
        for i, h in enumerate(headers):
            if h == "":
                continue  # 跳过空列头
            if h in seen:
                continue  # 跳过重复列头
            seen[h] = True
            record[h] = row[i] if i < len(row) else ""
        records.append(record)

    df = pd.DataFrame(records)
    # 缺列检查
    required_track_cols = [
        CONFIG["col_order_date"], CONFIG["col_acct"], CONFIG["col_zip"],
        CONFIG["col_so"], CONFIG["col_status"], CONFIG["col_note"],
        CONFIG["col_shipped_date"], CONFIG["col_qty"],
    ]
    missing = [c for c in required_track_cols if c not in df.columns]
    if missing:
        print(f"⚠️  追踪表缺少以下列（请检查 CONFIG 列名配置）：{missing}")
        print(f"   实际列名（前15个）：{list(df.columns[:15])}")
        # 缺列不直接退出，仅警告，让后续逻辑自行处理

    print(f"   读取到 {len(df)} 行，列：{list(df.columns[:8])}...")
    return df, ws, sh


# ══════════════════════════════════════════
# 核心分析
# ══════════════════════════════════════════

def analyze(tracking_df, history_df, config):
    today = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)
    cc = config
    window = cc["consolidate_window"]
    qty_threshold = cc["large_qty_threshold"]

    # ── 清洗追踪表 ──
    df = tracking_df.copy()
    df["_order_date"]  = df[cc["col_order_date"]].apply(_parse_date)
    df["_acct"]        = df[cc["col_acct"]].astype(str).str.strip()
    # ZIP 前导0保护：提取纯数字后补齐5位（防止 Google Sheet 把 00968 读成 968）
    df["_zip"] = (
        df[cc["col_zip"]].astype(str).str.strip()
        .str.extract(r"(\d+)")[0].fillna("")
        .apply(lambda z: z.zfill(5) if z else "")
    )
    df["_so_norm"]     = df[cc["col_so"]].apply(_normalize_so)
    df["_status"]      = df[cc["col_status"]].astype(str).str.strip().str.lower()
    df["_shipped_date"]= df[cc["col_shipped_date"]].apply(_parse_date)
    df["_qty"]         = pd.to_numeric(df[cc["col_qty"]], errors="coerce").fillna(0)

    col_send = cc.get("col_send_date", "")
    df["_send_date"] = df[col_send].apply(_parse_date) if col_send in df.columns else None

    # ── 筛选分析池（Status 白名单精确匹配）──
    # 白名单：精确匹配允许进入分析的状态（全部转小写比较）
    # stock order  → 未处理，可合并
    # fcsrepleshing → 补货中，可合并
    # hold          → 可能异常但仍可合并（等待中）
    # send          → 今天给了shipping，还来得及合并
    # shipped       → 已寄出，不参与

    STATUS_STOCK  = "stock order"
    STATUS_FCSREP = "fcsrepleshing"    # 确认拼写：Google Sheet 里就是 FcsRepleshing
    STATUS_HOLD   = "hold"
    STATUS_SEND   = "send"

    mask_stock   = df["_status"] == STATUS_STOCK
    mask_fcsrep  = df["_status"] == STATUS_FCSREP
    mask_hold    = df["_status"] == STATUS_HOLD
    mask_send_today = (
        (df["_status"] == STATUS_SEND) &
        df["_send_date"].apply(lambda d: d is not None and d.date() == today.date())
    )
    mask_candidate = mask_stock | mask_fcsrep | mask_hold | mask_send_today

    # SO# 为空的行（vlookup 留下的空行）直接排除
    mask_has_so = df["_so_norm"].str.strip() != ""
    mask_candidate = mask_candidate & mask_has_so

    # 排除不参与合单的账号
    exclude_accts = [str(a).strip() for a in cc.get("exclude_accts", [])]
    if exclude_accts:
        mask_exclude = df["_acct"].isin(exclude_accts)
        n_excluded = (mask_candidate & mask_exclude).sum()
        mask_candidate = mask_candidate & ~mask_exclude
    else:
        n_excluded = 0

    pending = df[mask_candidate].copy()

    # 用 SO# 去重计数（排除空值/重复），以 SO 为主数据
    so_counts = pending.groupby("_status")["_so_norm"].nunique()
    n_total   = pending["_so_norm"].nunique()
    print(f"\n📊 分析池：{n_total} 个 SO")
    print(f"   stock order：{so_counts.get(STATUS_STOCK, 0)}"
          f"  |  FcsRepleshing：{so_counts.get(STATUS_FCSREP, 0)}"
          f"  |  hold：{so_counts.get(STATUS_HOLD, 0)}"
          f"  |  今日SEND：{so_counts.get(STATUS_SEND, 0)}")
    if n_excluded:
        print(f"   已排除账号 {exclude_accts}：{n_excluded} 个 SO（不参与合单）")

    # ── 关联 IF 历史（取每个 SO 最新抓货日）──
    if len(history_df) > 0 and "so_norm" in history_df.columns:
        latest_if = (
            history_df.dropna(subset=["pick_date_dt"])
            .sort_values("pick_date_dt")
            .groupby("so_norm")["pick_date_dt"]
            .last()
            .reset_index()
            .rename(columns={"so_norm": "_so_norm", "pick_date_dt": "_if_date"})
        )
        latest_if["_so_norm"] = latest_if["_so_norm"].astype(str)
        pending["_so_norm"] = pending["_so_norm"].astype(str)
        pending = pending.merge(latest_if, on="_so_norm", how="left")
    else:
        pending["_if_date"] = None

    pending["_has_if"] = pending["_if_date"].notna()

    # ── 截止日 & 紧急度 ──
    def get_deadline(row):
        if row["_shipped_date"] is not None:
            return row["_shipped_date"]
        if row["_order_date"] is not None:
            return row["_order_date"] + timedelta(days=cc["expiry_days"])
        return None

    pending["_deadline"]    = pending.apply(get_deadline, axis=1)
    pending["_days_left"]   = pending["_deadline"].apply(
        lambda d: (d - today).days if d else None)
    pending["_expiry_str"]  = pending["_deadline"].apply(_fmt_md)

    def urgency(d):
        if d is None: return ""
        if d < 0:     return "⚠️ 已过期"
        if d <= cc["urgent_days"]: return "🔴 紧急"
        if d <= 4:    return "🟡 注意"
        return ""

    pending["_urgency"] = pending["_days_left"].apply(urgency)

    # ── 按 ACCT# + ZIP 分组，逐对判断合并 ──
    # 结果字典：so_norm → {"group_label", "merge_with": [...], "warning": ""}
    merge_result = {}   # so_norm → list of so_norm it can merge with
    warnings     = {}   # so_norm → warning message

    groups = pending.groupby(["_acct", "_zip"])

    group_label_counter = [0]
    group_label_map = {}  # frozenset(so_norms) → label

    def get_label(so_set):
        key = frozenset(so_set)
        if key not in group_label_map:
            i = group_label_counter[0]
            label = f"合单组{i + 1}"   # 数字编号，无上限
            group_label_map[key] = label
            group_label_counter[0] += 1
        return group_label_map[key]

    for (acct, zip_), grp in groups:
        if len(grp) < 2:
            continue

        sos = grp["_so_norm"].tolist()

        # 对组内每对 SO 判断是否可合并
        can_merge_pairs    = set()   # (so_a, so_b) 已判断可合并
        excluded_from_merge = set()  # 被排除的 SO（无IF大单），从所有边中移除

        for so_a, so_b in combinations(sos, 2):
            row_a = grp[grp["_so_norm"] == so_a].iloc[0]
            row_b = grp[grp["_so_norm"] == so_b].iloc[0]

            # ── 时间窗口检查 ──
            diff_ship = _date_diff_days(row_a["_deadline"], row_b["_deadline"])
            diff_if   = _date_diff_days(row_a["_if_date"],  row_b["_if_date"])

            time_ok = (
                (diff_ship is not None and diff_ship <= window) or
                (diff_if   is not None and diff_if   <= window)
            )
            if not time_ok:
                continue  # 时间窗口不符，跳过这对

            # ── 排除检查：一有IF一无IF且无IF那张QTY > threshold ──
            a_has_if = row_a["_has_if"]
            b_has_if = row_b["_has_if"]

            if a_has_if != b_has_if:
                no_if_row = row_b if a_has_if else row_a
                no_if_so  = so_b  if a_has_if else so_a
                has_if_so = so_a  if a_has_if else so_b

                if no_if_row["_qty"] > qty_threshold:
                    # 把该 SO 加入排除集合，后续所有含它的边都过滤掉
                    excluded_from_merge.add(no_if_so)
                    if no_if_so not in warnings:
                        warnings[no_if_so] = (
                            f"⚠️ QTY={int(no_if_row['_qty'])} 未抓货，"
                            f"与 {has_if_so} 放弃合并"
                        )
                    continue

            can_merge_pairs.add((min(so_a, so_b), max(so_a, so_b)))

        # 过滤掉含排除 SO 的所有边（选项三核心：BBB002 被排除后，BBB001+BBB003 仍可合并）
        can_merge_pairs = {
            (a, b) for a, b in can_merge_pairs
            if a not in excluded_from_merge and b not in excluded_from_merge
        }

        # ── 把可合并的对，整理成每张单视角的"可与谁合并"列表 ──
        if not can_merge_pairs:
            continue

        # ── Clique 逻辑：组内任意两张都必须满足时间窗口才算一组 ──
        # （替代 BFS：防止 A-B可合、B-C可合 但 A-C不可合 时被错误归为同组）
        all_sos_in_pairs = sorted({s for pair in can_merge_pairs for s in pair})
        remaining = set(all_sos_in_pairs)

        while remaining:
            # 从剩余中尝试找最大 clique（贪心：从连接数最多的节点开始扩展）
            best_clique = set()
            for start in sorted(remaining):
                clique = {start}
                for so in sorted(remaining - {start}):
                    # 检查 so 和 clique 里每一个成员都有合并边
                    if all((min(so, m), max(so, m)) in can_merge_pairs for m in clique):
                        clique.add(so)
                if len(clique) > len(best_clique):
                    best_clique = clique

            if len(best_clique) >= 2:
                label = get_label(best_clique)
                for so in best_clique:
                    others = sorted(best_clique - {so})
                    # 记录合并信息，包含 diff 信息供 Note 展示
                    merge_result[so] = {"label": label, "others": others}
                remaining -= best_clique
            else:
                break  # 剩余节点无法组成 clique，退出

    # ── 生成 Note 内容（含合单原因，方便操作员核实）──
    # 预先建立 so → row 的快速查找
    so_row_map = {r["_so_norm"]: r for _, r in pending.iterrows()}

    def build_note(row):
        so = row["_so_norm"]
        parts = []

        if so in merge_result:
            info   = merge_result[so]
            others = info["others"][:6]
            status_str = str(row[cc["col_status"]]).strip().lower()
            tag = "📦 今日SEND" if status_str == "send" else "✅"

            others_str = "、".join(f"SO{s}" for s in others)
            parts.append(f"可与 {others_str} 合发")

        if so in warnings:
            parts.append(warnings[so])

        if row["_has_if"]:
            parts.append(f"已抓货:{_fmt_md(row['_if_date'])}")

        return " | ".join(parts)

    pending["_new_note"] = pending.apply(build_note, axis=1)

    # ── 汇总打印 ──
    n_merge = len(merge_result)
    n_warn  = len(warnings)
    n_urgent = pending["_urgency"].str.contains("紧急|过期", na=False).sum()
    print(f"\n{'='*55}")
    print(f"  ✅ 可合单 SO：{n_merge} 张，共 {len(group_label_map)} 组")
    print(f"  ⚠️  放弃合并警告：{n_warn} 张")
    print(f"  🔴 紧急/过期：{n_urgent} 张")
    print(f"{'='*55}")

    if group_label_map:
        print("\n合单明细：")
        for so_set, label in group_label_map.items():
            sos_in_group = sorted(so_set)
            rows_in_group = pending[pending["_so_norm"].isin(sos_in_group)]
            urgencies = [u for u in rows_in_group["_urgency"].tolist() if u]
            urg = f"  {'  '.join(urgencies)}" if urgencies else ""
            print(f"  {label}{urg}")
            for so in sos_in_group:
                r = pending[pending["_so_norm"] == so]
                if len(r):
                    r = r.iloc[0]
                    has_if_tag = "📦有IF" if r["_has_if"] else "  无IF"
                    qty_str = f"QTY:{int(r['_qty'])}"
                    deadline = r["_expiry_str"]
                    print(f"    {has_if_tag}  {qty_str}  截止{deadline}  SO:{so}")

    return pending


# ══════════════════════════════════════════
# 写回 Google Sheet
# ══════════════════════════════════════════

def write_back(ws, pending_df, config):
    cc = config
    # 表头在第3行
    headers = ws.row_values(3)
    if cc["col_note"] not in headers or cc["col_so"] not in headers:
        missing = [c for c in [cc["col_note"], cc["col_so"]] if c not in headers]
        print(f"⚠️  Sheet 第3行找不到列：{missing}，跳过写回")
        print(f"   当前第3行列名：{headers[:10]}")
        return

    note_col = headers.index(cc["col_note"]) + 1
    so_col   = headers.index(cc["col_so"])   + 1

    # 脚本新增内容的标识前缀，用于判断是否已追加过
    AUTO_TAG = "[自动]"

    # 只收录这次有新建议的行；没建议的行完全不动原 Note
    note_map = {
        _normalize_so(row[cc["col_so"]]): row["_new_note"]
        for _, row in pending_df.iterrows()
        if row["_new_note"]
    }

    # 读取 SO 列和 Note 列（数据从第4行开始，跳过前3行）
    all_so   = ws.col_values(so_col)
    all_note = ws.col_values(note_col)

    updates = []
    # 数据从第4行开始（index 3），sheet row = index + 1
    for idx, so_val in enumerate(all_so[3:], start=3):
        row_idx = idx + 1  # sheet 行号（1-based）
        so_norm = _normalize_so(so_val)

        # 这次没有新建议 → 完全不动这行的 Note
        if so_norm not in note_map:
            continue
        new_content = note_map[so_norm]

        # 读取现有 Note
        existing = all_note[idx].strip() if idx < len(all_note) else ""

        # 保留人工填写部分，刷新 [自动] 之后内容
        if AUTO_TAG in existing:
            auto_start = existing.index(AUTO_TAG)
            human_part = existing[:auto_start].strip()
        else:
            human_part = existing

        # 拼接：人工内容（如有）+ 自动内容
        if human_part:
            merged = f"{human_part}  {AUTO_TAG}{new_content}"
        else:
            merged = f"{AUTO_TAG}{new_content}"

        # 内容没变就不写（减少 API 调用）
        if merged == existing:
            continue

        updates.append({
            "range": gspread.utils.rowcol_to_a1(row_idx, note_col),
            "values": [[merged]],
        })

    if updates:
        # 分批写入，每批100条，避免超过 Google API 单次限制
        batch_size = 100
        for i in range(0, len(updates), batch_size):
            ws.batch_update(updates[i:i + batch_size])
            print(f"   写入进度：{min(i + batch_size, len(updates))}/{len(updates)}")
        print(f"✅ 已更新 {len(updates)} 条备注到 Google Sheet（保留原有人工 Note）")
    else:
        print("⚠️  没有需要更新的行")


# ══════════════════════════════════════════
# 导出本地报告
# ══════════════════════════════════════════

def export_report(pending_df, config):
    cc = config
    cols = [cc["col_so"], cc["col_acct"], cc["col_zip"],
            cc["col_order_date"], "_expiry_str", "_days_left",
            "_urgency", "_has_if", "_qty", "_new_note"]
    cols = [c for c in cols if c in pending_df.columns]
    report = pending_df[pending_df["_new_note"].ne("")][cols].copy()
    report.columns = [c.lstrip("_") for c in report.columns]
    fname = f"consolidation_report_{datetime.today().strftime('%Y%m%d')}.csv"
    report.to_csv(fname, index=False, encoding="utf-8-sig")
    print(f"📄 本地报告：{fname}（{len(report)} 条）")


# ══════════════════════════════════════════
# 主程序
# ══════════════════════════════════════════

def main():
    print("=" * 55)
    print("  Order Consolidation Checker")
    print(f"  {datetime.today().strftime('%Y-%m-%d %H:%M')}")
    print("=" * 55)

    scopes = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
    ]
    try:
        creds = Credentials.from_service_account_file(
            CONFIG["credentials_file"], scopes=scopes)
        client = gspread.authorize(creds)
    except FileNotFoundError:
        print(f"❌ 找不到 {CONFIG['credentials_file']}，请先配置 Google Service Account")
        sys.exit(1)

    # 1. 读取追踪表
    tracking_df, tracking_ws, sh = load_tracking_sheet(client)

    # 2. 读取当日 IF 文件
    new_if_df = load_if_xls(CONFIG["if_xls_path"])

    # 3. 先算出追踪表里的有效 SO 集合（用于过滤 IF 数据）
    active_so_set = set(tracking_df[CONFIG["col_so"]].apply(_normalize_so).tolist())
    active_so_set.discard("")

    # 4. 加载 IF_history，只 upsert 追踪表里存在的 SO
    print("\n📂 管理 IF_history...")
    history_df, history_ws = load_if_history(sh)
    history_df, history_ws = upsert_if_history(history_ws, history_df, new_if_df, active_so_set)

    # 5. 清理 IF_history 中已不在追踪表的 SO
    cleanup_if_history(history_ws, history_df, active_so_set)

    # 6. 重新从 Sheet 读取（cleanup 后的最终状态）
    # 注意：直接用内存里的 history_df 即可，cleanup 只删了不在追踪表的行
    # history_df 本身已经是过滤过 active_so_set 的，无需再读一次
    # （避免二次读取触发 API 缓存问题）

    # 7. 分析合单
    pending_df = analyze(tracking_df, history_df, CONFIG)

    # 8. 写回 Google Sheet
    # 支持命令行参数：--write 直接写回，--dry-run 只分析不写回
    # 无参数则交互询问
    import sys as _sys
    args = _sys.argv[1:]
    if "--write" in args:
        print("\n[--write 模式] 直接写回 Google Sheet...")
        write_back(tracking_ws, pending_df, CONFIG)
    elif "--dry-run" in args:
        print("\n[--dry-run 模式] 仅分析，不写回 Google Sheet")
    else:
        ans = input("\n是否将结果写回 Google Sheet Note 列？(y/n): ").strip().lower()
        if ans == "y":
            write_back(tracking_ws, pending_df, CONFIG)

    # 9. 导出本地报告
    export_report(pending_df, CONFIG)
    print("\n🎉 完成！")


if __name__ == "__main__":
    main()