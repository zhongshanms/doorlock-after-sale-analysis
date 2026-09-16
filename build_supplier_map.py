# -*- coding: utf-8 -*-
"""
生成供应商-SKU 映射数据文件(data/supplier-map.json),供网页端"筛选供应商"功能使用。

数据源: C:\\Users\\DELL\\WorkBuddy\\质量管理数据库\\08-采购数据\\SKU当前供应商.xlsx
        (DLP 加密,自动调用 decrypt_xlsx.ps1 解密;表头行自动定位,数据取 SKU 以 MS 开头且供应商非空的行)

输出: data/supplier-map.json
      {"v": 生成日期, "src": 源文件名, "sup": [供应商名...], "m": {"SKU": 供应商索引}}

用法:
    python build_supplier_map.py                      # 默认路径
    python build_supplier_map.py <xlsx路径>            # 指定源文件
"""
import json
import os
import re
import subprocess
import sys
from datetime import date

PREFIX_LEN = 6   # SKU 前缀长度(前6位, 如 MS2093), 用于同系列供应商继承

BASE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE, "data")
OUT_JSON = os.path.join(DATA_DIR, "supplier-map.json")
DECRYPT_PS1 = os.path.join(BASE, "decrypt_xlsx.ps1")
DEFAULT_SRC = r"C:\Users\DELL\WorkBuddy\质量管理数据库\08-采购数据\SKU当前供应商.xlsx"
COMPACT_JSON = os.path.join(DATA_DIR, "after-sale-data-compact.json")
EXCLUSIONS_FILE = os.path.join(BASE, "exclusions.json")

# ── 停用排除清单(不合作供应商 / 非目标产品) ──
_excl = {"sku_prefixes": ["MSCG"], "skus": [], "suppliers": []}
try:
    with open(EXCLUSIONS_FILE, encoding="utf-8") as _f:
        _raw = json.load(_f)
    for _k in _excl:
        if _raw.get(_k):
            _excl[_k] = _raw[_k]
except FileNotFoundError:
    pass
except Exception as _e:
    print(f"[!] exclusions.json 读取失败,使用默认排除({_excl['sku_prefixes']}): {_e}")

_EXCL_PREFIXES = [str(p).strip().upper() for p in _excl["sku_prefixes"]]
_EXCL_SKUS = {str(s).strip().upper() for s in _excl["skus"]}
EXCLUDED_SUPPLIERS = [str(s).strip() for s in _excl["suppliers"]]


def is_excluded_sku(sku):
    up = str(sku or "").strip().upper()
    if not up:
        return True
    if up in _EXCL_SKUS:
        return True
    return any(up.startswith(p) for p in _EXCL_PREFIXES)


def is_excluded_supplier(name):
    n = str(name or "").strip()
    if not n:
        return False
    return any(n == s or (len(s) >= 4 and s in n) for s in EXCLUDED_SUPPLIERS)


def _decrypt(path):
    """DLP 加密文件 -> 明文临时文件路径;非加密直接返回"""
    with open(path, "rb") as f:
        head = f.read(4)
    if not (len(head) >= 3 and head[0] == 0x88 and head[1] == 0x7D and head[2] == 0x1C):
        return path
    print("  检测到 DLP 加密,调用 decrypt_xlsx.ps1 解密...")
    r = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", DECRYPT_PS1, "-SourcePath", path],
        capture_output=True, text=True, encoding="utf-8", errors="ignore",
    )
    lines = [l.strip() for l in (r.stdout or "").splitlines() if l.strip()]
    if r.returncode != 0 or not lines:
        raise RuntimeError(f"解密失败: {r.stdout} {r.stderr}")
    decrypted = lines[-1]
    if not os.path.exists(decrypted):
        raise RuntimeError(f"解密输出路径不存在: {decrypted}")
    return decrypted


def build(src_path=DEFAULT_SRC):
    import openpyxl

    if not os.path.exists(src_path):
        print(f"[X] 源文件不存在: {src_path}")
        return 1
    print(f"读取: {src_path}")
    path = _decrypt(src_path)

    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    # 定位表头行(含 "SKU" 与 "当前供应商")
    hdr_idx = None
    for i, row in enumerate(rows[:10]):
        cells = [str(c).strip() if c is not None else "" for c in row]
        if "SKU" in cells and any("当前供应商" in c for c in cells):
            hdr_idx = i
            break
    if hdr_idx is None:
        print("[X] 未找到表头行(需含 'SKU' 与 '当前供应商')")
        return 1

    headers = [str(c).strip() if c is not None else "" for c in rows[hdr_idx]]
    sku_col = headers.index("SKU")
    sup_col = next(i for i, h in enumerate(headers) if "当前供应商" in h)
    time_col = next((i for i, h in enumerate(headers) if "下单时间" in h), None)
    print(f"  表头行: {hdr_idx + 1} | SKU列={sku_col + 1} 供应商列={sup_col + 1} "
          f"下单时间列={(time_col + 1) if time_col is not None else '无'}")

    suppliers = []
    sup_index = {}

    def sup_idx(name):
        if name not in sup_index:
            sup_index[name] = len(suppliers)
            suppliers.append(name)
        return sup_index[name]

    # 收集记录, 同时建前缀索引: 前缀 -> {供应商: [最近下单时间, SKU数]}
    recs = []
    pfx_index = {}
    skipped = {"non_ms": 0, "excluded": 0, "excluded_sup": 0, "empty_sup": 0}
    for row in rows[hdr_idx + 1:]:
        if not row:
            continue
        sku = str(row[sku_col]).strip() if row[sku_col] is not None else ""
        sup = str(row[sup_col]).strip() if len(row) > sup_col and row[sup_col] is not None else ""
        t = (str(row[time_col]).strip()
             if time_col is not None and len(row) > time_col and row[time_col] is not None else "")
        if not sku:
            continue
        up = sku.upper()
        if not up.startswith("MS"):
            skipped["non_ms"] += 1
            continue
        if is_excluded_sku(sku):          # 停用清单: MSCG 非门锁 / MS3011 黑迪已不合作
            skipped["excluded"] += 1
            continue
        if not sup:
            skipped["empty_sup"] += 1
            continue
        if is_excluded_supplier(sup):     # 不合作供应商
            skipped["excluded_sup"] += 1
            continue
        recs.append((sku, sup, t))
        e = pfx_index.setdefault(up[:PREFIX_LEN], {}).setdefault(sup, ["", 0])
        if t > e[0]:
            e[0] = t
        e[1] += 1

    if not recs:
        print("[X] 未提取到任何 MS SKU 供应商映射,未覆盖输出文件")
        return 1

    # 1) 直接映射(该 SKU 有采购记录)
    sku_map = {}
    for sku, sup, _t in recs:
        sku_map[sku] = sup_idx(sup)
    direct_cnt = len(sku_map)

    # 2) 前缀继承: 业务数据中出现、但采购表无记录的 SKU,
    #    取前6位前缀相同的已归类 SKU; 多候选按【最近下单时间】优先,
    #    时间并列时取该前缀下 SKU 数更多者(产品换供应商时以新供应商为准)
    inherited = {}
    biz_base = set()
    if os.path.exists(COMPACT_JSON):
        with open(COMPACT_JSON, encoding="utf-8") as f:
            compact = json.load(f)
        for key in ("ar", "sr"):
            for r in compact.get(key, []):
                s = str(r.get("sku") or "").strip()
                up = s.upper()
                if up.startswith("MS") and not up.startswith("MSCG"):
                    biz_base.add(re.sub(r"-\d+$", "", s))

    for b in sorted(biz_base):
        if b in sku_map:
            continue
        cand = pfx_index.get(b[:PREFIX_LEN].upper())
        if not cand:
            continue
        best_sup, best_meta = None, None
        for sup, meta in cand.items():
            if best_meta is None or (meta[0], meta[1]) > (best_meta[0], best_meta[1]):
                best_sup, best_meta = sup, meta
        inherited[b] = (best_sup, best_meta[0], len(cand))
        sku_map[b] = sup_idx(best_sup)

    out = {
        "v": date.today().isoformat(),
        "src": os.path.basename(src_path),
        "sup": suppliers,
        "m": sku_map,
        "inh": len(inherited),
    }
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, separators=(",", ":"))

    print(f"  MS SKU 直接映射 {direct_cnt} 条; 前缀继承补充 {len(inherited)} 条 → 共 {len(sku_map)} 条, {len(suppliers)} 家供应商")
    print(f"  跳过: 非MS {skipped['non_ms']} 条, 停用清单(前缀/SKU) {skipped['excluded']} 条, "
          f"不合作供应商 {skipped['excluded_sup']} 条, 供应商为空 {skipped['empty_sup']} 条")
    for s, i in sorted(sup_index.items(), key=lambda kv: -sum(1 for v in sku_map.values() if v == kv[1])):
        cnt = sum(1 for v in sku_map.values() if v == i)
        print(f"    {cnt:>5}  {s}")
    print(f"  输出: {OUT_JSON} ({os.path.getsize(OUT_JSON) / 1024:.1f} KB)")

    if inherited:
        print(f"  前缀继承明细(按前{PREFIX_LEN}位同系列 + 最近下单时间):")
        for b in sorted(inherited):
            sup, t, ncand = inherited[b]
            tail = f", 并列 {ncand} 家候选" if ncand > 1 else ""
            print(f"    {b:<16} → {sup}  (最近下单 {t or '未知'}{tail})")

    # 覆盖率: 与业务数据的基础 SKU 比对(去掉 -N 包装后缀)
    if biz_base:
        matched = {s for s in biz_base if s in sku_map}
        miss = sorted(biz_base - matched)
        print(f"  覆盖率: 业务基础SKU {len(biz_base)} 个, 已归属 {len(matched)} 个 "
              f"({len(matched) / len(biz_base) * 100:.1f}%)")
        if miss:
            fam = {}
            for s in miss:
                fam.setdefault(s[:PREFIX_LEN], []).append(s)
            print(f"  仍未归属 {len(miss)} 个 / {len(fam)} 个系列(采购表中这些系列无任何记录): "
                  + ", ".join(f"{k}×{len(v)}" for k, v in sorted(fam.items())))

    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    src = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_SRC
    sys.exit(build(src))
