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
import subprocess
import sys
from datetime import date

BASE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE, "data")
OUT_JSON = os.path.join(DATA_DIR, "supplier-map.json")
DECRYPT_PS1 = os.path.join(BASE, "decrypt_xlsx.ps1")
DEFAULT_SRC = r"C:\Users\DELL\WorkBuddy\质量管理数据库\08-采购数据\SKU当前供应商.xlsx"
COMPACT_JSON = os.path.join(DATA_DIR, "after-sale-data-compact.json")


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
    print(f"  表头行: {hdr_idx + 1} | SKU列={sku_col + 1} 供应商列={sup_col + 1}")

    suppliers = []
    sup_index = {}
    sku_map = {}
    skipped = {"non_ms": 0, "empty_sup": 0}
    for row in rows[hdr_idx + 1:]:
        if not row:
            continue
        sku = str(row[sku_col]).strip() if row[sku_col] is not None else ""
        sup = str(row[sup_col]).strip() if len(row) > sup_col and row[sup_col] is not None else ""
        if not sku:
            continue
        if not sku.upper().startswith("MS"):
            skipped["non_ms"] += 1
            continue
        if not sup:
            skipped["empty_sup"] += 1
            continue
        if sup not in sup_index:
            sup_index[sup] = len(suppliers)
            suppliers.append(sup)
        sku_map[sku] = sup_index[sup]

    if not sku_map:
        print("[X] 未提取到任何 MS SKU 供应商映射,未覆盖输出文件")
        return 1

    out = {
        "v": date.today().isoformat(),
        "src": os.path.basename(src_path),
        "sup": suppliers,
        "m": sku_map,
    }
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, separators=(",", ":"))

    print(f"  MS SKU {len(sku_map)} 条 -> {len(suppliers)} 家供应商")
    print(f"  跳过: 非MS {skipped['non_ms']} 条, 供应商为空 {skipped['empty_sup']} 条")
    for s, i in sorted(sup_index.items(), key=lambda kv: -sum(1 for v in sku_map.values() if v == kv[1])):
        cnt = sum(1 for v in sku_map.values() if v == i)
        print(f"    {cnt:>5}  {s}")
    print(f"  输出: {OUT_JSON} ({os.path.getsize(OUT_JSON) / 1024:.1f} KB)")

    # 覆盖率: 与售后/销量数据的 SKU 比对
    if os.path.exists(COMPACT_JSON):
        with open(COMPACT_JSON, encoding="utf-8") as f:
            d = json.load(f)
        data_skus = set()
        for r in d.get("ar", []):
            if r.get("sku"):
                data_skus.add(str(r["sku"]).strip())
        for r in d.get("sr", []):
            if r.get("sku"):
                data_skus.add(str(r["sku"]).strip())
        ms_skus = {s for s in data_skus if s.upper().startswith("MS")}
        matched = {s for s in ms_skus if s in sku_map}
        miss = sorted(ms_skus - matched)
        print(f"  覆盖率: 业务数据 MS SKU {len(ms_skus)} 个, 已匹配供应商 {len(matched)} 个 "
              f"({len(matched) / len(ms_skus) * 100:.1f}%)")
        if miss:
            print(f"  未匹配 {len(miss)} 个(筛选时归入'未归类'): {', '.join(miss[:12])}{' ...' if len(miss) > 12 else ''}")

    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    src = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_SRC
    sys.exit(build(src))
