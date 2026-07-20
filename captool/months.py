"""月份標頭解析:容錯處理現行 Excel 的各種格式與 typo。"""
import re
from datetime import datetime

_PATTERNS = [
    re.compile(r"^(?P<y>\d{4})['./\-](?P<m>\d{1,2})$"),      # 2026'8, 2026-08, 2026/8
    re.compile(r"^(?P<y>\d{4})(?P<m>0[1-9]|1[0-2])$"),        # 202608
]


def parse_month(raw):
    """回傳 (正規化 'YYYY-MM' 或 None, 警告訊息或 None)。"""
    if raw is None:
        return None, "月份標頭為空"
    if isinstance(raw, datetime):
        return f"{raw.year:04d}-{raw.month:02d}", None
    s = str(raw).strip()
    for pat in _PATTERNS:
        m = pat.match(s)
        if m:
            month = int(m.group("m"))
            if 1 <= month <= 12:
                return f"{int(m.group('y')):04d}-{month:02d}", None
    # 容錯:如 "20267'" 之類的 typo(去除非數字後為 4 位年 + 1~2 位月)
    digits = re.sub(r"\D", "", s)
    if len(digits) in (5, 6):
        y, mo = int(digits[:4]), int(digits[4:])
        if 2000 <= y <= 2100 and 1 <= mo <= 12:
            return f"{y:04d}-{mo:02d}", f"月份標頭 '{s}' 格式異常,已解讀為 {y:04d}-{mo:02d},請確認"
    return None, f"無法解析月份標頭 '{s}'"
