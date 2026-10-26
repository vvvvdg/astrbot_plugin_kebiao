"""把课表画成一张图（Pillow），不依赖系统字体：容器里用 Noto Sans CJK。"""

from __future__ import annotations

import hashlib
import os
from typing import Any

from PIL import Image, ImageDraw, ImageFont

FONT_CANDIDATES = [
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/opentype/noto/NotoSerifCJK-Regular.ttc",
    "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
    "C:/Windows/Fonts/msyh.ttc",
    "C:/Windows/Fonts/simhei.ttf",
]
BOLD_CANDIDATES = [
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc",
    "/usr/share/fonts/opentype/noto/NotoSerifCJK-Bold.ttc",
    "C:/Windows/Fonts/msyhbd.ttc",
] + FONT_CANDIDATES

WEEKDAYS = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
WEEKDAY_KEYS = [f"星期{i}" for i in "一二三四五六日"]
BLOCKS = [1, 3, 5, 7, 9]  # 大节起点
BLOCK_LABEL = {1: "1-2节", 3: "3-4节", 5: "5-6节", 7: "7-8节", 9: "9-10节"}
BLOCK_TIME = {1: "08:00", 3: "10:00", 5: "14:00", 7: "16:00", 9: "19:00"}

PALETTE = [
    (232, 241, 255),
    (255, 240, 232),
    (233, 248, 236),
    (255, 246, 224),
    (243, 236, 253),
    (230, 247, 249),
    (252, 236, 243),
    (238, 243, 228),
]


def _font(size: int, bold: bool = False):
    for path in (BOLD_CANDIDATES if bold else FONT_CANDIDATES):
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size)
            except Exception:  # noqa: BLE001
                continue
    return ImageFont.load_default()


def _color(name: str) -> tuple[int, int, int]:
    digest = hashlib.md5(name.encode("utf-8")).hexdigest()
    return PALETTE[int(digest[:4], 16) % len(PALETTE)]


def _fmt_weeks(weeks) -> str:
    weeks = sorted({int(w) for w in weeks or []})
    if not weeks:
        return ""
    out, start, prev = [], weeks[0], weeks[0]
    for w in weeks[1:]:
        if w == prev + 1:
            prev = w
            continue
        out.append(f"{start}-{prev}" if start != prev else f"{start}")
        start = prev = w
    out.append(f"{start}-{prev}" if start != prev else f"{start}")
    return "、".join(out) + "周"


def _merge(courses: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """同一时段、同一门课的多条记录（不同周次/教室）合并成一张卡片"""
    merged: dict[tuple, dict[str, Any]] = {}
    order: list[tuple] = []
    for c in courses:
        k = (c.get("day"), c.get("period"), c.get("name"))
        if k not in merged:
            item = dict(c)
            item["_locs"] = [c.get("location")] if c.get("location") else []
            merged[k] = item
            order.append(k)
            continue
        m = merged[k]
        m["weeks"] = sorted(set(m.get("weeks") or []) | set(c.get("weeks") or []))
        loc = c.get("location")
        if loc and loc not in m["_locs"]:
            m["_locs"].append(loc)
    result = []
    for k in order:
        m = merged[k]
        locs = [x for x in m.get("_locs", []) if x]
        m["location"] = " / ".join(locs[:2]) + ("…" if len(locs) > 2 else "")
        result.append(m)
    return result


def _block(n: int) -> int:
    if n <= 2:
        return 1
    if n <= 4:
        return 3
    if n <= 6:
        return 5
    if n <= 8:
        return 7
    return 9


def _span(period: str) -> tuple[int, int]:
    import re

    nums = [int(x) for x in re.findall(r"\d+", str(period))]
    if not nums:
        return 1, 1
    return _block(nums[0]), _block(nums[-1])


def _wrap(draw, text: str, font, max_w: int) -> list[str]:
    lines, cur = [], ""
    for ch in text:
        if draw.textlength(cur + ch, font=font) <= max_w:
            cur += ch
        else:
            lines.append(cur)
            cur = ch
    if cur:
        lines.append(cur)
    return lines or [""]


def _vgrad(width: int, height: int, top=(253, 254, 255), bottom=(232, 239, 253)):
    img = Image.new("RGB", (width, height), top)
    d = ImageDraw.Draw(img)
    for y in range(height):
        t = y / max(1, height - 1)
        c = tuple(int(top[i] + (bottom[i] - top[i]) * t) for i in range(3))
        d.line([(0, y), (width, y)], fill=c)
    return img


def _crown(draw, cx: int, cy: int, color=(246, 190, 60)) -> None:
    pts = [
        (cx - 15, cy + 6),
        (cx - 15, cy - 7),
        (cx - 7, cy + 1),
        (cx, cy - 11),
        (cx + 7, cy + 1),
        (cx + 15, cy - 7),
        (cx + 15, cy + 6),
    ]
    draw.polygon(pts, fill=color)


def render_leaderboard(
    out_path: str,
    title: str,
    rows: list[tuple[str, str, float]],
    subtitle: str = "",
    footer: str = "",
    empty: str = "今天没人有课喵～",
) -> str:
    """上课排行榜：rows = [(昵称, 时长文本, 0~1 比例)]，前三名有奖牌样式。"""
    width = 900
    pad = 52
    medal = [
        (246, 190, 60),   # 金
        (188, 199, 214),  # 银
        (211, 148, 96),   # 铜
    ]
    tint = [
        (255, 249, 232),
        (244, 247, 252),
        (253, 243, 235),
    ]
    top_n = min(3, len(rows))
    row_h_top = 92
    row_h = 76
    head_h = 196
    body_h = sum(row_h_top for _ in range(top_n)) + sum(row_h for _ in rows[top_n:])
    body_h = max(body_h, 120)
    height = head_h + body_h + 46 + (34 if footer else 18)

    img = _vgrad(width, height)
    draw = ImageDraw.Draw(img)

    f_title = _font(46, True)
    f_sub = _font(23)
    f_hat = _font(21, True)
    f_name = _font(28, True)
    f_name_n = _font(25)
    f_val = _font(26, True)
    f_val_n = _font(23, True)
    f_num = _font(24, True)

    # 顶部：标题 + 副标题
    draw.text((pad, 40), title, font=f_title, fill=(27, 31, 42))
    if subtitle:
        draw.text((pad, 104), subtitle, font=f_sub, fill=(118, 125, 140))

    # 小标签
    chip = "上课时长榜"
    cw = draw.textlength(chip, font=f_hat) + 30
    draw.rounded_rectangle([pad, 146, pad + cw, 146 + 34], radius=17, fill=(233, 240, 255))
    draw.text((pad + 15, 153), chip, font=f_hat, fill=(64, 106, 214))

    y = head_h
    if not rows:
        draw.rounded_rectangle(
            [pad - 10, y, width - pad + 10, y + 100], radius=20, fill=(244, 247, 252)
        )
        draw.text((pad + 16, y + 32), empty, font=f_name_n, fill=(150, 157, 172))
    # 并列名次：时长相同的人拿同一个名次（1,1,3,3,5）
    ratios = [r[2] for r in rows]
    ranks = [1 + sum(1 for v in ratios if v > mine + 1e-9) for mine in ratios]

    for idx, (name, value, ratio) in enumerate(rows):
        rk = ranks[idx]
        is_top = rk <= 3
        rh = row_h_top if is_top else row_h
        if is_top:
            draw.rounded_rectangle(
                [pad - 14, y + 4, width - pad + 14, y + rh - 6],
                radius=20,
                fill=tint[rk - 1],
            )
        cy = y + rh // 2
        color = medal[rk - 1] if is_top else (176, 188, 206)

        # 名次圆牌（前三名有奖牌色 + 第一名皇冠）
        r = 26 if is_top else 21
        cx = pad + 20 + r
        draw.ellipse([cx - r, cy - r, cx + r, cy + r], fill=color)
        num = str(rk)
        tw = draw.textlength(num, font=f_num)
        draw.text((cx - tw / 2, cy - 15), num, font=f_num, fill=(255, 255, 255))
        if rk == 1:
            _crown(draw, cx, cy - r - 13)

        nx = cx + r + 20
        label = _wrap(draw, str(name), f_name if is_top else f_name_n, 300)[0]
        draw.text((nx, cy - 24), label, font=f_name if is_top else f_name_n, fill=(32, 38, 50))

        # 条形
        bar_x0 = nx
        bar_x1 = width - pad - 205
        draw.rounded_rectangle(
            [bar_x0, cy + 12, bar_x1, cy + 24], radius=6, fill=(236, 240, 247)
        )
        rr = max(0.0, min(1.0, float(ratio or 0)))
        if rr > 0:
            draw.rounded_rectangle(
                [bar_x0, cy + 12, bar_x0 + max(20, (bar_x1 - bar_x0) * rr), cy + 24],
                radius=6,
                fill=color if is_top else (150, 172, 205),
            )

        # 数值
        vt = str(value)
        fv = f_val if is_top else f_val_n
        vw = draw.textlength(vt, font=fv)
        draw.text((width - pad - vw, cy - 18), vt, font=fv, fill=(52, 60, 78))
        y += rh

    if footer:
        draw.text((pad, height - 38), footer, font=f_sub, fill=(160, 167, 180))

    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    img.save(out_path, "PNG")
    return out_path


def render_grid(
    courses: list[dict[str, Any]],
    out_path: str,
    title: str,
    subtitle: str = "",
    show_weeks: bool = False,
    footer: str = "",
    highlight: str | None = None,
    day_dates: dict[str, str] | None = None,
) -> str:
    """courses: [{day, period, name, location, weeks?, label?}]

    highlight:  要突出显示的那一天（"星期二" 这种 key），表头蓝底 + 两边亮边
    day_dates:  {"星期一": "10-27", ...}，给了就在表头每天下面写日期
    """
    col_w, row_h = 196, 150
    left_w, top_h = 122, 128
    head_h = 54 if day_dates else 44
    grid_top = top_h + head_h
    width = left_w + col_w * 7
    height = grid_top + row_h * len(BLOCKS) + 40
    img = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(img)

    f_title = _font(40, True)
    f_sub = _font(24)
    f_head = _font(26, True)
    f_label = _font(22, True)
    f_tiny = _font(15)

    # 标题
    draw.text((left_w // 2, 24), title, font=f_title, fill=(28, 32, 40))
    if subtitle:
        draw.text((left_w // 2, 78), subtitle, font=f_sub, fill=(110, 116, 128))

    # 表头
    f_head = _font(26, True)
    f_date = _font(14)
    for i, (label, day_key) in enumerate(zip(WEEKDAYS, WEEKDAY_KEYS)):
        x = left_w + i * col_w
        hot = day_key == highlight
        draw.rectangle([x, top_h, x + col_w, grid_top], fill=(64, 106, 214) if hot else (238, 242, 248))
        ink = (255, 255, 255) if hot else (40, 46, 58)
        date_txt = (day_dates or {}).get(day_key, "")
        tw = draw.textlength(label, font=f_head)
        if date_txt:
            draw.text((x + (col_w - tw) / 2, top_h + 4), label, font=f_head, fill=ink)
            dw = draw.textlength(date_txt, font=f_date)
            draw.text((x + (col_w - dw) / 2, top_h + 34), date_txt, font=f_date, fill=ink)
        else:
            draw.text((x + (col_w - tw) / 2, top_h + 9), label, font=f_head, fill=ink)

    # 行底色 + 左侧节次
    for r, base in enumerate(BLOCKS):
        y = grid_top + r * row_h
        if r % 2 == 1:
            draw.rectangle([0, y, width, y + row_h], fill=(250, 251, 253))
        draw.text((14, y + 32), BLOCK_LABEL[base], font=f_label, fill=(60, 66, 78))
        draw.text((14, y + 68), BLOCK_TIME[base], font=f_tiny, fill=(140, 146, 158))

    # 网格线
    for r in range(len(BLOCKS) + 1):
        y = grid_top + r * row_h
        draw.line([(0, y), (width, y)], fill=(222, 227, 235), width=1)
    for c in range(8):
        x = left_w + c * col_w
        draw.line([(x, top_h), (x, grid_top + len(BLOCKS) * row_h)], fill=(222, 227, 235), width=1)
    draw.line([(0, top_h), (width, top_h)], fill=(222, 227, 235), width=1)

    # 高亮那一列的两条竖边
    if highlight and highlight in WEEKDAY_KEYS:
        hi = WEEKDAY_KEYS.index(highlight)
        hx = left_w + hi * col_w
        bottom = grid_top + len(BLOCKS) * row_h
        draw.rectangle([hx, top_h, hx + 3, bottom], fill=(64, 106, 214))
        draw.rectangle([hx + col_w - 3, top_h, hx + col_w, bottom], fill=(64, 106, 214))

    # 同一格里的多门课 → 分道并排
    groups: dict[tuple[int, int, int], list[dict[str, Any]]] = {}
    for c in _merge(courses):
        if c.get("day") not in WEEKDAY_KEYS:
            continue
        day_idx = WEEKDAY_KEYS.index(c["day"])
        b0, b1 = _span(c.get("period", ""))
        if b0 not in BLOCKS or b1 not in BLOCKS:
            continue
        r0, r1 = BLOCKS.index(b0), BLOCKS.index(b1)
        groups.setdefault((day_idx, r0, r1), []).append(c)

    for (day_idx, r0, r1), items in groups.items():
        n = len(items)
        base_x = left_w + day_idx * col_w
        lane_w = (col_w - 8) / n
        y0 = grid_top + r0 * row_h + 4
        y1 = grid_top + (r1 + 1) * row_h - 4
        f_name = _font(17) if n == 1 else _font(16)
        for i, c in enumerate(items):
            x0 = base_x + 4 + i * lane_w
            x1 = x0 + lane_w - 3
            draw.rounded_rectangle([x0, y0, x1, y1], radius=8, fill=_color(c["name"]))
            pad = 6
            inner = int(x1 - x0 - pad * 2)
            ty = y0 + pad
            for line in _wrap(draw, c.get("name", ""), f_name, inner)[:5]:
                draw.text((x0 + pad, ty), line, font=f_name, fill=(30, 36, 46))
                ty += 20
            for line in _wrap(draw, c.get("location", ""), f_tiny, inner)[:3]:
                draw.text((x0 + pad, ty + 1), line, font=f_tiny, fill=(88, 96, 110))
                ty += 18
            if show_weeks:
                wk = _fmt_weeks(c.get("weeks"))
                if wk:
                    wk_lines = _wrap(draw, wk, f_tiny, inner)[:2]
                    wy = y1 - 8 - 17 * len(wk_lines)
                    for line in wk_lines:
                        draw.text((x0 + pad, wy), line, font=f_tiny, fill=(112, 120, 136))
                        wy += 17

    if footer:
        draw.text((left_w, height - 28), footer, font=f_tiny, fill=(150, 156, 168))

    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    img.save(out_path, "PNG")
    return out_path
