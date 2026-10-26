"""课表导入插件（MVP）

功能：
1. 把教务系统导出的 `学生个人课表_*.xls/.xlsx` 发给机器人 → 解析并按星期排好回复；
2. `/课表` 查看已保存的课表；`/课表 8` 只看第 8 周；`/课表 帮助` 看用法。

不做：定时提醒、网页上传、自动登录。
"""

from __future__ import annotations

import asyncio
import json
import os
import re
import time
from datetime import date, datetime, timedelta
from typing import Any

from astrbot.api import logger
from astrbot.api.event import AstrMessageEvent, filter
from astrbot.api.star import Context, Star, register
from astrbot.core.message.components import At, File, Image, Plain
from astrbot.core.message.message_event_result import MessageChain

try:  # 插件目录里的 image.py
    from .image import render_grid
except ImportError:  # pragma: no cover
    import sys

    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from image import render_grid  # type: ignore # noqa: E402

WEEKDAYS = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
WEEKDAY_KEYS = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]
PERIOD_LABEL = {
    "第一二节": "1-2",
    "第三四节": "3-4",
    "第五六节": "5-6",
    "第七八节": "7-8",
    "第九十节": "9-10",
    "第十一十二节": "11-12",
    "第十一十二节": "11-12",
}

# 一条课程：名字 / 老师 / 3-12([周])[01-02节] / 地点
ENTRY_RE = re.compile(
    r"(?P<name>[^\n]+)\n"
    r"(?P<teacher>[^\n]+)\n"
    r"(?P<weeks>[0-9,\-]+)\(\[周\]\)\[(?P<period>[^\]]*)\]"
    r"(?:\n(?P<loc>[^\n]*))?"
)

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
# 「开课就播报谁在上课」的开关状态（不存在视为开）
WHO_STATE = os.path.join(DATA_DIR, "_who_push.json")
# 上课前多少分钟提醒
REMIND_MIN = 20
# 插件自己的管理员名单（AstrBot 主配置里 admins_id 的人自动算管理员）
ADMIN_FILE = os.path.join(DATA_DIR, "_admins.json")
# 每天的行为计数（查课表 / 抓群友），排行榜用
STATS_FILE = os.path.join(DATA_DIR, "_stats.json")
# 排行榜上次推给哪个日期，防重复发
LB_STATE = os.path.join(DATA_DIR, "_rank_last.json")
# 数据目录里这些是插件自己的状态文件，不当成人来统计
INTERNAL_FILES = {"cmd_config.json", "_stats.json", "_who_push.json", "_admins.json", "_rank_last.json"}
# 第 1 周的周一（可在群里用 /课表 开学 2026-09-07 改）
DEFAULT_TERM_START = "2026-09-07"


def _load_admins() -> list[str]:
    """插件自己加的管理员（AstrBot 主配置 admins_id 之外的那些）"""
    try:
        with open(ADMIN_FILE, encoding="utf-8-sig") as f:
            d = json.load(f)
        return [str(x) for x in (d.get("admins") or [])]
    except Exception:  # noqa: BLE001
        return []


def _save_admins(ids: list[str]) -> None:
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(ADMIN_FILE, "w", encoding="utf-8") as f:
        json.dump({"admins": sorted({str(x) for x in ids})}, f, ensure_ascii=False, indent=2)


def _parse_weeks(text: str) -> list[int]:
    weeks: set[int] = set()
    for part in str(text).split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            try:
                a, b = part.split("-", 1)
                weeks.update(range(int(a), int(b) + 1))
            except ValueError:
                continue
        else:
            try:
                weeks.add(int(part))
            except ValueError:
                continue
    return sorted(weeks)


def _grid_from_xls(path: str) -> list[list[str]]:
    import xlrd

    book = xlrd.open_workbook(path)
    sheet = book.sheet_by_index(0)
    rows: list[list[str]] = []
    for r in range(sheet.nrows):
        row: list[str] = []
        for c in range(sheet.ncols):
            cell = sheet.cell(r, c)
            value = cell.value
            if isinstance(value, float) and value == int(value):
                value = int(value)
            row.append("" if value is None else str(value))
        rows.append(row)
    return rows


def _grid_from_xlsx(path: str) -> list[list[str]]:
    from openpyxl import load_workbook

    wb = load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    rows: list[list[str]] = []
    for row in ws.iter_rows(values_only=True):
        rows.append(["" if v is None else str(v) for v in row])
    return rows


def parse_schedule(path: str) -> dict[str, Any]:
    """解析教务导出的课表文件。"""
    ext = os.path.splitext(path)[1].lower()
    grid = _grid_from_xlsx(path) if ext in (".xlsx", ".xlsm") else _grid_from_xls(path)

    # 头部信息：学号/姓名/班级/学期
    head_text = "\n".join(" ".join(c for c in row if c) for row in grid[:4])
    student = ""
    m = re.search(r"([\u4e00-\u9fa5]{2,4})\s*学生个人课表", head_text)
    if m:
        student = m.group(1)
    term = ""
    m = re.search(r"学年学期[：:]\s*([0-9\-]+)", head_text)
    if m:
        term = m.group(1)
    klass = ""
    m = re.search(r"班级[：:]\s*([^\s]+)", head_text)
    if m:
        klass = m.group(1)

    # 找表头行
    header_idx = None
    day_cols: dict[int, str] = {}
    for i, row in enumerate(grid[:12]):
        hits = {j: wk for j, cell in enumerate(row) for wk in WEEKDAY_KEYS if wk in str(cell)}
        if len(hits) >= 5:
            header_idx = i
            day_cols = hits
            break
    if header_idx is None:
        raise ValueError("没找到星期表头，可能不是教务导出的课表文件")

    courses: list[dict[str, Any]] = []
    seen: set[tuple] = set()
    for row in grid[header_idx + 1 :]:
        label = str(row[0]).strip() if row else ""
        period_guess = PERIOD_LABEL.get(label, "")
        for col, weekday in sorted(day_cols.items()):
            cell = str(row[col]) if col < len(row) else ""
            if not cell.strip():
                continue
            for m in ENTRY_RE.finditer(cell):
                name = m.group("name").strip()
                teacher = m.group("teacher").strip()
                weeks = _parse_weeks(m.group("weeks"))
                period = m.group("period").strip()
                loc = (m.group("loc") or "").strip()
                if not weeks:
                    continue
                key = (name, weekday, tuple(weeks), period, loc)
                if key in seen:
                    continue
                seen.add(key)
                courses.append(
                    {
                        "name": name,
                        "teacher": teacher,
                        "day": weekday,
                        "weeks": weeks,
                        "period": period,
                        "period_label": period_guess,
                        "location": loc,
                    }
                )

    courses.sort(key=lambda c: (WEEKDAY_KEYS.index(c["day"]), _period_start(c["period"])))
    return {
        "student": student,
        "class": klass,
        "term": term,
        "courses": courses,
        "updated": int(time.time()),
    }


def _period_start(period: str) -> int:
    m = re.search(r"(\d+)", str(period))
    return int(m.group(1)) if m else 99


def _fmt_period(period: str) -> str:
    """把 01-02节 / 05-06-07-08-09-10节 收敛成 1-2节 / 5-10节"""
    nums = [int(n) for n in re.findall(r"\d+", str(period))]
    if not nums:
        return str(period).strip()
    if len(nums) == 1:
        return f"{nums[0]}节"
    consecutive = all(b == a + 1 for a, b in zip(nums, nums[1:]))
    return f"{nums[0]}-{nums[-1]}节" if consecutive else "-".join(str(n) for n in nums) + "节"


def _fmt_weeks(weeks: list[int]) -> str:
    if not weeks:
        return "?"
    out: list[str] = []
    start = prev = weeks[0]
    for w in weeks[1:]:
        if w == prev + 1:
            prev = w
            continue
        out.append(f"{start}" if start == prev else f"{start}-{prev}")
        start = prev = w
    out.append(f"{start}" if start == prev else f"{start}-{prev}")
    return ",".join(out)


# 每大节的上课时间（24 小时制），按学校实际可以改
PERIOD_START = {1: (8, 0), 3: (10, 0), 5: (14, 0), 7: (16, 0), 9: (19, 0)}


def _period_block(period: str) -> int:
    """把 "3-4节" 里的数字映射到 1/3/5/7/9 这几个大节的起点"""
    m = re.search(r"(\d+)", str(period))
    if not m:
        return 0
    n = int(m.group(1))
    if n <= 2:
        return 1
    if n <= 4:
        return 3
    if n <= 6:
        return 5
    if n <= 8:
        return 7
    return 9


def _period_dt(day: date, period: str) -> datetime | None:
    base = _period_block(period)
    if base not in PERIOD_START:
        return None
    hh, mm = PERIOD_START[base]
    return datetime(day.year, day.month, day.day, hh, mm)


PERIOD_MIN = 100  # 一大节（2 小节 + 课间）


def _period_span(period: str) -> tuple[str, str] | None:
    """把 "3-4节" 变成 ("10:00", "11:40")"""
    nums = [int(x) for x in re.findall(r"\d+", str(period))]
    if not nums:
        return None
    start_base = _period_block(str(nums[0]))
    end_base = _period_block(str(nums[-1]))
    st = PERIOD_START.get(start_base)
    et = PERIOD_START.get(end_base)
    if not st or not et:
        return None
    start = datetime(2000, 1, 1, *st)
    end = datetime(2000, 1, 1, *et) + timedelta(minutes=PERIOD_MIN)
    return f"{start:%H:%M}", f"{end:%H:%M}"


def next_lesson(data: dict[str, Any], after: datetime) -> tuple[datetime, dict[str, Any]] | None:
    """找 after 之后最近的一节课（最多往后找 21 天）"""
    term = data.get("term_start") or DEFAULT_TERM_START
    courses = data.get("courses", [])
    for offset in range(0, 22):
        day = after.date() + timedelta(days=offset)
        week = _week_of(term, day)
        if not week or week < 1:
            continue
        key = _weekday_key(day)
        best = None
        for c in courses:
            if c.get("day") != key or week not in c.get("weeks", []):
                continue
            dt = _period_dt(day, c.get("period", ""))
            if not dt or dt <= after:
                continue
            if best is None or dt < best[0]:
                best = (dt, c)
        if best:
            return best
    return None


def _human_delta(dt: datetime, now: datetime) -> str:
    mins = int((dt - now).total_seconds() // 60)
    if mins < 60:
        return f"{mins} 分钟后"
    h, m = divmod(mins, 60)
    if h < 24:
        return f"{h} 小时 {m} 分后" if m else f"{h} 小时后"
    d, h = divmod(h, 24)
    return f"{d} 天 {h} 小时后" if h else f"{d} 天后"


def render_lesson(title: str, dt: datetime, course: dict[str, Any], now: datetime | None = None) -> str:
    line = f"⏰ {title}  {dt:%m-%d %H:%M}"
    span = _period_span(course.get("period", ""))
    if span:
        first = f"{span[0]}-{span[1]}"
    else:
        first = f"{dt:%H:%M}"
    parts = [first, course.get("name", "")]
    if course.get("location"):
        parts.append(course["location"])
    out = [line, "   " + " · ".join(p for p in parts if p)]
    if now:
        out.append(f"   还有 {_human_delta(dt, now)}")
    return "\n".join(out)


def _span_dt(day: date, period: str) -> tuple[datetime, datetime] | None:
    """某天某节的开始/结束时间"""
    span = _period_span(period)
    if not span:
        return None
    h1, m1 = (int(x) for x in span[0].split(":"))
    h2, m2 = (int(x) for x in span[1].split(":"))
    return (
        datetime(day.year, day.month, day.day, h1, m1),
        datetime(day.year, day.month, day.day, h2, m2),
    )


def current_lesson(
    data: dict[str, Any], now: datetime
) -> tuple[datetime, datetime, dict[str, Any]] | None:
    """现在正在上的课"""
    term = data.get("term_start") or DEFAULT_TERM_START
    week = _week_of(term, now.date())
    if not week or week < 1:
        return None
    wd = _weekday_key(now.date())
    for c in data.get("courses", []):
        if c.get("day") != wd or week not in c.get("weeks", []):
            continue
        span = _span_dt(now.date(), c.get("period", ""))
        if not span:
            continue
        if span[0] <= now <= span[1]:
            return span[0], span[1], c
    return None


def groups_starting_now(now: datetime) -> dict[str, str]:
    """这一分钟里刚开课的群：umo -> platform"""
    wd = _weekday_key(now.date())
    targets: dict[str, str] = {}
    for fname in os.listdir(DATA_DIR):
        if not fname.endswith(".json") or fname == os.path.basename(WHO_STATE):
            continue
        try:
            with open(os.path.join(DATA_DIR, fname), encoding="utf-8-sig") as f:
                data = json.load(f)
        except Exception:  # noqa: BLE001
            continue
        umo = str(data.get("umo") or "")
        if not data.get("courses") or "GroupMessage" not in umo:
            continue
        term = data.get("term_start") or DEFAULT_TERM_START
        week = _week_of(term, now.date())
        if not week or week < 1:
            continue
        for c in data["courses"]:
            if c.get("day") != wd or week not in c.get("weeks", []):
                continue
            dt = _period_dt(now.date(), c.get("period", ""))
            if not dt:
                continue
            if 0 <= (now - dt).total_seconds() < 60:
                targets[umo] = str(data.get("platform") or "")
                break
    return targets


def who_lines(platform: str | None, now: datetime) -> list[str]:
    """「谁在上课」的正文：一个人一块，昵称在前"""
    people: list[tuple[datetime, str, str]] = []
    for fname in os.listdir(DATA_DIR):
        if not fname.endswith(".json"):
            continue
        try:
            with open(os.path.join(DATA_DIR, fname), encoding="utf-8-sig") as f:
                data = json.load(f)
        except Exception:  # noqa: BLE001
            continue
        if not data.get("courses"):
            continue
        if not data.get("visible", True):
            continue
        if platform and data.get("platform") != platform:
            continue
        found = current_lesson(data, now)
        if not found:
            continue
        st, en, c = found
        detail = " · ".join(
            p
            for p in (
                c.get("name", ""),
                f"{st:%H:%M}-{en:%H:%M}",
                c.get("location", ""),
            )
            if p
        )
        people.append((st, data.get("nickname") or "同学", detail))

    if not people:
        return ["现在这个点没人上课喵～ (￣▽￣)"]
    people.sort(key=lambda x: (x[0], x[1]))
    lines = [f"👀 现在上课的有 {len(people)} 位"]
    for _st, nick, detail in people:
        lines.append("")
        lines.append(nick)
        lines.append(f"  {detail}")
    lines.append("")
    lines.append("不想被看到就发 /隐身")
    return lines


def catch_lines(platform: str | None, now: datetime) -> list[str]:
    """「抓群友」：现在去哪个楼哪个教室能抓到谁（只给地点和人，不给课程名）"""
    rooms: dict[str, list[tuple[datetime, str]]] = {}
    for fname in os.listdir(DATA_DIR):
        if not fname.endswith(".json"):
            continue
        try:
            with open(os.path.join(DATA_DIR, fname), encoding="utf-8-sig") as f:
                data = json.load(f)
        except Exception:  # noqa: BLE001
            continue
        if not data.get("courses") or not data.get("visible", True):
            continue
        if platform and data.get("platform") != platform:
            continue
        found = current_lesson(data, now)
        if not found:
            continue
        _st, en, c = found
        room = str(c.get("location") or "").strip() or "（教室没写）"
        rooms.setdefault(room, []).append((en, str(data.get("nickname") or "同学")))

    if not rooms:
        return ["🕵️ 现在这个点，一个能抓的都没有喵～ (￣▽￣)"]

    total = sum(len(v) for v in rooms.values())
    lines = [f"🕵️ 现在能抓到 {total} 位"]
    for room in sorted(rooms):
        names = sorted({n for _e, n in rooms[room]})
        lines.append("")
        lines.append(f"📍 {room}")
        lines.append("  " + "、".join(names))
    lines.append("")
    lines.append("不想被看到就发 /隐身")
    return lines


def _stats_load() -> dict[str, Any]:
    try:
        with open(STATS_FILE, encoding="utf-8-sig") as f:
            d = json.load(f)
        if not isinstance(d.get("days"), dict):
            d["days"] = {}
        return d
    except Exception:  # noqa: BLE001
        return {"days": {}}


def _stats_save(d: dict[str, Any]) -> None:
    days = d.get("days") or {}
    if len(days) > 60:  # 只留最近两个月
        for k in sorted(days)[:-60]:
            days.pop(k, None)
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(STATS_FILE, "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False, indent=2)


def _bump(day: date, kind: str, qq: str, nick: str) -> None:
    """给某人某天的某一项（query / catch）+1"""
    if not qq:
        return
    d = _stats_load()
    slot = d["days"].setdefault(str(day), {})
    kind_map = slot.setdefault(kind, {})
    item = kind_map.setdefault(str(qq), {"n": 0, "nick": nick})
    item["n"] = int(item.get("n") or 0) + 1
    if nick:
        item["nick"] = nick
    _stats_save(d)


def day_load(data: dict[str, Any], day: date) -> tuple[int, int]:
    """某人某天上了多少分钟课 / 几门课的块数"""
    term = data.get("term_start") or DEFAULT_TERM_START
    week = _week_of(term, day)
    if not week or week < 1:
        return 0, 0
    wd = _weekday_key(day)
    mins, cnt = 0, 0
    for c in data.get("courses", []):
        if c.get("day") != wd or week not in c.get("weeks", []):
            continue
        span = _span_dt(day, c.get("period", ""))
        if not span:
            continue
        mins += int((span[1] - span[0]).total_seconds() // 60)
        cnt += 1
    return mins, cnt


def _fmt_min(m: int) -> str:
    m = int(m)
    h, mm = divmod(m, 60)
    if h and mm:
        return f"{h} 小时 {mm} 分"
    if h:
        return f"{h} 小时"
    return f"{mm} 分钟"


def rank_rows(day: date, limit: int = 8) -> list[tuple[str, str, float]]:
    """当天上课时长排行：[(昵称, 时长文本, 比例)]"""
    users: list[tuple[str, int]] = []
    for fname in os.listdir(DATA_DIR):
        if not fname.endswith(".json") or fname in INTERNAL_FILES:
            continue
        try:
            with open(os.path.join(DATA_DIR, fname), encoding="utf-8-sig") as f:
                data = json.load(f)
        except Exception:  # noqa: BLE001
            continue
        if not data.get("courses"):
            continue
        mins, _cnt = day_load(data, day)
        users.append((str(data.get("nickname") or "同学"), mins))
    if not users:
        return []
    users.sort(key=lambda u: (-u[1], u[0]))
    peak = max(u[1] for u in users) or 1
    return [(n, _fmt_min(m), m / peak) for n, m in users[:limit]]


def build_rank_image(day: date) -> str:
    """画上课排行榜，返回图片路径"""
    import image as _img  # noqa: PLC0415

    out = os.path.join(DATA_DIR, f"rank_{day:%Y%m%d}.png")
    weekday = "一二三四五六日"[day.weekday()]
    return _img.render_leaderboard(
        out,
        f"{day:%m-%d} 上课排行榜",
        rank_rows(day),
        subtitle=f"周{weekday} · 当天上课总时长",
        footer="按课表节次计算 · 每天 0 点自动发前一天的",
    )


def _parse_iso(text: str) -> date | None:
    m = re.search(r"(\d{4})[-/.年](\d{1,2})[-/.月](\d{1,2})", str(text))
    if not m:
        return None
    try:
        return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    except ValueError:
        return None


def _today() -> date:
    return datetime.now().date()


def _week_range(term_start: str | None, week: int) -> tuple[date, date] | None:
    start = _parse_iso(term_start) if term_start else None
    if not start:
        return None
    monday = start + timedelta(weeks=week - 1)
    return monday, monday + timedelta(days=6)


def _week_of(term_start: str | None, day: date) -> int | None:
    start = _parse_iso(term_start) if term_start else None
    if not start:
        return None
    delta = (day - start).days
    if delta < 0:
        return 0
    return delta // 7 + 1


def _weekday_key(day: date) -> str:
    return WEEKDAY_KEYS[day.weekday()]


def render_schedule(
    data: dict[str, Any],
    week: int | None = None,
    day: str | None = None,
    term_start: str | None = None,
    nickname: str | None = None,
) -> str:
    courses = data.get("courses", [])
    if week is not None:
        courses = [c for c in courses if week in c.get("weeks", [])]
    if day:
        courses = [c for c in courses if c.get("day") == day]

    who = nickname or data.get("nickname") or "同学"
    title = f"📅 {data.get('term') or '本学期'} · {who}"
    if week is not None:
        title += f" · 第 {week} 周"
        rng = _week_range(term_start, week)
        if rng:
            title += f"（{rng[0]:%m-%d} ~ {rng[1]:%m-%d}）"
    if day:
        title += f" · {WEEKDAYS[WEEKDAY_KEYS.index(day)]}"

    if not courses:
        if week is not None or day:
            return f"{title}\n\n这段时间没有课喵 (๑>ᴗ<๑)"
        return IMPORT_GUIDE

    lines = [title, ""]
    cur_day = None
    for c in courses:
        if c["day"] != cur_day:
            cur_day = c["day"]
            lines.append(f"【{WEEKDAYS[WEEKDAY_KEYS.index(cur_day)]}】")
        span = _period_span(c["period"])
        head = _fmt_period(c["period"])
        if span:
            head += f" {span[0]}-{span[1]}"
        parts = [head, c["name"]]
        if c["location"]:
            parts.append(c["location"])
        lines.append("  " + " · ".join(parts))
    lines.append("")
    names = {c["name"] for c in courses}
    lines.append(f"共 {len(names)} 门课 / {len(courses)} 个时间块")
    return "\n".join(lines)


HELP_TEXT = """📖 课表插件 · 命令总览

【查课表】
   /周课表         这周的课（出图）
   /课表 全部      整学期（出图，或 /课表 完整）
   /今日课表       今天上什么（文字）
   /下一节课       现实世界里下一节课

【同学】
   /谁在上课       看看现在谁在上课
   /抓群友         现在去哪个楼哪个教室能抓谁
   /排行榜         今天谁上课最多（前三名有奖牌）
   /排行榜 昨天    看昨天的
   /隐身           让别人查不到你
   /在线           取消隐身

【提醒】
   /上课提醒       查看提醒状态
   /上课提醒开     开启上课提醒
   /上课提醒关     关闭上课提醒

【播报】🔒 管理员
   /播报开         每节课一开始就在群里发「谁在上课」
   /播报关         关掉这个播报

【管理员】🔒
   /管理员          看这个插件认哪些人当管理员
   /管理员 添加 @某人
   /管理员 删除 @某人

【其他】
   /课表 开学 2026-09-07   设置第 1 周的周一
   /课表 删除              清掉已存课表

（这些词不加 / 直接打也能用：周课表、今日课表、下一节课、谁在上课、抓群友、排行榜、上课提醒、课表 全部）"""


IMPORT_GUIDE = """📖 还没有你的课表喵～ 三步就能导入 (๑>ᴗ<๑)

1️⃣ 打开湖南科技大学教务系统（电脑上更稳，手机可能下载不了）
https://kdjw.hnust.edu.cn/jsxsd/sso.jsp?jumpUrl=/xskb/xskb_list.do
用统一身份认证账号自己登录

2️⃣ 进「学期理论课表」→ 点右上角「导出」
会下载一个 `学生个人课表_学号.xls`

3️⃣ 把这个 xls 文件直接发给我（当文件发，别截图）
   我收到就自动解析排好

🔒 隐私怎么选
· 不想被别人看到 → **私聊发给我**
· 无所谓 → 群里直接发也行

导入完发 /课表 就能看啦～"""


@register("astrbot_plugin_kebiao", "sanyi", "课表导入与排课（MVP）", "0.1.0")
class KebiaoPlugin(Star):
    def __init__(self, context: Context, config: dict[str, Any] | None = None) -> None:
        super().__init__(context)
        self.config = config or {}
        os.makedirs(DATA_DIR, exist_ok=True)
        self._sent: set[str] = set()
        self._reminder_task = asyncio.create_task(self._reminder_loop())

    # ---------- 存储 ----------
    def _key(self, event: AstrMessageEvent) -> str:
        """按「平台 + QQ号」存，一个 QQ 一份课表（群里的课表不会互相覆盖）"""
        platform = ""
        sender = ""
        try:
            platform = str(event.get_platform_id() or "")
        except Exception:  # noqa: BLE001
            pass
        try:
            sender = str(event.get_sender_id() or "")
        except Exception:  # noqa: BLE001
            pass
        return f"{platform}_{sender}" if sender else event.unified_msg_origin

    def _platform(self, event: AstrMessageEvent) -> str:
        try:
            return str(event.get_platform_id() or "")
        except Exception:  # noqa: BLE001
            return ""

    # ---------- 管理员判定 ----------
    def _bot_admins(self) -> list[str]:
        """AstrBot 主配置里 admins_id 的人"""
        try:
            cfg = self.context.get_config()
            return [str(x) for x in (cfg.get("admins_id") or [])]
        except Exception:  # noqa: BLE001
            return []

    def _is_admin(self, event: AstrMessageEvent) -> bool:
        """三类人都算管理员：AstrBot 主配置管理员、插件名单、群里/群主"""
        uid = str(event.get_sender_id() or "").strip()
        if uid and uid in self._bot_admins():
            return True
        if uid and uid in _load_admins():
            return True
        try:
            if event.is_admin():
                return True
        except Exception:  # noqa: BLE001
            pass
        # 群里的群主 / 管理员也算
        try:
            raw = getattr(event.message_obj, "raw_message", None) or {}
            role = str(((raw.get("sender") or {}).get("role")) or "")
            if str(event.get_group_id() or "") and role in ("owner", "admin"):
                return True
        except Exception:  # noqa: BLE001
            pass
        return False

    def _need_admin(self, event: AstrMessageEvent) -> str | None:
        """不是管理员就返回提示文案"""
        if self._is_admin(event):
            return None
        return "🔒 这个只有管理员能用喵～\n（群里群主/管理员、或 /管理员 添加 过的人）"

    def _nick(self, event: AstrMessageEvent) -> str:
        """用群昵称代替真实姓名，课表里只显示昵称"""
        for getter in ("get_sender_name",):
            try:
                name = str(getattr(event, getter)() or "").strip()
            except Exception:  # noqa: BLE001
                name = ""
            if name:
                return name
        try:
            sender = str(event.get_sender_id() or "").strip()
        except Exception:  # noqa: BLE001
            sender = ""
        return sender or "同学"

    def _count(self, event: AstrMessageEvent, kind: str) -> None:
        """行为计数（排行榜用）：kind = query / catch"""
        try:
            _bump(
                _today(),
                kind,
                str(event.get_sender_id() or ""),
                self._nick(event),
            )
        except Exception:  # noqa: BLE001
            logger.exception(f"[kebiao] 计数失败 {kind}")

    def _path(self, key: str) -> str:
        safe = re.sub(r"[^0-9a-zA-Z_-]", "_", key)
        return os.path.join(DATA_DIR, f"{safe}.json")

    def _load(self, key: str) -> dict[str, Any] | None:
        p = self._path(key)
        if not os.path.exists(p):
            return None
        try:
            with open(p, encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:  # noqa: BLE001
            logger.warning(f"[kebiao] 读取失败 {p}: {e}")
            return None

    def _save(self, key: str, data: dict[str, Any]) -> None:
        with open(self._path(key), "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=1)

    # ---------- 会话/提醒信息 ----------
    def _touch(self, event: AstrMessageEvent, key: str, data: dict[str, Any]) -> None:
        """记住这个人最近在哪个会话、QQ 多少，提醒时往这里发"""
        changed = False
        umo = getattr(event, "unified_msg_origin", "") or ""
        if umo and data.get("umo") != umo:
            data["umo"] = umo
            changed = True
        qq = ""
        try:
            qq = str(event.get_sender_id() or "")
        except Exception:  # noqa: BLE001
            pass
        if qq and data.get("qq") != qq:
            data["qq"] = qq
            changed = True
        nick = self._nick(event)
        if nick and data.get("nickname") != nick:
            data["nickname"] = nick
            changed = True
        if data.get("platform") != self._platform(event):
            data["platform"] = self._platform(event)
            changed = True
        if "reminder" not in data:
            data["reminder"] = True
            changed = True
        if "visible" not in data:
            data["visible"] = True
            changed = True
        if changed:
            self._save(key, data)

    # ---------- 命令 ----------
    @filter.command("课表")
    async def kebiao_cmd(self, event: AstrMessageEvent):
        """查看课表 / 帮助"""
        arg = event.message_str.strip()
        arg = re.sub(r"^/?课表\s*", "", arg).strip()
        key = self._key(event)

        data = self._load(key)
        if not data:
            # 兼容旧版本：以前是按会话（umo）存的
            legacy = self._load(event.unified_msg_origin)
            if legacy:
                data = legacy
                self._save(key, data)
                try:
                    os.remove(self._path(event.unified_msg_origin))
                except OSError:
                    pass
        if not data:
            # 没课表的人，任何 /课表 都先看导入教程
            yield event.plain_result(IMPORT_GUIDE)
            return

        if arg in ("", "帮助", "help", "?", "？", "导入", "菜单"):
            yield event.plain_result(HELP_TEXT)
            return
        if arg in ("删除", "清除", "清除课表"):
            p = self._path(key)
            if os.path.exists(p):
                os.remove(p)
                yield event.plain_result("已清掉课表喵，随时重新导入～")
            else:
                yield event.plain_result("本来就没有存的课表呀 (・ω・)")
            return

        self._touch(event, key, data)

        if arg in ("全部", "完整", "整学期", "所有", "全"):
            yield event.image_result(self._render_image(event, key, data, None))
            return

        m = re.match(r"^(?:开学|设置|开学日)\s*(\S+)$", arg)
        if m:
            d = _parse_iso(m.group(1))
            if not d:
                yield event.plain_result("日期看不懂喵，格式像这样：/课表 开学 2026-09-07")
                return
            data["term_start"] = d.isoformat()
            self._save(key, data)
            wk = _week_of(d.isoformat(), _today())
            yield event.plain_result(
                f"✅ 已设置：第 1 周周一是 {d:%Y-%m-%d}\n今天（{_today():%m-%d}）是第 {wk} 周"
            )
            return

        # 其他参数一律回到帮助，避免半懂不懂的用法
        yield event.plain_result(HELP_TEXT)

    def _render_image(self, event: AstrMessageEvent, key: str, data: dict[str, Any], week: int | None) -> str:
        """把课表画成图：week=None 画整学期，否则只画那一周"""
        term_start = data.get("term_start") or DEFAULT_TERM_START
        courses: list[dict[str, Any]] = []
        for c in data.get("courses", []):
            if week is not None and week not in c.get("weeks", []):
                continue
            item = dict(c)
            item["label"] = _fmt_weeks(c.get("weeks", []))
            courses.append(item)
        term = data.get("term") or "本学期"
        who = self._nick(event)
        if week is None:
            subtitle = f"{who} · 整学期"
        else:
            rng = _week_range(term_start, week)
            span = f"（{rng[0]:%m-%d} ~ {rng[1]:%m-%d}）" if rng else ""
            subtitle = f"{who} · 第 {week} 周{span}"
        safe = re.sub(r"[^\w\-]", "_", f"{key}_{week if week is not None else 'all'}")[:60]
        out = os.path.join(DATA_DIR, "img", f"{safe}.png")
        return render_grid(
            courses,
            out,
            title=term,
            subtitle=subtitle,
            show_weeks=(week is None),
            footer="课表插件 · /周课表 看本周 · /今日课表 看今天",
        )

    # ---------- 周课表（图） ----------
    @filter.command("周课表")
    async def week_cmd(self, event: AstrMessageEvent):
        """本周课表（出图）"""
        self._count(event, "query")
        key = self._key(event)
        data = self._load(key) or self._load(event.unified_msg_origin)
        if not data:
            yield event.plain_result(IMPORT_GUIDE)
            return
        self._touch(event, key, data)
        week = _week_of(data.get("term_start") or DEFAULT_TERM_START, _today())
        if not week or week < 1:
            yield event.plain_result("还没开学喵 (・ω・) 用 /课表 开学 2026-09-07 设一下第一周")
            return
        courses = [c for c in data.get("courses", []) if week in c.get("weeks", [])]
        if not courses:
            yield event.plain_result(f"第 {week} 周没有课喵，放假去吧 (๑>ᴗ<๑)")
            return
        yield event.image_result(self._render_image(event, key, data, week))

    # ---------- 免 / 触发（照抄抽老婆插件的思路） ----------
    def _bare_action(self, text: str):
        """裸词 → 对应命令处理函数；不认识的返回 None"""
        if text in ("周课表", "本周课表", "这周课表"):
            return self.week_cmd
        if text in ("今日课表", "今天课表", "今天上什么"):
            return self.today_cmd
        if text in ("下一节课", "下一节", "下节课"):
            return self.next_cmd
        if text in ("谁在上课",):
            return self.who_cmd
        if text in ("抓群友", "抓人", "抓群友在哪"):
            return self.catch_cmd
        if text in ("排行榜", "看排行榜", "日报"):
            return self.rank_cmd
        if text == "上课提醒":
            return self.remind_cmd
        if text in ("上课提醒开", "开提醒"):
            return self.on_cmd
        if text in ("上课提醒关", "关提醒"):
            return self.off_cmd
        if text == "课表":
            return self.kebiao_cmd
        m = re.match(r"^课表\s*(\S+)$", text)
        if m and m.group(1) in ("全部", "完整", "整学期", "所有", "全", "帮助", "导入"):
            return self.kebiao_cmd
        return None

    @filter.event_message_type(filter.EventMessageType.ALL)
    async def bare_trigger(self, event: AstrMessageEvent):
        """群/私聊里直接打「周课表」这种词就能用，不用 / 也不用 @"""
        if not (self.config or {}).get("keyword_trigger_enabled", True):
            return
        text = (event.message_str or "").strip()
        if not text:
            return
        # 已经被 @ 或加前缀唤醒的，交给 @filter.command 处理，别双触发
        if event.is_at_or_wake_command:
            return
        if text.startswith(("/", "!", "！")):
            return
        handler = self._bare_action(text)
        if handler is None:
            return
        async for result in handler(event):
            yield result
        event.stop_event()

    # ---------- 今日 / 下一节 ----------
    @filter.command("今日课表")
    async def today_cmd(self, event: AstrMessageEvent):
        """今天的课表"""
        self._count(event, "query")
        key = self._key(event)
        data = self._load(key) or self._load(event.unified_msg_origin)
        if not data:
            yield event.plain_result(IMPORT_GUIDE)
            return
        self._touch(event, key, data)
        today = _today()
        term = data.get("term_start") or DEFAULT_TERM_START
        week = _week_of(term, today)
        yield event.plain_result(render_schedule(data, week, _weekday_key(today), term, self._nick(event)))

    @filter.command("下一节课")
    async def next_cmd(self, event: AstrMessageEvent):
        """现实世界里下一节课"""
        self._count(event, "query")
        key = self._key(event)
        data = self._load(key) or self._load(event.unified_msg_origin)
        if not data:
            yield event.plain_result(IMPORT_GUIDE)
            return
        self._touch(event, key, data)
        now = datetime.now()
        found = next_lesson(data, now)
        if not found:
            yield event.plain_result("接下来三周都没有课喵 (￣▽￣)")
            return
        dt, course = found
        yield event.plain_result(render_lesson("下一节课", dt, course, now))

    # ---------- 提醒开关 ----------
    @filter.command("上课提醒")
    async def remind_cmd(self, event: AstrMessageEvent):
        """提醒的二级界面"""
        key = self._key(event)
        data = self._load(key) or self._load(event.unified_msg_origin)
        if not data:
            yield event.plain_result(IMPORT_GUIDE)
            return
        self._touch(event, key, data)
        on = bool(data.get("reminder", True))
        state = "✅ 已开启" if on else "⏸️ 已关闭"
        target = data.get("remind_umo") or data.get("umo") or ""
        where = "群聊" if "GroupMessage" in str(target) else ("私聊" if target else "未确定")
        yield event.plain_result(
            f"⏰ 上课提醒\n\n当前状态：{state}\n"
            f"提醒发送到：{where}\n"
            f"开着的话，每节课开始前 {REMIND_MIN} 分钟我会在「开启提醒的那个会话」@你，"
            f"告诉你几点、在哪、上什么课。\n\n"
            f"/上课提醒开    开启（在哪开就往哪发）\n/上课提醒关    关闭"
        )

    @filter.command("上课提醒开")
    async def on_cmd(self, event: AstrMessageEvent):
        """开启提醒"""
        key = self._key(event)
        data = self._load(key) or self._load(event.unified_msg_origin)
        if not data:
            yield event.plain_result(IMPORT_GUIDE)
            return
        self._touch(event, key, data)
        data["reminder"] = True
        data["remind_umo"] = event.unified_msg_origin
        self._save(key, data)
        where = "群里" if "GroupMessage" in str(event.unified_msg_origin) else "私聊"
        yield event.plain_result(
            f"✅ 提醒已开启喵，上课前 {REMIND_MIN} 分钟我会在这个{where}@你 (๑>ᴗ<๑)"
        )

    @filter.command("上课提醒关")
    async def off_cmd(self, event: AstrMessageEvent):
        """关闭提醒"""
        key = self._key(event)
        data = self._load(key) or self._load(event.unified_msg_origin)
        if not data:
            yield event.plain_result(IMPORT_GUIDE)
            return
        self._touch(event, key, data)
        data["reminder"] = False
        self._save(key, data)
        yield event.plain_result("⏸️ 提醒已关闭，想开回来发 /上课提醒开 就行喵")

    # ---------- 开课播报开关 ----------
    @filter.command("播报开")
    async def who_push_on_cmd(self, event: AstrMessageEvent):
        """每节课开始就播报谁在上课（管理员）"""
        deny = self._need_admin(event)
        if deny:
            yield event.plain_result(deny)
            return
        self._set_who_push(True)
        yield event.plain_result(
            "📣 播报已开：以后每节课一到点，我就把「谁在上课」发到对应群里\n"
            "（不想看了发 /播报关）"
        )

    @filter.command("播报关")
    async def who_push_off_cmd(self, event: AstrMessageEvent):
        """关掉开课播报（管理员）"""
        deny = self._need_admin(event)
        if deny:
            yield event.plain_result(deny)
            return
        self._set_who_push(False)
        yield event.plain_result(f"🔇 播报已关，上课前 {REMIND_MIN} 分钟的单人提醒照常")

    @filter.command("管理员")
    async def admin_cmd(self, event: AstrMessageEvent):
        """管理这个插件的管理员名单"""
        text = (event.message_str or "").strip()
        text = re.sub(r"^[/!！]?\s*管理员", "", text).strip()
        tokens = text.split()
        sub = tokens[0] if tokens else ""

        if sub in ("", "列表", "查看"):
            cfg = self._bot_admins()
            extra = _load_admins()
            lines = ["👑 这个插件的管理员", ""]
            lines.append("① AstrBot 主配置里的：" + ("、".join(cfg) if cfg else "（没有）"))
            lines.append("② 插件里加的：" + ("、".join(extra) if extra else "（没有）"))
            lines.append("③ 所在群的群主 / 群管理员也算")
            lines += [
                "",
                "/管理员 添加 @某人    加一个",
                "/管理员 删除 @某人    去掉一个",
                "",
                "（播报开 / 播报关 只有管理员能用）",
            ]
            yield event.plain_result("\n".join(lines))
            return

        deny = self._need_admin(event)
        if deny:
            yield event.plain_result(deny)
            return
        if sub not in ("添加", "删除", "移除"):
            yield event.plain_result("用法：/管理员 添加 @某人  或  /管理员 删除 @某人")
            return

        targets: list[str] = []
        try:
            for seg in event.get_messages():
                if isinstance(seg, At):
                    targets.append(str(seg.qq))
        except Exception:  # noqa: BLE001
            pass
        for tok in tokens[1:]:
            if re.fullmatch(r"\d{5,12}", tok):
                targets.append(tok)
        targets = [t for t in dict.fromkeys(targets) if t and t != "all"]
        if not targets:
            yield event.plain_result("没看懂要加谁喵，@ 一下人或把 QQ 号写出来")
            return

        extra = _load_admins()
        if sub == "添加":
            added = [t for t in targets if t not in extra]
            _save_admins(extra + added)
            yield event.plain_result(
                ("✅ 加好了：" + "、".join(added)) if added else "这几个人本来就在名单里喵～"
            )
        else:
            keep = [t for t in extra if t not in targets]
            _save_admins(keep)
            yield event.plain_result("🗑️ 去掉了：" + "、".join(targets))


    # ---------- 谁在上课 / 隐身 ----------
    @filter.command("谁在上课")
    async def who_cmd(self, event: AstrMessageEvent):
        """现在有谁在上课"""
        now = datetime.now()
        platform = self._platform(event)
        yield event.plain_result("\n".join(who_lines(platform, now)))

    @filter.command("抓群友")
    async def catch_cmd(self, event: AstrMessageEvent):
        """现在去哪个楼哪个教室能抓到谁"""
        self._count(event, "catch")
        now = datetime.now()
        platform = self._platform(event)
        yield event.plain_result("\n".join(catch_lines(platform, now)))

    @filter.command("排行榜")
    async def rank_cmd(self, event: AstrMessageEvent):
        """今日排行榜（每天 0 点也会自动发昨天的）"""
        arg = (
            event.message_str.replace("排行榜", "").replace("/", "").strip()
        )
        day = _today()
        if arg in ("昨天", "昨日"):
            day = day - timedelta(days=1)
        elif arg:
            parsed = _parse_iso(arg)
            if parsed:
                day = parsed
        try:
            path = build_rank_image(day)
        except Exception:  # noqa: BLE001
            logger.exception("[kebiao] 排行榜出图失败")
            yield event.plain_result("图没画出来喵，去日志里看看")
            return
        yield event.image_result(path)

    @filter.command("at测试")
    async def at_test_cmd(self, event: AstrMessageEvent):
        """测一下真 @ 到底通不通（管理员）"""
        deny = self._need_admin(event)
        if deny:
            yield event.plain_result(deny)
            return
        uid = str(event.get_sender_id() or "")
        yield event.chain_result(
            [
                Plain("🔔 @ 测试："),
                At(qq=uid),
                Plain(" 如果上面是蓝色可点的名字，就是真 @；灰色字就是假 @。"),
            ]
        )

    @filter.command("隐身")
    async def hide_cmd(self, event: AstrMessageEvent):
        """隐身：不再出现在「谁在上课」里"""
        key = self._key(event)
        data = self._load(key) or self._load(event.unified_msg_origin)
        if not data:
            yield event.plain_result(IMPORT_GUIDE)
            return
        self._touch(event, key, data)
        data["visible"] = False
        self._save(key, data)
        yield event.plain_result("🫥 已隐身，别人用 /谁在上课 查不到你辽（提醒照常推送）")

    @filter.command("在线")
    async def show_cmd(self, event: AstrMessageEvent):
        """取消隐身"""
        key = self._key(event)
        data = self._load(key) or self._load(event.unified_msg_origin)
        if not data:
            yield event.plain_result(IMPORT_GUIDE)
            return
        self._touch(event, key, data)
        data["visible"] = True
        self._save(key, data)
        yield event.plain_result("👀 已上线，别人又能看到你在上什么课了喵")

    @filter.command("提醒测试")
    async def reminder_debug(self, event: AstrMessageEvent):
        """看看接下来一周的提醒时间算得对不对"""
        key = self._key(event)
        data = self._load(key) or self._load(event.unified_msg_origin)
        if not data:
            yield event.plain_result(IMPORT_GUIDE)
            return
        now = datetime.now()
        term = data.get("term_start") or DEFAULT_TERM_START
        week = _week_of(term, now.date())
        lines = [
            f"🕒 现在 {now:%m-%d %H:%M}（第 {week} 周）",
            f"开关：{'开' if data.get('reminder', True) else '关'} · 推送目标：{data.get('remind_umo') or '未设置'}",
            "",
        ]
        shown = 0
        for i in range(8):
            day = now.date() + timedelta(days=i)
            wk = _week_of(term, day)
            if not wk or wk < 1:
                continue
            wd = _weekday_key(day)
            same = [c for c in data.get("courses", []) if c.get("day") == wd and wk in c.get("weeks", [])]
            for c in sorted(same, key=lambda x: str(x.get("period", ""))):
                dt = _period_dt(day, c.get("period", ""))
                if not dt:
                    continue
                remind = dt - timedelta(minutes=REMIND_MIN)
                mark = "" if remind > now else "  ← 已过"
                lines.append(f"{remind:%m-%d %H:%M}  {c.get('name','')} {c.get('location','')}{mark}")
                shown += 1
        if not shown:
            lines.append("接下来一周没有课喵 (・ω・)")
        yield event.plain_result("\n".join(lines))

    # ---------- 定时提醒 ----------
    @filter.command("提醒自检")
    async def reminder_selftest(self, event: AstrMessageEvent):
        """立刻往推送目标发一条测试提醒，验证通道是否通"""
        key = self._key(event)
        data = self._load(key) or self._load(event.unified_msg_origin)
        if not data:
            yield event.plain_result(IMPORT_GUIDE)
            return
        umo = data.get("remind_umo") or data.get("umo")
        qq = data.get("qq")
        if not umo:
            yield event.plain_result("还没设置推送目标，先发一次课表给我，或开一次上课提醒喵")
            return
        chain = MessageChain(
            [
                Plain("⏰ 这是一条自检提醒喵～\n"),
                At(qq=qq, name=data.get("nickname") or "同学"),
                Plain("  如果你能在"),
                Plain(f" {umo.split(':')[-1]} "),
                Plain("收到这条，说明提醒通道是通的\n"),
            ]
        )
        try:
            await self.context.send_message(umo, chain)
        except Exception as e:  # noqa: BLE001
            logger.exception("[kebiao] 自检发送失败")
            yield event.plain_result(f"❌ 发送失败：{e.__class__.__name__}: {e}")
            return
        yield event.plain_result(f"✅ 已发到 {umo}，去看看收到没～")

    async def _reminder_loop(self) -> None:
        logger.info("[kebiao] 提醒循环已启动")
        await asyncio.sleep(15)
        ticks = 0
        while True:
            try:
                await self._tick_reminders()
                await self._tick_who()
                await self._tick_leaderboard()
                ticks += 1
                if ticks % 40 == 1:  # 每 20 分钟留一行心跳，方便排查
                    logger.info(f"[kebiao] 提醒心跳 #{ticks} · {datetime.now():%H:%M:%S}")
            except Exception:  # noqa: BLE001
                logger.exception("[kebiao] 提醒循环出错")
            await asyncio.sleep(30)

    async def _tick_reminders(self) -> None:
        now = datetime.now()
        bucket: dict[str, list[tuple[datetime, dict[str, Any], str, str]]] = {}
        for fname in os.listdir(DATA_DIR):
            if not fname.endswith(".json"):
                continue
            try:
                with open(os.path.join(DATA_DIR, fname), encoding="utf-8-sig") as f:
                    data = json.load(f)
            except Exception:  # noqa: BLE001
                continue
            if not data.get("courses"):
                continue
            if not data.get("reminder", True):
                continue
            umo = data.get("remind_umo") or data.get("umo")
            qq = data.get("qq")
            if not (umo and qq):
                continue
            term = data.get("term_start") or DEFAULT_TERM_START
            week = _week_of(term, now.date())
            if not week or week < 1:
                continue
            wd = _weekday_key(now.date())
            for c in data.get("courses", []):
                if c.get("day") != wd or week not in c.get("weeks", []):
                    continue
                dt = _period_dt(now.date(), c.get("period", ""))
                if not dt:
                    continue
                lead = (dt - now).total_seconds() / 60
                if not (REMIND_MIN - 1.0 <= lead <= REMIND_MIN + 0.5):
                    continue
                tag = f"{fname}|{now.date()}|{c.get('period')}|{c.get('name')}"
                if tag in self._sent:
                    continue
                self._sent.add(tag)
                bucket.setdefault(umo, []).append(
                    (dt, c, qq, data.get("nickname") or "同学")
                )

        for umo, items in bucket.items():
            items.sort(key=lambda x: (x[0], x[1].get("name", "")))
            chain = [Plain(f"⏰ {REMIND_MIN} 分钟后上课喵～\n")]
            is_group = "GroupMessage" in umo
            for dt, c, qq, nick in items:
                span = _period_span(c.get("period", ""))
                first = f"{span[0]}-{span[1]}" if span else f"{dt:%H:%M}"
                parts = [first, c.get("name", "")]
                if c.get("location"):
                    parts.append(c["location"])
                if is_group:
                    chain.append(At(qq=qq, name=nick))
                chain.append(Plain("  " + " · ".join(p for p in parts if p) + "\n"))
            try:
                await self.context.send_message(umo, MessageChain(chain))
                logger.info(f"[kebiao] 提醒已发 {umo} · {len(items)} 条")
            except Exception:  # noqa: BLE001
                logger.exception(f"[kebiao] 提醒发送失败 {umo}")

    # ---------- 开课即播报「谁在上课」 ----------
    def _who_push_on(self) -> bool:
        try:
            with open(WHO_STATE, encoding="utf-8-sig") as f:
                return bool(json.load(f).get("enabled", True))
        except Exception:  # noqa: BLE001
            return True

    def _set_who_push(self, on: bool) -> None:
        try:
            with open(WHO_STATE, "w", encoding="utf-8") as f:
                json.dump({"enabled": bool(on)}, f)
        except Exception:  # noqa: BLE001
            logger.exception("[kebiao] 播报开关写入失败")

    async def _tick_who(self) -> None:
        """每到有一节课刚开始，就把「谁在上课」发到那个群"""
        if not self._who_push_on():
            return
        now = datetime.now()
        targets = groups_starting_now(now)
        for umo, platform in targets.items():
            tag = f"who|{umo}|{now:%Y-%m-%d %H:%M}"
            if tag in self._sent:
                continue
            self._sent.add(tag)
            lines = who_lines(platform, now)
            if not lines or "没人上课" in lines[0]:
                logger.info(f"[kebiao] 播报跳过（这个点没人）{umo}")
                continue
            try:
                await self.context.send_message(umo, MessageChain([Plain("\n".join(lines))]))
                logger.info(f"[kebiao] 播报已发 {umo} · {now:%H:%M}")
            except Exception:  # noqa: BLE001
                logger.exception(f"[kebiao] 播报发送失败 {umo}")

    # ---------- 每天 0 点的排行榜 ----------
    async def _tick_leaderboard(self) -> None:
        now = datetime.now()
        if now.hour != 0 or now.minute > 3:
            return
        day = (now - timedelta(days=1)).date()
        stamp = str(day)
        try:
            with open(LB_STATE, encoding="utf-8-sig") as f:
                if json.load(f).get("last") == stamp:
                    return
        except Exception:  # noqa: BLE001
            pass
        try:
            path = build_rank_image(day)
        except Exception:  # noqa: BLE001
            logger.exception("[kebiao] 排行榜出图失败")
            return
        umos: set[str] = set()
        for fname in os.listdir(DATA_DIR):
            if not fname.endswith(".json") or fname in INTERNAL_FILES:
                continue
            try:
                with open(os.path.join(DATA_DIR, fname), encoding="utf-8-sig") as f:
                    data = json.load(f)
            except Exception:  # noqa: BLE001
                continue
            umo = str(data.get("umo") or "")
            if data.get("courses") and "GroupMessage" in umo:
                umos.add(umo)
        with open(LB_STATE, "w", encoding="utf-8") as f:
            json.dump({"last": stamp}, f)
        for umo in sorted(umos):
            try:
                await self.context.send_message(
                    umo, MessageChain([Image.fromFileSystem(path)])
                )
                logger.info(f"[kebiao] 排行榜已发 {umo} · {stamp}")
            except Exception:  # noqa: BLE001
                logger.exception(f"[kebiao] 排行榜发送失败 {umo}")


    # ---------- 收文件 ----------
    @filter.event_message_type(filter.EventMessageType.ALL)
    async def on_file(self, event: AstrMessageEvent):
        """收到课表文件就解析"""
        files = [c for c in event.message_obj.message if isinstance(c, File)]
        if not files:
            return
        key = self._key(event)
        handled = False
        for comp in files:
            name = (comp.name or "").lower()
            if not name:
                continue
            if not (name.endswith(".xls") or name.endswith(".xlsx")):
                continue
            handled = True
            try:
                path = await comp.get_file()
                if not path or not os.path.exists(path):
                    yield event.plain_result("文件没拿到，再发一次试试？")
                    continue
                if os.path.getsize(path) > 5 * 1024 * 1024:
                    yield event.plain_result("文件太大了，这不像课表喵 (￣ε￣)")
                    continue
                data = parse_schedule(path)
                data["source_name"] = comp.name
                data["nickname"] = self._nick(event)
                data["umo"] = getattr(event, "unified_msg_origin", "") or ""
                data["remind_umo"] = data["umo"]
                data["qq"] = str(event.get_sender_id() or "")
                data.setdefault("reminder", True)
                # 隐私：真实姓名/班级不留在服务器上
                data.pop("student", None)
                data.pop("class", None)
                self._save(key, data)
                logger.info(f"[kebiao] 导入成功 {comp.name} -> {len(data['courses'])} 条")
                names = {c["name"] for c in data["courses"]}
                yield event.plain_result(
                    f"✅ 导入成功：{len(names)} 门课 / {len(data['courses'])} 个时间块\n"
                    f"/课表 看整学期 · /今日课表 看今天 · /下一节课 看下一节\n"
                    f"/上课提醒 管理上课提醒（默认已开启）"
                )
            except Exception as e:  # noqa: BLE001
                logger.exception("[kebiao] 解析失败")
                yield event.plain_result(f"解析失败了 (｡•́︿•̀｡)\n{type(e).__name__}: {e}")
            finally:
                try:
                    if 'path' in dir() and path and os.path.exists(path):
                        os.remove(path)  # 原始文件不留在服务器
                except OSError:
                    pass
        if handled:
            event.stop_event()
