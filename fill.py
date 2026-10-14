"""Fill series (the fill handle): continue numbers, dates, day/month names and 'Item 7' style text."""
import calendar
import datetime as dt
import re

from formatting import parse_date, parse_number

CYCLES = [list(calendar.day_name), list(calendar.day_abbr),
          list(calendar.month_name)[1:], list(calendar.month_abbr)[1:]]
TEXT_NUM = re.compile(r"^(.*?)(\d+)(\D*)$", re.S)


def _num_text(x, ints):
    if ints and float(x).is_integer():
        return str(int(x))
    return f"{x:.15g}"


def _step(nums):
    """One value -> +1; several -> the average step between first and last."""
    return 1 if len(nums) == 1 else (nums[-1] - nums[0]) / (len(nums) - 1)


def _case(word, like):
    return word.upper() if like.isupper() and len(like) > 1 else word.lower() if like.islower() else word


def extend(values, count, copy=False):
    """Continue a run of cell texts by `count` cells.
    -> [(source index, text)] where text None means "copy that source cell"
    (the caller copies it, moving formula references)."""
    n = len(values)
    cycle = [(k % n, None) for k in range(count)]
    if copy or not any(values) or any(v.startswith("=") for v in values):
        return cycle

    nums = [parse_number(v) for v in values]
    if all(x is not None for x in nums):
        step, ints = _step(nums), all("." not in v and "e" not in v.lower() for v in values)
        ints = ints and float(step).is_integer()
        return [(k % n, _num_text(nums[-1] + step * (k + 1), ints)) for k in range(count)]

    dates = [parse_date(v) for v in values]
    if all(d is not None and len(v) == 10 for d, v in zip(dates, values)):
        days = [d.toordinal() for d in dates]
        step = round(_step(days))
        return [(k % n, (dates[-1] + dt.timedelta(days=step * (k + 1))).strftime("%Y-%m-%d"))
                for k in range(count)]

    for names in CYCLES:
        lower = [x.lower() for x in names]
        idx = [lower.index(v.lower()) if v.lower() in lower else None for v in values]
        if all(i is not None for i in idx):
            step = 1 if n == 1 else round(_step(idx))
            return [(k % n, _case(names[(idx[-1] + step * (k + 1)) % len(names)], values[-1]))
                    for k in range(count)]

    parts = [TEXT_NUM.match(v) for v in values]
    if all(parts) and len({(p.group(1), p.group(3)) for p in parts}) == 1:
        pre, _, post = parts[-1].groups()
        width = len(parts[0].group(2)) if parts[0].group(2).startswith("0") else 0
        nums = [int(p.group(2)) for p in parts]
        step = round(_step(nums))
        out = []
        for k in range(count):
            v = nums[-1] + step * (k + 1)
            out.append((k % n, f"{pre}{max(v, 0):0{width}d}{post}"))
        return out

    return cycle
