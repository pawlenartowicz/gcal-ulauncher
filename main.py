import re
from datetime import datetime, date, time, timedelta
from zoneinfo import ZoneInfo
from urllib.parse import urlencode, quote_plus

from ulauncher.api.client.Extension import Extension
from ulauncher.api.client.EventListener import EventListener
from ulauncher.api.shared.event import KeywordQueryEvent
from ulauncher.api.shared.item.ExtensionResultItem import ExtensionResultItem
from ulauncher.api.shared.action.RenderResultListAction import RenderResultListAction
from ulauncher.api.shared.action.OpenUrlAction import OpenUrlAction


WEEKDAYS = {
    "monday": 0, "mon": 0,
    "tuesday": 1, "tue": 1, "tues": 1,
    "wednesday": 2, "wed": 2,
    "thursday": 3, "thu": 3, "thurs": 3,
    "friday": 4, "fri": 4,
    "saturday": 5, "sat": 5,
    "sunday": 6, "sun": 6,
}

MONTHS = {
    "jan": 1, "january": 1,
    "feb": 2, "february": 2,
    "mar": 3, "march": 3,
    "apr": 4, "april": 4,
    "may": 5,
    "jun": 6, "june": 6,
    "jul": 7, "july": 7,
    "aug": 8, "august": 8,
    "sep": 9, "sept": 9, "september": 9,
    "oct": 10, "october": 10,
    "nov": 11, "november": 11,
    "dec": 12, "december": 12,
}

# Added "from" (for time ranges).
OPS = ["on", "at", "from", "with", "note", "desc", "details", "for", "in", "where"]


def _strip_ordinal(s: str) -> str:
    return re.sub(r"(\d+)(st|nd|rd|th)\b", r"\1", s, flags=re.IGNORECASE)


def parse_duration(s: str, default_minutes: int) -> int:
    if not s:
        return default_minutes
    s = s.strip().lower()

    if re.fullmatch(r"\d+", s):
        return max(1, int(s))

    total = 0
    parts = re.findall(r"(\d+)\s*(h|hr|hrs|hour|hours|m|min|mins|minute|minutes)", s)
    for num, unit in parts:
        n = int(num)
        if unit.startswith("h"):
            total += n * 60
        else:
            total += n

    return total if total > 0 else default_minutes


def next_weekday(d: date, target_wd: int, mode: str) -> date:
    wd = d.weekday()
    delta = (target_wd - wd) % 7

    if mode == "this":
        return d + timedelta(days=delta)
    if mode == "next":
        return d + timedelta(days=(delta if delta != 0 else 0) + 7)
    return d + timedelta(days=(delta if delta != 0 else 7))


def parse_date(s: str, now_local: datetime, date_order: str = "mdy") -> date | None:
    if not s:
        return None
    s0 = _strip_ordinal(s.strip().lower())
    today = now_local.date()

    if s0 == "today":
        return today
    if s0 in ("tomorrow", "tmr", "tmrw"):
        return today + timedelta(days=1)

    m = re.fullmatch(
        r"(this|next)?\s*(mon|monday|tue|tues|tuesday|wed|wednesday|thu|thurs|thursday|fri|friday|sat|saturday|sun|sunday)",
        s0
    )
    if m:
        mode = m.group(1) or "plain"
        wd = WEEKDAYS[m.group(2)]
        return next_weekday(today, wd, mode)

    m = re.fullmatch(r"(\d{1,2})[./-](\d{1,2})", s0)
    if m:
        if date_order == "dmy":
            dd = int(m.group(1))
            mm = int(m.group(2))
        else:
            mm = int(m.group(1))
            dd = int(m.group(2))
        yy = today.year
        try:
            candidate = date(yy, mm, dd)
        except ValueError:
            return None
        if candidate < today:
            try:
                candidate = date(yy + 1, mm, dd)
            except ValueError:
                return None
        return candidate

    m = re.fullmatch(r"([a-z]+)\s+(\d{1,2})", s0)
    if m and m.group(1) in MONTHS:
        mm = MONTHS[m.group(1)]
        dd = int(m.group(2))
        yy = today.year
        try:
            candidate = date(yy, mm, dd)
        except ValueError:
            return None
        if candidate < today:
            try:
                candidate = date(yy + 1, mm, dd)
            except ValueError:
                return None
        return candidate

    m = re.fullmatch(r"(\d{1,2})", s0)
    if m:
        dd = int(m.group(1))
        yy = today.year
        mm = today.month
        try:
            candidate = date(yy, mm, dd)
        except ValueError:
            return None
        if candidate < today:
            mm2 = mm + 1
            yy2 = yy
            if mm2 == 13:
                mm2 = 1
                yy2 += 1
            try:
                candidate = date(yy2, mm2, dd)
            except ValueError:
                return None
        return candidate

    return None


def _has_ampm(s: str) -> bool:
    if not s:
        return False
    s0 = s.strip().lower()
    return bool(re.search(r"\b(am|pm)\b", s0)) or bool(re.search(r"\d\s*(a|p)\b", s0))


def parse_time(s: str) -> time | None:
    if not s:
        return None
    s0 = s.strip().lower().replace(" ", "")
    s0 = s0.replace("a.m.", "am").replace("p.m.", "pm")

    ampm = None
    if s0.endswith("am"):
        ampm = "am"
        s0 = s0[:-2]
    elif s0.endswith("pm"):
        ampm = "pm"
        s0 = s0[:-2]
    elif s0.endswith("a"):
        ampm = "am"
        s0 = s0[:-1]
    elif s0.endswith("p"):
        ampm = "pm"
        s0 = s0[:-1]

    m = re.fullmatch(r"(\d{1,2}):(\d{2})", s0)
    if m:
        h = int(m.group(1))
        mi = int(m.group(2))
    else:
        m = re.fullmatch(r"(\d{1,4})", s0)
        if not m:
            return None
        digits = m.group(1)
        if len(digits) in (3, 4):
            h = int(digits[:-2])
            mi = int(digits[-2:])
        else:
            h = int(digits)
            mi = 0

    if mi < 0 or mi > 59 or h < 0 or h > 24:
        return None

    if ampm is None and h >= 13:
        return time(hour=h, minute=mi)

    if ampm:
        if h == 12:
            h = 0
        if ampm == "pm":
            h += 12
        return time(hour=h, minute=mi)

    # No am/pm heuristic: assume 1-8 => PM, 9-11 => AM
    if h == 12:
        return time(hour=12, minute=mi)
    if 1 <= h <= 8:
        return time(hour=h + 12, minute=mi)
    if 9 <= h <= 11:
        return time(hour=h, minute=mi)
    return time(hour=h, minute=mi)


def parse_time_range(s: str) -> tuple[time | None, time | None]:
    """
    Accept:
      "3 to 10"
      "11-1"
      "4p to 8p"
      "4:30p - 6p"
    """
    if not s:
        return None, None

    s0 = s.strip()
    s0 = s0.replace("—", "-").replace("–", "-")
    m = re.match(r"(.+?)\s*(?:\bto\b|\-)\s*(.+)$", s0, flags=re.IGNORECASE)
    if not m:
        return None, None

    left = m.group(1).strip()
    right = m.group(2).strip()

    t1 = parse_time(left)
    t2 = parse_time(right)
    if t1 is None or t2 is None:
        return None, None

    # If neither side specified AM/PM, and end looks earlier, push end to PM (common case: "3 to 10")
    if not _has_ampm(left) and not _has_ampm(right):
        if (t2.hour, t2.minute) <= (t1.hour, t1.minute) and t2.hour < 12:
            h = t2.hour + 12
            if h < 24:
                t2 = time(hour=h, minute=t2.minute)

    return t1, t2


def parse_aliases(text: str) -> dict[str, str]:
    out = {}
    if not text:
        return out
    for part in text.split(","):
        part = part.strip()
        if not part or "=" not in part:
            continue
        name, email = part.split("=", 1)
        name = name.strip().lower()
        email = email.strip()
        if name and email:
            out[name] = email
    return out


def split_with_tokens(s: str) -> list[str]:
    if not s:
        return []
    s = re.sub(r"\band\b", ",", s, flags=re.IGNORECASE)
    parts = [p.strip() for p in s.split(",")]
    out = []
    for p in parts:
        if not p:
            continue
        if "@" in p:
            out.append(p)
        else:
            for piece in re.split(r"\s{2,}|\s+(?=[A-Za-z])", p):
                piece = piece.strip()
                if piece:
                    out.append(piece)
    return out


def extract_emails(s: str) -> list[str]:
    if not s:
        return []
    emails = re.findall(r"[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}", s, flags=re.IGNORECASE)
    seen = set()
    out = []
    for e in emails:
        el = e.lower()
        if el not in seen:
            seen.add(el)
            out.append(e)
    return out


def guests_from_with(with_text: str, alias_map: dict[str, str]) -> tuple[list[str], list[str], list[str]]:
    """
    Returns (emails, unresolved_tokens, display_tokens)
    display_tokens preserves aliases (e.g., "Mom") when matched, else shows emails.
    """
    emails = []
    unresolved = []
    display = []

    tokens = split_with_tokens(with_text)
    for tok in tokens:
        if "@" in tok:
            for e in extract_emails(tok):
                emails.append(e)
                display.append(e)
            continue

        key = tok.strip().lower()
        if not key:
            continue
        if key in alias_map:
            emails.append(alias_map[key])
            display.append(tok.strip())  # show alias text
        else:
            unresolved.append(tok.strip())

    # de-dupe preserve order
    seen_em = set()
    emails_out = []
    for e in emails:
        el = e.lower()
        if el not in seen_em:
            seen_em.add(el)
            emails_out.append(e)

    seen_disp = set()
    display_out = []
    for d in display:
        dl = d.lower()
        if dl not in seen_disp:
            seen_disp.add(dl)
            display_out.append(d)

    return emails_out, unresolved, display_out


def extract_sections(query: str) -> dict:
    q = query.strip()

    if q.startswith(("\"", "'")):
        quote = q[0]
        end = q.find(quote, 1)
        if end != -1:
            title = q[1:end].strip()
            rest = q[end + 1:].strip()
            parts = {"title": title}
            if rest:
                parts.update(extract_sections(rest))
                parts["title"] = title
            return parts

    pattern = r"\b(" + "|".join(re.escape(o) for o in OPS) + r")\b"
    matches = list(re.finditer(pattern, q, flags=re.IGNORECASE))

    if not matches:
        return {"title": q.strip()}

    result = {}
    first = matches[0]
    result["title"] = q[:first.start()].strip()

    for i, m in enumerate(matches):
        key = m.group(1).lower()
        start = m.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(q)
        val = q[start:end].strip()

        if key in ("desc", "details"):
            key = "note"
        if key == "where":
            key = "in"
        result[key] = val

    return result


def infer_date_from_title(title: str) -> tuple[str, str]:
    t = title.strip()
    if not t:
        return t, ""

    patterns = [
        r"\b(today|tomorrow|tmr|tmrw)\b",
        r"\b(this|next)\s+(mon|monday|tue|tues|tuesday|wed|wednesday|thu|thurs|thursday|fri|friday|sat|saturday|sun|sunday)\b",
        r"\b(mon|monday|tue|tues|tuesday|wed|wednesday|thu|thurs|thursday|fri|friday|sat|saturday|sun|sunday)\b",
        r"\b(\d{1,2}[./-]\d{1,2})\b",
        r"\b(jan|january|feb|february|mar|march|apr|april|may|jun|june|jul|july|aug|august|sep|sept|september|oct|october|nov|november|dec|december)\s+\d{1,2}(st|nd|rd|th)?\b",
    ]

    for pat in patterns:
        matches = list(re.finditer(pat, t, flags=re.IGNORECASE))
        if matches:
            m = matches[-1]
            date_text = t[m.start():m.end()]
            new_title = (t[:m.start()] + t[m.end():]).strip()
            return new_title, date_text.strip()

    return t, ""


def build_url(base_url: str, params: dict) -> str:
    joiner = "&" if "?" in base_url else "?"
    return base_url + joiner + urlencode(params, quote_via=quote_plus)


def make_calendar_url(
    base_url: str,
    title: str,
    tzname: str,
    start_local: datetime | None,
    end_local: datetime | None,
    all_day_start: date | None,
    all_day_end: date | None,
    details: str,
    location: str,
    guests: list[str],
    src: str,
) -> str:
    params = {}

    if "/render" in base_url:
        params["action"] = "TEMPLATE"

    params["text"] = title
    params["ctz"] = tzname

    if src:
        params["src"] = src

    if details:
        params["details"] = details
    if location:
        params["location"] = location
    if guests:
        params["add"] = ",".join(guests)

    if start_local and end_local:
        params["dates"] = f"{start_local:%Y%m%dT%H%M%S}/{end_local:%Y%m%dT%H%M%S}"
    else:
        params["dates"] = f"{all_day_start:%Y%m%d}/{all_day_end:%Y%m%d}"

    return build_url(base_url, params)


def fmt_date(d: date, date_order: str = "mdy") -> str:
    # "Mar 4 2025" (mdy) or "4 Mar 2025" (dmy)
    if date_order == "dmy":
        return f"{d.day} {d.strftime('%b')} {d.year}"
    return f"{d.strftime('%b')} {d.day} {d.year}"


def fmt_time_short(t: time, time_display: str = "12h") -> str:
    # "4p" / "4:30p" (12h) or "16:30" (24h)
    if time_display == "24h":
        return f"{t.hour}:{t.minute:02d}"
    h24 = t.hour
    mi = t.minute
    ap = "a" if h24 < 12 else "p"
    h12 = h24 % 12
    if h12 == 0:
        h12 = 12
    if mi == 0:
        return f"{h12}{ap}"
    return f"{h12}:{mi:02d}{ap}"


def fmt_duration(minutes: int) -> str:
    if minutes <= 0:
        return "0m"
    h = minutes // 60
    m = minutes % 60
    if h and not m:
        return f"{h}h"
    if h and m:
        return f"{h}h{m}m"
    return f"{m}m"


class GCalQuickEventExtension(Extension):
    def __init__(self):
        super().__init__()
        self.subscribe(KeywordQueryEvent, KeywordQueryEventListener())


class KeywordQueryEventListener(EventListener):
    def on_event(self, event, extension):
        arg = (event.get_argument() or "").strip()
        used_kw = event.get_keyword()

        tzname = extension.preferences.get("tz", "America/Denver").strip() or "America/Denver"
        default_duration_raw = extension.preferences.get("default_duration", "60")
        try:
            default_duration = int(str(default_duration_raw).strip() or "60")
        except ValueError:
            default_duration = 60

        date_order = extension.preferences.get("date_order", "mdy").strip() or "mdy"
        time_display = extension.preferences.get("time_display", "12h").strip() or "12h"
        alias_map = parse_aliases(extension.preferences.get("guest_aliases", ""))

        kw_personal = extension.preferences.get("kw_personal", "event")
        kw_work = extension.preferences.get("kw_work", "wevent")
        kw_other = extension.preferences.get("kw_other", "oevent")

        personal_base = extension.preferences.get("personal_base_url", "https://calendar.google.com/calendar/u/0/r/eventedit").strip()
        work_base = extension.preferences.get("work_base_url", "https://calendar.google.com/calendar/u/1/r/eventedit").strip()
        other_base = extension.preferences.get("other_base_url", "https://calendar.google.com/calendar/u/2/r/eventedit").strip()

        personal_src = extension.preferences.get("personal_src", "").strip()
        work_src = extension.preferences.get("work_src", "").strip()
        other_src = extension.preferences.get("other_src", "").strip()

        if used_kw == kw_work:
            profile, base_url, src = "Work", work_base, work_src
        elif used_kw == kw_other:
            profile, base_url, src = "Other", other_base, other_src
        else:
            profile, base_url, src = "Personal", personal_base, personal_src

        tz = ZoneInfo(tzname)
        now_local = datetime.now(tz)

        if not arg:
            item = ExtensionResultItem(
                icon="images/icon.png",
                name=f"{profile} | Type to Create an Event",
                description='Try: event Family party tomorrow at 1 to 3 with Mom, Dad for 2h in Backyard',
                on_enter=OpenUrlAction("https://calendar.google.com/calendar"),
            )
            return RenderResultListAction([item])

        parts = extract_sections(arg)
        title = (parts.get("title") or "").strip() or "(No title)"

        # Allow "tomorrow/next fri/12.15" without "on"
        if not (parts.get("on") or "").strip():
            title2, inferred = infer_date_from_title(title)
            if inferred:
                parts["on"] = inferred
                title = title2 or title

        d = parse_date(parts.get("on", ""), now_local, date_order) or now_local.date()

        details = (parts.get("note", "") or "").strip()
        location = (parts.get("in", "") or "").strip()

        guests, unresolved, guest_display = guests_from_with(parts.get("with", ""), alias_map)

        # Time handling:
        # Priority:
        # 1) "from X to Y"
        # 2) "at X to Y" (range inside at)
        # 3) "at X" + duration from "for"
        start_t = None
        end_t = None

        from_text = (parts.get("from") or "").strip()
        at_text = (parts.get("at") or "").strip()

        if from_text:
            start_t, end_t = parse_time_range(from_text)

        if start_t is None and at_text:
            # "at 11 to 1"
            t1, t2 = parse_time_range(at_text)
            if t1 is not None:
                start_t, end_t = t1, t2
            else:
                start_t = parse_time(at_text)

        duration_min = parse_duration(parts.get("for", ""), default_duration)

        if start_t is None:
            # all-day
            all_day_start = d
            all_day_end = d + timedelta(days=1)
            url = make_calendar_url(
                base_url=base_url,
                title=title,
                tzname=tzname,
                start_local=None,
                end_local=None,
                all_day_start=all_day_start,
                all_day_end=all_day_end,
                details=details,
                location=location,
                guests=guests,
                src=src,
            )
            time_str = "all-day"
            dur_str = ""
        else:
            start_local = datetime(d.year, d.month, d.day, start_t.hour, start_t.minute, 0, tzinfo=tz)

            # If user typed today but time already passed, bump to tomorrow
            if d == now_local.date() and start_local < now_local:
                start_local += timedelta(days=1)
                d = start_local.date()

            if end_t is not None:
                end_local = datetime(d.year, d.month, d.day, end_t.hour, end_t.minute, 0, tzinfo=tz)
                if end_local <= start_local:
                    end_local += timedelta(days=1)  # allow overnight ranges
                duration_min = int((end_local - start_local).total_seconds() // 60)
            else:
                end_local = start_local + timedelta(minutes=duration_min)

            url = make_calendar_url(
                base_url=base_url,
                title=title,
                tzname=tzname,
                start_local=start_local,
                end_local=end_local,
                all_day_start=None,
                all_day_end=None,
                details=details,
                location=location,
                guests=guests,
                src=src,
            )

            time_str = f"{fmt_time_short(start_local.time(), time_display)}-{fmt_time_short(end_local.time(), time_display)}"
            dur_str = fmt_duration(duration_min)

        # Build the compact display line
        date_str = fmt_date(d, date_order)
        guests_count = len(guests)
        guests_list = ", ".join(guest_display) if guest_display else ""
        if guests_list:
            # keep it from exploding
            if len(guests_list) > 40:
                guests_list = guests_list[:37] + "..."

        # name (main line)
        pieces = [
            f"{profile}",
            title,
            f"🗓️: {date_str}",
            f"🕓: {time_str}",
        ]
        if dur_str:
            pieces.append(f"⏱: {dur_str}")
        if guests_count or guests_list:
            pieces.append(f"👥 {guests_count} {guests_list}".strip())

        name_line = " | ".join([p for p in pieces if p])

        # description (secondary line) – keep short
        desc_bits = []
        if location:
            desc_bits.append(f"📍 {location}")
        if unresolved:
            desc_bits.append(f"⚠ Unmatched: {', '.join(unresolved)[:40]}")
        description = " | ".join(desc_bits)[:250] if desc_bits else "Open prefilled event page"

        display_line = name_line  # the formatted line you built

        item = ExtensionResultItem(
            icon="images/icon.png",
            name=f"Create ({profile}): {title}"[:140],
            description=display_line[:250],
            on_enter=OpenUrlAction(url),
        )

        return RenderResultListAction([item])


if __name__ == "__main__":
    GCalQuickEventExtension().run()