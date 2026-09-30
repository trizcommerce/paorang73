"""앱 데이터(data.json) → 휴대폰 캘린더 구독용 calendar.ics

- 날짜마다 일정 하나: 그날 올릴 콘텐츠를 묶어서 제목·설명에 적고, 당일 오전 9시에 알림
- 공구 기간은 여러 날짜 일정 (알림 없음)
- 캘린더 탭의 일정(미팅·여행 등)은 알림 없는 일정으로
아이폰 캘린더는 이 파일을 구독하면 정해진 주기로 다시 받아 자동으로 바뀐다.
"""
import datetime as dt
import os
import re
from collections import defaultdict

ALARM_AT = "PT9H"  # 하루 일정 시작(0시)부터 9시간 뒤 = 오전 9시
FMT = [("광고", "광고"), ("릴스", "릴스"), ("캐러셀|게시글|게시물|피드", "피드"), ("무물", "무물"),
       ("라이브|라방", "라방"), ("자율|자유|일상", "자유"), ("스토리", "스토리")]


def plain(s):
    s = re.sub(r"⟪[^⟫]*⟫", "", s or "")
    return re.sub(r"⟦([^|⟧]*)\|[^⟧]*⟧", r"\1", s).strip()


def fmt_of(s):
    return next((label for pat, label in FMT if re.search(pat, s or "")), "")


def norm(s):
    return re.sub(r"[\s\W_]+", "", plain(s))


def _esc(s):
    return s.replace("\\", "\\\\").replace(";", "\\;").replace(",", "\\,").replace("\n", "\\n")


def _fold(line):
    """iCalendar 줄 길이 규칙 (75바이트마다 줄바꿈 + 공백)"""
    out, cur = [], b""
    for ch in line:
        b = ch.encode("utf-8")
        if len(cur) + len(b) > 74:
            out.append(cur.decode("utf-8"))
            cur = b" " + b
        else:
            cur += b
    out.append(cur.decode("utf-8"))
    return "\r\n".join(out)


def _day(d):
    return d.replace("-", "")


def _next(d):
    return (dt.date.fromisoformat(d) + dt.timedelta(days=1)).strftime("%Y%m%d")


def site_url():
    repo = os.environ.get("GITHUB_REPOSITORY", "")  # GitHub Actions 에서 "trizcommerce/imxux9"
    if "/" in repo:
        owner, name = repo.split("/", 1)
        return f"https://{owner.lower()}.github.io/{name}/"
    return ""


def build_ics(data, config, url=""):
    seller = config.get("seller", "")
    items = defaultdict(list)  # 날짜 → [(스토리 여부, 한 줄)]

    for p in data.get("products", []):
        pname = f"{p['name']} {p.get('round', '')}".strip()
        for s in p.get("slots", []):
            if not s.get("date"):
                continue
            kind = fmt_of(s.get("format", ""))
            title = plain(s.get("title", ""))
            if not title:
                first = next((v for it in s.get("items", []) for k, v in it.items() if not k.startswith("__")), "")
                title = plain(first).split("\n")[0][:40]
            head = f"[{pname}{' ' + s['label'] if s.get('label') else ''}]"
            if kind and (kind in title or norm(kind) in norm(title)):  # "자유 · 자유일상" 같은 반복 방지
                kind = "" if kind != "스토리" else kind
            items[s["date"]].append((kind == "스토리", " ".join(x for x in (head, kind, "·" if title and kind else "", title) if x)))

    for m in data.get("months", []):  # 월별 그리드(날짜별 기획)
        for d in m.get("days", []) if m.get("kind") == "grid" else []:
            for f in d.get("fields", []):
                if f["label"] in ("주제", "콘텐츠"):
                    t = plain(f["text"]).split("\n")[0]
                    items[d["date"]].append((False, f"[{m['name']} 기획] {t}"))
                elif f["label"] == "스토리":
                    items[d["date"]].append((True, f"[{m['name']} 기획] 스토리"))

    extra = []  # 캘린더 탭 일정
    for e in data.get("events", []):
        text = plain(e.get("text", "")).replace("\n", " ")
        if not text or text == "스토리":
            continue
        multi = e.get("end") and e["end"] != e["date"]
        if e.get("kind") != "content" or multi:
            extra.append((e["date"], e.get("end") or e["date"], text))
        elif not any(key and key in norm(line) for _, line in items[e["date"]]
                     for key in [norm(re.sub(r"커머스|D\s*[-+]\s*\d+|OPEN|오픈|마감|릴스|캐러셀|게시글|피드|스토리|광고|\[[^\]]*\]", "", text))[:6]]):
            items[e["date"]].append((False, text))

    now = dt.datetime.now(dt.timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    uid_base = re.sub(r"\W", "", config.get("sheet_id", "cal"))[:12]
    lines = [
        "BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//TRIZ//seller-content-calendar//KO", "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH", f"X-WR-CALNAME:{_esc(config.get('app_name', '콘텐츠 캘린더'))}",
        "X-WR-TIMEZONE:Asia/Seoul", "REFRESH-INTERVAL;VALUE=DURATION:PT1H", "X-PUBLISHED-TTL:PT1H",
    ]

    def event(uid, start, end, summary, desc="", alarm=False):
        lines.extend(["BEGIN:VEVENT", f"UID:{uid}@{uid_base}.trizcommerce", f"DTSTAMP:{now}",
                      f"DTSTART;VALUE=DATE:{_day(start)}", f"DTEND;VALUE=DATE:{_next(end)}",
                      f"SUMMARY:{_esc(summary)}", "TRANSP:TRANSPARENT"])
        if desc:
            lines.append(f"DESCRIPTION:{_esc(desc)}")
        if url:
            lines.append(f"URL:{url}")
        if alarm:
            lines.extend(["BEGIN:VALARM", "ACTION:DISPLAY", f"DESCRIPTION:{_esc(summary)}",
                          f"TRIGGER;RELATED=START:{ALARM_AT}", "END:VALARM"])
        lines.append("END:VEVENT")

    for date in sorted(items):
        rows = items[date]
        main = [t for story, t in rows if not story]
        stories = sum(1 for story, _ in rows if story)
        if main:
            summary = main[0] + (f" 외 {len(main) - 1}" if len(main) > 1 else "")
            summary += f" · 스토리 {stories}" if stories else ""
        else:
            summary = f"{seller} 스토리 {stories}개".strip()
        desc = "\n".join(f"• {t}" for _, t in rows) + (f"\n\n앱에서 보기: {url}" if url else "")
        event(f"day-{date}", date, date, summary, desc, alarm=True)

    for p in data.get("products", []):
        if p.get("period"):
            a, b = p["period"]
            pname = f"{p['name']} {p.get('round', '')}".strip()
            event(f"period-{p['id']}", a, b, f"🛒 {pname} 공구", f"공구 기간 {a} ~ {b}")

    for i, (a, b, text) in enumerate(extra):
        event(f"note-{a}-{norm(text)[:12] or i}", a, b, text)

    lines.append("END:VCALENDAR")
    return "\r\n".join(_fold(l) for l in lines) + "\r\n"
