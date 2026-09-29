"""시트 안의 링크와 이미지를 꺼내는 도구.

- 링크: xlsx 로 내보내면 한 셀에 링크가 여러 개일 때 하나만 남아서, 시트의 HTML 보기(htmlview)에서
  셀별 (링크 문구, 주소) 목록을 읽는다. 셀 글자 안에 ⟦문구|주소⟧ 표시로 끼워 넣으면 앱이 링크로 그린다.
- 이미지: xlsx 안의 그림(셀에 걸쳐 둔 이미지)을 꺼내 모바일용으로 줄여 media/ 에 저장하고,
  시트별 {(행, 열): [파일 경로]} 로 돌려준다.
"""
import hashlib
import html
import io
import posixpath
import re
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
import zipfile
from html.parser import HTMLParser

MAX_SIDE = 1080  # 이미지 긴 변 최대 픽셀
NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}
LINK_OPEN, LINK_SEP, LINK_CLOSE = "⟦", "|", "⟧"


# ---------------------------------------------------------------- 링크
def _get(url):
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=120) as r:
        return r.read().decode("utf-8", "replace")


def _unwrap(url):
    url = html.unescape(url)
    if url.startswith("https://www.google.com/url?"):
        q = urllib.parse.parse_qs(urllib.parse.urlparse(url).query).get("q")
        if q:
            return q[0]
    return url


class _SheetTable(HTMLParser):
    """htmlview 표 → {(행, 열): [(문구, 주소)]} (행·열은 1부터, 병합 칸 고려)"""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.links, self.row, self.col = {}, None, 0
        self.busy = {}  # 위에서 rowspan 으로 내려온 칸: row → set(col)
        self.in_th = self.in_td = False
        self.cell = None
        self.anchor = None

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag == "th" and (a.get("id") or "").rsplit("R", 1)[-1].isdigit() and "R" in (a.get("id") or ""):
            self.row = int(a["id"].rsplit("R", 1)[-1]) + 1
            self.col = 0
        elif tag == "td" and self.row is not None:
            if "freezebar" in (a.get("class") or ""):
                return
            self.col += 1
            while self.col in self.busy.get(self.row, ()):
                self.col += 1
            span_c, span_r = int(a.get("colspan", 1)), int(a.get("rowspan", 1))
            for dr in range(span_r):
                for dc in range(span_c):
                    if dr or dc:
                        self.busy.setdefault(self.row + dr, set()).add(self.col + dc)
            self.cell = (self.row, self.col)
            self.col += span_c - 1
            self.in_td = True
        elif tag == "a" and self.in_td and a.get("href"):
            self.anchor = [_unwrap(a["href"]), ""]
        elif tag == "br" and self.anchor is not None:
            self.anchor[1] += "\n"

    def handle_endtag(self, tag):
        if tag == "a" and self.anchor is not None:
            url, label = self.anchor
            self.links.setdefault(self.cell, []).append((label.strip(), url))
            self.anchor = None
        elif tag == "td":
            self.in_td = False
        elif tag == "tr":
            self.row = None

    def handle_data(self, data):
        if self.anchor is not None:
            self.anchor[1] += data


def list_tabs(sheet_id):
    """시트의 보이는 탭 [(이름, gid)] (시트에 보이는 순서)"""
    page = _get(f"https://docs.google.com/spreadsheets/d/{sheet_id}/htmlview")
    items = re.findall(r'items\.push\(\{name: "((?:[^"\\]|\\.)*)", pageUrl: "[^"]*", gid: "(\d+)"', page)
    return [(raw.encode().decode("unicode_escape").encode("latin-1").decode("utf-8").replace("\\/", "/"), gid)
            for raw, gid in items]


def fetch_links(sheet_id):
    """{시트 이름: {(행, 열): [(문구, 주소)]}} — 실패하면 빈 dict"""
    base = f"https://docs.google.com/spreadsheets/d/{sheet_id}"
    out = {}
    for raw_name, gid in list_tabs(sheet_id):
        name = raw_name.strip().strip("[]").strip()
        t = _SheetTable()
        t.feed(_get(f"{base}/htmlview/sheet?headers=true&gid={gid}"))
        out[name] = t.links
    return out


# ---------------------------------------------------------------- 글자 서식
# 앱으로 넘기는 글자에는 두 가지 표시를 끼워 넣는다.
#   ⟪b u c=ff0000 z=+1⟫ … ⟪/⟫  : 굵게(b) 기울임(i) 밑줄(u) 취소선(s) 글자색(c) 크기(z: -1 작게, +1 크게, +2 더 크게)
#   ⟦문구|주소⟧                 : 링크
# 시트 구조를 읽을 때는 표시 없는 글자를 쓰고, 화면에 보일 칸만 표시 있는 글자로 바꾼다.
FMT_OPEN, FMT_CLOSE_TAG = "⟪", "⟪/⟫"
PLAIN_COLORS = {"000000", "1F1F1F"}  # 기본 검정은 앱 기본 글자색으로
THEME = []  # 워크북 테마 색 (collect_formats 에서 채움)


def _theme_palette(wb):
    """테마 색 목록 — 엑셀 테마 번호 순서 (0,1 과 2,3 은 밝은/어두운 색이 바뀌어 있음)"""
    raw = getattr(wb, "loaded_theme", None)
    if not raw:
        return []
    raw = raw.decode("utf-8", "replace") if isinstance(raw, bytes) else raw
    names = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"]
    found = dict(re.findall(r"<a:(\w+)>\s*<a:(?:srgbClr val|sysClr [^>]*lastClr)=\"([0-9A-Fa-f]{6})\"", raw))
    order = ["lt1", "dk1", "lt2", "dk2"] + names[4:]
    return [found.get(n, "000000").upper() for n in order]


def _tint(rgb, tint):
    if not tint:
        return rgb
    ch = [int(rgb[i:i + 2], 16) for i in (0, 2, 4)]
    ch = [round(c + (255 - c) * tint) if tint > 0 else round(c * (1 + tint)) for c in ch]
    return "".join(f"{max(0, min(255, c)):02X}" for c in ch)


def _color(font):
    col = getattr(font, "color", None) if font is not None else None
    if col is None:
        return None
    rgb = getattr(col, "rgb", None)
    if getattr(col, "type", None) == "theme" and isinstance(getattr(col, "theme", None), int) and col.theme < len(THEME):
        rgb = _tint(THEME[col.theme], getattr(col, "tint", 0) or 0)
    if not isinstance(rgb, str) or len(rgb) < 6:
        return None
    rgb = rgb[-6:].upper()
    r, g, b = (int(rgb[i:i + 2], 16) for i in (0, 2, 4))
    if rgb in PLAIN_COLORS or (0.299 * r + 0.587 * g + 0.114 * b) > 225:
        return None  # 검정·흰색(칸 배경색 위 글자)은 앱 기본 글자색으로
    return rgb


def _style(font, cell_font):
    """글자 조각의 서식 (조각에 없는 속성은 칸 기본 서식을 따름)"""
    f = font if font is not None else cell_font

    def get(k):
        v = getattr(f, k, None)
        return v if v is not None else getattr(cell_font, k, None)

    has_color = getattr(f, "color", None) is not None
    size = get("sz")
    return {
        "b": bool(get("b")),
        "i": bool(get("i")),
        "u": bool(get("u")),
        "s": bool(get("strike")),
        "c": _color(f if has_color else cell_font),
        "sz": float(size) if size else None,
    }


def cell_runs(cell):
    """셀 → (글자, [(시작, 끝, 서식)])"""
    from openpyxl.cell.rich_text import CellRichText, TextBlock
    v = cell.value
    parts = []
    if isinstance(v, CellRichText):
        for part in v:
            if isinstance(part, TextBlock):
                parts.append((part.text, _style(part.font, cell.font)))
            else:
                parts.append((str(part), _style(None, cell.font)))
    elif isinstance(v, str):
        parts.append((v, _style(None, cell.font)))
    text, runs, pos = "", [], 0
    for t, st in parts:
        runs.append((pos, pos + len(t), st))
        text += t
        pos += len(t)
    # 크기는 칸 안에서 가장 많이 쓰인 크기를 기준으로 상대 크기만 남김
    weight = {}
    for a, b, st in runs:
        if st["sz"] and text[a:b].strip():
            weight[st["sz"]] = weight.get(st["sz"], 0) + (b - a)
    base = max(weight, key=weight.get) if weight else None
    for _, _, st in runs:
        d = (st.pop("sz") or base or 0) - (base or 0)
        st["z"] = "+2" if d >= 5 else "+1" if d >= 1.5 else "-1" if d <= -1.5 else ""
    return text, runs


def _token(st):
    bits = [k for k in ("b", "i", "u", "s") if st.get(k)]
    if st.get("c"):
        bits.append("c=" + st["c"])
    if st.get("z"):
        bits.append("z=" + st["z"])
    return " ".join(bits)


def _link_spans(text, links):
    spans, tail, pos = [], [], 0
    for label, url in links or []:
        key = label.split("\n")[0].strip()
        i = text.find(key, pos) if key else -1
        if i < 0:
            tail.append((key or "링크", url))
        else:
            spans.append((i, i + len(key), url))
            pos = i + len(key)
    return spans, tail


class CellFormat:
    """한 칸의 글자·서식·링크. render() 로 표시가 들어간 글자를 만든다."""

    def __init__(self, text, runs, links):
        self.text, self.runs = text, runs
        self.spans, self.tail = _link_spans(text, links)

    def render(self, start=0, end=None, with_tail=True):
        end = len(self.text) if end is None else end
        cuts = {start, end}
        for a, b, _ in self.runs:
            cuts.update(x for x in (a, b) if start < x < end)
        for a, b, _ in self.spans:
            cuts.update(x for x in (a, b) if start < x < end)
        cuts = sorted(cuts)
        out, open_tok = [], ""
        for a, b in zip(cuts, cuts[1:]):
            piece = self.text[a:b]
            if not piece:
                continue
            st = next((st for ra, rb, st in self.runs if ra <= a < rb), {})
            tok = _token(st) if piece.strip() or st.get("u") or st.get("s") else open_tok
            if tok != open_tok:
                if open_tok:
                    out.append(FMT_CLOSE_TAG)
                if tok:
                    out.append(f"{FMT_OPEN}{tok}⟫")
                open_tok = tok
            url = next((u for sa, sb, u in self.spans if sa <= a < sb), None)
            out.append(f"{LINK_OPEN}{piece}{LINK_SEP}{url}{LINK_CLOSE}" if url else piece)
        if open_tok:
            out.append(FMT_CLOSE_TAG)
        s = "".join(out)
        if with_tail and self.tail:
            s += "\n" + " ".join(f"{LINK_OPEN}{k}{LINK_SEP}{u}{LINK_CLOSE}" for k, u in self.tail)
        return s

    def render_text(self, sub):
        """칸 글자 중 sub 부분만 (없으면 sub 그대로)"""
        i = self.text.find(sub) if sub else -1
        return self.render(i, i + len(sub), with_tail=False) if i >= 0 else sub


def collect_formats(wb, links_by_sheet):
    """{(시트, 행, 열): CellFormat} 를 만들고, 셀 값은 표시 없는 글자로 되돌린다"""
    from openpyxl.cell.rich_text import CellRichText
    THEME[:] = _theme_palette(wb)
    out = {}
    for ws in wb.worksheets:
        name = ws.title.strip()
        links = links_by_sheet.get(name.strip("[]").strip())
        if links is None:  # htmlview 를 못 읽었으면 xlsx 의 셀 링크
            links = {(c.row, c.column): [(str(c.value or "").strip(), c.hyperlink.target)]
                     for row in ws.iter_rows() for c in row
                     if c.hyperlink is not None and c.hyperlink.target}
        for row in ws.iter_rows():
            for c in row:
                if not isinstance(c.value, (str, CellRichText)):
                    continue
                text, runs = cell_runs(c)
                cl = links.get((c.row, c.column))
                if cl or any(_token(st) for _, _, st in runs):
                    out[(name, c.row, c.column)] = CellFormat(text, runs, cl)
                if isinstance(c.value, CellRichText):
                    c.value = text
    return out


# ---------------------------------------------------------------- 이미지
def _rels(z, path):
    d, f = posixpath.split(path)
    rp = posixpath.join(d, "_rels", f + ".rels")
    if rp not in z.namelist():
        return {}
    out = {}
    for r in ET.fromstring(z.read(rp)).findall("rel:Relationship", NS):
        t = r.get("Target")
        out[r.get("Id")] = t if r.get("TargetMode") == "External" else posixpath.normpath(posixpath.join(d, t))
    return out


def _save_small(data, out_dir):
    from PIL import Image
    name = hashlib.sha1(data).hexdigest()[:16] + ".jpg"
    dest = out_dir / name
    if not dest.exists():
        im = Image.open(io.BytesIO(data))
        im.thumbnail((MAX_SIDE, MAX_SIDE))
        if im.mode in ("RGBA", "LA", "P"):
            im = im.convert("RGBA")
            bg = Image.new("RGB", im.size, "white")
            bg.paste(im, mask=im.split()[-1])
            im = bg
        im.convert("RGB").save(dest, "JPEG", quality=80, optimize=True, progressive=True)
    return name


def extract_images(xlsx_path, out_dir, only=None):
    """{시트 이름: {(행, 열): [파일 이름]}} — 이미지를 out_dir 에 jpg 로 저장 (only: 이 시트들만)"""
    out_dir.mkdir(exist_ok=True)
    z = zipfile.ZipFile(xlsx_path)
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    wb_rels = _rels(z, "xl/workbook.xml")
    saved, result = {}, {}
    for s in wb.find("m:sheets", NS):
        if only is not None and s.get("name").strip() not in only:
            continue
        path = wb_rels[s.get(f"{{{NS['r']}}}id")]
        s_rels = _rels(z, path)
        cells = {}
        for dr in ET.fromstring(z.read(path)).findall("m:drawing", NS):
            d_path = s_rels[dr.get(f"{{{NS['r']}}}id")]
            d_rels = _rels(z, d_path)
            anchors = []
            for anc in ET.fromstring(z.read(d_path)):
                fr, blip = anc.find("xdr:from", NS), anc.find(".//a:blip", NS)
                if fr is None or blip is None:
                    continue
                media = d_rels.get(blip.get(f"{{{NS['r']}}}embed"))
                if not media or media not in z.namelist():
                    continue
                row = int(fr.find("xdr:row", NS).text) + 1
                col = int(fr.find("xdr:col", NS).text) + 1
                off = int(fr.find("xdr:rowOff", NS).text or 0)
                anchors.append((row, col, off, media))
            for row, col, off, media in sorted(anchors):  # 같은 칸 안에서는 위에서 아래 순서
                if media not in saved:
                    try:
                        saved[media] = _save_small(z.read(media), out_dir)
                    except Exception as e:  # 깨진 이미지는 건너뜀
                        print("skip image", media, e)
                        saved[media] = None
                if saved[media]:
                    cells.setdefault((row, col), []).append(saved[media])
        result[s.get("name").strip().strip("[]").strip()] = cells
    return result
