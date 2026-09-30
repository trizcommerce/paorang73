"""duckduck_calendar.html → 설치형 웹앱(PWA) 폴더 + 업로드용 zip.

    python -X utf8 pwa_build.py

pwa/ 폴더(index.html, manifest, service worker, 아이콘)와 duckduck_app.zip 을 만든다.
Netlify / GitHub Pages 등 정적 호스팅에 pwa/ 폴더를 그대로 올리면 된다.
"""
import datetime as dt
import hashlib
import json
import re
import shutil
from pathlib import Path

import cal_feed

HERE = Path(__file__).parent
SRC = HERE / "duckduck_calendar.html"
DATA = HERE / "data.json"
ICONS = HERE / "icons"  # 미리 만들어 둔 아이콘 (서버에는 한글 폰트가 없어서 복사해서 씀)
PWA = HERE / "pwa"
ZIP = HERE / "duckduck_app"
CONFIG = json.loads((HERE / "config.json").read_text(encoding="utf-8"))
NAME = CONFIG["app_name"]
SHORT = CONFIG.get("short_name", NAME)
GREEN = "#2E6A27"
BG = "#F3F5F0"

HEAD = f"""<link rel="manifest" href="manifest.webmanifest">
<meta name="theme-color" content="{BG}" media="(prefers-color-scheme: light)">
<meta name="theme-color" content="#111512" media="(prefers-color-scheme: dark)">
<meta name="apple-mobile-web-app-title" content="{SHORT}">
<meta name="apple-mobile-web-app-status-bar-style" content="default">
<meta name="mobile-web-app-capable" content="yes">
<link rel="apple-touch-icon" href="icon-180.png">
<link rel="icon" type="image/png" href="icon-192.png">
<style>:root{{padding-top:env(safe-area-inset-top,0px);padding-bottom:0}}body{{margin:0}}img{{max-width:100%}}[hidden]{{display:none!important}}</style>
"""

REGISTER = """<script>
if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => navigator.serviceWorker.register("sw.js").catch(() => {}));
}
</script>
"""

SW = """// 화면은 항상 최신본을 먼저 받아오고(network-first), 오프라인일 때만 저장본을 보여준다.
const CACHE = "duckduck-__VER__";
const CORE = ["./", "index.html", "manifest.webmanifest", "icon-192.png", "icon-512.png", "icon-180.png"];
self.addEventListener("install", e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(CORE)).then(() => self.skipWaiting()));
});
self.addEventListener("activate", e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});
self.addEventListener("fetch", e => {
  const req = e.request;
  if (req.method !== "GET") return;
  e.respondWith(fetch(req).then(res => {
    if (res.ok || res.type === "opaque") {
      const copy = res.clone();
      caches.open(CACHE).then(c => c.put(req, copy));
    }
    return res;
  }).catch(() => caches.match(req).then(r => r || caches.match("index.html"))));
});
"""


def icon(size, path):
    from PIL import Image, ImageDraw, ImageFont
    s = size * 4  # 크게 그린 뒤 축소해서 가장자리를 부드럽게
    im = Image.new("RGB", (s, s), GREEN)
    d = ImageDraw.Draw(im)
    # 마스커블 아이콘 안전 영역(가운데 80%) 안에 달력 모양
    w, h = s * 0.50, s * 0.46
    x0, y0 = (s - w) / 2, (s - h) / 2 + s * 0.03
    r = s * 0.07
    d.rounded_rectangle([x0, y0, x0 + w, y0 + h], r, fill="#FFFFFF")
    d.rounded_rectangle([x0, y0, x0 + w, y0 + h * 0.26], r, fill="#CFE3C8")
    d.rectangle([x0, y0 + h * 0.15, x0 + w, y0 + h * 0.26], fill="#CFE3C8")
    for fx in (0.28, 0.72):
        cx = x0 + w * fx
        d.rounded_rectangle([cx - s * 0.018, y0 - s * 0.05, cx + s * 0.018, y0 + s * 0.05], s * 0.018, fill="#FFFFFF")
    font = ImageFont.truetype(r"C:\Windows\Fonts\malgunbd.ttf", int(s * 0.21))
    d.text((s / 2, y0 + h * 0.63), CONFIG.get("icon_char", NAME[0]), font=font, fill=GREEN, anchor="mm")
    im.resize((size, size), Image.LANCZOS).save(path)


def main():
    if PWA.exists():
        shutil.rmtree(PWA)
    PWA.mkdir()
    html = SRC.read_text(encoding="utf-8")
    html = html.replace("<title>", HEAD + "<title>", 1) + REGISTER
    (PWA / "index.html").write_text(html, encoding="utf-8")
    # 앱 데이터가 가리키는 시트 이미지만 복사
    used = set(re.findall(r"media/([0-9a-f]{16}\.jpg)", DATA.read_text(encoding="utf-8")))
    if used:
        (PWA / "media").mkdir()
        for name in used:
            shutil.copy(HERE / "media" / name, PWA / "media" / name)
    for n in (180, 192, 512):
        saved = ICONS / f"icon-{n}.png"
        if saved.exists():
            shutil.copy(saved, PWA / saved.name)
        else:
            icon(n, PWA / f"icon-{n}.png")
    data = json.loads(DATA.read_text(encoding="utf-8"))
    # 휴대폰 캘린더 구독용 파일 (당일 오전 9시 알림)
    (PWA / "calendar.ics").write_text(cal_feed.build_ics(data, CONFIG, cal_feed.site_url()), encoding="utf-8", newline="")
    # 시트 내용이 바뀌었는지 비교하는 용도 (업데이트 시각은 제외하고 해시)
    data.pop("updated", None)
    # 앱 코드가 바뀌어도 새로 배포되도록 화면·캘린더 코드도 함께 해시
    code = "".join((HERE / f).read_text(encoding="utf-8") for f in ("template.html", "cal_feed.py", "pwa_build.py")
                   if (HERE / f).exists())
    digest = hashlib.sha256((json.dumps(data, ensure_ascii=False, sort_keys=True) + code).encode()).hexdigest()[:16]
    (PWA / "version.txt").write_text(digest, encoding="utf-8")
    manifest = {
        "name": NAME, "short_name": SHORT, "lang": "ko",
        "start_url": "./", "scope": "./", "display": "standalone",
        "background_color": BG, "theme_color": BG,
        "icons": [
            {"src": "icon-192.png", "sizes": "192x192", "type": "image/png", "purpose": "any maskable"},
            {"src": "icon-512.png", "sizes": "512x512", "type": "image/png", "purpose": "any maskable"},
        ],
    }
    (PWA / "manifest.webmanifest").write_text(json.dumps(manifest, ensure_ascii=False, indent=1), encoding="utf-8")
    (PWA / "sw.js").write_text(SW.replace("__VER__", dt.datetime.now().strftime("%Y%m%d%H%M")), encoding="utf-8")
    if "--zip" in __import__("sys").argv:
        shutil.make_archive(str(ZIP), "zip", PWA)
    print("built", PWA, "version", digest)


if __name__ == "__main__":
    main()
