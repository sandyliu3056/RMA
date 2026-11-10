# -*- coding: utf-8 -*-
"""
從 i18n.py（＋web_glue.py 的網頁專用字串）產生 ui_strings.js，讓瀏覽器版不必等 Python 啟動
就能切換中文／English 與顯示名詞說明。

    python make_ui_strings.py          # 重新產生 ui_strings.js
    python make_ui_strings.py --check  # 只檢查是否過期（過期時結束碼 1）

改了 i18n.py 或 web_glue.py 的 WEB_EN 之後請重新執行一次，並把 ui_strings.js 一起提交。
--check 同時確認圖表子集字型 fonts/NotoSansTC-Regular.otf 涵蓋原始碼裡的所有中文字。
"""
import os
import sys
import json

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import i18n                      # noqa: E402
from web_glue import WEB_EN, tr  # noqa: E402

OUT = os.path.join(HERE, "ui_strings.js")


def build():
    data = {"ui": {}, "glossary": {}}
    keys = set(i18n.EN) | set(i18n.ZH) | set(WEB_EN)
    for lang in ("zh", "en"):
        i18n.set_lang(lang)
        data["ui"][lang] = {k: v for k in sorted(keys) if (v := tr(k)) != k}
        seen, items = set(), []
        for key in i18n.GLOSSARY:
            name = i18n.T(key)
            if name in seen:
                continue
            seen.add(name)
            items.append([name, i18n.define(key)])
        data["glossary"][lang] = items
    i18n.set_lang("zh")
    body = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    return ("/* 由 make_ui_strings.py 從 i18n.py 產生，請勿手改；改了 i18n.py 請重新執行 python make_ui_strings.py */\n"
            "window.RMA_UI = " + body + ";\n")


FONT = os.path.join(HERE, "fonts", "NotoSansTC-Regular.otf")
SOURCES = ["i18n.py", "rma_engine.py", "web_glue.py", "零件規劃平台.py"]


def check_font():
    """圖表字型是子集字型：確認原始碼裡出現的每個非 ASCII 字都有字形（少了會變成方塊）。"""
    try:
        from fontTools.ttLib import TTFont
    except ImportError:
        print("（未安裝 fontTools，略過字型檢查）")
        return 0
    if not os.path.exists(FONT):
        print("找不到字型檔：" + FONT)
        return 1
    cmap = TTFont(FONT).getBestCmap()
    used = set()
    for f in SOURCES:
        used |= set(open(os.path.join(HERE, f), encoding="utf-8").read())
    def cjk(c):   # 中日韓文字、全形符號與標點（圖表會用到的）；其他符號（例如 ▾）只在桌面程式介面出現
        o = ord(c)
        return 0x2E80 <= o <= 0x9FFF or 0xF900 <= o <= 0xFAFF or 0xFF00 <= o <= 0xFFEF
    missing = sorted(c for c in used if cjk(c) and ord(c) not in cmap)
    if missing:
        print("圖表字型缺少字形：" + "".join(missing) + "  → 請用 pyftsubset 重新產生 fonts/NotoSansTC-Regular.otf（見 README）")
        return 1
    print(f"圖表字型涵蓋原始碼全部字元（{len(cmap):,} 個字形）")
    return 0


def main():
    text = build()
    if "--check" in sys.argv:
        cur = open(OUT, encoding="utf-8").read() if os.path.exists(OUT) else ""
        if cur != text:
            print("ui_strings.js 已過期，請執行：python make_ui_strings.py")
            return 1
        print("ui_strings.js 是最新的")
        return check_font()
    with open(OUT, "w", encoding="utf-8") as fh:
        fh.write(text)
    n = {lang: len(v) for lang, v in json.loads(text.split("= ", 1)[1].rstrip(";\n"))["ui"].items()}
    print(f"已產生 {OUT}（{os.path.getsize(OUT) // 1024} KB；字串數 {n}）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
