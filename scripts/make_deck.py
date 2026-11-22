#!/usr/bin/env python3
"""把五年收入推演的结果 JSON 画成 8 页可编辑 PPT。

用法：python3 make_deck.py plan.json 五年收入推演.pptx
      python3 make_deck.py --check          # 用 assets/deck-example.json 自检
依赖：pip install python-pptx
"""
import json
import math
import os
import platform
import re
import sys
import tempfile

try:
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.enum.dml import MSO_THEME_COLOR as T
    from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
    from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT
    from pptx.oxml.ns import qn
    from pptx.oxml.xmlchemy import OxmlElement
    from pptx.util import Inches, Pt
except ImportError:
    sys.exit("缺少 python-pptx。先运行：python3 -m pip install python-pptx")

COLORS = dict(dk1="1C2422", lt1="FFFFFF", dk2="17352F", lt2="EEF3F1", accent1="D9662B",
              accent2="2F7D67", accent3="B83A3A", accent4="E3A52B", accent5="5F6E6A", accent6="C9DDD6")
INK, WHITE, DARK, PALE = T.DARK_1, T.LIGHT_1, T.DARK_2, T.LIGHT_2
NOW, GO, RISK, WARN, MUTED, MINT = T.ACCENT_1, T.ACCENT_2, T.ACCENT_3, T.ACCENT_4, T.ACCENT_5, T.ACCENT_6
BACK = {"easy": ("容易回头", GO), "costly": ("回头有代价", WARN), "hard": ("很难回头", RISK)}
QUESTIONS = (("want", "都能成，想要哪条"), ("regret", "没试会最后悔哪条"), ("accept", "错了也认哪条"))
WRONG = (("what", "怎样算选错"), ("when", "最早什么时候看出来"), ("loss", "会亏什么"), ("back", "退回哪里"))
DETAIL = (("earn", "钱靠什么挣"), ("cost", "要付的代价"), ("quit", "最可能在这时放弃"))
COUNTS = {"periods": 5, "routes": 3, "backcast": 7}
LIMITS = {"date": 10, "period": 10, "unit": 4, "year.goal": 16, "route.day": 70, "route.detail": 34, "route.key": 1, "direction.title": 11, "direction.detail": 32, "direction.reason": 48, "direction.confident": 80,
          "chain.label": 8, "chain.text": 10, "route.name": 12, "wrong": 20,
          "backcast.date": 10, "backcast.goal": 24, "backcast.gate": 8, "today.duration": 6,
          "today.title": 20, "today.message": 140, "today.done": 28,
          "week.date": 8, "week.action": 26, "week.pass": 30, "premise.title": 20, "premise.text": 85}


def tidy(value):
    """返回一份新数据：文字里的换行、控制字符换成空格，null 当空串。"""
    if isinstance(value, dict):
        return {k: tidy(v) for k, v in value.items()}
    if isinstance(value, list):
        return [tidy(v) for v in value]
    if isinstance(value, str):
        return re.sub(r"[\s\x00-\x1f]+", " ", value).strip()
    return "" if value is None else value


def text_fields(d):
    """列出所有上屏文字和它的上限种类，供校验用。"""
    yield "date", d["date"], "date"
    yield "unit", d["unit"], "unit"
    for i, label in enumerate(d["periods"]):
        yield f"periods[{i}]", label, "period"
    for key in ("title", "detail", "reason", "confident"):
        yield f"direction.{key}", d["direction"][key], f"direction.{key}"
    for i, c in enumerate(d["direction"]["chain"]):
        for key in ("label", "text"):
            yield f"direction.chain[{i}].{key}", c[key], f"chain.{key}"
    for i, r in enumerate(d["routes"]):
        yield f"routes[{i}].key", r["key"], "route.key"
        yield f"routes[{i}].name", r["name"], "route.name"
        for j, year in enumerate(r["years"]):
            yield f"routes[{i}].years[{j}].goal", year["goal"], "year.goal"
        yield f"routes[{i}].day", r["day"], "route.day"
        for key, _ in DETAIL:
            yield f"routes[{i}].{key}", r[key], "route.detail"
        for key in [k for k, _ in WRONG] + ["stop", "miss"]:
            yield f"routes[{i}].wrong.{key}", r["wrong"][key], "wrong"
    for i, b in enumerate(d["backcast"]):
        for key in ("date", "goal", "gate"):
            yield f"backcast[{i}].{key}", b.get(key, ""), f"backcast.{key}"
    for key in ("duration", "title", "message", "done"):
        yield f"today.{key}", d["today"][key], f"today.{key}"
    for group, keys, kind in (("week", ("date", "action", "pass"), "week"), ("premises", ("title", "text"), "premise")):
        for i, row in enumerate(d[group]):
            for key in keys:
                yield f"{group}[{i}].{key}", row[key], f"{kind}.{key}"


def validate(d):
    """返回错误列表；空列表表示可以出图。"""
    try:
        bad = [f"{path} 必须是文字" for path, value, _ in text_fields(d) if not isinstance(value, str)]
        bad += [f"choice.{k} 必须是列表" for k, _ in QUESTIONS if not isinstance(d["choice"][k], list)]
        if bad:
            return bad
        errors = [f"{k} 要正好 {n} 项" for k, n in COUNTS.items() if len(d[k]) != n]
        errors += [f"路径 {r['key']} 的 years 要正好 5 项（第 1 到第 5 年）" for r in d["routes"] if len(r["years"]) != 5]
        amounts = [d["now_income"]] + [y[k] for r in d["routes"] for y in r["years"] for k in ("low", "high")]
        if not all(isinstance(v, (int, float)) and not isinstance(v, bool) and math.isfinite(v) for v in amounts):
            errors.append("now_income 和每年的 low / high 必须是数字（单位写在 unit 里，亏损写负数）")
        else:
            errors += [f"路径 {r['key']} 第 {j + 1} 年的保守数（low）不能比乐观数（high）大"
                       for r in d["routes"] for j, y in enumerate(r["years"]) if y["low"] > y["high"]]
        errors += [f"路径 {r['key']} 的 wrong.level 只能是 {'/'.join(BACK)}"
                   for r in d["routes"] if r["wrong"]["level"] not in BACK]
        keys = {r["key"] for r in d["routes"]}
        if len(keys) != 3:
            errors.append("三条路的 key 不能重复")
        if d["direction"]["key"] not in keys:
            errors.append("direction.key 必须是三条路之一")
        errors += [f"choice.{k} 必须来自用户的回答，不能为空"
                   for k, _ in QUESTIONS if not d["choice"][k] or not set(d["choice"][k]) <= keys]
        if len(d["direction"]["chain"]) != 3:
            errors.append("direction.chain 要正好 3 项")
        if not 1 <= len(d["week"]) <= 4:
            errors.append("week 要 1–4 项（不含今天）")
        if not 1 <= len(d["premises"]) <= 4:
            errors.append("premises 要 1–4 项")
        errors += [f"premises 的 url 必须以 http 开头：{p['title']}"
                   for p in d["premises"] if p.get("url") and not re.match(r"https?://", p["url"])]
        errors += [f"{path} 不能为空" for path, value, kind in text_fields(d) if not value.strip() and kind != "backcast.gate"]
        errors += [f"{path} 超过 {LIMITS[kind]} 字（现在 {len(value)}），请删字"
                   for path, value, kind in text_fields(d) if len(value) > LIMITS[kind]]
        return errors
    except KeyError as e:
        return [f"缺少字段 {e.args[0]}。照着 assets/deck-example.json 写"]
    except (TypeError, AttributeError):
        return ["JSON 结构不对：有的字段该是列表或对象，却写成了别的。照着 assets/deck-example.json 写"]


def set_theme(prs):
    font = {"Darwin": "冬青黑体简体中文", "Windows": "Microsoft YaHei"}.get(platform.system(), "Noto Sans CJK SC")
    part = prs.slide_master.part.part_related_by(RT.THEME)
    xml = part.blob.decode("utf-8")
    for name, value in COLORS.items():
        xml = re.sub(rf"<a:{name}>.*?</a:{name}>", f'<a:{name}><a:srgbClr val="{value}"/></a:{name}>', xml, flags=re.S)
    for tag in ("majorFont", "minorFont"):
        block = re.search(rf"<a:{tag}>.*?</a:{tag}>", xml, re.S).group(0)
        new = re.sub(r'<a:latin typeface="[^"]*"', '<a:latin typeface="Arial"', block)
        new = re.sub(r'<a:ea typeface="[^"]*"', f'<a:ea typeface="{font}"', new)
        new = re.sub(r'(<a:font script="Hans" typeface=")[^"]*', rf"\g<1>{font}", new)
        xml = xml.replace(block, new)
    part._blob = xml.encode("utf-8")


def write(frame, paras, align=PP_ALIGN.LEFT):
    for i, (text, size, color, bold) in enumerate(paras):
        p = frame.paragraphs[0] if i == 0 else frame.add_paragraph()
        p.alignment = align
        p.space_before = Pt(3 if i else 0)
        run = p.add_run()
        run.text = text
        run.font.size, run.font.bold = Pt(size), bold
        run.font.color.theme_color = color
        rpr = run._r.get_or_add_rPr()
        rpr.set("lang", "zh-CN")  # 不标中文，LibreOffice / WPS 会拿日文字体，简体缺字
        ea = OxmlElement("a:ea")
        ea.set("typeface", "+mj-ea" if bold else "+mn-ea")
        rpr.append(ea)


def box(slide, rect, paras=(), fill=None, line=None, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
        align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE, pad=0.14):
    pos = [Inches(v) for v in rect]
    if fill is None and line is None:
        s = slide.shapes.add_textbox(*pos)
    else:
        s = slide.shapes.add_shape(shape, *pos)
        style = s._element.find(qn("p:style"))
        if style is not None:
            s._element.remove(style)  # 默认样式带阴影
        if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
            s.adjustments[0] = 0.07
        if fill is None:
            s.fill.background()
        else:
            s.fill.solid()
            s.fill.fore_color.theme_color = fill
        if line is None:
            s.line.fill.background()
        else:
            s.line.color.theme_color, s.line.width = line, Pt(1.5)
    frame = s.text_frame
    frame.word_wrap, frame.auto_size, frame.vertical_anchor = True, None, anchor
    frame.margin_left = frame.margin_right = Inches(pad)
    frame.margin_top = frame.margin_bottom = Inches(0.03)
    write(frame, paras, align)
    return s


def arrow(slide, x1, x2, y, color=MUTED):
    line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, *[Inches(v) for v in (x1, y, x2, y)])
    line.line.color.theme_color, line.line.width = color, Pt(2)
    head = OxmlElement("a:tailEnd")
    head.set("type", "triangle")
    line.line._get_or_add_ln().append(head)


def new_slide(prs, title, dark=False, rect=(0.55, 0.35, 12.2, 0.9), size=34):
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    bg = slide.background.fill
    bg.solid()
    bg.fore_color.theme_color = DARK if dark else WHITE
    t = slide.shapes.title
    t.left, t.top, t.width, t.height = [Inches(v) for v in rect]
    frame = t.text_frame
    frame.word_wrap, frame.vertical_anchor = True, MSO_ANCHOR.MIDDLE
    frame.margin_left = frame.margin_right = 0
    write(frame, [(title, size, WHITE if dark else DARK, True)])
    return slide


def overview(prs, d):
    g, now = d["direction"], d["today"]
    s = new_slide(prs, g["title"], True, (0.7, 1.55, 7.5, 1.3), 46)
    box(s, (0.7, 0.8, 7.3, 0.5), [(f"我最想走的路  ·  {d['date']}", 16, MINT, False)], pad=0)
    box(s, (0.7, 2.9, 7.3, 0.9), [(g["detail"], 20, MINT, False)], pad=0)
    for i, c in enumerate(g["chain"]):
        x = 0.7 + i * 2.5
        box(s, (x, 4.55, 2.15, 1.3), [(c["label"], 14, WHITE if i == 0 else MINT, False), (c["text"], 18, WHITE, True)],
            fill=GO if i == 0 else None, line=GO, align=PP_ALIGN.CENTER, pad=0.06)
        if i < 2:
            arrow(s, x + 2.17, x + 2.48, 5.2, MINT)
    box(s, (8.65, 0.8, 4.1, 5.9), [("今天只做这一件", 18, WHITE, True), (now["duration"], 54, WHITE, True),
        (now["title"], 24, WHITE, True), (" ", 12, WHITE, False), (f"做完的标准：{now['done']}", 16, WHITE, False)],
        fill=NOW, pad=0.35)


def amount(n):
    """数字写成上屏的样子：最多一位小数，整数不带小数点，负数用减号，不出科学计数法。"""
    r = round(n, 1)
    text = "0" if r == 0 else f"{r:.1f}".rstrip("0").rstrip(".")
    return text.replace("-", "−")


def span(low, high):
    return amount(low) if low == high else f"{amount(low)}–{amount(high)}"


def road_map(prs, d):
    s = new_slide(prs, "三条路，五年各走到哪、能挣多少")
    unit = d["unit"]
    box(s, (0.55, 1.18, 12.2, 0.4), [(f"现在一年约 {amount(d['now_income'])} {unit}  ·  每格下面是当年收入：保守–乐观，单位 {unit}",
                                      14, MUTED, False)], pad=0)
    for i, label in enumerate(d["periods"]):
        box(s, (3.3 + i * 1.9, 1.62, 1.78, 0.42), [(label, 14, MUTED, True)], align=PP_ALIGN.CENTER)
    for n, r in enumerate(d["routes"]):
        y, picked = 2.12 + n * 1.72, r["key"] == d["direction"]["key"]
        total = span(sum(v["low"] for v in r["years"]), sum(v["high"] for v in r["years"]))
        tag = [("最想走的", 13, WHITE, False)] if picked else []
        box(s, (0.55, y, 2.6, 1.55), [(f"{r['key']}  {r['name']}", 17, WHITE, True), (f"五年合计 {total} {unit}", 12, WHITE, False)] + tag,
            fill=NOW if picked else DARK)
        for i, v in enumerate(r["years"]):
            money = span(v["low"], v["high"])
            box(s, (3.3 + i * 1.9, y, 1.78, 1.55), [(v["goal"], 13, WHITE if picked else INK, False),
                (money, 19 if len(money) <= 9 else 15 if len(money) <= 12 else 12, WHITE if picked else DARK, True)],
                fill=GO if picked else PALE, align=PP_ALIGN.CENTER, pad=0.07)


def detail_cards(prs, d):
    s = new_slide(prs, "每条路过的是什么日子，要付什么代价")
    for i, r in enumerate(d["routes"]):
        x, picked = 0.55 + i * 4.15, r["key"] == d["direction"]["key"]
        box(s, (x, 1.35, 3.95, 5.9), fill=PALE, line=NOW if picked else None)
        box(s, (x + 0.1, 1.45, 3.75, 0.7), [(f"{r['key']}  {r['name']}", 17, DARK, True)], pad=0.05)
        box(s, (x + 0.15, 2.2, 3.65, 2.0), [("五年后普通的一天", 12, MUTED, False), (r["day"], 14, INK, False)],
            fill=WHITE, anchor=MSO_ANCHOR.TOP, pad=0.1)
        for j, (key, label) in enumerate(DETAIL):
            box(s, (x + 0.1, 4.3 + j * 0.95, 3.75, 0.9), [(label, 12, MUTED, False), (r[key], 14, RISK if key == "cost" else INK, True)],
                anchor=MSO_ANCHOR.TOP, pad=0.05)


def wrong_cards(prs, d):
    s = new_slide(prs, "每条选错了会怎样")
    for i, r in enumerate(d["routes"]):
        x, w = 0.55 + i * 4.15, r["wrong"]
        picked, (badge, color) = r["key"] == d["direction"]["key"], BACK[w["level"]]
        box(s, (x, 1.35, 3.95, 5.9), fill=PALE, line=NOW if picked else None)
        box(s, (x + 0.1, 1.45, 2.4, 0.8), [(f"{r['key']}  {r['name']}", 17, DARK, True)], pad=0.05)
        box(s, (x + 2.5, 1.62, 1.35, 0.46), [(badge, 13, INK if color == WARN else WHITE, True)],
            fill=color, align=PP_ALIGN.CENTER, pad=0.03)
        for j, (key, label) in enumerate(WRONG):
            box(s, (x + 0.1, 2.3 + j * 0.8, 3.75, 0.78), [(label, 12, MUTED, False), (w[key], 14, INK, True)],
                anchor=MSO_ANCHOR.TOP, pad=0.05)
        box(s, (x + 0.15, 5.55, 3.65, 0.82), [("止损线", 12, WHITE, False), (w["stop"], 14, WHITE, True)], fill=RISK, pad=0.1)
        box(s, (x + 0.1, 6.42, 3.75, 0.8), [("不选它会错过", 12, MUTED, False), (w["miss"], 14, GO, True)], pad=0.05)


def choice_grid(prs, d):
    g = d["direction"]
    s = new_slide(prs, f"最想的方向：{g['key']}  {g['title']}")
    for i, r in enumerate(d["routes"]):
        x, picked = 3.5 + i * 1.9, r["key"] == g["key"]
        box(s, (x, 1.45, 1.75, 1.05), [(r["key"], 24, WHITE if picked else DARK, True),
            (r["name"], 12, WHITE if picked else DARK, False)], fill=GO if picked else PALE, align=PP_ALIGN.CENTER, pad=0.05)
        for j, (key, _) in enumerate(QUESTIONS):
            y, hit = 2.65 + j * 1.15, r["key"] in d["choice"][key]
            box(s, (x + 0.45, y + 0.08, 0.85, 0.85), [("✓", 28, WHITE, True)] if hit else [("–", 22, MUTED, False)],
                fill=(GO if picked else MUTED) if hit else None, shape=MSO_SHAPE.OVAL, align=PP_ALIGN.CENTER, pad=0)
    for j, (_, label) in enumerate(QUESTIONS):
        box(s, (0.55, 2.65 + j * 1.15, 2.85, 1.0), [(label, 16, INK, True)], fill=PALE)
    box(s, (0.55, 6.2, 8.65, 0.85), [(g["reason"], 16, WHITE, True)], fill=GO)
    box(s, (9.55, 1.45, 3.23, 5.6), [("最想的，不等于最有把握的", 20, WHITE, True), (" ", 12, WHITE, False),
        (g["confident"], 16, MINT, False)], fill=DARK, pad=0.3)


def back_cast(prs, d):
    s = new_slide(prs, "从五年后，一级一级倒推到今天")
    box(s, (0.55, 1.22, 8, 0.45), [("从右往左推，从左往右做", 16, MUTED, False)], pad=0)
    for i, b in enumerate(d["backcast"]):
        x, top = 0.55 + i * 1.77, 4.5 - i * 0.43
        fill = NOW if i == 0 else DARK if i == 6 else GO
        box(s, (x, top, 1.68, 7.1 - top), [(b["date"], 16, WHITE, True), (" ", 8, WHITE, False),
            (b["goal"], 15, WHITE, False)], fill=fill, anchor=MSO_ANCHOR.TOP, pad=0.1)
        if b.get("gate"):
            box(s, (x, top - 0.56, 1.68, 0.48), [(b["gate"], 13, RISK, True)], line=RISK, align=PP_ALIGN.CENTER, pad=0.03)


def this_week(prs, d):
    now = d["today"]
    s = new_slide(prs, "今天做什么，这周做什么")
    box(s, (0.55, 1.4, 5.65, 5.7), fill=NOW)
    box(s, (0.85, 1.6, 5.05, 0.9), [(f"今天  ·  {now['duration']}", 16, WHITE, False), (now["title"], 24, WHITE, True)], pad=0)
    box(s, (0.85, 2.75, 5.05, 3.1), [(now["message"], 18, INK, False)], fill=WHITE, pad=0.25)
    box(s, (0.85, 6.05, 5.05, 0.8), [(f"做完的标准：{now['done']}", 16, WHITE, True)], pad=0)
    for i, row in enumerate(d["week"]):
        y = 1.4 + i * 1.45
        box(s, (6.55, y, 1.5, 1.3), [(row["date"], 16, WHITE, True)], fill=DARK, align=PP_ALIGN.CENTER, pad=0.05)
        box(s, (8.17, y, 4.6, 1.3), [(row["action"], 17, INK, True), (row["pass"], 14, MUTED, False)], fill=PALE)


def premises(prs, d):
    s = new_slide(prs, "动手之前，先确认这几件事", True)
    for i, p in enumerate(d["premises"]):
        x, y = 0.55 + (i % 2) * 6.2, 1.45 + (i // 2) * 2.85
        main = i == 0
        box(s, (x, y, 5.95, 2.6), [(p["title"], 21, WHITE if main else DARK, True), (" ", 8, INK, False),
            (p["text"], 16, WHITE if main else INK, False)], fill=RISK if main else PALE, pad=0.3)
        if p.get("url"):
            link = box(s, (x + 4.1, y + 2.08, 1.7, 0.4), [("看官方原文 ↗", 12, WHITE if main else GO, True)],
                       align=PP_ALIGN.RIGHT, pad=0)
            link.click_action.hyperlink.address = p["url"]


def build(raw, out):
    d = tidy(raw)
    errors = validate(d)
    if errors:
        raise ValueError("\n".join(errors))
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    prs.core_properties.title, prs.core_properties.author = "五年收入推演", ""
    set_theme(prs)
    for page in (overview, road_map, detail_cards, wrong_cards, choice_grid, back_cast, this_week, premises):
        page(prs, d)
    prs.save(out)


def check():
    if not __debug__:
        sys.exit("自检靠 assert，别带 -O 运行")
    example = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "assets", "deck-example.json")
    with open(example, encoding="utf-8") as f:
        d = json.load(f)
    with tempfile.TemporaryDirectory() as tmp:
        out = os.path.join(tmp, "check.pptx")
        build(d, out)
        assert len(Presentation(out).slides) == 8
    assert validate({**d, "choice": {**d["choice"], "regret": []}}), "没拦住替用户选"
    assert validate({**d, "today": {**d["today"], "title": "长" * 99}}), "没拦住超长文字"
    assert validate({}), "没拦住空 JSON"
    assert validate({k: v for k, v in d.items() if k != "date"}), "没拦住缺日期"
    assert validate({**d, "periods": [1, 2, 3, 4, 5]}), "没拦住非文字"
    flipped = [{**r, "years": [{**y, "low": y["high"] + 1} for y in r["years"]]} for r in d["routes"]]
    assert validate({**d, "routes": flipped}), "没拦住保守数比乐观数大"
    assert validate({**d, "now_income": "十万"}), "没拦住收入不是数字"
    assert validate({**d, "now_income": True}), "没拦住布尔值当收入"
    assert validate({**d, "now_income": float("nan")}), "没拦住 NaN"
    assert validate({**d, "routes": [{**r, "years": r["years"][:4]} for r in d["routes"]]}), "没拦住年数不对"
    assert validate({**d, "unit": ""}), "没拦住空文字"
    assert validate({**d, "routes": "三条路"}), "没拦住 routes 不是列表"
    assert amount(1234567) == "1234567" and amount(-0.04) == "0" and amount(-12.5) == "−12.5"
    assert amount(-3.0) == "−3" and span(16, 16) == "16" and span(9.5, 12) == "9.5–12"
    assert validate({**d, "choice": {**d["choice"], "want": "AC"}}), "没拦住 choice 不是列表"
    assert validate({**d, "routes": [d["routes"][0]] * 3}), "没拦住重复路线"
    no_gate = {**d, "backcast": [{**b, "gate": None} for b in d["backcast"]]}
    assert not validate(tidy(no_gate)), "gate 为 null 应该能过"
    print("check ok")


def main(argv):
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")  # Windows 管道默认不是 UTF-8
    if argv == ["--check"]:
        return check()
    if len(argv) != 2 or not argv[1].lower().endswith(".pptx"):
        sys.exit(__doc__)
    src, out = argv
    folder, name = os.path.split(out)
    if any(os.path.exists(os.path.join(folder, n)) for n in (f".~{name}", f"~${name}", f".~lock.{name}#")):
        sys.exit(f"{out} 正在被别的软件打开，关掉再生成")
    try:
        with open(src, encoding="utf-8") as f:
            build(json.load(f), out)
    except (OSError, ValueError) as e:
        sys.exit(f"没生成：\n{e}")
    print(f"已生成 {out}")


if __name__ == "__main__":
    main(sys.argv[1:])
