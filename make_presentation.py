#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Презентация «Великие географические открытия» — версия 3 (киношная).
Артём Платонов, 7 «Д».

Как устроено:
- на каждом слайде сзади во весь экран — настоящая старая карта или картина;
- текста мало, но он крупный, прямо поверх картинки, без белых плашек;
- текст появляется по кусочкам, по щелчку — как раз под рассказ;
- анимации комбинированные: появление, влёт снизу, проявление с увеличением;
- между слайдами стоят переходы.
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
import math, os

BASE = os.path.dirname(os.path.abspath(__file__))
A = lambda n: os.path.join(BASE, "assets", n)
OUT = os.path.join(BASE, "Великие_географические_открытия_Платонов_Артём_7Д.pptx")

WHITE = RGBColor(0xFF, 0xFF, 0xFF)
YEL = RGBColor(0xFF, 0xD9, 0x66)
ORANGE = RGBColor(0xFF, 0xA8, 0x4D)
GREEN = RGBColor(0x8F, 0xE6, 0x8F)
BLUE_L = RGBColor(0x9D, 0xC7, 0xFF)
GREY_L = RGBColor(0xD9, 0xD9, 0xD9)

prs = Presentation()
prs.slide_width = Inches(13.33)
prs.slide_height = Inches(7.5)
prs.core_properties.title = "Великие географические открытия"
prs.core_properties.author = "Артём Платонов Дмитриевич"
BLANK = prs.slide_layouts[6]
W, H = 13.33, 7.5


# ================= анимации =================
def _e(sld, tag, attrs=None):
    return sld.makeelement(qn(tag), attrs or {})


def _set_node(sld, nid, spid):
    set_ = _e(sld, 'p:set')
    cb = _e(sld, 'p:cBhvr')
    ctn = _e(sld, 'p:cTn', {'id': str(nid), 'dur': '1', 'fill': 'hold'})
    st = _e(sld, 'p:stCondLst'); cond = _e(sld, 'p:cond', {'delay': '0'}); st.append(cond); ctn.append(st)
    cb.append(ctn)
    tgt = _e(sld, 'p:tgtEl'); sp = _e(sld, 'p:spTgt', {'spid': str(spid)}); tgt.append(sp); cb.append(tgt)
    al = _e(sld, 'p:attrNameLst'); an = _e(sld, 'p:attrName'); an.text = 'style.visibility'; al.append(an); cb.append(al)
    set_.append(cb)
    to = _e(sld, 'p:to'); sv = _e(sld, 'p:strVal', {'val': 'visible'}); to.append(sv); set_.append(to)
    return set_


def _anim_num(sld, nid, spid, attr, frm, to, dur='600'):
    anim = _e(sld, 'p:anim', {'calcmode': 'lin', 'valueType': 'num'})
    cb = _e(sld, 'p:cBhvr', {'additive': 'base'} if attr in ('ppt_x', 'ppt_y') else {})
    ctn = _e(sld, 'p:cTn', {'id': str(nid), 'dur': dur, 'fill': 'hold'}); cb.append(ctn)
    tgt = _e(sld, 'p:tgtEl'); sp = _e(sld, 'p:spTgt', {'spid': str(spid)}); tgt.append(sp); cb.append(tgt)
    al = _e(sld, 'p:attrNameLst'); an = _e(sld, 'p:attrName'); an.text = attr; al.append(an); cb.append(al)
    anim.append(cb)
    tavl = _e(sld, 'p:tavLst')
    for tm, val, isflt in ((0, frm, False), (100000, to, False)):
        tav = _e(sld, 'p:tav', {'tm': str(tm)}); v = _e(sld, 'p:val')
        if attr in ('ppt_w', 'ppt_h') and tm == 0:
            fv = _e(sld, 'p:fltVal', {'val': '0'}); v.append(fv)
        else:
            sv = _e(sld, 'p:strVal', {'val': val}); v.append(sv)
        tav.append(v); tavl.append(tav)
    anim.append(tavl)
    return anim


def _effect_node(sld, nid, spid, kind, dur):
    ae = _e(sld, 'p:animEffect', {'transition': 'in', 'filter': 'wipe(down)' if kind == 'wipe' else 'fade', 'prLst': ''})
    cb = _e(sld, 'p:cBhvr')
    ctn = _e(sld, 'p:cTn', {'id': str(nid), 'dur': dur}); cb.append(ctn)
    tgt = _e(sld, 'p:tgtEl'); sp = _e(sld, 'p:spTgt', {'spid': str(spid)}); tgt.append(sp); cb.append(tgt)
    ae.append(cb)
    return ae


def add_animations(slide, entries):
    """entries: [(shape, kind)] kind = fade | fly | wipe | zoom. Каждый — по щелчку."""
    if not entries:
        return
    sld = slide._element
    for el in sld.findall(qn('p:timing')):
        sld.remove(el)
    nid = [1]
    def nxt():
        v = nid[0]; nid[0] += 1; return v
    timing = _e(sld, 'p:timing'); tnlst = _e(sld, 'p:tnLst'); timing.append(tnlst)
    root = _e(sld, 'p:par'); rctn = _e(sld, 'p:cTn', {'id': str(nxt()), 'restart': 'never', 'nodeType': 'tmRoot'})
    root.append(rctn); rchild = _e(sld, 'p:childTnLst'); rctn.append(rchild); tnlst.append(root)
    seq = _e(sld, 'p:seq', {'concurrent': '1', 'nextAc': 'seek'})
    sctn = _e(sld, 'p:cTn', {'id': str(nxt()), 'dur': 'indefinite', 'nodeType': 'mainSeq'})
    seq.append(sctn); schild = _e(sld, 'p:childTnLst'); sctn.append(schild); rchild.append(seq)
    PRESET = {'fade': ('10', '0', '600'), 'fly': ('2', '4', '550'), 'wipe': ('22', '4', '600'), 'zoom': ('53', '0', '650')}
    for grp, (shape, kind) in enumerate(entries):
        spid = shape.shape_id
        pid, sub, dur = PRESET[kind]
        par1 = _e(sld, 'p:par'); c1 = _e(sld, 'p:cTn', {'id': str(nxt()), 'fill': 'hold'})
        st = _e(sld, 'p:stCondLst'); st.append(_e(sld, 'p:cond', {'delay': 'indefinite'})); c1.append(st)
        ch1 = _e(sld, 'p:childTnLst'); c1.append(ch1); par1.append(c1)
        par2 = _e(sld, 'p:par'); c2 = _e(sld, 'p:cTn', {'id': str(nxt()), 'fill': 'hold'})
        st = _e(sld, 'p:stCondLst'); st.append(_e(sld, 'p:cond', {'delay': '0'})); c2.append(st)
        ch2 = _e(sld, 'p:childTnLst'); c2.append(ch2); par2.append(c2); ch1.append(par2)
        par3 = _e(sld, 'p:par'); c3 = _e(sld, 'p:cTn', {'id': str(nxt()), 'presetID': pid, 'presetClass': 'entr',
                                                        'presetSubtype': sub, 'fill': 'hold', 'grpId': str(grp), 'nodeType': 'clickEffect'})
        st = _e(sld, 'p:stCondLst'); st.append(_e(sld, 'p:cond', {'delay': '0'})); c3.append(st)
        if kind == 'zoom':
            it = _e(sld, 'p:iterate', {'type': 'lt'}); tp = _e(sld, 'p:tmPct', {'val': '0'}); it.append(tp); c3.append(it)
        ch3 = _e(sld, 'p:childTnLst'); c3.append(ch3); par3.append(c3); ch2.append(par3)
        ch3.append(_set_node(sld, nxt(), spid))
        if kind == 'fade' or kind == 'wipe':
            ch3.append(_effect_node(sld, nxt(), spid, kind, dur))
        elif kind == 'fly':
            ch3.append(_anim_num(sld, nxt(), spid, 'ppt_x', '#ppt_x', '#ppt_x', dur))
            ch3.append(_anim_num(sld, nxt(), spid, 'ppt_y', '1+#ppt_h/2', '#ppt_y', dur))
        elif kind == 'zoom':
            ch3.append(_anim_num(sld, nxt(), spid, 'ppt_w', '0', '#ppt_w', dur))
            ch3.append(_anim_num(sld, nxt(), spid, 'ppt_h', '0', '#ppt_h', dur))
            ch3.append(_effect_node(sld, nxt(), spid, 'fade', dur))
        schild.append(par1)
    for tag, evt in (('p:prevCondLst', 'onPrev'), ('p:nextCondLst', 'onNext')):
        lst = _e(sld, tag); cond = _e(sld, 'p:cond', {'evt': evt, 'delay': '0'})
        tgt = _e(sld, 'p:tgtEl'); tgt.append(_e(sld, 'p:sldTgt')); cond.append(tgt); lst.append(cond); seq.append(lst)
    bld = _e(sld, 'p:bldLst')
    for grp, (shape, kind) in enumerate(entries):
        bld.append(_e(sld, 'p:bldP', {'spid': str(shape.shape_id), 'grpId': str(grp), 'uiExpand': '0',
                                      'advAuto': 'indefinite', 'build': 'whole', 'animBg': '1'}))
    timing.append(bld)
    # timing ставится после transition / clrMapOvr / cSld
    idx = 0
    for i, ch in enumerate(list(sld)):
        if ch.tag in (qn('p:cSld'), qn('p:clrMapOvr'), qn('p:transition')):
            idx = i + 1
    sld.insert(idx, timing)


def set_transition(slide, kind="fade", dur=700):
    sld = slide._element
    for el in sld.findall(qn('p:transition')):
        sld.remove(el)
    trans = _e(sld, 'p:transition', {'spd': 'med', 'dur': str(dur)})
    if kind == "push":
        child = _e(sld, 'p:push', {'dir': 'l'})
    elif kind == "wipe":
        child = _e(sld, 'p:wipe', {'dir': 'l'})
    elif kind == "cover":
        child = _e(sld, 'p:cover', {'dir': 'r'})
    else:
        child = _e(sld, 'p:fade', {})
    trans.append(child)
    idx = 0
    for i, ch in enumerate(list(sld)):
        if ch.tag in (qn('p:cSld'), qn('p:clrMapOvr')):
            idx = i + 1
    sld.insert(idx, trans)


# ================= слайды =================
def bg(slide, img):
    slide.shapes.add_picture(A(img), Inches(0), Inches(0), Inches(W), Inches(H))


def txt(slide, x, y, w, h, runs, size=18, align=PP_ALIGN.LEFT, after=4, line=1.12):
    """runs: [(текст, цвет, bold)] — один абзац из кусочков"""
    tx = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tx.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = align; p.space_after = Pt(after); p.line_spacing = line
    for t, col, bold in runs:
        r = p.add_run(); r.text = t
        r.font.size = Pt(size); r.font.bold = bold; r.font.color.rgb = col; r.font.name = "Calibri"
    return tx


def plain(slide, x, y, w, h, text_, size=18, color=WHITE, bold=False, align=PP_ALIGN.LEFT):
    return txt(slide, x, y, w, h, [(text_, color, bold)], size=size, align=align)


def title(slide, n, text_, img, kicker=None):
    bg(slide, img)
    if kicker:
        plain(slide, 0.6, 0.42, 10, 0.35, kicker, size=11, color=YEL, bold=True)
        ty = 0.78
    else:
        ty = 0.55
    plain(slide, 0.6, ty, 12.1, 0.85, text_, size=31, color=WHITE, bold=True)
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.62), Inches(ty + 0.82), Inches(1.6), Pt(4))
    line.fill.solid(); line.fill.fore_color.rgb = YEL; line.line.fill.background(); line.shadow.inherit_shadow = False
    plain(slide, 0.6, 7.12, 7, 0.25, "Артём Платонов  •  7 «Д»", size=9, color=GREY_L)
    plain(slide, 12.05, 7.12, 0.75, 0.25, f"{n} / 20", size=9, color=WHITE, bold=True, align=PP_ALIGN.RIGHT)
    return ty + 1.15


def steps(slide, items, x=0.6, y=2.05, w=8.2, size=18, effects=None):
    """items: [(lead, rest)] — каждый пункт отдельным блоком, появляется по щелчку"""
    shapes = []
    effects = effects or ["fly", "fade", "fly", "fade", "fly", "fade"]
    yy = y
    for i, (lead, rest) in enumerate(items):
        full = (lead + " " + rest).strip()
        chars = max(30, int(w * 7.2))
        lines = max(1, math.ceil(len(full) / chars))
        h = lines * (size / 52) + 0.22
        runs = []
        if lead:
            runs.append((lead + "  ", YEL, True))
        if rest:
            runs.append((rest, WHITE, False))
        sh = txt(slide, x, yy, w, h, runs, size=size)
        shapes.append((sh, effects[i % len(effects)]))
        yy += h + 0.28
    return shapes


def foot_step(slide, text_, y=6.35, effect="zoom", size=16):
    sh = txt(slide, 0.6, y, 12.1, 0.65, [("Важно:  ", YEL, True), (text_, WHITE, True)], size=size)
    return (sh, effect)


def note(slide, text_):
    slide.notes_slide.placeholders[1].text = text_


slides = []
anim = []   # параллельно: список анимаций на слайд

# ---- 1 титул
s = prs.slides.add_slide(BLANK); slides.append(s)
bg(s, "b_title.jpg")
plain(s, 0.6, 0.7, 8.5, 0.4, "ПРОЕКТ ПО ИСТОРИИ", size=13, color=YEL, bold=True)
plain(s, 0.6, 1.25, 9.2, 1.9, "Великие географические открытия", size=42, color=WHITE, bold=True)
a = []
a.append((plain(s, 0.6, 3.15, 8.6, 0.9, "Как люди поплыли туда, где на картах было пусто — и что из этого вышло", size=19, color=GREY_L), "fade"))
a.append((txt(s, 0.6, 4.6, 9.0, 1.1, [("Выполнил:  ", YEL, True), ("Артём Платонов Дмитриевич, 7 «Д» класс", WHITE, True)], size=20), "zoom"))
a.append((plain(s, 0.6, 5.85, 9.0, 0.5, "Путешественники  •  Колумбов обмен  •  Что изменилось в мире", size=14, color=YEL), "fade"))
anim.append(a)
note(s, "Здравствуйте, меня зовут Артём Платонов, я из 7 «Д». Это мой проект про Великие географические открытия. Нажимаю дальше — и рассказываю по порядку.")

# ---- 2 обо мне и план
s = prs.slides.add_slide(BLANK); slides.append(s)
y = title(s, 2, "Привет! Коротко обо мне и о плане", "b_caravels.jpg")
anim.append(steps(s, [
    ("Я — Артём Платонов,", "учусь в 7 «Д» классе."),
    ("Почему эта тема?", "Мне нравятся карты. Было интересно, как люди плыли туда, где на карте ещё ничего нет."),
    ("Сначала расскажу про людей:", "Колумб, Васко да Гама, Магеллан и ещё трое."),
    ("Потом — про Колумбов обмен.", "Что Америка и Европа дали друг другу."),
    ("И в конце — что после этого изменилось", "в мире. Хорошего и, честно, плохого тоже."),
], y=y + 0.1, w=8.6))
note(s, "Представляюсь и даю план из трёх частей. Текст появляется по щелчку, поэтому я не читаю весь слайд сразу, а говорю по пунктам.")

# ---- 3 причины
s = prs.slides.add_slide(BLANK); slides.append(s)
y = title(s, 3, "Почему они вообще поплыли так далеко?", "b_map.jpg")
anim.append(steps(s, [
    ("Пряности стоили как золото.", "Перец и корицу везли через кучу посредников, и каждый набавлял цену."),
    ("Старая дорога стала неудобной.", "Через земли Османской империи торговать стало дорого и неспокойно."),
    ("А корабли уже умели.", "Каравелла могла идти против ветра, были компас и астролябия."),
    ("И королям хотелось славы.", "Испания и Португалия соревновались и давали деньги на экспедиции."),
], y=y + 0.1))
note(s, "Четыре причины, по одному щелчку на каждую. Главное, что всё совпало одновременно: и зачем плыть было понятно, и на чём уже было можно.")

# ---- 4 путешественники intro
s = prs.slides.add_slide(BLANK); slides.append(s)
y = title(s, 4, "Часть 1. Люди, которые поплыли", "b_caravels.jpg", "ТЕМА 1 ИЗ УЧЕБНИКА")
anim.append(steps(s, [
    ("В учебнике дают шестерых на выбор.", "Я взял всех — так виднее, как заполнялась карта."),
    ("1492 — Колумб.", "1498 — Васко да Гама в Индии. 1522 — вернулась «Виктория» Магеллана."),
    ("Дальше север:", "Картье в Канаде, Дрейк вокруг света, Баренц у Новой Земли."),
    ("Сто лет — и все океаны пройдены.", "Даже Северный Ледовитый попробовали."),
], y=y + 0.1, w=9.2))
note(s, "Переход к первой теме. Я не выбрал одного путешественника, а рассказываю про всех шестерых коротко.")

# ---- 5 Колумб
s = prs.slides.add_slide(BLANK); slides.append(s)
y = title(s, 5, "Колумб: хотел в Индию, а открыл Америку", "b_columbus.jpg", "1492 ГОД  •  СПРАВА — КАРТИНА ВАНДЕРЛИНА О ЕГО ВЫСАДКЕ")
anim.append(steps(s, [
    ("Он считал Землю маленькой.", "Решил плыть на запад, в Индию — и ошибся в расчётах."),
    ("3 августа 1492 года", "вышли «Санта-Мария», «Пинта» и «Нинья»."),
    ("12 октября увидели землю.", "Колумб решил, что это Индия, и назвал жителей индейцами."),
    ("Он плавал ещё три раза,", "но так и не понял, что открыл новый материк."),
], y=y + 0.15, w=9.3, size=16, effects=["fade", "fly", "zoom", "fade"]))
anim[-1].append(foot_step(s, "Америку назвали в честь Америго Веспуччи — он первым понял, что это новый материк.", y=6.45, effect="fade", size=14.5))
note(s, "Про Колумб главное — он ошибся и не понял своего открытия. Дата 12 октября 1492 года — её стоит запомнить. Картина на фоне — высадка Колумб, её нарисовал Джон Вандерлин в XIX веке.")

# ---- 6 Васко да Гама
s = prs.slides.add_slide(BLANK); slides.append(s)
y = title(s, 6, "Васко да Гама: а этот до Индии доплыл", "b_gama.jpg", "1497–1499  •  НА ФОНЕ — ЕГО ПРИБЫТИЕ В ИНДИЮ")
anim.append(steps(s, [
    ("Португальцы долго пытались обогнуть Африку.", "Все поворачивали назад, а Васко да Гама дошёл до конца."),
    ("Он обошёл мыс Доброй Надежды", "и пересёк Индийский океан."),
    ("В мае 1498 года — Каликут.", "Та самая настоящая Индия. Пряности теперь повезли прямо на кораблях."),
    ("Португалия на этом разбогатела.", "Правда, из команды домой вернулись далеко не все."),
], y=y + 0.15, w=7.9))
note(s, "Васко да Гама выполнил то задание, которое Колумб выполнить не смог, только другим путём — вокруг Африки. Литография на фоне показывает его прибытие в Каликут.")

# ---- 7 Магеллан
s = prs.slides.add_slide(BLANK); slides.append(s)
y = title(s, 7, "Магеллан: вокруг всей Земли", "b_victoria.jpg", "1519–1522  •  НА ФОНЕ — «ВИКТОРИЯ» СО СТАРОЙ КАРТЫ")
anim.append(steps(s, [
    ("1519 год.", "Из Испании вышли 5 кораблей и больше 250 человек."),
    ("1520 год.", "Нашли пролив внизу Южной Америки. Теперь он называется Магелланов."),
    ("Потом — Тихий океан.", "Шли почти четыре месяца. Еды почти не было."),
    ("1521 год.", "Магеллан погиб на Филиппинах. Кругосветку он не закончил."),
    ("1522 год.", "Элькано привёл домой один корабль, «Викторию», и 18 человек."),
], y=y + 0.1, w=8.4, size=17))
note(s, "История по годам, каждый год — отдельный щелчок. Вышли 250 человек, вернулись 18. Зато стало ясно: Земля круглая, а океаны соединены.")

# ---- 8 трое
s = prs.slides.add_slide(BLANK); slides.append(s)
y = title(s, 8, "Ещё трое: Канада, пират и Арктика", "b_caravels.jpg")
anim.append(steps(s, [
    ("Жак Картье, Франция.", "Исследовал Канаду и реку Святого Лаврентия. От индейского слова «каната», по одной версии, и пошло название Канада."),
    ("Фрэнсис Дрейк, Англия.", "Для испанцев пират, для англичан герой. На «Золотой лани» обошёл вокруг света и грабил испанские корабли с серебром."),
    ("Виллем Баренц, Голландия.", "Искал путь в Китай мимо Сибири. Зимовал на Новой Земле и погиб на обратном пути. Море там назвали Баренцевым."),
], y=y + 0.15, w=9.4, size=17.5))
note(s, "Трое оставшихся, по одному абзацу на каждого. Все искали короткий путь в Азию, а нашли совсем другие места.")

# ---- 9 все коротко
s = prs.slides.add_slide(BLANK); slides.append(s)
y = title(s, 9, "Все шесть — в шесть строк", "b_map.jpg", "ПОВТОРЕНИЕ ПЕРВОЙ ТЕМЫ")
anim.append(steps(s, [
    ("1492–1504, Колумб —", "Америка и Карибские острова."),
    ("1497–1499, Васко да Гама —", "морская дорога в Индию."),
    ("1519–1522, экспедиция Магеллана —", "первая кругосветка."),
    ("1534–1542, Картье —", "Канада."),
    ("1577–1580, Дрейк —", "вторая кругосветка."),
    ("1594–1597, Баренц —", "Новая Земля и поиск северного пути."),
], y=y + 0.05, w=9.6, size=16.5, effects=["fade"] * 6))
note(s, "Повторение первой темы. Каждая строка появляется отдельно — можно прямо по этому слайду отвечать учителю.")

# ---- 10 Колумбов обмен
s = prs.slides.add_slide(BLANK); slides.append(s)
y = title(s, 10, "Часть 2. Колумбов обмен — это что?", "b_map.jpg", "ТЕМА 2 ИЗ УЧЕБНИКА")
a = steps(s, [
    ("Объясняю своими словами:", "после 1492 года Старый Свет и Америка стали обмениваться всем подряд."),
    ("Растениями и животными.", "Болезнями. Людьми. И даже привычками и едой."),
    ("Название придумал историк Альфред Кросби.", "Но недавно, в 1972 году. А сам обмен идёт со времён Колумб."),
    ("До этого два мира жили отдельно", "и вообще не знали друг про друга."),
], y=y + 0.1, w=9.4)
anim.append(a)
note(s, "Вторая тема. Объясняю термин простыми словами и сразу говорю, что название новое, а явление старое.")

# ---- 11 растения
s = prs.slides.add_slide(BLANK); slides.append(s)
y = title(s, 11, "Растения, которые переехали через океан", "b_gama.jpg")
a = steps(s, [
    ("Из Америки к нам приехали:", "картофель, помидоры, кукуруза, фасоль, какао, перец, ананасы, арахис и подсолнух."),
    ("А в Америку из Старого Света:", "пшеница, рис, кофе, бананы, виноград, яблоки и сахарный тростник."),
    ("Картошка и кукуруза потом всех выручали:", "они урожайнее пшеницы и растут там, где хлеб не растёт."),
], y=y + 0.1, w=8.6, size=17.5)
a.append(foot_step(s, "Пицца, например: тесто из пшеницы Старого Света, а соус — из американских помидоров. До Колумб пиццы не было.", y=6.35, effect="fade", size=15))
anim.append(a)
note(s, "Сначала что приехало из Америки, потом что уехало туда. Пример с пиццей в конце — его все понимают и запоминают.")

# ---- 12 животные
s = prs.slides.add_slide(BLANK); slides.append(s)
y = title(s, 12, "Животные: обмен получился неравный", "b_caravels.jpg")
anim.append(steps(s, [
    ("Из Европы в Америку привезли", "лошадей, коров, свиней, овец и кур. Там таких животных просто не было."),
    ("Обратно в Европу широко уехала", "только индейка. Ламы так и остались жить в Андах."),
    ("Главное последствие — лошадь.", "Индейцы прерий научились ездить верхом, и охота на бизонов и вся их жизнь изменились."),
    ("А ещё зайцами приплывали крысы.", "Это уже не подарок, а вред."),
], y=y + 0.1, w=9.2))
note(s, "Животных Европа дала много, а получила в основном индейку. Но лошадь для Америки оказалась важнее всего остального.")

# ---- 13 болезни
s = prs.slides.add_slide(BLANK); slides.append(s)
y = title(s, 13, "Тяжёлая часть: болезни и люди", "b_teno.jpg")
anim.append(steps(s, [
    ("Вместе с кораблями приплыли оспа, корь и грипп.", ""),
    ("У европейцев организм хоть как-то умел с ними бороться.", "У индейцев защиты не было никакой."),
    ("Умирали целые селения.", "Иногда болезнь приходила в деревню раньше самих европейцев."),
    ("А потом на плантации стали силой привозить людей из Африки.", "Миллионы человек переселили не по их воле."),
], y=y + 0.15, w=9.4, effects=["fade"] * 4))
note(s, "Грустный слайд, тут анимации спокойные, без влётов. Без этого слайда рассказ про обмен получился бы неправдивым.")

# ---- 14 культуры
s = prs.slides.add_slide(BLANK); slides.append(s)
y = title(s, 14, "А ещё встретились культуры", "b_gama.jpg")
anim.append(steps(s, [
    ("Языки.", "В Латинской Америке говорят по-испански и по-португальски. Но индейские языки — кечуа, гуарани — живы до сих пор."),
    ("Вера и люди.", "Главной религией стало христианство. А от смешанных семей появились метисы и мулаты — теперь это большая часть населения."),
    ("Еда и города.", "Города строили по-европейски, а ели уже смешанное: кукурузные лепёшки и пшеничный хлеб на одном столе."),
    ("И карты стали настоящими.", "Оба берега Атлантики и Тихий океан наконец появились на картах."),
], y=y + 0.1, w=9.2, size=17))
note(s, "Культуры смешались: языки, религия, еда, состав населения. И Атлантика перестала быть краем света.")

# ---- 15 последствия +
s = prs.slides.add_slide(BLANK); slides.append(s)
y = title(s, 15, "Часть 3. Что хорошего из этого вышло", "b_map.jpg", "ТЕМА 3 ИЗ УЧЕБНИКА")
anim.append(steps(s, [
    ("Торговля стала мировой.", "Товары постоянно ходили между всеми материками, а не раз в сто лет."),
    ("Выросли новые порты:", "Лиссабон, Севилья, Амстердам, Лондон. Главным морем стала Атлантика."),
    ("Наука получила факты вместо догадок.", "Карты стали точными, а что Земля круглая, уже никто не оспаривал."),
    ("И еды стало больше.", "Картофель и кукуруза накормили растущее население Европы и Азии."),
], y=y + 0.1, w=9.4))
note(s, "Третья тема, сначала хорошая сторона последствий: торговля, порты, точные карты и новые продукты.")

# ---- 16 последствия −
s = prs.slides.add_slide(BLANK); slides.append(s)
y = title(s, 16, "А что плохого — честно", "b_teno.jpg", "НА ФОНЕ — ЗАВОЕВАНИЕ ТЕНОЧТИТЛАНА, КАРТИНА XVII ВЕКА")
anim.append(steps(s, [
    ("Кортес и Писарро завоевали государства ацтеков и инков.", "Испанцев было мало, но у них были лошади, сталь и союзники из индейцев."),
    ("От болезней и войн погибло огромное количество людей.", ""),
    ("На плантациях работали рабы из Африки.", "Работорговля длилась несколько веков."),
    ("И Европе досталось:", "от дешёвого серебра выросли цены, бедные стали ещё беднее."),
], y=y + 0.15, w=8.2, effects=["fade"] * 4))
note(s, "Обратная сторона на фоне картины завоевания Теночтитлана. Вывод: для Европы рывок вперёд, для Америки и Африки — беда. И то и другое правда.")

# ---- 17 книги и кино
s = prs.slides.add_slide(BLANK); slides.append(s)
y = title(s, 17, "Про это написали книги и сняли кино", "b_columbus.jpg", "СПИСОК — ИЗ УЧЕБНИКА, С. 38")
anim.append(steps(s, [
    ("Книги:", "Гуляев «По следам конкистадоров», Свет «Колумб», Кинжалов и Белов «Падение Теночтитлана»."),
    ("Приключения:", "Хаггард «Дочь Монтесумы», Кофман «Тьерра аделанте!»."),
    ("Мои рекомендации:", "Цвейг, «Подвиг Магеллана» и «Америго». Небольшие и читаются как роман."),
    ("Кино и памятники:", "фильм «Одиссея капитана Блада» (1991) и колонна Колумб в Барселоне."),
], y=y + 0.1, w=8.2, size=17))
note(s, "Задание из учебника — собрать книги и фильмы. Список со страницы 38, от себя добавил рекомендацию начать с Цвейга.")

# ---- 18 почему интересно
s = prs.slides.add_slide(BLANK); slides.append(s)
y = title(s, 18, "Почему про это до сих пор пишут и снимают?", "b_victoria.jpg")
anim.append(steps(s, [
    ("Потому что это про смелость.", "Уплыть в океан без точной карты — понятно и сейчас."),
    ("Потому что это встреча двух миров,", "которые ничего друг о друге не знали. Такой сюжет теперь даже в фантастике про космос используют."),
    ("Герои неоднозначные.", "Колумб — и первооткрыватель, и с него началось завоевание. Про таких интереснее читать."),
    ("И загадки остались.", "Куда делось золото инков, спорят до сих пор. А где спорят, там будут новые книги."),
], y=y + 0.1, w=9.2, size=17))
note(s, "Отвечаю на вопрос из учебника четырьмя причинами. Последняя фраза — моё мнение, могу его пояснить, если спросят.")

# ---- 19 вывод
s = prs.slides.add_slide(BLANK); slides.append(s)
y = title(s, 19, "Мой вывод", "b_title.jpg", "ОТВЕТ НА ГЛАВНЫЙ ВОПРОС ГЛАВЫ")
a = steps(s, [
    ("Первое. Мир стал одним целым.", "Товары, растения, деньги и люди стали постоянно ездить по всей планете."),
    ("Второе. У этого была цена.", "Европа получила торговлю и знания, Америка и Африка заплатили болезнями и рабством."),
    ("Третье. Всё решили техника и люди.", "Без каравеллы и компаса не вышло бы ничего. Но и без смельчаков корабли стояли бы в порту."),
], y=y + 0.15, w=9.6)
a.append((plain(s, 0.6, 6.15, 12.1, 0.6, "Спасибо за внимание! Отвечу на вопросы.", size=21, color=YEL, bold=True, align=PP_ALIGN.CENTER), "zoom"))
anim.append(a)
note(s, "Три пункта вывода, появляются по одному. Потом благодарю за внимание. Дальше — источники, если учитель попросит показать.")

# ---- 20 источники
s = prs.slides.add_slide(BLANK); slides.append(s)
y = title(s, 20, "Откуда я всё это брал", "b_map.jpg")
a = []
a.append((txt(s, 0.6, y + 0.1, 12.0, 1.5, [
    ("Учебник истории Нового времени, 7 класс — глава про открытия, задание и книги на с. 38.", WHITE, False),
], size=14), "fade"))
# источники тремя блоками поменьше (крупный текст тут не влезет, это справочный слайд)
blocks = [
    [("Большая российская энциклопедия — bigenc.ru", YEL, True), (" (статьи про Колумб, Васко да Гаму, Магеллана — там проверял даты)", WHITE, False)],
    [("Портал «История. РФ» — histrf.ru", YEL, True), (" (материалы к главе, на них QR-коды в учебнике)", WHITE, False)],
    [("Книги: ", YEL, True), ("Гуляев, Свет, Цвейг, Хаггард и другие со слайда 17. Фильм «Одиссея капитана Блада» (1991).", WHITE, False)],
    [("Картинки на фоне: ", YEL, True), ("старинная карта мира; высадка Колумб — картина Дж. Вандерлина; Васко да Гама в Каликуте — литография; «Виктория» с карты Ортелиуса; завоевание Теночтитлана — картина XVII века. Источники: Wikimedia Commons, worldhistory.org, ncpedia.org.", WHITE, False)],
    [("Текст ", YEL, True), ("писал сам по учебнику и энциклопедии. Совпадают только даты и имена — их иначе не напишешь.", WHITE, False)],
]
yy = y + 0.75
for b in blocks:
    sh = txt(s, 0.6, yy, 12.0, 0.95, b, size=13.5, line=1.15)
    a.append((sh, "fade"))
    yy += 1.0
anim.append(a)
note(s, "Источники списком, как положено в проекте. Отдельно сказал, откуда картинки на фоне — это старые карты и картины из открытых энциклопедий.")

# ---- переходы + анимации
trans_kinds = ["fade", "push", "wipe", "fade", "cover"] * 4
for sl, a, tk in zip(slides, anim, trans_kinds):
    set_transition(sl, tk)
    add_animations(sl, a)

prs.save(OUT)
print("Готово:", OUT, "| слайдов:", len(slides), "| анимированных блоков:", sum(len(a) for a in anim))
