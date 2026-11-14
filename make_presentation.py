#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Презентация v7: исправлены факты по рецензии, добавлен Васко да Гама, крупный плотный текст.
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

BLUE = RGBColor(0x1F, 0x38, 0x64)
BLUE2 = RGBColor(0x2E, 0x5A, 0x88)
ORANGE = RGBColor(0xB5, 0x4A, 0x12)
GREEN = RGBColor(0x2F, 0x5D, 0x2A)
RED = RGBColor(0x8F, 0x26, 0x26)
DARK = RGBColor(0x24, 0x24, 0x24)
GREY = RGBColor(0x5A, 0x5A, 0x5A)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
YEL = RGBColor(0xFF, 0xD9, 0x66)
YELLOW_BG = RGBColor(0xFF, 0xF2, 0xCC)
LIGHT = RGBColor(0xDE, 0xE6, 0xF2)
BORDER = RGBColor(0xB4, 0xC2, 0xD6)

prs = Presentation()
prs.slide_width = Inches(13.33)
prs.slide_height = Inches(7.5)
prs.core_properties.title = "Великие географические открытия"
prs.core_properties.author = "Артём Платонов Дмитриевич"
BLANK = prs.slide_layouts[6]
W, H = 13.33, 7.5


# ================= анимации (появление по щелчку) =================
def _e(sld, tag, attrs=None):
    return sld.makeelement(qn(tag), attrs or {})


def _set_node(sld, nid, spid):
    set_ = _e(sld, 'p:set'); cb = _e(sld, 'p:cBhvr')
    ctn = _e(sld, 'p:cTn', {'id': str(nid), 'dur': '1', 'fill': 'hold'})
    st = _e(sld, 'p:stCondLst'); st.append(_e(sld, 'p:cond', {'delay': '0'})); ctn.append(st); cb.append(ctn)
    tgt = _e(sld, 'p:tgtEl'); tgt.append(_e(sld, 'p:spTgt', {'spid': str(spid)})); cb.append(tgt)
    al = _e(sld, 'p:attrNameLst'); an = _e(sld, 'p:attrName'); an.text = 'style.visibility'; al.append(an); cb.append(al)
    set_.append(cb)
    to = _e(sld, 'p:to'); to.append(_e(sld, 'p:strVal', {'val': 'visible'})); set_.append(to)
    return set_


def _anim_num(sld, nid, spid, attr, frm, to, dur='600'):
    anim = _e(sld, 'p:anim', {'calcmode': 'lin', 'valueType': 'num'})
    cb = _e(sld, 'p:cBhvr', {'additive': 'base'} if attr in ('ppt_x', 'ppt_y') else {})
    cb.append(_e(sld, 'p:cTn', {'id': str(nid), 'dur': dur, 'fill': 'hold'}))
    tgt = _e(sld, 'p:tgtEl'); tgt.append(_e(sld, 'p:spTgt', {'spid': str(spid)})); cb.append(tgt)
    al = _e(sld, 'p:attrNameLst'); an = _e(sld, 'p:attrName'); an.text = attr; al.append(an); cb.append(al)
    anim.append(cb)
    tavl = _e(sld, 'p:tavLst')
    for tm, val in ((0, frm), (100000, to)):
        tav = _e(sld, 'p:tav', {'tm': str(tm)}); v = _e(sld, 'p:val')
        if attr in ('ppt_w', 'ppt_h') and tm == 0:
            v.append(_e(sld, 'p:fltVal', {'val': '0'}))
        else:
            v.append(_e(sld, 'p:strVal', {'val': val}))
        tav.append(v); tavl.append(tav)
    anim.append(tavl)
    return anim


def _effect_node(sld, nid, spid, kind, dur):
    ae = _e(sld, 'p:animEffect', {'transition': 'in', 'filter': 'wipe(down)' if kind == 'wipe' else 'fade', 'prLst': ''})
    cb = _e(sld, 'p:cBhvr'); cb.append(_e(sld, 'p:cTn', {'id': str(nid), 'dur': dur}))
    tgt = _e(sld, 'p:tgtEl'); tgt.append(_e(sld, 'p:spTgt', {'spid': str(spid)})); cb.append(tgt)
    ae.append(cb)
    return ae


def add_animations(slide, entries):
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
        spid = shape.shape_id; pid, sub, dur = PRESET[kind]
        par1 = _e(sld, 'p:par'); c1 = _e(sld, 'p:cTn', {'id': str(nxt()), 'fill': 'hold'})
        st = _e(sld, 'p:stCondLst'); st.append(_e(sld, 'p:cond', {'delay': 'indefinite'})); c1.append(st)
        ch1 = _e(sld, 'p:childTnLst'); c1.append(ch1); par1.append(c1)
        par2 = _e(sld, 'p:par'); c2 = _e(sld, 'p:cTn', {'id': str(nxt()), 'fill': 'hold'})
        st = _e(sld, 'p:stCondLst'); st.append(_e(sld, 'p:cond', {'delay': '0'})); c2.append(st)
        ch2 = _e(sld, 'p:childTnLst'); c2.append(ch2); par2.append(c2); ch1.append(par2)
        par3 = _e(sld, 'p:par')
        c3 = _e(sld, 'p:cTn', {'id': str(nxt()), 'presetID': pid, 'presetClass': 'entr', 'presetSubtype': sub,
                               'fill': 'hold', 'grpId': str(grp), 'nodeType': 'clickEffect'})
        st = _e(sld, 'p:stCondLst'); st.append(_e(sld, 'p:cond', {'delay': '0'})); c3.append(st)
        if kind == 'zoom':
            it = _e(sld, 'p:iterate', {'type': 'lt'}); it.append(_e(sld, 'p:tmPct', {'val': '0'})); c3.append(it)
        ch3 = _e(sld, 'p:childTnLst'); c3.append(ch3); par3.append(c3); ch2.append(par3)
        ch3.append(_set_node(sld, nxt(), spid))
        if kind in ('fade', 'wipe'):
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
    child = {'push': ('p:push', {'dir': 'l'}), 'wipe': ('p:wipe', {'dir': 'l'}),
             'cover': ('p:cover', {'dir': 'r'})}.get(kind, ('p:fade', {}))
    trans.append(_e(sld, child[0], child[1]))
    idx = 0
    for i, ch in enumerate(list(sld)):
        if ch.tag in (qn('p:cSld'), qn('p:clrMapOvr')):
            idx = i + 1
    sld.insert(idx, trans)


# ================= оформление =================
# ================= оформление v7 =================
# 1 слайд = 1 мысль, заголовок = вывод, шрифт крупный, интервалы плотные.

def bg(slide, img="bg_light.jpg"):
    slide.shapes.add_picture(A(img), Inches(0), Inches(0), Inches(W), Inches(H))

def rect(slide, x, y, w, h, fill=WHITE, line=BORDER, round_=True):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if round_ else MSO_SHAPE.RECTANGLE,
                                 Inches(x), Inches(y), Inches(w), Inches(h))
    if round_:
        try: shp.adjustments[0] = 0.06
        except Exception: pass
    shp.fill.solid(); shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line; shp.line.width = Pt(1)
    shp.shadow.inherit_shadow = False
    return shp

def _fill_tf(tf, paras, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = Inches(0.08); tf.margin_right = Inches(0.08)
    tf.margin_top = Inches(0.04); tf.margin_bottom = Inches(0.04)
    for i, (runs, al, spacing, space_after) in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = al if al is not None else align
        p.line_spacing = spacing
        if space_after: p.space_after = Pt(space_after)
        for t, col, bold, size, italic in runs:
            r = p.add_run(); r.text = t
            r.font.size = Pt(size); r.font.bold = bold; r.font.italic = italic
            r.font.color.rgb = col; r.font.name = "Calibri"
    return tf

def box(slide, x, y, w, h, paras, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    tx = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    _fill_tf(tx.text_frame, paras, align=align, anchor=anchor)
    return tx

def one(slide, x, y, w, h, text_, size=18, color=DARK, bold=False,
        align=PP_ALIGN.LEFT, italic=False, anchor=MSO_ANCHOR.TOP):
    return box(slide, x, y, w, h,
               [([(text_, color, bold, size, italic)], align, 1.05, 0)],
               align=align, anchor=anchor)

def shape_text(shp, paras, align=PP_ALIGN.CENTER):
    _fill_tf(shp.text_frame, paras, align=align, anchor=MSO_ANCHOR.MIDDLE)
    return shp

def header(slide, n, kicker, title, accent=ORANGE, title_size=32):
    bg(slide)
    one(slide, 0.6, 0.26, 12.1, 0.32, kicker.upper(), size=12, color=accent, bold=True)
    one(slide, 0.6, 0.56, 12.13, 1.08, title, size=title_size, color=BLUE, bold=True)
    rect(slide, 0.62, 1.70, 1.55, 0.055, fill=accent, line=None, round_=False)
    one(slide, 0.6, 7.16, 8, 0.25, "Платонов Артём  •  7 «Д»", size=9, color=GREY)
    one(slide, 12.0, 7.16, 0.73, 0.25, f"{n} / 20", size=9, color=BLUE, bold=True, align=PP_ALIGN.RIGHT)
    return 1.95

def photo(slide, img, x, y, w, h, caption):
    rect(slide, x - 0.07, y - 0.07, w + 0.14, h + 0.14, fill=WHITE, line=BORDER, round_=False)
    pic = slide.shapes.add_picture(A(img), Inches(x), Inches(y), Inches(w), Inches(h))
    one(slide, x - 0.07, y + h + 0.10, w + 0.14, 0.5, caption, size=9.5,
        color=GREY, italic=True, align=PP_ALIGN.CENTER)
    return pic

def bullets(slide, items, x=0.6, y=1.95, w=7.0, size=21, accent=BLUE):
    """Крупно и плотно: интервал между тезисами маленький."""
    out = []; yy = y
    for i, (lead, rest) in enumerate(items):
        full = (lead + " " + rest).strip()
        chars = max(26, int(w * 108 / size))
        lines = max(1, math.ceil(len(full) / chars))
        h = lines * (size / 45.0) + 0.12
        runs = [("▸  ", accent, True, size, False)]
        if lead: runs.append((lead + "  ", BLUE, True, size, False))
        if rest: runs.append((rest, DARK, False, size, False))
        sh = box(slide, x, yy, w, h, [(runs, PP_ALIGN.LEFT, 1.04, 0)])
        out.append((sh, "fly" if i % 2 == 0 else "fade"))
        yy += h + 0.14
    return out, yy

def fact(slide, x, y, w, h, big, small, fill=BLUE, big_size=42):
    shp = rect(slide, x, y, w, h, fill=fill, line=None)
    shape_text(shp, [
        ([(big, WHITE, True, big_size, False)], PP_ALIGN.CENTER, 1.0, 2),
        ([(small, WHITE, False, 14.5, False)], PP_ALIGN.CENTER, 1.05, 0),
    ])
    return shp

def tile(slide, x, y, w, h, title_, sub, fill=WHITE, tcol=BLUE, tsize=22, ssize=14.5):
    shp = rect(slide, x, y, w, h, fill=fill, line=BORDER)
    shape_text(shp, [
        ([(title_, tcol, True, tsize, False)], PP_ALIGN.CENTER, 1.0, 2),
        ([(sub, DARK, False, ssize, False)], PP_ALIGN.CENTER, 1.05, 0),
    ])
    return shp

def arrow(slide, x, y, w=0.85, kind=MSO_SHAPE.RIGHT_ARROW, fill=BLUE2, h=0.38):
    ar = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    ar.fill.solid(); ar.fill.fore_color.rgb = fill; ar.line.fill.background()
    ar.shadow.inherit_shadow = False
    return ar

def note(slide, text_):
    slide.notes_slide.placeholders[1].text = text_

slides = []; anim = []

# ---------- речь с исправленными фактами ----------
NOTES = {
1: "Здравствуйте, ребята и уважаемая [Имя Отчество]! Сегодня мы с вами отправимся в путешествие, которое навсегда изменило нашу планету. Мы поговорим об эпохе Великих географических открытий. И речь пойдет не просто о скучных датах и именах. Мы узнаем, как Европа и Америка «нашли» друг друга, почему картошка стала спасением для миллионов людей, как родилась глобализация, и почему приключенческие романы и фильмы до сих пор так популярны. Пристегните ремни, мы отправляемся в конец XV века!",
2: "В самом начале нашей главы учебник ставит главный вопрос: «Почему Великие географические открытия считаются переломным моментом в истории?» Вдумайтесь: до XV века люди жили, будто в разных параллельных мирах. Европейцы не знали, что есть Америка. Американские индейцы не знали, что есть Европа, Азия и Африка. Китай торговал с другими странами по Великому шёлковому пути и по морю, но прямых связей между Европой и Америкой не было ни у кого. Мир был огромным пазлом, кусочки которого не были соединены. И вот, благодаря отчаянным мореплавателям, этот пазл сложился. В конце своего рассказа я четко сформулирую три аргумента, почему же этот момент стал поворотным для всего человечества.",
3: "Итак, конец XV века. После падения Константинополя в 1453 году старый путь на восток стал ещё труднее. Европа сходит с ума по пряностям. Перец, корица, гвоздика стоят бешеных денег, потому что везут их из далекой Индии через посредников. Европейцы думают: «А давайте найдем свой, морской путь в Индию, чтобы не переплачивать!». Помогли и новые каравеллы, и приборы — компас и астролябия, и карты, которые стали точнее. А короли Испании и Португалии давали деньги на экспедиции, потому что соперничали друг с другом.",
4: "Христофор Колумб был уверен, что если поплыть на запад, можно быстрее добраться до богатой Азии. Он был отличным мореходом, но очень плохо считал. Он думал, что Земля гораздо меньше, чем есть на самом деле. В итоге, в 1492 году он приплыл не в Индию, а на Багамские острова, открыв для европейцев целый новый материк — Америку. Кстати, до конца своих дней он был уверен, что открыл именно Индию, поэтому местных жителей назвали «индейцами».",
5: "Португальцы искали путь в Индию вокруг Африки. Васко да Гама вышел из Лиссабона в 1497 году, обогнул Африку с юга, прошёл мыс Доброй Надежды, пересёк Индийский океан и в мае 1498 года дошёл до Каликута в Индии. Так был найден именно морской путь в Индию — та самая задача, ради которой всё начиналось. Плавание было тяжёлым: из четырёх кораблей домой вернулись только два, многие моряки погибли.",
6: "А вот Фернан Магеллан совершил настоящий подвиг. В 1519 году он отправился в плавание на 5 кораблях. Целью были Молуккские острова — Острова пряностей, а пролив за Южной Америкой был лишь дорогой к ним. Это был сущий ад: шторма, холод, бунты команды. Когда они наконец вышли в неизвестный океан, Магеллан назвал его Тихим, потому что за три месяца плавания они не встретили ни одного шторма. Но они не знали, что он огромный. Люди умирали от голода и цинги (болезни десен). Сам Магеллан погиб в 1521 году на Филиппинах и плавание не завершил. Экспедицию домой привёл Хуан Себастьян Элькано: в 1522 году вернулся лишь один корабль «Виктория» и 18 человек на нём. Но самое главное: они впервые обошли Землю и на практике подтвердили, что она круглая, а все океаны на планете соединены.",
7: "После открытия Америки начался процесс, который ученые называют «Колумбовым обменом». Это был грандиозный обмен растениями, животными, болезнями и даже людьми между Старым Светом (Европой) и Новым Светом (Америкой). Из Старого Света в Америку пошли пшеница и лошади, а обратно, в Европу, — картофель и томаты. Болезни, к сожалению, тоже «путешествовали», а люди перемещались между материками — кто добровольно, а кто насильно.",
8: "Из Америки в Европу приехали продукты, без которых мы сегодня не представляем свою жизнь. Это картофель, томаты (представьте, что до этого в Европе не было никакой пиццы и кетчупа!), кукуруза, какао (и шоколад!), подсолнечник, тыква и табак. Отдельно про картофель: сначала европейцы его не оценили и считали ядовитым. Есть история, которую часто рассказывают как легенду: будто его сначала выращивали для красоты на клумбах. Но потом поняли, что он очень сытный, и именно картошка стала спасением от голода во многих странах Европы.",
9: "Из Европы в Америку завезли пшеницу, рис, сахарный тростник, кофе и домашних животных. До прихода европейцев в Америке не было лошадей! Индейцы быстро научились на них ездить, и это полностью изменило их образ жизни и охоту.",
10: "Но у этого обмена была и очень трагичная, темная сторона. Вместе с товарами европейцы привезли в Америку болезни. У индейцев не было иммунитета к таким европейским заболеваниям, как оспа, корь и грипп. Для них это стало катастрофой. Оценки у историков сильно различаются: в разных регионах от болезней погибло от половины и, по самым высоким оценкам, до 90% коренного населения Америки.",
11: "Поскольку местных жителей для работы на плантациях катастрофически не хватало, европейцы начали массово завозить в Новый Свет африканских рабов. Так начался один из самых страшных периодов в истории — эпоха трансатлантической работорговли. Для Африки это тоже была катастрофа: континент терял миллионы людей, разрушались государства и обычная жизнь целых народов.",
12: "Великие географические открытия полностью перекроили мир. Во-первых, сместились торговые пути. Раньше главным центром торговли было Средиземное море (Италия, Венеция). Теперь же торговля переместилась в Атлантический океан. Испания, Португалия, Англия и Голландия стали новыми мировыми супердержавами.",
13: "За открытием Америки пошло её завоевание. Эрнан Кортес завоевал империю ацтеков: пала её столица Теночтитлан, а на её месте вырос город Мехико. Франсиско Писарро завоевал империю инков в Андах. Так началась эпоха колониализма: европейские страны поняли, что можно не просто торговать, а приплыть в чужую страну, объявить её своей и вывозить оттуда золото, серебро и драгоценности. Начался захват земель по всему миру.",
14: "В-третьих, это привело к революции цен. Огромное количество золота и серебра, хлынувшее из Америки в Европу, обесценило деньги и вызвало невероятную инфляцию. Это изменило экономику всей Европы. Цепочка простая: золота и серебра стало больше — деньги подешевели — цены выросли.",
15: "Но эпоха открытий подарила нам не только новые продукты и торговые пути. Она стала невероятным источником вдохновения для искусства. Представьте, как в Европе смотрели на карты, где были нарисованы неизвестные чудовища, и читали отчеты моряков о невиданных землях. Это же готовый сюжет для блокбастера!",
16: "В списке учебника на странице 38 есть фильм 1991 года «Одиссея капитана Блада». Это экранизация приключенческого романа Рафаэля Сабатини. Действие там происходит позже эпохи открытий, уже в XVII веке, на Карибском море. Поэтому это не сама эпоха открытий, а её отголосок: мир колоний, парусных флотов, пиратов и сокровищ появился именно благодаря открытиям XV–XVI веков. Художники тоже начали рисовать экзотические растения и животных, которых раньше никто не видел. А литература обогатилась жанром приключенческого романа. Мореплаватели стали настоящими супергероями своего времени.",
17: "Ну а теперь давайте вернемся к главному вопросу нашего учебника: почему Великие географические открытия — это переломный момент? Аргумент 1: Мир стал единым (Глобализация). Впервые в истории человечества появился единый мировой рынок. Люди, товары и идеи начали перемещаться по планете. До этого континенты развивались изолированно, а теперь они стали зависеть друг от друга.",
18: "Аргумент 2: Изменился сам человек и его быт («Колумбов обмен»). Благодаря обмену растениями и животными изменился наш рацион, а картофель во многих странах спасал людей от голода. Но самое главное — произошло глобальное смешение культур, народов и языков, из которого родились современные нации.",
19: "Аргумент 3: Научная революция. Что Земля — шар, учёные знали ещё с Античности: об этом писали Аристотель и Птолемей. Мореплаватели подтвердили это на практике: кругосветное плавание показало, что Землю можно обойти, а океаны соединены. Они разрушили старые мифы, создали точные карты и дали огромный толчок развитию географии, астрономии и других наук. Таким образом, Великие географические открытия — это тот самый момент, когда средневековый мир закончился, и начался мир Нового времени, в котором мы живем сегодня. Спасибо за внимание, я готов ответить на ваши вопросы!",
20: "Здесь все источники, которыми я пользовался. Учебник истории, страница 38. Большая российская энциклопедия — bigenc.ru. Портал «История.РФ» — histrf.ru. Книги из списка учебника с названиями: Гуляев — «По следам конкистадоров», Свет — «Колумб», Кинжалов и Белов — «Падение Теночтитлана», Кофман — «Тьерра аделанте!», Хаггард — «Дочь Монтесумы», Стефан Цвейг — «Подвиг Магеллана» и «Америго». Фильм — «Одиссея капитана Блада» 1991 года. Картинки — Wikimedia Commons, World History Encyclopedia и NCpedia.",
}

TX = 0.6; TW = 7.15
PX_R = 8.15
PX_L = 0.6; TX_R = 5.85; TW_R = 6.88

# ---- 1. Титул
s = prs.slides.add_slide(BLANK); slides.append(s)
bg(s, "b_title.jpg")
one(s, 0.6, 0.55, 10, 0.35, "ПРОЕКТ ПО ИСТОРИИ", size=14, color=YEL, bold=True)
one(s, 0.6, 1.05, 12.1, 1.6, "Великие географические\nоткрытия", size=44, color=WHITE, bold=True)
a = [
    (one(s, 0.6, 2.95, 12.0, 0.55, "Как Европа и Америка «нашли» друг друга — и мир стал единым",
         size=21, color=WHITE), "fade"),
    (fact(s, 0.6, 3.85, 3.9, 1.25, "1492", "год, с которого всё началось", fill=ORANGE), "zoom"),
]
one(s, 0.6, 5.65, 12.0, 0.5, "Платонов Артём Дмитриевич  •  7 «Д» класс", size=20, color=YEL, bold=True)
anim.append(a); note(s, NOTES[1])

# ---- 2. Главный вопрос (фото справа)
s = prs.slides.add_slide(BLANK); slides.append(s)
header(s, 2, "Главный вопрос учебника", "До XV века мир был как пазл,\nкоторый никто не собрал")
photo(s, "map.jpg", PX_R, 2.05, 4.55, 2.85, "Старый и Новый Свет на одной карте появились не сразу.")
a, yy = bullets(s, [
    ("Европейцы не знали,", "что существует Америка."),
    ("Индейцы не знали,", "что есть Европа, Азия и Африка."),
    ("Прямых связей", "между материками не было."),
    ("Открытия соединили", "части мира в один."),
], x=TX, w=TW, accent=ORANGE)
anim.append(a); note(s, NOTES[2])

# ---- 3. Причины (фото слева)
s = prs.slides.add_slide(BLANK); slides.append(s)
header(s, 3, "Часть 1 • Причины", "Европа искала свой морской\nпуть к пряностям")
photo(s, "caravels.jpg", PX_L, 2.05, 4.85, 3.05, "Каравеллы Колумба. Старинная открытка.")
a, yy = bullets(s, [
    ("1453 год:", "Константинополь пал, путь на восток стал труднее и дороже."),
    ("Пряности из Индии", "шли через посредников и стоили очень дорого."),
    ("Помогли каравеллы, компас,", "астролябия и более точные карты."),
    ("Короли Испании и Португалии", "давали деньги и соперничали друг с другом."),
], x=TX_R, w=TW_R, size=20, accent=ORANGE)
anim.append(a); note(s, NOTES[3])

# ---- 4. Колумб (фото справа)
s = prs.slides.add_slide(BLANK); slides.append(s)
header(s, 4, "Часть 1 • Путешественники", "Колумб ошибся в расчётах —\nи открыл Америку")
photo(s, "columbus.jpg", PX_R, 2.05, 4.55, 2.95, "Высадка Колумба. Картина Джона Вандерлина.")
a, yy = bullets(s, [
    ("Он был уверен:", "плыви на запад — и быстрее дойдёшь до Азии."),
    ("Ошибка в расчётах:", "Колумб сильно уменьшил размер Земли."),
    ("1492 год:", "вместо Индии — Багамские острова."),
    ("До конца жизни", "считал, что открыл Индию. Отсюда — «индейцы»."),
], x=TX, w=TW, accent=ORANGE)
anim.append(a); note(s, NOTES[4])

# ---- 5. Васко да Гама (фото слева)
s = prs.slides.add_slide(BLANK); slides.append(s)
header(s, 5, "Часть 1 • Путешественники", "Васко да Гама довёл Европу\nдо Индии морем")
photo(s, "gama.jpg", PX_L, 2.05, 4.85, 3.20, "Прибытие Васко да Гамы в Каликут, Индия.")
a, yy = bullets(s, [
    ("Португалия искала", "путь в Индию вокруг Африки."),
    ("1497 год:", "флот вышел из Лиссабона."),
    ("Он обогнул мыс Доброй Надежды", "и пересёк Индийский океан."),
    ("Май 1498 года: Каликут.", "Морской путь в Индию найден."),
], x=TX_R, w=TW_R, accent=ORANGE)
anim.append(a); note(s, NOTES[5])

# ---- 6. Магеллан и Элькано (фото справа)
s = prs.slides.add_slide(BLANK); slides.append(s)
header(s, 6, "Часть 1 • Путешественники", "Магеллан погиб, а Элькано\nдовёл корабль домой")
photo(s, "victoria.png", PX_R, 2.05, 4.55, 2.75, "Корабль «Виктория» на старинной карте.")
a, yy = bullets(s, [
    ("Цель — Молуккские острова.", "Пролив за Америкой — только дорога к ним."),
    ("1519 год:", "пять кораблей, штормы, бунты, голод и цинга."),
    ("Магеллан погиб в 1521 году", "на Филиппинах и плавание не завершил."),
], x=TX, w=TW, size=20, accent=ORANGE)
a.append((fact(s, TX, yy + 0.02, TW, 1.12, "18 человек", "вернулись с Элькано на «Виктории» в 1522 году", fill=ORANGE, big_size=34), "zoom"))
anim.append(a); note(s, NOTES[6])

# ---- 7. Колумбов обмен: схема с направлениями
s = prs.slides.add_slide(BLANK); slides.append(s)
header(s, 7, "Часть 2 • Колумбов обмен", "Два мира начали обмениваться всем", accent=GREEN, title_size=32)
left = rect(s, 0.6, 2.30, 3.95, 1.55, fill=LIGHT, line=BLUE2)
shape_text(left, [([("СТАРЫЙ СВЕТ", BLUE, True, 22, False)], PP_ALIGN.CENTER, 1.0, 2),
                  ([("Европа, Азия, Африка", DARK, False, 14.5, False)], PP_ALIGN.CENTER, 1.0, 0)])
right = rect(s, 8.78, 2.30, 3.95, 1.55, fill=RGBColor(0xFB,0xE8,0xDD), line=ORANGE)
shape_text(right, [([("НОВЫЙ СВЕТ", ORANGE, True, 22, False)], PP_ALIGN.CENTER, 1.0, 2),
                   ([("Америка", DARK, False, 14.5, False)], PP_ALIGN.CENTER, 1.0, 0)])
arrow(s, 4.85, 2.62, 3.65, MSO_SHAPE.RIGHT_ARROW, BLUE2)
arrow(s, 4.85, 3.28, 3.65, MSO_SHAPE.LEFT_ARROW, ORANGE)
p1 = rect(s, 0.6, 4.35, 5.95, 1.65, fill=LIGHT, line=BLUE2)
shape_text(p1, [([("Старый Свет → Новый Свет", BLUE, True, 18, False)], PP_ALIGN.CENTER, 1.0, 3),
                ([("пшеница, лошади, коровы;\nболезни — оспа и грипп", DARK, False, 15, False)], PP_ALIGN.CENTER, 1.08, 0)])
p2 = rect(s, 6.78, 4.35, 5.95, 1.65, fill=RGBColor(0xFB,0xE8,0xDD), line=ORANGE)
shape_text(p2, [([("Новый Свет → Старый Свет", ORANGE, True, 18, False)], PP_ALIGN.CENTER, 1.0, 3),
                ([("картофель, томаты, какао,\nкукуруза и табак", DARK, False, 15, False)], PP_ALIGN.CENTER, 1.08, 0)])
one(s, 0.6, 6.30, 12.13, 0.5, "Люди тоже перемещались между мирами: кто добровольно, а кто — насильно, в рабстве.",
    size=15.5, color=GREY, italic=True, align=PP_ALIGN.CENTER)
anim.append([(left, "fade"), (right, "fade"), (p1, "fly"), (p2, "fly")]); note(s, NOTES[7])

# ---- 8. Из Америки в Европу: группы + картофель
s = prs.slides.add_slide(BLANK); slides.append(s)
header(s, 8, "Часть 2 • Колумбов обмен", "Без Америки не было бы пиццы,\nшоколада и картошки", accent=GREEN)
rows = [
    ("Еда на каждый день", "картофель, томаты, кукуруза и тыква"),
    ("Вкус и сладости", "какао, а значит — шоколад"),
    ("Поля и привычки", "подсолнечник и табак"),
]
a = []
for i, (t_, sub) in enumerate(rows):
    a.append((tile(s, 0.6, 2.12 + i * 1.22, 12.13, 1.02, t_, sub, tcol=GREEN, tsize=20, ssize=15), "fly" if i % 2 == 0 else "fade"))
ban = rect(s, 0.6, 5.95, 12.13, 0.92, fill=YELLOW_BG, line=RGBColor(0xD9,0xB3,0x6A))
shape_text(ban, [([("Картофель: по легенде, его сначала сажали на клумбах, а потом он спасал Европу от голода.",
                    DARK, True, 16, False)], PP_ALIGN.CENTER, 1.08, 0)])
a.append((ban, "zoom"))
anim.append(a); note(s, NOTES[8])

# ---- 9. Из Европы в Америку
s = prs.slides.add_slide(BLANK); slides.append(s)
header(s, 9, "Часть 2 • Колумбов обмен", "До европейцев в Америке\nне было лошадей", accent=GREEN)
f9 = fact(s, PX_L, 2.15, 4.55, 2.4, "0 лошадей", "было в Америке\nдо прихода европейцев", fill=GREEN, big_size=36)
a, yy = bullets(s, [
    ("Растения:", "пшеница, рис, сахарный тростник и кофе."),
    ("Животные:", "лошади и другой домашний скот."),
    ("Индейцы быстро освоили лошадей:", "изменились охота и весь образ жизни."),
], x=TX_R, w=TW_R, accent=GREEN)
a.insert(0, (f9, "zoom"))
anim.append(a); note(s, NOTES[9])

# ---- 10. Болезни
s = prs.slides.add_slide(BLANK); slides.append(s)
header(s, 10, "Часть 2 • Тёмная сторона", "Болезни оказались страшнее оружия", accent=RED)
f = fact(s, 0.6, 2.15, 4.6, 2.9, "до 90%", "по оценкам историков —\nв разных регионах по-разному", fill=RED, big_size=40)
a, yy = bullets(s, [
    ("Европейцы привезли", "оспу, корь и грипп."),
    ("Иммунитета у индейцев не было:", "эти болезни были для них новыми."),
    ("Оценки разные:", "погибло от половины и до 90% населения."),
], x=5.65, w=7.08, accent=RED)
a.insert(0, (f, "zoom"))
anim.append(a); note(s, NOTES[10])

# ---- 11. Работорговля + Африка
s = prs.slides.add_slide(BLANK); slides.append(s)
header(s, 11, "Часть 2 • Тёмная сторона", "Нехватка рабочих привела\nк работорговле", accent=RED)
b1 = rect(s, 0.6, 2.35, 3.4, 1.45, fill=WHITE, line=RED)
shape_text(b1, [([("Америка", RED, True, 19, False)], PP_ALIGN.CENTER, 1.0, 2),
                ([("на плантациях работать некому", DARK, False, 14, False)], PP_ALIGN.CENTER, 1.05, 0)])
arrow(s, 4.15, 2.90, 0.75, MSO_SHAPE.RIGHT_ARROW, RED)
b2 = rect(s, 5.05, 2.35, 3.4, 1.45, fill=RGBColor(0xFB,0xE8,0xDD), line=RED)
shape_text(b2, [([("Африка", RED, True, 19, False)], PP_ALIGN.CENTER, 1.0, 2),
                ([("людей насильно везут через океан", DARK, False, 14, False)], PP_ALIGN.CENTER, 1.05, 0)])
arrow(s, 8.6, 2.90, 0.75, MSO_SHAPE.RIGHT_ARROW, RED)
b3 = rect(s, 9.5, 2.35, 3.23, 1.45, fill=LIGHT, line=BLUE2)
shape_text(b3, [([("Работорговля", BLUE, True, 17, False)], PP_ALIGN.CENTER, 1.0, 2),
                ([("через Атлантику", DARK, False, 13.5, False)], PP_ALIGN.CENTER, 1.05, 0)])
a, yy = bullets(s, [
    ("Местных жителей для плантаций", "катастрофически не хватало."),
    ("Европейцы начали массово завозить", "африканских рабов в Новый Свет."),
    ("Для Африки это была катастрофа:", "миллионы увезённых людей, разрушенные государства."),
], y=4.25, w=12.13, size=20, accent=RED)
anim.append([(b1, "fade"), (b2, "fade"), (b3, "fade")] + a); note(s, NOTES[11])

# ---- 12. Торговые пути: было / стало
s = prs.slides.add_slide(BLANK); slides.append(s)
header(s, 12, "Часть 3 • Последствия", "Торговля ушла из Средиземного моря\nв Атлантику")
c1 = rect(s, 0.6, 2.15, 5.85, 2.55, fill=WHITE, line=BORDER)
shape_text(c1, [([("БЫЛО:", GREY, True, 16, False)], PP_ALIGN.CENTER, 1.0, 4),
                ([("Средиземное море", BLUE, True, 21, False)], PP_ALIGN.CENTER, 1.0, 2),
                ([("центр торговли — Венеция и Италия", DARK, False, 15, False)], PP_ALIGN.CENTER, 1.05, 0)])
c2 = rect(s, 6.88, 2.15, 5.85, 2.55, fill=LIGHT, line=BLUE2)
shape_text(c2, [([("СТАЛО:", ORANGE, True, 16, False)], PP_ALIGN.CENTER, 1.0, 4),
                ([("Атлантический океан", BLUE, True, 21, False)], PP_ALIGN.CENTER, 1.0, 2),
                ([("лидеры: Испания, Португалия,\nАнглия и Голландия", DARK, False, 15, False)], PP_ALIGN.CENTER, 1.05, 0)])
a, yy = bullets(s, [("Главные торговые дороги мира", "прошли по океану, а не по суше.")],
                y=5.05, w=12.13)
anim.append([(c1, "fade"), (c2, "fly")] + a); note(s, NOTES[12])

# ---- 13. Конкистадоры (фото справа)
s = prs.slides.add_slide(BLANK); slides.append(s)
header(s, 13, "Часть 3 • Последствия", "Кортес и Писарро завоевали\nимперии ацтеков и инков", title_size=30)
photo(s, "tenochtitlan.png", PX_R, 2.05, 4.55, 3.0, "Завоевание Теночтитлана. Картина XVII века.")
a, yy = bullets(s, [
    ("Эрнан Кортес", "завоевал империю ацтеков и Теночтитлан."),
    ("Франсиско Писарро", "завоевал империю инков в Андах."),
    ("Дальше — колониализм:", "чужие земли объявляли своими и вывозили золото и серебро."),
], x=TX, w=TW, accent=ORANGE)
anim.append(a); note(s, NOTES[13])

# ---- 14. Революция цен: цепочка
s = prs.slides.add_slide(BLANK); slides.append(s)
header(s, 14, "Часть 3 • Последствия", "Золото из Америки обесценило\nденьги в Европе")
a, yy = bullets(s, [
    ("Из Америки в Европу", "хлынули золото и серебро."),
    ("Денег стало больше,", "и они начали дешеветь."),
    ("Цены выросли.", "Это назвали революцией цен."),
], x=TX, w=6.55, accent=ORANGE)
chain = []
steps_chain = [("Золота и серебра стало больше", BLUE2), ("Деньги подешевели", BLUE), ("Цены выросли", ORANGE)]
cy = 2.05
for i, (lab, col) in enumerate(steps_chain):
    b = rect(s, 7.45, cy, 5.28, 0.98, fill=col, line=None)
    shape_text(b, [([(lab, WHITE, True, 17.5, False)], PP_ALIGN.CENTER, 1.05, 0)])
    chain.append((b, "fade"))
    if i < 2:
        arrow(s, 9.85, cy + 1.02, 0.5, MSO_SHAPE.DOWN_ARROW, GREY, h=0.34)
    cy += 1.42
anim.append(a + chain); note(s, NOTES[14])

# ---- 15. Вдохновение (фото слева)
s = prs.slides.add_slide(BLANK); slides.append(s)
header(s, 15, "Часть 3 • Культура", "Неизвестные земли стали\nготовыми сюжетами")
photo(s, "map.jpg", PX_L, 2.05, 4.85, 2.7, "Карта мира после Великих открытий.")
a, yy = bullets(s, [
    ("На старых картах", "рисовали неизвестных чудовищ."),
    ("Отчёты моряков", "читались как приключения."),
    ("Невиданные земли и животные", "вдохновляли художников и писателей."),
], x=TX_R, w=TW_R, accent=ORANGE)
anim.append(a); note(s, NOTES[15])

# ---- 16. Капитан Блад: фильм, а не сама эпоха
s = prs.slides.add_slide(BLANK); slides.append(s)
header(s, 16, "Часть 3 • Культура", "«Одиссея капитана Блада» (1991) —\nотголосок эпохи парусов", title_size=29)
a, yy = bullets(s, [
    ("Это фильм по роману", "Рафаэля Сабатини о капитане Бладе."),
    ("Действие — XVII век,", "Карибское море и пираты. Это позже открытий."),
    ("Но мир колоний и парусных флотов", "создали именно открытия XV–XVI веков."),
], x=TX, w=TW, accent=ORANGE)
f = fact(s, PX_R, 2.15, 4.55, 2.4, "1991", "«Одиссея капитана Блада» —\nфильм из списка учебника (с. 38)", fill=BLUE)
anim.append(a + [(f, "zoom")]); note(s, NOTES[16])

# ---- 17. Аргумент 1
s = prs.slides.add_slide(BLANK); slides.append(s)
header(s, 17, "Вывод", "Мир стал единым:\nначалась глобализация")
a, yy = bullets(s, [
    ("Впервые появился", "единый мировой рынок."),
    ("Люди, товары и идеи", "начали перемещаться по всей планете."),
    ("Континенты перестали жить отдельно", "и стали зависеть друг от друга."),
], x=TX, w=TW, accent=BLUE)
f = fact(s, PX_R, 2.15, 4.55, 2.4, "1", "мировой рынок\nсвязал материки", fill=BLUE)
a.append((f, "zoom"))
anim.append(a); note(s, NOTES[17])

# ---- 18. Аргумент 2
s = prs.slides.add_slide(BLANK); slides.append(s)
header(s, 18, "Вывод", "Изменилась еда, быт\nи сами люди", accent=GREEN)
f = fact(s, PX_L, 2.15, 4.55, 2.4, "2", "еда, культуры и языки\nсмешались", fill=GREEN)
a, yy = bullets(s, [
    ("Изменился рацион:", "новые растения спасли от голода — вспомните картошку."),
    ("Смешались культуры,", "народы и языки."),
    ("Из этого смешения", "родились современные нации."),
], x=TX_R, w=TW_R, accent=GREEN)
a.insert(0, (f, "zoom"))
anim.append(a); note(s, NOTES[18])

# ---- 19. Аргумент 3 + финал
s = prs.slides.add_slide(BLANK); slides.append(s)
header(s, 19, "Вывод", "Подтвердили на практике то,\nо чём догадывались учёные", title_size=30)
a, yy = bullets(s, [
    ("Ещё Аристотель и Птолемей", "писали, что Земля — шар."),
    ("Кругосветное плавание это подтвердило:", "Землю можно обойти, океаны соединены."),
    ("Точные карты дали толчок", "географии и астрономии."),
    ("Средневековый мир закончился.", "Началось Новое время."),
], x=TX, w=TW, size=20, accent=BLUE)
f = fact(s, PX_R, 2.15, 4.55, 2.15, "3", "наука и точные карты", fill=BLUE)
a.append((f, "zoom"))
thanks = one(s, 0.6, 6.55, 12.13, 0.5, "Спасибо за внимание! Готов ответить на вопросы.",
             size=20, color=ORANGE, bold=True, align=PP_ALIGN.CENTER)
a.append((thanks, "zoom"))
anim.append(a); note(s, NOTES[19])

# ---- 20. Источники с названиями книг
s = prs.slides.add_slide(BLANK); slides.append(s)
header(s, 20, "Финал", "Источники", title_size=32)
one(s, 0.6, 1.95, 5.9, 0.4, "Материалы и ссылки", size=17, color=BLUE, bold=True)
one(s, 6.88, 1.95, 5.85, 0.4, "Книги, фильм и картинки", size=17, color=BLUE, bold=True)
left = one(s, 0.6, 2.42, 5.9, 4.4,
           "• Учебник истории, с. 38 — темы проектов, список книг и фильм\n\n"
           "• Большая российская энциклопедия — bigenc.ru\n\n"
           "• Портал «История.РФ» — histrf.ru",
           size=15.5)
right = one(s, 6.88, 2.42, 5.85, 4.4,
            "• Гуляев — «По следам конкистадоров»\n"
            "• Свет — «Колумб»\n"
            "• Кинжалов и Белов — «Падение Теночтитлана»\n"
            "• Кофман — «Тьерра аделанте!»\n"
            "• Хаггард — «Дочь Монтесумы»\n"
            "• Стефан Цвейг — «Подвиг Магеллана», «Америго»\n"
            "• Фильм: «Одиссея капитана Блада» (1991)\n"
            "• Картинки: commons.wikimedia.org,\n  worldhistory.org, ncpedia.org",
            size=14.5)
anim.append([(left, "fade"), (right, "fade")]); note(s, NOTES[20])

# ---- переходы и анимации
trans_kinds = ["fade", "push", "wipe", "fade", "cover"] * 4
for sl, a, tk in zip(slides, anim, trans_kinds):
    set_transition(sl, tk)
    if a:
        add_animations(sl, a)

prs.save(OUT)
print("Готово:", OUT, "| слайдов:", len(slides))
