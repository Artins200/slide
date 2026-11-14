#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Презентация v6: текст из нового варианта, 1 слайд = 1 мысль, заголовок-вывод, фото слева/справа по очереди.
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
# ================= оформление v6 =================
# Правило: 1 слайд = 1 мысль. Заголовок = вывод. 3-4 тезиса, крупный шрифт.
# Фото чередуется: справа / слева. Схемы вместо таблиц.

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
               [([(text_, color, bold, size, italic)], align, 1.08, 0)],
               align=align, anchor=anchor)

def shape_text(shp, paras, align=PP_ALIGN.CENTER):
    _fill_tf(shp.text_frame, paras, align=align, anchor=MSO_ANCHOR.MIDDLE)
    return shp

def header(slide, n, kicker, title, accent=ORANGE, title_size=29):
    bg(slide)
    one(slide, 0.6, 0.30, 12.1, 0.3, kicker.upper(), size=11, color=accent, bold=True)
    one(slide, 0.6, 0.60, 12.13, 1.05, title, size=title_size, color=BLUE, bold=True)
    ln = rect(slide, 0.62, 1.70, 1.55, 0.055, fill=accent, line=None, round_=False)
    one(slide, 0.6, 7.16, 8, 0.25, "Артём Платонов  •  7 «Д»", size=8.5, color=GREY)
    one(slide, 12.0, 7.16, 0.73, 0.25, f"{n} / 20", size=8.5, color=BLUE, bold=True, align=PP_ALIGN.RIGHT)
    return 1.98

def photo(slide, img, x, y, w, h, caption):
    rect(slide, x - 0.07, y - 0.07, w + 0.14, h + 0.14, fill=WHITE, line=BORDER, round_=False)
    pic = slide.shapes.add_picture(A(img), Inches(x), Inches(y), Inches(w), Inches(h))
    one(slide, x - 0.07, y + h + 0.10, w + 0.14, 0.55, caption, size=9,
        color=GREY, italic=True, align=PP_ALIGN.CENTER)
    return pic

def bullets(slide, items, x=0.6, y=1.98, w=7.0, size=19, accent=BLUE):
    """Каждый тезис — отдельный щелчок. lead = жирное начало."""
    out = []; yy = y
    for i, (lead, rest) in enumerate(items):
        full = (lead + " " + rest).strip()
        chars = max(28, int(w * 112 / size))
        lines = max(1, math.ceil(len(full) / chars))
        h = lines * (size / 48.0) + 0.14
        runs = [("▸  ", accent, True, size, False)]
        if lead: runs.append((lead + "  ", BLUE, True, size, False))
        if rest: runs.append((rest, DARK, False, size, False))
        sh = box(slide, x, yy, w, h, [(runs, PP_ALIGN.LEFT, 1.07, 0)])
        out.append((sh, "fly" if i % 2 == 0 else "fade"))
        yy += h + 0.26
    return out, yy

def fact(slide, x, y, w, h, big, small, fill=BLUE):
    shp = rect(slide, x, y, w, h, fill=fill, line=None)
    shape_text(shp, [
        ([(big, WHITE, True, 38, False)], PP_ALIGN.CENTER, 1.0, 2),
        ([(small, WHITE, False, 13.5, False)], PP_ALIGN.CENTER, 1.05, 0),
    ])
    return shp

def tile(slide, x, y, w, h, title_, sub, fill=WHITE, tcol=BLUE):
    shp = rect(slide, x, y, w, h, fill=fill, line=BORDER)
    shape_text(shp, [
        ([(title_, tcol, True, 19, False)], PP_ALIGN.CENTER, 1.0, 2),
        ([(sub, DARK, False, 12.5, False)], PP_ALIGN.CENTER, 1.05, 0),
    ])
    return shp

def arrow(slide, x, y, w=0.85, kind=MSO_SHAPE.RIGHT_ARROW, fill=BLUE2):
    ar = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(0.38))
    ar.fill.solid(); ar.fill.fore_color.rgb = fill; ar.line.fill.background()
    ar.shadow.inherit_shadow = False
    return ar

def note(slide, text_):
    slide.notes_slide.placeholders[1].text = text_

slides = []; anim = []

# ---------- речь: новый текст, разрезанный по слайдам (без добавок от себя) ----------
NOTES = {
1: "Здравствуйте, ребята и уважаемая [Имя Отчество]! Сегодня мы с вами отправимся в путешествие, которое навсегда изменило нашу планету. Мы поговорим об эпохе Великих географических открытий. И речь пойдет не просто о скучных датах и именах. Мы узнаем, как Европа и Америка «нашли» друг друга, почему картошка стала спасением для миллионов людей, как родилась глобализация, и почему приключенческие романы и фильмы до сих пор так популярны. Пристегните ремни, мы отправляемся в XVI век!",
2: "В самом начале нашей главы учебник ставит главный вопрос: «Почему Великие географические открытия считаются переломным моментом в истории?» Вдумайтесь: до XV века люди жили, будто в разных параллельных мирах. Европейцы не знали, что есть Америка. Американские индейцы не знали, что есть Европа, Азия и Африка. Китайцы жили своей жизнью за Великой стеной. Мир был огромным пазлом, кусочки которого не были соединены. И вот, благодаря отчаянным мореплавателям, этот пазл сложился. В конце своего рассказа я четко сформулирую три аргумента, почему же этот момент стал поворотным для всего человечества.",
3: "Итак, XV век. Европа сходит с ума по пряностям. Перец, корица, гвоздика стоят бешеных денег, потому что везут их из далекой Индии через арабских посредников. Европейцы думают: «А давайте найдем свой, морской путь в Индию, чтобы не переплачивать!».",
4: "Христофор Колумб был уверен, что если поплыть на запад, можно быстрее добраться до богатой Азии. Он был отличным мореходом, но очень плохо считал. Он думал, что Земля гораздо меньше, чем есть на самом деле. В итоге, в 1492 году он приплыл не в Индию, а на Багамские острова, открыв для европейцев целый новый материк — Америку. Кстати, до конца своих дней он был уверен, что открыл именно Индию, поэтому местных жителей назвали «индейцами».",
5: "А вот Фернан Магеллан совершил настоящий подвиг. В 1519 году он отправился в плавание на 5 кораблях. Они искали пролив, чтобы обогнуть Южную Америку. Это был сущий ад: шторма, холод, бунты команды. Когда они наконец вышли в неизвестный океан, Магеллан назвал его Тихим, потому что за три месяца плавания они не встретили ни одного шторма. Но они не знали, что он огромный. Люди умирали от голода и цинги (болезни десен). Из всей экспедиции выжило лишь 18 человек на одном корабле! Но самое главное: они доказали, что Земля круглая, и что все океаны на планете соединены.",
6: "После открытия Америки начался процесс, который ученые называют «Колумбовым обменом». Это был грандиозный обмен растениями, животными, болезнями и даже людьми между Старым Светом (Европой) и Новым Светом (Америкой).",
7: "Из Америки в Европу приехали продукты, без которых мы сегодня не представляем свою жизнь. Это картофель, томаты (представьте, что до этого в Европе не было никакой пиццы и кетчупа!), кукуруза, какао (и шоколад!), подсолнечник, тыква и табак.",
8: "Кстати, картофель сначала европейцы не оценили. Они считали его ядовитым и выращивали только для красоты на клумбах. Но потом поняли, что он очень сытный, и именно картошка спасла Европу от страшного голода в последующие века.",
9: "Из Европы в Америку завезли пшеницу, рис, сахарный тростник, кофе и домашних животных. До прихода европейцев в Америке не было лошадей! Индейцы быстро научились на них ездить, и это полностью изменило их образ жизни и охоту.",
10: "Но у этого обмена была и очень трагичная, темная сторона. Вместе с товарами европейцы привезли в Америку болезни. У индейцев не было иммунитета к таким европейским заболеваниям, как оспа, корь и грипп. Для них это стало катастрофой. По разным оценкам, от болезней погибло до 90% коренного населения Америки.",
11: "Поскольку местных жителей для работы на плантациях катастрофически не хватало, европейцы начали массово завозить в Новый Свет африканских рабов. Так начался один из самых страшных периодов в истории — эпоха трансатлантической работорговли.",
12: "Великие географические открытия полностью перекроили мир. Во-первых, сместились торговые пути. Раньше главным центром торговли было Средиземное море (Италия, Венеция). Теперь же торговля переместилась в Атлантический океан. Испания, Португалия, Англия и Голландия стали новыми мировыми супердержавами.",
13: "Во-вторых, началась эпоха колониализма. Европейские страны поняли, что можно не просто торговать, а приплыть в чужую страну, объявить её своей и вывозить оттуда золото, серебро и драгоценности. Начался захват земель по всему миру.",
14: "В-третьих, это привело к революции цен. Огромное количество золота и серебра, хлынувшее из Америки в Европу, обесценило деньги и вызвало невероятную инфляцию. Это изменило экономику всей Европы.",
15: "Но эпоха открытий подарила нам не только новые продукты и торговые пути. Она стала невероятным источником вдохновения для искусства. Представьте, как в Европе смотрели на карты, где были нарисованы неизвестные чудовища, и читали отчеты моряков о невиданных землях. Это же готовый сюжет для блокбастера!",
16: "Вспомните хотя бы роман и замечательный фильм 1991 года «Одиссея капитана Блада». История про благородного пирата и морские сражения — это прямое эхо той самой эпохи Великих географических открытий, когда море было зоной риска, приключений и несметных сокровищ. Художники тоже начали рисовать экзотические растения и животных, которых раньше никто не видел. А литература обогатилась жанром приключенческого романа. Мореплаватели стали настоящими супергероями своего времени.",
17: "Ну а теперь давайте вернемся к главному вопросу нашего учебника: почему Великие географические открытия — это переломный момент? Аргумент 1: Мир стал единым (Глобализация). Впервые в истории человечества появился единый мировой рынок. Люди, товары и идеи начали перемещаться по планете. До этого континенты развивались изолированно, а теперь они стали зависеть друг от друга.",
18: "Аргумент 2: Изменился сам человек и его быт («Колумбов обмен»). Благодаря обмену растениями и животными изменился наш рацион, мы победили голод (вспомните картошку). Но самое главное — произошло глобальное смешение культур, народов и языков, из которого родились современные нации.",
19: "Аргумент 3: Научная революция. Мореплаватели доказали, что Земля — шар. Они разрушили старые мифы, создали точные карты и дали огромный толчок развитию географии, астрономии и других наук. Таким образом, Великие географические открытия — это тот самый момент, когда средневековый мир закончился, и начался мир Нового времени, в котором мы живем сегодня. Спасибо за внимание, я готов ответить на ваши вопросы!",
20: "Здесь все источники, которыми я пользовался: учебник (страница 38), Большая российская энциклопедия, портал «История.РФ», Wikimedia Commons, Всемирная историческая энциклопедия и NCpedia. Книги — из списка учебника на странице 38, фильм — «Одиссея капитана Блада» 1991 года.",
}

TX = 0.6; TW = 7.15          # текст слева, фото справа
PX_R = 8.15                  # фото справа: x
PX_L = 0.6; TX_R = 5.85; TW_R = 6.88   # фото слева, текст справа

# ---- 1. Титул
s = prs.slides.add_slide(BLANK); slides.append(s)
bg(s, "b_title.jpg")
one(s, 0.6, 0.55, 10, 0.35, "ПРОЕКТ ПО ИСТОРИИ", size=13, color=YEL, bold=True)
one(s, 0.6, 1.05, 12.1, 1.6, "Великие географические\nоткрытия", size=42, color=WHITE, bold=True)
a = [
    (one(s, 0.6, 2.95, 12.0, 0.55, "Как Европа и Америка «нашли» друг друга — и мир стал единым",
         size=20, color=WHITE), "fade"),
    (fact(s, 0.6, 3.85, 3.9, 1.25, "1492", "год, с которого всё началось", fill=ORANGE), "zoom"),
]
one(s, 0.6, 5.65, 12.0, 0.5, "Артём Платонов Дмитриевич  •  7 «Д» класс", size=19, color=YEL, bold=True)
anim.append(a); note(s, NOTES[1])

# ---- 2. Главный вопрос (фото справа)
s = prs.slides.add_slide(BLANK); slides.append(s)
header(s, 2, "Главный вопрос учебника", "До XV века мир был как пазл,\nкоторый никто не собрал")
photo(s, "map.jpg", PX_R, 2.05, 4.55, 2.85, "Старый и Новый Свет на одной карте появились не сразу.")
a, yy = bullets(s, [
    ("Европейцы не знали,", "что существует Америка."),
    ("Индейцы не знали,", "что есть Европа, Азия и Африка."),
    ("Каждый мир жил сам по себе.", "Связи между материками не было."),
    ("В конце — три аргумента,", "почему открытия стали переломным моментом."),
], x=TX, w=TW, accent=ORANGE)
anim.append(a); note(s, NOTES[2])

# ---- 3. Причина: пряности (фото слева)
s = prs.slides.add_slide(BLANK); slides.append(s)
header(s, 3, "Часть 1 • Причины", "Европа искала свой морской\nпуть к пряностям")
photo(s, "caravels.jpg", PX_L, 2.05, 4.85, 3.05, "Каравеллы Колумба. Старинная открытка.")
a, yy = bullets(s, [
    ("Перец, корица, гвоздика", "стоили бешеных денег."),
    ("Их везли из Индии", "через арабских посредников — с переплатой."),
    ("Вывод европейцев:", "нужен свой морской путь в Индию."),
], x=TX_R, w=TW_R, accent=ORANGE)
a.append((fact(s, TX_R, yy + 0.05, TW_R, 1.15, "Цель — Индия", "а получилось — открытие нового мира", fill=BLUE), "zoom"))
anim.append(a); note(s, NOTES[3])

# ---- 4. Колумб (фото справа)
s = prs.slides.add_slide(BLANK); slides.append(s)
header(s, 4, "Часть 1 • Путешественники", "Колумб ошибся в расчётах —\nи открыл Америку")
photo(s, "columbus.jpg", PX_R, 2.05, 4.55, 2.95, "Высадка Колумба. Картина Джона Вандерлина.")
a, yy = bullets(s, [
    ("Он был уверен:", "плыви на запад — и быстрее дойдёшь до Азии."),
    ("Ошибка в цифрах:", "Колумб сильно уменьшил размер Земли."),
    ("1492 год:", "вместо Индии — Багамские острова."),
    ("До конца жизни", "считал, что открыл Индию. Отсюда — «индейцы»."),
], x=TX, w=TW, accent=ORANGE)
anim.append(a); note(s, NOTES[4])

# ---- 5. Магеллан (фото слева)
s = prs.slides.add_slide(BLANK); slides.append(s)
header(s, 5, "Часть 1 • Путешественники", "Вернулся один корабль из пяти —\nно Земля оказалась круглой")
photo(s, "victoria.png", PX_L, 2.05, 4.85, 2.75, "Корабль «Виктория» на старинной карте.")
a, yy = bullets(s, [
    ("1519 год:", "пять кораблей ушли искать пролив за Америкой."),
    ("Дальше — шторма, холод и бунты.", "А потом три месяца без единого шторма — Тихий океан."),
    ("Люди умирали", "от голода и цинги."),
], x=TX_R, w=TW_R, accent=ORANGE)
a.append((fact(s, TX_R, yy + 0.05, TW_R, 1.2, "18 человек", "выжили и доказали: океаны соединены, Земля — шар", fill=ORANGE), "zoom"))
anim.append(a); note(s, NOTES[5])

# ---- 6. Колумбов обмен: схема
s = prs.slides.add_slide(BLANK); slides.append(s)
header(s, 6, "Часть 2 • Колумбов обмен", "Два мира начали обмениваться всем", accent=GREEN, title_size=30)
left = rect(s, 0.6, 2.35, 4.15, 1.7, fill=LIGHT, line=BLUE2)
shape_text(left, [([( "СТАРЫЙ СВЕТ", BLUE, True, 21, False)], PP_ALIGN.CENTER, 1.0, 2),
                  ([("Европа, Азия, Африка", DARK, False, 14, False)], PP_ALIGN.CENTER, 1.0, 0)])
right = rect(s, 8.58, 2.35, 4.15, 1.7, fill=RGBColor(0xFB,0xE8,0xDD), line=ORANGE)
shape_text(right, [([("НОВЫЙ СВЕТ", ORANGE, True, 21, False)], PP_ALIGN.CENTER, 1.0, 2),
                   ([("Америка", DARK, False, 14, False)], PP_ALIGN.CENTER, 1.0, 0)])
arrow(s, 5.05, 2.70, 3.25, MSO_SHAPE.RIGHT_ARROW, BLUE2)
arrow(s, 5.05, 3.35, 3.25, MSO_SHAPE.LEFT_ARROW, ORANGE)
a = [(left, "fade"), (right, "fade")]
labels = ["растения", "животные", "болезни", "люди и культуры"]
for i, lab in enumerate(labels):
    t = tile(s, 0.6 + i * 3.13, 4.65, 2.85, 1.05, lab, "", fill=WHITE, tcol=GREEN)
    a.append((t, "fly"))
one(s, 0.6, 6.05, 12.13, 0.5, "Это и есть «Колумбов обмен» — обмен между Старым и Новым Светом после 1492 года.",
    size=16, color=GREY, italic=True, align=PP_ALIGN.CENTER)
anim.append(a); note(s, NOTES[6])

# ---- 7. Из Америки в Европу: плитки
s = prs.slides.add_slide(BLANK); slides.append(s)
header(s, 7, "Часть 2 • Колумбов обмен", "Без Америки не было бы пиццы,\nшоколада и картошки", accent=GREEN)
plants = [("Картофель", "главная еда Европы потом"), ("Томаты", "нет томатов — нет пиццы"),
          ("Кукуруза", "новая культура полей"), ("Какао", "а значит — и шоколад"),
          ("Подсолнечник", "и тыква — тоже оттуда"), ("Табак", "тоже приехал из Америки")]
a = []
for i, (t_, sub) in enumerate(plants):
    x = 0.6 + (i % 3) * 4.14; y = 2.15 + (i // 3) * 1.75
    a.append((tile(s, x, y, 3.85, 1.45, t_, sub, tcol=GREEN), "fly" if i % 2 == 0 else "fade"))
one(s, 0.6, 5.95, 12.13, 0.5, "Всё это в Европе появилось только после открытия Америки.",
    size=16, color=GREY, italic=True, align=PP_ALIGN.CENTER)
anim.append(a); note(s, NOTES[7])

# ---- 8. Картофель (текст слева, факт справа)
s = prs.slides.add_slide(BLANK); slides.append(s)
header(s, 8, "Часть 2 • Отдельная история", "Картошку сажали на клумбах —\nа потом она спасла Европу", accent=GREEN)
a, yy = bullets(s, [
    ("Сначала не оценили:", "считали картофель ядовитым."),
    ("Сажали для красоты,", "как цветы на клумбах."),
    ("Потом поняли главное:", "он сытный и растёт почти везде."),
], x=TX, w=TW, accent=GREEN)
f = fact(s, PX_R, 2.15, 4.55, 2.4, "Спасение\nот голода", "картофель выручал Европу\nв тяжёлые века", fill=GREEN)
# fact() ставит крупный размер; для многострочного большого текста это ок
a.append((f, "zoom"))
anim.append(a); note(s, NOTES[8])

# ---- 9. Из Европы в Америку (фото слева)
s = prs.slides.add_slide(BLANK); slides.append(s)
header(s, 9, "Часть 2 • Колумбов обмен", "В Америке не было лошадей —\nих привезли европейцы", accent=GREEN)
f9 = fact(s, PX_L, 2.15, 4.55, 2.4, "Лошадей\nне было", "до прихода европейцев\nв Америке", fill=GREEN)
a, yy = bullets(s, [
    ("Растения:", "пшеница, рис, сахарный тростник, кофе."),
    ("Животные:", "лошади и другой домашний скот."),
    ("Индейцы быстро освоили лошадей.", "Изменились охота и весь образ жизни."),
], x=TX_R, w=TW_R, accent=GREEN)
a.insert(0, (f9, "zoom"))
anim.append(a); note(s, NOTES[9])

# ---- 10. Болезни (большая цифра)
s = prs.slides.add_slide(BLANK); slides.append(s)
header(s, 10, "Часть 2 • Тёмная сторона", "Болезни оказались страшнее оружия", accent=RED)
f = fact(s, 0.6, 2.15, 4.6, 2.9, "до 90%", "коренного населения Америки\nпогибло от болезней", fill=RED)
a, yy = bullets(s, [
    ("Европейцы привезли", "оспу, корь и грипп."),
    ("Иммунитета у индейцев не было.", "Для них эти болезни были новыми."),
    ("Итог — катастрофа", "для целых народов Америки."),
], x=5.65, w=7.08, accent=RED)
a.insert(0, (f, "zoom"))
anim.append(a); note(s, NOTES[10])

# ---- 11. Работорговля (схема)
s = prs.slides.add_slide(BLANK); slides.append(s)
header(s, 11, "Часть 2 • Тёмная сторона", "Нехватка рабочих привела\nк работорговле", accent=RED)
b1 = rect(s, 0.6, 2.5, 3.4, 1.5, fill=WHITE, line=RED)
shape_text(b1, [([("Плантации Америки", RED, True, 18, False)], PP_ALIGN.CENTER, 1.0, 2),
                ([("работать некому", DARK, False, 13.5, False)], PP_ALIGN.CENTER, 1.0, 0)])
arrow(s, 4.15, 3.05, 0.75, MSO_SHAPE.RIGHT_ARROW, RED)
b2 = rect(s, 5.05, 2.5, 3.4, 1.5, fill=RGBColor(0xFB,0xE8,0xDD), line=RED)
shape_text(b2, [([("Африка", RED, True, 18, False)], PP_ALIGN.CENTER, 1.0, 2),
                ([("людей везут через океан насильно", DARK, False, 13.5, False)], PP_ALIGN.CENTER, 1.0, 0)])
arrow(s, 8.6, 3.05, 0.75, MSO_SHAPE.RIGHT_ARROW, RED)
b3 = rect(s, 9.5, 2.5, 3.23, 1.5, fill=LIGHT, line=BLUE2)
shape_text(b3, [([("Работорговля", BLUE, True, 16, False)], PP_ALIGN.CENTER, 1.0, 2),
                ([("через Атлантику", DARK, False, 12.5, False)], PP_ALIGN.CENTER, 1.0, 0)])
a, yy = bullets(s, [
    ("Местных жителей для плантаций", "катастрофически не хватало."),
    ("Европейцы начали массово завозить", "африканских рабов в Новый Свет."),
], y=4.55, w=12.13, accent=RED)
anim.append([(b1, "fade"), (b2, "fade"), (b3, "fade")] + a); note(s, NOTES[11])

# ---- 12. Торговые пути: было / стало
s = prs.slides.add_slide(BLANK); slides.append(s)
header(s, 12, "Часть 3 • Последствия", "Торговля ушла из Средиземного моря\nв Атлантику")
c1 = rect(s, 0.6, 2.2, 5.85, 2.6, fill=WHITE, line=BORDER)
shape_text(c1, [([("БЫЛО", GREY, True, 15, False)], PP_ALIGN.CENTER, 1.0, 4),
                ([("Средиземное море", BLUE, True, 20, False)], PP_ALIGN.CENTER, 1.0, 2),
                ([("центр торговли — Венеция и Италия", DARK, False, 14, False)], PP_ALIGN.CENTER, 1.05, 0)],
           align=PP_ALIGN.CENTER)
c2 = rect(s, 6.88, 2.2, 5.85, 2.6, fill=LIGHT, line=BLUE2)
shape_text(c2, [([("СТАЛО", ORANGE, True, 15, False)], PP_ALIGN.CENTER, 1.0, 4),
                ([("Атлантический океан", BLUE, True, 20, False)], PP_ALIGN.CENTER, 1.0, 2),
                ([("лидеры: Испания, Португалия,\nАнглия и Голландия", DARK, False, 14, False)], PP_ALIGN.CENTER, 1.05, 0)],
           align=PP_ALIGN.CENTER)
a, yy = bullets(s, [("Главные торговые дороги мира", "прошли по океану, а не по суше и Средиземному морю.")],
                y=5.15, w=12.13)
anim.append([(c1, "fade"), (c2, "fly")] + a); note(s, NOTES[12])

# ---- 13. Колониализм (фото справа)
s = prs.slides.add_slide(BLANK); slides.append(s)
header(s, 13, "Часть 3 • Последствия", "Европа начала захватывать\nчужие земли")
photo(s, "tenochtitlan.png", PX_R, 2.05, 4.55, 3.0, "Завоевание Теночтитлана. Картина XVII века.")
a, yy = bullets(s, [
    ("Логика изменилась:", "не торговать, а объявить землю своей."),
    ("Из колоний вывозили", "золото, серебро и драгоценности."),
    ("Так началась эпоха колониализма", "— захват земель по всему миру."),
], x=TX, w=TW, accent=ORANGE)
anim.append(a); note(s, NOTES[13])

# ---- 14. Революция цен (график bars)
s = prs.slides.add_slide(BLANK); slides.append(s)
header(s, 14, "Часть 3 • Последствия", "Золото из Америки обесценило\nденьги в Европе")
a, yy = bullets(s, [
    ("В Европу хлынули", "золото и серебро из Америки."),
    ("Денег стало много —", "и они начали дешеветь."),
    ("Цены резко выросли.", "Это назвали революцией цен."),
], x=TX, w=6.6, accent=ORANGE)
# простой график: три растущих столбца
bars = []
for i, (bh, lab) in enumerate([(0.9, "было"), (1.7, "стало больше"), (2.7, "цены вверх")]):
    b = rect(s, 7.75 + i * 1.65, 5.15 - bh, 1.25, bh, fill=ORANGE if i == 2 else BLUE2, line=None, round_=False)
    shape_text(b, [([(lab, WHITE, True, 11, False)], PP_ALIGN.CENTER, 1.0, 0)])
    bars.append((b, "wipe"))
one(s, 7.55, 5.35, 5.2, 0.4, "Как росли цены в Европе", size=13, color=GREY, italic=True, align=PP_ALIGN.CENTER)
anim.append(a + bars); note(s, NOTES[14])

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

# ---- 16. Капитан Блад
s = prs.slides.add_slide(BLANK); slides.append(s)
header(s, 16, "Часть 3 • Культура", "«Капитан Блад» — эхо эпохи\nпарусов и пиратов")
a, yy = bullets(s, [
    ("Фильм 1991 года", "«Одиссея капитана Блада» — про благородного пирата."),
    ("Морские сражения и сокровища", "пришли в кино из эпохи открытий."),
    ("Художники рисовали", "экзотических животных и растения."),
    ("Родился приключенческий роман.", "Мореплаватели стали как супергерои."),
], x=TX, w=TW, accent=ORANGE)
f = fact(s, PX_R, 2.15, 4.55, 2.4, "1991", "«Одиссея капитана Блада»\nфильм из списка учебника, с. 38", fill=BLUE)
a.append((f, "zoom"))
anim.append(a); note(s, NOTES[16])

# ---- 17. Аргумент 1
s = prs.slides.add_slide(BLANK); slides.append(s)
header(s, 17, "Вывод • Аргумент 1", "Мир стал единым:\nначалась глобализация")
a, yy = bullets(s, [
    ("Впервые появился", "единый мировой рынок."),
    ("Люди, товары и идеи", "начали перемещаться по всей планете."),
    ("Континенты перестали жить отдельно", "и стали зависеть друг от друга."),
], x=TX, w=TW, accent=BLUE)
f = fact(s, PX_R, 2.15, 4.55, 2.4, "1", "первый аргумент:\nмир соединился", fill=BLUE)
a.append((f, "zoom"))
anim.append(a); note(s, NOTES[17])

# ---- 18. Аргумент 2 (текст справа)
s = prs.slides.add_slide(BLANK); slides.append(s)
header(s, 18, "Вывод • Аргумент 2", "Изменилась еда, быт\nи сами люди", accent=GREEN)
f = fact(s, PX_L, 2.15, 4.55, 2.4, "2", "второй аргумент:\nКолумбов обмен в быту", fill=GREEN)
a, yy = bullets(s, [
    ("Изменился рацион:", "новые растения победили голод — вспомните картошку."),
    ("Смешались культуры,", "народы и языки."),
    ("Из этого смешения", "родились современные нации."),
], x=TX_R, w=TW_R, accent=GREEN)
a.insert(0, (f, "zoom"))
anim.append(a); note(s, NOTES[18])

# ---- 19. Аргумент 3 + финал
s = prs.slides.add_slide(BLANK); slides.append(s)
header(s, 19, "Вывод • Аргумент 3", "Открытия разрушили старые мифы\nи запустили науку")
a, yy = bullets(s, [
    ("Доказано на практике:", "Земля — шар, океаны соединены."),
    ("Появились точные карты.", "География и астрономия пошли вперёд."),
    ("Средневековый мир закончился.", "Началось Новое время — наш мир."),
], x=TX, w=TW, accent=BLUE)
f = fact(s, PX_R, 2.15, 4.55, 2.15, "3", "третий аргумент:\nнаучная революция", fill=BLUE)
a.append((f, "zoom"))
thanks = one(s, 0.6, 6.15, 12.13, 0.55, "Спасибо за внимание! Готов ответить на вопросы.",
             size=20, color=ORANGE, bold=True, align=PP_ALIGN.CENTER)
a.append((thanks, "zoom"))
anim.append(a); note(s, NOTES[19])

# ---- 20. Источники
s = prs.slides.add_slide(BLANK); slides.append(s)
header(s, 20, "Финал", "Источники", title_size=30)
one(s, 0.6, 1.95, 5.9, 0.4, "Материалы", size=16, color=BLUE, bold=True)
one(s, 6.88, 1.95, 5.85, 0.4, "Книги, фильм и картинки", size=16, color=BLUE, bold=True)
left = one(s, 0.6, 2.4, 5.9, 4.3,
           "• Учебник истории, с. 38 — темы проектов, список книг и фильм\n\n"
           "• Большая российская энциклопедия — bigenc.ru\n\n"
           "• Портал «История.РФ» — histrf.ru\n\n"
           "• Статьи по темам: Колумб, Магеллан, Колумбов обмен",
           size=14.5)
right = one(s, 6.88, 2.4, 5.85, 4.3,
            "• Книги со с. 38: Гуляев, Свет, Кинжалов и Белов, Кофман, Хаггард, Цвейг\n\n"
            "• Фильм: «Одиссея капитана Блада» (1991)\n\n"
            "• Картинки: Wikimedia Commons — commons.wikimedia.org\n"
            "• World History Encyclopedia — worldhistory.org\n"
            "• NCpedia — ncpedia.org",
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
