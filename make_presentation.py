#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Презентация v5: текст — речь Артёма (из его варианта), крупный шрифт, плотные интервалы.
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
def bg(slide, img="bg_light.jpg"):
    slide.shapes.add_picture(A(img), Inches(0), Inches(0), Inches(W), Inches(H))


def rect(slide, x, y, w, h, fill=WHITE, line=BORDER, round_=True):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if round_ else MSO_SHAPE.RECTANGLE,
                                 Inches(x), Inches(y), Inches(w), Inches(h))
    if round_:
        try: shp.adjustments[0] = 0.05
        except Exception: pass
    shp.fill.solid(); shp.fill.fore_color.rgb = fill
    shp.line.color.rgb = line; shp.line.width = Pt(1); shp.shadow.inherit_shadow = False
    return shp


def runs_box(slide, x, y, w, h, runs, size=17, align=PP_ALIGN.LEFT, line=1.05):
    tx = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tx.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = align; p.line_spacing = line
    for t, col, bold in runs:
        r = p.add_run(); r.text = t
        r.font.size = Pt(size); r.font.bold = bold; r.font.color.rgb = col; r.font.name = "Calibri"
    return tx


def plain(slide, x, y, w, h, text_, size=16, color=DARK, bold=False, align=PP_ALIGN.LEFT, italic=False):
    tx = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tx.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = align; p.line_spacing = 1.08
    r = p.add_run(); r.text = text_
    r.font.size = Pt(size); r.font.bold = bold; r.font.italic = italic; r.font.color.rgb = color; r.font.name = "Calibri"
    return tx


def head(slide, n, title, color=BLUE):
    bg(slide)
    plain(slide, 0.55, 0.28, 12.2, 0.6, title, size=27, color=color, bold=True)
    ln = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.57), Inches(0.88), Inches(12.2), Pt(3))
    ln.fill.solid(); ln.fill.fore_color.rgb = BLUE2; ln.line.fill.background(); ln.shadow.inherit_shadow = False
    plain(slide, 0.55, 7.18, 8, 0.25, "Артём Платонов  •  7 «Д»", size=8.5, color=GREY)
    plain(slide, 12.0, 7.18, 0.8, 0.25, f"{n} / 20", size=8.5, color=BLUE, bold=True, align=PP_ALIGN.RIGHT)
    return 1.15


def photo(slide, img, x, y, w, h, caption):
    rect(slide, x - 0.07, y - 0.07, w + 0.14, h + 0.14, fill=WHITE, line=BORDER, round_=False)
    slide.shapes.add_picture(A(img), Inches(x), Inches(y), Inches(w), Inches(h))
    plain(slide, x - 0.07, y + h + 0.08, w + 0.14, 0.55, caption, size=9, color=GREY, italic=True, align=PP_ALIGN.CENTER)


def steps(slide, items, x=0.55, y=1.35, w=12.1, size=17, lead_color=BLUE, dark=False, effects=None):
    """Плотно: маленький интервал между пунктами, шрифт крупный."""
    effects = effects or ["fly", "fade", "fly", "fade", "fly", "fade"]
    out = []; yy = y
    for i, (lead, rest) in enumerate(items):
        full = (lead + " " + rest).strip()
        chars = max(30, int(w * 118 / size))
        lines = max(1, math.ceil(len(full) / chars))
        h = lines * (size / 50) + 0.10
        runs = []
        if lead:
            runs.append((lead + "  ", YEL if dark else lead_color, True))
        if rest:
            runs.append((rest, WHITE if dark else DARK, False))
        sh = runs_box(slide, x, yy, w, h, runs, size=size)
        out.append((sh, effects[i % len(effects)]))
        yy += h + 0.16
    return out, yy


def sticky(slide, x, y, w, text_, size=13.5):
    rect(slide, x, y, w, 0.82, fill=YELLOW_BG, line=RGBColor(0xD9, 0xB3, 0x6A), round_=False)
    return plain(slide, x + 0.2, y + 0.12, w - 0.4, 0.6, text_, size=size, bold=True)


def note(slide, text_):
    slide.notes_slide.placeholders[1].text = text_


slides = []; anim = []

# ---------- заметки: полный текст выступления (вариант Артёма) ----------
NOTES = {
1: "Здравствуйте. Меня зовут Артём Платонов, я из 7 «Д». Мой проект — про Великие географические открытия. На странице 38 учебника даны три темы проектов, и я решил не брать одну, а соединить их. По-моему, они рассказывают одну общую историю: сначала люди поплыли в неизвестность, потом миры встретились, а потом всё изменилось.",
2: "Сначала я хотел взять только Колумба. Про него я что-то знал, и казалось, что так будет проще. Но когда перечитал все три темы, понял: по отдельности они скучнее. Поэтому у меня такой план: сначала причины плаваний, потом сами путешественники, потом обмен между материками и последствия. В конце покажу источники.",
3: "Причин я выделил четыре. Первая — пряности. Перец и другие специи были очень дорогими, потому что их перепродавали много раз, пока они доходили до Европы. Вторая — старый путь через восточные страны стал неудобным и опасным. Третья — появились новые корабли и приборы. Каравелла могла идти не только по ветру, но и против него, а моряки пользовались компасом и астролябией. Четвёртая причина — короли Испании и Португалии соперничали между собой и давали деньги на экспедиции.",
4: "Колумб считал, что Земля меньше, чем на самом деле. Поэтому он думал, что до Азии можно быстро доплыть с запада. Деньги на корабли дали Фердинанд и Изабелла. В августе 1492 года три корабля вышли в океан, а 12 октября показался остров Сан-Сальвадор. Колумб был уверен, что это окраина Индии, поэтому назвал местных жителей индейцами. Он плавал туда ещё три раза и умер в 1506 году, так и не узнав, что открыл новый материк. Запоминаем дату: 12 октября 1492 года.",
5: "Португалия пошла другим путём. Они решили не спорить, где находится Азия, а обогнуть Африку с юга. Васко да Гама вышел из Лиссабона в 1497 году, прошёл мыс Доброй Надежды, а через Индийский океан его провёл местный проводник. В мае 1498 года португальцы дошли до Каликута. Дорога в Индию была найдена. Правда, из четырёх кораблей домой вернулись только два, и людей погибло много.",
6: "Про экспедицию Магеллана мы много знаем благодаря дневнику Антонио Пигафетты. В 1519 году Магеллан повёл на запад пять кораблей. Он нашёл пролив у Южной Америки, теперь это Магелланов пролив. Потом они почти четыре месяца шли через огромный океан. Магеллан назвал его Тихим, потому что им повезло с погодой. Сам Магеллан погиб в 1521 году на Филиппинах. Домой в 1522 году вернулся один корабль «Виктория», и на нём было 18 человек. Главный итог: стало понятно, что Земля круглая, а океаны соединены.",
7: "Жак Картье искал проход в Азию с севера Америки. Прохода он не нашёл, но исследовал реку Святого Лаврентия и Канаду. Фрэнсис Дрейк был английским мореплавателем, а для Испании — пиратом. Он нападал на испанские корабли с серебром и совершил второе кругосветное плавание. Королева Елизавета даже сделала его рыцарем. Виллем Баренц искал путь в Азию мимо Сибири. У Новой Земли его корабль застрял во льдах, команда перезимовала в полярную ночь. Домой Баренц не вернулся. В его честь назвали Баренцево море.",
8: "Этот слайд для повторения. Колумб в 1492 году вышел к островам у Америки, но думал, что это Индия. Васко да Гама в 1498 году пришёл морем в Каликут. Команда Магеллана в 1519–1522 годах впервые обошла Землю вокруг. Картье в 1530-е исследовал Канаду. Дрейк в 1577–1580 годах совершил вторую кругосветку. Баренц в 1590-е пытался пройти северным путём и зимовал на Новой Земле.",
9: "После 1492 года начался обмен между Старым и Новым Светом. Это не только растения и животные, но ещё болезни и люди. Название «Колумбов обмен» придумал историк Альфред Кросби в 1972 году. То есть само явление старое, а термин довольно новый. Из Старого Света в Америку попали пшеница, лошади, коровы. Из Америки в Европу — картофель, томаты, какао.",
10: "Про растения мне проще рассказывать через еду. Картошка, помидоры, перец, кукуруза и какао появились в Америке. В Европе их раньше не было. Помидоры, кстати, сначала не ели, а держали для красоты. В обратную сторону уехали пшеница, рис, кофе, бананы и сахарный тростник. Картофель потом очень помог Европе, потому что давал урожай там, где пшенице было плохо.",
11: "С животными обмен получился неравным. В Америке до Колумба из домашних животных были ламы, индейки и морские свинки. Лошадей, коров и овец там не было. Европейцы привезли лошадей, коров, свиней, овец и кур. Особенно сильно на жизнь индейцев повлияла лошадь. Из Америки в Европу заметно уехала только индейка. А крысы, конечно, плавали сами.",
12: "Вместе с европейцами в Америку попали оспа, корь и грипп. Для местных жителей эти болезни были новыми, у них не было иммунитета. Поэтому умирали целые селения. Из-за этого на плантациях стало не хватать рабочих рук, и европейцы начали массово привозить порабощённых людей из Африки. Это трагическая часть Колумбова обмена, и про неё тоже надо говорить.",
13: "Испанцы разрушили столицу ацтеков Теночтитлан, а на её месте построили Мехико. Современный город буквально стоит на старом индейском. При этом индейские слова остались в языках: «томат», «шоколад», «какао». Население Латинской Америки стало смешанным. Там живут потомки европейцев, индейцев и африканцев. А карты мира впервые стали почти полными.",
14: "Торговля стала мировой. Раньше каждый регион торговал в основном с соседями, а теперь товары пошли между материками. Главные морские порты переместились на Атлантику: Лиссабон, Севилья, позже Амстердам и Лондон. Учёные получили много работы: надо было описать новые растения, животных и берега. География стала наукой с проверенными картами.",
15: "Но у этой эпохи была и тяжёлая цена. Государства ацтеков и инков не пережили встречи с испанцами. Отряды Кортеса и Писарро были небольшими, но у них были лошади, доспехи, оружие и союзники. Потом появилось рабство и работорговля через Атлантику. А американское серебро, которое пришло в Европу, обесценило деньги и вызвало рост цен. Поэтому назвать эту эпоху только хорошей нельзя.",
16: "По заданию учебника я собрал книги и фильм. Из книг там есть Гуляев, Свет, Кинжалов и Белов, Кофман, Хаггард и Цвейг. Я начал с Цвейга про Магеллана: он короткий и там есть дневник Пигафетты. Из кино в учебнике указан фильм «Одиссея капитана Блада» 1991 года. Ещё есть колонна Колумба в Барселоне.",
17: "Мне кажется, эта тема не отпускает потому, что в ней есть риск и неизвестность. Сейчас так же пишут про космос. Ещё герои здесь неудобные: Колумб вроде герой, но с его плаваний началось завоевание. И до сих пор есть споры, например, где именно Колумб высадился в первый раз. Пока учёные спорят, художники придумывают свои версии.",
18: "Первый: помидор в Европе долго боялись есть. Я думал, что новые продукты сразу стали популярными, но нет. Второй: кругосветку завершил не Магеллан, а Элькано. Про него почти никто не помнит, а ведь именно он привёл корабль домой. Третий: дом зимовки Баренца на Новой Земле действительно нашли. Он сохранился во льду, и вещи внутри тоже. История оказалась правдой до мелочей.",
19: "Мой главный вывод такой: после Великих географических открытий отдельные миры закончились и начался один общий. Это видно по растениям и животным, по мировой торговле и даже по ценам в Европе. Но я не могу сказать, что это было только хорошее время. Для Европы это был шаг вперёд, а для Америки и Африки — большая трагедия. В ответе надо держать обе части. Спасибо за внимание, я готов ответить на вопросы.",
20: "Я пользовался учебником, страницей 38, Большой российской энциклопедией и порталом «История.РФ». Ещё книги из списка учебника: Гуляев, Свет, Цвейг, Хаггард, Кинжалов и Белов. Фильм — «Одиссея капитана Блада» 1991 года. Картинки взяты с Wikimedia Commons, worldhistory.org и ncpedia.org. Текст я формулировал сам по этим материалам.",
}

# ---- 1. Титул
s = prs.slides.add_slide(BLANK); slides.append(s)
bg(s, "b_title.jpg")
plain(s, 0.6, 0.6, 9, 0.4, "ПРОЕКТ ПО ИСТОРИИ", size=13, color=YEL, bold=True)
plain(s, 0.6, 1.1, 11.5, 1.4, "Великие географические открытия", size=40, color=WHITE, bold=True)
a = [
    (plain(s, 0.6, 2.85, 11.5, 0.6, "Меня зовут Артём Платонов, я из 7 «Д»", size=21, color=WHITE, bold=True), "fade"),
    (plain(s, 0.6, 3.6, 11.8, 0.6, "На странице 38 учебника даны три темы — я соединил их в одну историю", size=18, color=WHITE), "fade"),
    (plain(s, 0.6, 4.4, 11.8, 0.6, "Сначала люди поплыли в неизвестность, потом миры встретились, а потом всё изменилось", size=18, color=YEL), "zoom"),
]
anim.append(a); note(s, NOTES[1])

# ---- 2. План (фото справа)
s = prs.slides.add_slide(BLANK); slides.append(s)
y = head(s, 2, "План")
a, yy = steps(s, [
    ("Сначала я хотел взять только Колумба.", "Про него я что-то знал, и казалось, что так будет проще."),
    ("Но когда перечитал все три темы, понял:", "по отдельности они скучнее."),
    ("Поэтому план такой:", "сначала причины плаваний, потом сами путешественники,"),
    ("потом обмен между материками и последствия.", "В конце покажу источники."),
], y=y + 0.2, w=7.4, size=16.5)
anim.append(a)
photo(s, "caravels.jpg", 8.45, 1.45, 4.3, 2.9, "Три корабля Колумба. Картинка со старой открытки.")
note(s, NOTES[2])

# ---- 3. Причины
s = prs.slides.add_slide(BLANK); slides.append(s)
y = head(s, 3, "Зачем европейцы полезли в океан")
a, yy = steps(s, [
    ("Первая — пряности.", "Перец и другие специи были очень дорогими, потому что их перепродавали много раз, пока они доходили до Европы."),
    ("Вторая — старый путь", "через восточные страны стал неудобным и опасным."),
    ("Третья — появились новые корабли и приборы.", "Каравелла могла идти не только по ветру, но и против него, а моряки пользовались компасом и астролябией."),
    ("Четвёртая — короли Испании и Португалии", "соперничали между собой и давали деньги на экспедиции."),
], y=y + 0.2, w=12.1, size=17)
anim.append(a); note(s, NOTES[3])

# ---- 4. Колумб (фото справа)
s = prs.slides.add_slide(BLANK); slides.append(s)
y = head(s, 4, "Колумб")
a, yy = steps(s, [
    ("Колумб считал, что Земля меньше, чем на самом деле.", "Поэтому он думал, что до Азии можно быстро доплыть с запада."),
    ("Деньги на корабли дали Фердинанд и Изабелла.", "В августе 1492 года три корабля вышли в океан."),
    ("12 октября показался остров Сан-Сальвадор.", "Колумб был уверен, что это окраина Индии, поэтому назвал местных жителей индейцами."),
    ("Он плавал туда ещё три раза", "и умер в 1506 году, так и не узнав, что открыл новый материк."),
], y=y + 0.2, w=7.4, size=16)
anim.append(a)
photo(s, "columbus.jpg", 8.35, 1.45, 4.4, 2.9, "Высадка Колумба на картине Джона Вандерлина.")
sticky(s, 0.55, 6.25, 12.23, "Запоминаем дату: 12 октября 1492 года.")
note(s, NOTES[4])

# ---- 5. Васко да Гама (фото слева)
s = prs.slides.add_slide(BLANK); slides.append(s)
y = head(s, 5, "Васко да Гама")
photo(s, "gama.jpg", 0.55, 1.45, 4.9, 3.65, "Прибытие португальцев в Каликут, Индия.")
a, yy = steps(s, [
    ("Португалия пошла другим путём.", "Они решили не спорить, где находится Азия, а обогнуть Африку с юга."),
    ("Васко да Гама вышел из Лиссабона в 1497 году,", "прошёл мыс Доброй Надежды, а через Индийский океан его провёл местный проводник."),
    ("В мае 1498 года португальцы дошли до Каликута.", "Дорога в Индию была найдена."),
    ("Правда, из четырёх кораблей домой вернулись только два,", "и людей погибло много."),
], x=5.85, y=y + 0.2, w=6.9, size=15.5)
anim.append(a); note(s, NOTES[5])

# ---- 6. Магеллан (фото справа)
s = prs.slides.add_slide(BLANK); slides.append(s)
y = head(s, 6, "Магеллан")
a, yy = steps(s, [
    ("Про экспедицию Магеллана мы много знаем", "благодаря дневнику Антонио Пигафетты."),
    ("В 1519 году Магеллан повёл на запад пять кораблей.", "Он нашёл пролив у Южной Америки, теперь это Магелланов пролив."),
    ("Потом они почти четыре месяца шли через огромный океан.", "Магеллан назвал его Тихим, потому что им повезло с погодой."),
    ("Сам Магеллан погиб в 1521 году на Филиппинах.", "Домой в 1522 году вернулся один корабль «Виктория», и на нём было 18 человек."),
], y=y + 0.2, w=7.4, size=15.5)
anim.append(a)
photo(s, "victoria.png", 8.35, 1.45, 4.4, 2.7, "Корабль «Виктория» на старой карте Ортелиуса.")
sticky(s, 0.55, 6.25, 12.23, "Главный итог: стало понятно, что Земля круглая, а океаны соединены.")
note(s, NOTES[6])

# ---- 7. Трое (колонки)
s = prs.slides.add_slide(BLANK); slides.append(s)
y = head(s, 7, "Картье, Дрейк, Баренц")
trio = [
    ("Жак Картье", BLUE, "Искал проход в Азию с севера Америки. Прохода не нашёл, но исследовал реку Святого Лаврентия и Канаду."),
    ("Фрэнсис Дрейк", ORANGE, "Английский мореплаватель, а для Испании — пират. Нападал на испанские корабли с серебром и совершил второе кругосветное плавание. Королева Елизавета даже сделала его рыцарем."),
    ("Виллем Баренц", GREEN, "Искал путь в Азию мимо Сибири. У Новой Земли его корабль застрял во льдах, команда перезимовала в полярную ночь. Домой Баренц не вернулся. В его честь назвали Баренцево море."),
]
x = 0.55; a = []
for title_, col, body in trio:
    rect(s, x, 1.4, 3.95, 4.5)
    plain(s, x + 0.22, 1.6, 3.5, 0.4, title_, size=17, color=col, bold=True)
    sh = plain(s, x + 0.22, 2.1, 3.5, 3.6, body, size=13.5)
    a.append((sh, "fade")); x += 4.14
anim.append(a); note(s, NOTES[7])

# ---- 8. Повторение
s = prs.slides.add_slide(BLANK); slides.append(s)
y = head(s, 8, "Повторение")
a, yy = steps(s, [
    ("Колумб, 1492 —", "вышел к островам у Америки, но думал, что это Индия."),
    ("Васко да Гама, 1498 —", "пришёл морем в Каликут."),
    ("Команда Магеллана, 1519–1522 —", "впервые обошла Землю вокруг."),
    ("Картье, 1530-е —", "исследовал Канаду."),
    ("Дрейк, 1577–1580 —", "совершил вторую кругосветку."),
    ("Баренц, 1590-е —", "пытался пройти северным путём и зимовал на Новой Земле."),
], y=y + 0.2, w=12.1, size=17, effects=["fade"] * 6)
anim.append(a); note(s, NOTES[8])

# ---- 9. Колумбов обмен + схема
s = prs.slides.add_slide(BLANK); slides.append(s)
y = head(s, 9, "Колумбов обмен", color=GREEN)
a, yy = steps(s, [
    ("После 1492 года начался обмен между Старым и Новым Светом.", ""),
    ("Это не только растения и животные,", "но ещё болезни и люди."),
    ("Название «Колумбов обмен» придумал историк Альфред Кросби в 1972 году.", "То есть само явление старое, а термин довольно новый."),
], y=y + 0.2, w=12.1, size=17, lead_color=GREEN)
anim.append(a)
rect(s, 0.55, 4.75, 4.4, 1.05, fill=LIGHT, line=BLUE2, round_=False)
plain(s, 0.75, 5.0, 4.0, 0.6, "Из Старого Света: пшеница, лошади, коровы", size=13, color=BLUE, bold=True, align=PP_ALIGN.CENTER)
rect(s, 8.38, 4.75, 4.4, 1.05, fill=RGBColor(0xFB, 0xE8, 0xDD), line=ORANGE, round_=False)
plain(s, 8.58, 5.0, 4.0, 0.6, "Из Америки: картофель, томаты, какао", size=13, color=ORANGE, bold=True, align=PP_ALIGN.CENTER)
for yy2, kind in ((4.95, MSO_SHAPE.RIGHT_ARROW), (5.4, MSO_SHAPE.LEFT_ARROW)):
    ar = s.shapes.add_shape(kind, Inches(5.15), Inches(yy2), Inches(3.05), Inches(0.32))
    ar.fill.solid(); ar.fill.fore_color.rgb = BLUE2 if kind == MSO_SHAPE.RIGHT_ARROW else ORANGE
    ar.line.fill.background(); ar.shadow.inherit_shadow = False
note(s, NOTES[9])

# ---- 10. Еда / растения
s = prs.slides.add_slide(BLANK); slides.append(s)
y = head(s, 10, "Про растения — через еду", color=GREEN)
a, yy = steps(s, [
    ("Картошка, помидоры, перец, кукуруза и какао появились в Америке.", "В Европе их раньше не было."),
    ("Помидоры, кстати, сначала не ели, а держали для красоты.", ""),
    ("В обратную сторону уехали пшеница, рис, кофе, бананы и сахарный тростник.", ""),
    ("Картофель потом очень помог Европе,", "потому что давал урожай там, где пшенице было плохо."),
], y=y + 0.25, w=12.1, size=17.5, lead_color=GREEN)
anim.append(a); note(s, NOTES[10])

# ---- 11. Животные
s = prs.slides.add_slide(BLANK); slides.append(s)
y = head(s, 11, "Животные", color=GREEN)
a, yy = steps(s, [
    ("С животными обмен получился неравным.", "В Америке до Колумба из домашних животных были ламы, индейки и морские свинки."),
    ("Лошадей, коров и овец там не было.", "Европейцы привезли лошадей, коров, свиней, овец и кур."),
    ("Особенно сильно на жизнь индейцев повлияла лошадь.", ""),
    ("Из Америки в Европу заметно уехала только индейка.", "А крысы, конечно, плавали сами."),
], y=y + 0.25, w=12.1, size=17.5, lead_color=GREEN)
anim.append(a); note(s, NOTES[11])

# ---- 12. Болезни
s = prs.slides.add_slide(BLANK); slides.append(s)
y = head(s, 12, "Болезни", color=RED)
a, yy = steps(s, [
    ("Вместе с европейцами в Америку попали оспа, корь и грипп.", ""),
    ("Для местных жителей эти болезни были новыми,", "у них не было иммунитета. Поэтому умирали целые селения."),
    ("Из-за этого на плантациях стало не хватать рабочих рук,", "и европейцы начали массово привозить порабощённых людей из Африки."),
    ("Это трагическая часть Колумбова обмена,", "и про неё тоже надо говорить."),
], y=y + 0.25, w=12.1, size=17.5, lead_color=RED, effects=["fade"] * 4)
anim.append(a); note(s, NOTES[12])

# ---- 13. Встреча культур
s = prs.slides.add_slide(BLANK); slides.append(s)
y = head(s, 13, "Встреча культур", color=GREEN)
a, yy = steps(s, [
    ("Испанцы разрушили столицу ацтеков Теночтитлан,", "а на её месте построили Мехико. Современный город буквально стоит на старом индейском."),
    ("При этом индейские слова остались в языках:", "«томат», «шоколад», «какао»."),
    ("Население Латинской Америки стало смешанным.", "Там живут потомки европейцев, индейцев и африканцев."),
    ("А карты мира впервые стали почти полными.", ""),
], y=y + 0.25, w=12.1, size=17.5, lead_color=GREEN)
anim.append(a); note(s, NOTES[13])

# ---- 14. Хорошее
s = prs.slides.add_slide(BLANK); slides.append(s)
y = head(s, 14, "Что изменилось: хорошее")
a, yy = steps(s, [
    ("Торговля стала мировой.", "Раньше каждый регион торговал в основном с соседями, а теперь товары пошли между материками."),
    ("Главные морские порты переместились на Атлантику:", "Лиссабон, Севилья, позже Амстердам и Лондон."),
    ("Учёные получили много работы:", "надо было описать новые растения, животных и берега."),
    ("География стала наукой с проверенными картами.", ""),
], y=y + 0.25, w=12.1, size=17.5)
anim.append(a); note(s, NOTES[14])

# ---- 15. Тяжёлое (фото справа)
s = prs.slides.add_slide(BLANK); slides.append(s)
y = head(s, 15, "Что изменилось: тяжёлое", color=RED)
a, yy = steps(s, [
    ("Но у этой эпохи была и тяжёлая цена.", "Государства ацтеков и инков не пережили встречи с испанцами."),
    ("Отряды Кортеса и Писарро были небольшими,", "но у них были лошади, доспехи, оружие и союзники."),
    ("Потом появилось рабство и работорговля через Атлантику.", ""),
    ("А американское серебро обесценило деньги в Европе и вызвало рост цен.", "Поэтому назвать эту эпоху только хорошей нельзя."),
], y=y + 0.2, w=7.4, size=15.5, lead_color=RED, effects=["fade"] * 4)
anim.append(a)
photo(s, "tenochtitlan.png", 8.35, 1.45, 4.4, 2.95, "Штурм Теночтитлана на картине XVII века.")
note(s, NOTES[15])

# ---- 16. Книги и фильм
s = prs.slides.add_slide(BLANK); slides.append(s)
y = head(s, 16, "Книги и фильм по заданию учебника")
rect(s, 0.55, 1.4, 6.0, 4.6)
plain(s, 0.8, 1.6, 5.5, 0.4, "Книги со страницы 38:", size=15, color=BLUE, bold=True)
left = plain(s, 0.8, 2.05, 5.5, 3.7,
             "Гуляев «По следам конкистадоров»\nСвет «Колумб»\nКинжалов и Белов «Падение Теночтитлана»\nКофман «Тьерра аделанте!»\nХаггард «Дочь Монтесумы»\nЦвейг «Америго» и «Подвиг Магеллана»", size=14)
rect(s, 6.78, 1.4, 6.0, 4.6)
plain(s, 7.03, 1.6, 5.5, 0.4, "Что я выбрал и что посмотреть:", size=15, color=BLUE, bold=True)
right = plain(s, 7.03, 2.05, 5.5, 3.7,
              "Я начал с Цвейга про Магеллана: он короткий и там есть дневник Пигафетты.\n\nФильм из учебника — «Одиссея капитана Блада» 1991 года.\n\nЕщё есть колонна Колумба в Барселоне.", size=14)
anim.append([(left, "fade"), (right, "fade")]); note(s, NOTES[16])

# ---- 17. Почему тема интересна
s = prs.slides.add_slide(BLANK); slides.append(s)
y = head(s, 17, "Почему тема до сих пор интересна")
a, yy = steps(s, [
    ("Мне кажется, эта тема не отпускает потому,", "что в ней есть риск и неизвестность."),
    ("Сейчас так же пишут про космос.", ""),
    ("Ещё герои здесь неудобные:", "Колумб вроде герой, но с его плаваний началось завоевание."),
    ("И до сих пор есть споры,", "например, где именно Колумб высадился в первый раз."),
    ("Пока учёные спорят, художники придумывают свои версии.", ""),
], y=y + 0.2, w=12.1, size=17)
anim.append(a); note(s, NOTES[17])

# ---- 18. Удивило меня
s = prs.slides.add_slide(BLANK); slides.append(s)
y = head(s, 18, "Три факта, которые удивили меня", color=ORANGE)
a, yy = steps(s, [
    ("Первый: помидор в Европе долго боялись есть.", "Я думал, что новые продукты сразу стали популярными, но нет."),
    ("Второй: кругосветку завершил не Магеллан, а Элькано.", "Про него почти никто не помнит, а ведь именно он привёл корабль домой."),
    ("Третий: дом зимовки Баренца на Новой Земле действительно нашли.", "Он сохранился во льду, и вещи внутри тоже. История оказалась правдой до мелочей."),
], y=y + 0.3, w=12.1, size=18, lead_color=ORANGE)
anim.append(a); note(s, NOTES[18])

# ---- 19. Вывод (тёмный)
s = prs.slides.add_slide(BLANK); slides.append(s)
bg(s, "b_map.jpg")
plain(s, 0.6, 0.45, 12.1, 0.65, "Вывод", size=30, color=WHITE, bold=True)
ln = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.62), Inches(1.15), Inches(1.6), Pt(4))
ln.fill.solid(); ln.fill.fore_color.rgb = YEL; ln.line.fill.background(); ln.shadow.inherit_shadow = False
a, yy = steps(s, [
    ("Мой главный вывод такой:", "после Великих географических открытий отдельные миры закончились и начался один общий."),
    ("Это видно по растениям и животным,", "по мировой торговле и даже по ценам в Европе."),
    ("Но это было не только хорошее время.", "Для Европы — шаг вперёд, а для Америки и Африки — большая трагедия."),
    ("В ответе надо держать обе части.", ""),
], y=1.45, w=11.8, size=18, dark=True)
a.append((plain(s, 0.6, 6.35, 12.1, 0.55, "Спасибо за внимание, я готов ответить на вопросы.", size=19, color=YEL, bold=True, align=PP_ALIGN.CENTER), "zoom"))
anim.append(a); note(s, NOTES[19])

# ---- 20. Источники
s = prs.slides.add_slide(BLANK); slides.append(s)
y = head(s, 20, "Источники")
srcs = [
    "Я пользовался учебником, страницей 38, Большой российской энциклопедией и порталом «История.РФ».",
    "Ещё книги из списка учебника: Гуляев, Свет, Цвейг, Хаггард, Кинжалов и Белов.",
    "Фильм — «Одиссея капитана Блада» 1991 года.",
    "Картинки взяты с Wikimedia Commons, worldhistory.org и ncpedia.org.",
    "Текст я формулировал сам по этим материалам.",
]
a, yy = steps(s, [(src, "") for src in srcs], y=y + 0.25, w=12.1, size=16, effects=["fade"] * 5)
anim.append(a); note(s, NOTES[20])

# ---- переходы и анимации
trans_kinds = ["fade", "push", "wipe", "fade", "cover"] * 4
for sl, a, tk in zip(slides, anim, trans_kinds):
    set_transition(sl, tk)
    if a:
        add_animations(sl, a)

prs.save(OUT)
print("Готово:", OUT, "| слайдов:", len(slides))
