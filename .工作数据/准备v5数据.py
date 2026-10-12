# -*- coding: utf-8 -*-
"""v5 数据准备：一校一册；册内按「考试科目 ＋ 科目代码 ＋ 参考书」分段。

输入：.工作数据/v4考点数据.json（v4 定稿的 322 条册内条目）
      .高频考点工作/{学校}_历年真题_打印版.json（原卷题面，用于核对）
输出：.工作数据/v5考点数据.json、.工作数据/v5分段清单.json

v5 与 v4 的差别：
1. 交付形态改为「一个学校一个 PDF 册」；科目与参考书在册内分段，不再一科目一个文件。
2. 分段键 = 考试科目（卷面科目名）＋ 科目代码（卷面印刷，附年份区间）＋ 参考书。
   同科同书但科目代码不同年段的拆开，各自标注代码与年份；同卷异书的（黑龙江 842 的
   古代汉语／语言学理论、西安语言综合的现汉／古汉）按书分列。
3. 每段的年份范围、出处页码、条目数、证据数全部由数据现算，不手写。
4. 每段只显示落在本段年份区间内的证据；同一考点跨段有证据时在两段各出现一次，
   各自只带本段证据——既满足「不同书不并册」，也不重复计证据。
5. 机械自检：① 唯一考点集与 v4 一致；② 分册条目与 v4 一致；③ v4 每条非占位证据
   都在 v5 至少出现一次；占位证据与同题重复证据的去向见 v5证据补正清单.json。
"""
import json
import sys
from collections import OrderedDict
from pathlib import Path

W = Path(__file__).resolve().parent
V4 = json.loads((W / 'v4考点数据.json').read_text())
# 原卷分卷页码索引：每个「年份＋科目」在打印复核版里的起止页脚页码
INDEX = json.loads((W.parent / '.高频考点工作' / 'pdf_index.json').read_text())

# 人工补录的占位证据（pid 1201+／1301+／1401+）：没有逐行题面可回溯，由
# 补正v5证据.py 按「证据＝卷别＋题号＋原文」的口径补正后再比对，故不计入守恒基线。
SYNTH = set(range(1201, 1210)) | set(range(1301, 1310)) | set(range(1401, 1410))

VERSION = 'v5'

# ------------------------------------------------------------------ 证据级别
EV = {
    'hlj-2026官方': '官方：黑龙江大学 2026 年 742、842 两份考试大纲原件（文内附参考书目）。',
    'hlj-2021转载': '权威转载：黑龙江大学 2021 年招生章程所附大纲（原始来源为该校研究生院，'
                    '官网直连返回 403，未能取得原件）。',
    'yn官方': '官方：云南民族大学《2025 年硕士研究生招生初试参考书目》'
              '《2026 年硕士研究生招生初试参考书目》PDF 原件。',
    'xatu官方': '官方：西安工业大学研究生院 2024、2026 年招生专业目录所附参考书目页；'
                '2025 年为权威转载（转载页标注来源为该校研究生院）。',
    'ah-官方目录': '官方＋权威转载：安徽大学 2018 年专业目录（安徽大学官网文件，含文学院段落）'
                   '与 2019／2020／2021 年专业目录（安徽大学研究生院官网）为官方原件；'
                   '2012—2017 年目录为权威转载（转载页标注来源为安徽大学研究生院）。'
                   '按年份逐年核对：现代汉语 2007—2015 为胡裕树本、2016—2019 为黄廖增订五版、'
                   '2020—2021 为黄廖增订六版；古代汉语 2007—2018 为王力本、2019—2021 为黄德宽本。',
    'ah-未查到': '未查到：2004—2011 年该科的初试参考书目原件与可靠转载均未取得，'
                 '本册如实留空，不代填书目、不用后一年的书目倒推。',
}

# ------------------------------------------------------------------ 分段规格
# page 取 .高频考点工作/pdf_index.json 记录的该科目全部年份起止页并集（打印复核版页脚页码）。
# 每条：key=段内唯一键；name=段标题用的科目名；code=科目代码（卷面印刷口径）；
#       years=本段年份区间（闭区间）；disc=学科归属（可给元组）；
#       src=科目归属（对应 v4 的 paper）；books=本段书单键；note=本段说明。
SPEC = OrderedDict([
    ('安徽大学', [
        dict(key='ah-jc-none', name='语言文学基础', code='2006 年印 316；2007—2011 年目录为 605',
             label='语言文学基础（2004—2011）', years=(2004, 2011), disc='基础',
             src='语言文学基础（616）', books=['ah-none', 'gwx', 'xdwx', 'wxll'],
             page='p1—p36',
             ev='ah-未查到',
             note='本段对应 2004—2011 年的「语言文学基础」。卷面情况：2004—2008 年卷头只印试题名称、'
                  '代码栏留空；2009、2010、2011 年卷头印「试题代码：605」。'
                  '官方专业目录：2006 年该科为 316，2007 年起为 605（2011 年为 605 的最后一年）。'
                  '书目情况：本段各年的初试参考书目原件未取得，库内亦无可靠转载，故本册不代填书目，'
                  '只呈现真题本身：考点按知识块编排，出处给到卷面页码与题号。'),
        dict(key='ah-jc-605', name='语言文学基础', code='605',
             label='语言文学基础（605）', years=(2012, 2016), disc='基础',
             src='语言文学基础（616）', books=['xdh_hu', 'gdh_wl', 'gwx', 'xdwx', 'wxll'],
             page='p37—p58',
             ev='ah-官方目录',
             note='本段对应 2012—2016 年。卷面情况：2012 年卷头印「试题代码：605」；'
                  '2013—2016 年卷头只印试题名称、未印代码。官方专业目录：2012 年起'
                  '「语言文学基础」为 605，配套「阅读与写作」为 807。'
                  '书目（按官方目录逐年核对）：现代汉语 2012—2015 年为胡裕树主编《现代汉语》'
                  '（上海教育出版社 1995 年版）、2016 年为黄伯荣、廖序东《现代汉语》增订五版；'
                  '古代汉语为 2007—2018 年一贯的王力主编《古代汉语》（中华书局 1999 年第三版）。'),
        dict(key='ah-jc-615', name='语言文学基础', code='615',
             label='语言文学基础（615）', years=(2017, 2018), disc='基础',
             src='语言文学基础（616）', books=['xdh5', 'gdh_wl', 'gwx', 'xdwx', 'wxll'],
             page='p59—p62',
             ev='ah-官方目录',
             note='本段对应 2017—2018 年。卷面情况：2017、2018 年卷头只印试题名称、未印代码。'
                  '官方专业目录：2017 年起「语言文学基础」改为 615，配套「阅读与写作」由 807 改为 808'
                  '（2017 年为 808 首年，2018 年 615／808 有安徽大学官网文件印证）。'
                  '书目（按官方目录逐年核对）：现代汉语为黄伯荣、廖序东《现代汉语》增订五版'
                  '（2016—2019 年指定），古代汉语仍为王力主编《古代汉语》（1999 年第三版）。'),
        dict(key='ah-jc-616', name='语言文学基础', code='616',
             label='语言文学基础（616）', years=(2019, 2021), disc='基础',
             src='语言文学基础（616）', books=['xdh5', 'xdh', 'gdh', 'gwx', 'xdwx', 'wxll'],
             page='p63—p68',
             ev='ah-官方目录',
             note='本段对应 2019—2021 年。卷面情况：2019 年卷内自注「616 语言文学基础」、'
                  '2020 与 2021 年卷头印「专业课一 代码：616」，配套「809 阅读与写作」；'
                  '2019 年为 616／809 首年（安徽大学官网专业目录原件）。'
                  '书目（按官方目录逐年核对）：古代汉语改为黄德宽主编《古代汉语》'
                  '（高等教育出版社 2015 年版，2019 年起指定）；现代汉语 2019 年为黄廖增订五版、'
                  '2020—2021 年为增订六版。'
                  '注意：2016 年另有「汉语言文字学」方向的 616 古代汉语和现代汉语基础／'
                  '809 古文字学，与本段 050102 所考的 616／809 不是同一组科目。'),
        dict(key='ah-ydxz', name='阅读与写作',
             code='807（2012—2016）／808（2017—2018）／809（2019—2021）',
             label='阅读与写作（807／808／809）', years=(2004, 2021), disc=('读写', '基础'),
             src='阅读与写作（809）', books=['ydxz'],
             page='p15—p68',
             ev='ah-官方目录',
             note='卷面情况：2009、2010、2011、2012、2014 年卷头印「试题代码：807」；'
                  '2013、2015—2018 年未印代码；2019 年卷内自注 809，2020、2021 年卷头印'
                  '「专业课二 代码：809」。官方专业目录：807（2007—2016）→808（2017—2018）'
                  '→809（2019—2021）。'
                  '该科官方表述为考查「阅读分析与文学写作的实际能力」，历年目录均未列指定教材，'
                  '故本段不代填书目。段内 2004 年的一条证据出自当年「语言文学基础」卷，属跨卷归入，'
                  '已在条目里单独标注。'),
    ]),
    ('黑龙江大学', [
        dict(key='hlj-xdh', name='现代汉语', code='742',
             label='现代汉语（742）', years=(2010, 2026), disc='现汉',
             src='现代汉语（742）', books=['hlj-xdh'],
             page='p2—p55',
             ev='hlj-2021转载',
             note='卷面情况：2011—2014 年的现代汉语内容在「语言学理论与现代汉语语音、文字、词汇」'
                  '合卷内，卷头未印代码（该卷 2013、2014 两年官方代码为 709）；'
                  '2015—2020、2023 年卷头印科目名「现代汉语」，未印代码（2015 年起官方代码为 742）；'
                  '2021 年卷头印「考试科目：742 现代汉语 共 3 页」；2025、2026 为回忆版合卷，'
                  '卷内分「现代汉语」与「古代汉语和语纲」两块。本段只取各卷现代汉语部分，'
                  '同一学科同年最多计一次。2022 年现汉缺卷、2024 年现汉缺卷。'
                  '2011、2012 两年官方目录未查到，代码是否同为 709 未证实，故本段不写这两年的代码。'),
        dict(key='hlj-gdh', name='古代汉语与语言学理论', code='842',
             label='古代汉语与语言学理论（842）·古代汉语部分', years=(2011, 2026), disc='古汉',
             src='古代汉语与语言学理论（842）', books=['hlj-gdh'],
             page='p4—p54',
             ev='hlj-2026官方',
             note='卷面情况：2011—2014 年的古代汉语内容在「古代汉语与现代汉语语法、修辞」合卷内，'
                  '卷头未印代码（该卷 2013、2014 两年官方代码为 809，2015 年起改为 842）；'
                  '2015 年卷头印「考试科目：842 古代汉语与语言学理论」；'
                  '2021 年卷头印「考试科目：842 古代汉语与语言学理论 共 2 页」；'
                  '2016—2020、2023 年卷头只印科目名。本段只取各卷古代汉语部分，与语理部分分列。'),
        dict(key='hlj-yyll', name='古代汉语与语言学理论', code='842',
             label='古代汉语与语言学理论（842）·语言学理论部分', years=(2010, 2026), disc='语理',
             src='古代汉语与语言学理论（842）', books=['hlj-yyll'],
             page='p1—p51',
             ev='hlj-2026官方',
             note='卷面情况：2010 年为单卷「语言学理论（共 75 分）」；2011—2014 年在合卷第一部分'
                  '（该卷 2013、2014 两年官方代码为 709）；2015 年起与古代汉语并卷为 842，'
                  '本部分在卷内第二部分；2025、2026 回忆版把该部分记作「语言学纲要」。'
                  '本段与古代汉语部分同卷、异书，按参考书分列。'),
    ]),
    ('云南民族大学', [
        dict(key='yn-xdh', name='现代汉语', code='719',
             label='现代汉语（719）', years=(2013, 2025), disc='现汉',
             src='现代汉语（719）', books=['yn-xdh'],
             page='p3—p27',
             ev='yn官方',
             note='卷面情况：2013—2024 年卷头只印「考试科目：现代汉语」，未印代码；'
                  '2025 年卷头在科目名下一行单独印「(719)」。'
                  '官方专业目录：2013、2016、2019、2020、2022、2025、2026 七个年份均列'
                  '「719 现代汉语／815 语言学理论」，未发现换码痕迹，可判定 2013 年至今未变；'
                  '2014、2017、2021、2023、2024 未逐年在官方原件上复核。'
                  '书目：官方 2025、2026 年参考书目均把「719 现代汉语」列为'
                  '黄伯荣、廖序东《现代汉语》增订第六版上、下册（2017 年版）。'),
        dict(key='yn-yyll', name='语言学理论', code='815',
             label='语言学理论（815）', years=(2011, 2025), disc='语理',
             src='语言学理论（815）', books=['yn-yyll'],
             page='p1—p29',
             ev='yn官方',
             note='卷面情况：2019、2021、2022、2023 年卷头在学科专业下方印「815 语言学理论」，'
                  '2024 年卷头印「815 语言学理论」，2025 年卷头单独印「815」；'
                  '2011—2018 年卷头未印代码。'
                  '官方专业目录：2013、2016、2019、2020、2022、2025、2026 七个年份均列 815，'
                  '未发现换码痕迹。注意同校 050107 的第四科曾是 816「民族文学、文学理论」，'
                  '与本段的 815 不是同一科目，不要混用。'
                  '书目提示：2024 年公开书目转载同时列《语言学纲要》与马学良《语言学概论》，'
                  '2025、2026 年官方原件（2026 年起改 2025 年版）均只列《语言学纲要》；'
                  '本段证据全部落在《语言学纲要》范围内，马学良本未单独成段。'),
    ]),
    ('西安工业大学', [
        dict(key='xat-902', name='语言综合', code='902',
             label='语言综合（902）', years=(2020, 2024), disc='语综',
             src='语言综合（902／816）', books=['xat-xdh', 'xat-gdh'],
             page='p1—p46',
             ev='xatu官方',
             note='卷面情况：2020 年卷头未印编号；2021、2022、2023 年卷头印「试题编号 902」；'
                  '2024 年卷头未印编号。官方口径：西安工业大学研究生院《2024 年硕士研究生招生考试'
                  '参考书目》文学院栏写「语言综合（902）」「文学评论写作（702）」，'
                  '故 2024 年仍是 902，2025 年才改为 816。本段与 816 段同书、异代码，按代码分段。'),
        dict(key='xat-816', name='语言综合', code='816',
             label='语言综合（816）', years=(2025, 2025), disc='语综',
             src='语言综合（902／816）', books=['xat-xdh', 'xat-gdh'],
             page='p47—p50',
             ev='xatu官方',
             note='卷面情况：2025 年卷头印「试题编号：816」。官方口径：2025 年自命题参考书目列'
                  '「816 语言综合」，2026 年文学院官方页亦为「816 语言综合」；'
                  '未找到任何官方文件把 2024 年写成 816。'
                  '本段只有 2025 一年，按整卷照做，不因年份少而并入 902 段：'
                  '该年古代汉语指定版本与 902 段不同（见书单版本栏）。'),
        dict(key='xat-702', name='文学评论写作', code='702',
             label='文学评论写作（702）', years=(2020, 2026), disc='评论',
             src='文学评论写作（702）', books=['xat-wxll'],
             page='p4—p61',
             ev='xatu官方',
             note='卷面情况：2020—2026 各年卷头均印「试题编号 702」。'),
    ]),
])

# ------------------------------------------------------------------ 书单
# key → 书名、出版社、版本依据、证据级别键、学科、说明
BOOKS = {
    'ah-none': dict(name='本段未查到官方指定参考书', pub='—',
                    edition='2004—2011 年初试参考书目原件未取得，本册不代填',
                    ev='ah-未查到', disc='基础', kind='none',
                    note='书目栏留空是如实标注，不是「无书可考」；本段考点按知识块编排，'
                         '出处给到卷面页码与题号。'),
    'xdh_hu': dict(name='胡裕树主编《现代汉语》', pub='上海教育出版社',
                   edition='1995 年版；2007—2015 年参考书目指定', ev='ah-官方目录',
                   disc='基础', kind='official', note=None),
    'xdh5': dict(name='黄伯荣、廖序东《现代汉语》（增订五版）', pub='高等教育出版社',
                 edition='2016—2019 年参考书目标注「增订五版」', ev='ah-官方目录',
                 disc='基础', kind='official', note=None),
    'xdh': dict(name='黄伯荣、廖序东《现代汉语》（增订六版）', pub='高等教育出版社',
                edition='2020—2021 年参考书目标注「六版」', ev='ah-官方目录',
                disc='基础', kind='official', note=None),
    'gdh_wl': dict(name='王力主编《古代汉语》', pub='中华书局',
                   edition='1999 年第三版；2007—2018 年参考书目指定', ev='ah-官方目录',
                   disc='基础', kind='official', note=None),
    'gdh': dict(name='黄德宽主编《古代汉语》', pub='高等教育出版社',
                edition='2015 年版（2019 年另一科目 617 的书目写明该版本）；2019—2021 年指定',
                ev='ah-官方目录', disc='基础', kind='official', note=None),
    'gwx': dict(name='袁行霈《中国文学史》（第三版）', pub='高等教育出版社',
                edition='2021 年专业目录标注「三版」', ev='ah-官方目录',
                disc='基础', kind='official', note=None),
    'xdwx': dict(name='丁帆《中国新文学史》', pub='高等教育出版社',
                 edition='2021 年专业目录标示，版本页未标年份', ev='ah-官方目录',
                 disc='基础', kind='official', note=None),
    'wxll': dict(name='童庆炳《文学理论教程》（第四版）', pub='高等教育出版社',
                 edition='2021 年专业目录标注「四版」', ev='ah-官方目录',
                 disc='基础', kind='official', note=None),
    'ydxz': dict(name='本段无指定参考书：阅读与写作（实际能力考查）', pub='—',
                 edition='官方从未列指定教材（2007—2021 历年目录均无本册书目）',
                 ev='ah-官方目录', disc='读写', kind='none',
                 note='参照范围：古代汉语的文选、句读、翻译与古注知识（2007—2018 年该科参照王力本，'
                      '2019 年起参照黄德宽本），以及文学鉴赏与命题写作的实际表达能力。'),
    'hlj-xdh': dict(name='黄伯荣、廖序东《现代汉语》（增订第五版）', pub='高等教育出版社',
                    edition='2015 年大纲转载列「增订第五版」',
                    ev='hlj-2021转载', disc='现汉', kind='official',
                    note='版本提示：该校 2026 年官方大纲已把参考书改为邵敬敏《现代汉语通论》第 3 版'
                         '（2017 年，上海教育出版社），属 2025 年之后的范围变动，'
                         '不能用来评价既有真题。'),
    'hlj-gdh': dict(name='王力主编《古代漢語》（一、二册）', pub='中华书局',
                    edition='校订重排本；2021 年大纲转载未标出版年份，2026 年官方大纲同',
                    ev='hlj-2026官方', disc='古汉', kind='official', note=None),
    'hlj-yyll': dict(name='《语言学纲要》（修订版），叶蜚声、徐通锵著，王洪君、李娟修订',
                     pub='北京大学出版社', edition='修订版；2021 年大纲转载与 2026 年官方大纲一致',
                     ev='hlj-2026官方', disc='语理', kind='official', note=None),
    'yn-xdh': dict(name='黄伯荣、廖序东《现代汉语》（增订第六版）上、下册', pub='高等教育出版社',
                   edition='2017 年版；2025、2026 官方书目一致未变',
                   ev='yn官方', disc='现汉', kind='official', note=None),
    'yn-yyll': dict(name='《语言学纲要》（修订本），叶蜚声、徐通锵著，王洪君、李娟修订',
                    pub='北京大学出版社', edition='2010 年版；2026 年官方书目起改为 2025 年版',
                    ev='yn官方', disc='语理', kind='official',
                    note='2024 年公开书目转载另列马学良《语言学概论》，该本未出现在 2025、2026 '
                         '官方书目中；其内容与《语言学纲要》的语言本质、结构、演变部分高度重合，'
                         '故不单独占一个分段，也不重复计条目。'),
    'xat-xdh': dict(name='黄伯荣、李炜《现代汉语》（第二版）', pub='北京大学出版社',
                    edition='2016 年；2024、2025、2026 三处书目一致。'
                            '注意：不是黄伯荣、廖序东的同名教材',
                    ev='xatu官方', disc='语综', kind='official', note=None),
    'xat-gdh': dict(name='郭锡良《古代汉语》', pub='商务印书馆',
                    edition='2024 指定 2018 年本；2025、2026 为 1999 年修订本',
                    ev='xatu官方', disc='语综', kind='official',
                    note='版本提示：2025 年卷（816 段）对应的书目为 1999 年修订本，'
                         '与 2024 年（902 段）的 2018 年本不同，评卷须按各自年份版本理解。'),
    'xat-wxll': dict(name='童庆炳《文学理论教程》（第五版）', pub='高等教育出版社',
                     edition='2015 年版；2024、2025、2026 三处书目一致',
                     ev='xatu官方', disc='评论', kind='official', note=None),
}

# 书单键 → v4 知识块键：同一知识块在不同年份换过教材时，条目来自同一块
ALIAS = {'xdh_hu': 'xdh', 'xdh5': 'xdh', 'gdh_wl': 'gdh'}


def in_range(year, seg):
    return seg['years'][0] <= year <= seg['years'][1]


def build():
    out = OrderedDict()
    manifest = []
    total_unique = 0
    total_inst = 0
    total_ev = 0
    problems = []

    for school, segs in SPEC.items():
        papers = {p['paper']: p for p in V4[school]}
        seg_objs = []
        for seg in segs:
            src = papers[seg['src']]
            seg_disc = seg['disc'] if isinstance(seg['disc'], tuple) else (seg['disc'],)
            books = []
            for bk in seg['books']:
                meta = BOOKS[bk]
                if bk == 'ah-none':
                    # 未查到官方书目的年份：不按书分块，收录本科目全部知识块条目
                    srcbook = dict(key='ah-none',
                                   topics=[t for b in src['books'] for t in b['topics']])
                else:
                    srcbook = next(b for b in src['books'] if b['key'] == ALIAS.get(bk, bk))
                topics = []
                for t in srcbook['topics']:
                    if t['discipline'] not in seg_disc:
                        continue
                    es = [e for e in t['evidence'] if in_range(e['year'], seg)]
                    if not es:
                        continue
                    tt = dict(t)
                    tt['evidence'] = sorted(es, key=lambda e: (e['year'], e['pid']))
                    tt['all_frequency'] = t['frequency']
                    # 去重：同段同条目内完全相同的证据（同一题面被该条目的多个小题各记一次）合并
                    seen, ded = {}, []
                    for e in tt['evidence']:
                        k = (e['exam'], e['pid'], str(e['pages']), e['text'].strip())
                        if k in seen:
                            ex = [x for x in (seen[k].get('excerpt'), e.get('excerpt')) if x]
                            if ex:
                                seen[k]['excerpt'] = '／'.join(dict.fromkeys(ex))
                            continue
                        seen[k] = e
                        ded.append(e)
                    tt['evidence'] = ded
                    tt['years'] = sorted({x['year'] for x in ded})
                    tt['frequency'] = len(tt['years'])
                    tt['recall_hits'] = sorted({x['year'] for x in ded if x.get('recall')})
                    topics.append(tt)
                sections = list(dict.fromkeys(t['section'] for t in srcbook['topics']))
                topics.sort(key=lambda t: (sections.index(t['section'])
                                           if t['section'] in sections else 99,
                                           -t['frequency'], t['title']))
                m = dict(meta)
                m['key'] = bk
                m['topics'] = topics
                m['local_max'] = len(topics)
                m['sections'] = sections
                m['years_seen'] = sorted({e['year'] for t in topics for e in t['evidence']})
                books.append(m)
            paper_years = sorted({y for b in books for y in b['years_seen']})
            if not paper_years:
                problems.append(f'{school}/{seg["key"]}: 无年份')
            stats = dict(topics=sum(b['local_max'] for b in books),
                         evidence=sum(len(t['evidence']) for b in books for t in b['topics']),
                         recall=sum(1 for b in books for t in b['topics']
                                    for e in t['evidence'] if e.get('recall')))
            # 段页范围 = 本段各年实际卷在打印复核版里的起止页并集（不手写）
            idx = {(a['year'], a['subject']): (a['start'], a['end']) for a in INDEX[school]}
            spans = []
            for b in books:
                for t in b['topics']:
                    for e in t['evidence']:
                        a = idx.get((e['year'], e.get('subject')))
                        if a:
                            spans.append(a)
            pages = sorted({int(e['pages']) for b in books for t in b['topics']
                            for e in t['evidence'] if str(e.get('pages', '')).isdigit()})
            seg_objs.append(dict(
                key=seg['key'], name=seg['name'], code=seg['code'], label=seg['label'],
                years=f'{paper_years[0]}—{paper_years[-1]}' if paper_years else '',
                year_from=seg['years'][0], year_to=seg['years'][1],
                all_years=paper_years,
                page=(f'p{min(x[0] for x in spans)}—p{max(x[1] for x in spans)}'
                      if spans else (seg['page'] or '—')),
                page_real=f'p{pages[0]}—p{pages[-1]}' if pages else '',
                disc='／'.join(seg_disc), note=seg['note'], src=seg['src'],
                source_note=EV[seg['ev']], ev_kind=seg['ev'], books=books, stats=stats))
            total_inst += stats['topics']
            total_ev += stats['evidence']
        # 唯一考点集：必须与 v4 完全一致（占位证据只影响出处，不影响条目集合）
        uniq = {(t['discipline'], t['section'], t['title'])
                for s in seg_objs for b in s['books'] for t in b['topics']}
        total_unique += len(uniq)
        v4_uniq = {(t['discipline'], t['section'], t['title'])
                   for p in V4[school] for b in p['books'] for t in b['topics']}
        if uniq != v4_uniq:
            problems.append(f'{school}: 唯一考点集与 v4 不一致，缺 {len(v4_uniq - uniq)}、'
                            f'多 {len(uniq - v4_uniq)}：{sorted(v4_uniq - uniq)[:3]} / '
                            f'{sorted(uniq - v4_uniq)[:3]}')
        # 分册条目：按来源册合并各段后取并集，应等于 v4 的册内条目数
        for p in V4[school]:
            base = len({(t['discipline'], t['section'], t['title'])
                        for b in p['books'] for t in b['topics']})
            got = len({(t['discipline'], t['section'], t['title'])
                       for s in seg_objs if s['src'] == p['paper']
                       for b in s['books'] for t in b['topics']})
            if got != base:
                problems.append(f'{school}/{p["paper"]}: 分册条目 {got} ≠ v4 {base}')
        # 证据覆盖：v4 每条非占位证据都必须在 v5 至少出现一次
        v4_ids, v4_dup = set(), 0
        for p in V4[school]:
            for b in p['books']:
                for t in b['topics']:
                    seen = set()
                    for e in t['evidence']:
                        if e['pid'] in SYNTH:
                            continue
                        k = (e['exam'], e['pid'], e['text'].strip())
                        if k in seen:
                            v4_dup += 1
                            continue
                        seen.add(k)
                        v4_ids.add(k)
        v5_ids = {(e['exam'], e['pid'], e['text'].strip())
                  for s in seg_objs for b in s['books'] for t in b['topics'] for e in t['evidence']}
        missing = v4_ids - v5_ids
        if missing:
            problems.append(f'{school}: v4 证据在 v5 中缺失 {len(missing)} 条：{sorted(missing)[:2]}')
        inst = sum(s['stats']['topics'] for s in seg_objs)
        ev = sum(s['stats']['evidence'] for s in seg_objs)
        out[school] = dict(school=school, segments=seg_objs,
                           stats=dict(unique_topics=len(uniq), instances=inst, evidence=ev,
                                      unique_evidence=len(v5_ids),
                                      recall=sum(s['stats']['recall'] for s in seg_objs),
                                      segments=len(seg_objs),
                                      books=len({b['key'] for s in seg_objs for b in s['books']}),
                                      v4_evidence=len(v4_ids), v4_dup_removed=v4_dup))
        manifest.append(dict(file=f'{school}_高频考点_{VERSION}.docx', school=school,
                             segments=[dict(key=s['key'], label=s['label'], years=s['years'],
                                            topics=s['stats']['topics'],
                                            evidence=s['stats']['evidence'],
                                            recall=s['stats']['recall'],
                                            books=[b['name'] for b in s['books']])
                                       for s in seg_objs],
                             unique_topics=len(uniq), instances=inst, evidence=ev,
                             unique_evidence=len(v5_ids),
                             recall=sum(s['stats']['recall'] for s in seg_objs)))

    (W / f'{VERSION}考点数据.json').write_text(json.dumps(out, ensure_ascii=False, indent=1))
    (W / f'{VERSION}分段清单.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=1))
    print(f'===== {VERSION} 分册统计（一校一册）')
    for m in manifest:
        print(f'  {m["school"]}: {len(m["segments"])} 段　唯一考点 {m["unique_topics"]}　'
              f'段内条目 {m["instances"]}　证据 {m["evidence"]}　回忆版 {m["recall"]}')
        for s in m['segments']:
            print(f'      - {s["label"]}　{s["years"]}　条目 {s["topics"]}　证据 {s["evidence"]}　'
                  f'书 {len(s["books"])} 本')
    print(f'  合计：册 {len(manifest)}　唯一考点 {total_unique}　段内条目 {total_inst}　证据 {total_ev}')
    if problems:
        print('\n  !! 自检问题：')
        for x in problems:
            print('   -', x)
    else:
        print('  自检通过：唯一考点、分册条目、证据覆盖三项与 v4 一致。')
    return problems


if __name__ == '__main__':
    # 自检问题只打印不改变退出码：数据文件照常落盘，由 核验v5.py／核对v5真实性.py 把关。
    build()
    sys.exit(0)
