# /// script
# requires-python = ">=3.10"
# dependencies = ["openpyxl>=3.1"]
# ///
"""Write the Dir' Public Safety Authority file pack (labs/files/) and the reference answers
(labs/solutions/reference/*.json and labs/solutions/D*_*.md).

    uv run labs/_tools/prepare_files.py

Deterministic: no randomness and fixed file times, so the same bytes on every run. Everything is
synthetic. The case itself (systems, records, roles, policy, gates) is in case.py and
case_docs.py; labs/solutions/check.py checks every key against the written files.
"""
from __future__ import annotations

import json
import re
import sys
import zipfile
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

sys.path.insert(0, str(Path(__file__).resolve().parent))
import case as C  # noqa: E402
import case_docs as K  # noqa: E402

LABS = Path(__file__).resolve().parents[1]
OUT = LABS / "files"
SOL = LABS / "solutions"
REF = SOL / "reference"
FIXED = datetime(2026, 11, 1, 9, 0, 0)
HEAD_FILL = PatternFill("solid", fgColor="1C355E")
FILL_FILL = PatternFill("solid", fgColor="FFF7E0")


# ------------------------------------------------------------------ helpers
def write_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes((text.strip() + "\n").encode("utf-8"))


def put_ref(name: str, data) -> None:
    REF.mkdir(parents=True, exist_ok=True)
    (REF / name).write_bytes((json.dumps(data, ensure_ascii=False, indent=1) + "\n").encode("utf-8"))


def sheet(wb, title, header, rows, fill_cols=(), widths=None, validations=None):
    ws = wb.create_sheet(title)
    ws.sheet_view.rightToLeft = True
    ws.append(header)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = HEAD_FILL
        c.alignment = Alignment(wrap_text=True, vertical="top")
    for r in rows:
        ws.append(list(r))
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
            if c.column - 1 in fill_cols:
                c.fill = FILL_FILL
    for i, w in enumerate(widths or [18] * len(header)):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i + 1)].width = w
    ws.freeze_panes = "A2"
    for col, values in (validations or {}).items():
        dv = DataValidation(type="list", formula1='"' + ",".join(values) + '"', allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f"{col}2:{col}{max(2, len(rows) + 1)}")
    return ws


def save(wb, rel: str) -> None:
    del wb[wb.sheetnames[0]]
    wb.properties.creator = C.AUTHORITY
    wb.properties.lastModifiedBy = C.AUTHORITY
    wb.properties.created = wb.properties.modified = FIXED
    path = OUT / rel
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".tmp")
    wb.save(tmp)
    with zipfile.ZipFile(tmp) as z:
        items = [(i.filename, z.read(i.filename)) for i in z.infolist()]
    stamp = FIXED.strftime("%Y-%m-%dT%H:%M:%SZ").encode()
    items = [(n, re.sub(rb"(<dcterms:(?:created|modified)[^>]*>)[^<]*(</dcterms)", rb"\g<1>" + stamp + rb"\g<2>", d))
             if n == "docProps/core.xml" else (n, d) for n, d in items]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in items:
            info = zipfile.ZipInfo(name, date_time=FIXED.timetuple()[:6])
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, data)
    tmp.unlink()


def note_sheet(wb, lines):
    ws = wb.create_sheet("اقرأني")
    ws.sheet_view.rightToLeft = True
    for ln in [C.NOTE] + lines:
        ws.append([ln])
    ws.column_dimensions["A"].width = 120
    return ws


def table(cols, rows):
    out = ["| " + " | ".join(cols) + " |", "|" + "|".join("---" for _ in cols) + "|"]
    out += ["| " + " | ".join(str(x) for x in r) + " |" for r in rows]
    return "\n".join(out)


def sol(name, title, start, body):
    write_text(SOL / name, f"# الحل المرجعي: المختبر {title}\n\nنقطة البداية: {start}. التحقق: `uv run labs/solutions/check.py`.\n\n{body}")


LV = C.LEVEL_NAMES
INV_FIELDS = ["رقم النظام", "اسم النظام", "الإدارة", "المالك (المنصب)", "الغرض", "المستخدمون والمتأثرون", "البيانات",
              "تصنيف البيانات", "مصدر النموذج", "دور النظام في القرار", "المتأثرون سنويًا", "أثر الخطأ وإمكانية تداركه",
              "مرحلة دورة الحياة", "مصرّح عنه", "مصدر الاكتشاف", "مستوى المخاطر", "مبرر المستوى", "آخر تأكيد من المالك"]
assert len(INV_FIELDS) == 18
FOUND_AR = {"interview": "مقابلة", "procurement": "سجل المشتريات", "expense": "مطالبات المصروفات", "network": "سجل الشبكة"}


def inv_row(s, with_tier=False, confirmed=""):
    return [s["id"], s["name"], C.DEPTS[s["dept"]][0], s["owner"], s["purpose"], s["users"], s["data"], s["data_class"],
            C.SOURCES[s["source"]], C.AUTONOMY_TEXT[s["role"]], s["people"], s["reversibility"], C.STAGES[s["stage"]],
            "نعم" if s["found"] == "interview" else "لا", FOUND_AR[s["found"]],
            LV[s["level"]] if with_tier else "", s["why"] if with_tier else "", confirmed]


# ------------------------------------------------------------------ the brief
def case_brief():
    people = "\n".join(f"- {n}: {r}" for n, r in C.PEOPLE.values())
    depts = "\n".join(f"- {d[0]}" for d in C.DEPTS.values())
    write_text(OUT / "case_brief.txt", f"""
{C.AUTHORITY} · موجز الحالة
{C.NOTE}

من نحن
هيئة دِرع جهة عامة افتراضية مسؤولة عن السلامة العامة في منطقة افتراضية: تفتيش السلامة من الحرائق في المنشآت،
وإصدار شهادات السلامة للمباني والفعاليات، ومركز للبلاغات والعمليات يستقبل البلاغات الطارئة ويوجه الفرق،
وحملات التوعية. يعمل فيها نحو 4,000 موظف.

لماذا الآن
أنشأت الهيئة هذا العام مكتبًا لحوكمة الذكاء الاصطناعي ولجنة لها. تقدّر الإدارة العليا أن في الهيئة نحو ثلاثين نظام ذكاء
اصطناعي قيد الاستخدام أو التطوير، لكن لا أحد يعرف قائمتها الكاملة، ولا من يملك كلًّا منها، ولا أيها يحتاج موافقة اللجنة.
تريد د. ريما الحقباني في نهاية الشهر «حزمة حوكمة تأسيسية» تعرضها على اللجنة: سجل الأنظمة، وتصنيف مخاطرها، ومصفوفة
المسؤوليات، وسياسة الاستخدام المقبول، ومواصفات بوابات المراجعة. أنت تعمل في مكتبها.

الإدارات
{depts}

الأشخاص
{people}

قاعدة المقرر
كل الملفات اصطناعية. لا ترفع ملفًا من جهة عملك الحقيقية إلى أي أداة. في المختبرات يُستخدم المساعد مرشدًا يناقش استدلالك،
لا ليعطيك مستويات المخاطر، وفي المحاولة الأولى للتصنيف (المختبر 2.1) لا يُستخدم الذكاء الاصطناعي.
""")


# ------------------------------------------------------------------ day 1
def lab_1_1():
    wb = openpyxl.Workbook()
    note_sheet(wb, ["المختبر 1.1: صنّف كل عبارة في طبقتها، ثم افحص كل ضابط: هل له مالك ومحفّز وتكرار واختبار؟",
                    "الطبقات: " + "، ".join(K.LAYERS)])
    rows = [(i, src, txt, "", "", "") for i, src, txt, _, _ in K.STATEMENTS]
    sheet(wb, "العبارات", ["الرقم", "المصدر", "العبارة", "الطبقة", "إن كان ضابطًا: ما الناقص", "إعادة الصياغة"], rows,
          fill_cols=(3, 4, 5), widths=[8, 24, 70, 14, 28, 50], validations={"D": K.LAYERS})
    save(wb, "governance_statements.xlsx")
    key = [{"id": i, "layer": lay, "missing": miss} for i, _, _, lay, miss in K.STATEMENTS]
    put_ref("lab_1_1.json", key)
    rows = [(i, txt, lay, "، ".join(miss) if miss else ("قابل للتشغيل" if lay == "ضابط" else ""))
            for i, _, txt, lay, miss in K.STATEMENTS]
    sol("D1_1_layers.md", "1.1 · ست طبقات وضابط يعمل", "`files/governance_statements.xlsx`", f"""
{table(["الرقم", "العبارة", "الطبقة", "الضابط: الناقص"], rows)}

العدد في كل طبقة: مبدآن، وثلاث سياسات، ومعياران، وثلاثة إجراءات، وستة ضوابط، ودليلان.

الضوابط غير القابلة للتشغيل: G13 وG14 وG15 وG16. مثالان لإعادة الصياغة:

- G16: «يعرض كل روبوت محادثة عام عبارة إفصاح في أول رسالة. يختبرها مكتب حوكمة الذكاء الاصطناعي في بداية كل شهر بمحادثة تجريبية ويحفظ لقطة الشاشة.»
- G14: «قبل كل إعادة تدريب يشغّل مالك النظام فحوص الاكتمال والصلاحية المعتمدة على بيانات التدريب، ولا تبدأ إعادة التدريب إذا نقص الاكتمال عن 98%، ويُحفظ تقرير الفحص في ملف النظام.»

الفخ: G10 إجراء وليس ضابطًا (يصف خطوات طلب، لا فحصًا متكررًا يثبت أن شيئًا تحقق). وG17 وG18 دليلان: ما يبقى بعد أن يعمل الضابط.
""")


def lab_1_2():
    wb = openpyxl.Workbook()
    note_sheet(wb, ["المختبر 1.2: لكل مبدأ اختر ضابطًا من الكتالوج، وأكمل المالك والمحفّز والتكرار والاختبار والدليل، ثم اربطه بمبدأ من مبادئ OECD ووظيفة من وظائف إطار NIST."])
    sheet(wb, "الضوابط", ["الرقم", "الضابط", "المالك", "المحفّز", "التكرار", "الاختبار", "الدليل"],
          [list(c) for c in K.CONTROLS], widths=[8, 60, 26, 22, 14, 34, 26])
    sheet(wb, "مبادئ OECD", ["الرمز", "المبدأ"], K.OECD, widths=[8, 70])
    sheet(wb, "وظائف NIST", ["الوظيفة", "ما تعنيه"], K.NIST, widths=[12, 70])
    save(wb, "control_catalogue.xlsx")
    wb = openpyxl.Workbook()
    note_sheet(wb, ["الخريطة من صفحة واحدة. المعيار: خمسة مبادئ من سبعة على الأقل لكل منها ضابط له مالك ومحفّز واختبار."])
    sheet(wb, "الخريطة", ["الرمز", "المبدأ الوطني", "رقم الضابط", "المالك", "المحفّز", "التكرار", "الاختبار", "الدليل",
                          "مبدأ OECD", "وظيفة NIST"],
          [(p, n, "", "", "", "", "", "", "", "") for p, n in K.PRINCIPLES], fill_cols=range(2, 10),
          widths=[6, 26, 10, 22, 20, 12, 28, 22, 10, 12],
          validations={"C": [c[0] for c in K.CONTROLS], "I": [o[0] for o in K.OECD], "J": [n[0] for n in K.NIST]})
    save(wb, "principle_control_map.xlsx")
    put_ref("lab_1_2.json", {"operable": K.OPERABLE_CONTROLS,
                             "map": {p: {"controls": c, "oecd": o, "nist": n} for p, (c, o, n) in K.MAP_KEY.items()}})
    ctl = {c[0]: c for c in K.CONTROLS}
    rows = []
    for p, name in K.PRINCIPLES:
        cs, o, n = K.MAP_KEY[p]
        c = ctl[cs[0]]
        rows.append((name, " أو ".join(cs), c[2], c[3], c[5], c[6], " أو ".join(o), " أو ".join(n)))
    sol("D1_2_principle_map.md", "1.2 · خريطة المبادئ والضوابط", "`files/control_catalogue.xlsx` و`files/principle_control_map.xlsx`", f"""
{table(["المبدأ", "الضابط", "المالك", "المحفّز", "الاختبار", "الدليل", "OECD", "NIST"], rows)}

الضوابط غير القابلة للتشغيل في الكتالوج: CC03 وCC07 (شعاران)، وCC10 (إطار مرجعي وليس ضابطًا)، وCC11 (لا محفّز ولا اختبار ولا دليل).
CC12 مقياس يُرفع إلى اللجنة، ويقبل بديلًا لمبدأ المساءلة.

معيار النجاح في التقييم: خمسة مبادئ من سبعة على الأقل لكل منها ضابط له مالك ومحفّز واختبار. الخلاف في عمودي OECD وNIST
لا يُحسب خطأ إذا كان التعليل معقولًا: الإطاران سند، لا مصدر الضابط.
""")


def interviews_text():
    parts = [f"{C.AUTHORITY} · ملاحظات مقابلات حصر أنظمة الذكاء الاصطناعي (أكتوبر 2026)", C.NOTE,
             "كتبها فريق مكتب حوكمة الذكاء الاصطناعي بعد مقابلة كل إدارة. الملاحظات كما دُوِّنت، ولم يُتحقق منها بعد."]
    for i, d in enumerate(K.INTERVIEW_ORDER, 1):
        parts.append(f"\n{i}. {C.DEPTS[d][0]}\n{K.INTERVIEW_INTRO[d]}")
        for s in C.S:
            if s["dept"] == d and s["found"] == "interview":
                stage = C.STAGES[s["stage"]]
                parts.append(f"- «{s['name']}»: {s['purpose']}. يستخدمه أو يتأثر به: {s['users']}. البيانات: {s['data']}. "
                             f"مصدره: {C.SOURCES[s['source']]}. المرحلة: {stage}. في القرار: {C.AUTONOMY_TEXT[s['role']]}. "
                             f"نحو {s['people']:,} شخص في السنة. إذا أخطأ: {s['reversibility']}.")
        for x, xd, name, desc in K.NOT_AI:
            if xd == d:
                parts.append(f"- «{name}»: ذكروه ضمن «أنظمة الذكاء الاصطناعي». {desc}.")
        for sid, dd, alias in K.DUPLICATES:
            if dd == d:
                s = C.SYS[sid]
                parts.append(f"- «{alias}»: قالوا إنه يعمل داخل منصة تشغلها تقنية المعلومات لصالح {C.DEPTS[s['dept']][0]}. {s['purpose']}.")
        if d == "AUD":
            parts.append("- لم تذكر الإدارة أي نظام.")
    write_text(OUT / "interview_notes.txt", "\n".join(parts))


def lab_1_3():
    interviews_text()
    wb = openpyxl.Workbook()
    note_sheet(wb, ["قالب سجل أنظمة الذكاء الاصطناعي: صف واحد لكل نظام، لا لكل عقد ولا لكل إدارة.",
                    "الحقول الثمانية عشر إلزامية عند اكتمال السجل؛ مستوى المخاطر ومبرره يُكتبان في اليوم الثاني."])
    sheet(wb, "السجل", INV_FIELDS, [], fill_cols=range(18), widths=[8, 26, 22, 22, 40, 26, 30, 10, 18, 26, 10, 30, 12, 8, 14, 16, 30, 12])
    save(wb, "inventory_template.xlsx")
    declared = [s for s in C.S if s["found"] == "interview"]
    put_ref("lab_1_3.json", {"declared": [s["id"] for s in declared], "not_ai": [x[2] for x in K.NOT_AI],
                             "duplicates": {a: sid for sid, _, a in K.DUPLICATES}})
    rows = [(s["id"], s["name"], s["owner"], C.SOURCES[s["source"]], C.STAGES[s["stage"]]) for s in declared]
    sol("D1_3_inventory_interviews.md", "1.3 · السجل من المقابلات", "`files/interview_notes.txt` و`files/inventory_template.xlsx`", f"""
تذكر الملاحظات 31 بندًا: 25 نظامًا مختلفًا، وأربعة بنود ليست ذكاءً اصطناعيًا، وبندان مكرران.

{table(["الرقم", "النظام", "المالك (المنصب)", "مصدر النموذج", "المرحلة"], rows)}

ليست أنظمة ذكاء اصطناعي: {"، ".join("«" + x[2] + "»" for x in K.NOT_AI)}. كلها قواعد أو صيغ ثابتة لا تتعلم ولا تستنتج.

مكرران: «{K.DUPLICATES[0][2]}» هو S01، و«{K.DUPLICATES[1][2]}» هو S25. المالك هو الإدارة التي تقرر بالمخرجات، لا تقنية المعلومات التي تشغّل المنصة.

يدخل المقترحان S13 وS17 السجل وهما في مرحلة الاستقبال: السجل يعدّ ما يُستخدم وما يُطوَّر وما يُقترح.

هذه القائمة نفسها هي `files/inventory_start.xlsx`، نقطة البداية للمختبر 1.4.
""")


def lab_1_4():
    declared = [s for s in C.S if s["found"] == "interview"]
    wb = openpyxl.Workbook()
    note_sheet(wb, ["السجل بعد المقابلات: 25 نظامًا مصرحًا عنها. أضف ما تجده في سجلات المشتريات والمصروفات والشبكة."])
    sheet(wb, "السجل", INV_FIELDS, [inv_row(s) for s in declared], fill_cols=(),
          widths=[8, 26, 22, 22, 40, 26, 30, 10, 18, 26, 10, 30, 12, 8, 14, 16, 30, 12])
    save(wb, "inventory_start.xlsx")
    wb = openpyxl.Workbook()
    note_sheet(wb, ["أوامر الشراء من يناير إلى سبتمبر 2026 (مقتطف)."])
    sheet(wb, "المشتريات", ["أمر الشراء", "المورّد", "البند", "الإدارة", "القيمة (ريال)", "التاريخ", "ملاحظات"],
          [(p[0], p[1], p[2], C.DEPTS[p[3]][0], p[4], p[5], p[6]) for p in K.PROCUREMENT], widths=[12, 28, 56, 26, 12, 12, 50])
    save(wb, "procurement_records.xlsx")
    wb = openpyxl.Workbook()
    note_sheet(wb, ["مطالبات المصروفات المعتمدة من يوليو إلى سبتمبر 2026 (مقتطف)."])
    sheet(wb, "المصروفات", ["رقم المطالبة", "صفة الموظف", "الإدارة", "البند", "المبلغ (ريال)", "الشهر"],
          [(e[0], e[1], C.DEPTS[e[2]][0], e[3], e[4], e[5]) for e in K.EXPENSES], widths=[12, 18, 26, 60, 12, 10])
    save(wb, "expense_claims.xlsx")
    wb = openpyxl.Workbook()
    note_sheet(wb, ["ملخص بوابة الوكيل (proxy) لشهر سبتمبر 2026: الطلبات الصادرة من أجهزة الهيئة مجمعة حسب الإدارة والنطاق.",
                    "النطاقات المنتهية بـ .example أسماء وهمية."])
    sheet(wb, "الشبكة", ["الإدارة", "النطاق", "الفئة بحسب البوابة", "الطلبات في الشهر", "البيانات المرفوعة (ميغابايت)"],
          [(C.DEPTS[n[0]][0], n[1], n[2], n[3], n[4]) for n in K.NETWORK], widths=[26, 30, 36, 14, 18])
    save(wb, "network_summary.xlsx")
    put_ref("lab_1_4.json", {"undeclared": {sid: {"source": k, "record": r} for sid, (k, r) in K.FOUND_IN.items()},
                             "total": len(C.S)})
    rows = [(sid, C.SYS[sid]["name"], FOUND_AR[k], r, C.SYS[sid]["owner"]) for sid, (k, r) in K.FOUND_IN.items()]
    sol("D1_4_undeclared.md", "1.4 · البحث عن الخمسة", "`files/inventory_start.xlsx` وسجلات المشتريات والمصروفات والشبكة", f"""
{table(["الرقم", "النظام", "وُجد في", "السطر", "المالك (المنصب)"], rows)}

بعد الإضافة: 30 نظامًا، خمسة منها غير مصرح عنها (اثنان في المشتريات، واثنان في الشبكة، وواحد في المصروفات).

الفخاخ:
- الإضاءة الذكية والطفايات المتصلة وبطاقات الدخول «الذكية» والشاحن «الذكي»: كلمة «ذكي» ليست ذكاءً اصطناعيًا.
- رسوم المؤتمر والكتاب: نفقات عن الذكاء الاصطناعي، لا أنظمة.
- حركة الشبكة إلى منصة الاجتماعات (IT) لا تكشف وحدها أن خاصية التفريغ مفعلة؛ الدليل في أمر الشراء PO-26-0302.
- إدارة الموارد البشرية قالت في المقابلة إنها لا تستخدم الذكاء الاصطناعي في التوظيف، وأمر الشراء PO-26-0151 يقول غير ذلك.
- سجل الشبكة يثبت الاستخدام ولا يثبت الغرض: الغرض والمالك يُؤكَّدان بسؤال الإدارة.

المالك في كل حالة هو مدير الإدارة التي تقرر بمخرجات النظام؛ لخاصية الاجتماعات المالك مدير تقنية المعلومات لأنها تخدم كل الإدارات.
""")


# ------------------------------------------------------------------ day 2
RUBRIC = """
{auth} · سلم تصنيف مخاطر أنظمة الذكاء الاصطناعي (الإصدار 1، معتمد من لجنة الحوكمة)
{note}

هذا السلم أداة الهيئة الداخلية لتطبيق مستويات المخاطر الأربعة في مبادئ أخلاقيات الذكاء الاصطناعي الصادرة عن سدايا.
المستويات وطنية؛ العوامل الخمسة وقاعدة القرار من وضع الهيئة.

العوامل (درجة من 1 إلى 3 لكل عامل)
1. الأثر في الناس (I)
   1: لا يمس حقوق الأفراد أو سلامتهم أو مصالحهم المباشرة.
   2: يؤثر في خدمة أو معاملة لشخص، ويمكن تداركه.
   3: قد يمس سلامة الناس أو مصدر رزقهم أو حقًا من حقوقهم.
2. الاستقلالية (A)
   1: {a1}.
   2: {a2}.
   3: {a3}.
3. النطاق (S): عدد المتأثرين في السنة
   1: أقل من 5,000.   2: من 5,000 إلى 100,000.   3: أكثر من 100,000.
4. صعوبة التراجع (R)
   1: يُصحح الخطأ بسهولة ودون ضرر باقٍ.
   2: يُصحح بجهد أو عبر تظلم ومراجعة، مع ضرر مؤقت.
   3: يصعب تداركه أو لا يمكن: تأخر استجابة، ضرر جسدي، فرصة لا تعود.
5. حساسية البيانات (D)
   1: بيانات غير شخصية مصنفة «عام».
   2: بيانات شخصية أو مصنفة «مقيد».
   3: بيانات شخصية حساسة (صحية، بيومترية، أمنية وجنائية، موقع دقيق لأشخاص) أو مصنفة «سري» فأعلى.

قاعدة القرار (طبّقها بالترتيب، وتوقف عند أول قاعدة تنطبق)
1. إن كان النظام في فئة محظورة: مخاطر غير مقبولة.
2. إن كان الأثر 3 والاستقلالية 2 أو أكثر، أو الأثر 3 وصعوبة التراجع 3: مخاطر عالية.
3. إن كان مجموع الدرجات 12 أو أكثر: مخاطر عالية.
4. إن كان المجموع 9 أو أكثر، أو الأثر 3، أو حساسية البيانات 3: مخاطر محدودة.
5. غير ذلك: مخاطر بسيطة أو منعدمة.

الفئات المحظورة في سياسة الهيئة
{prohibited}

ملاحظة: المستوى يصف النظام إذا عمل كما صُمم. مخالفة سياسة الاستخدام (مثل إدخال بيانات مقيدة في أداة خارجية) مسألة مستقلة
تُعالج بالسياسة، ولا ترفع المستوى ولا تخفضه.
"""


def lab_2_1():
    write_text(OUT / "tiering_rubric.txt", RUBRIC.format(auth=C.AUTHORITY, note=C.NOTE, a1=C.AUTONOMY_TEXT[1], a2=C.AUTONOMY_TEXT[2],
                                                         a3=C.AUTONOMY_TEXT[3], prohibited="\n".join(f"- {p}" for p in C.PROHIBITED)))
    wb = openpyxl.Workbook()
    note_sheet(wb, ["سجل الأنظمة الثلاثين بعد الحصر، بلا مستويات. صنّف كل نظام في ورقة «التصنيف» دون مساعد ذكاء اصطناعي.",
                    "إذا لم تكمل مختبرات اليوم الأول، فهذا الملف نقطة بدايتك."])
    sheet(wb, "السجل", INV_FIELDS, [inv_row(s) for s in C.S], widths=[8, 26, 22, 22, 40, 26, 30, 10, 18, 26, 10, 30, 12, 8, 14, 16, 30, 12])
    ws = sheet(wb, "التصنيف", ["رقم النظام", "اسم النظام", "محظور؟", "I", "A", "S", "R", "D", "المجموع", "المستوى", "المبرر في سطر واحد"],
               [(s["id"], s["name"], "", "", "", "", "", "", "", "", "") for s in C.S], fill_cols=(2, 3, 4, 5, 6, 7, 9, 10),
               widths=[8, 30, 8, 5, 5, 5, 5, 5, 8, 20, 50],
               validations={"C": ["نعم", "لا"], "J": [LV[x] for x in C.LEVELS]})
    for r in range(2, len(C.S) + 2):
        ws.cell(row=r, column=9, value=f"=IF(COUNT(D{r}:H{r})=5,SUM(D{r}:H{r}),\"\")")
    save(wb, "inventory_day2.xlsx")
    put_ref("lab_2_1.json", {s["id"]: {"f": s["f"], "prohibited": s["prohibited"], "level": s["level"]} for s in C.S})
    rows = [(s["id"], s["name"], *[s["f"][k] for k in C.FACTORS], sum(s["f"].values()), LV[s["level"]], s["why"]) for s in C.S]
    counts = {lv: sum(1 for s in C.S if s["level"] == lv) for lv in C.LEVELS}
    sol("D2_1_tiering.md", "2.1 · تصنيف ثلاثين نظامًا", "`files/inventory_day2.xlsx` و`files/tiering_rubric.txt`", f"""
المحاولة الأولى دون ذكاء اصطناعي. لا تُعرض هذه الصفحة على المشاركين قبل الخطوة 4 من المختبر.

{table(["الرقم", "النظام", "I", "A", "S", "R", "D", "المجموع", "المستوى", "المبرر"], rows)}

التوزيع: غير مقبولة {counts['U']}، عالية {counts['H']}، محدودة {counts['L']}، بسيطة أو منعدمة {counts['N']}.

معيار النجاح في التقييم: يتفق المستوى مع المرجع في 24 نظامًا من 30 على الأقل (80%)، وكل اختلاف معلل كتابةً.
الخلاف في درجة عامل واحد لا يهم إن لم يغيّر المستوى.

حالات تختلف فيها القاعة عادة: S02 (المجموع 13 مع أثر 2)، وS06 (يقول بعضهم إن المشرف يقرر)، وS10 (رفض آلي يؤخر رزقًا)،
وS18 (مرشح البريد يتصرف وحده لكنه بسيط المخاطر، كما في مثال المبادئ الوطنية)، وS21 (المخالفة لا ترفع المستوى).
""")


def lab_2_2():
    lines = [f"{C.AUTHORITY} · ثلاث حالات حدّية وتغيير في الوقائع", C.NOTE, ""]
    extra = {
        "S02": ["الملخص يُحفظ في سجل الحادث آليًا بعد انتهاء المكالمة، ولا يُطلب من الموظف مراجعته.",
                "يعتمد محققو الحوادث على الملخص لاحقًا، ويعودون إلى التسجيل الصوتي إذا شكوا فقط.",
                "في التجربة الحالية تُرسل التسجيلات إلى خدمة خارجية لتحويل الكلام إلى نص.",
                "قال مدير المركز: «هي مجرد ملاحظات، القرار في المكالمة نفسها»."],
        "S06": ["يُصدَر جدول التفتيش السنوي من ترتيب النموذج كما هو؛ يستطيع المشرف إضافة منشأة بعد بلاغ، ونادرًا ما يفعل.",
                "المنشآت في آخر الترتيب قد لا تُفتَّش مرتين متتاليتين.",
                "لا يعلم أصحاب المنشآت أن النموذج يحدد موعد التفتيش.",
                "قال مدير التفتيش: «المفتش هو الذي يقرر المخالفة، النموذج يرتب فقط»."],
        "S10": ["الطلب الذي ينقصه مستند يُرفض آليًا وتصل صاحب المنشأة رسالة بالرفض.",
                "يعيد صاحب المنشأة التقديم، ويستغرق ذلك ثلاثة أسابيع في المتوسط.",
                "لا تُفتح المنشأة دون شهادة السلامة.",
                "رصد القسم أن 6% من حالات الرفض كانت لمستندات موجودة لكن بصيغة لم يقرأها النظام."],
    }
    for sid in C.BORDERLINE:
        s = C.SYS[sid]
        lines.append(f"الحالة {sid}: {s['name']}")
        lines += [f"- {x}" for x in extra[sid]]
        lines.append("")
    lines.append("تغيير الوقائع (كما تفعل اللجنة في المناقشة)")
    for sid, k, v, fact in C.CHANGED:
        lines.append(f"- {sid} ({C.SYS[sid]['name']}): {fact}. أعد التصنيف، وسمِّ العامل الذي تغير.")
    write_text(OUT / "borderline_cases.txt", "\n".join(lines))
    ch = []
    for sid, k, v, fact in C.CHANGED:
        f = dict(C.SYS[sid]["f"]); f[k] = v
        ch.append({"id": sid, "factor": k, "value": v, "before": C.SYS[sid]["level"], "after": C.tier(f, False)})
    put_ref("lab_2_2.json", {"borderline": {sid: C.SYS[sid]["level"] for sid in C.BORDERLINE}, "changed": ch})
    sol("D2_2_borderline.md", "2.2 · ثلاث حالات حدّية", "`files/borderline_cases.txt` و`files/inventory_day2.xlsx`", f"""
## الحالات الثلاث

- S02 تفريغ المكالمات: **مخاطر عالية**. الأثر 2 لا 3 (الملخص لا يقرر الاستجابة)، لكن الاستقلالية 3 (يُحفظ دون مراجعة)،
  والنطاق 3، والبيانات 3 (صحية وموقع). المجموع 13 فالقاعدة 3. حجة مدير المركز تخلط بين القرار في المكالمة والقرار بالسجل لاحقًا.
- S06 أولوية التفتيش: **مخاطر عالية** بالقاعدة 2. الأثر 3 (سلامة رواد المنشآت)، والاستقلالية 2 لا 1: الجدول يصدر من الترتيب كما هو،
  وإضافة المشرف استثناء نادر. «المفتش يقرر المخالفة» صحيح، لكن النموذج يقرر من يُفتَّش.
- S10 فحص الطلبات: **مخاطر عالية** بالقاعدة 2. الرفض يؤخر افتتاح منشأة (مصدر رزق: أثر 3)، والاستقلالية 3. نسبة 6% تعني أن الخطأ واقع فعلًا.

## تغيير الوقائع

{table(["النظام", "العامل", "القيمة الجديدة", "قبل", "بعد"], [(c["id"], c["factor"], c["value"], LV[c["before"]], LV[c["after"]]) for c in ch])}

S10 بعد التعديل: النظام ينبّه إلى الوثائق الناقصة ولا يوصي بالرفض، والموظف يفحص ويقرر، فالاستقلالية 1 بنص الدرجة 1 في السلم.
الأثر 3 والاستقلالية 1 وصعوبة التراجع 2، فلا تنطبق القاعدة 2؛ المجموع 10 فهي محدودة. لو بقي النظام يوصي بالرفض لكانت الاستقلالية 2 والمستوى عاليًا.
S05 بعد ربطه بالتوجيه: الأثر 3 والاستقلالية 3، فهي عالية. البوابات التي تتغير: S05 يحتاج موافقة اللجنة على النشر (G2) وإعادة الموافقة الربعية (G3).
""")


def lab_2_3():
    write_text(OUT / "org_chart.txt", f"{C.AUTHORITY}\n{C.NOTE}\n" + K.ORG_CHART)
    wb = openpyxl.Workbook()
    note_sheet(wb, ["المختبر 2.3: لكل قرار ضع A في دور واحد فقط (المساءل)، وR في دور أو أكثر (المسؤول عن التنفيذ).",
                    "المختبر 2.4: أكمل C (يُستشار) وI (يُبلَّغ)."])
    sheet(wb, "المصفوفة", ["الرقم", "القرار"] + [r[1] for r in K.ROLES], [(d[0], d[1]) + ("",) * len(K.ROLES) for d in K.DECISIONS],
          fill_cols=range(2, 2 + len(K.ROLES)), widths=[6, 44] + [12] * len(K.ROLES))
    save(wb, "decisions.xlsx")
    put_ref("lab_2_3.json", {d[0]: {"A": d[2], "R": d[3], "C": d[4], "I": d[5]} for d in K.DECISIONS})
    rows = [(d[0], d[1], K.ROLE[d[2]], "، ".join(K.ROLE[r] for r in d[3])) for d in K.DECISIONS]
    sol("D2_3_accountable.md", "2.3 · المساءل والمسؤول", "`files/org_chart.txt` و`files/decisions.xlsx`", f"""
{table(["الرقم", "القرار", "المساءل (A)", "المسؤول (R)"], rows)}

القواعد التي تحسم الخلافات:
- دور مساءل واحد لكل قرار. «اللجنة ومكتب الحوكمة» معًا ليسا إجابة.
- D09: الإيقاف بعد حادث قرار مالك النظام، لأن اللجنة لا تجتمع في ساعات؛ اللجنة تُبلَّغ.
- D10: اعتماد أدوات الذكاء الاصطناعي وتفعيلها لمكتب إدارة البيانات، كما في الهيكل وفي مبادئ الذكاء الاصطناعي التوليدي للجهات الحكومية.
- D12: رئيس الهيئة يعتمد خطة أخلاقيات الذكاء الاصطناعي وتقريرها السنوي، كما في مبادئ أخلاقيات الذكاء الاصطناعي.
- المراجعة الداخلية لا تكون A ولا R في أي قرار، رغم أنها طلبت ذلك في المقابلة: هي خط الدفاع الثالث.

مقبول في R: إضافة تقنية المعلومات إلى D05 وD08، أو مكتب الحوكمة إلى D02. غير مقبول: تغيير A.
""")


def lab_2_4():
    wb = openpyxl.Workbook()
    note_sheet(wb, ["نقطة البداية: A وR من المرجع. أكمل C وI، ثم تتبع الحوادث الثلاث في incidents.txt عبر مسار التصعيد في org_chart.txt."])
    rows = []
    for d in K.DECISIONS:
        cells = []
        for rid, _ in K.ROLES:
            cells.append("A" if rid == d[2] else ("R" if rid in d[3] else ""))
        rows.append([d[0], d[1]] + cells)
    sheet(wb, "المصفوفة", ["الرقم", "القرار"] + [r[1] for r in K.ROLES], rows, fill_cols=range(2, 2 + len(K.ROLES)),
          widths=[6, 44] + [12] * len(K.ROLES))
    save(wb, "raci_start.xlsx")
    wb = openpyxl.Workbook()
    note_sheet(wb, ["المصفوفة المرجعية الكاملة: نقطة البداية للمختبر 3.3 ولحزمة المشروع الختامي."])
    rows = []
    for d in K.DECISIONS:
        cells = []
        for rid, _ in K.ROLES:
            cells.append("A" if rid == d[2] else "R" if rid in d[3] else "C" if rid in d[4] else "I" if rid in d[5] else "")
        rows.append([d[0], d[1]] + cells)
    sheet(wb, "المصفوفة", ["الرقم", "القرار"] + [r[1] for r in K.ROLES], rows, widths=[6, 44] + [12] * len(K.ROLES))
    save(wb, "raci_reference.xlsx")
    write_text(OUT / "incidents.txt", f"{C.AUTHORITY} · ثلاثة حوادث\n{C.NOTE}\n\n" +
               "\n\n".join(f"{i}: {t}\nتتبع الحادث: من يلاحظ، ومن يقرر الإيقاف، ومن يُبلَّغ ومتى، وهل يلزم اجتماع طارئ للجنة؟" for i, t, _, _ in K.INCIDENTS))
    put_ref("lab_2_4.json", {"incidents": {i: {"path": p, "key": k} for i, _, p, k in K.INCIDENTS}})
    rows = [(d[0], K.ROLE[d[2]], "، ".join(K.ROLE[r] for r in d[4]), "، ".join(K.ROLE[r] for r in d[5])) for d in K.DECISIONS]
    inc = "\n".join(f"- {i}: {' ← '.join(p)}. {k}" for i, _, p, k in K.INCIDENTS)
    sol("D2_4_raci.md", "2.4 · المصفوفة الكاملة ومسارات التصعيد", "`files/raci_start.xlsx` و`files/incidents.txt` و`files/org_chart.txt`", f"""
{table(["الرقم", "A", "يُستشار (C)", "يُبلَّغ (I)"], rows)}

الفحص البنيوي (معيار النجاح): في كل صف حرف A واحد فقط، ولا صف بلا A. والمراجعة الداخلية في عمود I فقط.

## الحوادث

{inc}
""")


# ------------------------------------------------------------------ day 3
MANDATE = """
{auth} · الاختصاص (مقتطف من التنظيم)
{note}

تختص الهيئة بما يلي:
1. وضع اشتراطات السلامة من الحرائق للمنشآت والمباني والفعاليات، والتفتيش على الالتزام بها.
2. إصدار شهادات السلامة وتجديدها وإلغاؤها.
3. استقبال البلاغات الطارئة وتوجيه فرق الاستجابة.
4. توعية الجمهور بالسلامة.

قيود تعمل الهيئة ضمنها:
- تخضع بياناتها لسياسة تصنيف البيانات الوطنية: عام، مقيد، سري، سري للغاية. معظم بيانات البلاغات والتفتيش «مقيد» أو «سري».
- تعالج بيانات شخصية وحساسة (صحية، موقع) فتلتزم بنظام حماية البيانات الشخصية ولائحته التنفيذية.
- تلتزم بمبادئ أخلاقيات الذكاء الاصطناعي ومبادئ الذكاء الاصطناعي التوليدي للجهات الحكومية الصادرة عن سدايا، ومنها:
  ألا يُدخل الموظف في أدوات الذكاء الاصطناعي التوليدي إلا بيانات مصنفة «عام»، وأن يراجع الإنسان المخرجات، وأن يُبلَّغ المستفيد
  إذا كان المحتوى مولدًا، وأن تتاح له قناة غير آلية.
"""

AUP_TEMPLATE = """
# سياسة الاستخدام المقبول للذكاء الاصطناعي · {auth}

> قالب. لا تتجاوز السياسة أربع صفحات. اكتب كل بند بحيث يستطيع موظف أن يحسم به حالة دون أن يسأل أحدًا.

## 1. النطاق
من تسري عليه السياسة، وما الأنظمة التي تشملها (لا تنسَ الخصائص في البرمجيات المشتراة والأتمتة التي يبنيها الموظفون).

## 2. الاستخدامات المحظورة
فئات لا استثناء منها.

## 3. الاستخدامات المقيدة
ما يحتاج موافقة مسبقة: من يوافق، وبأي شرط. (أدوات الذكاء الاصطناعي التوليدي، ووكلاء البرمجة، والأتمتة، والوسائط الاصطناعية، والقرارات عن الأفراد.)

## 4. الاستخدامات المسموحة بشروط
ما يجوز دون موافقة، والشروط: تصنيف البيانات، والمراجعة، والوسم.

## 5. البيانات وموطنها
أي تصنيف يجوز إدخاله في أي خدمة، وأين تُعالج البيانات.

## 6. الاستثناءات
من يطلب، ومن يعتمد، ولأي مدة، وأين تُسجل.

## 7. المراجعة
تاريخ المراجعة القادمة ومحفّزاتها.
"""


def lab_3_1():
    write_text(OUT / "mandate.txt", MANDATE.format(auth=C.AUTHORITY, note=C.NOTE))
    write_text(OUT / "aup_template.md", AUP_TEMPLATE.format(auth=C.AUTHORITY))
    put_ref("lab_3_1.json", {"clauses": [{"n": n, "head": h} for n, h, _ in K.AUP]})
    body = "\n\n".join(f"**{n} {h}.** {t}" for n, h, t in K.AUP)
    words = sum(len(t.split()) for _, _, t in K.AUP)
    sol("D3_1_aup.md", "3.1 · مسودة السياسة", "`files/aup_template.md` و`files/mandate.txt`", f"""
سياسة مرجعية في {words} كلمة تقريبًا، أي أقل من صفحتين. ليست الصياغة الوحيدة الصحيحة؛ المهم أن يحسم كل بند حالة.

{body}

ما يبحث عنه المدرب في مسودة المشارك: الفئات المحظورة بلا استثناء؛ وقاعدة التصنيف للأدوات الخارجية (عام فقط)؛
ووكلاء البرمجة والأتمتة والخصائص المدمجة؛ ومن يعتمد الاستثناء ولكم من الوقت؛ وتاريخ مراجعة محدد.
""")


def lab_3_2():
    write_text(OUT / "staff_scenarios.txt", f"{C.AUTHORITY} · عشرة سيناريوهات للموظفين\n{C.NOTE}\n\n" +
               "لكل سيناريو: القرار (مسموح، مسموح بشروط، مقيد بموافقة، محظور) ورقم البند الذي حسمه. إن لم يحسمه بند، فهي فجوة.\n\n" +
               "\n".join(f"{s[0]}: {s[1]}" for s in K.SCENARIOS))
    keep = {"1", "2.1", "2.4", "3.4", "3.5", "4.1", "5", "6"}
    draft = [f"# سياسة الاستخدام المقبول للذكاء الاصطناعي · {C.AUTHORITY} (مسودة زميل، الإصدار 0.3)", "",
             f"> {C.NOTE}", ""]
    for n, h, t in K.AUP:
        if n in keep:
            draft += [f"**{n} {h}.** {t}", ""]
    draft += ["**ملاحظة الكاتب.** لم أكتب شيئًا بعد عن أدوات البرمجة ولا عن الأتمتة؛ أظنها تدخل في البند 4.1."]
    write_text(OUT / "aup_colleague_draft.md", "\n".join(draft))
    drafted = {n for n in keep}
    resolved = [s[0] for s in K.SCENARIOS if s[3] in drafted]
    assert resolved == K.DRAFT_RESOLVES, resolved
    genai_ok = [i for i in C.GENAI if i not in C.GENAI_BREACH]
    put_ref("lab_3_2.json", {s[0]: {"decision": s[2], "clause": s[3]} for s in K.SCENARIOS} |
            {"draft_resolves": resolved, "genai_breaches": C.GENAI_BREACH, "genai_ok": genai_ok})
    rows = [(s[0], s[2], s[3], "نعم" if s[0] in resolved else "لا: فجوة") for s in K.SCENARIOS]
    sol("D3_2_scenarios.md", "3.2 · عشرة سيناريوهات", "`files/staff_scenarios.txt` و`files/aup_colleague_draft.md` أو مسودتك", f"""
{table(["السيناريو", "القرار", "البند المرجعي", "تحسمه مسودة الزميل؟"], rows)}

مسودة الزميل تحسم 7 من 10: معيار النجاح ثمانية، فلا تكفي. الفجوات: SC03 (وكلاء البرمجة)، وSC05 (صوت شخص حقيقي)،
وSC07 (أتمتة تراسل السكان). إضافة البنود 3.1 و3.2 و3.3 تجعلها 10 من 10. ملاحظة الكاتب خاطئة: البند 4.1 للبيانات العامة،
والأتمتة التي تراسل السكان ليست كذلك.

SC08 فخ: الخاصية نفسها مقيدة (3.4)، لكن تشغيلها في اجتماعات «سري» محظور بنص البند نفسه.
SC04 مسموح بشروط: صياغة عامة دون إدخال اسم الموظف أو بياناته، والحكم يكتبه المدير (3.5).

## السجل تحت البند 2.4 (بعد المختبر، في شريحة «سياستك على سجلك»)

الأنظمة التوليدية في سجل الهيئة: {'، '.join(C.GENAI)}. ستة منها تعالج بيانات غير مصنفة «عام»، فتخالف اليوم قاعدة
مبادئ الذكاء الاصطناعي التوليدي للجهات الحكومية والبند 2.4، ولا يرفع المخالفة في الهيئة اعتماد ولا استثناء، لأن البند 5 من سياستها لا يجيز استثناءً من محظور:

{table(["النظام", "تصنيف البيانات", "ما يلزم"], [(f"{i} {C.SYS[i]['name']}", C.SYS[i]['data_class'], K.GENAI_ACTIONS[i]) for i in C.GENAI_BREACH])}

لا مخالفة في S14 وS22 (بيانات عامة) ولا في S12 (لوائح منشورة للعموم)؛ لكن الاستثناء المغلق EXC-07 سمح له يومًا بلائحة
داخلية لم تُنشر، وكان استثناءً من محظور. وS27 بحث دلالي لا يولّد نصًا: تحكمه موافقة البيانات وضوابط الأمن السيبراني، لا البند 2.4.
للمدرب: هذه القراءة الصارمة للقاعدة الوطنية تجعل وكيل البرمجة على شيفرة داخلية مصنفة «مقيد» مخالفة، مهما كانت ضوابط البند 3.2؛
ناقشها مع القاعة، فاللجنة الصورية قد تسأل عنها.
""")


def lab_3_3():
    wb = openpyxl.Workbook()
    note_sheet(wb, ["لكل بوابة ومستوى: معايير الدخول، والأدلة المطلوبة، والدور الذي يوافق. خذ الموافق من المصفوفة المرجعية (raci_reference.xlsx).",
                    "المخاطر غير المقبولة تقف عند بوابة الاستقبال."])
    rows = [(g, gn, LV[lv], "", "", "") for g, gn in K.GATES for lv in ["N", "L", "H"]]
    sheet(wb, "البوابات", ["البوابة", "الاسم", "المستوى", "معايير الدخول", "الأدلة المطلوبة", "الموافق"], rows,
          fill_cols=(3, 4, 5), widths=[8, 30, 22, 40, 50, 30], validations={"F": [r[1] for r in K.ROLES] + ["لا تلزم"]})
    save(wb, "gate_template.xlsx")
    ref_rows = []
    for g, gn in K.GATES:
        for lv in ["N", "L", "H"]:
            spec = K.GATE_SPEC.get((g, lv))
            ref_rows.append((g, gn, LV[lv]) + ((spec[0], spec[1], K.ROLE[spec[2]]) if spec else ("لا تلزم", "", "")))
    ref_rows.append(("G0", "الاستقبال والفرز", LV["U"], K.GATE_SPEC[("G0", "U")][0], K.UNACCEPTABLE_RULE, K.ROLE["R03"]))
    wb = openpyxl.Workbook()
    note_sheet(wb, ["مواصفات البوابات المرجعية: نقطة البداية للمختبر 3.4 ولحزمة المشروع الختامي."])
    sheet(wb, "البوابات", ["البوابة", "الاسم", "المستوى", "معايير الدخول", "الأدلة المطلوبة", "الموافق"], ref_rows,
          widths=[8, 30, 22, 40, 50, 30])
    save(wb, "gate_spec_reference.xlsx")
    put_ref("lab_3_3.json", {f"{g}|{lv}": {"approver": v[2], "decision": v[3]} for (g, lv), v in K.GATE_SPEC.items()})
    sol("D3_3_gates.md", "3.3 · مواصفات البوابات", "`files/gate_template.xlsx` و`files/raci_reference.xlsx`", f"""
{table(["البوابة", "الاسم", "المستوى", "معايير الدخول", "الأدلة", "الموافق"], ref_rows)}

القاعدة: الموافق في كل خلية هو المساءل (A) عن القرار نفسه في المصفوفة: D01 للاستقبال، وD03 للبيانات، وD04 وD05 وD08 للعالية، وD06 للمحدودة، وD11 لإيقاف المحدودة.
التناسب: البسيطة تمر بالاستقبال والتسجيل والتأكيد السنوي فقط؛ والعالية تمر بالبوابات الخمس كلها وبموافقة اللجنة.
تخطيط الإيقاف بالتفصيل في المقرر LDS-AIG-203، وتقييم الأثر الكامل في LDS-AIG-303.
""")


def lab_3_4():
    for sid, d in K.DOSSIERS.items():
        s = C.SYS[sid]
        lines = [f"{C.AUTHORITY} · ملف النظام {sid}: {s['name']}", C.NOTE, "",
                 f"المستوى في السجل: {LV[s['level']]}. المرحلة: {C.STAGES[s['stage']]}. المالك: {s['owner']}.",
                 f"البوابة المطلوبة: {d['gate']}."]
        if d["change"]:
            lines.append(f"التغيير المطلوب اعتماده: {d['change']}.")
        lines.append("الأدلة المرفقة بالطلب:")
        lines += [f"- {x}" for x in d["have"]]
        lines.append("ملاحظة المالك: «الملف مكتمل ونرجو العرض على أقرب جهة توافق».")
        write_text(OUT / f"dossier_{sid}.txt", "\n".join(lines))
    wb = openpyxl.Workbook()
    note_sheet(wb, [f"سجل البوابات والاستثناءات حتى {C.AS_OF}. احسب المقاييس الثلاثة."])
    sheet(wb, "السجل", ["رقم النظام", "اسم النظام", "مدخل مكتمل (مالك مؤكد ومستوى مسجل)"],
          [(s["id"], s["name"], "لا" if s["id"] in K.INCOMPLETE else "نعم") for s in C.S], widths=[10, 36, 22])
    sheet(wb, "البوابات", ["رقم النظام", "البوابة", "تاريخ التقديم", "تاريخ القرار"],
          [(e[0], e[1], e[2], e[3] or "") for e in K.GATE_EVENTS], widths=[10, 10, 14, 14])
    sheet(wb, "الاستثناءات", ["الرقم", "النظام", "الاستثناء", "تاريخ الاعتماد", "ينتهي في", "الحالة"],
          K.EXCEPTIONS, widths=[10, 8, 50, 14, 14, 10])
    save(wb, "gate_log.xlsx")
    from datetime import date
    d = lambda x: date.fromisoformat(x)
    complete = len(C.S) - len(K.INCOMPLETE)
    cycles = sorted((d(e[3]) - d(e[2])).days for e in K.GATE_EVENTS if e[3])
    n = len(cycles)
    median = cycles[n // 2] if n % 2 else (cycles[n // 2 - 1] + cycles[n // 2]) / 2
    open_exc = [e for e in K.EXCEPTIONS if e[5] == "مفتوح"]
    overdue = [e[0] for e in open_exc if d(e[4]) < d(C.AS_OF)]
    pending = [(e[0], e[1], (d(C.AS_OF) - d(e[2])).days) for e in K.GATE_EVENTS if not e[3]]
    m = {"coverage": [complete, len(C.S)], "cycle_days": cycles, "median_days": median, "decided": n,
         "open_exceptions": len(open_exc), "overdue_exceptions": overdue, "pending": pending}
    put_ref("lab_3_4.json", {"dossiers": K.DOSSIERS, "measures": m})
    pend = "، ".join(f"{p[0]} عند {p[1]} منذ {p[2]} يومًا" for p in pending)
    sol("D3_4_two_systems.md", "3.4 · نظامان عبر البوابات", "`files/gate_spec_reference.xlsx` و`files/dossier_S06.txt` و`files/dossier_S07.txt` و`files/gate_log.xlsx`", f"""
## النظامان

- S06 أولوية التفتيش (عالية، تغيير جوهري): البوابة G3. القرار **تعليق**: ينقص الملف موافقة مكتب إدارة البيانات على المصدر الجديد
  (القرار D03). الموافق على إعادة الموافقة لجنة حوكمة الذكاء الاصطناعي (D08)، لا مكتب الحوكمة ولا «أقرب جهة». مصدر بيانات جديد
  عن المنشآت يُعامل تغييرًا جوهريًا، فلا يكفي التأكيد السنوي.
- S07 مساعد مراجعة المخططات (محدودة، تصميم): البوابة G1. القرار **عبور**: مذكرة التصميم وموافقة البيانات موجودتان. الموافق
  الرئيسة التنفيذية للبيانات (D03). البوابة التالية G2 تحتاج تقرير الاختبار وإشعار المستخدمين، ويوافق مكتب الحوكمة (D06).

## المقاييس حتى {C.AS_OF}

- تغطية السجل: {complete} من {len(C.S)} مدخلًا مكتملًا ({round(100 * complete / len(C.S))}%). الناقص: الخمسة غير المصرح عنها، وS28.
- زمن دورة البوابة: الوسيط {median} يومًا على {n} قرارًا (من {cycles[0]} إلى {cycles[-1]} يومًا). المتوسط يضلل هنا لأن S29 وS10 أخذا {cycles[-2]} و{cycles[-1]} يومًا.
- الاستثناءات المفتوحة: {len(open_exc)}، انتهت مدة {len(overdue)} منها ولم تُغلق ({"، ".join(overdue)}): علامة حوكمة شكلية.
- طلبات بلا قرار: {pend}.
""")


def answer_key():
    counts = {lv: sum(1 for s in C.S if s["level"] == lv) for lv in C.LEVELS}
    write_text(SOL / "answer_key.md", f"""
# مفتاح الإجابة للمدرب · {C.AUTHORITY}

كل الملفات تولّدها `labs/_tools/prepare_files.py` من `case.py` و`case_docs.py`، ويتحقق `labs/solutions/check.py` من كل مفتاح.

## المشكلات المزروعة، وأين تُكتشف

| المشكلة | الملف | المختبر |
|---|---|---|
| أربعة ضوابط بلا مالك أو محفّز أو اختبار (G13 إلى G16) | governance_statements.xlsx | 1.1 |
| شعارات وإطار مرجعي في كتالوج الضوابط (CC03، CC07، CC10، CC11) | control_catalogue.xlsx | 1.2 |
| أربعة بنود ليست ذكاءً اصطناعيًا، ونظامان مذكوران مرتين | interview_notes.txt | 1.3 |
| خمسة أنظمة غير مصرح عنها: S11 وS30 في المشتريات، وS02 وS21 في الشبكة، وS15 في المصروفات | procurement_records.xlsx، network_summary.xlsx، expense_claims.xlsx | 1.4 |
| الموارد البشرية تنفي استخدام الذكاء الاصطناعي في التوظيف | interview_notes.txt مقابل PO-26-0151 | 1.4 |
| مقترحان محظوران (S13، S17) | inventory_day2.xlsx | 2.1 |
| ثلاث حالات حدّية (S02، S06، S10) وتغييران في الوقائع | borderline_cases.txt | 2.2 |
| المراجعة الداخلية تطلب صلاحية موافقة | interview_notes.txt، org_chart.txt | 2.3 |
| حادث على نظام عالي المخاطر يحتاج إعادة موافقة | incidents.txt | 2.4 |
| مسودة زميل تسكت عن وكلاء البرمجة والأتمتة والوسائط الاصطناعية | aup_colleague_draft.md | 3.2 |
| ملف S06 ينقصه موافقة البيانات، ويطلب «أقرب جهة توافق» | dossier_S06.txt | 3.4 |
| استثناءان مفتوحان انتهت مدتهما | gate_log.xlsx | 3.4 |
| ستة أنظمة توليدية تعالج بيانات غير «عام» (S02، S09، S15، S19، S21، S30)، والاستثناء EXC-01 من محظور | inventory_day2.xlsx، gate_log.xlsx | بعد 3.2 |

## الأرقام

توزيع المستويات: غير مقبولة {counts['U']}، عالية {counts['H']}، محدودة {counts['L']}، بسيطة أو منعدمة {counts['N']}.
الحلول المرجعية لكل مختبر في `labs/solutions/D*_*.md`.

## السلم وقاعدة القرار

السلم ذو العوامل الخمسة وقاعدة القرار من وضع الهيئة الافتراضية، وليسا من نص سدايا. المستويات الأربعة وأمثلتها وطنية،
ويتحقق المولّد من أن مرشح البريد العشوائي (S18) بسيط المخاطر وأن التصنيف الاجتماعي (S17) غير مقبول كما في نص المبادئ.
""")


def main():
    for p in list(OUT.rglob("*")) if OUT.exists() else []:
        if p.is_file():
            p.unlink()
    case_brief()
    lab_1_1(); lab_1_2(); lab_1_3(); lab_1_4()
    lab_2_1(); lab_2_2(); lab_2_3(); lab_2_4()
    lab_3_1(); lab_3_2(); lab_3_3(); lab_3_4()
    answer_key()
    print("files:", len([p for p in OUT.rglob("*") if p.is_file()]), "reference:", len(list(REF.glob("*.json"))))


if __name__ == "__main__":
    main()
