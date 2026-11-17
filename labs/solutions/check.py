# /// script
# requires-python = ">=3.10"
# dependencies = ["openpyxl>=3.1"]
# ///
"""Run the reference solution of every lab against the course files and check it.

    uv run labs/solutions/check.py            # all labs
    uv run labs/solutions/check.py 2.1 3.4    # some labs

Every lab of this course is done in a spreadsheet, a document or a diagram; none calls a model.
This script does on the files what the reference solution does (reads the records, applies the
tiering rule, checks the matrix, resolves the scenarios, computes the measures) and checks the
keys in solutions/reference/ against the files, so a key the files do not support fails here.
It prints the numbers the debrief slides quote and PASS per lab. Nothing is written to disk.
"""
from __future__ import annotations

import json
import re
import statistics
import sys
from collections import Counter
from datetime import date
from pathlib import Path

import openpyxl

HERE = Path(__file__).resolve().parent
FILES = HERE.parent / "files"
REF = HERE / "reference"
LABS: dict = {}
LEVEL_AR = {"N": "مخاطر بسيطة أو منعدمة", "L": "مخاطر محدودة", "H": "مخاطر عالية", "U": "مخاطر غير مقبولة"}


def lab(lab_id):
    def wrap(fn):
        LABS[lab_id] = fn
        return fn
    return wrap


def ref(name):
    return json.loads((REF / name).read_text(encoding="utf-8"))


def text(rel):
    return (FILES / rel).read_text(encoding="utf-8")


def rows(rel, sheet):
    ws = openpyxl.load_workbook(FILES / rel, read_only=True)[sheet]
    it = ws.iter_rows(values_only=True)
    head = next(it)
    return [dict(zip(head, r)) for r in it if any(v not in (None, "") for v in r)]


def rule(f, prohibited):
    """The decision rule as written in files/tiering_rubric.txt."""
    if prohibited:
        return "U"
    total = sum(f.values())
    if f["I"] == 3 and (f["A"] >= 2 or f["R"] == 3):
        return "H"
    if total >= 12:
        return "H"
    if total >= 9 or f["I"] == 3 or f["D"] == 3:
        return "L"
    return "N"


@lab("1.1")
def l11():
    key = ref("lab_1_1.json")
    got = rows("governance_statements.xlsx", "العبارات")
    assert [r["الرقم"] for r in got] == [k["id"] for k in key]
    layers = Counter(k["layer"] for k in key)
    non_op = [k["id"] for k in key if k["layer"] == "ضابط" and k["missing"]]
    assert non_op == ["G13", "G14", "G15", "G16"], non_op
    for r, k in zip(got, key):  # a statement keyed as a control reads like one: no step list, no signature record
        if k["layer"] == "إجراء":
            assert ":" in r["العبارة"], r["الرقم"]
    print("  layers:", dict(layers), "| not operable:", non_op)


@lab("1.2")
def l12():
    key = ref("lab_1_2.json")
    cat = {r["الرقم"]: r for r in rows("control_catalogue.xlsx", "الضوابط")}
    oecd = {r["الرمز"] for r in rows("control_catalogue.xlsx", "مبادئ OECD")}
    nist = {r["الوظيفة"] for r in rows("control_catalogue.xlsx", "وظائف NIST")}
    operable = [c for c, r in cat.items() if r["المالك"] and r["المحفّز"] and r["الاختبار"]]
    assert operable == key["operable"], operable
    traced = 0
    for p, m in key["map"].items():
        c = cat[m["controls"][0]]
        assert c["المالك"] and c["المحفّز"] and c["الاختبار"], p
        assert set(m["oecd"]) <= oecd and set(m["nist"]) <= nist, p
        traced += 1
    principles = [r["المبدأ الوطني"] for r in rows("principle_control_map.xlsx", "الخريطة")]
    assert len(principles) == 7
    print(f"  principles traced in the reference: {traced} of 7 (bar: 5) | operable controls: {len(operable)} of {len(cat)}")


@lab("1.3")
def l13():
    key = ref("lab_1_3.json")
    notes = text("interview_notes.txt")
    items = re.findall(r"^- «([^»]+)»", notes, flags=re.M)
    inv = {r["رقم النظام"]: r for r in rows("inventory_day2.xlsx", "السجل")}
    declared = [inv[s]["اسم النظام"] for s in key["declared"]]
    for n in declared:
        assert items.count(n) == 1, n
    for n in key["not_ai"]:
        assert n in items, n
    for alias in key["duplicates"]:
        assert alias in items, alias
    undeclared = [r["اسم النظام"] for r in inv.values() if r["مصرّح عنه"] == "لا"]
    for n in undeclared:
        assert n not in notes, n
    assert "لا نستخدم الذكاء الاصطناعي في التوظيف" in notes
    print(f"  items in the notes: {len(items)} = {len(declared)} systems + {len(key['not_ai'])} not AI + {len(key['duplicates'])} duplicates")
    assert len(items) == len(declared) + len(key["not_ai"]) + len(key["duplicates"])


@lab("1.4")
def l14():
    key = ref("lab_1_4.json")
    start = rows("inventory_start.xlsx", "السجل")
    proc = {r["أمر الشراء"]: r for r in rows("procurement_records.xlsx", "المشتريات")}
    exp = {r["رقم المطالبة"]: r for r in rows("expense_claims.xlsx", "المصروفات")}
    net = {r["النطاق"]: r for r in rows("network_summary.xlsx", "الشبكة")}
    found = Counter()
    for sid, u in key["undeclared"].items():
        table = {"procurement": proc, "expense": exp, "network": net}[u["source"]]
        rec = table[u["record"]]
        blob = " ".join(str(v) for v in rec.values())
        assert "ذكاء" in blob or "ذكي" in blob or "تحويل الكلام" in blob or "توليدي" in blob, sid
        assert sid not in [r["رقم النظام"] for r in start], sid
        found[u["source"]] += 1
    assert len(start) + len(key["undeclared"]) == key["total"] == 30
    months = sorted(r["الشهر"] for r in exp.values() if "تطبيق كتابة" in r["البند"])
    print(f"  start: {len(start)} declared; found: {dict(found)}; total {key['total']}; writing app claimed in {months}")


@lab("2.1")
def l21():
    key = ref("lab_2_1.json")
    inv = {r["رقم النظام"]: r for r in rows("inventory_day2.xlsx", "السجل")}
    rubric = text("tiering_rubric.txt")
    assert "12 أو أكثر" in rubric and "9 أو أكثر" in rubric
    for sid, k in key.items():
        r = inv[sid]
        n = r["المتأثرون سنويًا"]
        band = 1 if n < 5000 else 2 if n <= 100000 else 3
        assert band == k["f"]["S"], (sid, n)
        auto = {"مخرجاته معلومات": 1, "يوصي بقرار": 2, "يتصرف أو يقرر": 3}
        a = next(v for t, v in auto.items() if r["دور النظام في القرار"].startswith(t))
        assert a == k["f"]["A"], sid
        assert rule(k["f"], k["prohibited"]) == k["level"], sid
        if r["تصنيف البيانات"] == "سري":
            assert k["f"]["D"] == 3, sid
    assert key["S18"]["level"] == "N" and key["S17"]["level"] == "U"
    dist = Counter(k["level"] for k in key.values())
    sheet = rows("inventory_day2.xlsx", "التصنيف")
    assert len(sheet) == 30 and all(not r["المستوى"] for r in sheet)
    print("  levels:", {LEVEL_AR[x]: dist[x] for x in "UHLN"}, "| pass bar: 24 of 30")


@lab("2.2")
def l22():
    key = ref("lab_2_2.json")
    base = ref("lab_2_1.json")
    cases = text("borderline_cases.txt")
    for sid, lv in key["borderline"].items():
        assert f"الحالة {sid}" in cases and base[sid]["level"] == lv
    for c in key["changed"]:
        f = dict(base[c["id"]]["f"]); f[c["factor"]] = c["value"]
        assert rule(f, False) == c["after"] and base[c["id"]]["level"] == c["before"], c
        assert c["before"] != c["after"]
    print("  borderline:", {k: LEVEL_AR[v] for k, v in key["borderline"].items()},
          "| changed:", [(c["id"], LEVEL_AR[c["before"]], LEVEL_AR[c["after"]]) for c in key["changed"]])


def raci(rel):
    return rows(rel, "المصفوفة")


@lab("2.3")
def l23():
    key = ref("lab_2_3.json")
    start = raci("raci_start.xlsx")
    roles = [h for h in start[0].keys() if h not in ("الرقم", "القرار")]
    assert len(start) == 12 and len(roles) == 12
    for r in start:
        a = [h for h in roles if r[h] == "A"]
        assert len(a) == 1, r["الرقم"]
    org = text("org_chart.txt")
    assert "مالك النظام" in org and "يوقف نظامه فورًا" in org and "تفعيل خصائصه في البرمجيات المشتراة" in org
    by_a = Counter(k["A"] for k in key.values())
    print("  accountable roles used:", dict(by_a))


@lab("2.4")
def l24():
    full = raci("raci_reference.xlsx")
    roles = [h for h in full[0].keys() if h not in ("الرقم", "القرار")]
    audit = next(h for h in roles if "المراجعة الداخلية" in h)
    for r in full:
        letters = [r[h] for h in roles if r[h]]
        assert letters.count("A") == 1, r["الرقم"]
        assert r[audit] in (None, "", "I"), r["الرقم"]
    inc = text("incidents.txt")
    assert inc.count("INC-") == 3
    print(f"  12 decisions, each with exactly one A; internal audit only informed; {sum(1 for r in full for h in roles if r[h])} cells filled")


@lab("3.1")
def l31():
    key = ref("lab_3_1.json")
    md = (HERE / "D3_1_aup.md").read_text(encoding="utf-8")
    for c in key["clauses"]:
        assert f"**{c['n']} {c['head']}.**" in md, c
    body = md.split("أي أقل من صفحتين")[1].split("ما يبحث عنه المدرب")[0]
    words = len(re.sub(r"[*][*][^*]+[*][*]", "", body).split())
    assert words < 1600, words  # four pages of about 400 words
    print(f"  {len(key['clauses'])} clauses; about {words} words in the reference page")


@lab("3.2")
def l32():
    key = ref("lab_3_2.json")
    md = (HERE / "D3_1_aup.md").read_text(encoding="utf-8")
    draft = text("aup_colleague_draft.md")
    sc = text("staff_scenarios.txt")
    resolved_ref = resolved_draft = 0
    for sid, k in key.items():
        if not sid.startswith("SC"):
            continue
        assert f"{sid}:" in sc
        assert f"**{k['clause']} " in md, sid
        resolved_ref += 1
        if f"**{k['clause']} " in draft:
            resolved_draft += 1
    assert resolved_ref == 10 and resolved_draft == len(key["draft_resolves"]) == 7
    inv = {r["رقم النظام"]: r for r in rows("inventory_day2.xlsx", "السجل")}
    for sid in key["genai_breaches"]:
        assert inv[sid]["تصنيف البيانات"] != "عام", sid
        assert f"| {sid} " in (HERE / "D3_2_scenarios.md").read_text(encoding="utf-8"), sid
    for sid in key["genai_ok"]:
        assert inv[sid]["تصنيف البيانات"] == "عام", sid
    print(f"  register against clause 2.4: generative systems with non-public data {key['genai_breaches']}")
    print(f"  reference policy resolves {resolved_ref} of 10; colleague draft {resolved_draft} of 10 (bar: 8)")


@lab("3.3")
def l33():
    key = ref("lab_3_3.json")
    dec = ref("lab_2_3.json")
    full = raci("raci_reference.xlsx")
    roles = [h for h in full[0].keys() if h not in ("الرقم", "القرار")]
    a_of = {r["الرقم"]: next(h for h in roles if r[h] == "A") for r in full}
    spec = rows("gate_spec_reference.xlsx", "البوابات")
    names = {"R01": "رئيس الهيئة", "R02": "لجنة حوكمة", "R03": "مكتب حوكمة", "R04": "مالك النظام", "R05": "الرئيسة التنفيذية للبيانات"}
    linked = 0
    for gl, v in key.items():
        if v["decision"]:
            assert dec[v["decision"]]["A"] == v["approver"], gl
            assert a_of[v["decision"]].startswith(names[v["approver"]]), gl
            linked += 1
    assert len(spec) == 16
    print(f"  {len(key)} gate cells with an approver; {linked} tied to a RACI decision, all equal to its A")


@lab("3.4")
def l34():
    key = ref("lab_3_4.json")
    m = key["measures"]
    reg = rows("gate_log.xlsx", "السجل")
    ev = rows("gate_log.xlsx", "البوابات")
    exc = rows("gate_log.xlsx", "الاستثناءات")
    as_of = date(2026, 10, 31)
    complete = sum(1 for r in reg if r["مدخل مكتمل (مالك مؤكد ومستوى مسجل)"] == "نعم")
    cycles = sorted((date.fromisoformat(r["تاريخ القرار"]) - date.fromisoformat(r["تاريخ التقديم"])).days for r in ev if r["تاريخ القرار"])
    med = statistics.median(cycles)
    open_ = [r for r in exc if r["الحالة"] == "مفتوح"]
    overdue = [r["الرقم"] for r in open_ if date.fromisoformat(r["ينتهي في"]) < as_of]
    assert [complete, len(reg)] == m["coverage"] and cycles == m["cycle_days"] and med == m["median_days"]
    assert len(open_) == m["open_exceptions"] and overdue == m["overdue_exceptions"]
    d6 = text("dossier_S06.txt")
    assert "موافقة مكتب إدارة البيانات" not in d6 and "G3" in d6
    d7 = text("dossier_S07.txt")
    assert "موافقة مكتب إدارة البيانات على المصادر" in d7 and "G1" in d7
    print(f"  coverage {complete}/{len(reg)} = {round(100 * complete / len(reg))}%; median gate cycle {med} days over {len(cycles)} decisions "
          f"(mean {round(statistics.mean(cycles), 1)}); open exceptions {len(open_)}, past expiry {overdue}")


def main(argv):
    sys.stdout.reconfigure(encoding="utf-8")
    wanted = argv or list(LABS)
    failed = 0
    for lab_id in wanted:
        print(f"Lab {lab_id}")
        try:
            LABS[lab_id]()
            print("  PASS")
        except AssertionError as e:
            failed += 1
            print(f"  FAIL {e!r}")
    print(f"{len(wanted) - failed} of {len(wanted)} labs PASS")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
