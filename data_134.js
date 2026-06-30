#!/usr/bin/env python3
"""
سكريبت تحديث لوحة مشروع الإرشاد السياحي + حساب BVA النسبي
الاستخدام: python3 update_tourism.py --excel ملف_الاكسيل.xlsx
"""
import openpyxl, json, sys, os
from datetime import datetime
from collections import defaultdict

# ═══ الميزانيات الثابتة (العقد — 300 دورة) ═══
TOTAL_COURSES_PLAN = 300
BVA_BUDGETS = [
    ('trainer_fee',     'أجر المدرب',      1800000),
    ('coordinator_fee', 'أجر المنسق',       315000),
    ('bus_fee',         'الباص',             810000),
    ('hotel_snack',     'الفندق/السناك',   2925000),
    ('printing',        'الطباعة',            36000),
    ('banners',         'البنرات',             5100),
    ('shipping',        'الشحن',              12000),
    ('sms',             'SMS',                36000),
    ('cert_fee',        'الشهادة الدولية', 1530000),
    ('staff',           'الموظفين',          585000),
]

def n(v, d=0):
    try: return float(v) if v is not None else d
    except: return d

def main(excel_path):
    print(f'📂 قراءة: {excel_path}')
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb[wb.sheetnames[0]]

    ar_m = ["يناير","فبراير","مارس","أبريل","مايو","يونيو",
            "يوليو","أغسطس","سبتمبر","أكتوبر","نوفمبر","ديسمبر"]

    # ═══ استخراج الدورات ═══
    courses = []
    for row in range(8, 400):
        num_raw = ws.cell(row=row, column=1).value
        proj = ws.cell(row=row, column=5).value
        if not num_raw or not proj or str(proj).strip() != 'الوزارة': continue
        num = int(n(num_raw))
        if num > 135: continue
        reg = int(n(ws.cell(row=row, column=10).value))
        if reg == 0: continue

        bus_dt = ws.cell(row=row, column=4).value
        dt_raw = ws.cell(row=row, column=3).value
        bd = None
        if isinstance(bus_dt, datetime): bd = bus_dt
        elif isinstance(dt_raw, datetime): bd = dt_raw
        elif isinstance(dt_raw, str):
            try: bd = datetime.strptime(dt_raw, '%d-%m-%Y')
            except: pass
        if not bd: continue
        if bd.year == 2025 and 4 <= bd.month <= 11:
            bd = bd.replace(year=2026)

        courses.append({
            'num': num,
            'code': str(ws.cell(row=row, column=2).value or '').strip(),
            'bus_date': bd.strftime('%Y-%m-%d'),
            'city': str(ws.cell(row=row, column=6).value or '').strip(),
            'hotel': str(ws.cell(row=row, column=7).value or '').strip(),
            'trainer': str(ws.cell(row=row, column=8).value or '').strip(),
            'coordinator': str(ws.cell(row=row, column=9).value or '').strip(),
            'registered': reg,
            'withdrawn_before': int(n(ws.cell(row=row, column=11).value)),
            'certified': int(n(ws.cell(row=row, column=12).value)),
            'passed': int(n(ws.cell(row=row, column=13).value)),
            'failed': int(n(ws.cell(row=row, column=14).value)),
            'withdrawn_during': int(n(ws.cell(row=row, column=15).value)),
            'trainer_fee': n(ws.cell(row=row, column=16).value),
            'coordinator_fee': n(ws.cell(row=row, column=17).value),
            'bus_fee': n(ws.cell(row=row, column=18).value),
            'hotel_snack': n(ws.cell(row=row, column=19).value),
            'printing': n(ws.cell(row=row, column=20).value),
            'banners': round(n(ws.cell(row=row, column=21).value), 2),
            'shipping': n(ws.cell(row=row, column=22).value),
            'sms': n(ws.cell(row=row, column=23).value),
            'cert_fee': n(ws.cell(row=row, column=24).value),
            'staff': n(ws.cell(row=row, column=25).value),
            'total_expenses': round(n(ws.cell(row=row, column=26).value), 2),
            'revenue': round(n(ws.cell(row=row, column=27).value), 2),
            'profit': round(n(ws.cell(row=row, column=28).value), 2),
            'profit_pct': round(n(ws.cell(row=row, column=29).value), 2),
        })

    courses.sort(key=lambda x: x['num'])
    NC = len(courses)
    ratio = NC / TOTAL_COURSES_PLAN

    # ═══ التجميعات ═══
    cd = defaultdict(lambda: {'courses':0,'registered':0,'certified':0,'passed':0,'expenses':0,'revenue':0,'profit':0})
    for c in courses:
        d=cd[c['city']]; d['courses']+=1; d['registered']+=c['registered']; d['certified']+=c['certified']
        d['passed']+=c['passed']; d['expenses']+=c['total_expenses']; d['revenue']+=c['revenue']; d['profit']+=c['profit']
    cities = sorted([{'city':k,**{kk:round(vv) for kk,vv in v.items()}} for k,v in cd.items()], key=lambda x:-x['profit'])

    td = defaultdict(lambda: {'courses':0,'passed':0,'certified':0})
    for c in courses:
        t=td[c['trainer']]; t['courses']+=1; t['passed']+=c['passed']; t['certified']+=c['certified']
    trainers = sorted([{'name':k,**v,'rate':round(100*v['passed']/v['certified']) if v['certified'] else 0} for k,v in td.items()], key=lambda x:-x['courses'])

    mo = {}
    for c in courses:
        d=datetime.strptime(c['bus_date'],'%Y-%m-%d')
        lbl=f'{ar_m[d.month-1]} {d.year}'; sk=f'{d.year}-{d.month:02d}'
        if lbl not in mo: mo[lbl]={'courses':0,'revenue':0,'expenses':0,'profit':0,'_s':sk}
        m=mo[lbl]; m['courses']+=1; m['revenue']+=c['revenue']; m['expenses']+=c['total_expenses']; m['profit']+=c['profit']
    months = sorted([{'label':k,'courses':v['courses'],'revenue':round(v['revenue']),'expenses':round(v['expenses']),'profit':round(v['profit'])} for k,v in mo.items()], key=lambda x:mo[x['label']]['_s'])

    # ═══ BVA النسبي (Proportional Budget Variance) ═══
    bva = []
    for key, label, budget in BVA_BUDGETS:
        actual = sum(c[key] for c in courses)
        prorated = budget * ratio
        diff = prorated - actual
        pct = (diff / prorated * 100) if prorated else 0
        status = 'pos' if diff > 10 else ('neg' if diff < -10 else 'zero')
        bva.append({
            'key': key, 'label': label,
            'budget_total': budget,
            'budget_prorated': round(prorated, 2),
            'actual': round(actual, 2),
            'diff': round(diff, 2),
            'pct': round(pct, 2),
            'status': status,
        })

    # ═══ كتابة data.js ═══
    js = f'const COURSES = {json.dumps(courses, ensure_ascii=False)};\n'
    js += f'const CITIES = {json.dumps(cities, ensure_ascii=False)};\n'
    js += f'const TRAINERS = {json.dumps(trainers, ensure_ascii=False)};\n'
    js += f'const MONTHS_DATA = {json.dumps(months, ensure_ascii=False)};\n'
    js += f'const BVA = {json.dumps(bva, ensure_ascii=False)};\n'
    js += f'const BVA_META = {json.dumps({"courses_done":NC,"courses_plan":TOTAL_COURSES_PLAN,"ratio":round(ratio,4),"ratio_pct":round(ratio*100,2)}, ensure_ascii=False)};\n'
    with open('data.js','w',encoding='utf-8') as f:
        f.write(js)

    # ═══ المراجعة ═══
    T = lambda f: sum(c[f] for c in courses)
    trev=T('revenue'); texp=T('total_expenses'); tprof=T('profit')
    ex_rev=round(n(ws.cell(row=5, column=27).value))
    ex_exp=round(n(ws.cell(row=5, column=26).value))

    print(f'\n{"═"*60}')
    print(f'  {NC} دورة — {ratio*100:.2f}% من {TOTAL_COURSES_PLAN}')
    print(f'{"═"*60}')
    print(f'  إيراد: {trev:>12,.0f} {"✅" if abs(trev-ex_rev)<1 else "❌"}')
    print(f'  مصاريف: {texp:>11,.0f} {"✅" if abs(texp-ex_exp)<10 else "❌"}')
    print(f'  ربح: {tprof:>13,.0f}')
    print(f'  هامش: {tprof/trev*100:.1f}%')
    print(f'  مسجل: {T("registered"):,} | معتمد: {T("certified"):,} | مجتاز: {T("passed"):,}')

    # ═══ تقرير BVA النسبي ═══
    print(f'\n{"═"*60}')
    print(f'  تحليل BVA النسبي (المخصص لـ{NC} دورة = {ratio*100:.2f}% من الميزانية)')
    print(f'{"═"*60}')
    print(f'  {"البند":<22} {"المخصص":>12} {"الفعلي":>12} {"الفرق":>12}  الحالة')
    print(f'  {"-"*22} {"-"*12} {"-"*12} {"-"*12}  -----')
    for b in bva:
        emoji = '✅' if b['status']=='pos' else ('❌' if b['status']=='neg' else '⚖️')
        sign = '+' if b['diff']>=0 else ''
        print(f'  {b["label"]:<22} {b["budget_prorated"]:>12,.0f} {b["actual"]:>12,.0f} {sign+f"{b["diff"]:,.0f}":>12}  {emoji}')

    tot_p = sum(b['budget_prorated'] for b in bva)
    tot_a = sum(b['actual'] for b in bva)
    tot_d = tot_p - tot_a
    tot_pct = (tot_d/tot_p*100) if tot_p else 0
    print(f'  {"-"*22} {"-"*12} {"-"*12} {"-"*12}')
    sign = '+' if tot_d>=0 else ''
    emoji = '✅' if tot_d>0 else '❌'
    print(f'  {"الإجمالي":<22} {tot_p:>12,.0f} {tot_a:>12,.0f} {sign+f"{tot_d:,.0f}":>12}  {emoji} ({tot_pct:+.2f}%)')

    # ═══ توليد ملف الـ spans الجاهز للصق ═══
    with open('bva_spans.txt','w',encoding='utf-8') as f:
        f.write(f'<!-- BVA spans — جاهزة للصق بعد رقم الفعلي في كل صف -->\n')
        f.write(f'<!-- محسوبة عند {NC} دورة ({ratio*100:.2f}%) -->\n\n')
        for b in bva:
            sign = '+' if b['diff']>=0 else '−'
            val = f'{abs(b["diff"]):,.0f}'
            f.write(f'<!-- {b["label"]:<22} -->\n')
            f.write(f'<span class="bva-diff {b["status"]}">{sign}{val}</span>\n\n')

    print(f'\n  ✅ data.js جاهز')
    print(f'  ✅ bva_spans.txt جاهز (انسخ والصق في HTML)')
    print(f'\n  ⚠️ لا تنسَ تحديث HTML إذا تغيّر عدد الدورات أو قيم الـ spans!')

if __name__ == '__main__':
    if len(sys.argv) < 3 or sys.argv[1] != '--excel':
        print('الاستخدام: python3 update_tourism.py --excel ملف.xlsx')
        sys.exit(1)
    main(sys.argv[2])
