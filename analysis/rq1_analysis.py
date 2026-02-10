#!/usr/bin/env python3
"""RQ1 analysis for CPA pursuit factors.

Note: In this execution environment, pandas/openpyxl/matplotlib are unavailable.
This script uses pure Python stdlib to keep the workflow reproducible.
"""
from __future__ import annotations
import csv, json, math, os, re, statistics, zipfile, zlib
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict

ROOT = os.path.dirname(os.path.dirname(__file__))
DATA_CANDIDATES = [
    os.path.join(ROOT, "data"),
    ROOT,
]

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
LIKERT5 = {
    "Strongly disagree": 1,
    "Somewhat disagree": 2,
    "Neither agree nor disagree": 3,
    "Somewhat agree": 4,
    "Strongly agree": 5,
}
LIKELIHOOD5 = {
    "Very unlikely": 1,
    "Somewhat unlikely": 2,
    "Neither likely nor unlikely": 3,
    "Somewhat likely": 4,
    "Very likely": 5,
}
IMPORTANCE5 = {
    "Not at all important": 1,
    "Slightly important": 2,
    "Moderately important": 3,
    "Very important": 4,
    "Extremely important": 5,
}
YESNO = {"No": 0, "Yes": 1}
Q55 = {"Definitely not": 1, "Probably not": 2, "Might or might not": 3, "Probably yes": 4, "Definitely yes": 5}
Q6 = {"Very Negative": 1, "Somewhat Negative": 2, "Neutral": 3, "Somewhat Positive": 4, "Very Positive": 5}
Q30 = {
    "It had no influence on my decision to pursue a graduate program.": 1,
    "It was a minor factor in my decision to pursue a graduate program.": 2,
    "It was a significant factor among others in my decision to pursue a graduate program.": 3,
    "It was the primary factor in my decision to pursue a graduate program.": 4,
    "It was the only reason I chose to pursue a graduate program.": 5,
}

FACTOR_MAP = {
    "Q30": ("150-credit education requirement influenced graduate enrollment", Q30),
    "Q39_1": ("Importance of CPA exam preparation in program value", IMPORTANCE5),
    "Q55": ("Belief that graduate degree increases lifetime earnings (ROI)", Q55),
    "Q53": ("Awareness of alternative CPA pathway before survey", YESNO),
    "Q49": ("Employer requires/encourages graduate degree", YESNO),
    "Q6": ("Perception of alternative pathway (fewer credits + extra work year)", Q6),
}

OPEN_ENDED_FIELDS = ["Q32", "Q36", "Q38", "Q45", "Q50", "Q56", "Q59", "Q9"]


def find_excel_files():
    files = []
    for base in DATA_CANDIDATES:
        if not os.path.isdir(base):
            continue
        for name in os.listdir(base):
            if name.lower().endswith(".xlsx"):
                files.append(os.path.join(base, name))
    return sorted(set(files))


def col_key(col):
    return (len(col), col)


def parse_xlsx(path):
    with zipfile.ZipFile(path) as z:
        ss = []
        sroot = ET.fromstring(z.read("xl/sharedStrings.xml"))
        for si in sroot.findall(f"{NS}si"):
            ss.append("".join((t.text or "") for t in si.iter(f"{NS}t")))
        sheet = ET.fromstring(z.read("xl/worksheets/sheet1.xml"))
        rows = []
        for r in sheet.findall(f".//{NS}row"):
            rec = {}
            for c in r.findall(f"{NS}c"):
                ref = c.attrib["r"]
                col = "".join(ch for ch in ref if ch.isalpha())
                t = c.attrib.get("t")
                v = c.find(f"{NS}v")
                raw = "" if v is None else (v.text or "")
                rec[col] = ss[int(raw)] if t == "s" and raw.isdigit() else raw
            rows.append(rec)

    cols = sorted(rows[1].keys(), key=col_key)
    names = {c: rows[1].get(c, "") for c in cols}
    questions = {names[c]: rows[2].get(c, "") for c in cols}

    data = []
    for rr in rows[4:]:
        d = {names[c]: (rr.get(c, "") or "").strip() for c in cols}
        data.append(d)
    return data, questions


def classify_field(values, name, qtext):
    nonempty = [v for v in values if v != ""]
    if not nonempty:
        return "other"
    uniq = set(nonempty)
    if name.endswith("_TEXT") or "Please explain" in qtext or "Please briefly" in qtext or "Please share" in qtext or max(len(v) for v in nonempty) > 120:
        return "text"
    if uniq.issubset(set(LIKERT5)|set(LIKELIHOOD5)|set(IMPORTANCE5)|set(Q55)|set(Q6)|set(Q30)):
        return "Likert"
    if uniq.issubset({"Yes","No"}) or len(uniq)<=10:
        return "multi-choice"
    return "other"


def pearson(x,y):
    if len(x) < 3:
        return float('nan')
    mx, my = statistics.mean(x), statistics.mean(y)
    num = sum((a-mx)*(b-my) for a,b in zip(x,y))
    denx = math.sqrt(sum((a-mx)**2 for a in x))
    deny = math.sqrt(sum((b-my)**2 for b in y))
    if denx == 0 or deny == 0:
        return float('nan')
    return num/(denx*deny)


def quote_theme(text):
    t = text.lower()
    if any(k in t for k in ["tuition", "afford", "loan", "cost", "debt", "expensive"]): return "cost"
    if any(k in t for k in ["time", "years", "delay", "full-time", "hours", "workload"]): return "time"
    if any(k in t for k in ["exam", "pass rate", "cpa prep", "study"]): return "exam difficulty"
    if any(k in t for k in ["employer", "job offer", "promotion", "firm"]): return "employer support"
    if any(k in t for k in ["150", "credit hour", "graduate degree", "master", "macc", "mba"]): return "education requirement"
    if any(k in t for k in ["experience", "work experience", "2 years", "extra year"]): return "work experience requirement"
    if any(k in t for k in ["earn", "salary", "lifetime", "return", "roi", "payoff", "career ladder"]): return "ROI/value"
    if any(k in t for k in ["aware", "know", "confus", "pathway", "understand"]): return "awareness/confusion"
    if any(k in t for k in ["low-income", "family", "access", "equity", "children"]): return "equity/access"
    return "other"


def redact(text):
    text = re.sub(r"[\w\.-]+@[\w\.-]+", "[REDACTED_EMAIL]", text)
    text = re.sub(r"\b\+?\d?[\d\-\(\) ]{8,}\b", "[REDACTED_PHONE]", text)
    return text.strip()


def png_bar_chart(values, labels, title, outpath, width=1200, height=700):
    # Minimal PNG renderer with simple bars/axes (no external libs).
    bg=(255,255,255); axis=(40,40,40); bar=(66,135,245)
    img=[ [list(bg) for _ in range(width)] for _ in range(height)]
    lm,rm,tm,bm=120,40,60,120
    plot_w=width-lm-rm; plot_h=height-tm-bm
    # axes
    for x in range(lm,lm+plot_w+1): img[tm+plot_h][x]=list(axis)
    for y in range(tm,tm+plot_h+1): img[y][lm]=list(axis)
    n=len(values); maxv=max(values) if values else 1
    bw=max(10,int(plot_w/(n*1.4)))
    gap=bw//2
    x=lm+gap
    for v in values:
        h=0 if maxv==0 else int((v/maxv)*(plot_h-10))
        for yy in range(tm+plot_h-h, tm+plot_h):
            for xx in range(x, min(x+bw,lm+plot_w)):
                img[yy][xx]=list(bar)
        x+=bw+gap
    # write png
    raw=b''
    for row in img:
        raw+=b'\x00'+bytes([c for px in row for c in px])
    def chunk(tag,data):
        return len(data).to_bytes(4,'big')+tag+data+zlib.crc32(tag+data).to_bytes(4,'big')
    png=b'\x89PNG\r\n\x1a\n'
    png+=chunk(b'IHDR', width.to_bytes(4,'big')+height.to_bytes(4,'big')+b'\x08\x02\x00\x00\x00')
    png+=chunk(b'tEXt', f'Title\x00{title}'.encode('latin1','ignore'))
    png+=chunk(b'IDAT', zlib.compress(raw,9))
    png+=chunk(b'IEND', b'')
    with open(outpath,'wb') as f: f.write(png)


def main():
    files = find_excel_files()
    if not files:
        raise SystemExit("No .xlsx files found in /data or repo root")

    all_data=[]; question_text={}
    for f in files:
        d,q=parse_xlsx(f)
        all_data.extend(d)
        question_text.update(q)

    rows=[r for r in all_data if r.get('Finished','')=='1']

    # data dictionary
    fields=sorted(rows[0].keys())
    with open(os.path.join(ROOT,'analysis','data_dictionary.csv'),'w',newline='') as fh:
        w=csv.writer(fh)
        w.writerow(['field_name','question_text_if_available','type','value_labels_if_any','missing_rate'])
        for field in fields:
            vals=[r.get(field,'') for r in rows]
            miss=sum(1 for v in vals if v=='')/len(rows)
            q=question_text.get(field,'') or field
            t=classify_field(vals, field, q)
            uniq=[]
            for v in vals:
                if v and v not in uniq: uniq.append(v)
            value_labels=' | '.join(uniq[:12])
            w.writerow([field,q,t,value_labels,f"{miss:.4f}"])

    # top factors
    intent_scores=[]
    for r in rows:
        v=r.get('Q29','')
        if v in LIKELIHOOD5: intent_scores.append((r,LIKELIHOOD5[v]))
    factor_out=[]
    for field,(label,mapv) in FACTOR_MAP.items():
        vals=[]; intents=[]
        for r,iscore in intent_scores:
            v=r.get(field,'')
            if v in mapv:
                vals.append(mapv[v]); intents.append(iscore)
        if not vals: continue
        cnt=Counter(vals)
        topbox=sum(1 for x in vals if x>=4)/len(vals)
        factor_out.append({
            'factor_field':field,
            'factor_label':label,
            'n':len(vals),
            'mean':round(statistics.mean(vals),3),
            'median':statistics.median(vals),
            'top_box_pct':round(topbox*100,1),
            'pearson_with_intent':round(pearson(vals,intents),3),
            'abs_corr':round(abs(pearson(vals,intents)),3) if not math.isnan(pearson(vals,intents)) else '',
            'distribution':json.dumps(dict(sorted(cnt.items())))
        })
    factor_out.sort(key=lambda x:(x['abs_corr'] if x['abs_corr']!='' else -1), reverse=True)
    with open(os.path.join(ROOT,'analysis','rq1_top_factors.csv'),'w',newline='') as fh:
        w=csv.DictWriter(fh,fieldnames=list(factor_out[0].keys()))
        w.writeheader(); w.writerows(factor_out)

    # figures
    top5=factor_out[:5]
    png_bar_chart([r['abs_corr'] for r in top5],[r['factor_field'] for r in top5], 'Top factors by absolute correlation with CPA intent', os.path.join(ROOT,'figures','rq1_top_factors_corr.png'))

    # segmentation full-time vs part-time using intent mean
    seg=defaultdict(list)
    for r,_ in intent_scores:
        if r.get('Q16','') in ('Full-time','Part-time'):
            seg[r['Q16']].append(LIKELIHOOD5[r['Q29']])
    seg_labels=sorted(seg)
    seg_vals=[round(statistics.mean(seg[k]),3) for k in seg_labels]
    png_bar_chart(seg_vals,seg_labels,'Mean CPA intent by enrollment status', os.path.join(ROOT,'figures','rq1_intent_by_status.png'))

    # awareness segmentation
    seg2=defaultdict(list)
    for r,_ in intent_scores:
        if r.get('Q53','') in ('Yes','No'):
            seg2[r['Q53']].append(LIKELIHOOD5[r['Q29']])
    s2_labels=['No','Yes']
    s2_vals=[round(statistics.mean(seg2[k]),3) if seg2.get(k) else 0 for k in s2_labels]
    png_bar_chart(s2_vals,s2_labels,'Mean CPA intent by pathway awareness', os.path.join(ROOT,'figures','rq1_intent_by_awareness.png'))

    # quotes
    quotes=[]
    qcount=defaultdict(int)
    max_per_theme=4
    for idx,r in enumerate(rows, start=5):
        rid=r.get('ResponseId','') or f'row_{idx}'
        for f in OPEN_ENDED_FIELDS:
            txt=redact(r.get(f,''))
            if len(txt)<30: continue
            theme=quote_theme(txt)
            if qcount[theme] >= max_per_theme: continue
            sent=txt.split('\n')[0].strip()
            if len(sent)>280: sent=sent[:277]+'...'
            quotes.append({'ResponseID_or_row':rid,'question_field':f,'quote':sent,'theme_label':theme})
            qcount[theme]+=1
    # enforce 2-4 for major themes where possible
    with open(os.path.join(ROOT,'analysis','rq1_quotes.csv'),'w',newline='') as fh:
        w=csv.DictWriter(fh, fieldnames=['ResponseID_or_row','question_field','quote','theme_label'])
        w.writeheader(); w.writerows(quotes)

    print('Created files:')
    for p in [
        'analysis/data_dictionary.csv','analysis/rq1_top_factors.csv','analysis/rq1_quotes.csv',
        'figures/rq1_top_factors_corr.png','figures/rq1_intent_by_status.png','figures/rq1_intent_by_awareness.png'
    ]:
        print('-',p)

if __name__=='__main__':
    main()
