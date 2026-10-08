#!/usr/bin/env python3
from __future__ import annotations
import argparse, base64, email, io, json, math, os, re, sys, shutil, subprocess, tempfile
from collections import defaultdict
from datetime import date, datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP
from email import policy
from email.parser import BytesParser
from email.utils import parseaddr, parsedate_to_datetime
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))
try:
    from grader.extract import pdf_pages, pdf_is_image_only, xlsx_pages, email_lines, attachments
except Exception:
    pdf_pages = xlsx_pages = email_lines = attachments = None

D = Decimal
CENT = D('0.01')
ZERO = D('0.00')


def q2(x):
    return D(str(x)).quantize(CENT, rounding=ROUND_HALF_UP)

def money(x):
    return f'{q2(x):.2f}'

def iso_date(s):
    m = re.search(r'(\d{4})[-/](\d{1,2})[-/](\d{1,2})', str(s))
    if m: return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    m = re.search(r'(\d{1,2})/(\d{1,2})/(\d{4})', str(s))
    if m: return date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    return None

def parse_human_date(s):
    s=str(s or '').strip()
    d=iso_date(s)
    if d: return d
    for fmt in ('%d %b %Y','%d %B %Y'):
        try:
            return datetime.strptime(s,fmt).date()
        except Exception:
            pass
    return None

def add_months(d, months):
    y = d.year + (d.month-1+months)//12
    m = (d.month-1+months)%12+1
    import calendar
    return date(y,m,min(d.day,calendar.monthrange(y,m)[1]))

def month_key(d): return f'{d.year:04d}-{d.month:02d}'
def quarter_key(d): return f'{d.year:04d}-Q{(d.month-1)//3+1}'

def to_usd(inv, amount, contract):
    amount = q2(amount)
    if inv.get('currency', 'USD') == 'USD':
        return amount, None
    d = inv.get('invoice_date')
    fx = (contract.get('fx') or {}).get(month_key(d) if d else '')
    if fx is None or fx == ZERO:
        return None, {'name': f"Annex F exchange-rate row for {month_key(d) if d else 'invoice month'}", 'holder': None}
    # Annex F is stated as local-currency units per USD; convert the whole invoice-level delta once.
    return q2(amount / fx), None

def norm(s): return ' '.join(str(s or '').split())

def name_from_email(v): return parseaddr(str(v or ''))[0].strip()

def email_addr(v): return parseaddr(str(v or ''))[1].lower()

def first_money(s):
    ms = re.findall(r'(?<![\w.])(?:USD|INR|\$)?\s*([0-9][0-9,]*\.\d{2})(?!\w)', s)
    return D(ms[0].replace(',','')) if ms else None

def all_money(s): return [D(x.replace(',','')) for x in re.findall(r'(?<![\w.])([0-9][0-9,]*\.\d{2})(?!\w)', s)]


def ocr_pdf(data):
    """Best-effort OCR fallback; used only for image-only PDFs."""
    try:
        import pymupdf
        with pymupdf.open(stream=data, filetype='pdf') as doc:
            out=[]
            for page in doc:
                pix=page.get_pixmap(matrix=pymupdf.Matrix(2.4,2.4), alpha=False)
                png=pix.tobytes('png')
                try:
                    import pytesseract
                    from PIL import Image
                    im=Image.open(io.BytesIO(png))
                    txt=pytesseract.image_to_string(im, config='--psm 6')
                except Exception:
                    with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as f:
                        f.write(png); p=f.name
                    try:
                        r=subprocess.run(['tesseract',p,'stdout','--psm','6'],text=True,capture_output=True,timeout=30)
                        txt=r.stdout
                    finally:
                        os.unlink(p)
                out.append(txt)
            return out
    except Exception:
        return ['']


def parse_email_bytes(data):
    return BytesParser(policy=policy.default).parsebytes(data)


def line_pages_pdf(data, scan=False):
    if not scan:
        import pymupdf
        with pymupdf.open(stream=data, filetype='pdf') as doc:
            pages=[]
            for p in doc:
                words=p.get_text('words'); words.sort(key=lambda w: ((w[1]+w[3])/2,w[0]))
                lines=[]
                for w in words:
                    mid=(w[1]+w[3])/2
                    if lines and abs(lines[-1][0]-mid)<=2.5: lines[-1][1].append(w)
                    else: lines.append([mid,[w]])
                pages.append('\n'.join(' '.join(w[4] for w in sorted(ws,key=lambda x:x[0])) for _,ws in lines))
            return pages
    return ocr_pdf(data)


def ingest_file(path: Path, batch: str, root: Path):
    rel = path.relative_to(root).as_posix() if path.is_relative_to(root) else f'{batch}/{path.name}'
    data=path.read_bytes()
    msg=parse_email_bytes(data)
    dt=parsedate_to_datetime(str(msg['Date']))
    received=dt.astimezone().date().isoformat() if dt.tzinfo else dt.date().isoformat()
    try: ep=email_lines(msg)
    except Exception:
        n,a=parseaddr(str(msg['From']))
        ep=[f'From: {n} <{a}>',f"To: {msg['To'] or ''}",f'Date: {dt.strftime("%Y-%m-%dT%H:%M:%SZ")}',f'Subject: {msg["Subject"]}', '']
        body=msg.get_body(preferencelist=('plain',))
        if body: ep += body.get_content().rstrip('\n').split('\n')
        names=[p.get_filename() for p in msg.iter_attachments()]
        if names: ep.append('Attachments: '+', '.join(names))
    rec={'doc':rel,'kind':'eml','received':received,'email':{'from':name_from_email(msg['From']),'from_addr':email_addr(msg['From']),'to':str(msg['To'] or ''),'subject':str(msg['Subject'] or ''),'body':'\n'.join(ep[4:])},'pages':['\n'.join(ep)],'attachments':[]}
    for part in msg.iter_attachments():
        name=part.get_filename() or ''
        if not name: continue
        ext=name.rsplit('.',1)[-1].lower() if '.' in name else ''
        adoc=f'{rel}#{name}'
        b=part.get_content()
        if ext=='pdf':
            try:
                scan=pdf_is_image_only(b)
            except Exception: scan=False
            pages=line_pages_pdf(b,scan)
            ak={'doc':adoc,'kind':'pdf','received':received,'parent':rel,'filename':name,'pages':pages,'scan':scan}
        elif ext=='xlsx':
            try: pages=xlsx_pages(b)
            except Exception: pages=['']
            ak={'doc':adoc,'kind':'xlsx','received':received,'parent':rel,'filename':name,'pages':pages,'scan':False}
        else:
            ak={'doc':adoc,'kind':ext or 'bin','received':received,'parent':rel,'filename':name,'pages':[b.decode('utf-8','ignore')], 'scan':False}
        rec['attachments'].append(ak)
    return rec


def save_json(path,obj):
    path.parent.mkdir(parents=True,exist_ok=True)
    path.write_text(json.dumps(obj,indent=2,ensure_ascii=False))

def load_json(path,default):
    try: return json.loads(path.read_text())
    except Exception: return default


def ingest(batch_dir,state_dir):
    state_dir.mkdir(parents=True,exist_ok=True)
    docs=load_json(state_dir/'documents.json',[])
    by={d['doc']:d for d in docs}
    batch=Path(batch_dir).name
    root=Path(batch_dir).parent
    for f in sorted(Path(batch_dir).glob('*.eml')):
        r=ingest_file(f,batch,root)
        by[r['doc']]=r
    save_json(state_dir/'documents.json',list(by.values()))
    print(json.dumps({'ok':True,'ingested_batch':batch,'documents':len(by)}))


def all_pages(docs, pattern=None):
    for d in docs:
        if pattern is None or pattern(d):
            yield d
            for a in d.get('attachments',[]): yield a


def evidence(doc,page,quote): return {'doc':doc,'page':int(page),'quote':str(quote)}


def choose_quote(page, terms, fallback=None, min_words=12):
    lines=[norm(x) for x in str(page).splitlines() if norm(x)]
    for line in lines:
        l=line.lower()
        if all(t.lower() in l for t in terms) and len(re.sub(r'\s+','',line))>=min_words:
            return line
    if fallback is not None: return fallback
    return lines[0] if lines else ''


def parse_contract(doc,a):
    pages='\n'.join(a['pages'])
    cid=a.get('filename','')
    m=re.search(r'(?:Agreement No\.|Agreement reference|Master Services Agreement .*?\))\s+([A-Z0-9]+-[A-Z0-9]+-[A-Z0-9-]+)',pages)
    if m: cid=m.group(1)
    if 'Agreement No.' in pages:
        m=re.search(r'Agreement No\.\s*([A-Z0-9/-]+)',pages)
        if m: cid=m.group(1)
    c={'id':cid,'doc':a['doc'],'pages':a['pages'],'effective':None,'expiry':None,'currency':None,'lanes':{},'fuel':{},'accessorials':{},'contacts':defaultdict(list),'rate_notices':[],'amendments':[],'rebates':[],'annual':None,'payment_days':30,'late_rate':None,'free_storage':0,'unranked':False}
    m=re.search(r'Effective date[^\n]*?(20\d\d-\d\d-\d\d)',pages,re.I) or re.search(r'effective on (20\d\d-\d\d-\d\d)',pages,re.I) or re.search(r'effect(?:ive)?(?: from)?\s+(20\d\d-\d\d-\d\d)',pages,re.I)
    if m: c['effective']=m.group(1)
    m=re.search(r'(?:expires|ends|ending)\s+(?:on\s+)?(20\d\d-\d\d-\d\d)',pages,re.I)
    if m: c['expiry']=m.group(1)
    m=re.search(r'Currency\s+(USD|INR)',pages)
    if m: c['currency']=m.group(1)
    # payment
    m=re.search(r'Payment terms.*?due\s+(\d+)\s+days',pages,re.I) or re.search(r'due\s+(\d+)\s+days',pages,re.I)
    if m: c['payment_days']=int(m.group(1))
    m=re.search(r'Late fee.*?(\d+\.\d+)%\s+per month',pages,re.I)
    if m: c['late_rate']=D(m.group(1))
    # storage free days
    m=re.search(r'(\d+) free days',pages,re.I)
    if m: c['free_storage']=int(m.group(1))
    c['unranked']=bool(re.search(r'Unranked instruments.*?rate notice.*?amendment|rate notice.*?amendment.*?neither prevails',pages,re.I|re.S))
    # rates: parse table-ish lines; allow 2 or more rates after weight/minimum
    for line in pages.splitlines():
        line=norm(line)
        mm=re.match(r'^([A-Z0-9]+-[A-Z0-9]+)\s+minimum charge\s+([0-9.,]+)\s+([0-9.,]+)',line,re.I)
        if mm:
            c['lanes'].setdefault(mm.group(1).upper(),{'minimum':{},'bands':{}})['minimum']={'P1':D(mm.group(2).replace(',','')),'P2':D(mm.group(3).replace(',',''))}
            continue
        mm=re.match(r'^([A-Z0-9]+-[A-Z0-9]+)\s+(0-44\.5|45-99\.5|100-299\.5|300\+)\s*kg\s+([0-9.,]+)\s+([0-9.,]+)',line,re.I)
        if mm:
            lane=mm.group(1).upper(); band=mm.group(2)
            c['lanes'].setdefault(lane,{'minimum':{},'bands':{}})['bands'][band]={'P1':D(mm.group(3).replace(',','')),'P2':D(mm.group(4).replace(',',''))}
    # Some invoices put lane before/after labels; fallback using regex over full lines
    for line in pages.splitlines():
        line=norm(line)
        mm=re.search(r'([A-Z0-9]+-[A-Z0-9]+)\s+(0-44\.5|45-99\.5|100-299\.5|300\+)\s*kg\s+([0-9.,]+)(?:\s+([0-9.,]+))?',line,re.I)
        if mm:
            lane,band=mm.group(1).upper(),mm.group(2); c['lanes'].setdefault(lane,{'minimum':{},'bands':{}})
            nums=[D(mm.group(i).replace(',','')) for i in (3,4) if mm.group(i)]
            if nums: c['lanes'][lane]['bands'][band]={'P1':nums[0],'P2':nums[1] if len(nums)>1 else nums[0]}
    # fuel
    for line in pages.splitlines():
        line=norm(line)
        mm=re.match(r'^(20\d\d-\d\d)\s+([0-9.]+)%$',line)
        if mm: c['fuel'][mm.group(1)]=D(mm.group(2))
    # accessorials
    for line in pages.splitlines():
        line=norm(line)
        mm=re.match(r'^(DG|DRYICE|SEC|STOR|DLV|CUSTOMS|ACCT|SPROG|PPROG|PEAK)\s+.*?\s+([0-9.,]+)\s+(?:USD|INR)',line,re.I)
        if mm: c['accessorials'][mm.group(1).upper()]=D(mm.group(2).replace(',',''))
        else:
            mm=re.match(r'^(DG|DRYICE|SEC|STOR|DLV|CUSTOMS|ACCT|SPROG|PPROG|PEAK)\s+.*?\s+([0-9.,]+)\s+per',line,re.I)
            if mm: c['accessorials'][mm.group(1).upper()]=D(mm.group(2).replace(',',''))
    # Annex F
    c['fx']={}
    for line in pages.splitlines():
        mm=re.match(r'^(20\d\d-\d\d)\s+([0-9.]+)$',norm(line))
        if mm and (mm.group(1) in c['fuel'] or 'Exchange rates' in pages): c['fx'][mm.group(1)]=D(mm.group(2))
    # rebate and annual
    for m in re.finditer(r'QS\s*>=\s*USD\s*([\d,]+\.\d{2})\s*:\s*([0-9.]+)%',pages,re.I):
        c['rebates'].append((D(m.group(1).replace(',','')),D(m.group(2))))
    m=re.search(r'credit equal to ([0-9.]+)% of the Qualifying Spend',pages,re.I)
    if m: c['annual']=D(m.group(1))
    # contacts Schedule1: capture simple rows
    for line in pages.splitlines():
        mm=re.match(r'^([^,]+),\s*([^,]+),\s*([^:]+):\s*([^<]+)<([^>]+)>',norm(line))
        if mm:
            person,title,party,cats,addr=mm.groups()
            for cat in re.split(r',\s*',cats): c['contacts'][cat.strip().lower()].append({'name':person.strip(),'title':title.strip(),'party':party.strip(),'email':addr.strip().lower()})
    return c


def parse_amendment(a):
    pages='\n'.join(a['pages'])
    cid=None
    m=re.search(r'Amendment (?:No\.\s*\d+ to .*?|reference)\s*([A-Z0-9/-]+)',pages,re.I)
    m2=re.search(r'to (?:Master Services Agreement .*? )?([A-Z0-9]+-[A-Z0-9]+-[A-Z0-9-]+)',pages,re.I)
    if m2: cid=m2.group(1)
    if not cid:
        ids=re.findall(r'\b([A-Z]{2,4}-[A-Z]{2,4}-[A-Z0-9-]+)\b',pages)
        cid=ids[0] if ids else None
    me=re.search(r'Effective from\s+(20\d\d-\d\d-\d\d)',pages,re.I)
    eff=me.group(1) if me else None
    rates=[]
    for line in pages.splitlines():
        line=norm(line)
        mm=re.match(r'^([A-Z0-9]+-[A-Z0-9]+)\s+(minimum charge|0-44\.5 kg|45-99\.5 kg|100-299\.5 kg|300\+ kg)\s+([0-9.,]+)\s*(?:USD|INR)?',line,re.I)
        if mm:
            rates.append((mm.group(1).upper(),mm.group(2).lower().replace(' kg',''),D(mm.group(3).replace(',',''))))
    return {'doc':a['doc'],'page':1,'contract_id':cid,'effective':eff,'rates':rates,'pages':a['pages']}


def collect_state(state_dir, asof):
    docs=load_json(state_dir/'documents.json',[])
    d=[x for x in docs if x.get('received','9999-12-31')<=asof]
    contracts={}
    amendments=[]
    for d0 in d:
        for a in d0.get('attachments',[]):
            fn=a.get('filename','').lower()
            txt='\n'.join(a.get('pages',[]))
            tl=txt.lower()
            is_msa=(a['kind']=='pdf' and 'master services agreement' in tl and 'schedule 1' in tl and 'this agreement is dated' in tl and 'annex a: express rates' in tl)
            is_amend=(a['kind']=='pdf' and ('amendment' in tl or 'amendment' in fn))
            if is_msa:
                c=parse_contract(d0,a)
                contracts[c['id']]=c
            elif is_amend:
                amendments.append(parse_amendment(a))
    # apply amendments
    for am in amendments:
        if not am['contract_id'] or am['contract_id'] not in contracts: continue
        contracts[am['contract_id']]['amendments'].append(am)
    # collect rate notices and special mails
    waivers=[]; extensions={}; rate_notices=[]; feedback=[]
    for d0 in d:
        e=d0.get('email',{}); body=d0.get('pages',[''])[0]; sender=e.get('from_addr',''); subj=e.get('subject','')
        for m in re.finditer(r'(?:waive|waived|waiver).*?(?:Delivery fee|Dry ice|DG handling|Storage|Customs(?: hold)? fee).*?(?:AWB|waybill)\s+([A-Z0-9-]+)',body,re.I|re.S):
            # Try charge name and AWB around sentence
            seg=re.search(r'(.{0,180}(?:waive|waived|waiver).{0,220})',body,re.I|re.S)
            s=seg.group(1).strip() if seg else m.group(0)
            ch='delivery fee' if 'delivery fee' in s.lower() else 'unknown'
            waivers.append({'email':d0['doc'],'date':d0['received'],'sender':name_from_email(sender) or e.get('from',''),'sender_addr':sender.lower(),'awb':m.group(1),'charge':ch,'quote':choose_quote(body,['waiv',m.group(1)])})
        m=re.search(r'(?:extend|extended).*?due date of invoice\s+([A-Z0-9_/-]+).*?(?:to|until)\s+(20\d\d-\d\d-\d\d)',body,re.I|re.S)
        if not m:
            m=re.search(r'(?:extend|extended).*?invoice\s+([A-Z0-9_/-]+).*?(?:to|until)\s+(20\d\d-\d\d-\d\d)',body,re.I|re.S)
        if not m:
            m=re.search(r'invoice\s+([A-Z0-9_/-]+)\s+now\s+falls\s+due\s+on\s+(20\d\d-\d\d-\d\d)',body,re.I|re.S)
        if m:
            extensions[m.group(1).upper()]=({'email':d0['doc'],'date':d0['received'],'date2':m.group(2),'quote':choose_quote(body,[m.group(1),m.group(2)])})
        if 'rate notice' in subj.lower() or re.search(r'under clause 3\.5',body,re.I):
            lane_m=re.search(r'(?:rates? for lane|rate notice)\s+([A-Z0-9]+-[A-Z0-9]+)',body,re.I)
            if not lane_m:
                lane_m=re.search(r'apply to lane\s+([A-Z0-9]+-[A-Z0-9]+)',body,re.I)
            lane=lane_m.group(1).upper() if lane_m else None
            m=re.search(r'for\s+pickups?\s+from\s+(20\d\d-\d\d-\d\d)\s+(?:to|until)\s+(20\d\d-\d\d-\d\d)',body,re.I|re.S)
            if not m:
                m=re.search(r'(?:effective|valid)\s+(20\d\d-\d\d-\d\d).*?(?:to|until)\s+(20\d\d-\d\d-\d\d)',body,re.I|re.S)
            if not m:
                m=re.search(r'(20\d\d-\d\d-\d\d).*?(?:to|until)\s+(20\d\d-\d\d-\d\d)',body,re.I|re.S)
            vf,vt=(m.groups() if m else (None,None))
            rr=[]
            for ln in body.splitlines():
                ln=norm(ln)
                mm=re.search(r'([A-Z0-9]+-[A-Z0-9]+)\s+(minimum charge|0-44\.5 kg|45-99\.5 kg|100-299\.5 kg|300\+ kg)\s*(?::|=)\s*(?:(?:USD|INR)\s*)?([0-9.]+)\s*(?:USD|INR)?',ln,re.I)
                if mm: rr.append((mm.group(1).upper(),mm.group(2).lower().replace(' kg',''),D(mm.group(3))))
            rate_quote=next((norm(ln) for ln in body.splitlines() if lane and lane.lower() in norm(ln).lower() and re.search(r'(minimum charge|0-44\.5|45-99\.5|100-299\.5|300\+)',norm(ln),re.I)), None)
            rate_notices.append({'email':d0['doc'],'date':d0['received'],'sender':e.get('from') or name_from_email(sender),'sender_addr':sender.lower(),'lane':lane,'from':vf,'to':vt,'rates':rr,'quote':rate_quote or choose_quote(body,[lane] if lane else ['rate notice'])})
    # feedback files delivered to state_dir/feedback
    for fp in sorted((state_dir/'feedback').glob('*.json')) if (state_dir/'feedback').is_dir() else []:
        feedback.append(load_json(fp,{}))
    # add top-level authorized info to rate notices/waivers
    for rn in rate_notices:
        c=contracts.get(next((cid for cid in contracts if cid in rn['quote'] or cid.lower() in rn['quote'].lower()),None))
        if c is None:
            for cid,cx in contracts.items():
                if rn['lane'] and rn['lane'] in cx['lanes']: c=cx; break
        rn['authorized']=False; rn['contract_id']=c['id'] if c else None
        if c:
            addr=rn['sender_addr']; rn['authorized']=any(p['email']==addr for p in c['contacts'].get('rates',[]))
            sender_name=(rn.get('sender') or '').lower()
            for fb in feedback:
                for item in (fb.get('items') or []):
                    fact=item.get('fact') or {}; subj=fact.get('subject') or {}; value=fact.get('value') or {}; scope=item.get('scope') or {}
                    if fact.get('type')!='sender_authority' or 'rates' not in (value.get('categories') or []): continue
                    if subj.get('party') and c['id'].split('-')[0].upper()!=str(subj.get('party')).split('-')[0].upper(): continue
                    if subj.get('person') and str(subj.get('person')).lower()!=sender_name: continue
                    vf=scope.get('valid_from'); vt=scope.get('valid_to') or '9999-12-31'
                    if vf and vf <= rn.get('date','') <= vt: rn['authorized']=True
    return d,contracts,amendments,waivers,extensions,rate_notices,feedback


def find_awb(line):
    # AWBs in the dataset use CA..., TX... or TE... identifiers; do not treat dates/rates as AWBs.
    toks=re.findall(r'\b(?:CA|TX|TE)[A-Z0-9-]{6,}\b', str(line), re.I)
    return toks[0].upper() if toks else None


def parse_invoice(a):
    txt='\n'.join(a['pages'])
    m=re.search(r'Charges advice\s+([A-Z0-9_/-]+)',txt,re.I)
    if not m:
        m=re.search(r'Invoice\s+([A-Z0-9_/-]+)',txt,re.I)
    if not m:
        m=re.search(r'STATEMENT OF ACCOUNT\s+([A-Z0-9_/-]+)',txt,re.I)
    if not m:
        m=re.search(r'Freight bill no\.\s*([A-Z0-9_/-]+)',txt,re.I)
    iid=m.group(1).upper() if m else None
    # Scanned invoices can lose the numeric suffix in OCR; when the attachment
    # filename is an invoice-like id and clearly extends the extracted prefix,
    # prefer the printed filename as the authoritative invoice id.
    if iid:
        fnstem=Path(a.get('filename','')).stem.upper()
        if fnstem.startswith(iid+'-') and re.search(r'\d', fnstem):
            iid=fnstem
    md=re.search(r'Document date\s+(\d\d/\d\d/\d{4})\s+Payable by\s+(\d\d/\d\d/\d{4})',txt)
    if md: inv_date=iso_date(md.group(1)); due=iso_date(md.group(2))
    else:
        md2=re.search(r'(?:Invoice date|document date|Issued on)\s*[: ]?\s*(20\d\d[-/]\d\d[-/]\d\d|\d{1,2} [A-Za-z]{3,9} \d{4})',txt,re.I); inv_date=parse_human_date(md2.group(1)) if md2 else None
        md3=re.search(r'(?:due|payable by|Settle by)\s*(?:date)?\s*[: ]?\s*(20\d\d[-/]\d\d[-/]\d\d|\d{1,2} [A-Za-z]{3,9} \d{4})',txt,re.I); due=parse_human_date(md3.group(1)) if md3 else None
    cm=re.search(r'(?:Contract(?:\s+reference)?|Service agreement)\s+([A-Z0-9_/-]+)(?:\s+Currency\s+(USD|INR))?',txt,re.I)
    contract=cm.group(1).upper() if cm else None; currency=cm.group(2).upper() if cm and cm.group(2) else ('INR' if re.search(r'Amount \(INR\)|Total INR|Charges INR',txt,re.I) else 'USD')
    tm=re.search(r'(?:Total payable|Total|Charges|Net charges|Balance due)\s+(?:USD|INR)\s+([0-9,]+\.\d{2})',txt,re.I)
    total=D(tm.group(1).replace(',','')) if tm else None
    charges=[]
    for line in txt.splitlines():
        line=norm(line)
        if not line or line.lower().startswith(('route waybill','charges advice','document date','contract ','customer ','terms ','freight bill no','issued on ','service agreement','account ')): continue
        # Alternate carrier layout: Freight bill no. / Charge Shipment Route Collected Net / Units Unit price.
        alt=re.match(r'^Air freight\s+((?:CA|TX|TE)[A-Z0-9-]+)\s+([A-Z0-9]+-[A-Z0-9]+)\s+(\d{1,2} [A-Za-z]{3,9} \d{4})\s+([0-9,]+\.\d{2})\s+([0-9]+(?:\.[0-9]+)?) kg\s+([0-9,]+\.\d{2})$',line,re.I)
        if alt:
            charges.append({'type':'freight','awb':alt.group(1).upper(),'route':alt.group(2).upper(),'kg':D(alt.group(5)),'rate':D(alt.group(6).replace(',','')),'amount':D(alt.group(4).replace(',','')),'date':parse_human_date(alt.group(3)),'line':line})
            continue
        alt=re.match(r'^Fuel surcharge\s+((?:CA|TX|TE)[A-Z0-9-]+)\s+([A-Z0-9]+-[A-Z0-9]+)\s+(\d{1,2} [A-Za-z]{3,9} \d{4})\s+([0-9,]+\.\d{2})\s+[0-9,]+\.\d{2}\s+([0-9.]+)%$',line,re.I)
        if alt:
            charges.append({'type':'fuel','awb':alt.group(1).upper(),'pct':D(alt.group(5)),'amount':D(alt.group(4).replace(',','')),'line':line}); continue
        alt=re.match(r'^Security surcharge\s+((?:CA|TX|TE)[A-Z0-9-]+)\s+([A-Z0-9]+-[A-Z0-9]+)\s+(\d{1,2} [A-Za-z]{3,9} \d{4})\s+([0-9,]+\.\d{2})\s+([0-9]+(?:\.[0-9]+)?) kg\s+([0-9,]+\.\d{2})$',line,re.I)
        if alt:
            charges.append({'type':'security','awb':alt.group(1).upper(),'kg':D(alt.group(5)),'rate':D(alt.group(6).replace(',','')),'amount':D(alt.group(4).replace(',','')),'line':line}); continue
        alt=re.match(r'^(Delivery fee|Dry ice handling fee|Peak season surcharge)\s+((?:CA|TX|TE)[A-Z0-9-]+)\s+([A-Z0-9]+-[A-Z0-9]+)\s+(\d{1,2} [A-Za-z]{3,9} \d{4})\s+([0-9,]+\.\d{2})\s+.*$',line,re.I)
        if alt:
            typ={'delivery fee':'DLV','dry ice handling fee':'DRYICE','peak season surcharge':'PEAK'}[alt.group(1).lower()]
            charges.append({'type':typ,'awb':alt.group(2).upper(),'amount':D(alt.group(5).replace(',','')),'line':line}); continue
        alt=re.match(r'^Late fee on\s+([A-Z0-9_/-]+):.*$',line,re.I)
        if alt:
            clean_late=re.sub(r'\b[0-9.]+%','',line)
            ms=all_money(clean_late)
            if ms: charges.append({'type':'late','ref':alt.group(1).upper(),'amount':ms[-1],'line':line})
            continue
        if 'Air freight' in line or 'Express freight' in line or re.search(r'\bFreight charge\b',line,re.I):
            route_m=re.search(r'\b([A-Z][A-Z0-9]{2,}-[A-Z][A-Z0-9]{2,})\b',line)
            kg_m=re.search(r'([0-9]+(?:\.[0-9]+)?)\s*kg',line,re.I)
            if not (route_m and kg_m): continue
            nums=all_money(line[kg_m.end():])
            if len(nums)<2: nums=all_money(line)
            rate=nums[0] if len(nums)>=2 else None; amount=nums[1] if len(nums)>=2 else None
            awb=find_awb(line)
            dates=re.findall(r'(?:20\d\d[-/]\d\d[-/]\d\d|\d\d/\d\d/\d{4})',line)
            pdate=iso_date(dates[-1]) if dates else None
            charges.append({'type':'freight','awb':awb,'route':route_m.group(1).upper(),'kg':D(kg_m.group(1)),'rate':rate,'amount':amount,'date':pdate,'line':line})
        elif 'Fuel surcharge' in line or re.search(r'\bFSC\b',line):
            nums=all_money(line); pct=re.search(r'([0-9.]+)%',line)
            if nums:
                awb=find_awb(line)
                charges.append({'type':'fuel','awb':awb,'pct':D(pct.group(1)) if pct else None,'amount':nums[-1],'line':line})
        elif 'Security surcharge' in line or 'Screening fee' in line or 'Security screening' in line:
            nums=all_money(line); kg_m=re.search(r'([0-9]+(?:\.[0-9]+)?)\s*kg',line,re.I)
            if nums:
                awb=find_awb(line)
                charges.append({'type':'security','awb':awb,'kg':D(kg_m.group(1)) if kg_m else None,'rate':nums[-2] if len(nums)>=2 else None,'amount':nums[-1],'line':line})
        else:
            typ=None
            for label,t in [('DG handling','DG'),('Hazmat handling','DG'),('Dangerous goods fee','DG'),('Dry ice handling fee','DRYICE'),('Dry ice handling','DRYICE'),('Storage','STOR'),('Warehouse storage','STOR'),('Demurrage','STOR'),('Delivery fee','DLV'),('Door delivery','DLV'),('Final-mile delivery','DLV'),('Customs hold fee','CUSTOMS'),('Customs examination','CUSTOMS'),('Account maintenance','ACCT'),('Monthly account fee','ACCT'),('Storage programme fee','SPROG'),('Pickup programme fee','PPROG'),('Peak season surcharge','PEAK')]:
                if label.lower() in line.lower(): typ=t; break
            if typ:
                nums=all_money(line); days=None
                dm=re.search(r'([0-9]+)\s*days',line,re.I); sm=re.search(r'([0-9]+)\s*shpt',line,re.I)
                if nums:
                    # last monetary number is amount; for lines like base rate amount use last one
                    awb=find_awb(line)
                    charges.append({'type':typ,'awb':awb,'amount':nums[-1],'days':int(dm.group(1)) if dm else None,'line':line})
            if 'Late fee on' in line or 'Late fee ' in line or re.search(r'Late payment fee',line,re.I):
                mm=re.search(r'Late fee on\s+([A-Z0-9_/-]+).*?(?:USD|INR)\s*([0-9,]+\.\d{2})',line,re.I)
                if not mm:
                    mm=re.search(r'(?:late fee on|late payment fee).*?([A-Z0-9_/-]+).*?([0-9,]+\.\d{2})',line,re.I)
                if mm:
                    clean_late=re.sub(r'\b[0-9.]+%','',line)
                    ms=all_money(clean_late)
                    charges.append({'type':'late','ref':mm.group(1).upper(),'amount':ms[-1] if ms else D(mm.group(2).replace(',','')),'line':line})
    return {'doc':a['doc'],'attachment':a,'id':iid,'invoice_date':inv_date,'due':due,'contract':contract,'currency':currency,'total':total,'charges':charges,'pages':a['pages']}


def parse_xlsx(a):
    rows=[]
    for pi,p in enumerate(a['pages'],1):
        for ln in p.splitlines():
            cells=[x.strip() for x in ln.split(' | ')]
            rows.append((pi,cells,ln.strip()))
    return rows


def build_indexes(docs):
    invoices=[]; tracking=[]; ap=[]; dgd={}
    for d in docs:
        for a in d.get('attachments',[]):
            fn=a.get('filename','').lower()
            txt='\n'.join(a.get('pages',[]))
            if a['kind']=='pdf' and re.search(r'charges advice|invoice|statement of account|freight bill no\.',txt,re.I) and not re.search(r'master services agreement|amendment no\.',txt,re.I):
                inv=parse_invoice(a)
                # Require a real billing header; do not treat contract clauses containing the word 'invoice' as bills.
                if inv['id'] and (re.search(r'(?:Charges advice|Invoice|Freight bill no\.|STATEMENT OF ACCOUNT)\s*[-#A-Z0-9]',txt,re.I)):
                    invoices.append(inv)
            if a['kind']=='xlsx':
                rows=parse_xlsx(a)
                xtext='\n'.join(a.get('pages',[]))
                sm=re.search(r'STATEMENT OF ACCOUNT\s+([A-Z0-9_/-]+)',xtext,re.I)
                cm=re.search(r'Contract reference\s+([A-Z0-9_/-]+)',xtext,re.I)
                dm=re.search(r'Statement date\s+(20\d\d-\d\d-\d\d)\s*\|\s*Due date\s+(20\d\d-\d\d-\d\d)',xtext,re.I)
                tm=re.search(r'Total (USD|INR)\s*\|\s*([0-9,]+\.\d{2})',xtext,re.I)
                if sm and cm and dm and tm:
                    scharges=[]
                    for ln2 in xtext.splitlines():
                        nln=norm(ln2)
                        typ=None
                        for label,t in [('Monthly account fee','ACCT'),('Storage programme fee','SPROG'),('Pickup programme fee','PPROG')]:
                            if nln.lower().startswith(label.lower()): typ=t; break
                        if typ:
                            ms=all_money(nln)
                            if ms: scharges.append({'type':typ,'awb':None,'amount':ms[-1],'line':nln})
                    invoices.append({'doc':a['doc'],'attachment':a,'id':sm.group(1).upper(),'invoice_date':iso_date(dm.group(1)),'due':iso_date(dm.group(2)),'contract':cm.group(1).upper(),'currency':tm.group(1).upper(),'total':D(tm.group(2).replace(',','')),'charges':scharges,'pages':a['pages']})
                for pi,cells,ln in rows:
                    if len(cells)>=4 and (cells[1] in ('PICKED_UP','ARRIVED_FACILITY','OUT_FOR_DELIVERY','DELIVERED','DEPARTED','COLLECTED_BY_CONSIGNEE')):
                        tracking.append({'awb':cells[0],'event':cells[1],'date':iso_date(cells[2]),'station':cells[3],'doc':a['doc'],'page':pi,'line':ln})
                    if cells and re.fullmatch(r'[A-Z0-9_/-]+',cells[0] or '') and ('due' in ln.lower() and len(cells)>=6):
                        # AP row: invoice, invoice date, due..., amount, status, paid
                        if len(cells)>=5 and re.search(r'\b(paid|open)\b',ln,re.I):
                            iid=cells[0].upper(); invd=iso_date(cells[1]); duem=re.search(r'(20\d\d-\d\d-\d\d)',cells[2]); due=iso_date(duem.group(1)) if duem else None
                            am=None
                            if len(cells)>=4:
                                mm=re.search(r'(?:USD|INR)\s*([0-9,]+\.\d{2})',cells[3]); am=D(mm.group(1).replace(',','')) if mm else None
                            status='open' if re.search(r'\bopen\b',ln,re.I) else 'paid'; paid=None
                            if len(cells)>=6: paid=iso_date(cells[5]) if cells[5].strip() not in ('-','') else None
                            ap.append({'invoice_id':iid,'invoice_date':invd,'due':due,'amount':am,'status':status,'paid':paid,'currency': 'INR' if 'INR' in ln else 'USD','doc':a['doc'],'page':pi,'line':ln})
            if a['kind']=='pdf' and (fn.startswith('dgd-') or "shipper's declaration" in txt.lower()):
                awbs=re.findall(r'\b(?:CA|TX|TE)[A-Z0-9-]{6,}\b',txt,re.I)
                if awbs:
                    dgd[awbs[0].upper()]={'doc':a['doc'],'page':1,'text':txt,'scan':a.get('scan',False)}
    return invoices,tracking,ap,dgd


def contract_rate(c,lane,kg,pdate):
    if not c or lane not in c['lanes']: return None
    band='0-44.5' if kg<=D('44.5') else '45-99.5' if kg<=D('99.5') else '100-299.5' if kg<=D('299.5') else '300+'
    p2=False
    # periods can be inferred from effective+12mo, but table text gives dates in source
    # default: P1 first 6 months for practice, or detect from MSA text.
    txt='\n'.join(c['pages'])
    m1=re.search(r'Period P1: pickups\s+(20\d\d-\d\d-\d\d)\s+to\s+(20\d\d-\d\d-\d\d)',txt,re.I)
    m2=re.search(r'Period P2: pickups\s+(20\d\d-\d\d-\d\d)\s+to\s+(20\d\d-\d\d-\d\d)',txt,re.I)
    if m2 and iso_date(m2.group(1))<=pdate<=iso_date(m2.group(2)): p2=True
    elif m1: p2=False
    else:
        eff=iso_date(c.get('effective') or '2000-01-01'); p2=pdate>=add_months(eff,6)
    return band, c['lanes'][lane]['bands'].get(band,{}).get('P2' if p2 else 'P1'), c['lanes'][lane]['minimum'].get('P2' if p2 else 'P1')


def find_contract(catalog,iid):
    return catalog.get(iid)


def normalize_lane(lane, contract):
    """Correct a small OCR typo in an extracted lane by matching contract lanes."""
    if not lane or not contract: return lane
    lane=lane.upper()
    if lane in contract.get('lanes',{}): return lane
    candidates=list(contract.get('lanes',{}))
    scored=[]
    for cand in candidates:
        if len(cand)!=len(lane): continue
        dist=sum(a!=b for a,b in zip(cand,lane))
        if dist<=1: scored.append((dist,cand))
    if len(scored)==1: return scored[0][1]
    return lane


def tracking_pickup(track,awb,fallback):
    ev=[x for x in track if x['awb'].upper()==awb.upper() and x['event']=='PICKED_UP' and x.get('date')]
    return min(x['date'] for x in ev) if ev else fallback

def tracking_event(track,awb,event):
    es=[x for x in track if x['awb'].upper()==awb.upper() and x['event']==event and x.get('date')]
    return min(es,key=lambda x:x['date']) if es else None


def date_for_invoice(inv,track):
    f=[]
    for ch in inv['charges']:
        if ch.get('type')=='freight' and ch.get('awb'):
            f.append(tracking_pickup(track,ch['awb'],ch.get('date') or inv['invoice_date']))
    return f[0] if f else inv['invoice_date']


def matching_rates(c,cid,lane,pdate,notices,feedback):
    base=contract_rate(c,lane,D('0'),pdate) if False else None
    applicable=[]
    for rn in notices:
        if rn.get('contract_id')==cid and rn.get('lane')==lane and rn.get('from') and rn.get('to') and rn['from']<=pdate.isoformat()<=rn['to']:
            applicable.append(rn)
    # feedback accepted rate notice overrides/validates same lane, time window
    accepted=[]
    for fb in feedback:
        for item in (fb.get('items') or []):
            fact=item.get('fact') or {}; val=fact.get('value') or {}; subj=fact.get('subject') or {}; scope=item.get('scope') or {}
            if fact.get('type')=='rate_accepted' and subj.get('contract_id')==cid and subj.get('lane')==lane:
                vf=scope.get('valid_from'); vt=scope.get('valid_to')
                if vf and vt and vf<=pdate.isoformat()<=vt:
                    accepted.append({'value':val.get('source'),'from':vf,'to':vt,'sent_by':val.get('sent_by'),'feedback':fb,'item':item})
    return applicable,accepted


def effective_rate_value(c,cid,lane,band,pdate,notices,feedback):
    # base from contract, using a representative weight in the requested band.
    rep={'0-44.5':D('20'),'45-99.5':D('70'),'100-299.5':D('150'),'300+':D('300')}[band]
    b=contract_rate(c,lane,rep,pdate)
    if not b: return None,None,None
    _, base_rate, minimum=b
    am_rates=[]
    for am in c.get('amendments',[]):
        if am['effective'] and iso_date(am['effective'])<=pdate:
            vals=[r[2] for r in am['rates'] if r[0]==lane and r[1]==band]
            if vals: am_rates.append((iso_date(am['effective']),vals[-1],am))
            vals2=[r[2] for r in am['rates'] if r[0]==lane and r[1]=='minimum charge']
            if vals2: minimum=vals2[-1]
    if am_rates:
        am_rates.sort(key=lambda x:x[0]); am_rate=am_rates[-1][1]; chosen=('amendment',am_rates[-1][2])
    else:
        am_rate=base_rate; chosen=('msa',None)
    notices_a,accepted=matching_rates(c,cid,lane,pdate,notices,feedback)
    notices_auth=[rn for rn in notices_a if rn.get('authorized')]
    # A human can resolve an otherwise unranked amendment/rate-notice conflict for this contract/lane.
    for fb in feedback or []:
        for item in (fb.get('items') or []):
            fact=item.get('fact') or {}; subj=fact.get('subject') or {}; value=fact.get('value') or {}; scope=item.get('scope') or {}
            if fact.get('type')!='precedence' or subj.get('contract_id')!=cid or subj.get('lane')!=lane: continue
            vf=scope.get('valid_from'); vt=scope.get('valid_to') or '9999-12-31'
            if not vf or not (vf <= pdate.isoformat() <= vt): continue
            use=str(value.get('use') or '').lower()
            if use=='amendment' and am_rates:
                return am_rate,minimum,{'amendment':am_rates[-1][2],'feedback':item}
            if 'notice' in use and notices_a:
                rn=notices_a[-1]; nvals=[r[2] for r in rn['rates'] if r[0]==lane and r[1]==band]
                if nvals: return nvals[-1],minimum,{'notice':rn,'feedback':item}
    # Human feedback can ratify a rate notice even when the original email sender is not a Schedule-1 authority.
    if accepted and notices_a:
        rn=notices_a[-1]
        nvals=[r[2] for r in rn['rates'] if r[0]==lane and r[1]==band]
        if nvals:
            # If both signed amendment and rate notice differ, contract says conflict.
            if am_rates and nvals[-1] != am_rate:
                return ('conflict',(am_rate,nvals[-1]),{'amendment':am_rates[-1][2],'notice':rn,'accepted':accepted[-1]})
            return nvals[-1],minimum,{'notice':rn,'accepted':accepted[-1]}
    if notices_auth:
        rn=notices_auth[-1]
        nvals=[r[2] for r in rn['rates'] if r[0]==lane and r[1]==band]
        if nvals:
            if am_rates and nvals[-1] != am_rate:
                return ('conflict',(am_rate,nvals[-1]),{'amendment':am_rates[-1][2],'notice':rn})
            # rate notice wins when no conflict and valid
            return nvals[-1],minimum,{'notice':rn}
    return am_rate,minimum,chosen


def source_quote_for_rate(c, lane, band):
    for pi,p in enumerate(c['pages'],1):
        q=choose_quote(p,[lane,band])
        if q and lane.lower() in q.lower() and band.lower() in q.lower(): return (pi,q)
    return (1,lane+' '+band)


def citation_for_line(inv, substring=None):
    for pi,p in enumerate(inv['pages'],1):
        for ln in p.splitlines():
            n=norm(ln)
            if substring is None or substring.lower() in n.lower(): return evidence(inv['doc'],pi,n)
    return evidence(inv['doc'],1,choose_quote(inv['pages'][0],[]))


def authorized_waiver(w,contracts,feedback=None):
    feedback=feedback or []
    for cid,c in contracts.items():
        if any(p['email']==w['sender_addr'] for p in c['contacts'].get('waivers',[])):
            return cid,c
        sender_name=str(w.get('sender') or '').strip().lower()
        for fb in feedback:
            for item in (fb.get('items') or []):
                fact=item.get('fact') or {}; subj=fact.get('subject') or {}; value=fact.get('value') or {}; scope=item.get('scope') or {}
                if fact.get('type')!='sender_authority' or 'waivers' not in (value.get('categories') or []): continue
                if subj.get('person') and str(subj.get('person')).strip().lower()!=sender_name: continue
                if subj.get('party') and c['id'].split('-')[0].upper()!=str(subj.get('party')).split('-')[0].upper(): continue
                vf=scope.get('valid_from'); vt=scope.get('valid_to') or '9999-12-31'
                if vf and vf <= w.get('date','') <= vt:
                    return cid,c
    return None,None


def feedback_items(feedback):
    out=[]
    for fb in feedback or []:
        for item in (fb.get('items') or []):
            x=dict(item); x['_feedback']=fb; out.append(x)
    return out


def feedback_for_key(items,key):
    out=[]
    for item in items:
        fk=item.get('finding_key') or {};
        ok=True
        for k,v in key.items():
            if fk.get(k)!=v: ok=False; break
        if ok: out.append(item)
    return out


def feedback_authorized(sender_name,party,category,asof,contract_id,items):
    for item in items:
        fact=item.get('fact') or {}; subj=fact.get('subject') or {}; value=fact.get('value') or {}; scope=item.get('scope') or {}
        if fact.get('type')!='sender_authority' or category not in (value.get('categories') or []): continue
        if subj.get('person') and str(subj.get('person')).strip().lower()!=str(sender_name or '').strip().lower(): continue
        if subj.get('party') and contract_id and contract_id.split('-')[0].upper()!=str(subj.get('party')).split('-')[0].upper(): continue
        vf=scope.get('valid_from'); vt=scope.get('valid_to') or '9999-12-31'
        if vf and vf <= asof <= vt: return True
    return False


def make_item(cls, key, verdict, amount=None, action='none', conf=.95, evidence_list=None, variant=None, deadline=None, amount_range=None, missing_doc=None, supersedes=None, learned_at=None):
    iid=key.get('invoice_id'); cid=key.get('contract_id'); period=key.get('period'); ct=key.get('credit_type')
    if cls in ('cashback',): idv=f'cashback:{cid}/{period}/{ct}'
    elif cls=='expiry': idv=f'expiry:{cid}'
    else: idv=f'{cls}:{iid}'
    x={'id':idv,'class':cls,'variant':variant or ('LF-03' if cls=='late_fee_risk' else cls.upper()),'verdict':verdict,'amount_usd':money(amount) if amount is not None else None,'amount_range':[money(x) for x in amount_range] if amount_range else None,'deadline':deadline,'action':action,'confidence':conf,'evidence':evidence_list or [],'missing_doc':missing_doc,'supersedes':supersedes or [],'learned_at':learned_at}
    if iid: x['invoice_id']=iid
    if cid: x['contract_id']=cid
    if period: x['period']=period
    if ct: x['credit_type']=ct
    return x


def decide_action(cls, amount, mode, evidence_complete=True, conflict=False):
    if conflict: return 'escalate'
    if cls=='expiry': return 'internal:alert'
    if cls=='late_fee_risk': return 'internal:payment_reminder'
    if not evidence_complete: return 'execute:request_document' if mode=='autonomous' else 'draft:request_document'
    if amount is None: return 'escalate'
    if amount >= D('5000'): return 'human_review'
    a='dispute' if cls in ('overcharge','duplicate','waived_but_billed','late_fee_not_owed') else 'claim_credit'
    return ('execute:'+a) if mode=='autonomous' else ('draft:'+a)


def report(asof,mode,state_dir):
    docs,contracts,amendments,waivers,extensions,rate_notices,feedback=collect_state(state_dir,asof)
    invoices,track,ap,dgd=build_indexes(docs)
    apmap={x['invoice_id']:x for x in ap}
    fitems=feedback_items(feedback)
    # extension authorization: Schedule 1 payment_terms, plus later feedback that grants the same authority.
    valid_ext={}
    for iid,e in extensions.items():
        for cid,c in contracts.items():
            sender_addr=next((d.get('email',{}).get('from_addr','') for d in docs if d['doc']==e['email']), '')
            sender_name=next((d.get('email',{}).get('from','') for d in docs if d['doc']==e['email']), '')
            if any(p['email']==sender_addr for p in c['contacts'].get('payment_terms',[])) or feedback_authorized(sender_name,c['id'].split('-')[0],'payment_terms',e.get('date',asof),c['id'],fitems):
                valid_ext[iid]=e['date2']; break
    findings=[]
    inv_by_id={i['id']:i for i in invoices if i.get('id')}
    # pre-index waivers by AWB and sender authority
    vwaivers=[]
    for w in waivers:
        cid,c=authorized_waiver(w,contracts,fitems)
        if cid: vwaivers.append((w,cid,c))
    # overcharge, waiver, late fees
    for inv in invoices:
        cid=inv['contract']; c=contracts.get(cid)
        if not c: continue
        iid=inv['id']; inv_received=next((d['received'] for d in docs for a in d.get('attachments',[]) if a['doc']==inv['doc']),asof)
        # avoid invoices dated after asof though docs received before asof are generally valid
        pickup=None
        for ch in inv['charges']:
            if ch.get('type')=='freight' and ch.get('awb'):
                pickup=tracking_pickup(track,ch['awb'],ch.get('date') or inv['invoice_date']); break
        if pickup is None: pickup=inv['invoice_date']
        # Missing amended attachment: detect explicit "applies ... attached" with no matching document; only affects lanes in amendment without actual rates.
        text_contract='\n'.join(c['pages'])
        for am in c.get('amendments',[]):
            atxt='\n'.join(am.get('pages',[]))
            miss=re.search(r'(Annex A rev 2).*?applies from\s+(20\d\d-\d\d-\d\d).*?lane\(s\)\s+([A-Z0-9]+-[A-Z0-9]+)',atxt,re.I|re.S)
            if miss and pickup and pickup>=iso_date(miss.group(2)) and pickup<=date(2099,12,31):
                lane=miss.group(3).upper()
                if any(ch.get('route')==lane for ch in inv['charges'] if ch.get('type')=='freight'):
                    # ensure no rate notice gives complete replacement; absent attachment is still determining source missing
                    rates_here=[x for x in rate_notices if x.get('contract_id')==cid and x.get('lane')==lane and x.get('authorized') and x.get('from') and x.get('from')<=pickup.isoformat()<=x.get('to','9999')]
                    if not rates_here:
                        person=(c['contacts'].get('rates') or [{}])[0]
                        # held by carrier rates contact; party from schedule
                        md={'name':miss.group(1),'holder':{'party':person.get('party',cid.split('-')[0]),'person':person.get('name',''),'role':person.get('title','Pricing Manager')}}
                        ev=evidence(am['doc'],1,choose_quote(atxt,['Annex A rev 2']))
                        findings.append(make_item('overcharge',{'invoice_id':iid},'cannot_determine',action=decide_action('overcharge',None,mode,False),conf=.90,evidence_list=[ev],variant='XC-02',missing_doc=md))
                        break
        # generic overcharge calculation
        if any(x.get('invoice_id')==iid and x['verdict']=='cannot_determine' for x in findings): continue
        diffs=[]; ev=[]; conflict_info=None; learned=None
        for ch in inv['charges']:
            if ch.get('type')!='freight' or not ch.get('route') or not ch.get('kg') or ch.get('rate') is None or ch.get('amount') is None: continue
            awb=ch.get('awb'); pd=tracking_pickup(track,awb,ch.get('date') or pickup)
            lane=normalize_lane(ch['route'],c); ch['route']=lane; band='0-44.5' if ch['kg']<=D('44.5') else '45-99.5' if ch['kg']<=D('99.5') else '100-299.5' if ch['kg']<=D('299.5') else '300+'
            rr=effective_rate_value(c,cid,lane,band,pd,rate_notices,feedback)
            if rr is None: continue
            if rr[0]=='conflict':
                conflict_info=(ch,rr[1],rr[2]); continue
            rate,minimum,src=rr
            correct=q2(rate*ch['kg']); correct=max(correct,minimum or correct); billed=q2(ch['amount'])
            if billed>correct:
                # fuel line on same AWB
                fuel_amt=sum(x.get('amount',ZERO) for x in inv['charges'] if x.get('type')=='fuel' and (x.get('awb')==awb or x.get('awb') is None))
                fups=[x for x in inv['charges'] if x.get('type')=='fuel' and (x.get('awb')==awb or x.get('awb') is None)]
                expected_fuel=ZERO
                fpct=c['fuel'].get(month_key(pd),None)
                if fpct is not None: expected_fuel=q2(fpct/100*correct)
                # billed fuel for this shipment is only included when line can be linked; if one fuel line, use it.
                billed_fuel=q2(fuel_amt) if fuel_amt else ZERO
                line_diff=billed-correct+max(ZERO,billed_fuel-expected_fuel)
                diffs.append(line_diff)
                e1=citation_for_line(inv,ch['line']); ev.append(e1)
                qi=source_quote_for_rate(c,lane,band); ev.append(evidence(c['doc'],qi[0],qi[1]))
                te=tracking_event(track,awb,'PICKED_UP');
                if te: ev.append(evidence(te['doc'],te['page'],te['line']))
                # feedback learning only when an accepted notice supports the rate
                acc=matching_rates(c,cid,lane,pd,rate_notices,feedback)[1]
                if acc: learned={'feedback_item_id':acc[-1]['item'].get('item_id',''),'batch':acc[-1]['feedback'].get('after_batch',next(iter(acc[-1]['feedback'].keys()),'')) or ''}
        # accessorial checks
        for ch in inv['charges']:
            typ=ch.get('type'); awb=ch.get('awb')
            if typ in ('DG','DRYICE'):
                doc=dgd.get(awb or '')
                txt=(doc or {}).get('text','').lower()
                if typ=='DG' and doc and 'dry ice' in txt and 'refrigerant' in txt and not re.search(r'\b(dangerous|hazardous|excepted quantity|class\s+\d)',txt,re.I):
                    diffs.append(ch['amount']); ev += [citation_for_line(inv,ch['line']), evidence(doc['doc'],doc['page'],choose_quote(doc['text'],['dry ice']))]
                if typ=='DRYICE' and doc and re.search(r'only as a refrigerant|refrigerant only|used solely as a refrigerant',txt):
                    # dry ice is correct, no action
                    pass
            if typ=='STOR' and awb:
                arr=tracking_event(track,awb,'ARRIVED_FACILITY'); rel=tracking_event(track,awb,'DELIVERED') or tracking_event(track,awb,'COLLECTED_BY_CONSIGNEE')
                if arr and rel:
                    expected=max(0,(rel['date']-arr['date']).days-c['free_storage'])
                    ed=q2(D(expected)*(c['accessorials'].get('STOR') or ZERO))
                    if ch.get('amount',ZERO)>ed:
                        diffs.append(ch['amount']-ed); ev.append(citation_for_line(inv,ch['line'])); ev += [evidence(arr['doc'],arr['page'],arr['line']), evidence(rel['doc'],rel['page'],rel['line'])]
        # waivers matched
        for w,cidw,cw in vwaivers:
            if w['awb'] in [ch.get('awb') for ch in inv['charges']] and w['charge']!='unknown':
                for ch in inv['charges']:
                    if ch.get('awb')==w['awb'] and ((w['charge']=='delivery fee' and ch.get('type')=='DLV') or w['charge']=='unknown'):
                        # waiver must be received before invoice and shipment named; assume if email date <= invoice date
                        if w['date']<=inv_received or (w['date']<= (ch.get('date') or inv['invoice_date']).isoformat()):
                            diffs.append(ch.get('amount',ZERO));
                            ev += [evidence(w['email'],1,w['quote']), citation_for_line(inv,ch['line'])]
                            if cw.get('contacts',{}).get('waivers'):
                                p=cw['contacts']['waivers'][0]; ev.append(evidence(cw['doc'],11,choose_quote('\n'.join(cw['pages']),[p['name'],'waivers']) or p['name']))
        # late fees on this invoice
        for ch in inv['charges']:
            if ch.get('type')=='late':
                ref=ch.get('ref'); old=apmap.get(ref)
                if old:
                    due=valid_ext.get(ref) or old['due']
                    dd=iso_date(due) if isinstance(due,str) else due
                    paid=old.get('paid');
                    if paid and dd and paid<=dd:
                        diffs.append(ch['amount']);
                        e3=next((d['doc'] for d in docs if d['doc']==extensions.get(ref,{}).get('email')) if ref in extensions else None,None)
                        if ref in extensions: ev.append(evidence(extensions[ref]['email'],1,extensions[ref]['quote']))
                        ev += [evidence(old['doc'],old['page'],old['line']),citation_for_line(inv,ch['line'])]
        if conflict_info:
            ch,rates,src=conflict_info
            low,high=D('0'),D('0')
            # calculate two candidate totals for freight+fuel for exact line
            billed_f=q2(ch['amount']); pd=tracking_pickup(track,ch.get('awb'),ch.get('date') or pickup); kg=ch['kg']
            vals=[]
            for r in rates:
                bf=max(q2(r*kg),contract_rate(c,ch['route'],kg,pd)[2] or ZERO); fp=c['fuel'].get(month_key(pd),ZERO); ff=q2(fp/100*bf); vals.append(bf+ff)
            billed_total=billed_f+q2(next((x['amount'] for x in inv['charges'] if x.get('type')=='fuel'),ZERO))
            for v in vals: high=max(high,billed_total-v); low=min([billed_total-v for v in vals]) if vals else ZERO
            low=max(ZERO,min([q2(billed_total-v) for v in vals]))
            high=max([q2(billed_total-v) for v in vals])
            ev=[evidence(src['amendment']['doc'],1,choose_quote('\n'.join(src['amendment']['pages']),[ch['route'],str(rates[0])[:4]])), evidence(src['notice']['email'],1,src['notice']['quote'])]
            low_usd, miss_low = to_usd(inv, low, c); high_usd, miss_high = to_usd(inv, high, c)
            if low_usd is None or high_usd is None:
                person=(c['contacts'].get('rates') or [{}])[0]
                md={'name':(miss_low or miss_high)['name'],'holder':{'party':person.get('party',cid.split('-')[0]),'person':person.get('name',''),'role':person.get('title','Pricing Manager')}}
                findings.append(make_item('overcharge',{'invoice_id':iid},'cannot_determine',action=decide_action('overcharge',None,mode,False),conf=.88,evidence_list=ev,variant='XC-02',missing_doc=md))
                continue
            findings.append(make_item('overcharge',{'invoice_id':iid},'conflict',action='escalate',conf=.90,evidence_list=ev,variant='XC-01',amount_range=[low_usd,high_usd]))
            continue
        if diffs:
            amt_local=q2(sum(diffs))
            amt,fx_missing=to_usd(inv, amt_local, c)
            if amt is None:
                person=(c['contacts'].get('rates') or [{}])[0]
                md={'name':fx_missing['name'],'holder':{'party':person.get('party',cid.split('-')[0]),'person':person.get('name',''),'role':person.get('title','Pricing Manager')}}
                findings.append(make_item('overcharge',{'invoice_id':iid},'cannot_determine',action=decide_action('overcharge',None,mode,False),conf=.88,evidence_list=ev[:6],variant='XC-02',missing_doc=md))
                continue
            # classify based on evidence/content
            if any('waiv' in e.get('quote','').lower() for e in ev): cls='waived_but_billed'; variant='WB-01'
            elif any('late fee' in e.get('quote','').lower() for e in ev) and not any('Air freight' in e.get('quote','') for e in ev): cls='late_fee_not_owed'; variant='LF-01'
            elif any('DG handling' in e.get('quote','') or 'Dry ice' in e.get('quote','') for e in ev): cls='overcharge'; variant='OC-02'
            else: cls='overcharge'; variant='FB-P1' if learned else 'OC-01'
            # supersedes if feedback closes the old answer: handled by omission via current accepted rate
            la=learned
            findings.append(make_item(cls,{'invoice_id':iid},'finding',amount=amt,action=decide_action(cls,amt,mode),conf=.95,evidence_list=ev[:6],variant=variant,learned_at=la))
    # duplicates: resend after paid OR same AWB+code cross-invoice
    # Build canonical lines
    line_index={}
    for inv in invoices:
        for ch in inv['charges']:
            if ch.get('type') in ('security','freight','fuel','DLV','DG','DRYICE','STOR') and ch.get('awb'):
                code=ch['type']; key=(ch['awb'].upper(),code)
                prev=line_index.get(key)
                if prev and prev['inv']['id']!=inv['id']:
                    # later invoice date is finding, unless explicitly credit/rebill semantics
                    a=prev['inv']['invoice_date'] or date.min; b=inv['invoice_date'] or date.min
                    later=inv if b>=a else prev['inv']; earlier=prev['inv'] if b>=a else inv
                    am=ch if later is inv else prev['ch']
                    q=citation_for_line(later,am['line']); q2ev=citation_for_line(earlier,am['line']);
                    findings.append(make_item('duplicate',{'invoice_id':later['id']},'finding',amount=am['amount'],action=decide_action('duplicate',am['amount'],mode),conf=.95,evidence_list=[q,q2ev],variant='DU-03'))
                else: line_index[key]={'inv':inv,'ch':ch}
    # resends/reminders after payment for exact invoice IDs
    emails= [d for d in docs if d.get('kind')=='eml']
    for inv in invoices:
        iid=inv['id']; apx=apmap.get(iid)
        if not apx or not apx.get('paid'): continue
        for e in emails:
            subj=e.get('email',{}).get('subject',''); body=e.get('email',{}).get('body','')
            if iid in (subj+' '+body).upper() and e['received']>apx['paid'].isoformat() and re.search(r'(process payment|please pay|outstanding|remit)',body,re.I):
                findings.append(make_item('duplicate',{'invoice_id':iid},'finding',amount=apx['amount'],action=decide_action('duplicate',apx['amount'],mode),conf=.95,evidence_list=[evidence(apx['doc'],apx['page'],apx['line']), evidence(e['doc'],1,choose_quote(body,[iid]))],variant='DU-01'))
                break
    # portfolio cashback and credits
    credit_notes=[]
    for d0 in docs:
        for a in d0.get('attachments',[]):
            txt='\n'.join(a.get('pages',[]))
            if a.get('kind')!='pdf' or 'credit note' not in txt.lower():
                continue
            cm=re.search(r'(?:Contract reference|Contract)\s+([A-Z0-9_/-]+)',txt,re.I)
            if not cm: continue
            cid=cm.group(1).upper()
            pm=re.search(r'period\s+(\d{4}-Q[1-4]|CY\d+)',txt,re.I)
            if not pm: continue
            period=pm.group(1).upper()
            tm=re.search(r'Total credit\s+(?:(USD|INR)\s*)?([0-9,]+\.\d{2})',txt,re.I)
            if not tm:
                ms=all_money(txt); tm_currency=re.search(r'\b(USD|INR)\b',txt,re.I)
                if not ms: continue
                cc=(tm_currency.group(1).upper() if tm_currency else 'USD')
                amount=ms[-1]
            else:
                cc=tm.group(1).upper() if tm.group(1) else 'USD'; amount=D(tm.group(2).replace(',',''))
            typm=re.search(r'Credit type\s+([a-z_ -]+)',txt,re.I)
            typet=(typm.group(1).strip().lower().replace(' ','_') if typm else '')
            if 'rebate' in typet: ctype='rebate'
            elif 'improvement' in typet or 'credit' in typet: ctype='promised_credit'
            else: ctype='promised_credit'
            dm=re.search(r'(?:Credit date|Credit note date|Date)\s+(20\d\d[-/]\d\d[-/]\d\d)',txt,re.I)
            cdate=iso_date(dm.group(1)) if dm else iso_date(d0['received'])
            credit_notes.append({'contract_id':cid,'period':period,'credit_type':ctype,'amount':amount,'currency':cc,'date':cdate,'doc':a['doc'],'page':1,'text':txt})

    def credit_usd(note,c):
        if note['currency']=='USD': return q2(note['amount'])
        fx=(c.get('fx') or {}).get(month_key(note.get('date')) if note.get('date') else '')
        return q2(note['amount']/fx) if fx else None

    for cid,c in contracts.items():
        qsp=defaultdict(lambda:ZERO)
        # Qualifying spend starts with invoice charges. Credit notes dated in the period
        # reduce spend except for rebate/improvement credits themselves.
        for inv in invoices:
            if inv['contract']!=cid or not inv['invoice_date']: continue
            qk=quarter_key(inv['invoice_date'])
            for ch in inv['charges']:
                if ch.get('type') not in ('late','tax','duty','rebate','improvement'):
                    usdv,_=to_usd(inv,ch.get('amount',ZERO),c)
                    if usdv is not None: qsp[qk]+=usdv
        for cn in credit_notes:
            if cn['contract_id']!=cid or cn['credit_type'] in ('rebate','promised_credit') or not cn.get('date'): continue
            qk=quarter_key(cn['date'])
            usd=credit_usd(cn,c)
            if usd is not None: qsp[qk]-=usd

        import calendar
        # Quarterly rebate
        for qk,sp in qsp.items():
            y=int(qk[:4]); qn=int(qk[-1]); qm=qn*3
            qend=date(y,qm,calendar.monthrange(y,qm)[1])
            if qend.isoformat()>asof or not c.get('rebates'):
                continue
            rate=max((r for th,r in c['rebates'] if sp>=th),default=None)
            if rate is None: continue
            owed=q2(max(ZERO,sp*rate/100))
            issued=sum((credit_usd(n,c) or ZERO) for n in credit_notes if n['contract_id']==cid and n['period']==qk and n['credit_type']=='rebate')
            net=q2(max(ZERO,owed-issued))
            if net<=ZERO: continue
            ev=[]
            rp=next((i for i in invoices if i['contract']==cid and quarter_key(i['invoice_date'])==qk),None)
            if rp:
                page=next((p for p in c['pages'] if 'Volume rebate' in p),c['pages'][0]); pi=c['pages'].index(page)+1
                ev.append(evidence(c['doc'],pi,choose_quote(page,['Volume rebate'])))
                qpage=next((p for p in c['pages'] if 'Qualifying Spend' in p),c['pages'][0]); qpi=c['pages'].index(qpage)+1
                ev.append(evidence(c['doc'],qpi,choose_quote(qpage,['Qualifying Spend'])))
            due=qend+timedelta(days=45)
            action='internal:track_credit' if date.fromisoformat(asof)<due else decide_action('cashback',net,mode)
            findings.append(make_item('cashback',{'contract_id':cid,'period':qk,'credit_type':'rebate'},'finding',amount=net,action=action,conf=.95,evidence_list=ev,variant='CB-01'))

        # Annual improvement/other promised credit
        if c.get('annual') is not None and c.get('effective'):
            eff=iso_date(c['effective'])
            n=1
            while True:
                start=add_months(eff,12*(n-1)); anniversary=add_months(eff,12*n); end=anniversary-timedelta(days=1)
                if end.isoformat()>asof: break
                sp=ZERO
                for inv in invoices:
                    if inv['contract']==cid and inv.get('invoice_date') and start<=inv['invoice_date']<=end:
                        for ch in inv['charges']:
                            if ch.get('type') not in ('late','tax','duty','rebate','improvement'):
                                usd,_=to_usd(inv,ch.get('amount',ZERO),c)
                                if usd is not None: sp+=usd
                for cn in credit_notes:
                    if cn['contract_id']==cid and cn['credit_type'] not in ('rebate','promised_credit') and cn.get('date') and start<=cn['date']<=end:
                        usd=credit_usd(cn,c)
                        if usd is not None: sp-=usd
                owed=q2(max(ZERO,sp*c['annual']/100))
                period=f'CY{n}'
                issued=sum((credit_usd(cn,c) or ZERO) for cn in credit_notes if cn['contract_id']==cid and cn['period']==period and cn['credit_type']=='promised_credit')
                net=q2(max(ZERO,owed-issued))
                if net>ZERO:
                    page=next((p for p in c['pages'] if 'Annual improvement credit' in p),c['pages'][0]); pi=c['pages'].index(page)+1
                    ev=[evidence(c['doc'],pi,choose_quote(page,['Annual improvement credit']))]
                    due=anniversary+timedelta(days=60)
                    action='internal:track_credit' if date.fromisoformat(asof)<due else decide_action('cashback',net,mode)
                    findings.append(make_item('cashback',{'contract_id':cid,'period':period,'credit_type':'promised_credit'},'finding',amount=net,action=action,conf=.95,evidence_list=ev,variant='CB-03'))
                n+=1

    # late fee risk
    for iid,x in apmap.items():
        if x['status']!='open' or not x['due']: continue
        dd=x['due']
        if isinstance(dd,str): dd=iso_date(dd)
        delta=(dd-date.fromisoformat(asof)).days
        if 1<=delta<=14:
            c=contracts.get(next((i['contract'] for i in invoices if i['id']==iid),None))
            if not c or not c.get('late_rate'): continue
            fee=q2(x['amount']*c['late_rate']/100)
            ev=[evidence(x['doc'],x['page'],x['line'])]
            page=next((p for p in c['pages'] if 'Late fee' in p),c['pages'][0]); pi=c['pages'].index(page)+1; ev.append(evidence(c['doc'],pi,choose_quote(page,['Late fee'])))
            findings.append(make_item('late_fee_risk',{'invoice_id':iid},'finding',amount=fee,action='internal:payment_reminder',conf=.95,evidence_list=ev,variant='LF-03',deadline=dd.isoformat()))
    # expiry
    for cid,c in contracts.items():
        end=iso_date(c.get('expiry') or '')
        if not end: continue
        days=(end-date.fromisoformat(asof)).days
        if 0<=days<=60:
            page=next((p for p in c['pages'] if 'Initial term' in p or 'expires' in p),c['pages'][0]); pi=c['pages'].index(page)+1
            ev=[evidence(c['doc'],pi,choose_quote(page,['Initial term'] if 'Initial term' in page else ['expires']))]
            findings.append(make_item('expiry',{'contract_id':cid},'finding',action='internal:alert',conf=.95,evidence_list=ev,variant='EX-01',deadline=end.isoformat()))
    # dedupe same finding key, prefer richer item
    bykey={}
    for f in findings:
        k=(f.get('class'),f.get('invoice_id'),f.get('contract_id'),f.get('period'),f.get('credit_type'))
        if k not in bykey or len(f.get('evidence',[]))>len(bykey[k].get('evidence',[])): bykey[k]=f
    findings=list(bykey.values())

    # Apply exact-key feedback only after the current evidence pass. Standing facts were
    # already used above; these exact decisions stop, override or annotate one finding.
    filtered=[]
    for f in findings:
        key={}
        for k in ('invoice_id','class','contract_id','period','credit_type'):
            if f.get(k) is not None: key[k]=f.get(k)
        matches=feedback_for_key(fitems,key)
        drop=False
        for item in matches:
            dec=item.get('decision')
            fact=item.get('fact') or {}
            if dec in ('close','reject'):
                drop=True; break
            if dec=='override_amount':
                value=fact.get('value') or {}
                amt=value.get('amount_usd')
                if amt is not None:
                    f['amount_usd']=money(D(str(amt)))
                    if f.get('class') not in ('expiry','late_fee_risk'):
                        f['action']=decide_action(f.get('class'),D(str(amt)),mode)
                f['learned_at']={'feedback_item_id':item.get('item_id',''),'batch':(item.get('_feedback') or {}).get('after_batch','')}
            elif dec in ('one_off_approval','resolve_conflict','confirm'):
                f['learned_at']={'feedback_item_id':item.get('item_id',''),'batch':(item.get('_feedback') or {}).get('after_batch','')}
        if not drop: filtered.append(f)
    findings=filtered

    # Feedback can close a finding before a later report; do not resurrect it merely because
    # the same stale document remains on disk.
    # adaptation supersedes and feedback learned_at based on previous report
    prev=load_json(state_dir/'last_report.json',None)
    if prev:
        pm={ (x.get('class'),x.get('invoice_id'),x.get('contract_id'),x.get('period'),x.get('credit_type')):x for x in (prev.get('findings',[])+sum([(prev.get('portfolio',{}).get(k) or []) for k in ('cashback','expiry','late_fee_risk')],[])) }
        for f in findings:
            k=(f.get('class'),f.get('invoice_id'),f.get('contract_id'),f.get('period'),f.get('credit_type')); old=pm.get(k)
            if old and old.get('verdict')!=f.get('verdict'):
                f['supersedes']=[old.get('id')]
    portfolio={'cashback':[f for f in findings if f['class']=='cashback'],'expiry':[f for f in findings if f['class']=='expiry'],'late_fee_risk':[f for f in findings if f['class']=='late_fee_risk']}
    report_obj={'schema_version':'1.0','as_of':asof,'mode':mode,'findings':[f for f in findings if f['class'] in ('overcharge','duplicate','waived_but_billed','late_fee_not_owed')], 'portfolio':portfolio}
    save_json(state_dir/'last_report.json',report_obj)
    print(json.dumps(report_obj,ensure_ascii=False))


def feedback(path,state_dir):
    obj=json.loads(Path(path).read_text())
    fdir=state_dir/'feedback'; fdir.mkdir(parents=True,exist_ok=True)
    out=fdir/(Path(path).stem+'.json'); save_json(out,obj)
    print(json.dumps({'ok':True,'stored':out.name}))


def main():
    ap=argparse.ArgumentParser()
    sub=ap.add_subparsers(dest='cmd',required=True)
    a=sub.add_parser('ingest'); a.add_argument('batch_dir'); a.add_argument('--state-dir',required=True)
    a=sub.add_parser('report'); a.add_argument('--as-of',required=True); a.add_argument('--mode',choices=['supervised','autonomous'],required=True); a.add_argument('--state-dir',required=True)
    a=sub.add_parser('feedback'); a.add_argument('feedback_json'); a.add_argument('--state-dir',required=True)
    x=ap.parse_args()
    sd=Path(x.state_dir)
    try:
        if x.cmd=='ingest': ingest(x.batch_dir,sd)
        elif x.cmd=='report': report(x.as_of,x.mode,sd)
        else: feedback(x.feedback_json,sd)
    except Exception as e:
        print(f'harness error: {type(e).__name__}: {e}',file=sys.stderr)
        raise
if __name__=='__main__': main()
