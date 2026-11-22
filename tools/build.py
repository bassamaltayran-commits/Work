"""Build an assignment in the UPM template; move any numbered section that would split across pages to a new page."""
import sys,os,json,subprocess,glob,tempfile
TPL,content,out,header,course,assign,date,short=sys.argv[1:9]
here=os.path.dirname(os.path.abspath(__file__))
heads=[b['text'] for b in json.load(open(content)) if b['type']=='h1']
soffice=glob.glob('/root/.claude/skills/synced/*/docx/scripts/office/soffice.py')[0]
tmp=tempfile.mkdtemp(); brk=[]
def norm(x): return ' '.join(x.split())
for it in range(8):
    env=dict(os.environ,UPM_BREAKS='|'.join(brk))
    subprocess.run([sys.executable,f'{here}/upm.py',TPL,content,out,header,course,assign,date],env=env,check=True,capture_output=True)
    subprocess.run([sys.executable,f'{here}/hdr.py',out,short],check=True,capture_output=True)
    subprocess.run([sys.executable,soffice,'--headless','--convert-to','pdf','--outdir',tmp,out],capture_output=True)
    pdf=os.path.join(tmp,os.path.basename(out)[:-5]+'.pdf')
    pages=subprocess.run(['pdftotext','-layout',pdf,'-'],capture_output=True,text=True).stdout.split('\f')
    pages=[[norm(l) for l in pg.splitlines() if norm(l) and not any(k in l for k in ('Kingdom of Saudi','Prince Mugrin Bin','Al-Madinah','Academic Year 1448',short[:25]))] for pg in pages]
    def page_of(h):
        for i,pg in enumerate(pages):
            if any(l==h or l.startswith(h) for l in pg): return i
    new=None
    for k,h in enumerate(heads):
        if h in brk: continue
        p1=page_of(h); p2=page_of(heads[k+1]) if k+1<len(heads) else len([x for x in pages if x])-1
        if p1 is None or p2 is None: continue
        nxt=heads[k+1] if k+1<len(heads) else None
        spills = p2>p1 and pages[p1+1] and not (nxt and pages[p1+1][0].startswith(nxt))
        if spills: new=h; break
    if not new: break
    brk.append(new)
print(os.path.basename(out),'breaks before:',brk,'pages:',len([x for x in pages if x]))
