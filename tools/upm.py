import sys,json,copy,zipfile,shutil,os
import docx
from docx.shared import Pt,Emu,Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH as A
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
TPL,content,out,header,course,assign,date=sys.argv[1:8]
d=docx.Document(TPL)
body=d.element.body
paras=d.paragraphs
# keep cover (paragraphs 0..11)
keep=set(id(p._p) for p in paras[:12])
for el in list(body):
    if el.tag==qn('w:sectPr') or id(el) in keep: continue
    body.remove(el)
sect=body.find(qn('w:sectPr'))
def setfield(p,label,val):
    r=p.runs; r[0].text=label
    if val is None: 
        for x in r[1:]: x._r.getparent().remove(x._r)
        return
    r[1].text=val
    for x in r[2:]: x._r.getparent().remove(x._r)
setfield(paras[0],"University of Prince Mugrin - College of Business and Tourism",None)
setfield(paras[1],"International Hospitality Management",None)
setfield(paras[2],"Course: ",course)
setfield(paras[3],"Assignment No: ",assign)
setfield(paras[4],"Student’s Name: ","Bassam K. Altayran")
setfield(paras[5],"Student’s ID No.: ","4411839")
ins=copy.deepcopy(paras[5]._p); paras[5]._p.addnext(ins)
ip=docx.text.paragraph.Paragraph(ins,paras[5]._parent); setfield(ip,"Instructor: ","Dr. Malik Mohammed")
setfield(paras[6],"Date of submission: ",date)

def fmt(p,align=A.JUSTIFY,after=8,line=1.15,left=None,first=None,before=0,keep=False):
    pf=p.paragraph_format; p.alignment=align; pf.space_after=Pt(after); pf.space_before=Pt(before); pf.line_spacing=line
    if left is not None: pf.left_indent=Emu(left)
    if first is not None: pf.first_line_indent=Emu(first)
    if keep: pf.keep_with_next=True
def addrun(p,text,bold=False,italics=False,size=12):
    r=p.add_run(text); r.bold=bold; r.italic=italics; r.font.name="Times New Roman"; r.font.size=Pt(size)
    r._r.get_or_add_rPr().get_or_add_rFonts().set(qn('w:eastAsia'),"Times New Roman"); return r
def newp():
    p=d.add_paragraph(); body.remove(p._p); sect.addprevious(p._p); return p
def borders(tbl):
    tblPr=tbl._tbl.tblPr; b=OxmlElement('w:tblBorders')
    for e in ['top','left','bottom','right','insideH','insideV']:
        x=OxmlElement('w:'+e); x.set(qn('w:val'),'single'); x.set(qn('w:sz'),'4'); x.set(qn('w:color'),'1AA7A7'); b.append(x)
    tblPr.append(b)
def shade(cell,fill):
    tcPr=cell._tc.get_or_add_tcPr(); s=OxmlElement('w:shd'); s.set(qn('w:val'),'clear'); s.set(qn('w:color'),'auto'); s.set(qn('w:fill'),fill); tcPr.append(s)
TW=9566
sect.find(qn('w:pgMar')).set(qn('w:top'),'2950')
p=newp(); fmt(p,A.CENTER,12,keep=True); addrun(p,header.split(': ',1)[1] if ': ' in header else header,True,size=14)
# page break before title so content starts on page 2
p.paragraph_format.page_break_before=True
blocks=json.load(open(content))
for i,bk in enumerate(blocks):
    t=bk['type']
    if t=='h1':
        p=newp(); fmt(p,A.CENTER,6,before=6,keep=True); addrun(p,bk['text'],True)
        if i>0 and blocks[i-1]['type']=='pagebreak': p.paragraph_format.page_break_before=True
    elif t=='h2':
        p=newp(); fmt(p,A.LEFT,6,before=4,keep=True); addrun(p,bk['text'],True,True)
    elif t in('p','bullet','ref'):
        p=newp()
        if t=='p': fmt(p)
        elif t=='bullet': fmt(p,after=4,left=265430,first=-164465); addrun(p,"• ")
        else: fmt(p,A.LEFT,8,left=457200,first=-457200)
        for r in bk['runs']: addrun(p,r['text'],r.get('bold',False),r.get('italics',False))
    elif t=='sp':
        continue
    elif t=='pagebreak':
        continue
    elif t=='table':
        rows=bk['rows']; n=len(rows[0]); w=bk['widths']; s=sum(w)
        tb=d.add_table(rows=len(rows),cols=n); body.remove(tb._tbl); sect.addprevious(tb._tbl)
        borders(tb); tb.autofit=False
        for ri,row in enumerate(rows):
            for ci,val in enumerate(row):
                cell=tb.cell(ri,ci); cell.width=Twips(int(w[ci]*TW/s))
                cp=cell.paragraphs[0]; fmt(cp,A.RIGHT if (bk['right'] and ci>0 and ri>0) else A.LEFT,0,1.0)
                if ri<len(rows)-1: cp.paragraph_format.keep_with_next=True
                hb= ri==0 or ri in bk['bold']
                r=addrun(cp,str(val),hb)
                if ri==0: shade(cell,'00507F'); r.font.color.rgb=docx.shared.RGBColor(0xFF,0xFF,0xFF)
                elif ri in bk['bold']: shade(cell,'FBEFD5')
                elif ri%2==0: shade(cell,'EAF6F6')
        cw=[int(x*TW/s) for x in w]
        grid=tb._tbl.tblGrid
        for gc,wv in zip(grid.findall(qn('w:gridCol')),cw): gc.set(qn('w:w'),str(wv))
        tblPr=tb._tbl.tblPr
        tw=OxmlElement('w:tblW'); tw.set(qn('w:w'),str(sum(cw))); tw.set(qn('w:type'),'dxa')
        for old in tblPr.findall(qn('w:tblW')): tblPr.remove(old)
        tblPr.append(tw)
        lay=OxmlElement('w:tblLayout'); lay.set(qn('w:type'),'fixed'); tblPr.append(lay)
        for row in tb.rows:
            trPr=row._tr.get_or_add_trPr(); cs=OxmlElement('w:cantSplit'); trPr.append(cs)
        p=newp(); fmt(p,after=6)
    if globals().get('pb') and t!='pagebreak':
        pass
# handle page breaks: set page_break_before on heading following a pagebreak block
d.save(out)
