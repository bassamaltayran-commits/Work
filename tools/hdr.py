import sys,zipfile,shutil
f,new=sys.argv[1],sys.argv[2]
old="Assignment 3: Strategic Brand Management – The Ritz-Carlton"
zin=zipfile.ZipFile(f); tmp=f+'.tmp'; zout=zipfile.ZipFile(tmp,'w',zipfile.ZIP_DEFLATED); n=0
for it in zin.infolist():
    if it.filename.endswith('/'): continue
    data=zin.read(it.filename)
    if it.filename.startswith('word/header'):
        s=data.decode(); n+=s.count(old); s=s.replace(old,new); data=s.encode()
    if it.filename=='docProps/core.xml':
        import re; s=data.decode(); s=re.sub(r'<dc:title>.*?</dc:title>','',s); data=s.encode()
    zout.writestr(it,data)
zout.close(); shutil.move(tmp,f); print(f,n)
