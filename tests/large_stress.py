from pathlib import Path
import json,time,sys,subprocess,sqlite3,tempfile
from reportlab.pdfgen import canvas
root=Path(tempfile.mkdtemp(prefix='dte-4000-page-check-'))
f=root/'source.pdf'
t=time.monotonic();c=canvas.Canvas(str(f),pageCompression=0)
for n in range(4000):
    c.drawString(30,800,f'Shop {n+1} Amount 123.45')
    c.drawString(30,760,'Detailed source text. '*320)
    c.showPage()
c.save()
(root/'job.json').write_text(json.dumps({'columns':['Shop','Amount'],'name':'synthetic-4000.pdf','requirements':'extract shops','pages':4000}))
script=str(Path(__file__).resolve().parents[1]/'scripts/large_pdf.py')
def run(command,limit=4000):
    start=time.monotonic();r=subprocess.run([sys.executable,script,command,'--dir',str(root),'--limit',str(limit)],capture_output=True,text=True,check=True)
    value=json.loads(r.stdout);value.pop('preview',None);print(json.dumps({'step':command,'seconds':round(time.monotonic()-start,2),**value}),flush=True)
print(json.dumps({'bytes':f.stat().st_size,'pages':4000}),flush=True)
run('extract');run('export');run('extract')
from openpyxl import load_workbook
w=load_workbook(root/'result.xlsx',read_only=True)
rows=sum(1 for _ in w['原文'].rows)-1
pages=sqlite3.connect(root/'pages.sqlite').execute('select count(*) from pages').fetchone()[0]
assert pages==4000 and rows==16000,(pages,rows)
print(json.dumps({'verifiedPages':pages,'verifiedRawRows':rows,'totalSeconds':round(time.monotonic()-t,2)}),flush=True)
