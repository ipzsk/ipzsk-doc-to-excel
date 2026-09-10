"""Disk-backed PDF processing. Each page commits independently; no model calls here."""
import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import sqlite3
import sys
from doc_to_excel import local_ocr


def connect(directory, readonly=False):
    db = sqlite3.connect(Path(directory) / 'pages.sqlite', timeout=30)
    db.row_factory = sqlite3.Row
    if readonly:
        db.execute('PRAGMA query_only=ON')
        return db
    db.execute('PRAGMA journal_mode=WAL')
    db.executescript('''
    CREATE TABLE IF NOT EXISTS pages (page INTEGER PRIMARY KEY, text TEXT, warning TEXT, failed INTEGER);
    CREATE TABLE IF NOT EXISTS records (id TEXT PRIMARY KEY, page INTEGER, part INTEGER, text TEXT);
    CREATE TABLE IF NOT EXISTS decisions (id TEXT PRIMARY KEY, data TEXT);
    CREATE INDEX IF NOT EXISTS record_page ON records(page,part);
    ''')
    return db


def reader(directory):
    from pypdf import PdfReader
    pdf = PdfReader(Path(directory) / 'source.pdf')
    if pdf.is_encrypted and not pdf.decrypt(''):
        raise ValueError('PDF 已加密，请先提供解密副本')
    return pdf


def stats(db):
    return dict(extracted=db.execute('SELECT COUNT(*) FROM pages').fetchone()[0],
                failed=db.execute('SELECT COUNT(*) FROM pages WHERE failed=1').fetchone()[0],
                total=db.execute('SELECT COUNT(*) FROM records').fetchone()[0],
                pending=db.execute('SELECT COUNT(*) FROM records r LEFT JOIN decisions d USING(id) WHERE d.id IS NULL').fetchone()[0])


def extract(db, directory, limit):
    pdf = reader(directory)
    count = len(pdf.pages)
    done = {r[0] for r in db.execute('SELECT page FROM pages')}
    for page in range(1, min(limit, count) + 1):
        if page in done:
            continue
        warning = ''; failed = 0
        try:
            text = (pdf.pages[page-1].extract_text() or '').strip()
            if not text:
                text, warning = local_ocr(path=Path(directory) / 'source.pdf', page=page)
                failed = int(not text)
        except Exception as error:
            text = ''; warning = '页面提取失败：' + str(error)[:200]; failed = 1
        with db:
            db.execute('INSERT INTO pages VALUES(?,?,?,?)', (page, text, warning, failed))
            for part, offset in enumerate(range(0, len(text), 1800)):
                rid = f'p{page}-s{part+1}'
                db.execute('INSERT INTO records VALUES(?,?,?,?)', (rid, page, part+1, text[offset:offset+1800]))
    return {**stats(db), 'pages': count}


def batch(db, limit):
    records = []
    for r in db.execute('SELECT r.* FROM records r LEFT JOIN decisions d USING(id) WHERE d.id IS NULL AND page<=? ORDER BY page,part LIMIT 2', (limit,)):
        previous = db.execute('SELECT text FROM records WHERE page<? OR (page=? AND part<?) ORDER BY page DESC,part DESC LIMIT 1', (r['page'],r['page'],r['part'])).fetchone()
        following = db.execute('SELECT text FROM records WHERE page>? OR (page=? AND part>?) ORDER BY page,part LIMIT 1', (r['page'],r['page'],r['part'])).fetchone()
        records.append({'id':r['id'], 'page':r['page'], 'excerpt':r['text'],
                        'previousContext':previous[0][-250:] if previous else '',
                        'nextContext':following[0][:250] if following else ''})
    return {'records': records}


def apply(db, directory):
    items = json.loads((Path(directory) / 'accepted.json').read_text())
    with db:
        for item in items:
            if not db.execute('SELECT 1 FROM records WHERE id=?',(item['id'],)).fetchone():
                raise ValueError('未知记录')
            db.execute('INSERT OR REPLACE INTO decisions VALUES(?,?)', (item['id'],json.dumps(item,ensure_ascii=False)))
    return stats(db)


def preview(db):
    result = []
    for r in db.execute('SELECT r.*, d.data, p.warning FROM records r LEFT JOIN decisions d USING(id) JOIN pages p USING(page) ORDER BY page,part LIMIT 50'):
        d = json.loads(r['data']) if r['data'] else {}
        result.append({'id':r['id'], 'page':r['page'], 'text':r['text'][:500], 'decision':d,'warning':r['warning']})
    return {**stats(db), 'preview': result}


def export(db, directory):
    from openpyxl import Workbook
    from openpyxl.cell import WriteOnlyCell
    from openpyxl.styles import Font
    config = json.loads((Path(directory)/'job.json').read_text())
    columns = config['columns'] or ['整理内容']
    book = Workbook(write_only=True)
    overview = book.create_sheet('总览')
    sheets = {name:book.create_sheet(name) for name in ['提取结果','待核对','原文','页面异常']}
    sheet_rows={}
    def append(sheet, values):
        sheet_rows[sheet.title]=sheet_rows.get(sheet.title,0)+1
        if sheet_rows[sheet.title]>1048576:
            raise ValueError('结果超出单个 Excel 工作表行数限制，原文及进度仍在本地数据库中')
        cells=[]
        for value in values:
            if isinstance(value, str):
                value = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', value)
            cell=WriteOnlyCell(sheet,value=value)
            if isinstance(value,str): cell.data_type='s'
            cells.append(cell)
        sheet.append(cells)
    headers=['记录ID',*columns,'来源文件','页码','原文片段ID','核对说明']
    for name in ['提取结果','待核对']:
        sheets[name].freeze_panes='A2'; append(sheets[name],headers)
    append(sheets['原文'],['片段ID','页码','原始文字','识别备注','处理状态'])
    append(sheets['页面异常'],['页码','说明'])
    output_rows=0; review=0
    for r in db.execute('SELECT r.*, d.data,p.warning FROM records r LEFT JOIN decisions d USING(id) JOIN pages p USING(page) ORDER BY page,part'):
        d=json.loads(r['data']) if r['data'] else None
        append(sheets['原文'],[r['id'],r['page'],r['text'],r['warning'],'已整理' if d else '未整理'])
        if d is None:
            append(sheets['待核对'],[r['id'],*['' for _ in columns],config['name'],r['page'],r['id'],'未整理，见原文']); review+=1
            continue
        for n, item in enumerate(d['rows'],1):
            reason=item['review']
            if r['warning']: reason=(reason+'；'+r['warning']).strip('；')
            row=[r['id']+f'-r{n}',*[item['values'][c] for c in columns],config['name'],r['page'],r['id'],reason]
            append(sheets['提取结果'],row); output_rows+=1
            if reason: append(sheets['待核对'],row); review+=1
        if not d['rows']:
            append(sheets['待核对'],[r['id'],*['' for _ in columns],config['name'],r['page'],r['id'],'模型未提取到匹配记录：'+d['note']]); review+=1
    for p in db.execute('SELECT page,warning FROM pages WHERE warning<>"" ORDER BY page'):
        append(sheets['页面异常'],list(p))
    s=stats(db)
    for row in [['项目','值'],['源文件',config['name']],['总页数',config.get('pages',0)],['已解析页数',s['extracted']],['解析失败页数',s['failed']],['未整理片段',s['pending']],['提取记录',output_rows],['待核对记录',review],['要求',config['requirements']],['说明','页码为 PDF 物理页码；跨页/跨片段记录需人工核对合并。未整理和无匹配片段均保留。所有字段为文本，未执行数据库入库。']]:
        append(overview,row)
    temporary=Path(directory)/'export.tmp.xlsx'
    book.save(temporary)
    os.replace(temporary,Path(directory)/'result.xlsx')
    return {'rows':output_rows,'review':review}


def main():
    p=argparse.ArgumentParser(); p.add_argument('command',choices=['info','extract','batch','apply','status','export','retry']);p.add_argument('--dir',required=True);p.add_argument('--limit',type=int,default=10)
    a=p.parse_args(); db=connect(a.dir, readonly=a.command in ('status','batch','export'))
    try:
        if a.command=='info': result={'pages':len(reader(a.dir).pages)}
        elif a.command=='extract': result=extract(db,a.dir,a.limit)
        elif a.command=='batch': result=batch(db,a.limit)
        elif a.command=='apply': result=apply(db,a.dir)
        elif a.command=='status': result=preview(db)
        elif a.command=='export': result=export(db,a.dir)
        else:
            with db: db.execute('DELETE FROM pages WHERE failed=1')
            result=stats(db)
        print(json.dumps(result,ensure_ascii=False))
    finally: db.close()

if __name__=='__main__': main()
