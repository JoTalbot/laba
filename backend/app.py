import shutil, subprocess, tempfile, re
from pathlib import Path
import fitz
from docx import Document
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
app=FastAPI(title='111ЛАБ API',version='1.0.0')
app.add_middleware(CORSMiddleware,allow_origins=['*'],allow_methods=['*'],allow_headers=['*'])
BASE=Path(__file__).resolve().parent
TEMPLATE=BASE/'template'/'шаблонdocx.docx'
def extract_pdf(data):
    with fitz.open(stream=data,filetype='pdf') as doc:return '\n'.join(p.get_text() for p in doc)
def clean(t):return re.sub(r'\s+',' ',t).strip()
def parse_lab(text):
    text=clean(text)
    def m(p,d=''):
        x=re.search(p,text,re.I);return x.group(1).strip() if x else d
    return {'lab_number':m(r'Лабораторна\s+робота\s+№\s*(\d+)','1'),'topic':m(r'Тема:\s*(.*?)(?:\s+Мета:|$)','Не визначено'),'goal':m(r'Мета:\s*(.*?)(?:\s+Теоретичні відомості|$)','Не визначено'),'raw':text}
def replace_all(doc,mapping):
    for p in doc.paragraphs:
        for old,new in mapping.items():
            if old in p.text:
                for r in p.runs:r.text=r.text.replace(old,new)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for old,new in mapping.items():
                        if old in p.text:
                            for r in p.runs:r.text=r.text.replace(old,new)
def fill_template(parsed,student,teacher,discipline,conclusion,out):
    doc=Document(TEMPLATE)
    replace_all(doc,{'???':discipline,'ЛАБОРАТОРНОЇ РОБОТИ № 1':f"ЛАБОРАТОРНОЇ РОБОТИ № {parsed['lab_number']}",'«???»':f"«{parsed['topic']}»"})
    doc.add_paragraph(f"Мета: {parsed['goal']}")
    doc.add_paragraph('Завдання та рішення')
    doc.add_paragraph('Рішення завдань формуються solver-модулем після анализа текста лабораторної роботи.')
    doc.add_paragraph(f'Виконав: {student}')
    doc.add_paragraph(f'Перевірив: {teacher}')
    doc.add_paragraph('Висновок')
    doc.add_paragraph(conclusion)
    doc.save(out)
def render(docx,pdf):
    soffice=shutil.which('libreoffice') or shutil.which('soffice')
    if not soffice:raise RuntimeError('LibreOffice не установлен на backend-сервере')
    subprocess.run([soffice,'--headless','--convert-to','pdf','--outdir',str(pdf.parent),str(docx)],check=True,stdout=subprocess.PIPE,stderr=subprocess.PIPE)
    (pdf.parent/(docx.stem+'.pdf')).replace(pdf)
@app.get('/health')
def health():return {'status':'ok','template':TEMPLATE.exists(),'renderer':bool(shutil.which('libreoffice') or shutil.which('soffice'))}
@app.post('/api/inspect')
async def inspect(file:UploadFile=File(...)):
    if file.content_type!='application/pdf':raise HTTPException(400,'Нужен PDF')
    return parse_lab(extract_pdf(await file.read()))
@app.post('/api/generate')
async def generate(file:UploadFile=File(...),student:str=Form(...),teacher:str=Form(...),discipline:str=Form(...),conclusion:str=Form(...)):
    if file.content_type!='application/pdf':raise HTTPException(400,'Нужен PDF')
    parsed=parse_lab(extract_pdf(await file.read()));work=Path(tempfile.mkdtemp(prefix='laba-'));docx=work/'111ЛАБ.docx';pdf=work/'111ЛАБ.pdf'
    try:
        fill_template(parsed,student,teacher,discipline,conclusion,docx);render(docx,pdf)
        return FileResponse(pdf,media_type='application/pdf',filename='111ЛАБ.pdf')
    except Exception as e:raise HTTPException(500,str(e))
