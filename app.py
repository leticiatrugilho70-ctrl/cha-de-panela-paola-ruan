from flask import Flask, render_template, jsonify, request, send_from_directory, redirect, url_for, send_file
import sqlite3, json, os, uuid
from pathlib import Path
from datetime import datetime
from werkzeug.utils import secure_filename
from openpyxl import Workbook

BASE=Path(__file__).resolve().parent
DB=BASE/'presentes.db'
IMG=BASE/'images'
EXPORT=BASE/'escolhas.xlsx'
ALLOWED={'.png','.jpg','.jpeg','.webp','.gif'}
app=Flask(__name__)
app.config['MAX_CONTENT_LENGTH']=8*1024*1024

def conn():
    c=sqlite3.connect(DB,timeout=10); c.row_factory=sqlite3.Row; return c

def export_choices(c=None):
    own=c is None
    if own:c=conn()
    rows=c.execute('''SELECT h.id,h.guest_name,h.gift_name,h.chosen_at,h.active
                      FROM choice_history h ORDER BY h.id DESC''').fetchall()
    wb=Workbook(); ws=wb.active; ws.title='Escolhas'
    ws.append(['Nome do convidado','Presente escolhido','Data e hora','Status'])
    for r in rows:
        ws.append([r['guest_name'],r['gift_name'],r['chosen_at'],'Ativa' if r['active'] else 'Liberada'])
    ws.column_dimensions['A'].width=28; ws.column_dimensions['B'].width=42; ws.column_dimensions['C'].width=22; ws.column_dimensions['D'].width=14
    wb.save(EXPORT)
    if own:c.close()

def init_db():
    c=conn()
    c.execute('''CREATE TABLE IF NOT EXISTS gifts(id INTEGER PRIMARY KEY AUTOINCREMENT,name TEXT NOT NULL,link TEXT,image TEXT,status TEXT NOT NULL DEFAULT 'disponivel',chosen_by TEXT,chosen_at TEXT)''')
    c.execute('CREATE TABLE IF NOT EXISTS settings(key TEXT PRIMARY KEY,value TEXT)')
    c.execute('''CREATE TABLE IF NOT EXISTS choice_history(id INTEGER PRIMARY KEY AUTOINCREMENT,gift_id INTEGER,guest_name TEXT NOT NULL,gift_name TEXT NOT NULL,chosen_at TEXT NOT NULL,active INTEGER NOT NULL DEFAULT 1)''')
    if c.execute('SELECT COUNT(*) n FROM gifts').fetchone()['n']==0:
        data=json.loads((BASE/'items.json').read_text(encoding='utf-8'))
        c.executemany('INSERT INTO gifts(id,name,link,image) VALUES(:id,:name,:link,:image)',data)
    defaults={'title':'Lista de Presentes','subtitle':'Chá de Panela','couple':'Paola & Ruan','intro':'Cada presente escolhido com amor levará um pedacinho de vocês para o nosso novo lar, fazendo parte dos sonhos e da linda história que estamos construindo juntos.','primary':'#82a8c5','secondary':'#d8edf9','background':'#fbfaf5'}
    for k,v in defaults.items():c.execute('INSERT OR IGNORE INTO settings VALUES(?,?)',(k,v))
    c.commit(); export_choices(c); c.close()

@app.get('/images/<path:name>')
def images(name):return send_from_directory(IMG,name)

@app.get('/')
def home():
    c=conn(); gifts=[dict(r) for r in c.execute('SELECT id,name,link,image,status FROM gifts ORDER BY id')]; settings={r['key']:r['value'] for r in c.execute('SELECT key,value FROM settings')}; c.close()
    return render_template('index.html',gifts=gifts,settings=settings)

@app.get('/api/gifts')
def gifts_api():
    c=conn(); data=[dict(r) for r in c.execute('SELECT id,name,link,image,status FROM gifts ORDER BY id')]; c.close(); return jsonify(data)

@app.post('/api/reserve/<int:gift_id>')
def reserve(gift_id):
    name=(request.json or {}).get('name','').strip()
    if not name:return jsonify(ok=False,message='Digite seu nome.'),400
    c=conn()
    try:
        c.execute('BEGIN IMMEDIATE'); row=c.execute('SELECT name,status FROM gifts WHERE id=?',(gift_id,)).fetchone()
        if not row:c.rollback(); return jsonify(ok=False,message='Presente não encontrado.'),404
        if row['status']!='disponivel':c.rollback(); return jsonify(ok=False,message='Esse presente acabou de ser escolhido por outra pessoa. Escolha outro presente.'),409
        now=datetime.now().strftime('%d/%m/%Y %H:%M:%S')
        cur=c.execute("UPDATE gifts SET status='escolhido',chosen_by=?,chosen_at=? WHERE id=? AND status='disponivel'",(name,now,gift_id))
        if cur.rowcount!=1:c.rollback(); return jsonify(ok=False,message='Esse presente acabou de ser escolhido por outra pessoa. Escolha outro presente.'),409
        c.execute('INSERT INTO choice_history(gift_id,guest_name,gift_name,chosen_at) VALUES(?,?,?,?)',(gift_id,name,row['name'],now)); c.commit(); export_choices(c)
        return jsonify(ok=True,message='Presente escolhido com carinho! Obrigada por fazer parte desse momento tão especial.')
    finally:c.close()

def save_upload(file):
    if not file or not file.filename:return None
    ext=Path(secure_filename(file.filename)).suffix.lower()
    if ext not in ALLOWED:raise ValueError('Formato de imagem não permitido.')
    name=f'gift-{uuid.uuid4().hex}{ext}'; file.save(IMG/name); return f'images/{name}'

@app.route('/admin',methods=['GET','POST'])
def admin():
    c=conn()
    if request.method=='POST':
        for k in ['title','subtitle','couple','intro','primary','secondary','background']:
            c.execute('INSERT OR REPLACE INTO settings VALUES(?,?)',(k,request.form.get(k,'')))
        c.commit(); c.close(); return redirect(url_for('admin'))
    gifts=[dict(r) for r in c.execute('SELECT * FROM gifts ORDER BY id')]; settings={r['key']:r['value'] for r in c.execute('SELECT key,value FROM settings')}; history=[dict(r) for r in c.execute('SELECT * FROM choice_history ORDER BY id DESC LIMIT 100')]; c.close()
    return render_template('admin.html',gifts=gifts,settings=settings,history=history)

@app.post('/admin/gift/add')
def add_gift():
    name=request.form.get('name','').strip(); link=request.form.get('link','').strip()
    if not name:return redirect(url_for('admin'))
    try:image=save_upload(request.files.get('image'))
    except ValueError:image=None
    c=conn(); c.execute('INSERT INTO gifts(name,link,image) VALUES(?,?,?)',(name,link,image)); c.commit(); c.close(); return redirect(url_for('admin'))

@app.post('/admin/gift/<int:gift_id>/edit')
def edit_gift(gift_id):
    c=conn(); row=c.execute('SELECT image FROM gifts WHERE id=?',(gift_id,)).fetchone()
    if not row:c.close(); return redirect(url_for('admin'))
    image=row['image']
    try:new=save_upload(request.files.get('image')); image=new or image
    except ValueError:pass
    c.execute('UPDATE gifts SET name=?,link=?,image=? WHERE id=?',(request.form.get('name','').strip(),request.form.get('link','').strip(),image,gift_id)); c.commit(); c.close(); return redirect(url_for('admin'))

@app.post('/admin/gift/<int:gift_id>/delete')
def delete_gift(gift_id):
    c=conn(); c.execute('DELETE FROM gifts WHERE id=?',(gift_id,)); c.commit(); c.close(); return redirect(url_for('admin'))

@app.post('/admin/reset/<int:gift_id>')
def reset(gift_id):
    c=conn(); c.execute("UPDATE gifts SET status='disponivel',chosen_by=NULL,chosen_at=NULL WHERE id=?",(gift_id,)); c.execute('UPDATE choice_history SET active=0 WHERE gift_id=? AND active=1',(gift_id,)); c.commit(); export_choices(c); c.close(); return redirect(url_for('admin'))

@app.get('/admin/escolhas.xlsx')
def download_choices():
    export_choices(); return send_file(EXPORT,as_attachment=True,download_name='escolhas-cha-de-panela.xlsx')

init_db()
if __name__=='__main__':app.run(host='0.0.0.0',port=int(os.environ.get('PORT',5000)),debug=False)
