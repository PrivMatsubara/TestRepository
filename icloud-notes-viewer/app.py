import subprocess
import json
import html2text
import io
import zipfile
import re
from flask import Flask, render_template, jsonify, send_file, request, abort

app = Flask(__name__)


def run_jxa(script: str) -> str:
    """osascript -l JavaScript でスクリプトを実行し stdout を返す。"""
    proc = subprocess.run(
        ['osascript', '-l', 'JavaScript', '-e', script],
        capture_output=True, text=True, timeout=120
    )
    if proc.returncode != 0:
        raise RuntimeError(proc.stderr.strip() or 'osascript error')
    return proc.stdout.strip()


def get_all_notes() -> list:
    """Notes.app から全メモのメタ情報を取得する。"""
    script = '''
var Notes = Application('Notes');
var all = Notes.notes();
var out = [];
for (var i = 0; i < all.length; i++) {
    var n = all[i];
    var folder = '';
    try { folder = n.container().name(); } catch(e) {}
    out.push({
        id: n.id(),
        name: n.name(),
        modified: n.modificationDate() ? n.modificationDate().toISOString() : null,
        folder: folder
    });
}
JSON.stringify(out);
'''
    return json.loads(run_jxa(script))


def get_note_body(note_id: str) -> str:
    """指定 ID のメモ HTML 本文を取得する。"""
    sid = json.dumps(note_id)
    script = (
        "var Notes = Application('Notes');\n"
        "var all = Notes.notes();\n"
        "var body = '';\n"
        "for (var i = 0; i < all.length; i++) {\n"
        f"    if (all[i].id() === {sid}) {{\n"
        "        body = all[i].body();\n"
        "        break;\n"
        "    }\n"
        "}\n"
        "body;\n"
    )
    return run_jxa(script)


def get_all_notes_with_body() -> list:
    """全メモをメタ情報 + HTML 本文ごと一括取得（download-all 用）。"""
    script = '''
var Notes = Application('Notes');
var all = Notes.notes();
var out = [];
for (var i = 0; i < all.length; i++) {
    var n = all[i];
    var folder = '';
    try { folder = n.container().name(); } catch(e) {}
    var body = '';
    try { body = n.body(); } catch(e) {}
    out.push({
        id: n.id(),
        name: n.name(),
        modified: n.modificationDate() ? n.modificationDate().toISOString() : null,
        folder: folder,
        body: body
    });
}
JSON.stringify(out);
'''
    return json.loads(run_jxa(script))


def to_markdown(html_body: str) -> str:
    h = html2text.HTML2Text()
    h.body_width = 0
    h.unicode_snob = True
    h.ignore_links = False
    return h.handle(html_body)


def safe_fname(name: str) -> str:
    s = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '', name).strip()
    s = re.sub(r'\s+', '_', s)
    return (s[:80] or 'note') + '.md'


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/notes')
def api_notes():
    try:
        notes = get_all_notes()
        notes.sort(key=lambda n: n.get('modified') or '', reverse=True)
        return jsonify({'ok': True, 'notes': notes})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500


@app.route('/api/note')
def api_note():
    nid = request.args.get('id')
    if not nid:
        return jsonify({'ok': False, 'error': 'id required'}), 400
    try:
        html = get_note_body(nid)
        md = to_markdown(html)
        return jsonify({'ok': True, 'html': html, 'markdown': md})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500


@app.route('/api/download')
def api_download():
    nid = request.args.get('id')
    if not nid:
        abort(400)
    try:
        notes = get_all_notes()
        note = next((n for n in notes if n['id'] == nid), None)
        name = note['name'] if note else 'note'
        html = get_note_body(nid)
        content = f'# {name}\n\n{to_markdown(html)}'
        buf = io.BytesIO(content.encode('utf-8'))
        buf.seek(0)
        return send_file(buf, mimetype='text/markdown', as_attachment=True,
                         download_name=safe_fname(name))
    except Exception:
        abort(500)


@app.route('/api/download-all')
def api_download_all():
    try:
        notes = get_all_notes_with_body()
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
            for note in notes:
                try:
                    md = to_markdown(note.get('body') or '')
                    content = f"# {note['name']}\n\n{md}"
                    folder = note.get('folder') or 'Notes'
                    zf.writestr(f"{folder}/{safe_fname(note['name'])}", content)
                except Exception:
                    continue
        buf.seek(0)
        return send_file(buf, mimetype='application/zip', as_attachment=True,
                         download_name='icloud-notes.zip')
    except Exception:
        abort(500)


if __name__ == '__main__':
    print('\n  iCloud Notes Viewer → http://localhost:5000\n')
    app.run(debug=True, port=5000)
