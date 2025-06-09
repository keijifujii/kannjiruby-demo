from flask import Flask, render_template, request, send_file, redirect, url_for, flash
from docx import Document as DocxReader, Document as DocxWriter
import io, datetime, re, csv
from fugashi import Tagger

app = Flask(__name__)
app.secret_key = "replace-this"

# -------------------------------------------
# 2) 漢字かどうかを判定する共通正規表現
#    （CJK統合漢字ブロック：一～鿿）
# -------------------------------------------
KANJI_CHAR_RE = re.compile(r'[一-鿿]')

# -------------------------------------------
# 3) 漢字→学年マッピング読み込み
#    CSV 形式：kanji_grade.csv（「漢字,grade」の2列）
# -------------------------------------------
def load_kanji_grade_mapping(csv_path='kanji_grade.csv'):
    mapping = {}
    try:
        with open(csv_path, encoding='utf-8') as f:
            reader = csv.reader(f)
            for row in reader:
                if len(row) < 2:
                    continue
                kanji = row[0].strip()
                try:
                    grade = int(row[1])
                except ValueError:
                    continue
                if kanji and 1 <= grade <= 9:
                    mapping[kanji] = grade
    except FileNotFoundError:
        print(f"[Warning] {csv_path} が見つかりません。すべての漢字にフリガナを振ります。")
    return mapping

KANJI_GRADE = load_kanji_grade_mapping()

# -------------------------------------------
# 4) MeCab Tagger
# -------------------------------------------

tagger = Tagger()

# -------------------------------------------
# 5) 注釈関数：形態素トークンごとに読みを取得し、
#    漢字部分のみにルビを振る
# -------------------------------------------
def annotate_by_grade(text: str, threshold: int) -> str:
    output = []
    for token in tagger(text):
        surface = token.surface
        # token.feature からカタカナ読みを後方検索
        reading = None
        for feat in reversed(token.feature):
            if feat and re.fullmatch(r'[\u30A0-\u30FF]+', feat):
                reading = feat
                break
        if reading:
            try:
                import jaconv
                reading = jaconv.kata2hira(reading)
            except ImportError:
                pass
        if not reading or not KANJI_CHAR_RE.search(surface):
            output.append(surface)
            continue
        m = re.match(r'^([一-鿿]+)(.*)$', surface)
        if m:
            kanji_part, rest = m.group(1), m.group(2)
        else:
            kanji_part, rest = surface, ''
        if rest and len(kanji_part) == 1:
            kanji_reading = reading[:-len(rest)]
        else:
            kanji_reading = reading
        grades = [KANJI_GRADE.get(ch) for ch in kanji_part]
        if any(g is None for g in grades):
            attach = True
        elif all(g < threshold for g in grades):
            attach = False
        else:
            attach = True
        if not attach:
            output.append(surface)
        else:
            annotated = f"{kanji_part}（{kanji_reading}）{rest}"
            output.append(annotated)
    return ''.join(output)

# -------------------------------------------
# 6) .docm 出力用関数（テンプレートのマクロを保持）
# -------------------------------------------
import zipfile

def make_docm(raw_text: str) -> io.BytesIO:
    """
    template.docm のマクロ有効部分を保持しつつ、
    word/document.xml の内容だけを差し替えて新規 docm を作成します。
    """
    template_path = 'template.docm'
    # テンプレート docm 内のすべてのパーツを読み込む
    with zipfile.ZipFile(template_path, 'r') as zin:
        parts = {name: zin.read(name) for name in zin.namelist()}

    # 一時ドキュメントで新しい document.xml を生成
    tmp_doc = DocxWriter()
    for line in raw_text.splitlines():
        tmp_doc.add_paragraph(line)
    tmp_io = io.BytesIO()
    tmp_doc.save(tmp_io)
    tmp_io.seek(0)
    # docx として生成された一時 ZIP から document.xml を抽出
    with zipfile.ZipFile(tmp_io, 'r') as ztmp:
        new_doc_xml = ztmp.read('word/document.xml')

    # テンプレートの document.xml を置き換え
            # XML をパースして rubyPr にフォントサイズを指定
        import xml.etree.ElementTree as ET
        # 名前空間定義
        namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        document_xml = new_doc_xml.decode('utf-8')
        root = ET.fromstring(document_xml)
        # ルビ（w:ruby）要素をすべて処理
        for ruby in root.findall('.//w:ruby', namespaces):
            # rubyPr 要素を取得または作成
            rubyPr = ruby.find('w:rubyPr', namespaces)
            if rubyPr is None:
                rubyPr = ET.SubElement(ruby, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rubyPr')
            # フォントサイズ指定（半ポイント単位、例: 18 -> 9pt）
            sz = ET.SubElement(rubyPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz')
            sz.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '18')
            szCs = ET.SubElement(rubyPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}szCs')
            szCs.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '18')
        # 変更をバイト列に戻す
        new_doc_xml = ET.tostring(root, encoding='utf-8')
        parts['word/document.xml'] = new_doc_xml

    # 新規 docm をメモリ上で作成
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w') as zout:
        for name, data in parts.items():
            zout.writestr(name, data)
    buf.seek(0)
    return buf

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        grade_str = request.form.get('grade', '').strip()
        if not grade_str.isdecimal():
            flash('学年を正しく選択してください', 'danger')
            return redirect(url_for('index'))
        threshold_grade = int(grade_str)

        uploaded = request.files.get('text_file')
        raw = ''
        if uploaded and uploaded.filename:
            filename = uploaded.filename.lower()
            data = uploaded.read()
            if filename.endswith('.docx'):
                stream = io.BytesIO(data)
                reader = DocxReader(stream)
                paragraphs = [p.text for p in reader.paragraphs]
                raw = "\n".join(paragraphs)
            else:
                for enc in ('utf-8', 'cp932'):
                    try:
                        raw = data.decode(enc)
                        break
                    except UnicodeDecodeError:
                        continue
                else:
                    flash('テキストファイルが UTF-8 でも CP932 でもデコードできません', 'danger')
                    return redirect(url_for('index'))
        else:
            raw = request.form.get('source_text', '').strip()

        if not raw:
            flash('ファイルをアップロードするか、テキストを入力してください', 'danger')
            return redirect(url_for('index'))

        annotated = annotate_by_grade(raw, threshold_grade)
        docm_io = make_docm(annotated)
        filename_out = f"annotated_{datetime.datetime.now():%Y%m%d_%H%M%S}.docm"
        return send_file(
            docm_io,
            as_attachment=True,
            download_name=filename_out,
            mimetype='application/vnd.ms-word.document.macroEnabled.12'
        )
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True, port=5000)
