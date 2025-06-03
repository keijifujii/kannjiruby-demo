# app.py

from flask import Flask, render_template, request, send_file, redirect, url_for, flash
from docx import Document as DocxReader, Document as DocxWriter
import io, datetime, re, csv
from pykakasi import kakasi

app = Flask(__name__)
app.secret_key = "replace-this"

# -------------------------------------------
# 1) pykakasi セットアップ（漢字→ひらがな）
# -------------------------------------------
_kks = kakasi()
_kks.setMode("J", "H")  # J: 漢字 → H: ひらがな
_converter = _kks.getConverter()

# -------------------------------------------
# 2) 漢字かどうかを判定する共通正規表現
#    （CJK統合漢字ブロック：一～龯）
# -------------------------------------------
KANJI_CHAR_RE = re.compile(r'[\u4E00-\u9FFF]')

# -------------------------------------------
# 3) アプリ起動時に「漢字→学年」のマッピングを読み込む
#    CSV 形式：kanji_grade.csv（「漢字,grade」の2列）
# -------------------------------------------
KANJI_GRADE = {}
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

KANJI_GRADE = load_kanji_grade_mapping('kanji_grade.csv')

# -------------------------------------------
# 4) 修正後：連続漢字列をまとめて熟語単位で処理し、
#    学年 threshold に応じてルビを付与する関数
# -------------------------------------------
def annotate_by_grade(text: str, threshold: int) -> str:
    """
    1) 正規表現 r'[\u4E00-\u9FFF]+' で
       連続する漢字の塊（熟語候補）をマッチさせる。
    2) マッチする each 「漢字列 s」について:
       - 各文字 ch の学年 grade = KANJI_GRADE.get(ch) を取得
       - ・mapping に ch がなければ grade is None → 常用漢字外 とみなし、必ずルビを振る
         ・mapping があって、かつすべての grade < threshold → 置き換えずそのまま s を返す
         ・上記以外（少なくとも1文字が grade ≥ threshold） → s の熟語全体にルビを付与
    3) kakasi で s の読み（ひらがな）を取り、"s（読み）" に置き換える
    """
    def replace_match(match: re.Match) -> str:
        s = match.group()  # 例：'仮想環境' や '学校' など、連続する漢字の塊
        grades = [KANJI_GRADE.get(ch) for ch in s]

        # ① もし「どれか1文字でも mapping にない＝grade is None」であれば、
        #    常用漢字外を含む とみなして、一律でルビを付ける
        if any(g is None for g in grades):
            reading = _converter.do(s)
            return f"{s}（{reading}）"

        # ② すべての文字が「grade < threshold（学年より前に学ぶ漢字）」ならそのまま
        if all(g < threshold for g in grades):
            return s

        # ③ それ以外（少なくとも1文字が grade ≥ threshold）なら
        #    熟語全体を kakasi で読みをとり、"熟語（読み）" に置き換える
        reading = _converter.do(s)
        return f"{s}（{reading}）"

    # テキスト全体から「連続漢字列」をすべて置き換える
    return re.sub(r'[\u4E00-\u9FFF]+', replace_match, text)


# -------------------------------------------
# 5) .docx 出力用関数（変更なし）
# -------------------------------------------
def make_docx(raw_text: str) -> io.BytesIO:
    doc = DocxWriter()
    for line in raw_text.splitlines():
        doc.add_paragraph(line)
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# -------------------------------------------
# 6) Flask ルート定義（変更なし）
# -------------------------------------------
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # 学年（1～9）を取得
        grade_str = request.form.get('grade', '').strip()
        if not grade_str.isdecimal():
            flash('学年を正しく選択してください', 'danger')
            return redirect(url_for('index'))
        threshold_grade = int(grade_str)

        # ファイル or テキスト入力から raw を取得
        uploaded = request.files.get('text_file')
        raw = ""
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

        # フリガナ付与処理を実行
        annotated = annotate_by_grade(raw, threshold_grade)

        # .docx を作って返す
        docx_io = make_docx(annotated)
        filename_out = f"annotated_{datetime.datetime.now():%Y%m%d_%H%M%S}.docx"
        return send_file(
            docx_io,
            as_attachment=True,
            download_name=filename_out,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

    return render_template('index.html')


if __name__ == '__main__':
    app.run(debug=True, port=5000)
