# app.py 2025-06-05

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
# 5) 改良版２：漢字＋送り仮名（ひらがな）がある場合は
#    振り仮名を「漢字部分だけ」に付与し、送り仮名はそのまま残す
# -------------------------------------------
def annotate_by_grade(text: str, threshold: int) -> str:
    """
    ① 「漢字列＋直後のひらがな（送り仮名）」を優先的にマッチ。
       振り仮名を付けるのは漢字部分だけとし、送り仮名そのものはルビ外に出す。
    ② マッチしない「純粋な連続漢字」は、従来どおり漢字丸ごとルビを付与する。
    """
    # 1) パターン：グループ1＝漢字+ひらがな、グループ2＝連続漢字
    pattern = re.compile(r'([\u4E00-\u9FFF]+)([ぁ-ゔ]+)|([\u4E00-\u9FFF]+)')

    def replace_match(match: re.Match) -> str:
        # ── グループ1：漢字＋送り仮名（ひらがな）がマッチしている場合 ──
        if match.group(1) and match.group(2):
            kanji_part = match.group(1)      # 例：「政治的」, 「議会」, 「予算」など
            okurigana  = match.group(2)      # 例：「に」, 「は」, 「局は」など
            full_word  = kanji_part + okurigana

            # 学年判定は「漢字部分だけ」で行う
            grades = [KANJI_GRADE.get(ch) for ch in kanji_part]

            # （A）もし漢字部分にマッピングがない文字（常用外など）があれば、
            #     漢字部分だけにルビを付与し、送り仮名はそのまま後ろに付ける
            if any(g is None for g in grades):
                reading = _converter.do(kanji_part)  # 例："せいじてき"
                return f"{kanji_part}（{reading}）{okurigana}"

            # （B）すべての漢字が threshold 未満（もっと早い学年で習う）なら、
            #     漢字＋送り仮名をそのまま返す（ルビ不要）
            if all(g < threshold for g in grades):
                return full_word

            # （C）それ以外（少なくとも1文字が grade ≥ threshold）なら、
            #     漢字部分だけを kakasi で読みを取得し、ルビを付ける
            reading = _converter.do(kanji_part)  # 例："せいじてき"
            return f"{kanji_part}（{reading}）{okurigana}"

        # ── グループ1 がマッチせず、グループ2：純粋な連続漢字がマッチした場合 ──
        else:
            s = match.group(3)  # 例：「学校」「環境」「勉強」など

            grades = [KANJI_GRADE.get(ch) for ch in s]
            # （A）マッピングにない文字があれば、漢字全部にルビ
            if any(g is None for g in grades):
                reading = _converter.do(s)
                return f"{s}（{reading}）"
            # （B）すべての文字が threshold 未満なら、そのまま返す（ルビ不要）
            if all(g < threshold for g in grades):
                return s
            # （C）それ以外は、漢字全体を kakasi で読みを取得し、ルビを付与
            reading = _converter.do(s)
            return f"{s}（{reading}）"

    return pattern.sub(replace_match, text)




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
