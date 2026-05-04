import streamlit as st
from openai import OpenAI
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from io import BytesIO
import json

# =========================
# PAGE CONFIG
# =========================
st.set_page_config(
    page_title="Sabaq AI",
    page_icon="📘",
    layout="wide"
)

client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])


# =========================
# DESIGN
# =========================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800;900&display=swap');

* {
    font-family: 'Inter', sans-serif;
}

.stApp {
    background:
        radial-gradient(circle at top left, rgba(124,58,237,0.25), transparent 35%),
        radial-gradient(circle at top right, rgba(6,182,212,0.25), transparent 30%),
        linear-gradient(135deg, #F8FAFC 0%, #EEF2FF 50%, #ECFEFF 100%);
}

#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}

.block-container {
    padding-top: 2rem;
    padding-bottom: 3rem;
}

.hero {
    background: linear-gradient(135deg, #2563EB, #7C3AED, #EC4899);
    padding: 38px;
    border-radius: 34px;
    color: white;
    box-shadow: 0 24px 60px rgba(37, 99, 235, 0.35);
    text-align: center;
    margin-bottom: 28px;
}

.hero-title {
    font-size: 54px;
    font-weight: 900;
}

.hero-subtitle {
    font-size: 20px;
    opacity: 0.95;
    margin-top: 8px;
}

.badge {
    display: inline-block;
    padding: 9px 16px;
    border-radius: 999px;
    background: rgba(255,255,255,0.20);
    color: white;
    font-weight: 800;
    margin: 6px;
}

.feature-card {
    background: rgba(255,255,255,0.92);
    backdrop-filter: blur(14px);
    border-radius: 28px;
    padding: 26px;
    text-align: center;
    box-shadow: 0 18px 45px rgba(15, 23, 42, 0.12);
    border: 2px solid #E0E7FF;
    min-height: 185px;
    margin-bottom: 18px;
}

.feature-icon {
    font-size: 44px;
    margin-bottom: 10px;
}

.feature-title {
    font-size: 22px;
    font-weight: 900;
    color: #1E3A8A;
}

.feature-text {
    color: #475569;
    font-size: 15px;
    margin-top: 8px;
}

.result-box {
    background: white;
    padding: 28px;
    border-radius: 28px;
    box-shadow: 0 18px 45px rgba(15, 23, 42, 0.13);
    border-left: 8px solid #7C3AED;
    margin-top: 20px;
}

section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #FFFFFF 0%, #EEF2FF 100%);
    border-right: 1px solid #DBEAFE;
}

section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3 {
    color: #1D4ED8;
    font-weight: 900;
}

.stTextInput input,
.stTextArea textarea,
.stSelectbox div[data-baseweb="select"],
.stMultiSelect div[data-baseweb="select"] {
    border-radius: 16px !important;
    border: 2px solid #BFDBFE !important;
    background: white !important;
}

.stButton > button {
    border-radius: 20px;
    height: 64px;
    font-size: 18px;
    font-weight: 900;
    color: white;
    border: none;
    background: linear-gradient(135deg, #2563EB, #7C3AED, #EC4899);
    box-shadow: 0 14px 30px rgba(124, 58, 237, 0.35);
}

.stButton > button:hover {
    transform: translateY(-3px);
    color: white;
}

.stDownloadButton > button {
    border-radius: 18px;
    height: 55px;
    font-size: 16px;
    font-weight: 800;
    background: linear-gradient(135deg, #10B981, #06B6D4);
    color: white;
    border: none;
}
</style>
""", unsafe_allow_html=True)


st.markdown("""
<div class="hero">
    <div class="hero-title">📘 Sabaq AI</div>
    <div class="hero-subtitle">
        Мұғалімдерге арналған ҚМЖ, жұмыс парағы және ойын генераторы
    </div>
    <div style="margin-top:18px;">
        <span class="badge">⚡ ҚМЖ генератор</span>
        <span class="badge">📄 Word кесте</span>
        <span class="badge">🎮 EdTech ойындар</span>
        <span class="badge">🇰🇿 Қазақша</span>
    </div>
</div>
""", unsafe_allow_html=True)


# =========================
# SIDEBAR INPUTS
# =========================
with st.sidebar:
    st.header("⚙️ Сабақ мәліметтері")

    school = st.text_input("ББҰ / мектеп атауы", "№1 Хромтау орта мектебі")
    teacher = st.text_input("Педагогтің Т.А.Ә.", "Сисекенова А.М.")
    subject = st.text_input("Пән", "Биология")
    grade = st.text_input("Сынып", "7")
    section = st.text_input("Бөлім", "7.3 C Координация және реттелу")
    date = st.text_input("Күні", "")
    topic = st.text_area("Сабақ тақырыбы")
    objective = st.text_area("Оқу бағдарламасына сәйкес оқыту мақсаттары")

    lesson_type = st.selectbox(
        "Сабақ түрі",
        ["Жаңа сабақ", "Ашық сабақ", "Зертханалық жұмыс", "Практикалық сабақ", "Қайталау сабағы"]
    )

    platforms = st.multiselect(
        "Қолданылатын платформалар",
        ["Wordwall", "Genially", "Kahoot", "Quizizz", "LearningApps", "Mozaik 3D", "Canva", "BilimClass", "Padlet"]
    )

    ebq = st.radio("ЕБҚ тапсырмасы қосылсын ба?", ["Иә", "Жоқ"])


# =========================
# FEATURE CARDS
# =========================
c1, c2, c3 = st.columns(3)

with c1:
    st.markdown("""
    <div class="feature-card">
        <div class="feature-icon">📘</div>
        <div class="feature-title">ҚМЖ генератор</div>
        <div class="feature-text">Ресми құрылым, сабақ барысы, бағалау, дескриптор, ресурстар.</div>
    </div>
    """, unsafe_allow_html=True)

with c2:
    st.markdown("""
    <div class="feature-card">
        <div class="feature-icon">📄</div>
        <div class="feature-title">Жұмыс парағы</div>
        <div class="feature-text">Оқушыға арналған тапсырмалар, PISA, ЕБҚ, рефлексия.</div>
    </div>
    """, unsafe_allow_html=True)

with c3:
    st.markdown("""
    <div class="feature-card">
        <div class="feature-icon">🎮</div>
        <div class="feature-title">Ойын идеялары</div>
        <div class="feature-text">Wordwall, Kahoot, Genially, LearningApps тапсырмалары.</div>
    </div>
    """, unsafe_allow_html=True)


# =========================
# WORD HELPERS
# =========================
def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def set_cell_text(cell, text, bold=False, size=9):
    cell.text = ""
    paragraph = cell.paragraphs[0]
    run = paragraph.add_run(str(text))
    run.bold = bold
    run.font.name = "Times New Roman"
    run.font.size = Pt(size)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def set_table_borders(table):
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    borders = OxmlElement("w:tblBorders")

    for border_name in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        border = OxmlElement(f"w:{border_name}")
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), "8")
        border.set(qn("w:space"), "0")
        border.set(qn("w:color"), "000000")
        borders.append(border)

    tbl_pr.append(borders)


def add_paragraph(doc, text, bold=False, size=12, align="left"):
    p = doc.add_paragraph()
    run = p.add_run(str(text))
    run.bold = bold
    run.font.name = "Times New Roman"
    run.font.size = Pt(size)

    if align == "center":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == "right":
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT

    return p


# =========================
# AI FUNCTIONS
# =========================
def generate_qmj_json():
    prompt = f"""
Сен кәсіби әдіскерсің. Қазақ тілінде толық ҚМЖ жаса.

Тек JSON форматында жауап бер. Markdown қолданба. Артық мәтін жазба.

Берілген мәлімет:
ББҰ: {school}
Педагог: {teacher}
Пән: {subject}
Сынып: {grade}
Бөлім: {section}
Күні: {date}
Сабақ тақырыбы: {topic}
Оқу мақсаты: {objective}
Сабақ түрі: {lesson_type}
Платформалар: {", ".join(platforms)}
ЕБҚ: {ebq}

JSON құрылымы:

{{
  "school": "",
  "section": "",
  "teacher": "",
  "date": "",
  "grade": "",
  "topic": "",
  "learning_objectives": "",
  "lesson_goal": "",
  "ebq_goal": "",
  "values": "",
  "lesson_flow": [
    {{
      "stage": "",
      "teacher_action": "",
      "student_action": "",
      "assessment": "",
      "resources": ""
    }}
  ],
  "homework": "",
  "reflection": ""
}}

lesson_flow ішінде 9 кезең болсын:
1. Сабақтың басы — 5 минут
2. Өткен білімді еске түсіру — 5 минут
3. Жаңа тақырыпты ашу — 5 минут
4. Мұғалім түсіндірмесі — 7 минут
5. Топтық жұмыс — 8 минут
6. Жеке жұмыс — 5 минут
7. PISA тапсырмасы — 5 минут
8. ЕБҚ тапсырмасы — 3 минут
9. Сабақ соңы, рефлексия — 2 минут

Әр кезеңде:
- педагог нақты не айтады, не көрсетеді, қандай сұрақ қояды
- оқушы нақты не істейді
- бағалау дескриптор және баллмен
- ресурстар

Тапсырмалар жоғары деңгейлі болсын: талдау, салыстыру, дәлелдеу, қорытынды жасау.
"""

    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": "Сен қазақ тіліндегі кәсіби ҚМЖ құрастырушы әдіскерсің."},
            {"role": "user", "content": prompt}
        ],
        temperature=0.6
    )

    content = response.choices[0].message.content.strip()

    try:
        return json.loads(content)
    except Exception:
        st.error("Қате: AI JSON форматты бұзып жіберді. Қайта басып көріңіз.")
        st.text(content)
        return None


def generate_worksheet_text():
    prompt = f"""
Қазақ тілінде оқушыға арналған толық жұмыс парағын жаса.

Пән: {subject}
Сынып: {grade}
Сабақ тақырыбы: {topic}
Оқу мақсаты: {objective}
ЕБҚ тапсырмасы: {ebq}

Құрылымы:
# ОҚУШЫҒА АРНАЛҒАН ЖҰМЫС ПАРАҒЫ

Пән:
Сынып:
Тақырып:
Оқу мақсаты:
Оқушының аты-жөні:
Күні:

1-тапсырма. Терминдерді сәйкестендір.
2-тапсырма. Бос орындарды толықтыр.
3-тапсырма. Сызба немесе суретпен жұмыс.
4-тапсырма. Жұптық жұмыс.
5-тапсырма. PISA форматындағы жағдаят.
6-тапсырма. Жоғары деңгейлі сұрақ.
7-тапсырма. ЕБҚ тапсырмасы.
8-тапсырма. Өзін-өзі бағалау.
9-тапсырма. Рефлексия.

Әр тапсырмаға дескриптор және балл қос.
"""

    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": "Сен кәсіби жұмыс парағын құрастырушы әдіскерсің."},
            {"role": "user", "content": prompt}
        ],
        temperature=0.6
    )
    return response.choices[0].message.content


def generate_games_text():
    prompt = f"""
Қазақ тілінде сабаққа арналған цифрлық ойындар мен платформалар идеяларын жаса.

Пән: {subject}
Сынып: {grade}
Сабақ тақырыбы: {topic}
Оқу мақсаты: {objective}
Платформалар: {", ".join(platforms)}

Мына бөлімдермен бер:
1. Wordwall ойыны
2. Genially интерактив тапсырмасы
3. Kahoot викторинасы
4. Quizizz тесті
5. LearningApps сәйкестендіру
6. Mozaik 3D қолдану
7. Офлайн ойын

Әрқайсысына нақты дайын тапсырма мәтінін жаз.
"""

    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": "Сен EdTech әдіскері және ойын тапсырмаларын жасаушы мамансың."},
            {"role": "user", "content": prompt}
        ],
        temperature=0.7
    )
    return response.choices[0].message.content


# =========================
# DOCX GENERATORS
# =========================
def create_qmj_docx(data):
    doc = Document()

    sec = doc.sections[0]
    sec.top_margin = Cm(1.5)
    sec.bottom_margin = Cm(1.5)
    sec.left_margin = Cm(1.3)
    sec.right_margin = Cm(1.3)

    doc.styles["Normal"].font.name = "Times New Roman"
    doc.styles["Normal"].font.size = Pt(10)

    add_paragraph(doc, "ҚЫСҚА МЕРЗІМДІ ЖОСПАР", bold=True, size=14, align="center")

    info_table = doc.add_table(rows=10, cols=2)
    info_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(info_table)

    info = [
        ("ББҰ", data.get("school", "")),
        ("Бөлім", data.get("section", "")),
        ("Педагогтің Т.А.Ә.", data.get("teacher", "")),
        ("Күні", data.get("date", "")),
        ("Сынып", data.get("grade", "")),
        ("Сабақтың тақырыбы", data.get("topic", "")),
        ("Оқу бағдарламасына сәйкес оқыту мақсаттары", data.get("learning_objectives", "")),
        ("Сабақтың мақсаты", data.get("lesson_goal", "")),
        ("ЕБҚЕ мақсаты", data.get("ebq_goal", "")),
        ("Құндылықтар", data.get("values", "")),
    ]

    for i, (label, value) in enumerate(info):
        set_cell_text(info_table.cell(i, 0), label, bold=True, size=9)
        set_cell_text(info_table.cell(i, 1), value, size=9)
        set_cell_shading(info_table.cell(i, 0), "EAF2F8")

    doc.add_paragraph("")
    add_paragraph(doc, "Сабақтың барысы", bold=True, size=13, align="center")

    table = doc.add_table(rows=1, cols=5)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table)

    headers = [
        "Сабақтың кезеңі және уақыт",
        "Педагогтің әрекеті",
        "Оқушының әрекеті",
        "Бағалау, дескрипторлар, балл",
        "Ресурстар"
    ]

    for i, header in enumerate(headers):
        set_cell_text(table.cell(0, i), header, bold=True, size=8)
        set_cell_shading(table.cell(0, i), "D9EAF7")

    for item in data.get("lesson_flow", []):
        row = table.add_row().cells
        set_cell_text(row[0], item.get("stage", ""), bold=True, size=8)
        set_cell_text(row[1], item.get("teacher_action", ""), size=8)
        set_cell_text(row[2], item.get("student_action", ""), size=8)
        set_cell_text(row[3], item.get("assessment", ""), size=8)
        set_cell_text(row[4], item.get("resources", ""), size=8)

    doc.add_paragraph("")
    add_paragraph(doc, "Үй тапсырмасы:", bold=True, size=11)
    add_paragraph(doc, data.get("homework", ""), size=11)

    add_paragraph(doc, "Рефлексия:", bold=True, size=11)
    add_paragraph(doc, data.get("reflection", ""), size=11)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


def create_text_docx(text, title):
    doc = Document()

    sec = doc.sections[0]
    sec.top_margin = Cm(1.5)
    sec.bottom_margin = Cm(1.5)
    sec.left_margin = Cm(1.5)
    sec.right_margin = Cm(1.5)

    doc.styles["Normal"].font.name = "Times New Roman"
    doc.styles["Normal"].font.size = Pt(12)

    add_paragraph(doc, title, bold=True, size=14, align="center")

    for line in text.split("\n"):
        add_paragraph(doc, line, size=12)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


def check_inputs():
    if not topic or not objective:
        st.warning("Алдымен сабақ тақырыбы мен оқу мақсатын толтырыңыз.")
        return False
    return True


# =========================
# BUTTONS
# =========================
col1, col2, col3 = st.columns(3)

make_qmj = col1.button("📘 ҚМЖ жасау + Word кесте", use_container_width=True)
make_sheet = col2.button("📄 Жұмыс парағын жасау", use_container_width=True)
make_games = col3.button("🎮 Ойындар табу", use_container_width=True)


if make_qmj:
    if check_inputs():
        with st.spinner("ҚМЖ жасалып жатыр..."):
            qmj_data = generate_qmj_json()

        if qmj_data:
            st.success("✅ ҚМЖ дайын!")

            st.markdown('<div class="result-box">', unsafe_allow_html=True)
            st.subheader("📘 ҚМЖ қысқаша қарау")
            st.write("**Сабақ тақырыбы:**", qmj_data.get("topic", ""))
            st.write("**Сабақ мақсаты:**", qmj_data.get("lesson_goal", ""))
            st.write("**Құндылықтар:**", qmj_data.get("values", ""))
            st.write("### Сабақ барысы")
            st.dataframe(qmj_data.get("lesson_flow", []), use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

            docx_file = create_qmj_docx(qmj_data)

            st.download_button(
                label="📥 ҚМЖ Word жүктеу (.docx)",
                data=docx_file,
                file_name=f"QMJ_{topic[:25]}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )


if make_sheet:
    if check_inputs():
        with st.spinner("Жұмыс парағы жасалып жатыр..."):
            text = generate_worksheet_text()

        st.success("✅ Жұмыс парағы дайын!")
        st.markdown('<div class="result-box">', unsafe_allow_html=True)
        st.markdown(text)
        st.markdown('</div>', unsafe_allow_html=True)

        docx_file = create_text_docx(text, "ОҚУШЫҒА АРНАЛҒАН ЖҰМЫС ПАРАҒЫ")

        st.download_button(
            label="📥 Жұмыс парағын Word жүктеу",
            data=docx_file,
            file_name=f"Worksheet_{topic[:25]}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )


if make_games:
    if check_inputs():
        with st.spinner("Ойын идеялары жасалып жатыр..."):
            text = generate_games_text()

        st.success("✅ Ойын идеялары дайын!")
        st.markdown('<div class="result-box">', unsafe_allow_html=True)
        st.markdown(text)
        st.markdown('</div>', unsafe_allow_html=True)

        docx_file = create_text_docx(text, "САБАҚҚА АРНАЛҒАН ОЙЫН ИДЕЯЛАРЫ")

        st.download_button(
            label="📥 Ойын идеяларын Word жүктеу",
            data=docx_file,
            file_name=f"Games_{topic[:25]}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
