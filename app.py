
import io
import re
import streamlit as st
import docx
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from pypdf import PdfReader
import google.generativeai as genai

# 1. Налаштування сторінки
st.set_page_config(
    page_title="LegalCounterDraft (Gemini) — Зустрічні договори",
    page_icon="⚖️",
    layout="wide"
)

# 2. Допоміжні функції для зчитування файлів
def extract_text_from_docx(file_bytes) -> str:
    doc = docx.Document(io.BytesIO(file_bytes))
    full_text = [para.text for para in doc.paragraphs if para.text.strip()]
    return "\n".join(full_text)

def extract_text_from_pdf(file_bytes) -> str:
    reader = PdfReader(io.BytesIO(file_bytes))
    text = ""
    for page in reader.pages:
        extracted = page.extract_text()
        if extracted:
            text += extracted + "\n"
    return text

# 3. Розумний конвертер Markdown у професійний Word (.docx)
def markdown_to_legal_docx(md_text: str, doc_title: str = "") -> io.BytesIO:
    doc = Document()
    
    # Налаштування стилю документа за стандартами діловодства
    normal_style = doc.styles['Normal']
    font = normal_style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)
    font.color.rgb = RGBColor(0, 0, 0)
    
    for section in doc.sections:
        section.top_margin = Inches(0.79)     # 2 см
        section.bottom_margin = Inches(0.79)  # 2 см
        section.left_margin = Inches(0.98)    # 2.5 см
        section.right_margin = Inches(0.59)   # 1.5 см
        
    if doc_title:
        title_p = doc.add_paragraph()
        title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = title_p.add_run(doc_title)
        run.bold = True
        run.font.size = Pt(14)
        title_p.paragraph_format.space_after = Pt(12)
        
    lines = md_text.split('\n')
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if not line:
            i += 1
            continue
            
        # Якщо натрапили на таблицю (початок з | )
        if line.startswith('|') and line.endswith('|'):
            table_rows = []
            while i < len(lines) and lines[i].strip().startswith('|') and lines[i].strip().endswith('|'):
                row_line = lines[i].strip()
                # Пропускаємо технічні роздільники на зразок |---|---|
                if not re.match(r'^\|[\s\-:|]+\|$', row_line):
                    cols = [c.strip() for c in row_line.strip('|').split('|')]
                    table_rows.append(cols)
                i += 1
                
            if table_rows:
                max_cols = max(len(r) for r in table_rows)
                table = doc.add_table(rows=len(table_rows), cols=max_cols)
                table.style = 'Table Grid'
                table.alignment = WD_TABLE_ALIGNMENT.CENTER
                
                for r_idx, row_data in enumerate(table_rows):
                    row = table.rows[r_idx]
                    for c_idx in range(max_cols):
                        cell_text = row_data[c_idx] if c_idx < len(row_data) else ""
                        cell = row.cells[c_idx]
                        cell.text = cell_text
                        for p in cell.paragraphs:
                            p.paragraph_format.line_spacing = 1.0
                            if r_idx == 0:
                                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                for r in p.runs:
                                    r.bold = True
                            else:
                                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                doc.add_paragraph()  # Відступ після таблиці
            continue
            
        # Заголовки розділів або додатків
        if line.startswith('###') or line.startswith('##') or line.startswith('#'):
            clean_title = re.sub(r'^#+\s*', '', line)
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if ("ДОДАТОК" in clean_title.upper() or "ДОГОВІР" in clean_title.upper()) else WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(clean_title)
            run.bold = True
            run.font.size = Pt(12)
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(3)
        else:
            # Звичайний пункт договору
            p = doc.add_paragraph()
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(3)
            
            # Підтримка напівжирного тексту **текст**
            parts = re.split(r'(\*\*.*?\*\*)', line)
            for part in parts:
                if part.startswith('**') and part.endswith('**'):
                    run = p.add_run(part[2:-2])
                    run.bold = True
                else:
                    p.add_run(part)
        i += 1
        
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

# 4. Системний промпт
SYSTEM_PROMPT = """
Ти — провідний український юрист-договірник (Senior Contract Lawyer), експерт із цивільного та господарського права України.
Твоє завдання — проаналізувати наданий проєкт цивільно-правового договору, захистити юридичні та фінансові інтереси визначеної Сторони та підготувати:
1. Детальний Протокол розбіжностей (з обов'язковими посиланнями на чинне законодавство України).
2. Повний, юридично грамотний зустрічний проєкт договору, готовий до підписання або направлення контрагенту.

### ПРАВИЛА ТА ОБМЕЖЕННЯ:
1. ПРИОРИТЕТ ІНТЕРЕСІВ: Усі зміни мають бути виключно на користь Сторони, яку визначив користувач, з урахуванням обраної Стратегії та Особливих вимог.
2. СУВОРЕ ДОТРИМАННЯ ЗАКОНОДАВСТВА:
   - ОБОВ'ЯЗКОВО вказувати точні статті та назви актів (Цивільний кодекс України, Господарський кодекс України, профільні закони України).
   - ЗАБОРОНЕНО вигадувати неіснуючі статті чи посилатися на недіючі норми. Якщо норма диспозитивна — зазначай: «Норма диспозитивна (ст. Х ЦКУ), зміна вноситься для нівелювання комерційного ризику».
3. ЗБЕРЕЖЕННЯ ЦІЛІСНОСТІ СТРУКТУРИ:
   - Не змінюй без потреби нумерацію та назви розділів і додатків договору.
   - Якщо в договорі є Додатки (специфікації, графіки, переліки), зберігай їхню структуру. Якщо це таблиця — оформлюй таблицею Markdown.

### СТРУКТУРА ТВОЄЇ ВІДПОВІДІ:
Твоя відповідь ОБОВ'ЯЗКОВО повинна містити чітко розмежовані блоки:

### РОЗДІЛ 1. ПРОТОКОЛ РОЗБІЖНОСТЕЙ
Оформи у вигляді Markdown-таблиці з колонками:
| № та оригінальний пункт | Запропонована зустрічна редакція | Рівень ризику (Високий/Середній/Низький) | Юридичне обґрунтування та норма чинного законодавства |

### РОЗДІЛ 2. ГОТОВИЙ ЗУСТРІЧНИЙ ПРОЄКТ ДОГОВОРУ
Надай ПОВНИЙ суцільний текст договору (від Преамбули до Реквізитів і Додатків) з інтегрованими новими редакціями пунктів. Без закреслень — лише чистий юридичний текст зустрічної версії, готовий до підписання.
"""

# 5. Бокова панель
with st.sidebar:
    st.header("⚙️ Налаштування Gemini")
    api_key = st.text_input("Google AI Studio API Key:", type="password", help="Вставте ваш ключ AIzaSy...")
    selected_model = st.selectbox(
        "Модель:",
        ["gemini-3.8-flash"],
        index=0,
        help="Найновіша швидка модель Gemini."
    )
    st.success("✅ Google API діє безкоштовно")

# 6. Головний екран
st.title("⚖️ Генератор зустрічних договорів (Gemini)")
st.markdown("Сервіс аналізує договір, виявляє ризики за нормами ЦКУ/ГКУ та формує зустрічний проєкт під ваші вимоги.")

col1, col2 = st.columns(2)

with col1:
    role_option = st.selectbox(
        "Чию сторону захищаємо?",
        ["Замовник / Покупець", "Виконавець / Підрядник / Постачальник", "Орендар", "Орендодавець", "Інше (вказати вручну)"]
    )
    if role_option == "Інше (вказати вручну)":
        target_role = st.text_input("Вкажіть вашу сторону:")
    else:
        target_role = role_option

with col2:
    strategy = st.selectbox(
        "Стратегія правок:",
        [
            "Максимальний захист моїх інтересів (жорсткі правки)",
            "Збалансований варіант (ринковий компроміс)",
            "Тільки приведення у відповідність до імперативних норм"
        ]
    )

custom_requirements = st.text_area(
    "Особливі вимоги / Комерційні умови (опціонально):",
    placeholder="Наприклад: штрафи не більше 10%; підсудність у м. Києві; оплата протягом 7 робочих днів.",
    height=80
)

# 7. Введення договору
st.subheader("📄 Текст проєкту договору")
tab_file, tab_text = st.tabs(["Завантажити файл (.docx / .pdf)", "Вставити текст вручну"])

contract_text = ""

with tab_file:
    uploaded_file = st.file_uploader("Оберіть файл договору", type=["docx", "pdf"])
    if uploaded_file is not None:
        file_bytes = uploaded_file.read()
        if uploaded_file.name.endswith(".docx"):
            contract_text = extract_text_from_docx(file_bytes)
        elif uploaded_file.name.endswith(".pdf"):
            contract_text = extract_text_from_pdf(file_bytes)
        st.success(f"Файл успішно завантажено. Довжина тексту: {len(contract_text)} символів.")

with tab_text:
    manual_text = st.text_area("Вставте текст договору сюди:", height=250)
    if manual_text.strip():
        contract_text = manual_text

# 8. Запуск аналізу
if st.button("🚀 Сформувати зустрічний проєкт", type="primary"):
    if not api_key:
        st.error("Будь ласка, вставте ваш Google AI Studio API Key у ліву бокову панель!")
    elif not contract_text.strip():
        st.error("Будь ласка, завантажте файл або вставте текст договору!")
    else:
        with st.spinner("Gemini аналізує договір за законодавством України та створює зустрічний проєкт... Зачекайте 15–30 секунд."):
            try:
                genai.configure(api_key=api_key)
                model = genai.GenerativeModel(
                    model_name=selected_model,
                    system_instruction=SYSTEM_PROMPT
                )
                
                user_prompt = f"""
                ПАРАМЕТРИ АНАЛІЗУ:
                - СТОРОНА: {target_role}
                - СТРАТЕГІЯ: {strategy}
                - ОСОБЛИВІ ВИМОГИ: {custom_requirements if custom_requirements.strip() else 'Загальний захист інтересів сторони'}
                
                ТЕКСТ ПРОЄКТУ ДОГОВОРУ:
                {contract_text}
                """
                
                response = model.generate_content(
                    user_prompt,
                    generation_config={"temperature": 0.2}
                )
                
                result_text = response.text
                st.session_state["result"] = result_text

            except Exception as e:
                st.error(f"Помилка при зверненні до Google Gemini: {e}")

# 9. Відображення результату та окремий експорт
if "result" in st.session_state:
    st.divider()
    st.success("✅ Зустрічний проєкт та Протокол успішно сформовано!")
    
    full_result = st.session_state["result"]
    
    if "### РОЗДІЛ 2. ГОТОВИЙ ЗУСТРІЧНИЙ ПРОЄКТ ДОГОВОРУ" in full_result:
        parts = full_result.split("### РОЗДІЛ 2. ГОТОВИЙ ЗУСТРІЧНИЙ ПРОЄКТ ДОГОВОРУ")
        protocol_part = parts[0].replace("### РОЗДІЛ 1. ПРОТОКОЛ РОЗБІЖНОСТЕЙ", "").strip()
        counter_contract_text = parts[1].strip()
    else:
        protocol_part = full_result
        counter_contract_text = full_result

    res_tab1, res_tab2 = st.tabs(["📊 Протокол розбіжностей та ризики", "📝 Чистий зустрічний договір"])
    
    with res_tab1:
        st.markdown(protocol_part)
        
        # Експорт Протоколу у Word
        protocol_docx = markdown_to_legal_docx(protocol_part, doc_title="ПРОТОКОЛ РОЗБІЖНОСТЕЙ ТА ОЦІНКА РИЗИКІВ")
        st.download_button(
            label="📥 Завантажити Протокол розбіжностей (.docx)",
            data=protocol_docx,
            file_name="Протокол_розбіжностей.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        
    with res_tab2:
        st.markdown(counter_contract_text)
        
        # Експорт Договору у Word
        contract_docx = markdown_to_legal_docx(counter_contract_text)
        st.download_button(
            label="📥 Завантажити зустрічний договір (.docx)",
            data=contract_docx,
            file_name="Зустрічний_проєкт_договору.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )