import streamlit as st
import pandas as pd
import io
import re
import zipfile
import spacy
from docx import Document
import fitz  # PyMuPDF
import firebase_admin
from firebase_admin import credentials, firestore


# --- 1. ІНІЦІАЛІЗАЦІЯ FIREBASE ---
def init_firebase():
    if not firebase_admin._apps:
        try:
            fb_secrets = dict(st.secrets["firebase"])
            fb_secrets["private_key"] = fb_secrets["private_key"].replace("\\n", "\n")
            cred = credentials.Certificate(fb_secrets)
            firebase_admin.initialize_app(cred)
        except Exception as e:
            st.error(f"Firebase Error: {e}")
            st.stop()
    return firestore.client()


db = init_firebase()


# --- 2. ЗАВАНТАЖЕННЯ МОДЕЛІ NLP ---
@st.cache_resource
def load_nlp():
    return spacy.load("pl_core_news_sm")


nlp = load_nlp()


# --- 3. ЛОГІКА ТОКЕНІЗАЦІЇ (REGEX + NLP) ---
def get_token(value, entity_type, mapping):
    if value not in mapping:
        idx = len([v for v in mapping.values() if entity_type in v]) + 1
        mapping[value] = f"<{entity_type}_{idx}>"
    return mapping[value]


def process_text(text, custom_list, mapping):
    if not text or str(text) == "nan":
        return text

    doc = nlp(str(text))
    result_text = str(text)

    # 1. Пошук сутностей через spaCy (Імена)
    # Збираємо збіги, щоб замінювати з кінця (щоб не збивати індекси)
    matches = []
    for ent in doc.ents:
        # Ігноруємо дуже короткі сутності та ті, що не схожі на імена/міста
        if len(ent.text) < 3:
            continue

        if ent.label_ in ["PERSON", "ORG", "GPE"]:
            # Додаткова перевірка: якщо слово у нижньому регістрі — це навряд чи ім'я
            if ent.text[0].isupper():
                matches.append((ent.start_char, ent.end_char, ent.label_))
    # 2. Пошук через Regex (Пошта, Телефон, IBAN, Картки)
    patterns = {
        "EMAIL": r"[\w\.-]+@[\w\.-]+\.\w+",
        "IBAN": r"[A-Z]{2}\d{2}[A-Z0-9]{11,30}",
        "CARD": r"\b(?:\d[ -]*?){13,16}\b",
        "PHONE": r"(\+48|\+380|0)\s?[\d\-\s]{7,12}\b",  # Більш точний паттерн для UA/PL
        "SWIFT": r"\b[A-Z]{6}[A-Z0-9]{2}([A-Z0-9]{3})?\b"  # Додаємо SWIFT
    }

    for label, pattern in patterns.items():
        for match in re.finditer(pattern, result_text):
            matches.append((match.start(), match.end(), label))

    # 3. Пошук вашого кастомного списку з БД
    if custom_list:
        for word in custom_list:
            if word.lower() in result_text.lower():
                for m in re.finditer(re.escape(word), result_text, re.IGNORECASE):
                    matches.append((m.start(), m.end(), "CUSTOM"))

    # Сортуємо та видаляємо дублікати координат
    matches = sorted(list(set(matches)), key=lambda x: x[0], reverse=True)

    # Заміна
    for start, end, label in matches:
        original = result_text[start:end]
        token = get_token(original, label, mapping)
        result_text = result_text[:start] + token + result_text[end:]

    return result_text


# --- 4. ОБРОБНИКИ ФАЙЛІВ ---

def handle_pdf(file_bytes, custom_list, mapping):
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    for page in doc:
        # Для PDF ми просто замальовуємо знайдені слова з кастомного списку та Regex
        # (NLP в PDF складніше через координати, тому фокусуємось на точному пошуку)
        for label, pattern in {"EMAIL": r"\S+@\S+", "IBAN": r"[A-Z]{2}\d{2}.+"}.items():
            for m in page.search_for(pattern):  # Тут можна додати більше пошуку
                page.add_redact_annot(m, fill=(0, 0, 0))

        if custom_list:
            for word in custom_list:
                for area in page.search_for(word):
                    token = get_token(word, "CUSTOM", mapping)
                    page.add_redact_annot(area, fill=(0, 0, 0))
                    page.apply_redactions()
                    page.insert_text(area.tl, token, color=(1, 1, 1), fontsize=8)
    return doc.tobytes()


def handle_docx(file_bytes, custom_list, mapping):
    doc = Document(io.BytesIO(file_bytes))
    for p in doc.paragraphs:
        p.text = process_text(p.text, custom_list, mapping)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                cell.text = process_text(cell.text, custom_list, mapping)
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def handle_xlsx(file_bytes, custom_list, mapping):
    df = pd.read_excel(io.BytesIO(file_bytes))
    for col in df.columns:
        df[col] = df[col].astype(str).apply(lambda x: process_text(x, custom_list, mapping))
    out = io.BytesIO()
    df.to_excel(out, index=False)
    return out.getvalue()


# --- 5. ІНТЕРФЕЙС ---

st.title("🛡️ Secure Tokenizer Pro")

with st.sidebar:
    st.header("Налаштування списку")
    ref = db.collection("settings").document("blacklist")
    db_names = ref.get().to_dict().get("names", "") if ref.get().exists else ""

    input_names = st.text_area("Введіть назви через кому:", value=db_names)
    if st.button("Зберегти в БД"):
        ref.set({"names": input_names})
        st.success("Оновлено!")

    blacklist = [n.strip() for n in input_names.split(",") if n.strip()]

files = st.file_uploader("Завантажте файли", accept_multiple_files=True, type=['pdf', 'docx', 'xlsx'])

if files and st.button("Обробити"):
    session_mapping = {}
    zip_buf = io.BytesIO()

    with zipfile.ZipFile(zip_buf, 'w') as zf:
        for f in files:
            data = f.read()
            ext = f.name.split(".")[-1].lower()

            if ext == 'pdf':
                res = handle_pdf(data, blacklist, session_mapping)
            elif ext == 'docx':
                res = handle_docx(data, blacklist, session_mapping)
            elif ext == 'xlsx':
                res = handle_xlsx(data, blacklist, session_mapping)

            zf.writestr(f"anonymized_{f.name}", res)

        # Звіт
        report = pd.DataFrame(session_mapping.items(), columns=["Оригінал", "Токен"])
        rep_buf = io.BytesIO()
        report.to_excel(rep_buf, index=False)
        zf.writestr("mapping_report.xlsx", rep_buf.getvalue())

    st.success("Готово!")
    st.download_button("Завантажити архів", zip_buf.getvalue(), "result.zip")