import streamlit as st
import pandas as pd
import io
import re
import zipfile
from docx import Document
import fitz  # PyMuPDF
from presidio_analyzer import AnalyzerEngine, PatternRecognizer, RecognizerRegistry
from presidio_anonymizer import AnonymizerEngine
import firebase_admin
from firebase_admin import credentials, firestore


# --- 1. ІНІЦІАЛІЗАЦІЯ FIREBASE ---
def init_firebase_connection():
    try:
        firebase_admin.get_app()
    except ValueError:
        # Вкажіть шлях до вашого json-ключа
        cred = credentials.Certificate("firebase-key.json")
        firebase_admin.initialize_app(cred)
    return firestore.client()


db = init_firebase_connection()


# --- 2. ДВИГУН АНАЛІЗУ (БЕЗ SHADOWING) ---
@st.cache_resource
def load_anonymization_engines(custom_names_list):
    registry = RecognizerRegistry()
    registry.load_predefined_recognizers()

    if custom_names_list:
        list_regex = r"\b(" + "|".join(map(re.escape, custom_names_list)) + r")\b"
        list_rec = PatternRecognizer(
            supported_entity="MANUAL_LIST",
            patterns=[{"name": "custom_list", "score": 1.0, "regex": list_regex}]
        )
        registry.add_recognizer(list_rec)

    # Додаємо IBAN
    iban_rec = PatternRecognizer(
        supported_entity="IBAN",
        patterns=[{"name": "iban", "score": 1.0, "regex": r"[A-Z]{2}\d{2}[A-Z0-9]{11,30}"}]
    )
    registry.add_recognizer(iban_rec)

    # Додаємо Credit Card
    card_rec = PatternRecognizer(
        supported_entity="CREDIT_CARD",
        patterns=[{"name": "card", "score": 1.0, "regex": r"\b(?:\d[ -]*?){13,16}\b"}]
    )
    registry.add_recognizer(card_rec)

    _analyzer = AnalyzerEngine(registry=registry, default_score_threshold=0.4)
    _anonymizer = AnonymizerEngine()
    return _analyzer, _anonymizer


# --- 3. ДОПОМІЖНІ ФУНКЦІЇ ---

def get_token_for_value(original_value, entity_type, mapping_dict):
    if original_value not in mapping_dict:
        # Рахуємо скільки вже є токенів цього типу для індексу
        idx = len([v for v in mapping_dict.values() if entity_type in v]) + 1
        mapping_dict[original_value] = f"<{entity_type}_{idx}>"
    return mapping_dict[original_value]


def process_text_content(raw_text, analyzer_engine, mapping_dict):
    if not raw_text or str(raw_text) == "nan":
        return raw_text

    analysis_results = analyzer_engine.analyze(text=str(raw_text), language='en')
    sorted_res = sorted(analysis_results, key=lambda x: x.start, reverse=True)

    final_text = str(raw_text)
    for res in sorted_res:
        original_chunk = final_text[res.start:res.end]
        token = get_token_for_value(original_chunk, res.entity_type, mapping_dict)
        final_text = final_text[:res.start] + token + final_text[res.end:]
    return final_text


# --- 4. ОБРОБКА ФАЙЛІВ ---

def handle_excel(file_bytes, analyzer_engine, mapping_dict):
    df = pd.read_excel(io.BytesIO(file_bytes))
    for col in df.columns:
        df[col] = df[col].astype(str).apply(lambda x: process_text_content(x, analyzer_engine, mapping_dict))

    out_buffer = io.BytesIO()
    df.to_excel(out_buffer, index=False)
    return out_buffer.getvalue()


def handle_docx(file_bytes, analyzer_engine, mapping_dict):
    doc = Document(io.BytesIO(file_bytes))
    for p in doc.paragraphs:
        p.text = process_text_content(p.text, analyzer_engine, mapping_dict)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                cell.text = process_text_content(cell.text, analyzer_engine, mapping_dict)

    out_buffer = io.BytesIO()
    doc.save(out_buffer)
    return out_buffer.getvalue()


def handle_pdf(file_bytes: bytes, analyzer_engine: AnalyzerEngine, mapping_dict: dict) -> bytes:
    # Явно відкриваємо документ
    doc = fitz.open(stream=file_bytes, filetype="pdf")

    try:
        for page in doc:
            page_text = page.get_text()
            results = analyzer_engine.analyze(text=page_text, language='en')

            if results:
                for res in results:
                    original_val = page_text[res.start:res.end].strip()
                    if not original_val:
                        continue

                    token_val = get_token_for_value(original_val, res.entity_type, mapping_dict)

                    # Пошук координат
                    areas = page.search_for(original_val)
                    for area in areas:
                        # Важливо: деякі версії fitz очікують apply_redactions без контексту
                        page.add_redact_annot(area, fill=(0, 0, 0))

                # Викликаємо один раз на сторінку
                page.apply_redactions()

        return doc.tobytes()
    finally:
        doc.close()  # Гарантоване закриття ресурсу

# --- 5. UI СТРІМЛІТ ---

st.set_page_config(page_title="Data Masking Tool", layout="wide")
st.title("🛡️ Корпоративний Токенізатор")

with st.sidebar:
    st.header("Керування списком")
    settings_ref = db.collection("settings").document("blacklist")
    current_data = settings_ref.get().to_dict() or {"names": ""}

    input_names = st.text_area("Назви компаній/ПІБ (через кому):", value=current_data.get("names", ""))

    if st.button("Зберегти у Firestore"):
        settings_ref.set({"names": input_names})
        st.success("Базу оновлено")

    active_list = [n.strip() for n in input_names.split(",") if n.strip()]

# Завантажуємо двигуни
engine_analyzer, engine_anonymizer = load_anonymization_engines(active_list)

uploaded_files = st.file_uploader("Виберіть файли", type=["xlsx", "docx", "pdf"], accept_multiple_files=True)

if uploaded_files and st.button("Почати обробку"):
    session_mapping = {}
    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(zip_buffer, 'w') as zf:
        for uploaded_file in uploaded_files:
            ext = uploaded_file.name.split(".")[-1].lower()
            file_data = uploaded_file.read()  # Читаємо байти один раз

            with st.spinner(f"Обробка {uploaded_file.name}..."):
                if ext == "xlsx":
                    processed = handle_excel(file_data, engine_analyzer, session_mapping)
                elif ext == "docx":
                    processed = handle_docx(file_data, engine_analyzer, session_mapping)
                elif ext == "pdf":
                    processed = handle_pdf(file_data, engine_analyzer, session_mapping)
                else:
                    continue

                zf.writestr(f"masked_{uploaded_file.name}", processed)

        # Створюємо звіт
        if session_mapping:
            report_df = pd.DataFrame(list(session_mapping.items()), columns=["Оригінальний текст", "Токен"])
            report_buf = io.BytesIO()
            report_df.to_excel(report_buf, index=False)
            zf.writestr("token_mapping_report.xlsx", report_buf.getvalue())

    st.success("Готово!")
    st.download_button(
        label="📥 Завантажити архів з результатом",
        data=zip_buffer.getvalue(),
        file_name="processed_documents.zip",
        mime="application/zip"
    )