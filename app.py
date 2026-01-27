from flask import Flask, render_template, request, send_file, redirect, url_for
from docxtpl import DocxTemplate
import os
import pandas as pd
import zipfile
from datetime import datetime
import io

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['GENERATED_FOLDER'] = 'generated'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['GENERATED_FOLDER'], exist_ok=True)

# Все поля шаблона (для валидации и примера CSV)
TEMPLATE_FIELDS = [
    'college_name', 'commission_name', 'approval_position', 'approval_signature',
    'approval_date', 'module_code', 'module_name', 'specialty_code', 'specialty_name',
    'year', 'fgos_specialty_code', 'fgos_date', 'fgos_order', 'example_program_date',
    'example_program_order', 'study_plan_date', 'pck_protocol_number', 'pck_protocol_date',
    'pck_chair', 'employer_position', 'employer_signature', 'method_council_protocol',
    'developer_name', 'developer_category', 'field_of_study'
]

@app.route('/')
def index():
    return render_template('index.html', fields=TEMPLATE_FIELDS)

@app.route('/single', methods=['GET', 'POST'])
def single():
    if request.method == 'POST':
        context = {field: request.form.get(field, '') for field in TEMPLATE_FIELDS}
        return generate_and_download(context, prefix='program')
    return render_template('single.html', fields=TEMPLATE_FIELDS)

@app.route('/batch', methods=['GET', 'POST'])
def batch():
    if request.method == 'POST':
        if 'file' not in request.files:
            return 'Файл не загружен', 400

        file = request.files['file']
        if file.filename == '':
            return 'Файл не выбран', 400

        # Сохраняем временно
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], f"batch_{datetime.now().strftime('%Y%m%d_%H%M%S')}{os.path.splitext(file.filename)[1]}")
        file.save(filepath)

        # Читаем данные
        try:
            if filepath.endswith('.csv'):
                df = pd.read_csv(filepath)
            else:
                df = pd.read_excel(filepath)
        except Exception as e:
            os.remove(filepath)
            return f'Ошибка чтения файла: {str(e)}', 400

        # Проверяем наличие всех нужных колонок
        missing = set(TEMPLATE_FIELDS) - set(df.columns)
        if missing:
            os.remove(filepath)
            return f'В файле отсутствуют колонки: {", ".join(missing)}', 400

        # Генерируем документы
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for idx, row in df.iterrows():
                context = {field: str(row[field]) if pd.notna(row[field]) else '' for field in TEMPLATE_FIELDS}

                # Генерируем имя файла из ключевых полей
                filename_base = f"{context['module_code'].replace('.', '_')}_{context['specialty_code']}_{idx+1}"
                doc = DocxTemplate('template.docx')
                doc.render(context)

                # Сохраняем во временный буфер
                doc_buffer = io.BytesIO()
                doc.save(doc_buffer)
                doc_buffer.seek(0)

                zipf.writestr(f"{filename_base}.docx", doc_buffer.read())

        os.remove(filepath)
        zip_buffer.seek(0)

        return send_file(
            zip_buffer,
            mimetype='application/zip',
            as_attachment=True,
            download_name=f'batch_programs_{datetime.now().strftime("%Y%m%d_%H%M%S")}.zip'
        )

    return render_template('batch.html', fields=TEMPLATE_FIELDS)

def generate_and_download(context, prefix='program'):
    doc = DocxTemplate('template.docx')
    doc.render(context)

    filename = f"{prefix}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
    filepath = os.path.join(app.config['GENERATED_FOLDER'], filename)
    doc.save(filepath)

    return send_file(filepath, as_attachment=True, download_name=filename)

@app.route('/example-csv')
def example_csv():
    """Генерирует пример CSV для скачивания"""
    example = pd.DataFrame([{field: f"Пример_{field}" for field in TEMPLATE_FIELDS}])
    buffer = io.BytesIO()
    example.to_csv(buffer, index=False, encoding='utf-8-sig')
    buffer.seek(0)
    return send_file(
        buffer,
        mimetype='text/csv',
        as_attachment=True,
        download_name='example_template.csv'
    )

@app.route('/example-xlsx')
def example_xlsx():
    """Генерирует пример XLSX для скачивания"""
    example = pd.DataFrame([{field: f"Пример_{field}" for field in TEMPLATE_FIELDS}])
    buffer = io.BytesIO()
    example.to_excel(buffer, index=False, engine='openpyxl')
    buffer.seek(0)
    return send_file(
        buffer,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='example_template.xlsx'
    )

if __name__ == '__main__':
    app.run(debug=True)
