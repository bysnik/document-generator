from flask import Flask, render_template, request, send_file
from docxtpl import DocxTemplate
import os
import pandas as pd
import zipfile
from datetime import datetime, timedelta
import io
import uuid

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['GENERATED_FOLDER'] = 'generated'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # Ограничение: 16 МБ
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['GENERATED_FOLDER'], exist_ok=True)

# Все поля шаблона
TEMPLATE_FIELDS = [
    'college_name', 'commission_name', 'approval_position', 'approval_signature',
    'approval_date', 'module_code', 'module_name', 'specialty_code', 'specialty_name',
    'year', 'fgos_specialty_code', 'fgos_date', 'fgos_order', 'example_program_date',
    'example_program_order', 'study_plan_date', 'pck_protocol_number', 'pck_protocol_date',
    'pck_chair', 'employer_position', 'employer_signature', 'method_council_protocol',
    'developer_name', 'developer_category', 'field_of_study'
]

# Очистка старых файлов (каждый раз при запуске)
def cleanup_old_files():
    now = datetime.now()
    for folder in [app.config['UPLOAD_FOLDER'], app.config['GENERATED_FOLDER']]:
        if os.path.exists(folder):
            for fname in os.listdir(folder):
                fpath = os.path.join(folder, fname)
                if os.path.isfile(fpath):
                    mtime = datetime.fromtimestamp(os.path.getmtime(fpath))
                    if (now - mtime) > timedelta(hours=1):
                        try:
                            os.remove(fpath)
                            print(f"🧹 Удалён старый файл: {fname}")
                        except Exception as e:
                            print(f"⚠️ Не удалось удалить {fname}: {e}")

@app.route('/')
def index():
    cleanup_old_files()  # Очищаем старые файлы при открытии главной страницы
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
            return '❌ Файл не загружен', 400

        file = request.files['file']
        if file.filename == '':
            return '❌ Файл не выбран', 400

        # Проверяем расширение
        ext = os.path.splitext(file.filename)[1].lower()
        if ext not in ['.csv', '.xlsx', '.xls']:
            return '❌ Неподдерживаемый формат файла. Используйте CSV или XLSX', 400

        # Генерируем уникальное имя для сохранения
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        unique_id = uuid.uuid4().hex[:8]
        filename = f"batch_{unique_id}_{timestamp}{ext}"
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        
        # Сохраняем файл
        file.save(filepath)

        # Читаем данные
        try:
            if ext == '.csv':
                df = pd.read_csv(filepath, encoding='utf-8-sig')
            else:
                df = pd.read_excel(filepath)
        except Exception as e:
            os.remove(filepath)
            return f'❌ Ошибка чтения файла: {str(e)}<br><br>Проверьте формат файла и кодировку.', 400

        # Проверяем наличие всех нужных колонок
        missing = set(TEMPLATE_FIELDS) - set(df.columns)
        if missing:
            os.remove(filepath)
            available = set(df.columns) - set(TEMPLATE_FIELDS)
            return f'''
            ❌ В файле отсутствуют обязательные колонки:<br>
            <strong>{", ".join(sorted(missing))}</strong><br><br>
            
            Доступные колонки в файле:<br>
            {", ".join(sorted(df.columns))}<br><br>
            
            <a href="/example-csv" style="color:#0066cc;">Скачать пример шаблона CSV</a> | 
            <a href="/example-xlsx" style="color:#0066cc;">Скачать пример шаблона XLSX</a>
            ''', 400

        # Генерируем документы
        zip_buffer = io.BytesIO()
        success_count = 0
        error_count = 0
        
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for idx, row in df.iterrows():
                try:
                    context = {field: str(row[field]) if pd.notna(row[field]) else '' for field in TEMPLATE_FIELDS}

                    # Генерируем имя файла
                    filename_base = f"{context['module_code'].replace('.', '_')}_{context['specialty_code']}_{idx+1}"
                    doc = DocxTemplate('template.docx')
                    doc.render(context)

                    # Сохраняем в архив
                    doc_buffer = io.BytesIO()
                    doc.save(doc_buffer)
                    doc_buffer.seek(0)

                    zipf.writestr(f"{filename_base}.docx", doc_buffer.read())
                    success_count += 1
                except Exception as e:
                    error_count += 1
                    print(f"⚠️ Ошибка при генерации документа {idx+1}: {e}")

        # Удаляем временный файл
        os.remove(filepath)
        
        zip_buffer.seek(0)

        # Формируем имя архива
        archive_name = f'batch_programs_{success_count}docs_{timestamp}.zip'
        
        return send_file(
            zip_buffer,
            mimetype='application/zip',
            as_attachment=True,
            download_name=archive_name
        )

    return render_template('batch.html', fields=TEMPLATE_FIELDS)

def generate_and_download(context, prefix='program'):
    doc = DocxTemplate('template.docx')
    doc.render(context)

    filename = f"{prefix}_{uuid.uuid4().hex}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
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
    print("=" * 60)
    print("🚀 Запуск приложения...")
    print(f"📁 Папка загрузок: {app.config['UPLOAD_FOLDER']}")
    print(f"📁 Папка результатов: {app.config['GENERATED_FOLDER']}")
    print("=" * 60)
    app.run(debug=True)