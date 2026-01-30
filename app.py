from flask import Flask, render_template, request, send_file, redirect, url_for
from docxtpl import DocxTemplate
import os
import pandas as pd
import zipfile
from datetime import datetime
import io
import uuid

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['GENERATED_FOLDER'] = 'generated'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['GENERATED_FOLDER'], exist_ok=True)

# Основные поля шаблона с человекочитаемыми названиями
TEMPLATE_FIELDS = [
    ('college_name', 'Название колледжа'),
    ('commission_name', 'Название цикловой комиссии'),
    ('approval_position', 'Должность утверждающего'),
    ('approval_signature', 'ФИО утверждающего (подпись)'),
    ('approval_date', 'Дата утверждения'),
    ('module_code', 'Код модуля (ПМ.01)'),
    ('module_name', 'Название модуля'),
    ('specialty_code', 'Код специальности (09.02.06)'),
    ('specialty_name', 'Название специальности'),
    ('year', 'Год разработки программы'),
    ('fgos_specialty_code', 'Код специальности в ФГОС'),
    ('fgos_date', 'Дата приказа ФГОС'),
    ('fgos_order', 'Номер приказа ФГОС'),
    ('example_program_date', 'Дата примерной программы'),
    ('example_program_order', 'Номер приказа примерной программы'),
    ('study_plan_date', 'Дата утверждения учебного плана'),
    ('pck_protocol_number', 'Номер протокола ПЦК'),
    ('pck_protocol_date', 'Дата протокола ПЦК'),
    ('pck_chair', 'Председатель ПЦК (ФИО)'),
    ('employer_position', 'Должность представителя работодателя'),
    ('employer_signature', 'ФИО представителя работодателя'),
    ('method_council_protocol', 'Протокол методического совета'),
    ('developer_name', 'ФИО разработчика'),
    ('developer_category', 'Категория разработчика'),
    ('field_of_study', 'Область техники'),
]

@app.route('/')
def index():
    """Главная страница"""
    return render_template('index.html', fields=TEMPLATE_FIELDS)

@app.route('/single', methods=['GET', 'POST'])
def single():
    """Одиночная генерация"""
    if request.method == 'POST':
        # Собираем данные из формы
        context = {}
        for field, _ in TEMPLATE_FIELDS:
            context[field] = request.form.get(field, '')
        
        # Генерируем и сохраняем документ
        filename = f"program_{uuid.uuid4().hex}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
        filepath = os.path.join(app.config['GENERATED_FOLDER'], filename)
        
        doc = DocxTemplate('template.docx')
        doc.render(context)
        doc.save(filepath)
        
        # Перенаправляем на страницу результата
        return redirect(url_for('single_result', filename=filename))
    
    return render_template('single.html', fields=TEMPLATE_FIELDS)

@app.route('/single/result/<filename>')
def single_result(filename):
    """Страница результата генерации"""
    return render_template('single_result.html', filename=filename)

@app.route('/single/download/<filename>')
def single_download(filename):
    """Скачивание сгенерированного документа"""
    filepath = os.path.join(app.config['GENERATED_FOLDER'], filename)
    if os.path.exists(filepath):
        return send_file(filepath, as_attachment=True, download_name=filename)
    return 'Файл не найден', 404

@app.route('/batch', methods=['GET', 'POST'])
def batch():
    """Пакетная генерация"""
    if request.method == 'POST':
        if 'file' not in request.files:
            return '❌ Файл не загружен', 400

        file = request.files['file']
        if file.filename == '':
            return '❌ Файл не выбран', 400

        ext = os.path.splitext(file.filename)[1].lower()
        if ext not in ['.csv', '.xlsx', '.xls']:
            return '❌ Неподдерживаемый формат файла. Используйте CSV или XLSX', 400

        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        unique_id = uuid.uuid4().hex[:8]
        filename = f"batch_{unique_id}_{timestamp}{ext}"
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        
        file.save(filepath)

        try:
            if ext == '.csv':
                df = pd.read_csv(filepath, encoding='utf-8-sig')
            else:
                df = pd.read_excel(filepath)
        except Exception as e:
            os.remove(filepath)
            return f'❌ Ошибка чтения файла: {str(e)}<br><br>Проверьте формат файла и кодировку.', 400

        # Извлекаем только технические названия полей
        field_names = [field for field, _ in TEMPLATE_FIELDS]
        
        missing = set(field_names) - set(df.columns)
        if missing:
            os.remove(filepath)
            return f'''
            ❌ В файле отсутствуют обязательные колонки:<br>
            <strong>{", ".join(sorted(missing))}</strong><br><br>
            
            Доступные колонки в файле:<br>
            {", ".join(sorted(df.columns))}<br><br>
            
            <a href="/example-csv" style="color:#0066cc;">Скачать пример шаблона CSV</a> | 
            <a href="/example-xlsx" style="color:#0066cc;">Скачать пример шаблона XLSX</a>
            ''', 400

        # Генерируем имя архива
        archive_name = f'batch_programs_{len(df)}docs_{timestamp}.zip'
        archive_path = os.path.join(app.config['GENERATED_FOLDER'], archive_name)
        
        success_count = 0
        
        # Создаём архив на диске
        with zipfile.ZipFile(archive_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for idx, row in df.iterrows():
                try:
                    context = {field: str(row[field]) if pd.notna(row[field]) else '' for field in field_names}

                    filename_base = f"{context['module_code'].replace('.', '_')}_{context['specialty_code']}_{idx+1}"
                    doc = DocxTemplate('template.docx')
                    doc.render(context)

                    doc_buffer = io.BytesIO()
                    doc.save(doc_buffer)
                    doc_buffer.seek(0)

                    zipf.writestr(f"{filename_base}.docx", doc_buffer.read())
                    success_count += 1
                except Exception as e:
                    print(f"⚠️ Ошибка при генерации документа {idx+1}: {e}")

        os.remove(filepath)
        
        # Отправляем архив пользователю
        return send_file(
            archive_path,
            mimetype='application/zip',
            as_attachment=True,
            download_name=archive_name
        )

    return render_template('batch.html', fields=TEMPLATE_FIELDS)

@app.route('/example-csv')
def example_csv():
    """Генерирует пример CSV для скачивания"""
    # Извлекаем только технические названия полей
    field_names = [field for field, _ in TEMPLATE_FIELDS]
    example = pd.DataFrame([{field: f"Пример_{field}" for field in field_names}])
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
    # Извлекаем только технические названия полей
    field_names = [field for field, _ in TEMPLATE_FIELDS]
    example = pd.DataFrame([{field: f"Пример_{field}" for field in field_names}])
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
    print("💾 Все сгенерированные файлы сохраняются на сервере")
    print("=" * 60)
    app.run(debug=True)