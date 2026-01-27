from flask import Flask, render_template, request, send_file
from docxtpl import DocxTemplate
import os
from datetime import datetime

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'generated'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Поля формы (соответствуют плейсхолдерам в шаблоне)
FIELDS = [
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
    ('fgos_specialty_code', 'Код специальности в ФГОС (09.02.01)'),
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
    ('developer_category', 'Категория разработчика (первой/высшей)'),
    ('field_of_study', 'Область техники (вычислительная техника)'),
]

@app.route('/', methods=['GET'])
def index():
    return render_template('form.html', fields=FIELDS)

@app.route('/generate', methods=['POST'])
def generate():
    # Собираем данные из формы
    context = {key: request.form.get(key, '') for key, _ in FIELDS}

    # Генерируем имя файла
    filename = f"program_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)

    # Заполняем шаблон
    doc = DocxTemplate('templates_docx/template.docx')
    doc.render(context)
    doc.save(filepath)

    return send_file(filepath, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
