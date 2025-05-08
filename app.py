from flask import Flask, render_template, request, send_file
from docxtpl import DocxTemplate
import os
import uuid

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def form():
    if request.method == 'POST':
        name = request.form['name']
        email = request.form['email']
        start_date = request.form['start_date']
        end_date = request.form['end_date']
        letter_date = request.form['letter_date']
        template_file = request.form['template']

        doc = DocxTemplate(f"templates/word_templates/{template_file}")

        context = {
            'name': name,
            'email': email,
            'start_date': start_date,
            'end_date': end_date,
            'letter_date': letter_date
        }

        output_filename = f"generated_offer_{uuid.uuid4().hex}.docx"
        output_path = os.path.join("templates", output_filename)
        doc.render(context)
        doc.save(output_path)

        return send_file(output_path, as_attachment=True)

    return render_template("form.html")

if __name__ == '__main__':
    app.run(debug=True)