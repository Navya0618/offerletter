from flask import Flask, render_template, request, send_file
from docxtpl import DocxTemplate
import os
import uuid
import datetime

app = Flask(__name__)

def format_date_with_suffix(date_str):
    date_obj = datetime.datetime.strptime(date_str, "%Y-%m-%d")
    day = int(date_obj.strftime("%d"))
    suffix = 'th' if 11 <= day <= 13 else {1: 'st', 2: 'nd', 3: 'rd'}.get(day % 10, 'th')
    return f"{day}{suffix} {date_obj.strftime('%B, %Y')}"

@app.route('/', methods=['GET', 'POST'])
def form():
    if request.method == 'POST':
        name = request.form['name']
        role = request.form['role']
        email = request.form['email']
        start_date_raw = request.form['start_date']
        end_date_raw = request.form['end_date']
        letter_date_raw = request.form['letter_date']
        template_file = request.form['template']

        doc = DocxTemplate(f"templates/word_templates/{template_file}")

        # Format dates
        start_date = format_date_with_suffix(start_date_raw)
        end_date = format_date_with_suffix(end_date_raw)
        letter_date = format_date_with_suffix(letter_date_raw)

        context = {
            'name': name,
            'role': role,
            'email': email,
            'start_date': start_date,
            'end_date': end_date,
            'letter_date': letter_date
        }

        # Sanitize name for filename
        safe_name = name.replace(" ", "_").lower()
        safe_role = role.replace(" ", "_").lower()
        output_filename = f"offer_{safe_name}_{safe_role}_{uuid.uuid4().hex[:6]}.docx"
        output_path = os.path.join("generated_letters", output_filename)

        os.makedirs("generated_letters", exist_ok=True)

        doc.render(context)
        doc.save(output_path)

        return send_file(output_path, as_attachment=True)

    return render_template("form.html")

if __name__ == '__main__':
    app.run(debug=True)
