from flask import Flask, request, render_template, send_file
import requests
from bs4 import BeautifulSoup
import pandas as pd
from collections import defaultdict
import os
from io import BytesIO

app = Flask(__name__)


def generate_image_ids(last_number):
    image_ids = []

    for i in range(1, min(last_number + 1, 100)):
        image_ids.append(f"{i:02d}".upper())

    max_letter_index = (last_number - 99) // 10
    for i in range(max_letter_index + 1):
        letter = chr(ord('A') + i)
        for j in range(1, 10):
            if len(image_ids) >= last_number + 1:
                break
            image_ids.append(f"{letter}{j}".upper())

    for i in range(100, last_number + 1):
        image_ids.append(f"{i:02d}".upper())

    return image_ids


def fetch_student_data(student_id, examid):
    url = f'http://www.srkrexams.in/Result.aspx?Id={examid}'
    session = requests.Session()
    response = session.get(url)

    soup = BeautifulSoup(response.content, 'html.parser')

    viewstate = soup.find('input', {'name': '__VIEWSTATE'}).get('value')
    eventvalidation = soup.find('input', {'name': '__EVENTVALIDATION'}).get('value')

    payload = {
        '__VIEWSTATE': viewstate,
        '__EVENTVALIDATION': eventvalidation,
        'ctl00$ContentPlaceHolder1$txtRegNo': student_id,
        'ctl00$ContentPlaceHolder1$btnSubmit': 'Submit'
    }

    headers = {
        'Content-Type': 'application/x-www-form-urlencoded'
    }

    response = session.post(url, data=payload, headers=headers)

    soup = BeautifulSoup(response.content, 'html.parser')

    marks_table = soup.find('table', id='ContentPlaceHolder1_dgvStudentHistory')
    data = []
    student_name = soup.find('input', {'id': 'ContentPlaceHolder1_txtStudentName'}).get('value', '').strip()
    if marks_table:
        rows = marks_table.find_all('tr')[1:]

        for row in rows:
            columns = row.find_all('td')
            if columns:
                data.append([student_id] + [column.text.strip() for column in columns])

    sgpa_cgpa_table = soup.find('table', id='ContentPlaceHolder1_gvSGPA_CGPA')
    sgpa_cgpa = []

    if sgpa_cgpa_table:
        sgpa_cgpa_row = sgpa_cgpa_table.find_all('tr')[1]
        columns = sgpa_cgpa_row.find_all('td')
        if columns:
            sgpa_cgpa = [col.text.strip() for col in columns]

    return data, sgpa_cgpa, student_name


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files['file']
        exam_id = request.form['exam_id']
        year = request.form['year']
        branch_code = request.form['branch_code']
        last_number = int(request.form['last_number'] or 0)

        if file:
            df = pd.read_excel(file)
            print(df)
            student_ids = df['Student ID'].tolist()
        else:
            image_ids = generate_image_ids(last_number)
            student_ids = [f"{year}B91A{branch_code}{image_id}" for image_id in image_ids]

        all_data = defaultdict(dict)
        sgpa_cgpa_data = {}
        subjects = set()
        student_names = {}

        for student_id in student_ids:
            marks_data, sgpa_cgpa, student_name = fetch_student_data(student_id, exam_id)
            student_names[student_id] = student_name

            for row in marks_data:
                subject = row[2]
                grade = row[4]
                all_data[student_id][subject] = grade
                subjects.add(subject)

            if sgpa_cgpa:
                sgpa_cgpa_data[student_id] = sgpa_cgpa

        subjects_list = sorted(subjects)
        marks_headers = ['Student ID', 'Student Name'] + subjects_list + ['SGPA', 'CGPA']

        consolidated_data = []
        for student_id, subject_grades in all_data.items():
            row = [student_id, student_names.get(student_id, '')]
            for subject in subjects_list:
                row.append(subject_grades.get(subject, ''))

            sgpa_cgpa = sgpa_cgpa_data.get(student_id, [''] * 6)
            sgpa = sgpa_cgpa[1] if len(sgpa_cgpa) > 1 else ''
            cgpa = sgpa_cgpa[2] if len(sgpa_cgpa) > 2 else ''

            row += [sgpa, cgpa]
            consolidated_data.append(row)

        marks_df = pd.DataFrame(consolidated_data, columns=marks_headers)

        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            marks_df.to_excel(writer, sheet_name='Marks Data', index=False)

        output.seek(0)
        return send_file(output, download_name=f'student_results_{branch_code}_{year}_{last_number}.xlsx', as_attachment=True)

    return render_template('index.html')


if __name__ == "__main__":
    app.run(debug=True)
