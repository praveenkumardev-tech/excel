from flask import Flask, render_template, request, redirect, send_file
import pandas as pd
import random
import io
import os

# Initialize Flask app
app = Flask(__name__)

# Home page to upload file and enter column names
@app.route('/')
def upload_file():
    return render_template('index.html')

# Process the uploaded file
@app.route('/process', methods=['POST'])
def process_file():
    if 'file' not in request.files:
        return "No file uploaded!", 400

    file = request.files['file']
    if file.filename == '':
        return "No file selected!", 400

    # Read the uploaded Excel file
    df = pd.read_excel(file)

    # Get column names from form
    colm1 = request.form['colm1']
    colm2 = request.form['colm2']

    # Check if columns exist
    if colm1 not in df.columns or colm2 not in df.columns:
        return "Invalid column names!", 400

    # Extract total marks and student names
    total_mark = df[colm1].tolist()
    st_name = df[colm2].tolist()

    # Initialize rows and columns
    row = []

    for i in range(len(total_mark)):
        # Calculate Part A and Part B marks
        if 0 <= total_mark[i] <= 10:
            partA = total_mark[i]
            partB = 0
        elif 11 <= total_mark[i] <= 30:
            partA = random.randint(13, 20)
            partB = total_mark[i] - partA
        else:
            min_val = 20 - (50 - total_mark[i])
            partA = random.randint(min_val, 20)
            partB = total_mark[i] - partA

        # Split Part A marks
        A = []
        x = partA // 2
        y = partA % 2
        z = 10 - (x + y)
        for j in range(x):
            A.append(2)
        for j in range(y):
            A.append(1)
        for j in range(z):
            A.append(0)
        random.shuffle(A)

        # Split Part B marks for 3 questions
        B = []
        a = partB // 6
        b = partB % 6
        for k in range(b):
            B.append(a + 1)
        for k in range(6 - b):
            B.append(a)
        random.shuffle(B)

        # Combine data for each row
        t_mark = [total_mark[i]]
        s_name = [st_name[i]]
        split_marks = s_name + A + B + t_mark
        row.append(split_marks)

    # Define column names
    ques1 = ['Roll no', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10']
    ques2 = ['11.A', '11.B', '12.A', '12.B', '13.A', '13.B', 'Total mark']
    col = ques1 + ques2

    # Create a DataFrame
    df_result = pd.DataFrame(row, columns=col)

    # Save to a BytesIO object
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df_result.to_excel(writer, index=False, sheet_name='Results')
    writer.close()
    output.seek(0)

    # Return the file for download
    return send_file(output, as_attachment=True, download_name="Result.xlsx", mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

# Run the Flask app
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)


