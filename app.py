from flask import Flask, request, render_template, send_file
import pandas as pd
from io import BytesIO
import os

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('form.html')

@app.route('/submit', methods=['POST'])
def submit():
    # Get form data
    name = request.form['name']
    email = request.form['email']
    age = request.form['age']

    # Load existing data if file exists, else create a new DataFrame
    file_path = 'Users.xlsx'
    if os.path.exists(file_path):
        df = pd.read_excel(file_path)
    else:
        df = pd.DataFrame(columns=['Name', 'Email', 'Age'])

    # Append the new data
    new_data = pd.DataFrame([[name, email, age]], columns=['Name', 'Email', 'Age'])
    df = pd.concat([df, new_data], ignore_index=True)

    # Save the updated DataFrame back to the same Excel file
    df.to_excel(file_path, index=False)

    # Optional: Return the updated file for download
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    output.seek(0)

    return send_file(output, download_name="Users.xlsx", as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
