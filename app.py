from flask import Flask, render_template, request, send_file
import pandas as pd
import os
import shutil

app = Flask(__name__)

# Path to the uploads folder
UPLOAD_FOLDER = 'uploads'

# Home Page Route
@app.route('/')
def home():
    return render_template('home.html')


# Upload Page Route
@app.route('/upload', methods=['GET', 'POST'])
def upload():
    if request.method == 'POST':
        if 'file' not in request.files:
            return 'No file part'

        file = request.files['file']

        if file.filename == '':
            return 'No selected file'

        # Create uploads folder if it doesn't exist
        if not os.path.exists(UPLOAD_FOLDER):
            os.makedirs(UPLOAD_FOLDER)

        # Save the uploaded Excel file
        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)

        # Convert to CSV with pipe delimiter
        csv_filepath = filepath.rsplit('.', 1)[0] + '.csv'
        df = pd.read_excel(filepath)

        # Save the DataFrame to a CSV file
        df.to_csv(csv_filepath, sep='|', index=False)

        # Send the CSV file to the user
        response = send_file(csv_filepath, as_attachment=True)

        # After sending the file, delete all files in the uploads folder, including the CSV
        clear_uploads_folder(UPLOAD_FOLDER)

        return response

    return render_template('upload.html')


# Function to delete all files in the uploads folder
def clear_uploads_folder(folder):
    for filename in os.listdir(folder):
        file_path = os.path.join(folder, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)  # Remove file or link
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)  # Remove directory
        except Exception as e:
            print(f'Failed to delete {file_path}. Reason: {e}')


if __name__ == '__main__':
    app.run(debug=True)
