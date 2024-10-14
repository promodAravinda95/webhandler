from flask import Flask, render_template, request, send_file, flash, redirect
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

@app.route('/validate_plan', methods=['GET', 'POST'])
def validate_plan():
    if not os.path.exists(UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER)
    else:
        clear_uploads_folder(UPLOAD_FOLDER)
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)

        file = request.files['file']

        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)

        # Save the uploaded Excel file
        filepath = os.path.join('uploads', file.filename)
        file.save(filepath)

        # Process the Excel file for validation
        try:
            df = pd.read_excel(filepath)

            # Ensure the PLAN_CREATION_DATE column is in datetime format
            df['PLAN_CREATION_DATE'] = pd.to_datetime(df['PLAN_CREATION_DATE'], errors='coerce')

            # Filter data where PLAN_CREATION_DATE is greater than '2024-09-01' and PLAN_CREATION_TYPE is in specific values
            filtered_df = df[df['PLAN_CREATION_TYPE'].isin(['MP', 'SV', 'SC', 'DM', 'DY'])].copy()

            # Create 'outletyearmonthwid' by concatenating PLAN_CREATION_DATE year-month with USL_OUTLET_CODE
            filtered_df['outletyearmonthwid'] = filtered_df['PLAN_CREATION_DATE'].dt.strftime('%Y-%m') + '_' + filtered_df['USL_OUTLET_CODE'].astype(str)

            # Group by 'outletyearmonthwid' and count PLAN_CREATION_TYPE
            grouped_df = filtered_df.groupby('outletyearmonthwid')['PLAN_CREATION_TYPE'].count().reset_index(name='countPLAN_CREATION_TYPE')

            # Filter groups having count greater than 1
            result_df = grouped_df[grouped_df['countPLAN_CREATION_TYPE'] > 1]

            # Save the result to CSV
            result_csv_filepath = filepath.rsplit('.', 1)[0] + '_validation_result.csv'
            result_df.to_csv(result_csv_filepath, index=False)


            # Send the result file back to the user
            return send_file(result_csv_filepath, as_attachment=True)

        except Exception as e:
            flash(f"Error processing the file: {e}")
            return redirect(request.url)

    return render_template('validate_plan.html')

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
    app.run()
