from flask import Flask, render_template, request, send_file, jsonify, session
import subprocess
import os
import logging
import uuid  # for generating unique identifiers
import pandas as pd # for importing xlsx file

app = Flask(__name__)
app.secret_key = os.urandom(24)  # Secret key for session

# Configure logging
logging.basicConfig(filename='app.log', level=logging.INFO)

# Define upload status constants
UPLOAD_SUCCESS = "File uploaded and processed successfully"
NO_FILE_SELECTED = "No file selected"


@app.route("/")
def index():
    return render_template('index.html', upload_status="")



@app.route('/upload', methods=['POST'])
def upload():
    try:
        if 'file' not in request.files:
            return jsonify({'No file selected'}), 400

        file = request.files['file']
        if file.filename == '':
            return jsonify({'No file selected'}), 400
        if not file.filename.endswith('.xlsx') and not file.filename.endswith('.xls'):
            return jsonify({'File must be in Excel format (.xlsx or .xls)'}), 400

        # Generate a unique identifier
        unique_identifier = str(uuid.uuid4())[:8]  # Adjust the length of the unique code as needed

        # Get the filename and extension
        filename, extension = os.path.splitext(file.filename)

        # Append the unique identifier to the filename
        filename_with_identifier = f"{filename}_{unique_identifier}{extension}"

        # Save the file with the new filename
        file_path = os.path.join('uploads', filename_with_identifier)
        file.save(file_path)
        logging.info("Excel file saved successfully.")

        # Convert Excel to CSV
        excel_df = pd.read_excel(file_path)
        csv_filename = f"{filename}_{unique_identifier}.csv"
        csv_file_path = os.path.join('uploads', csv_filename)
        excel_df.to_csv(csv_file_path, index=False)
        logging.info("CSV file saved successfully.")

        # Store the file path in session
        session['file_path'] = csv_file_path
        logging.info(f"Uploaded filepath: {csv_file_path}")
        

        # Construct the filename without the extension
        filename_without_extension = os.path.splitext(filename_with_identifier)[0]
        logging.info(f"Filename without extension: {filename_without_extension}")

        # Call the script with the uploaded file as an argument
        subprocess.run(['python3', 'RunAll.py', csv_file_path])
        logging.info("RunAll.py is being executed.")

        return jsonify({'message': 'Upload successful'}), 200
    except Exception as e:
        # Log any exceptions that occur
        logging.error(f'An error occurred: {str(e)}')
        return jsonify({'error': 'An error occurred while processing the file'}), 500
        
        
@app.route('/download_all_files')
def download_all_files():
    try:
        # Get the file path from session
        file_path = session.get('file_path')
        
        if file_path is None:
            logging.error("File path is not found in session.")
            # Log session contents for debugging
            logging.info(f"Session contents: {session}")
            return jsonify({'error': 'File path not found in session.'}), 404

        logging.info(f"file path from session: {file_path}")

        # Specify the path to the PowerPoint file
        filename_with_identifier = os.path.basename(file_path)
        filename_without_extension = os.path.splitext(filename_with_identifier)[0]
        combined_file = os.path.join('uploads', f"{filename_without_extension}_combined.pptx")

        if os.path.exists(combined_file):
            # Return the file as an attachment
            logging.info("Combined file found and will be downloaded.")
            return send_file(combined_file, as_attachment=True)
        else:
            logging.error("Combined file not found.")
            return jsonify({'error': 'Combined file not found.'}), 404

    except Exception as e:
        # Handle exceptions
        logging.error(f'An error occurred while downloading all files: {str(e)}')
        return jsonify({'error': 'An error occurred while downloading all files.'}), 500

if __name__ == "__main__":
    app.run(debug=True, use_reloader=False, threaded=False)
