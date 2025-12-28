from flask import Flask, send_file
import os
from main_script import create_final_custom_height_replica

app = Flask(__name__)

@app.route('/')
def home():
    return '''
    <div style="text-align: center; margin-top: 50px; font-family: Arial;">
        <h1>Mediation Form Generator</h1>
        <p>Click below to generate the exact replica Word Document.</p>
        <a href="/download">
            <button style="padding: 15px 30px; font-size: 18px; background-color: #007BFF; color: white; border: none; cursor: pointer; border-radius: 5px;">
                Download Form A
            </button>
        </a>
    </div>
    '''

@app.route('/download')
def download_file():
    # Run your script to create the file
    create_final_custom_height_replica()

    # Must match the filename in main_script.py
    filename = "Form_A_Mediation_Replica.docx"

    try:
        return send_file(filename, as_attachment=True)
    except Exception as e:
        return str(e)

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)