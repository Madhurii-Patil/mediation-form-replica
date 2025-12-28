from flask import Flask, send_file
import os
from main_script import create_final_custom_height_replica

app = Flask(__name__)

@app.route('/')
def home():
    # 1. Script run karke file generate karo
    create_final_custom_height_replica()
    
    # 2. File ka Path dhoondo (Absolute Path use kar rahe hain taaki error na aaye)
    filename = "Form_A_Mediation_Replica.docx"
    file_path = os.path.join(os.getcwd(), filename)
    
    # 3. Direct Download (Browser mein link kholte hi file milegi)
    try:
        if os.path.exists(file_path):
            return send_file(file_path, as_attachment=True)
        else:
            return f"Error: File generate nahi hui. Path: {file_path}"
    except Exception as e:
        return f"Error: {str(e)}"

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)