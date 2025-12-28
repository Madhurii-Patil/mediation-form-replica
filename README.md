# Python Automation: PDF to Word Replica

## 📝 Project Summary
This project is an automation tool designed to create an exact MS Word (`.docx`) replica of a given PDF file. Instead of just extracting text, I focused on recreating the precise **layout, formatting, and table structure** of the original legal form using Python.

## 🛠️ Tech Stack Used
* **Python 3.x**
* **python-docx:** To build the document structure from scratch.
* **Flask:** To create a simple web interface for generating the file.
* **Gunicorn:** For production deployment on Render.

## 💡 My Approach
The biggest challenge was matching the PDF's strict alignment. Here is how I solved it:

1.  **Grid System:** I realized standard paragraphs wouldn't work for this complex layout. I used a **table-based grid** structure, merging cells where necessary to align text perfectly (Left, Center, Middle).
2.  **Precision Sizing:** I measured the columns and rows in `Inches` to ensure the Word document looks identical to the PDF on A4 paper.
3.  **Custom Borders (XML):** The `python-docx` library has limited border options. I dug into the **XML layer (`OxmlElement`)** to force specific single-line borders on individual cells (e.g., Top/Bottom only) as required by the form.
4.  **Dynamic Ready:** I didn't just hardcode text. I inserted placeholders (like `{{client_name}}`) so this script can be used as a template for generating real client forms in the future.

## 🚀 How to Run Locally
1.  **Clone the repo and install libraries:**
    ```bash
    pip install -r requirements.txt
    ```
2.  **Start the server:**
    ```bash
    python app.py
    ```
3.  **Use the tool:**
    Open `http://127.0.0.1:5000` in your browser and click "Download".

## 🌐 Live Demo
You can test the live application here:
* **Live Link:** [PASTE YOUR RENDER LINK HERE]