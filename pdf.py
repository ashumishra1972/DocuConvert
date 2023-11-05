import sys
from flask import Flask, render_template, request, send_file, url_for
import tempfile
import pdf2docx
import subprocess

app = Flask(__name__, static_folder='static')

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/pdf-to-word')
def pdftoword():    
    return render_template('pdf-to-word.html')

@app.route('/word-to-pdf')
def wordtopdf():
    return render_template('word-to-pdf.html')

@app.route('/jpeg-to-pdf')
def jpegtopdf():
    return render_template('jpeg-to-pdf.html')

@app.route('/pptx-to-pdf')
def pptxtopdf():
    return render_template('pptx-to-pdf.html')

@app.route('/excel-to-pdf')
def exceltopdf():   
    return render_template('excel-to-pdf.html')

@app.route('/compress-pdf')
def compresspdf():  
    return render_template('compress-pdf.html')
    

# @app.route('/convertword', methods=['POST'])
# def convertwordtopdf():
#     file = request.files['file']
#     if not file.filename.endswith(('.doc', '.docx')):
#         return "Invalid file format. Please upload a .doc or .docx file."
    
#     sys.stderr.write("Starting conversion for " + file.filename + "\n")
#     word_path = tempfile.NamedTemporaryFile(suffix='.docx').name
#     file.save(word_path)
#     pdf_path = f"{word_path}.pdf"
#     command = ['unoconv', '-f', 'pdf', word_path]
#     subprocess.run(command)
    
#     return send_file(pdf_path, as_attachment=True)

@app.route('/convertpdf', methods=['POST'])
def convert():
    # Get the uploaded file and save it to disk
    file = request.files['file']
    if not file.filename.endswith('.pdf'):
        return "Invalid file format. Please upload a .pdf file."

    sys.stderr.write("Starting conversion for " + file.filename + "\n")

    pdf_path = tempfile.NamedTemporaryFile().name
    file.save(pdf_path)

    # Convert the PDF file to a Word document
    docx_path = f"{pdf_path}.docx"
    pdf2docx.parse(pdf_path, docx_path)

    # Return the converted Word document to the user
    return send_file(docx_path, as_attachment=True)




@app.route('/convertword', methods=['POST'])
def convertwordtopdf():
    file = request.files['file']
    if not file.filename.endswith(('.doc', '.docx')):
        return "Invalid file format. Please upload a .doc or .docx file."

    sys.stderr.write("Starting conversion for " + file.filename + "\n")
    word_path = tempfile.NamedTemporaryFile(suffix='.docx').name
    pxssssdf_path = tempfile.NamedTemporaryFile(suffix='.pdf').name
    file.save(word_path)
    # Convert Word to PDF using unoconv
    command = ['unoconv', '-f', 'pdf', '-o', pdf_path, word_path]
    subprocess.run(command, check=True)
    return send_file(pdf_path, as_attachment=True)

@app.route('/convertpptx',methods=['POST'])
def convertpptxtopdf():
    file = request.files['file']
    if not file.filename.endswith(('.pptx','.ppt')):
        return "Invalid file format. Please upload a .pptx or .ppt file."
    
    sys.stderr.write("Starting conversion for " + file.filename + "\n")
    ppt_path=tempfile.NamedTemporaryFile(suffix='.pptx').name
    pdf_path=tempfile.NamedTemporaryFile(suffix='.pdf').name
    file.save(ppt_path)
    command = ['unoconv', '-f', 'pdf', '-o', pdf_path, ppt_path]
    subprocess.run(command, check=True)
    return send_file(pdf_path, as_attachment=True)


@app.route('/convertjpeg', methods=['POST'])
def convertjpeg():
    file=request.files['file']
    if not file.filename.endswith(('.jpeg','.jpg')):
        return "Invalid file format. Please upload a .jpeg or .jpg Image File ."
    
    jpeg_path=tempfile.NamedTemporaryFile(suffix='.jpeg').name
    pdf_path=tempfile.NamedTemporaryFile(suffix=".pdf").name
    file.save(jpeg_path)
    command = ['unoconv', '-f','pdf','-o', pdf_path,jpeg_path]
    subprocess.run(command,check=True)
    return send_file(pdf_path,as_attachment=True)

@app.route('/convertexcel', methods=['POST'])
def convertexceltopdf():
    file = request.files['file']
    if not file.filename.endswith(('.xls', '.xlsx')):
        return "Invalid file format. Please upload a .xls or .xlsx file."

    sys.stderr.write("Starting conversion for " + file.filename + "\n")
    excel_path = tempfile.NamedTemporaryFile(suffix='.xls').name
    pdf_path = tempfile.NamedTemporaryFile(suffix='.pdf').name
    file.save(excel_path)
    # Convert excel to PDF using unoconv
    command = ['unoconv', '-f', 'pdf', '-o', pdf_path, excel_path]
    subprocess.run(command, check=True)
    return send_file(pdf_path, as_attachment=True)

@app.route('/compresspdf', methods=['POST'])
def compresspdfs():  
    file = request.files['file']
    if not file.filename.endswith('.pdf'):
        return "Invalid file format. Please upload a .pdf file."

    sys.stderr.write("Starting conversion for " + file.filename + "\n")
    input_path = tempfile.NamedTemporaryFile(suffix='.pdf   ').name
    output_path = tempfile.NamedTemporaryFile(suffix='.pdf').name
    file.save(input_path)
    # Convert excel to PDF using unoconv
    command = ['unoconv', '-f', 'pdf', '-e', 'CompressionLevel=2', '-o', output_path, input_path]
    subprocess.run(command, check=True)
    return send_file(output_path, as_attachment=True)



@app.route('/robots.txt')
def robots():
    return app.send_static_file('robots.txt')

@app.errorhandler(404)
def not_found(error):
    return render_template('error.html'), 404

if __name__ == '__main__':
    app.run(debug=True    ,threaded=True)