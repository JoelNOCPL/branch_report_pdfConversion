from distutils.log import debug
from fileinput import filename
from flask import *
from auto_branchReport import excel2pdf
import datetime
app = Flask(__name__,template_folder='views')  
  
@app.route('/')  
def main():  
    return render_template("index.html",utc_dt=None)#datetime.datetime.utcnow() )  
  
@app.route('/success', methods = ['POST'])  
def success():  
    if request.method == 'POST':
        global f
        f = request.files['file']
        f.save(f.filename)
        #if f.filename[-5:]=='.xlsx':
        return render_template("Acknowledgement.html", name = f.filename)
        #else:
            #return redirect(url_for('index'))
            #return render_template("index.html")  
            






@app.route('/download')
def download_file():
    #path = "html2pdf.pdf"
    #path = "info.xlsx"
    global pdfName
    pdfName =  excel2pdf(f.filename)
    path = pdfName#f.filename #"simple.docx"
    #path = "sample.txt"
    return send_file(path, as_attachment=True)

  
if __name__ == '__main__':  
    app.run(host='localhost', port='5002',debug=True)# 
