
from zipfile import Path
from flask import Flask, render_template , request, flash , url_for, redirect
import os
from flask.globals import current_app
from flask.helpers import send_file, send_from_directory
from werkzeug.utils import secure_filename
from os import remove
from formato_excel import excelLibro




#iicializar y def archivo principal
#
app=Flask( __name__)
app.secret_key = b'_5#y2L"F4Q8z\n\xec]/'
app.config['ACTUALIZAR_ARCHIVO']="./archivo/"
#creando ruta de nav
@app.route('/')
def home():
    return render_template('inicio.html')

@app.route('/login',methods=['POST'])
def login():
    #obtener info
    error='Credenciales invalidas'
    nombreU = request.form['nombreU']
    contrasena = request.form['contrasena']
    if nombreU == 'admin'and contrasena == 'ser123':
        print(f'sddasdds')
        folder = './archivo/' 
        for the_file in os.listdir(folder): 
            file_path = os.path.join(folder, the_file)  
            if os.path.isfile(file_path) and file_path != folder+'plantilla.xlsx' : 
                    print(f'{file_path}#####')
                    os.unlink(file_path) 

        return redirect(url_for('control'))     
    return render_template("inicio.html",error=error)        
    
@app.route('/control')
def control():
    return render_template('panel.html')


@app.route('/formato', methods=['POST'])
def formato():
        #obtener info
    error='Credenciales invalidas'  
    numMan = request.form['numMan']
    file = request.files['file']
    

    if request.method == 'POST' and numMan != '' and file != '':
        sf=secure_filename(file.filename)
        
        print(f'ruta----{file.filename}')
        ruta = './archivo/'+file.filename 
        if not os.path.exists(ruta):            
            file.save(os.path.join(app.config['ACTUALIZAR_ARCHIVO'],sf)) 
        clase1 = excelLibro(ruta)
        btnDes=clase1.extraer_datos(int(numMan))
        if btnDes: 
            clase1.imprimir()
            print(f'ENTRAND O=================')
            path = clase1.guardar_info("./archivo/plantilla.xlsx")
            path = clase1.crear_pdf(path)
            return send_file(path, as_attachment=True)
        else:
            print(f'elemento no exite')
            remove(ruta)
            error='No se encotro el numero de manifiesto'
            
                        
        
          
    return render_template("panel.html",error = error)   
         
@app.route('/download')
def downloadFile ():
    #For windows you need to use drive name [ex: F:/Example.pdf]
    path = "./archivo/xlsx-to-pdf.xlsx"
    return send_file(path, as_attachment=True)


if __name__ == '__main__':
    app.run(debug=True)