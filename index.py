
from flask import Flask, render_template, request, url_for, redirect
import os
from flask.helpers import send_file
from werkzeug.utils import secure_filename
from os import remove
from formato_excel import excelLibro


# iicializar y def archivo principal
#
app = Flask(__name__)
app.secret_key = b'_5#y2L"F4Q8z\n\xec]/'
app.config['ACTUALIZAR_ARCHIVO'] = "./archivo/"


def eliminar(lista):
    print(f'eliminando!!')
    folder = './archivo/'
    contador=True
    for the_file in os.listdir(folder):
        file_path = os.path.join(folder, the_file)

        if os.path.isfile(file_path) :  # folder+'plantilla.xlsx'
            for item in lista:
                if file_path == item:
                    contador=False
            if contador == True:
                print(f'{file_path}#####')
                os.unlink(file_path)
            else:
                contador=True       
            

# creando ruta de nav


@app.route('/')
def home():
    return render_template('inicio.html')


@app.route('/login', methods=['POST'])
def login():
    # obtener info
    error = 'Credenciales invalidas'
    nombreU = request.form['nombreU']
    contrasena = request.form['contrasena']
    if nombreU == 'admin' and contrasena == 'ser123':
        print(f'sddasdds')

        return redirect(url_for('control'))
    return render_template("inicio.html", error=error)


@app.route('/control')
def control():
    eliminar(lista=["./archivo/plantilla.xlsx"])
    return render_template('panel.html')


@app.route('/formato', methods=['POST'])
def formato():
    # obtener info
    error = 'Credenciales invalidas'
    numMan = request.form['numMan']
    file = request.files['file']

    if request.method == 'POST' and numMan != '' and file != '':
        sf = secure_filename(file.filename)

        print(f'ruta----{file.filename}')
        ruta = './archivo/'+file.filename
        if not os.path.exists(ruta):
            file.save(os.path.join(app.config['ACTUALIZAR_ARCHIVO'], sf))
        clase1 = excelLibro(ruta)
        btnDes = clase1.extraer_datos(int(numMan))
        if btnDes:
            clase1.imprimir()
            print(f'ENTRANDo++')
            path1 = clase1.guardar_info("./archivo/plantilla.xlsx")
            path = clase1.crear_pdf(path1)
            print(f'val222')
            print(f'{path}/// {ruta}/// {path1}')
            eliminar(lista=[path, ruta, path1, "./archivo/plantilla.xlsx"])
            return send_file(path, as_attachment=True)
        else:
            print(f'elemento no exite')
            remove(ruta)
            error = 'No se encotro el numero de manifiesto'

    return render_template("panel.html", error=error)


@app.route('/download')
def downloadFile():
    # For windows you need to use drive name [ex: F:/Example.pdf]
    path = "./archivo/xlsx-to-pdf.xlsx"
    return send_file(path, as_attachment=True)


#if __name__ == '__main__':
 #   app.run(debug=True)
