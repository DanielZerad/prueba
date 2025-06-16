from flask import Flask, request, render_template, redirect, url_for, flash, send_from_directory
import pandas as pd
import pyodbc


app = Flask(__name__)
app.secret_key = 'your_secret_key'



#Modulo de conexion de datos
def connect_db():
    servidor = '10.109.10.36'
    base = 'Cigarrillos'
    usuario = 'PEOD97AL'
    contrasena = 'Daniel3232'
    
    connection_string = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={servidor};DATABASE={base};UID={usuario};PWD={contrasena}'
    return pyodbc.connect(connection_string)

#Conexion a la ruta index
@app.route('/')
def index():
    return render_template('index.html')

#Conexion a la ruta import
@app.route('/import', methods=['POST'])
def importar_excel():
    file = request.files['file']
    if not file:
        flash('No file selected')
        return redirect(url_for('index'))

    conn = None
    cursor = None
    try:
        df = pd.read_excel(file)
        df.fillna(0, inplace=True)

        conn = connect_db()
        cursor = conn.cursor()

        row_count = 0
        for index, row in df.iterrows():
            if row['Id_Aduana'] != 0 and row['Id_Mes'] != 0:
                cursor.execute(
                    "INSERT INTO td_cigarrillos (Id_Aduana, Id_Mes, Id_Año, NO_PAQ_AS, NO_CAJ_AS, NO_CIG_AS, FEC_ASE_AS, "
                    "NO_PAQ_DES, NO_CAJ_DES, NO_CIG_DES, FEC_ASE_DES, NO_PAQ_PROC, NO_CAJ_PROC, NO_CIG_PROC, "
                    "NO_PAQ_PEND, NO_CAJ_PEND, NO_CIG_PEND) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                    tuple(row))
                row_count += 1
        
        conn.commit()
        flash(f'Datos importados correctamente. Número de registros importados: {row_count}')
    except pyodbc.Error as db_err:
        flash(f"Error de base de datos: {str(db_err)}")
    except Exception as e:
        flash(f"Error: {str(e)}")
    finally:
        if cursor is not None:
            cursor.close()
        if conn is not None:
            conn.close()

    return redirect(url_for('index'))


@app.route('/static/<path:filename>')
def static_files(filename):
    return send_from_directory('static', filename)

#Conexion a la pagina data
@app.route('/data')
def mostrar_datos():
    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM td_cigarrillos") #Seleccion de tablas
    rows = cursor.fetchall()
    cursor.close()
    conn.close()
    return render_template('data.html', rows=rows)

if __name__ == '__main__':
    from waitress import serve
    serve(app, host='10.150.24.13', port=8000)


