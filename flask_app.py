# Importando las bibliotecas necesarias
from flask import Flask, render_template, redirect, url_for

from flask import Flask, render_template_string, request
import pandas as pd

from flask_mail import Mail, Message


# Creación de la aplicación Flask
app = Flask(__name__)

# Configuración de Flask-Mail
app.config['MAIL_SERVER'] = 'smtp.gmail.com'
app.config['MAIL_PORT'] = 587
app.config['MAIL_USERNAME'] = 'miguel.lizcano@academicos.udg.mx'
app.config['MAIL_PASSWORD'] = 'yadz boql yfrk ycuc'
app.config['MAIL_USE_TLS'] = True
app.config['MAIL_USE_SSL'] = False

mail = Mail(app)


# Ruta para el inicio / página principal
@app.route('/')
def home():
    return render_template('home.html')

# Rutas para la sección de Registro
@app.route('/registro/alumnos')
def registro_alumnos():
    return render_template('registro_alumnos.html')

@app.route('/registro/carrera')
def registro_carrera():
    return render_template('registro_carrera.html')

@app.route('/registro/cursos')
def registro_cursos():
    return render_template('registro_cursos.html')

@app.route('/registro/profesor')
def registro_profesor():
    return render_template('registro_profesor.html')

# Rutas para la sección de Asistencias
@app.route('/asistencias/pasar_lista', methods=['GET', 'POST'])
def asistencias_pasar_lista():
    archivo_excel = 'Clases2024A.xlsx'
    xls = pd.ExcelFile(archivo_excel)
    cursos = xls.sheet_names
    curso_seleccionado = request.form.get('curso') if request.method == 'POST' else cursos[0]

    # Leer todas las hojas del archivo Excel
    dfs = pd.read_excel(archivo_excel, sheet_name=None)

    if 'fecha' in request.form:
        fecha = request.form['fecha']
        asistencias = request.form.getlist('asistencia')
        #df_curso = dfs[curso_seleccionado]
        try:
            df_curso = dfs[curso_seleccionado]
        except KeyError:
            # Manejo del error si curso_seleccionado no es una clave válida en dfs
            return "Curso seleccionado no válido. Por favor, seleccione un curso válido."
        df_curso[fecha] = asistencias

        # Actualizar solo la hoja del curso seleccionado
        dfs[curso_seleccionado] = df_curso

        # Escribir todas las hojas de nuevo en el archivo Excel
        with pd.ExcelWriter(archivo_excel) as writer:
            for sheet_name, df_sheet in dfs.items():
                df_sheet.to_excel(writer, sheet_name=sheet_name, index=False)

        return 'Asistencia guardada con éxito en el curso ' + curso_seleccionado

    df_curso = dfs[curso_seleccionado]
    tutores = df_curso['Tutor'].tolist() if 'Tutor' in df_curso.columns else []
    codigos = df_curso['Codigo'].tolist() if 'Codigo' in df_curso.columns else []
    nombres = df_curso['Nombre'].tolist() if 'Nombre' in df_curso.columns else []


    #return render_template_string(html, cursos=cursos, nombres=nombres, codigos=codigos, tutores=tutores, curso_seleccionado=curso_seleccionado)
    #return render_template('asistencias_pasar_lista.html')
    return render_template('asistencias_pasar_lista.html', cursos=cursos, nombres=nombres, codigos=codigos, tutores=tutores, curso_seleccionado=curso_seleccionado)



#@app.route('/asistencias/justificar_falta')
#def asistencias_justificar_falta():
#    return render_template('asistencias_justificar_falta.html')


@app.route('/asistencias/asistencias_justificar_falta', methods=['GET', 'POST'])
def asistencias_justificar_falta():
    archivo_excel = 'Clases2024A.xlsx'
    xls = pd.ExcelFile(archivo_excel)
    cursos = xls.sheet_names

    if request.method == 'POST':
        curso_seleccionado = request.form['curso']
        df = pd.read_excel(archivo_excel, sheet_name=curso_seleccionado)
        # Calcular el total de faltas y seleccionar columnas

        #df['Total Faltas'] = (df.iloc[:, 4:] == 'Falta').sum(axis=1)
        #columnas_requeridas = ['Tutor', 'Email', 'Nombre', 'CRN', 'Total Faltas']
        #df = df[columnas_requeridas]
        #return render_template('mostrar_asistencia_curso.html', tabla=df.to_html(classes='asistencia'), curso=curso_seleccionado)


        

        # Calcular el total de faltas contando 'Falta' y 'F'
        df['Total Faltas'] = df.iloc[:, 4:].apply(lambda x: (x == 'Falta') | (x == 'F')).sum(axis=1)

        # Seleccionar todas las columnas de asistencia más las columnas requeridas
        columnas_asistencia = df.columns[4:-1]  # Todas las columnas de asistencia
        columnas_requeridas = ['Tutor', 'Email', 'Nombre', 'CRN'] + list(columnas_asistencia) + ['Total Faltas']
        df = df[columnas_requeridas]

        return render_template('mostrar_asistencia_curso.html', tabla=df.to_html(classes='asistencia'), curso=curso_seleccionado)





    return render_template('asistencias_justificar_falta.html', cursos=cursos)




@app.route('/asistencias/enviar_correo', methods=['GET', 'POST'])
def asistencias_enviar_correo():
    archivo_excel = 'clases2024a.xlsx'
    xls = pd.ExcelFile(archivo_excel)
    cursos = xls.sheet_names

    curso_seleccionado = request.form.get('curso') if request.method == 'POST' else cursos[0]
    df_curso = pd.read_excel(archivo_excel, sheet_name=curso_seleccionado)

    # Calcular totales de asistencia, faltas, retardos y justificaciones
    # Suponiendo que las columnas de fechas tienen valores como 'A', 'F', 'R', 'J'
    totales = {
        'Asistencia': (df_curso == 'A').sum(axis=1),
        'Faltas': (df_curso == 'F').sum(axis=1),
        'Retardos': (df_curso == 'R').sum(axis=1),
        'Justificaciones': (df_curso == 'J').sum(axis=1)
    }

    if request.method == 'POST' and 'enviar_correos' in request.form:
        for index, alumno in df_curso.iterrows():
            if request.form.get(f'correo_{alumno["Codigo"]}'):
                # Prepara y envía el correo para este alumno
                # Identifica las fechas de las faltas
                fechas_faltas = ', '.join(df_curso.columns[(df_curso.iloc[index] == 'F')])
                mensaje = f"Estimado/a {alumno['Nombre']},\nTienes faltas en los siguientes días: {fechas_faltas}.\nEn la curso de: { curso_seleccionado } \nPor favor, justifica estas faltas según el reglamento del alumno.\n\nEstimado/a Tutor, se le notifica las inasistencia de su tutorado, para su seguimiento.\n\nAtentamente\n\nDr. Miguel Lizcano Sánchez.\nProfesor del curso de { curso_seleccionado }"

                # Configura el mensaje del correo
                msg = Message('Faltas por Justificar',
                              sender=app.config['MAIL_USERNAME'],
                              recipients=[alumno['Email'], alumno['EmailTutor']])
                msg.body = mensaje

                # Enviar el correo
                mail.send(msg)

        return 'Correos enviados con éxito.'

    #return render_template_string(html, cursos=cursos, datos=df_curso, totales=totales, curso_seleccionado=curso_seleccionado)
    #return render_template('asistencias_enviar_correo.html')
    return render_template('asistencias_enviar_correo.html', cursos=cursos, datos=df_curso, totales=totales, curso_seleccionado=curso_seleccionado)
# Ruta para la sección de Tareas
@app.route('/tareas')
def tareas():
    return render_template('tareas.html')

# Rutas para la sección de Mensaje
@app.route('/mensaje/whatsapp')
def mensaje_whatsapp():
    return render_template('mensaje_whatsapp.html')

@app.route('/mensaje/texto')
def mensaje_texto():
    return render_template('mensaje_texto.html')

# Ruta para Salir
@app.route('/salir')
def salir():
    # Aquí iría la lógica para cerrar la sesión o salir de la aplicación
    return redirect(url_for('home'))


if __name__ == '__main__':
    app.run(debug=True, port=56440)
