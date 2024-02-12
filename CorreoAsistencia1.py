import pandas as pd
from flask import Flask, render_template, request, redirect, url_for

# Aquí puedes incluir las funciones y configuraciones necesarias para enviar correos electrónicos

# Inicializar la aplicación Flask
app = Flask(__name__)

# Ruta de inicio para seleccionar un curso
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Obtener el curso seleccionado del formulario
        selected_course = request.form.get('course')
        # Redirigir a la página de asistencia con el curso seleccionado
        return redirect(url_for('attendance', course=selected_course))

    # Listar los cursos disponibles (hojas en el archivo Excel)
    courses = read_excel_and_list_sheets('cursos2024a.xlsx')
    return render_template('index.html', courses=courses)

# Ruta de asistencia para mostrar y registrar asistencia
@app.route('/attendance', methods=['GET', 'POST'])
def attendance():
    selected_course = request.args.get('course')
    attendance_date = None  # Inicializar con un valor predeterminado

    if request.method == 'POST':
        attendance_date = request.form.get('date')
        df = pd.read_excel('cursos2024a.xlsx', sheet_name=selected_course)

        # Procesar los datos de asistencia enviados
        for key, value in request.form.items():
            if key.startswith('attendance_'):
                student_code = key.split('_')[1]
                attendance_status = value

                # Inicializar student_name como una cadena vacía
                student_name = ""

                if df['CodigoAlumno'].isin([student_code]).any():
                    student_name = df.loc[df['CodigoAlumno'] == student_code, 'NombreAlumno'].iloc[0]
                    df.loc[df['CodigoAlumno'] == student_code, attendance_date] = attendance_status

                # Mover la impresión fuera del if para evitar el UnboundLocalError
                print(f"Código: {student_code}, Nombre: {student_name}, Asistencia: {attendance_status}")

        # Guardar los cambios en el archivo Excel
        with pd.ExcelWriter('cursos2024a.xlsx', mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=selected_course, index=False)

        return redirect(url_for('index'))

    # Si es un GET request, simplemente renderiza la plantilla sin usar 'attendance_date'
    else:
        # Leer los datos de los alumnos del curso seleccionado
        df = pd.read_excel('cursos2024a.xlsx', sheet_name=selected_course)
        students = df[['CodigoAlumno', 'NombreAlumno']].to_dict(orient='records')
        return render_template('attendance.html', students=students, course=selected_course)

# Función para leer el archivo Excel y listar las hojas disponibles
def read_excel_and_list_sheets(file_path):
    xls = pd.ExcelFile(file_path)
    return xls.sheet_names

# Ejecutar la aplicación Flask
app.run(debug=True, port=56440)
