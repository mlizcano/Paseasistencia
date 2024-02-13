from flask import Flask, render_template_string, request
import pandas as pd

# Define la función zip_lists aquí
def zip_lists(*args):
    return zip(*args)

app = Flask(__name__)

# Añadir la función zip al entorno global de Jinja2
app.jinja_env.globals.update(zip=zip)

@app.route('/', methods=['GET', 'POST'])
def index():
    archivo_excel = 'Clases2024A.xlsx'
    xls = pd.ExcelFile(archivo_excel)
    cursos = xls.sheet_names
    curso_seleccionado = request.form.get('curso') if request.method == 'POST' else cursos[0]

    # Leer todas las hojas del archivo Excel
    dfs = pd.read_excel(archivo_excel, sheet_name=None)

    if 'fecha' in request.form:
        fecha = request.form['fecha']
        asistencias = request.form.getlist('asistencia')
        df_curso = dfs[curso_seleccionado]
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

    #return render_template_string(html, cursos=cursos, nombres=nombres, curso_seleccionado=curso_seleccionado)
    return render_template_string(html, cursos=cursos, nombres=nombres, codigos=codigos, tutores=tutores)
    #return render_template_string(html, cursos=cursos, nombres=nombres, codigos=codigos, tutores=tutores, zip_lists=zip_lists)



html = """
<!DOCTYPE html>
<html>
<head>
    <title>Registro de Asistencia</title>
    <style>
        body {
            font-family: Arial, sans-serif;
        }
        table {
            width: 100%;
            border-collapse: collapse;
        }
        th, td {
            border: 1px solid #ddd;
            padding: 8px;
            text-align: left;
        }
        th {
            background-color: #4CAF50;
            color: white;
        }
        tr:nth-child(even) {
            background-color: #f2f2f2;
        }
        tr:hover {
            background-color: #ddd;
        }
        input[type="submit"] {
            background-color: #4CAF50;
            color: white;
            padding: 10px 20px;
            border: none;
            border-radius: 4px;
            cursor: pointer;
        }
        input[type="submit"]:hover {
            background-color: #45a049;
        }
    </style>
</head>
<body>
    <h2>Registro de Asistencias</h2>
    <form method="post">
        <label for="curso">Seleccione el curso:</label>
        <select name="curso" required>
            {% for curso in cursos %}
            <option value="{{ curso }}" {% if curso == curso_seleccionado %}selected{% endif %}>{{ curso }}</option>
            {% endfor %}
        </select>
        <input type="submit" value="Seleccionar Curso">
    </form>

    {% if nombres %}
    <form method="post">
        <input type="hidden" name="curso" value="{{ curso_seleccionado }}">
        <label for="fecha">Seleccione la fecha:</label>
        <input type="date" id="fecha" name="fecha" required><br><br>
        <table>
        <table>
            <tr>
                <th>Tutor</th>
                <th>Código</th>
                <th>Nombre</th>
                <th>Asistencia</th>
            </tr>
            {% for i in range(nombres|length) %}
            <tr>
                <td>{{ tutores[i] }}</td>
                <td>{{ codigos[i] }}</td>
                <td>{{ nombres[i] }}</td>
                <td>
                    <select name="asistencia">
                        <option value="A">Asistencia</option>
                        <option value="F">Falta</option>
                        <option value="R">Retardo</option>
                        <option value="J">Justificada</option>
                    </select>
                </td>
            </tr>
            {% endfor %}
        </table>

        <br>
        <input type="submit" value="Guardar Asistencias">
    </form>
    {% endif %}
</body>
</html>
"""

if __name__ == '__main__':
    app.run(debug=True, port=56440)
