from flask import Flask, render_template, request, redirect, url_for, send_file
import pdfplumber
import re
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd
import io  # Cambiamos la importación de 'os' a 'io'

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    if 'pdf_file' not in request.files:
        return redirect(request.url)

    pdf_file = request.files['pdf_file']

    if pdf_file.filename == '':
        return redirect(request.url)

    if pdf_file:
        # Abre el archivo PDF
        with pdfplumber.open(pdf_file) as pdf:
            # Variables para almacenar los datos adicionales
            ruc_data = None
            comprobante_data = None
            autorizacion_data = None
            fecha_hora_data = None

            # Itera a través de las páginas del PDF
            for page in pdf.pages:
                # Extrae el texto de la página
                text = page.extract_text()

                # Utiliza expresiones regulares para buscar los datos específicos
                ruc_pattern = r"R\.U\.C\.: (\d{13})"
                comprobante_pattern = r"COMPROBANTE DE RETENCIÓN\nNo\. (\d+-\d+-\d+)"
                autorizacion_pattern = r"NÚMERO DE AUTORIZACIÓN\n(\d+)"
                fecha_hora_pattern = r"FECHA Y HORA DE\n(.+)"

                # Busca los datos en el texto
                ruc_match = re.search(ruc_pattern, text)
                comprobante_match = re.search(comprobante_pattern, text)
                autorizacion_match = re.search(autorizacion_pattern, text)
                fecha_hora_match = re.search(fecha_hora_pattern, text)

                # Almacena los datos encontrados, si están disponibles
                if ruc_match:
                    ruc_data = ruc_match.group(1)
                if comprobante_match:
                    comprobante_data = comprobante_match.group(1)
                if autorizacion_match:
                    autorizacion_data = autorizacion_match.group(1)
                if fecha_hora_match:
                    fecha_hora_data = fecha_hora_match.group(1)

            # Extraer datos de la tabla
            table_pattern = r"(\d{13})\s+FACTURA\s+(\d{2}/\d{2}/\d{4})\s+(\d{2}/\d{4})\s+([\d.]+)\s+([^0-9]+)\s+([\d.]+)\s+([\d.]+)(?:\s+(\d+))?"
            table_matches = re.findall(table_pattern, text)

            # Crear un DataFrame para los datos de la tabla
            data = []
            headers = ["Comprobante", "Fecha de Emisión", "Ejercicio Fiscal", "Base Imponible para la Retención", "Impuesto", "Porcentaje de Retención", "Valor Retenido"]
            for match in table_matches:
                comprobante = match[0] + match[7] if match[7] else match[0]
                fecha_emision = match[1]
                ejercicio_fiscal = match[2]
                base_imponible = match[3]
                impuesto = match[4]
                porcentaje_retencion = match[5]
                valor_retenido = match[6]
                data.append([comprobante, fecha_emision, ejercicio_fiscal, base_imponible, impuesto, porcentaje_retencion, valor_retenido])

            df = pd.DataFrame(data, columns=headers)

            # Guardar el archivo Excel en un objeto en memoria (en lugar de en una ubicación en disco)
            excel_filename = request.form['excel_filename']
            excel_buffer = io.BytesIO()  # Usamos un buffer en memoria para el archivo Excel
            workbook = Workbook()
            sheet = workbook.active

            # Agregar los datos adicionales al archivo Excel
            sheet.append(["R.U.C.", ruc_data])
            sheet.append(["Número de Comprobante de Retención", comprobante_data])
            sheet.append(["Número de Autorización", autorizacion_data])
            sheet.append(["Fecha y Hora de Autorización", fecha_hora_data])
            sheet.append([])  # Agregar una fila en blanco como separación
            for row in dataframe_to_rows(df, index=False, header=True):
                sheet.append(row)

            # Guardar el archivo Excel en el buffer
            workbook.save(excel_buffer)
            excel_buffer.seek(0)

            # Devolver el archivo Excel como respuesta para descarga
            return send_file(excel_buffer, as_attachment=True, download_name=excel_filename + '.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == '__main__':
    app.run(debug=True)

