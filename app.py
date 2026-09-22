from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import pandas as pd
import io
from openpyxl import load_workbook

app = Flask(__name__)
CORS(app, expose_headers=["Content-Disposition"])

def limpiar_moneda(x):
    if pd.isna(x): return 0.0
    if isinstance(x, str):
        x = x.replace('.', '').replace(',', '.')
    return float(x)

@app.route('/procesar-almaviva', methods=['POST'])
def procesar_archivo():
    try: # 🚨 Envolvemos TODO en un Try para atrapar el error real y no desconectarnos
        if 'file' not in request.files:
            return jsonify({"error": "No se subió ningún archivo"}), 400
        
        file = request.files['file']
        tipo_reporte = request.form.get('tipo_reporte', 'admin')
        file_content = file.read()
        
        # --- LECTOR SÚPER INTELIGENTE ---
        # utf-8-sig es el formato exacto que usa Excel al guardar CSV.
        # sep=None hace que Python detecte automáticamente el separador.
        try:
            df = pd.read_csv(io.BytesIO(file_content), sep=None, engine='python', encoding='utf-8-sig')
        except Exception:
            df = pd.read_csv(io.BytesIO(file_content), sep=None, engine='python', encoding='latin1')
        
        # Verificamos que las columnas clave existan (a veces Excel les pone espacios sin querer)
        df.columns = [str(col).strip() for col in df.columns]

        # Validar que todas las columnas maestras existan antes de hacer cálculos
        if tipo_reporte == 'admin':
            columnas_esperadas = ['ID de reserva', 'Hora de Pickup (Local)', 'Recoger', 'Destino', 'Nombre', 'Precio', 'Costo', 'Vehículo', 'Programa de viaje']
        else:
            columnas_esperadas = ['ID de reserva', 'Hora de Pickup (Local)', 'Recoger', 'Destino', 'Nombre', 'Precio', 'Programa de viaje']
            
        columnas_faltantes = [col for col in columnas_esperadas if col not in df.columns]
        if columnas_faltantes:
            return jsonify({"error": f"Faltan estas columnas exactas en tu Excel (revisa los títulos): {', '.join(columnas_faltantes)}"}), 400

        # --- FILTRO DE AUDITORÍA ESTRICTO ---
        vehiculos_str = df.get('Vehículo', pd.Series(dtype=str)).astype(str).str.strip().str.lower()
        programas_str = df.get('Programa de viaje', pd.Series(dtype=str)).astype(str).str.strip().str.lower()
        
        sin_vehiculo = df[(df.get('Vehículo').isna()) | (vehiculos_str == '') | (vehiculos_str == 'nan') | (vehiculos_str == 'null')]
        sin_programa = df[(df.get('Programa de viaje').isna()) | (programas_str == '') | (programas_str == 'nan') | (programas_str == 'null')]
        
        errores = []
        if not sin_programa.empty:
            reservas_p = ", ".join(sin_programa['ID de reserva'].astype(str).tolist()[:15])
            errores.append(f"Falta CENTRO DE COSTO en ID: {reservas_p}")
            
        if not sin_vehiculo.empty:
            if tipo_reporte == 'admin': # Solo bloqueamos por vehículo si es el reporte admin
                reservas_v = ", ".join(sin_vehiculo['ID de reserva'].astype(str).tolist()[:15])
                errores.append(f"Falta VEHÍCULO (Conductor) en ID: {reservas_v}")

        if errores:
            mensaje_final = "CORRIGE EN AUTOCAB ANTES DE PROCESAR:\n\n" + "\n".join(errores)
            return jsonify({"error": mensaje_final}), 400
        # --- FIN DEL FILTRO ---

        # Limpieza
        df['Precio'] = df['Precio'].apply(limpiar_moneda)
        if 'Costo' in df.columns:
            df['Costo'] = df['Costo'].apply(limpiar_moneda)
        if 'Vehículo' in df.columns:
            df['Vehículo'] = df['Vehículo'].astype(str).str.strip()

        try:
            fechas = pd.to_datetime(df['Hora de Pickup (Local)'], dayfirst=True, errors='coerce')
            fechas_validas = fechas.dropna()
            if not fechas_validas.empty:
                min_date = fechas_validas.min()
                max_date = fechas_validas.max()
                meses = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic']
                mes_str = meses[max_date.month - 1]
                prefijo = "Admin" if tipo_reporte == 'admin' else "Cliente"
                nombre_archivo = f"{prefijo}_Almaviva_{min_date.day}_al_{max_date.day}_{mes_str}.xlsx"
            else:
                nombre_archivo = "Reporte_Almaviva.xlsx"
        except:
            nombre_archivo = "Reporte_Almaviva.xlsx"

        output = io.BytesIO()
        df_base = df[columnas_esperadas].copy()
        programas = df['Programa de viaje'].dropna().unique()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for prog in programas:
                nombre_pestana = str(prog).split(' - ')[0][:30]
                df_prog = df_base[df_base['Programa de viaje'] == prog].copy()
                df_prog.to_excel(writer, sheet_name=nombre_pestana, index=False)

        output.seek(0)
        wb = load_workbook(output)
        formato_moneda = '"$"#,##0.00'
        
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            max_row = ws.max_row
            fila_inicio = max_row + 2
            
            if tipo_reporte == 'admin':
                ws.column_dimensions['F'].width = 16
                ws.column_dimensions['G'].width = 16
                ws.column_dimensions['H'].width = 12
                for row in range(2, max_row + 1):
                    ws[f'F{row}'].number_format = formato_moneda
                    ws[f'G{row}'].number_format = formato_moneda
                
                ws[f'E{fila_inicio+3}'] = 'SUBTOTAL'
                ws[f'F{fila_inicio+3}'] = f'=SUM(F2:F{max_row})'

                ws[f'E{fila_inicio}'] = 'DTS'
                ws[f'F{fila_inicio}'] = f'=SUMIF(H2:H{max_row}, "*8979*", F2:F{max_row})'
                
                ws[f'E{fila_inicio+1}'] = 'TRC'
                ws[f'F{fila_inicio+1}'] = f'=SUM(G2:G{max_row}) - SUMIF(H2:H{max_row}, "*8979*", G2:G{max_row})'
                
                ws[f'E{fila_inicio+2}'] = 'ING'
                ws[f'F{fila_inicio+2}'] = f'=F{fila_inicio+3} - F{fila_inicio} - F{fila_inicio+1}'
                
                ws[f'E{fila_inicio+4}'] = 'COMISIÓN 5%'
                ws[f'F{fila_inicio+4}'] = f'=F{fila_inicio+3} * 0.05'
                ws[f'E{fila_inicio+5}'] = 'IVA 19%'
                ws[f'F{fila_inicio+5}'] = f'=F{fila_inicio+4} * 0.19'
                ws[f'E{fila_inicio+6}'] = 'TOTAL'
                ws[f'F{fila_inicio+6}'] = f'=F{fila_inicio+3} + F{fila_inicio+4} + F{fila_inicio+5}'
                
                for i in range(fila_inicio, fila_inicio + 7):
                    ws[f'F{i}'].number_format = formato_moneda

            else: 
                ws.column_dimensions['F'].width = 16
                ws.column_dimensions['G'].width = 30
                
                for row in range(2, max_row + 1):
                    ws[f'F{row}'].number_format = formato_moneda
                    
                ws[f'E{fila_inicio}'] = 'subtotal'
                ws[f'F{fila_inicio}'] = f'=SUM(F2:F{max_row})'
                ws[f'E{fila_inicio+1}'] = 'comision 5%'
                ws[f'F{fila_inicio+1}'] = f'=F{fila_inicio} * 0.05'
                ws[f'E{fila_inicio+2}'] = 'iva 19%'
                ws[f'F{fila_inicio+2}'] = f'=F{fila_inicio+1} * 0.19'
                ws[f'E{fila_inicio+3}'] = 'total'
                ws[f'F{fila_inicio+3}'] = f'=F{fila_inicio} + F{fila_inicio+1} + F{fila_inicio+2}'
                
                for i in range(fila_inicio, fila_inicio + 4):
                    ws[f'F{i}'].number_format = formato_moneda

        final_output = io.BytesIO()
        wb.save(final_output)
        final_output.seek(0)
        
        return send_file(
            final_output, 
            as_attachment=True, 
            download_name=nombre_archivo,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    except Exception as e:
        # Si ALGO falla en Python, se lo enviamos a Blogger para que nos diga QUÉ es
        return jsonify({"error": f"Error interno al leer tu archivo Excel: {str(e)}"}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
