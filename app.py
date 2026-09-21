from flask import Flask, request, send_file
from flask_cors import CORS
import pandas as pd
import io
from openpyxl import load_workbook

app = Flask(__name__)
# Permitimos que Blogger pueda leer el nombre del archivo generado
CORS(app, expose_headers=["Content-Disposition"])

def limpiar_moneda(x):
    if pd.isna(x): return 0.0
    if isinstance(x, str):
        x = x.replace('.', '').replace(',', '.')
    return float(x)

@app.route('/procesar-almaviva', methods=['POST'])
def procesar_archivo():
    if 'file' not in request.files:
        return "No archivo", 400
    
    file = request.files['file']
    df = pd.read_csv(file, quotechar='"')
    
    df['Precio'] = df['Precio'].apply(limpiar_moneda)
    df['Costo'] = df['Costo'].apply(limpiar_moneda)
    df['Vehículo'] = df['Vehículo'].astype(str).str.strip()

    # --- LÓGICA PARA EL NOMBRE DEL ARCHIVO CON FECHAS ---
    try:
        fechas = pd.to_datetime(df['Hora de Pickup (Local)'], dayfirst=True, errors='coerce')
        fechas_validas = fechas.dropna()
        if not fechas_validas.empty:
            min_date = fechas_validas.min()
            max_date = fechas_validas.max()
            meses = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic']
            mes_str = meses[max_date.month - 1]
            nombre_archivo = f"Almaviva_{min_date.day}_al_{max_date.day}_{mes_str}.xlsx"
        else:
            nombre_archivo = "Almaviva_Organizado.xlsx"
    except:
        nombre_archivo = "Almaviva_Organizado.xlsx"

    output = io.BytesIO()
    df_base = df[['ID de reserva', 'Hora de Pickup (Local)', 'Recoger', 'Destino', 'Nombre', 'Precio', 'Costo', 'Vehículo', 'Programa de viaje']].copy()
    programas = df['Programa de viaje'].dropna().unique()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for prog in programas:
            nombre_pestana = str(prog).split(' - ')[0][:30]
            df_prog = df_base[df_base['Programa de viaje'] == prog].copy()
            df_prog.to_excel(writer, sheet_name=nombre_pestana, index=False)

    output.seek(0)
    wb = load_workbook(output)
    
    # Formato oficial de Excel para moneda
    formato_moneda = '"$"#,##0.00'
    
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        max_row = ws.max_row
        
        # Ajustamos el ancho
        ws.column_dimensions['F'].width = 16
        ws.column_dimensions['G'].width = 16
        ws.column_dimensions['H'].width = 12
        
        # --- NUEVO: Le damos formato de precio/moneda a todas las filas de Precio y Costo ---
        for row in range(2, max_row + 1):
            ws[f'F{row}'].number_format = formato_moneda
            ws[f'G{row}'].number_format = formato_moneda
            
        fila_inicio = max_row + 2
        
        # Fórmulas de los totales
        ws[f'E{fila_inicio}'] = 'DTS'
        ws[f'F{fila_inicio}'] = f'=SUMIF(H2:H{max_row}, "8979", F2:F{max_row})'
        
        ws[f'E{fila_inicio+1}'] = 'TRC'
        ws[f'F{fila_inicio+1}'] = f'=SUMIF(H2:H{max_row}, "<>8979", G2:G{max_row})'
        
        ws[f'E{fila_inicio+2}'] = 'ING'
        ws[f'F{fila_inicio+2}'] = f'=SUMIF(H2:H{max_row}, "<>8979", F2:F{max_row}) - F{fila_inicio+1}'
        
        ws[f'E{fila_inicio+3}'] = 'SUBTOTAL'
        ws[f'F{fila_inicio+3}'] = f'=F{fila_inicio} + F{fila_inicio+1} + F{fila_inicio+2}'
        
        ws[f'E{fila_inicio+4}'] = 'COMISIÓN 5%'
        ws[f'F{fila_inicio+4}'] = f'=F{fila_inicio+3} * 0.05'
        
        ws[f'E{fila_inicio+5}'] = 'IVA 19%'
        ws[f'F{fila_inicio+5}'] = f'=F{fila_inicio+4} * 0.19'
        
        ws[f'E{fila_inicio+6}'] = 'TOTAL'
        ws[f'F{fila_inicio+6}'] = f'=F{fila_inicio+3} + F{fila_inicio+4} + F{fila_inicio+5}'

        # Le damos el mismo formato de moneda a la tabla de resultados
        for i in range(fila_inicio, fila_inicio + 7):
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

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
