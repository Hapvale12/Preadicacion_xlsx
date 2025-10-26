import pandas as pd
import re
import sys
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

print("----------------------------------------------------------")
print("  PROGRAMA DE ASIGNACIONES - EXTRACCIÓN DE TEXTO")
print("----------------------------------------------------------")
print("PASO 1: Copia el texto completo de las asignaciones de WhatsApp.")
print("PASO 2: Pégalo aquí (puedes pegar varias líneas a la vez).")
print("PASO 3: Una vez pegado, presiona Enter y luego Ctrl+Z (o Ctrl+D) y luego Enter de nuevo para finalizar la entrada.")
print("----------------------------------------------------------")

try:
    import sys
    texto_whatsapp = sys.stdin.read()
    if not texto_whatsapp.strip():
        raise EOFError 
except EOFError:
    print("\nNo se detectó entrada. Por favor, pega el texto de WhatsApp manualmente:")
    texto_whatsapp = input("Pega aquí el texto de WhatsApp: ")
    
REGEX_PATTERN = re.compile(
    r"^(?P<Dia>\w+)\s+"
    r"(?P<Hora>\d{1,2}:\d{2})\s*(?:am|pm)?\.?\s*" 
    r"(?P<Cuerpo>.+)$" 
)

datos_limpios = []
for linea in texto_whatsapp.strip().split('\n'):
    linea = linea.strip()
    if not linea.startswith(('Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo')):
        continue

    match = REGEX_PATTERN.search(linea)
    
    if match:
        data = match.groupdict()
        cuerpo = data['Cuerpo'].strip()

        final_pattern = r"(.*)\sTerritorio\s(\d+\s*.*?)[\.]?\s*(?P<Conductor>[A-Z][a-z]+\s+[A-Z][a-z]+.*)$"
        final_match = re.search(final_pattern, cuerpo)


        if final_match:
            # 1. El conductor es el último grupo
            conductor = final_match.group('Conductor').strip().strip('.')
            
            # 2. El Territorio es el segundo grupo
            territorio_bruto = final_match.group(2).strip().strip('.').replace(' y ', ', ')
            
            # 3. El Lugar es el primer grupo
            lugar_bruto = final_match.group(1).strip().strip('.').replace('Avenida', 'Av.')

            # Lógica de normalización de TERRITORIO
            normalized_pattern = r'(\d+\s*parte\s+[a-zA-Z])\s*(.*)'
            norm_match = re.search(normalized_pattern, territorio_bruto, flags=re.IGNORECASE)

            if norm_match:
                terr_base = norm_match.group(1).strip() 
                descripcion_adicional = norm_match.group(2).strip().strip('.') 
                
                terr_final = re.sub(r'\s*parte\s+([a-zA-Z]).*$', r'\1', terr_base, flags=re.IGNORECASE)
                terr_final = 'T.' + ''.join(terr_final.split()).upper()
                if terr_final == 'T.19A' : 
                    terr_final = 'T.19Animas'
                
            else:
                # Territorio simple (Ej: 18, 3)
                terr_final = 'T.' + ''.join(territorio_bruto.split()).upper().strip('.')
                terr_final = terr_final.replace('y', ', ')
        
        else:
            print(f"⚠️ Alerta: No se pudo identificar Conductor/Territorio: {linea}. Revisando Fallback...")
            
            parts = re.split(r'\sTerritorio\s', cuerpo)

            if len(parts) > 1:
                terr_cond_part = parts[-1].strip().strip('.')
                
                last_split = re.split(r'(\d+\s*parte\s*[a-zA-Z]\s*.*|\d+)\s*\.\s*', terr_cond_part, 1)

                if len(last_split) >= 3:
                    territorio_bruto = last_split[1].strip()
                    conductor = last_split[2].strip().strip('.')
                    lugar_bruto = parts[0].strip().strip('.')
                    terr_final = 'T.' + ''.join(territorio_bruto.split()).upper().strip('.')
                else:
                    lugar_bruto, conductor, terr_final = "ERROR", "ERROR", "T.XX"
            else:
                 lugar_bruto, conductor, terr_final = "ERROR", "ERROR", "T.XX"
        
        hora_final = data['Hora']
        if data['Dia'].lower() in ['lunes', 'martes', 'miércoles', 'jueves', 'viernes'] and '7:00' in data['Hora']:
             hora_final += 'pm'
        elif data['Dia'].lower() not in ['sábado', 'domingo']: 
             hora_final += 'am'

        datos_limpios.append({
            'DÍA': data['Dia'].capitalize(),
            'HORA': hora_final,
            'CONDUCTOR': conductor,
            'LUGAR_BRUTO': lugar_bruto, 
            'TERRITORIO': terr_final
        })

df = pd.DataFrame(datos_limpios)

mapa_lugar = {
    'Avenida Surco con santo cristo': 'Santo Cristo / San Juan',
    'Avenida Santo Cristo con San Juan': 'Santo Cristo / San Juan',
    'Avenida Surco con San Juan': 'San Juan / Surco',
    'Parque Familia Duarte': 'Santo Cristo / San Juan',
    'Parque Familia Manta': 'Santo Cristo / San Juan',
    'Parque El Palmar': 'San Juan / Surco',
    'Parque el Palmar': 'San Juan / Surco',
    'Salón del Reino': 'SALON DEL REINO',
    'Salón del reino': 'SALON DEL REINO'
}

df['LUGAR DE PREDICACION'] = df['LUGAR_BRUTO'].apply(lambda x: next((v for k, v in mapa_lugar.items() if k in x), x))

df['MÉTODO DE PREDICACIÓN'] = 'PREDICACION Y NO EN CASA P.'
df['ID_ASIGNACION'] = range(1, len(df) + 1) 

columnas_finales = ['ID_ASIGNACION', 'DÍA', 'HORA', 'CONDUCTOR', 'LUGAR DE PREDICACION', 'MÉTODO DE PREDICACIÓN', 'TERRITORIO']
df_final = df[columnas_finales]


PLANTILLA_FILE = "./template/template.xlsx" 
OUTPUT_FILE = f"./output/ROL_FINAL_{datetime.now().strftime('%d-%m-%Y')}.xlsx"
HOJA_DATOS = 'Datos_Crudos'

try:
    print("\nIniciando exportación a Excel...")
    wb = load_workbook(PLANTILLA_FILE)

    if HOJA_DATOS in wb.sheetnames:
        del wb[HOJA_DATOS]
    
    ws_datos = wb.create_sheet(HOJA_DATOS)
    
    for r_idx, row in enumerate(dataframe_to_rows(df_final, header=True, index=False)):
        ws_datos.append(row)
        
    wb.save(OUTPUT_FILE)
    
    print(f"\n✅ Proceso completado exitosamente.")
    print(f"Archivo generado: {OUTPUT_FILE}")

except FileNotFoundError:
    print(f"\n❌ ERROR: No se encontró el archivo de plantilla: {PLANTILLA_FILE}")
except Exception as e:
    print(f"\n❌ ERROR INESPERADO: {e}")