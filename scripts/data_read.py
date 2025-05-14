import pandas as pd

# Ruta del archivo EXCEL
archivo_xlsx = 'C:/Users/activ/workspace/granja-reportes-automatizados/data/produccion_semanal.xlsx'

try:
    # Leer el archivo EXCEL
    data = pd.read_excel(archivo_xlsx)
    
    # Guardar los datos en un archivo CSV temporal
    archivo_csv = 'C:/Users/activ/workspace/granja-reportes-automatizados/data/produccion_semanal.csv'
    data.to_csv(archivo_csv, index=False)
    
    print(f"Datos leídos y guardados en {archivo_csv}")
except Exception as e:
    print(f"Error al leer el archivo: {e}")