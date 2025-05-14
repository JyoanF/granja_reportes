import pandas as pd
import openpyxl

#Ruta del archivo CSV intermedio
archivo_csv = 'C:/Users/activ/workspace/granja-reportes-automatizados/data/produccion_semanal.csv'

try:

    data = pd.read_csv(archivo_csv)
    data['N DOSIS 100'].replace("-","0",inplace=True)
    data['N DOSIS 100'] = data['N DOSIS 100'].astype(int)

    reporte = openpyxl.Workbook()
    reporte.save(r'C:/Users/activ/workspace/granja-reportes-automatizados/reportes/dosis_desechada.xlsx')

    pvt_colecciones = pd.pivot_table(data[data['N DOSIS 100']!=0],index='ARETE', values='N DOSIS 100', columns='SPZ/DOSIS (MILLONES)', aggfunc='sum')
    pvt_colecciones['TOTAL GENERAL'] = pvt_colecciones.sum(axis=1)
    pvt_colecciones.insert(0,'ARETES',pvt_colecciones.index)
    with pd.ExcelWriter("C:/Users/activ/workspace/granja-reportes-automatizados/reportes/dosis_desechada.xlsx", engine='openpyxl', mode='a') as writer:
      pvt_colecciones.to_excel(writer, sheet_name='Tabla',index=False)

    print("Listo!")
    reporte.save(r'C:/Users/activ/workspace/granja-reportes-automatizados/reportes/dosis_desechadas.xlsx')

except Exception as e:
    print(f"Error al obtener informe: {e}")