import pandas as pd
import openpyxl

#Ruta del archivo CSV intermedio
archivo_csv = 'C:/Users/activ/workspace/granja-reportes-automatizados/data/produccion_semanal.csv'

try:

  data = pd.read_csv(archivo_csv)
  animales = data['ARETE'].unique().tolist()
  animales.sort()

  reporte = openpyxl.Workbook()
  reporte.save(r'C:/Users/activ/workspace/granja-reportes-automatizados/reportes/evaluacion_animales.xlsx')

  print(f"Obteniendo datos de {len(animales)} animales", end="\t\t\t")
  for animal in animales:

    eval_animal = data[data['ARETE']==animal][['FECHA','FECHA PIC','MOTILIDAD  %','T','VOL /ML SEMEN TOTAL','FR', 'SPZ/ML', 'ANOMALIAS %','AGLUT','N DOSIS 100','SPZ/DOSIS (MILLONES)']]
    eval_animal['OBSERVACIONES'] = ["BAJA MOTILIDAD Y ALTA ANOMALIA" if cantidad == "-" else None for cantidad in eval_animal['N DOSIS 100']]
    eval_animal.sort_values(by="FECHA", ascending=True, inplace=True)

  with pd.ExcelWriter("C:/Users/activ/workspace/granja-reportes-automatizados/reportes/evaluacion_animales.xlsx", engine='openpyxl', mode='a') as writer:
        eval_animal.to_excel(writer, sheet_name=animal,index=False)

  print("Listo!")
  reporte.save(r'C:/Users/activ/workspace/granja-reportes-automatizados/reportes/evaluacion_animales.xlsx')

except Exception as e:
    print(f"Error al transformar los datos: {e}")