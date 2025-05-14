import pandas as pd
import openpyxl

#Ruta del archivo CSV intermedio
archivo_csv = 'C:/Users/activ/workspace/granja-reportes-automatizados/data/produccion_semanal.csv'


def getLinea(macho):
  df = pd.read_excel("C:/Users/activ/workspace/granja-reportes-automatizados/data/datos_animales.xlsx")

  if(len(macho)>6):
    macho = separarMachos(macho)[0]

  linea = df[df["ARETE"]==macho]["LINEA"].values[0]
  return linea

def separarMachos(pool):
  machos=[]
  macho1=''
  macho2=''
  if 'AS' in pool:
    for i in range(len(pool)):
        if(pool[i]=='A' and pool[i+1]=='S'):
          macho2=pool[i]
          macho1=pool[0:i]
        else:
          macho2+=pool[i]
  else:
    for i in range(len(pool)):
        if(pool[i]=='A' or pool[i]=='S' or pool[i]=='Z'):
          macho2=pool[i]
          macho1=pool[0:i]
        else:
          macho2+=pool[i]
  machos.append(macho1.strip())
  machos.append(macho2.strip())
  return machos

try:

  data = pd.read_csv(archivo_csv)

  # Calculo de DOSIS POR LINEA Y TIPO
  df_terminales = data[data['N DOSIS 100']!='-'][['ARETE','N DOSIS 100','SPZ/DOSIS (MILLONES)']]
  df_terminales = df_terminales.groupby(['ARETE','SPZ/DOSIS (MILLONES)'], as_index=False)['N DOSIS 100'].sum()
  df_terminales['LINEA']=[getLinea(animal) for animal in df_terminales['ARETE']]
  pvt_table_terminales = pd.pivot_table(df_terminales,index='LINEA', values='N DOSIS 100', columns='SPZ/DOSIS (MILLONES)', aggfunc='sum')
  pvt_table_terminales.loc['TOTAL'] = pvt_table_terminales.sum(axis=0)

  # Exportar a Excel
  archivo_excel = 'C:/Users/activ/workspace/granja-reportes-automatizados/reportes/produccion_por_linea_-_tipo.xlsx'
  pvt_table_terminales.to_excel(archivo_excel, index=False)
  print(f"Datos exportados exitosamente a {archivo_excel}")


except Exception as e:
    print(f"Error al obtener informe: {e}")