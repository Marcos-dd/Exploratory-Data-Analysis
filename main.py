
from selenium.webdriver.support.ui import Select
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By

from time import sleep
import os
import shutil
from pathlib import Path

import pandas as pd
from functools import reduce
from plotly.offline import init_notebook_mode
init_notebook_mode(connected=True)

import funciones as fun

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service)

url = "http://app.bde.es/rss_www/"
driver.get(url)

# damos tiempo a que carge la página
sleep(2)

# Del desplegable 'Tipo de entidad', seleccionamos 'Investigador independiente'

tipo_entidad = Select(driver.find_element(By.ID, value = 'entidad'))
tipo_entidad.select_by_visible_text('Investigador independiente')

# Del desplegable 'Objetivo del estudio' -> 'Analisis económico sectorial'

objetivo_estudio = Select(driver.find_element(By.ID, value = 'objetivo'))
objetivo_estudio.select_by_visible_text('Análisis económico sectorial')

# Del desplegable 'País' -> 'España'

objetivo_estudio = Select(driver.find_element(By.ID, value = 'paisRegistro'))
objetivo_estudio.select_by_visible_text('España')


# Extraemos sector C26 -> Fabricación de productos informáticos, electrónicos y ópticos
C26 = driver.find_element(By.XPATH, '/html/body/table[3]/tbody/tr/td[2]/table[2]/tbody/tr[2]/td/select/option[110]')
C26.click()
fun.iter_down(driver)

# Extrameos sector J -> Información y comunicaciones
J = driver.find_element(By.XPATH, '/html/body/table[3]/tbody/tr/td[2]/table[2]/tbody/tr[2]/td/select/option[235]')
J.click()
fun.iter_down(driver)

# Extraemos el subsector J62 -> Programación, consultoría y otras actividades relacionadas con la informática
J62 = driver.find_element(By.XPATH, '/html/body/table[3]/tbody/tr/td[2]/table[2]/tbody/tr[2]/td/select/option[250]')
J62.click()
fun.iter_down(driver)

# Extraemos el subsector J631 -> Procesos de datos, hosting y actividades relacionadas; portales web
J631 = driver.find_element(By.XPATH, '/html/body/table[3]/tbody/tr/td[2]/table[2]/tbody/tr[2]/td/select/option[252]')
J631.click()
fun.iter_down(driver)

# Extraemos el sector N -> Servicios administrativos y auxiliares
N = driver.find_element(By.XPATH, '/html/body/table[3]/tbody/tr/td[2]/table[2]/tbody/tr[2]/td/select/option[283]')
N.click()
fun.iter_down(driver)

# Extraemos el sector P -> Educación
P = driver.find_element(By.XPATH, '/html/body/table[3]/tbody/tr/td[2]/table[2]/tbody/tr[2]/td/select/option[309]')
P.click()
fun.iter_down(driver)

# Extraemos el sector Q -> Sanidad y Servicios Sociales
Q = driver.find_element(By.XPATH, '/html/body/table[3]/tbody/tr/td[2]/table[2]/tbody/tr[2]/td/select/option[317]')
Q.click()
fun.iter_down(driver)

# Extraemos el sector I -> Hostelería
I = driver.find_element(By.XPATH, '/html/body/table[3]/tbody/tr/td[2]/table[2]/tbody/tr[2]/td/select/option[225]')
I.click()
fun.iter_down(driver)

# esta parte del programa mueve todos los archivos desde la carpeta download hasta la carpeta destino:
#       - 1º Crea carpeta 'datasets'
#       - 2º Crea subcarpetas con los nombres de los sectores
#       - 3º En cada subcarpeta guardará los excel del sector correspondiente 

# después realizará una serie de acciones sobre los excel de cada subcarpeta y nos creará dos excel nuevos en cada una de ellas:
#       - Uno que mostrará la evolución del número de empresas de ese sector durante los años en los que hay registro (2000-2020)
#           - Para este caso se ha seleccionado el indicador de ratio 'Cifra neta de negocios / Total activo' ya que que muestra el total.
#           - El resto de indicadores de ratio o son iguales o están por debajo del indicado (respecto a número de empresas)
#       - Y otro que mostrará la evolución de todos los ratios sectoriales en ese mismo periodo


# empezamos creando la lista de sectores que queramos estudiar. Las siguientes versiones traeran esta opción definida a través de un input
#   - todas las acciones que se muestran a continuación se completarán una vez para cada sector
sector_list = ['C26','J','J62','J631','N','P','Q','I']

# C26-> Fabricación de productos informáticos, electrónicos y ópticos
# J -> Información y comunicaciones
# J62-> Programación, consultoría y otras actividades relacionadas con la informática
# J631-> Procesos de datos, hosting y actividades relacionadas; portales web
# N-> Servicios administrativos y auxiliares
# P-> Educación
# Q-> Sanidad y Servicios Sociales
# I-> Hostelería

# creamos una serie de diccionarios y listas vacias que iremos usando a lo largo de la ejecución
dict_titles = {'C26':' -> Fabricación de productos informáticos, electrónicos y ópticos',
                    'J':' -> Información y comunicaciones',
                    'J62':' -> Programación, consultoría y otras actividades relacionadas con la informática',
                    'J631':' -> Procesos de datos, hosting y actividades relacionadas; portales web',
                    'N':' -> Servicios administrativos y auxiliares',
                    'P':' -> Educación',
                    'Q':' -> Sanidad y Servicios Sociales',
                    'I':' -> Hostelería'}
TVM_rentabilidad_list = []
TVM_rendimiento_list = []
TVM_index = []
TVM_titulos = []
dict_titulos = {'C26':'Productos informáticos',
                    'J':' -> Información y comunicaciones',
                    'J62':'Programación, consultoría',
                    'J631':'Procesos de datos, hosting',
                    'N':'Servicios administrativos y auxiliares',
                    'P':'Educación',
                    'Q':'Sanidad y Servicios Sociales',
                    'I':'Hostelería'}

# obtiene la ruta de la carpeta de descargas y la ruta actual
downloadfd_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Downloads')
downloadfd = os.path.join(downloadfd_path)
current_dir = os.getcwd()

# crea la carpeta 'datasets' en el current directory (si ya existe, no la vuelve a crear)  
try:
    if os.path.exists(os.path.join(current_dir, 'data')):
        destiny_folder_path = os.path.join(current_dir, 'data')
    
    else:
        destiny_folder = os.mkdir('data')
        destiny_folder_path = os.path.join(current_dir, 'data')  
except:
    destiny_folder_path = os.path.join(current_dir, 'data') 

# crea la carpeta 'graficas' en el current directory

if os.path.exists(os.path.join(current_dir, 'graficas')):
    graphics_folder_path = os.path.join(current_dir, 'graficas')
    
else:
    graphic_folder = os.mkdir('graficas')
    graphics_folder_path = os.path.join(current_dir, 'graficas')

# Iteramos en la lista de sectores y para cada uno de ellos se realizaran las acciones indicadas al incio de este bloque
for sector in sector_list:
    input_sector = f'_{sector}_'

    # crea la carpeta con el nombre del sector
    try:        
        previous_path = Path(destiny_folder_path)
        sector_path = previous_path.joinpath(input_sector)
        sector_folder = os.mkdir(sector_path)
        
    except Exception:
        previous_path = Path(destiny_folder_path)
        sector_path = previous_path.joinpath(input_sector)

            
    # lista todos los archivos presentes en la carpeta Download
    files = os.listdir(downloadfd)
 
    # vamos a crear en paralelo una lista con los nombres de los archivos para iterarla más tarde y montar el dataframe    
    datasets_list = []
    dataframes = []

    # mueve los archivos .xls pertenecientes al sector y que nos acabamos de descargar, en la carpeta 'datasets/sector'. Para ello:
    #       - busca en la carpeta descargas los excel con el código del sector
    #       - comprueba que el archivo no exista ya en la carpeta destino
    #       - comprueba que no se haya incluido ya en la lista de datasets
    #       - de los archivos resultantes, descarta aquellos duplicados (en los que en su nombre figura '(1).xlsx', '(2).xlsx', etc)
    
    # itera sobre todos los archivos de la carpeta download y ejecuta las acciones comentadas
    for f in files:
        if os.path.exists(os.path.join(sector_path, f)):
            datasets_list.append(f)
                
        else:
            if f.endswith('xls') and input_sector in f:
                text = f.split('_',2)
                ref_code = f'{text[0]}_{text[1]}'
                if ref_code not in datasets_list and len(f)<25:                
                    datasets_list.append(f)
                    shutil.move(os.path.join(downloadfd, f),  # .copy para testar el programa. -> .move para ejecutar
                            os.path.join(sector_path, f))  
          

    # ordenamos primero la lista por si se hubieran descargado o almacenado desordenados (que no debería)                           
    datasets_list = sorted(datasets_list)

    # montamos una lista de dataframes. Uno para cada objeto de estudio (número de empresas y medias).
    #        - así podremos trabajar por separado en cada uno de ellos, ya que cada necesitaremos cosas distintas de cada uno.
    dataframes_median = []
    dataframes_nterp = []

    # itera sobre los excel que han pasado el filtro anterior y que se han ido guardado en la subcarpeta y en la lista vacia que creamos
    for dataset in datasets_list:
    # para cada uno de los excel:
        # fragmenta el nombre del archivo y cogemos el primer valor (año) que nos servirá para nombrar las columnas de dataframe más adelante
        text = dataset.split('_')
        year = text[0]

        # crea el dataframe de las medianas (Q2)
        # lee el excel, se queda con las columnas 'Nombre de Ratio' y 'Q2' (mediana), elimina las filas en blanco, y renombra la columna de 'Q2'
        df_median = pd.read_excel(os.path.join(sector_path, dataset), 
        header = 11,
        nrows = 35,
        usecols = "B,E",
        names = ["Nombre de Ratio",f"Q2_{year}_{sector}"],        
        keep_default_na = False).drop([0,10,15,20,25,33],axis = "index")

        # sobreescribe el índice implícito por la columna de 'Nombre de Ratio' y añade el resultado a la lista de dataframes
        df_median.set_index('Nombre de Ratio', inplace=True)
        dataframes_median.append(df_median)

        # crea el dataframe con el número de empresas
        # lee el excel, se queda con las columnas 'Nombre de Ratio' y 'Empresas', elimina las filas en blanco, y renombra la columna de 'Empresas'
        df_nterp = pd.read_excel(os.path.join(sector_path, dataset), 
        header = 11,
        nrows = 35,
        usecols = "B,C",
        names = ["Nombre de Ratio",f"Empresas_{year}_{sector}"],        
        keep_default_na = False).drop([0,10,15,20,25,33],axis = "index")

        # sobreescribe el índice implícito por la columna de 'Nombre de Ratio' y añade el resultado a la lista de dataframes
        df_nterp.set_index('Nombre de Ratio', inplace=True)
        dataframes_nterp.append(df_nterp)

    # une todos los dataframes de medianas por un lado y los de número de empresas por otro
    df_merged_median = reduce(lambda  left,right: pd.concat([left,right],join='outer',
                                            axis=1), dataframes_median)                                       
    df_merged_nterp = reduce(lambda  left,right: pd.concat([left,right],join='outer',
                                            axis=1), dataframes_nterp)

    # transpone los dataframe resultantes
    df_merged_median_T = df_merged_median.transpose()
    df_merged_nterp_T = df_merged_nterp.transpose()

    # añade una columna con los años de los ejercicios en cada dataframe
    df_merged_median_T.insert(1,'Ejercicio',[2000, 2001, 2002, 2003, 2004, 2005, 2006, 2007, 2008, 2009, 2010,
       2011, 2012, 2013, 2014, 2015, 2016, 2017, 2018, 2019, 2020])
    df_merged_nterp_T.insert(1,'Ejercicio',[2000, 2001, 2002, 2003, 2004, 2005, 2006, 2007, 2008, 2009, 2010,
       2011, 2012, 2013, 2014, 2015, 2016, 2017, 2018, 2019, 2020])

    # como se indicaba al inicio, coge sólo la columna 'Cifra neta de negocios / Total activo' que el que mejor muestra el número total de empresas
    df_merged_nterp_T = df_merged_nterp_T['Cifra neta de negocios / Total activo']
    df_merged_nterp_T.rename('Numero de empresas', inplace=True)

    # crea los excel a partir de los dataframes y los guarda en la subcarpeta propia del sector
    with pd.ExcelWriter(f'{sector_path}\{sector}_median.xlsx') as writer: 
        df_merged_median_T.to_excel(writer , sheet_name=sector)
    with pd.ExcelWriter(f'{sector_path}\{sector}_entreprises.xlsx') as writer: 
        df_merged_nterp_T.to_excel(writer , sheet_name=sector)

    # lee los excel creados y opera en ellos para unirlos en uno sólo sobre el que se trabajará
    df1 = pd.read_excel(f'{sector_path}\{sector}_entreprises.xlsx')
    df2 = pd.read_excel(f'{sector_path}\{sector}_median.xlsx')

    df1.drop(columns=['Unnamed: 0'], inplace=True)

    df1.reset_index(inplace=True)
    df2.reset_index(inplace=True)

    final_sector_df = df2.merge(df1, how='left')
    final_sector_df.drop(columns=['index'], inplace= True)

    # nos quedamos sólo con los ratios 'Cifra neta de negocios / Total activo' (rentabilidad) y 'Resultado económico neto / Total activo' (rendimiento)
    final_sector_df = final_sector_df.loc[:,['Ejercicio','Cifra neta de negocios / Total activo','Resultado económico neto / Total activo','Numero de empresas']]

    with pd.ExcelWriter(f'{sector_path}\{sector}_total.xlsx') as writer: 
        final_sector_df.to_excel(writer , sheet_name=sector)

    # calculamos las tasas de variacion media de todos los sectores
for sector in sector_list:
    input_sector = f'_{sector}_'
    fun.calculo_tvm(TVM_rentabilidad_list,TVM_rendimiento_list,TVM_index,TVM_titulos,dict_titulos,sector,input_sector)

# montamos el DataFrame de TVM
TVM_df = pd.DataFrame({'Rentabilidad':TVM_rentabilidad_list, 'Rendimiento':TVM_rendimiento_list, 'Sector':TVM_index,'Titulo':TVM_titulos})
  
# creamos y guardamos la gráfica de TVM
data_dict = {}
# creamos un diccionario con los datasets a medida que los vamos leyendo -> {clave = nombre del dataset, valor = dataset}
# aplicamos la función 'grafica' a cada uno de los sectores
for sector in sector_list:
    input_sector = f'_{sector}_'
    sector_df = pd.read_excel(f'data\{input_sector}\{sector}_total.xlsx', index_col='Unnamed: 0')
    pair = {sector:sector_df}
    data_dict.update(pair)
    fun.grafica_ratios(data_dict,sector,graphics_folder_path,dict_titles)
    fun.grafica_empresas(data_dict,sector,graphics_folder_path,dict_titles)

fun.grafica_tvm(TVM_df,graphics_folder_path)