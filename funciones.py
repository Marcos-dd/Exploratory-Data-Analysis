from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
import plotly.graph_objs as go

import pandas as pd


# definimos una función para iterar sobre todos los años dentro del desplegable 'Ejercicio' de la web
# la usaremos para cada uno de los sectores que queramos incluir en la muestra
def iter_down(driver):
    
    for i in range(21):
        ejercicio = Select(driver.find_element(By.ID, value = 'ejercicio'))
        año = str(2020 - i)
        ejercicio.select_by_visible_text(año)
        download_excel = driver.find_element(By.XPATH, '/html/body/table[3]/tbody/tr/td[2]/table[3]/tbody/tr/td/input[2]')
        download_excel.click()


# definimos una función que pinte las gráficas y las guarde como archivo .html en la carpeta 'graficas' creada
# para rendimiento y rentabilidad
def grafica_ratios(data_dict,sector,graphics_folder_path,dict_titles):
   
    rentabilidad = go.Scatter(
                        x = data_dict[sector]['Ejercicio'],                    
                        y = data_dict[sector]['Cifra neta de negocios / Total activo'],
                        name = 'rentabilidad',
                        mode = 'lines',
                        marker = dict(color = 'rgba(94, 0, 210, 0.92)'),
                        text = data_dict[sector]['Cifra neta de negocios / Total activo'])

    rendimiento = go.Scatter(
                        x = data_dict[sector]['Ejercicio'],                    
                        y = data_dict[sector]['Resultado económico neto / Total activo']*20,
                        name = 'rendimiento',
                        mode = 'lines',
                        marker = dict(color = 'rgba(255, 0, 24, 0.92)'),
                        text = data_dict[sector]['Resultado económico neto / Total activo'])

    sector_graph = [rentabilidad, rendimiento]

    layout = dict(title = f'Resumen Estados financieros EUROPA      -SECTOR -{sector} {dict_titles[sector]}',
                 xaxis= dict(title= 'Ejercicios',dtick = 2 ,ticklen= 5, ticks = 'outside', tickson = 'boundaries'))

    fig = go.Figure(data = sector_graph, layout=layout)
    fig.write_html(f'{graphics_folder_path}\{sector}_ratios_graph.html')

# para número de empresas
def grafica_empresas(data_dict,sector,graphics_folder_path,dict_titles):

    empresas = go.Scatter(
                        x = data_dict[sector]['Ejercicio'],                   
                        y = data_dict[sector]['Numero de empresas'],
                        name = 'Número de empresas',
                        mode = 'lines+markers',
                        marker = dict(color = 'rgba(238, 0, 224, 0.92)'),
                        text = data_dict[sector]['Numero de empresas'])
    sector_graph = [empresas]

    layout = dict(title = f'Evolución Número de empresas EUROPA      -SECTOR -{sector} {dict_titles[sector]}',
                 xaxis= dict(title= 'Ejercicios',dtick = 2 ,ticklen= 5, ticks = 'outside', tickson = 'boundaries'))

    fig = go.Figure(data = sector_graph, layout=layout)
    fig.write_html(f'{graphics_folder_path}\{sector}_nterp_graph.html')

# función para el cálculo de la Tasa de Variación Media (TVM)
def calculo_tvm(TVM_rentabilidad_list,TVM_rendimiento_list,TVM_index,TVM_titulos,dict_titulos,sector,input_sector):

    
    ex = pd.read_excel(f'data\{input_sector}\{sector}_total.xlsx', index_col=0)
    ex_rentabilidad = ex['Cifra neta de negocios / Total activo']
    ex_rendimiento = ex['Resultado económico neto / Total activo']
    TVM_rentabilidad = round((ex_rentabilidad[20]-ex_rentabilidad[0])/20, 2)
    TVM_rendimiento = round((ex_rendimiento[20]-ex_rendimiento[0])/20, 2)
    
    TVM_rentabilidad_list.append(TVM_rentabilidad)
    TVM_rendimiento_list.append(TVM_rendimiento)
    TVM_index.append(sector)
    TVM_titulos.append(dict_titulos[sector])
    TVM_df = pd.DataFrame({'Rentabilidad':TVM_rentabilidad_list, 'Rendimiento':TVM_rendimiento_list, 'Sector':TVM_index,'Titulo':TVM_titulos} )

    return TVM_df

# función para crear la gráfica con la TVM
def grafica_tvm(TVM_df,graphics_folder_path):
    rentabilidad = go.Bar(x = TVM_df['Titulo'],
                y = TVM_df['Rentabilidad'],
                name = 'Rentabilidad',
                marker = dict(color = 'rgba(94, 0, 210, 0.92)',
                line = dict(color='rgb(0,0,0)', width = 1.5)),
                text = TVM_df['Sector'])

    rendimiento = go.Bar(x = TVM_df['Titulo'],
               y = TVM_df['Rendimiento'],
               name = 'Rendimiento',
               marker = dict(color = 'rgba(255, 255, 128, 0.5)',
                            line = dict(color='rgb(0,0,0)', width = 1.5)),
               text = TVM_df['Sector'])


    data = [rendimiento, rentabilidad]

    layout = go.Layout(barmode = "group")

    fig = go.Figure(data = data, layout = layout)
    fig.write_html(f'{graphics_folder_path}\TVM.html')
