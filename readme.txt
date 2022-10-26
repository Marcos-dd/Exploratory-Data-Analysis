Extract, Move & Merge - EMM Local API  (Beta Version)
-----------------------------------------------------
-----------------------------------------------------

EMM es un programa de extracción de datos automatizado. Dicha extracción la realiza desde la página web del banco de españa http://app.bde.es/rss_www/ , desde donde se descargarán
todos los archivos excel solicitados de todos los ejercicios (años) disponibles en la web. Actualmente del 2000 al 2020.

EMM ofrece una gráfica por sector donde se muestra la evolución de los ratios sectoriales R16 (Cifra Neta de Negocio / Total Activo) y R10 (Resultado Económico Neto / Total Activo), 
otra gráfica por sector donde se muestra la evolución del número de empresas y una última gráfica comparativa conjunta de las tasas de variación de R16 y R10 (TVM).

Dichos ratios pueden ofrecer una simplificación suficiente para entender la evolución de los sectores y el impacto de las crisis sobre ellos.

En esta versión Beta se estudia el ratio R16((Ingresos-Gastos)/Total de Activos) que se muestra como Rentabilidad y el R10(Ventas/Total de Activos) que se muestra 
como Rendimiento y vienen predefinidos para su estudio los siguientes sectores:

# C26-> Fabricación de productos informáticos, electrónicos y ópticos
# J -> Información y comunicaciones
# J62-> Programación, consultoría y otras actividades relacionadas con la informática
# J631-> Procesos de datos, hosting y actividades relacionadas; portales web
# N-> Servicios administrativos y auxiliares
# P-> Educación
# Q-> Sanidad y Servicios Sociales
# I-> Hostelería

"En la versión definitiva se puodrán seleccionar todos los sectores sobre los que se quiera realizar el estudio."

-----------------------------------------------------
Cómo ejecutarlo:
1.- Guardar el archivo main.py y funciones.py en local.
2.- Configurar el path de descargas del navegador a la carpeta de descargas en local (Download por defecto)
3.- Ejecutar el archivo main.py


Una vez ejecutado, el programa creará dos carpetas, 'data' y 'graficas', en la carpeta donde se hayan guardado los archivos .py.

En /data creará tantas subcarpetas como sectores se hayan escogido para el estudio. 
Cada subcarpeta llevará el nombre del sector y contendrá los informes del 2000 al 2020 con todos los ratios sectoriales e información complementaria de dicho sector.
Además creará tres archivos más:
	- {sector}/enterprises.xlsx: Muestra el número de empresas del sector sometidas al estudio en cada año. 
					     Se ha elegido el ratio R16:Cifra neta de negocios/Total activo para obtener estas cifras ya que posee los valores máximos en todos los informes 
					     demostrando ser el valor más representativo.
	- {sector}/median.xlsx: Ofrece una muestra de todos los valores Q2 de todos los ratios por cada año.
	- {sector}/total.xlsx: Tabla resumen que muestra el número de empresas y los ratios sometidos a estudio por cada año

En /graficas se crearán, en formato .html, los gráficos iteractivos correspondientes a los excel comentados para cada sector así como un gráfico TVM.html donde se mostrarán las 
correspodientes tasas de variación media.
