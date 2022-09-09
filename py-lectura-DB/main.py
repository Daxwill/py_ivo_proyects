import pandas as pd
import pyodbc
from fast_to_sql import fast_to_sql as fts

archivo = pd.read_excel("nomina-personal-jurado-de-enjuiciamiento-julio-2022.xlsx")

server = input("Escriba el nombre de su servidor de SQL: ")
nom_data_base = 'master'

try:
    conn = pyodbc.connect('DRIVER={SQL Server};SERVER=' + server + ';DATABASE=' + nom_data_base + ';Trusted_Connection=yes;',
                          autocommit=True)
    cursor= conn.cursor()

    nom_data_base = input("Ingrese el nombre de la nueva Base de Datos: ")
    cursor.execute('create database ' + nom_data_base + ';')
    conn.close()

    # conexion a la base de datos creada
    conn = pyodbc.connect(
        'DRIVER={SQL Server};SERVER=' + server + ';DATABASE=' + nom_data_base + ';Trusted_Connection=yes;',
        autocommit=True)
    print("Conexion a base de datos nueva exitosa")

    nom_table = input("Ingrese el nombre de la nueva tabla: ")

    subirDB = fts.fast_to_sql(archivo,nom_table,conn,if_exists="append",custom=None,temp=False)

    archivo1 = pd.read_sql('Select * from ' + nom_table + ';', conn)

    print (archivo1)

    conn.commit()

except Exception as ex:
    print(ex)
finally:
    conn.close()
    print("Conexion finalizada")

"""
    
"""