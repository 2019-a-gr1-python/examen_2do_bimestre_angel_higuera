import numpy as np
import pandas as pd
import os

# Archivos de entrada
path_doc_info_general="./di/INFORMACIÓN-GENERAL.xlsx"
path_doc_articulos="./di/articulos.xlsx"
# Archivos de salida
path_log_bodegas="./df/bodegas-py.xlsx"
path_log_clientes="./df/clientes-py.xlsx"
path_log_contacto="./df/contacto-cliente.xlsx"
path_log_empleados="./df/empleados-py.xlsx"
path_log_empresas="./df/empresas-py.xlsx"
path_log_grupos="./df/grupos-subgrupos.xlsx"
path_log_productos="./df/productos-py.xlsx"
path_log_vendedores="./df/vendedores-py.xlsx"
# funciones auxiliares

def format_ruc(ruc):
  ruc=f"{ruc}"
  ruc=ruc.replace("´", "")
  return ruc

def format_contact_numbers(contact):
  contact=f"{contact}"
  contact=contact.replace("´", "")
  contact=contact.replace("-", "")
  contact=contact.replace("nan", "")
  contact=contact.replace("\n", "/")
  return contact

# filtrar contacto-cliente
df_read_reventa = pd.read_excel(path_doc_info_general, sheet_name="REVENTA")
df_read_reventa.columns = df_read_reventa.iloc[0]
df_read_reventa = df_read_reventa.iloc[1:]
all_columns = df_read_reventa.columns

"""
0: 'CÓDIGO'
1: 'TIPO DE CLIENTE'
2: 'RUC'
3: 'NOMBRE DEL CLIENTE'
4: 'TIPO DE CLIENTE (REVENTA)'
5: 'NOMBRE COMERCIAL'
6: 'NOMBRE DEL REPRESENTATE  LEGAL'
7: 'PROVINCIA'
8: 'CIUDAD'
9: 'CIUDAD, DIRECCIÓN '
10: 'TELÉFONO '
11: 'CORREO ELECTRÓNICO 1'
12: 'CORREO ELECTRÓNICO 2'
13: 'CONTACTO 1'
14: 'CARGO'
15: 'CONTACTO 2'
16: 'CARGO',
17: 'OBSERVACIONES'
18: 'Latitud'
19: 'Longitud'
"""

df_read_reventa[all_columns[2]] = df_read_reventa[all_columns[2]].map(format_ruc)
df_read_reventa[all_columns[11]] = df_read_reventa[all_columns[11]] + "/"+ df_read_reventa[all_columns[12]]
df_read_reventa[all_columns[7]] = df_read_reventa[all_columns[7]] + "/"+ df_read_reventa[all_columns[8]]
df_read_reventa[all_columns[10]] = df_read_reventa[all_columns[10]].map(format_contact_numbers)

selected_columns_contactos = df_read_reventa.columns[[2, 3, 11, 10, 0]]
df_contactos_clientes = df_read_reventa.loc[0:][selected_columns_contactos]

final_columns_names = [
    "Identificador contacto", #0-2
    "Contacto", #1-3
    "Correo(opcional)", #2-11
    "Telefono(opcional)", #3-10
    "Cliente(codigo)" #4-0
]
df_contactos_clientes.columns=final_columns_names
df_contactos_clientes.to_excel(path_log_contacto, index=False)

# filtrar contacto-vendedor
selected_columns_clientes = df_read_reventa.columns[[0, 1, 1, 3, 4, 5, 7, 18, 19, 10, 11]]
df_clientes = df_read_reventa.loc[0:][selected_columns_clientes]

final_columns_clientes = [
  "Codigo", #0-0
  "Tipo cliente", #1-1
  "Ruc", #2-2
  "Nombre del Cliente", #3-3
  "Tipo cliente Reventa", #4-4
  "Nombre comercial", #5-5
  "Provincia/Ciudad", #6-7/8
  "latitud direccion", #7-X
  "longitud direccion", #8-X
  "Telefono", #9-10
  "correo electronico" #10-11
]
df_clientes.columns=final_columns_clientes
df_clientes.to_excel(path_log_clientes, index=False)



