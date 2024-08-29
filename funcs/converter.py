import os
import pandas as pd
from unidecode import unidecode
import re


# FUNCIÓN PARA BR SI EXISTE UN DIRECTORIO
def check_valid_dir(dirpath):
    if not os.path.exists(dirpath):
        os.makedirs(dirpath)
        return dirpath
    else:
        return dirpath


# FUNCIÓN PARA TENER UN PATH VÁLIDO PARA PANDAS
def valid_path(path):
    return re.sub("\s*", "", re.sub('[&%"]', '', path.replace("\\", "/")).replace("'", ""), count=1)


# FUNCIÓN PARA TENER NOMBRES DE COLUMNAS FÁCILES DE MANEJAR
def validname(name):
    return unidecode(''.join(filter(str.isalpha, name)).lower())


# FUNCIÓN PRINCIPAL -> CONVERTIR ARCHIVOS A PANDA'S DATAFRAME
def converter(filename, sheetname="Hoja1"):

    dataframe = pd.DataFrame(pd.read_excel(valid_path(filename)))
    dataframe.dropna(how="all", inplace=True)
    dataframe.columns = [validname(column) for column in dataframe.columns]

    return dataframe