# std
import re
from json import dumps
from typing import Any
from inspect import stack
from unicodedata import normalize
from dataclasses import dataclass
# externo
from pandas import DataFrame

def remover_acentuacao(string: str) -> str:
    """Remover a acentuação de uma string"""
    nfkd = normalize('NFKD', string)
    ascii = nfkd.encode('ASCII', 'ignore')
    return ascii.decode("utf8")

def normalizar(string: str) -> str:
    """Strip, lower, replace espaços por underline e remoção de acentuação"""
    return remover_acentuacao( string.strip().replace(" ", "_").lower() )

def mapear_dtypes(df: DataFrame) -> dict:
    """Criar um dicionário { coluna: tipo } de um dataframe"""
    mapa = {}
    for colunaTipo in df.dtypes.to_string().split("\n"):
        coluna, tipo, *_ = re.split(r"\s+", colunaTipo)
        mapa[coluna] = tipo
    return mapa
@dataclass
class Arquivo:
    nome: str
    caminho: str
    funcao: str
    linha: int
def info_stack(index = 1) -> Arquivo:
    """Obter informações presente no stack dos callers.\n
    Padrão = Arquivo que chamou o info_stack()"""
    linha = stack()[index].lineno
    funcao = stack()[index].function
    filename = stack()[index].filename
    nome = filename[filename.rfind("\\") + 1 : ]
    caminho = filename[0 : filename.rfind("\\")]
    return Arquivo(nome, caminho, funcao, linha)

def to_json(item) -> str:
    """Retorna o `item` na forma de JSON"""
    return dumps(item, ensure_ascii=False, indent=4)

__all__ = [
    "remover_acentuacao",
    "normalizar",
    "mapear_dtypes",
    "info_stack",
    "to_json"
]