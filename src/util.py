# std
import re
from inspect import stack
from unicodedata import normalize
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

def informacoes_filename(index = 1) -> tuple[str, str]:
    """Obter o caminho absoluto e o nome do arquivo presente no stack dos callers.\n
    Padrão = Arquivo que chamou o informacoes_filename()"""
    filename = stack()[index].filename
    nome = filename[filename.rfind("\\") + 1:]
    caminho = filename[0 : filename.rfind("\\")]
    return (caminho, nome)

__all__ = [
    "remover_acentuacao",
    "normalizar",
    "mapear_dtypes",
    "informacoes_filename"
]