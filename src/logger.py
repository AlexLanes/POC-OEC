# std
import logging
from sys import exc_info
# interno
from src.util import *

ENCODING = "utf-8"
NOME_ARQUIVO = "execução.log"
FORMATO_DATA = "%Y-%m-%dT%H:%M:%S"

class Logger:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | id(%(process)d) | level(%(levelname)s) | %(message)s",
        datefmt=FORMATO_DATA,
        filename=NOME_ARQUIVO,
        encoding=ENCODING,
        filemode="a"
    )
    
    @staticmethod
    def informar(mensagem: str) -> None:
        """Log nível 'INFO'"""
        logging.info(f"arquivo({ info_stack(2).nome }) | função({ info_stack(2).funcao }) | linha({ info_stack(2).linha }) | { mensagem }")
    @staticmethod
    def avisar(mensagem: str) -> None:
        """Log nível 'WARNING'"""
        logging.warning(f"arquivo({ info_stack(2).nome }) | função({ info_stack(2).funcao }) | linha({ info_stack(2).linha }) | { mensagem }")
    @staticmethod
    def erro(mensagem: str) -> None:
        """Log nível 'ERROR'"""
        logging.error(f"arquivo({ info_stack(2).nome }) | função({ info_stack(2).funcao }) | linha({ info_stack(2).linha }) | { mensagem }", exc_info=exc_info())

__all__ = [
    "Logger"
]

"""Remover linhas de log acima do `TEMPO_MAXIMO`
Executar como `python -m src.logger` na pasta raiz"""
if __name__ == "__main__":
    import re
    from datetime import datetime, timedelta
    TEMPO_MAXIMO: timedelta = timedelta(hours=1)
    Logger.informar("Buscando por Logs que estão acima do tempo máximo configurado")

    agora = datetime.now()
    linhasParaRemover: int = 0
    with open(NOME_ARQUIVO, "r+", encoding=ENCODING) as log:
        linhas: list[str] = log.readlines()
        for linha in linhas:
            dataHora = re.search(r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}", linha)
            if not dataHora: 
                linhasParaRemover = linhasParaRemover + 1
                continue
            dataHora = datetime.strptime(dataHora.group(), FORMATO_DATA)
            if dataHora + TEMPO_MAXIMO < agora: 
                linhasParaRemover = linhasParaRemover + 1
            else: break
        
        if linhasParaRemover > 0:
            log.seek(0)     # mover cursor pro começo do arquivo
            log.truncate()  # limpar arquivo
            log.writelines(linhas[linhasParaRemover:]) # linhas fora do tempo máximo
        
    Logger.informar(f"{ 'Nenhuma linha removida' if linhasParaRemover == 0 else f'Removido { linhasParaRemover } linhas' } do arquivo '{ NOME_ARQUIVO }' após o tempo máximo '{ TEMPO_MAXIMO }'")