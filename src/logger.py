# std
import logging
from sys import exc_info
from inspect import stack
# interno
from src.util import *

class Logger:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | id(%(process)d) | level(%(levelname)s) | %(message)s",
        datefmt="%Y-%m-%dT%H:%M:%S",
        filename="execução.log",
        encoding="utf-8",
        filemode="w"
    )
    
    @staticmethod
    def informar(mensagem: str) -> None:
        """Log nível 'INFO'"""
        logging.info(f"arquivo({ informacoes_filename(2)[1] }) | função({ stack()[1].function }) | { mensagem }")
    @staticmethod
    def avisar(mensagem: str) -> None:
        """Log nível 'WARNING'"""
        logging.warning(f"arquivo({ informacoes_filename(2)[1] }) | função({ stack()[1].function }) | { mensagem }")
    @staticmethod
    def erro(mensagem: str) -> None:
        """Log nível 'ERROR'"""
        logging.error(f"arquivo({ informacoes_filename(2)[1] }) | função({ stack()[1].function }) | { mensagem }", exc_info=exc_info())

__all__ = [
    "Logger"
]