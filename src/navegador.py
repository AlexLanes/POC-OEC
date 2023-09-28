# std
from time import sleep
from typing import Literal, overload, Callable
# interno
from src.util import *
from src.logger import *
# externo
from selenium.webdriver.common.by import By
from selenium.webdriver import Ie, IeOptions
from selenium.common import TimeoutException
from selenium.webdriver.edge.service import Service
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support.wait import WebDriverWait as Wait

ESTRATEGIAS = Literal["id", "xpath", "link text", "name", "tag name", "class name", "css selector", "partial link text"]
CAMINHO_EDGE = "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe"

class Navegador:
    driver: Ie
    def __init__(self):
        """Iniciar o driver do Edge. Usar com o `with`"""
        self.options = IeOptions()
        self.options.attach_to_edge_chrome = True
        self.options.add_argument("--ignore-certificate-errors")
        self.options.edge_executable_path = CAMINHO_EDGE
        self.TIMEOUT = 120
    
    def __enter__(self):
        self.driver = Ie(self.options, Service())
        # self.driver.maximize_window() # BUG Travando o navegador 
        self.driver.implicitly_wait(self.TIMEOUT)
        Logger.informar("Navegador iniciado")
        return self

    def __exit__(self, *args):
        self.driver.quit()
        Logger.informar("Navegador fechado")

    @property
    def abas(self):
        """ID das abas abertas"""
        return self.driver.window_handles

    def pesquisar(self, url: str):
        """Pesquisar o url na aba focada"""
        Logger.informar(f"Pesquisado o url '{ url }'")
        self.driver.get(url)

    def nova_aba(self):
        """Abrir uma nova aba e alterar o foco para ela"""
        self.driver.switch_to.new_window("tab")
        Logger.informar("Aberto uma nova aba")

    def fechar_aba(self):
        """Fechar a aba focada e alterar o foco para a anterior"""
        titulo = self.driver.title
        self.driver.close()
        self.driver.switch_to.window(self.abas[-1])
        Logger.informar(f"Fechado a aba '{ titulo }'")
    
    def focar_aba(self, aba: str = None):
        """Focar na aba informada.\n
        Default é a última aba `self.abas[-1]`"""
        aba = self.abas[-1] if not aba else aba
        self.driver.switch_to.window(aba)
        Logger.informar(f"O navegador focou na aba '{ self.driver.title }'")

    @overload
    def encontrar(self, estrategia: ESTRATEGIAS, localizador: str) -> WebElement | None:
        """Encontrar o primeiro elemento na aba atual"""
    @overload
    def encontrar(self, estrategia: ESTRATEGIAS, localizador: str, primeiro=False) -> list[WebElement] | None:
        """Encontrar os elementos na aba atual"""
    def encontrar(self, estrategia: str, localizador: str, primeiro=True) -> WebElement | list[WebElement] | None:
        Logger.informar(f"Procurando { '1 elemento' if primeiro else '+1 elementos' } no navegador: ('{ estrategia }', '{ localizador }')")
        elementos = self.driver.find_elements(estrategia, localizador)
        Logger.informar(f"Encontrado { len(elementos) } elemento(s)")
        if len(elementos) == 0: return None
        return elementos[0] if primeiro else elementos
        
    def aguardar(self, condicao: Callable[[], bool], mensagemErro: str = None):
        """Repete a condição até que resulte em `True` ou `TimeoutException` com a `mensagemErro`"""
        Logger.informar(f"Aguardando uma condição")
        Wait(self.driver, self.TIMEOUT / 2).until(lambda _: condicao(), mensagemErro)
        Logger.informar(f"Condição atendida")

__all__ = [
    "Navegador",
    "TimeoutException"
]