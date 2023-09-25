# std
from typing import Literal, overload
# interno
from src.logger import *
# externo
from selenium.webdriver.common.by import By
from selenium.webdriver import Ie, IeOptions
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.service import Service
from selenium.webdriver.remote.webelement import WebElement
import selenium.webdriver.support.expected_conditions as Expect
from selenium.webdriver.support.wait import WebDriverWait as Wait

ESTRATEGIAS = Literal["id", "xpath", "link text", "name", "tag name", "class name", "css selector", "partial link text"]
CAMINHO_EDGE = "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe"

class Navegador:
    navegador: Ie
    def __init__(self):
        """Iniciar o driver do Edge. Usar com o `with`"""
        self.options = IeOptions()
        self.options.attach_to_edge_chrome = True
        self.options.add_argument("--ignore-certificate-errors")
        self.options.edge_executable_path = CAMINHO_EDGE
        self.TIMEOUT = 30
    
    def __enter__(self):
        self.navegador = Ie(self.options, Service())  
        self.navegador.maximize_window()
        self.navegador.implicitly_wait(self.TIMEOUT)
        # fechar a pagina aberta automaticamento e abrir uma nova
        self.navegador.switch_to.new_window("tab")
        self.navegador.switch_to.window(self.abas[0])
        self.navegador.close()
        self.navegador.switch_to.window(self.abas[0])
        Logger.informar("Navegador iniciado")
        return self

    def __exit__(self, *args):
        self.navegador.quit()
        Logger.informar("Navegador fechado")

    @property
    def abas(self):
        """ID das abas abertas"""
        return self.navegador.window_handles

    def pesquisar(self, url: str):
        """Pesquisar o url na aba focada"""
        self.navegador.get(url)
        Logger.informar(f"Pesquisado o url '{ url }'")

    def nova_aba(self):
        """Abrir uma nova aba e alterar o foco para ela"""
        self.navegador.switch_to.new_window("tab")
        Logger.informar("Aberto uma nova aba")

    def fechar_aba(self):
        """Fechar a aba focada e alterar para a anterior"""
        titulo = self.navegador.title
        self.navegador.close()
        self.navegador.switch_to.window(self.abas[-1])
        Logger.informar(f"Fechado a aba '{ titulo }'")

    @overload
    def encontrar(self, estrategia: ESTRATEGIAS, localizador: str) -> WebElement | None:
        """Encontrar o primeiro elemento na aba atual"""
    @overload
    def encontrar(self, estrategia: ESTRATEGIAS, localizador: str, primeiro=False) -> list[WebElement] | None:
        """Encontrar os elementos na aba atual"""
    def encontrar(self, estrategia: str, localizador: str, primeiro=True) -> WebElement | list[WebElement] | None:
        Logger.informar(f"Procurando { '1 elemento' if primeiro else '+1 elementos' } no navegador: ('{ estrategia }', '{ localizador }')")
        elementos = self.navegador.find_elements(estrategia, localizador)
        Logger.informar(f"Encontrado { len(elementos) } elemento(s)")
        if len(elementos) == 0: return None
        return elementos[0] if primeiro else elementos

__all__ = [
    "Navegador"
]