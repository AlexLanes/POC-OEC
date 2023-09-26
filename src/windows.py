# std
from time import sleep, time
from dataclasses import dataclass
from typing import overload, Literal
# interno
from src.util import *
# externo
import pyautogui as AutoGui
from pygetwindow import Win32Window

AutoGui.PAUSE = 0.5
AutoGui.FAILSAFE = True
MOUSE = Literal["left", "middle", "right"]
PORCENTAGENS = Literal["0.9", "0.8", "0.7", "0.6", "0.5", "0.4", "0.3", "0.2", "0.1"]
TECLADO = Literal['\t', '\n', '\r', ' ', '!', '"', '#', '$', '%', '&', "'", '(', ')', '*', '+', ',', '-', '.', '/', '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', ':', ';', '<', '=', '>', '?', '@', '[', '\\', ']', '^', '_', '`', 'a', 'b', 'c', 'd', 'e','f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z', '{', '|', '}', '~', 'accept', 'add', 'alt', 'altleft', 'altright', 'apps', 'backspace', 'browserback', 'browserfavorites', 'browserforward', 'browserhome', 'browserrefresh', 'browsersearch', 'browserstop', 'capslock', 'clear', 'convert', 'ctrl', 'ctrlleft', 'ctrlright', 'decimal', 'del', 'delete', 'divide', 'down', 'end', 'enter', 'esc', 'escape', 'execute', 'f1', 'f10', 'f11', 'f12', 'f13', 'f14', 'f15', 'f16', 'f17', 'f18', 'f19', 'f2', 'f20', 'f21', 'f22', 'f23', 'f24', 'f3', 'f4', 'f5', 'f6', 'f7', 'f8', 'f9', 'final', 'fn', 'hanguel', 'hangul', 'hanja', 'help', 'home', 'insert', 'junja', 'kana', 'kanji', 'launchapp1', 'launchapp2', 'launchmail', 'launchmediaselect', 'left', 'modechange', 'multiply', 'nexttrack', 'nonconvert', 'num0', 'num1', 'num2', 'num3', 'num4', 'num5', 'num6', 'num7', 'num8', 'num9', 'numlock', 'pagedown', 'pageup', 'pause', 'pgdn', 'pgup', 'playpause', 'prevtrack', 'print', 'printscreen', 'prntscrn', 'prtsc', 'prtscr', 'return', 'right', 'scrolllock', 'select', 'separator', 'shift', 'shiftleft', 'shiftright', 'sleep', 'space', 'stop', 'subtract', 'tab', 'up', 'volumedown', 'volumemute', 'volumeup', 'win', 'winleft', 'winright', 'yen', 'command', 'option', 'optionleft', 'optionright']

@dataclass
class Coordenadas:
    x: int
    y: int
    comprimento: int
    altura: int

    def __iter__(self):
        """Utilizar com o tuple()"""
        yield self.x
        yield self.y
        yield self.comprimento
        yield self.altura
    
    @overload
    def transformar(self) -> tuple[int, int]:
        """Transformar as cordenadas para a posição (x, y) central horizontalmente e verticalmente"""
    @overload
    def transformar(self, xOffset: float, yOffset: float) -> tuple[int, int]:
        """Transformar as cordenadas para a posição (x, y) de acordo com a porcentagem xOffset e yOffset.\n
        `xOffset` (esquerda, centro, direita) = (0, 0.5, 1)\n
        `yOffset` (topo, centro, baixo) = (0, 0.5, 1)"""
    def transformar(self, xOffset=0.5, yOffset=0.5) -> tuple[int, int]:
        # enforça o range entre 0.0 e 1.0
        xOffset, yOffset = max(0.0, min(1.0, xOffset)), max(0.0, min(1.0, yOffset))
        return ( 
            self.x + int(self.comprimento * xOffset), 
            self.y + int(self.altura * yOffset) 
        )

    def __len__(self):
        return 4

def obter_x_y(coordenadas: tuple[int, int] | Coordenadas) -> tuple[int|None, int|None]:
    """Obter coordenadas (x, y) do item recebido"""
    x = y = None
    if isinstance(coordenadas, tuple): 
        x, y = coordenadas
    elif isinstance(coordenadas, Coordenadas):
        x, y = coordenadas.transformar()
    return (x, y)

@dataclass
class Janela:
    janela: Win32Window
    def maximizar(self) -> None:
        """Maximizar janela"""
        self.janela.maximize()
    def maximizado(self) -> bool:
        """Checar se a janela está maximizada"""
        self.janela.isMaximized
    def minimizar(self) -> None:
        """Minimizar janela"""
        self.janela.minimize()
    def minimizado(self) -> bool:
        """Checar se a janela está minimizada"""
        self.janela.isMinimized
    def fechar(self) -> None:
        """Fechar janela"""
        self.janela.close()
    def focado(self) -> bool:
        """Checar se a janela está focada"""
        return self.janela.isActive
    def focar(self) -> None:
        """Focar a janela minimizando e maximizando"""
        if not self.focado():
            self.minimizar()
            self.maximizar()

class Windows:
    @staticmethod
    def procurar_imagem(caminhoImagem: str, porcentagemConfianca: PORCENTAGENS = "0.9", segundosProcura=1) -> Coordenadas | None:
        """Procurar imagem na tela com `porcentagemConfianca`% de confiança na procura durante `segundosProcura` segundos"""
        box = AutoGui.locateOnScreen(
            image=caminhoImagem, 
            minSearchTime=segundosProcura,
            confidence=porcentagemConfianca
        )
        return Coordenadas(box.left, box.top, box.width, box.height) if box else None
        
    @staticmethod
    def titulos_janelas() -> list[str]:
        """Listar os titulos das janelas abertas"""
        return [ titulo for titulo in AutoGui.getAllTitles() if titulo != "" ]

    @staticmethod
    def titulo_janela_focada() -> str:
        """Obter o titulo da janela em foque"""
        return AutoGui.getActiveWindowTitle()
    
    @staticmethod
    def janela_focada() -> Janela:
        """Obter a janela focada"""
        janela: Win32Window = AutoGui.getActiveWindow()
        return Janela(janela)
    
    @staticmethod
    def buscar_janela(titulo: str) -> Janela | None:
        """Obter a primeira janela que possua o `titulo`"""
        janelas: list[Win32Window] = AutoGui.getAllWindows()
        for janela in janelas:
            if titulo.lower() in janela.title.lower(): return Janela(janela)

    @staticmethod
    def aguardar_janela(titulo: str, timeout=10) -> Janela | TimeoutError:
        """Aguardar até que uma janela de título `titulo` esteja aberta ou `TimeoutError` após `timeout` segundos"""
        inicio = time()
        while True:
            if time() - inicio >= timeout: 
                raise TimeoutError(f"Janela de título '{ titulo }' não foi encontrada depois de { timeout } segundos")
            janela = Windows.buscar_janela(titulo)
            if janela: return janela
            else: sleep(0.25)
    
    @staticmethod
    def mover_mouse(coordenadas: tuple[int, int] | Coordenadas) -> None:
        """Mover o mouse até as cordenadas"""
        x, y = obter_x_y(coordenadas)
        AutoGui.moveTo(x, y)
    
    @overload
    @staticmethod
    def clicar_mouse() -> None:
        """Clicar com o botão esquerdo do mouse 1 vez na posição atual"""
    @overload
    @staticmethod
    def clicar_mouse(coordenadas: Coordenadas | tuple[int, int]) -> None:
        """Clicar com o botão esquerdo do mouse 1 vez nas `coordenadas`"""
    @overload
    @staticmethod
    def clicar_mouse(coordenadas: Coordenadas | tuple[int, int], botao: MOUSE, quantidade=1) -> None:
        """Clicar com o botão `botao` do mouse `quantidade` vez(es) nas `coordenadas`"""
    def clicar_mouse(coordenadas: Coordenadas | tuple[int, int] = None, botao="left", quantidade=1) -> None:
        x, y = obter_x_y(coordenadas)
        AutoGui.click(x, y, quantidade, 0.25, botao)
    
    @staticmethod
    def rgb_mouse() -> tuple[int, int, int]:
        """Obter o RGB das coordenadas atual do mouse"""
        return AutoGui.pixel( *AutoGui.position() )
    
    @staticmethod
    def atalho(teclas: list[TECLADO]) -> None:
        """Apertar as teclas sequencialmente e depois soltá-las em ordem reversa"""
        AutoGui.hotkey(teclas, interval=0.25)
    
    @staticmethod
    def digitar(texto: str, delay = 0.01) -> None:
        """Digitar o texto pressionando cada tecla do texto e soltando em seguida a cada `delay` segundos"""
        AutoGui.write(texto, delay)

__all__ = [
    "Windows",
    "AutoGui"
]
