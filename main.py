# std
import re
from time import sleep
from enum import Enum, unique
# interno
from src.util import *
from src.logger import *
from src.windows import *
from src.navegador import *
from src.planilhas import *

CAMINHO_EXCEL = rf"{ informacoes_filename()[0] }\documentos\RECURSOS_DEPARTAMENTOS.xlsx"
SITE_EBS = "https://o2qa.odebrecht.com"
USUARIO = "dclick"
SENHA = "senha123"

@unique
class Localizadores(Enum):
    usuario = "input#usernameField"
    senha = "input#passwordField"
    efetuar_login = "button#SubmitButton"
    texto_home = "Home Page"
    navegacao_dclick = "a#AppsNavLink"
    navegacao_dclick_table = "table.x9t"
    recurso = "a#N55"

@unique
class Imagens(Enum):
    organizacoes = "./screenshots/aba_organizacoes.png"
    recursos = "./screenshots/aba_recursos.png"

@unique
class Offsets(Enum):
    organizacoes_acn = (0.06, 0.56)
    organizacoes_ok = (0.58, 0.9)
    recursos_recurso = (0.3, 0.07)
    recursos_descricao = (0.3, 0.115)
    recursos_udm = (0.9, 0.16)
    recursos_processamento_externo = (0.04, 0.34)
    recursos_processamento_externo_item = (0.15, 0.40)

def preencher_recurso(recurso: Recurso):
    """Preencher a aba Recursos com os dados do `Recurso`"""
    coordenadas = Windows.procurar_imagem(Imagens.recursos.value, segundosProcura=10)
    assert coordenadas != None, "Aba dos Recursos não encontrada"
    # recurso
    Windows.clicar_mouse( coordenadas.transformar(*Offsets.recursos_recurso.value) )
    Windows.digitar(recurso.Geral.recurso)
    # descricao
    Windows.clicar_mouse( coordenadas.transformar(*Offsets.recursos_descricao.value) )
    Windows.digitar(recurso.Geral.descricao)
    # tipo
    # tipo de encargo
    # udm
    Windows.clicar_mouse( coordenadas.transformar(*Offsets.recursos_udm.value) )
    Windows.digitar(recurso.Geral.descricao)
    # processamento externo
    if recurso.ProcessamentoExterno.processamentoExterno:
        Windows.clicar_mouse( coordenadas.transformar(*Offsets.recursos_processamento_externo.value) )
        # item
        Windows.clicar_mouse( coordenadas.transformar(*Offsets.recursos_processamento_externo_item.value) )
        Windows.digitar(recurso.ProcessamentoExterno.item)
    
def abrir_organizacao_acn():
    """Clicar na organização Código 'ACN' e depois em 'OK'"""
    coordenadas = Windows.procurar_imagem(Imagens.organizacoes.value, segundosProcura=10)
    assert coordenadas != None, "Aba das organizações não encontrada"
    Windows.clicar_mouse( coordenadas.transformar(*Offsets.organizacoes_acn.value) )
    Windows.clicar_mouse( coordenadas.transformar(*Offsets.organizacoes_ok.value) )

def efetuar_login(navegador: Navegador):
    """Efetuar o login no `SITE_EBS`"""
    Logger.informar("Efetuando o login")
    navegador.pesquisar(SITE_EBS)
    navegador.aguardar(lambda: "login" in navegador.driver.title.lower())
    # usuario
    elemento = navegador.encontrar("css selector", Localizadores.usuario.value)
    assert elemento != None, "Campo do usuario não encontrado"
    elemento.send_keys(USUARIO)
    # senha
    elemento = navegador.encontrar("css selector", Localizadores.senha.value)
    assert elemento != None, "Campo de senha não encontrado"
    elemento.send_keys(SENHA)
    # efetuar login
    elemento = navegador.encontrar("css selector", Localizadores.efetuar_login.value)
    assert elemento != None, "Botão para efeutar login não encontrado"
    elemento.click()
    # aguardar a pagina 'Home' carregar
    navegador.aguardar(lambda: "home page" in navegador.driver.title.lower())
    Logger.informar("Login efetuado e Home Page carregada")

def main(navegador: Navegador, recursos: list[Recurso], departamentos: list[Departamento]):
    """Fluxo principal"""
    # efetuar_login(navegador)
    # abrir_organizacao_acn()
    # preencher_recurso(recursos[11])
    
if __name__ == "__main__":
    Logger.informar("Executando")
    recursos = parse_recursos(CAMINHO_EXCEL)
    departamentos = parse_departamentos(CAMINHO_EXCEL)
    
    try:
        with Navegador() as navegador:
            main(navegador, recursos, departamentos)
        Logger.informar("Finalizado execução com sucesso")

    except TimeoutException as erro:
        Logger.erro(f"Erro de timeout na espera de alguma condição/elemento: { erro }")
        exit(1)
    except AssertionError as erro:
        Logger.erro(f"Erro de validação pré-execução de algum passo no fluxo: { erro }")
        exit(1)
    except Exception as erro:
        Logger.erro(f"Erro inesperado no fluxo: { erro }")
        exit(1)

"""
    Processo passo a passo se encontra no arquivo "../documentos/Processo de automação.txt"
"""