# std
from re import search
from time import sleep
from enum import Enum, unique
# interno
from src.util import *
from src.logger import *
from src.windows import *
from src.navegador import *
from src.planilhas import *

CAMINHO_EXCEL = rf"{ info_stack().caminho }\documentos\RECURSOS_DEPARTAMENTOS.xlsx"
SITE_EBS = "https://o2qa.odebrecht.com"
USUARIO = "dclick"
SENHA = "senha123"

@unique
class Localizadores(Enum):
    recurso = "a#N55"
    texto_login = "Login"
    texto_home = "Home Page"
    senha = "input#passwordField"
    usuario = "input#usernameField"
    navegacao_dclick = "a#AppsNavLink"
    efetuar_login = "button#SubmitButton"
    texto_aplicativo_oracle = "Aplicativos Oracle"

@unique
class Imagens(Enum):
    recursos = "./screenshots/aba_recursos.png"
    organizacoes = "./screenshots/aba_organizacoes.png"
    erro_para_fechar = "./screenshots/erro_para_fechar.png"
    botoes_padroes_abas = "./screenshots/botoes_padroes_abas.png"
    recursos_custeado_desmarcado = "./screenshots/aba_recursos_custeado.png"
    recursos_processamento_desmarcado = "./screenshots/aba_recursos_processamento_externo.png"

@unique
class Offsets(Enum):
    recursos_udm = (0.9, 0.16)
    recursos_tipo = (0.3, 0.16)
    organizacoes_ok = (0.58, 0.9)
    erro_para_fechar = (0.8, 0.2)
    recursos_taxas = (0.45, 0.69)
    recursos_recurso = (0.3, 0.07)
    organizacoes_acn = (0.06, 0.56)
    botoes_padroes_abas = (0.8, 0.5)
    recursos_descricao = (0.3, 0.115)
    recursos_custeado = (0.037, 0.495)
    recursos_taxa_padrao = (0.4, 0.53)
    recursos_tipo_encargo = (0.3, 0.205)
    recursos_conta_absorcao = (0.3, 0.58)
    recursos_conta_variacao = (0.3, 0.629)
    recursos_processamento_externo = (0.037, 0.34)
    recursos_processamento_externo_item = (0.15, 0.40)

def fechar_possiveis_erros() -> bool:
    """Fechar possíveis erros na tela retornando um `bool` informando caso algum tenha sido encontrado e fechado"""
    erroAtual, errosMaximos = 1, 4
    while erroAtual <= errosMaximos:
        erro_para_fechar = Windows.procurar_imagem(Imagens.erro_para_fechar.value, "0.7", 2)
        if erro_para_fechar: 
            Windows.clicar_mouse( erro_para_fechar.transformar(*Offsets.erro_para_fechar.value) )
            erroAtual = erroAtual + 1
        else: break
    return erroAtual > 1

def preencher_recurso(recurso: Recurso) -> bool:
    """Preencher a aba Recursos com os dados do `Recurso`.\n
    Retornar um `bool` indicando se o recurso foi preenchido corretamente"""
    Logger.informar("Iniciado o preenchimento de um Recurso")
    coordenadas = Windows.procurar_imagem(Imagens.recursos.value, segundosProcura=10)
    assert coordenadas != None, "Aba dos Recursos não encontrada"
    
    # recurso
    # necessário validar se o recurso já existe
    Windows.clicar_mouse( coordenadas.transformar(*Offsets.recursos_recurso.value) )
    Windows.digitar(recurso.Geral.recurso)
    Windows.atalho(["tab"])
    if fechar_possiveis_erros():
        Logger.avisar(f"Este recurso foi ignorado devido a falha do campo 'Geral.recurso'. Recurso: { to_json(recurso.__dict__()) }")
        return False
    
    # descricao
    Windows.clicar_mouse( coordenadas.transformar(*Offsets.recursos_descricao.value) )
    Windows.digitar(recurso.Geral.descricao)
    
    # tipo
    # necessita validação se o tipo do recurso existe
    Windows.clicar_mouse( coordenadas.transformar(*Offsets.recursos_tipo.value) )
    match normalizar(recurso.Geral.tipo):
        case "pessoa": pass
        case "diversos": Windows.atalho(["up"])
        case "maquina": Windows.atalho(["up", "up"])
        case "moeda": Windows.atalho(["up", "up", "up"])
        case "quantia": Windows.atalho(["up", "up", "up", "up"])
        case _:
            Logger.avisar(f"Este recurso foi ignorado devido a falha do campo 'Geral.tipo'. Recurso: { to_json(recurso.__dict__()) }")
            return False
    Windows.atalho(["enter"])
    
    # tipo de encargo
    # necessita validação se o tipo de encargo do recurso existe
    Windows.clicar_mouse( coordenadas.transformar(*Offsets.recursos_tipo_encargo.value) )
    match normalizar(recurso.Geral.tipoEncargo):
        case "movimentacao_de_wip": pass
        case "movimentacao_de_oc": Windows.atalho(["up"])
        case "recebimento_da_oc": Windows.atalho(["up", "up"])
        case "manual": Windows.atalho(["up", "up", "up"])
        case _:
            Logger.avisar(f"Este recurso foi ignorado devido a falha do campo 'Geral.tipoEncargo'. Recurso: { to_json(recurso.__dict__()) }")
            return False
    Windows.atalho(["enter"])
    
    # udm 
    # necessita validação se o campo foi aceito ou está vazio
    # TODO - implementar uma checagem de mais pixeis ao redor do mouse na hora de checar pelo azul. O texto inserido dentro pode ficar na frente do azul
    Windows.clicar_mouse( coordenadas.transformar(*Offsets.recursos_udm.value) )
    Windows.digitar(recurso.Geral.udm)
    Windows.atalho(["tab"])
    if Windows.rgb_mouse()[2] == 255 or recurso.Geral.udm == "":
        Logger.avisar(f"Este recurso foi ignorado devido a falha do campo 'Geral.udm'. Recurso: { to_json(recurso.__dict__()) }")
        return False
    
    # processamento externo
    # necessário ativar para preencher
    # necessário desativar caso estiver ativo e não houver processamento externo
    if recurso.ProcessamentoExterno.processamentoExterno:
        # ativar
        processamentoDesmarcado = Windows.procurar_imagem(Imagens.recursos_processamento_desmarcado.value, "0.95")
        if processamentoDesmarcado: Windows.clicar_mouse( coordenadas.transformar(*Offsets.recursos_processamento_externo.value) )
        # item
        # necessita validação se o campo foi aceito
        Windows.clicar_mouse( coordenadas.transformar(*Offsets.recursos_processamento_externo_item.value) )
        Windows.digitar(recurso.ProcessamentoExterno.item)
        Windows.atalho(["tab"])
        if fechar_possiveis_erros():
            Logger.avisar(f"Este recurso foi ignorado devido a falha do campo 'ProcessamentoExterno.item'. Recurso: { to_json(recurso.__dict__()) }")
            return False
    else:
        # desativar
        processamentoDesmarcado = Windows.procurar_imagem(Imagens.recursos_processamento_desmarcado.value, "0.95")
        if not processamentoDesmarcado: Windows.clicar_mouse( coordenadas.transformar(*Offsets.recursos_processamento_externo.value) )
    
    # custeado
    # necessário ativar para preencher
    # necessário desativar caso estiver ativo e não for custeado
    if recurso.Faturamento.custeado:
        # ativar
        custeadoDesmarcado = Windows.procurar_imagem(Imagens.recursos_custeado_desmarcado.value, "0.95")
        if custeadoDesmarcado: Windows.clicar_mouse( coordenadas.transformar(*Offsets.recursos_custeado.value) )
        # conta de absorção
        Windows.clicar_mouse( coordenadas.transformar(*Offsets.recursos_conta_absorcao.value) )
        Windows.digitar(recurso.Faturamento.contaAbsorcao)
        Windows.atalho(["tab"])
        if fechar_possiveis_erros() or recurso.Faturamento.contaAbsorcao == "":
            Logger.avisar(f"Este recurso foi ignorado devido a falha do campo 'Faturamento.contaAbsorcao'. Recurso: { to_json(recurso.__dict__()) }")
            return False
        # conta de variação
        Windows.clicar_mouse( coordenadas.transformar(*Offsets.recursos_conta_variacao.value) )
        Windows.digitar(recurso.Faturamento.contaVariacao)
        if recurso.Faturamento.contaVariacao != "": 
            Windows.atalho(["tab"])
        if fechar_possiveis_erros():
            Logger.avisar(f"Este recurso foi ignorado devido a falha do campo 'Faturamento.contaVariacao'. Recurso: { to_json(recurso.__dict__()) }")
            return False
        # taxas
        if recurso.Faturamento.taxaPadrao and len(recurso.Taxas) > 0:
            # ativar e abrir a tela das taxas
            Windows.clicar_mouse( coordenadas.transformar(*Offsets.recursos_taxa_padrao.value) )
            Windows.clicar_mouse( coordenadas.transformar(*Offsets.recursos_taxas.value) )
            # necessário preencher as taxas
            # validações necessárias para os campos que serão preenchidos
            for taxa in recurso.Taxas:
                if taxa.tipoCusto == "" or not search(r"^\d+(,\d+)?$", taxa.custoUnitarioRecurso):
                    Logger.avisar(f"A taxa no index '{ recurso.Taxas.index(taxa) }' foi ignorada devido a uma falha de validação. Recurso: { to_json(recurso.__dict__()) }")
                    continue
                # tipo de custo
                Windows.digitar(taxa.tipoCusto)
                Windows.atalho(["tab"])
                if fechar_possiveis_erros():
                    Windows.atalho(["ctrl", "up"]) # limpar registro
                    Logger.avisar(f"Uma taxa foi ignorada devido a falha do campo 'Taxas[{ recurso.Taxas.index(taxa) }].tipoCusto'. Recurso: { to_json(recurso.__dict__()) }")
                    continue
                # custo unitário
                Windows.digitar(taxa.custoUnitarioRecurso)
                Windows.atalho(["tab"])
            # fechar a aba das taxas
            # SHIFT + PageUp Não funciona por algum motivo
            # retornar o foco para a aba dos recursos
            botoes = Windows.procurar_imagem(Imagens.botoes_padroes_abas.value, "0.8", 2)
            Windows.clicar_mouse( botoes.transformar(*Offsets.botoes_padroes_abas.value) )
    else:
        # desativar
        custeadoDesmarcado = Windows.procurar_imagem(Imagens.recursos_custeado_desmarcado.value, "0.95")
        if not custeadoDesmarcado: Windows.clicar_mouse( coordenadas.transformar(*Offsets.recursos_custeado.value) )
    
    Logger.informar(f"Finalizado o preenchido do Recurso '{ recurso.Geral.recurso }'")
    return True

def abrir_organizacao_acn():
    """Clicar na organização Código 'ACN' e depois em 'OK'"""
    coordenadas = Windows.procurar_imagem(Imagens.organizacoes.value, segundosProcura=30)
    assert coordenadas != None, "Aba das organizações não encontrada"
    Windows.clicar_mouse( coordenadas.transformar(*Offsets.organizacoes_acn.value) )
    Windows.clicar_mouse( coordenadas.transformar(*Offsets.organizacoes_ok.value) )

def fechar_aplicativo_oracle(aplicativoOracle: Janela):
    """Fechar a janela aberta do Aplicativo Oracle e aguardar o Edge voltar ao foco"""
    Logger.informar("Fechando o Aplicativo Oracle")
    tituloAplicativoOracle = aplicativoOracle.titulo()
    
    aplicativoOracle.fechar()
    if Windows.titulo_janela_focada().lower() == tituloAplicativoOracle.lower():
        Windows.atalho(["enter"])
    
    Windows.aguardar(
        lambda: tituloAplicativoOracle.lower() != Windows.titulo_janela_focada(),
        "O Aplicativo Oracle não foi fechado corretamente"
    )
    Logger.informar("Aplicativo Oracle fechado")

def abrir_aplicativo_oracle(navegador: Navegador):
    """Clicar em `AUTOMACAO DCLICK`, `Recursos` e esperar o aplicativo oracle ficar focado"""
    Logger.informar("Abrindo o Aplicativo Oracle")
    # aba "AUTOMACAO DCLICK"
    elemento = navegador.encontrar("css selector", Localizadores.navegacao_dclick.value)
    assert elemento != None, "Navegação 'AUTOMACAO DCLICK' não encontrada"
    elemento.click()
    # elemento "Recurso"
    elemento = navegador.encontrar("css selector", Localizadores.recurso.value)
    assert elemento != None, "Elemento 'Recurso' não encontrado"
    elemento.click()
    # aguardar o aplicativo oracle
    Windows.aguardar(
        lambda: Localizadores.texto_aplicativo_oracle.value.lower() in Windows.titulo_janela_focada().lower(),
        f"O texto '{ Localizadores.texto_aplicativo_oracle.value }' não foi encontrado no titulo da janela focada"
    )
    Logger.informar("Aplicativo Oracle aberto")

def efetuar_login(navegador: Navegador):
    """Efetuar o login no `SITE_EBS` e esperar a página Home carregar"""
    Logger.informar("Efetuando o login")
    navegador.pesquisar(SITE_EBS)
    navegador.aguardar(
        lambda: Localizadores.texto_login.value.lower() in navegador.driver.title.lower(), 
        f"Texto '{ Localizadores.texto_login.value }' não encontrado no título do navegador"
    )
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
    navegador.aguardar(
        lambda: Localizadores.texto_home.value.lower() in navegador.driver.title.lower(), 
        f"Texto '{ Localizadores.texto_home.value }' não encontrado no título do navegador"
    )
    Logger.informar("Login efetuado e Home Page carregada")

def main():
    """Fluxo principal.\n
    Mapeamentos passo a passo se encontra no arquivo '../documentos/Mapeamentos de automação.txt'"""
    recursos = parse_recursos(CAMINHO_EXCEL)
    departamentos = parse_departamentos(CAMINHO_EXCEL)

    # # TODO - nome de recurso aleatorio para não dar conflito
    # from uuid import uuid4
    # for recurso in recursos:
    #     recurso.Geral.recurso = uuid4().__str__().split("-")[0]
    
    try:
        # abrir navegador no modo Internet Explorer
        with Navegador() as navegador:
            # maximizar janela do navegador
            # efetuar login no `SITE_EBS`
            Windows.janela_focada().maximizar()
            efetuar_login(navegador)
            
            # abrir o aplicativo oracle e maximiza-lo
            # aguardar um tempo para iniciar corretamente
            abrir_aplicativo_oracle(navegador)
            sleep(20)
            janelaAplicativoOracle = Windows.janela_focada()
            janelaAplicativoOracle.maximizar()
            
            # abrir as organizações Código 'ACN'
            abrir_organizacao_acn()
            
            # realizar o preenchimento de cada recurso
            for recurso in recursos:
                preenchidoSemErro = preencher_recurso(recurso)
                if preenchidoSemErro: Windows.atalho(["ctrl", "s"])
                # limpar para o próximo recurso
                Windows.atalho(["f6"])
            
            # fechar o Aplicativo Oracle sutilmente
            fechar_aplicativo_oracle(janelaAplicativoOracle)

    except (TimeoutException, TimeoutError) as erro:
        Logger.erro(f"Erro de timeout na espera de alguma condição/elemento/janela: { erro }")
        exit(1)
    except AssertionError as erro:
        Logger.erro(f"Erro de validação pré-execução de algum passo no fluxo: { erro }")
        exit(1)
    except Exception as erro:
        Logger.erro(f"Erro inesperado no fluxo: { erro }")
        exit(1)
    
if __name__ == "__main__":
    Logger.informar("--- Iniciado execução do fluxo ---")
    main()
    Logger.informar("--- Finalizado execução com sucesso ---")