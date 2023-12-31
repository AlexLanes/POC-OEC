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
    departamento = "a#N61"
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
    departamentos = "./screenshots/aba_departamentos.png"
    erro_para_fechar = "./screenshots/erro_para_fechar.png"
    botoes_padroes_abas = "./screenshots/botoes_padroes_abas.png"
    preocupacao = "./screenshots/preocupacao_departamentos_recursos.png"
    recursos_custeado_desmarcado = "./screenshots/aba_recursos_custeado.png"
    aba_departamentos_recursos = "./screenshots/aba_departamentos_recursos.png"
    erro_organizacoes_definidas = "./screenshots/erro_organizacoes_definidas.png"
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
    departamentos_local = (0.407, 0.48)
    recursos_tipo_encargo = (0.3, 0.205)
    departamentos_recursos = (0.87, 0.92)
    recursos_conta_absorcao = (0.3, 0.58)
    recursos_conta_variacao = (0.3, 0.629)
    departamentos_descricao = (0.45, 0.25)
    departamentos_departamento = (0.45, 0.172)
    departamentos_categoria_custo = (0.45, 0.33)
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

def preencher_departamento(departamento: Departamento) -> bool:
    """Preencher a aba Departamentos com os dados do `departamento`.\n
    Retornar um `bool` indicando se o departamento foi preenchido corretamente"""
    Logger.informar("Iniciado o preenchimento de um Departamento")
    coordenadas = Windows.procurar_imagem(Imagens.departamentos.value, segundosProcura=10)
    assert coordenadas != None, "Aba dos Departamentos não encontrada"
    
    # departamento
    # necessário validar se o departamento já existe ou está vazio
    Windows.clicar_mouse( coordenadas.transformar(*Offsets.departamentos_departamento.value) )
    Windows.digitar(departamento.Departamento.departamento)
    Windows.atalho(["tab"])
    if fechar_possiveis_erros() or departamento.Departamento.departamento == "":
        Logger.avisar(f"Este departamento foi ignorado devido a falha do campo 'Departamento.departamento'. Departamento: { to_json(departamento.__dict__()) }")
        return False
    
    # descrição
    Windows.clicar_mouse( coordenadas.transformar(*Offsets.departamentos_descricao.value) )
    Windows.digitar(departamento.Departamento.descricao)
    
    # categoria de custo
    # necessário validar se existe a opção
    Windows.clicar_mouse( coordenadas.transformar(*Offsets.departamentos_categoria_custo.value) )
    Windows.digitar(departamento.Departamento.categoria_de_custo)
    Windows.atalho(["tab"]) 
    if fechar_possiveis_erros():
        Logger.avisar(f"Este departamento foi ignorado devido a falha do campo 'Departamento.categoria_de_custo'. Departamento: { to_json(departamento.__dict__()) }")
        return False
    
    # local
    # necessário validar se existe a opção
    # TODO - implementar uma checagem de mais pixeis ao redor do mouse na hora de checar pelo azul. O texto inserido dentro pode ficar na frente do azul
    Windows.clicar_mouse( coordenadas.transformar(*Offsets.departamentos_local.value) )
    Windows.digitar(departamento.Departamento.local)
    Windows.atalho(["tab"]) 
    if fechar_possiveis_erros() or Windows.rgb_mouse()[2] == 255:
        print(Windows.rgb_mouse())
        Logger.avisar(f"Este departamento foi ignorado devido a falha do campo 'Departamento.local'. Departamento: { to_json(departamento.__dict__()) }")
        return False

    # recursos
    if len(departamento.Recursos) > 0:
        # abrir a tela dos recursos
        # aguardar carregar
        Windows.clicar_mouse( coordenadas.transformar(*Offsets.departamentos_recursos.value) )
        Windows.procurar_imagem(Imagens.aba_departamentos_recursos.value, segundosProcura=10)
        # preencher os recursos
        for recurso in departamento.Recursos:
            # abrir um novo registro (não necessário para o primeiro, mas para os próximos)
            # validação de campos e formatos
            Windows.atalho(["ctrl", "down"])
            if recurso.recurso == "" or not search(r"^\d+(,\d+)?$", recurso.unidades):
                Logger.avisar(f"O recurso no index '{ departamento.Recursos.index(recurso) }' foi ignorado devido a uma falha de validação. Departamento: { to_json(departamento.__dict__()) }")
                continue
            # recurso
            Windows.digitar(recurso.recurso)
            Windows.atalho(["tab"]) 
            if fechar_possiveis_erros():
                Logger.avisar(f"Um recurso foi ignorado devido a falha do campo 'Recursos[{ departamento.Recursos.index(recurso) }].recurso'. Departamento: { to_json(departamento.__dict__()) }")
                # limpar registro
                Windows.atalho(["ctrl", "up"]) 
                Windows.procurar_imagem(Imagens.preocupacao.value, segundosProcura=5)
                Windows.atalho(["enter"]) 
                continue
            # disponível 24 horas
            # padrão do checkbox é marcado
            if not recurso.disponibilizar_horas:
                Windows.atalho(["space"])
            # unidades
            Windows.atalho(["tab", "tab"])
            Windows.digitar(recurso.unidades)
        # fechar a aba dos recursos
        # SHIFT + PageUp Não funciona por algum motivo
        # retornar o foco para a aba dos departamentos
        botoes = Windows.procurar_imagem(Imagens.botoes_padroes_abas.value, "0.8", 2)
        Windows.clicar_mouse( botoes.transformar(*Offsets.botoes_padroes_abas.value) )

    Logger.informar(f"Finalizado o preenchido do Departamento '{ departamento.Departamento.departamento }'")
    return True

def preencher_recurso(recurso: Recurso) -> bool:
    """Preencher a aba Recursos com os dados do `Recurso`.\n
    Retornar um `bool` indicando se o recurso foi preenchido corretamente"""
    Logger.informar("Iniciado o preenchimento de um Recurso")
    coordenadas = Windows.procurar_imagem(Imagens.recursos.value, segundosProcura=10)
    assert coordenadas != None, "Aba dos Recursos não encontrada"
    
    # recurso
    # necessário validar se o recurso já existe ou está vazio
    Windows.clicar_mouse( coordenadas.transformar(*Offsets.recursos_recurso.value) )
    Windows.digitar(recurso.Geral.recurso)
    Windows.atalho(["tab"])
    if fechar_possiveis_erros() or recurso.Geral.recurso == "":
        Logger.avisar(f"Este recurso foi ignorado devido a falha do campo 'Geral.recurso'. Recurso: { to_json(recurso.__dict__()) }")
        return False
    
    # descrição
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
    """Clicar na organização Código 'ACN' e depois em 'OK'.\n
    Maximizar a janela do Aplicativo Oracle"""
    Logger.informar(f"Maximizando janela focada '{ Windows.titulo_janela_focada() }'")
    Windows.janela_focada().maximizar()

    # checando por erro fatal
    erro = Windows.procurar_imagem(Imagens.erro_organizacoes_definidas.value, segundosProcura=5)
    assert erro == None, "Erro das organizações não definidas encontrado" 
    
    Logger.informar("Abrindo a organização 'ACN'")
    coordenadas = Windows.procurar_imagem(Imagens.organizacoes.value, segundosProcura=60)
    assert coordenadas != None, "Aba das organizações não encontrada"
    Windows.clicar_mouse( coordenadas.transformar(*Offsets.organizacoes_acn.value) )
    Windows.clicar_mouse( coordenadas.transformar(*Offsets.organizacoes_ok.value) )
    
    sleep(1)
    Logger.informar("Organização 'ACN' aberta")

def fechar_aplicativo_oracle():
    """Fechar a janela aberta do Aplicativo Oracle e aguardar o Edge voltar ao foco"""
    Logger.informar("Fechando o Aplicativo Oracle")
    tituloAplicativoOracle = Windows.titulo_janela_focada()
    Windows.janela_focada().fechar()
    sleep(1)
    
    if Windows.titulo_janela_focada().lower() == tituloAplicativoOracle.lower():
        Windows.atalho(["enter"])
    
    Windows.aguardar(
        lambda: tituloAplicativoOracle.lower() != Windows.titulo_janela_focada(),
        "O Aplicativo Oracle não foi fechado corretamente"
    )
    Logger.informar("Aplicativo Oracle fechado")

def abrir_aplicativo_oracle_departamento(navegador: Navegador):
    """Alterar a aba para a da `Home Page`, clicar em `Departamento` e esperar o aplicativo oracle ficar focado"""
    Logger.informar("Abrindo o Aplicativo Oracle")
    # fechar a aba aberta pelo Recurso e retornar para a Home Page
    navegador.focar_aba()
    navegador.fechar_aba()
    sleep(1)
    # elemento "Departamento"
    elemento = navegador.encontrar("css selector", Localizadores.departamento.value)
    assert elemento != None, "Elemento 'Departamento' não encontrado"
    elemento.click()
    # aguardar o Aplicativo Oracle
    Windows.aguardar(
        lambda: "edge" not in Windows.titulo_janela_focada().lower() and Localizadores.texto_aplicativo_oracle.value.lower() in Windows.titulo_janela_focada().lower(),
        f"O Aplicativo Oracle não foi inicializado corretamente"
    )
    Logger.informar(f"Aplicativo Oracle aberto. Título '{ Windows.titulo_janela_focada() }'")
    # TODO implementar a procura pelo erro do "../screenshots/erro_java_tm_blocked.png"
    # informar os passos necessários para resolver o erro

def abrir_aplicativo_oracle_recurso(navegador: Navegador):
    """Clicar em `AUTOMACAO DCLICK`, `Recurso` e esperar o aplicativo oracle ficar focado"""
    Logger.informar("Abrindo o Aplicativo Oracle")
    # aba "AUTOMACAO DCLICK"
    elemento = navegador.encontrar("css selector", Localizadores.navegacao_dclick.value)
    assert elemento != None, "Navegação 'AUTOMACAO DCLICK' não encontrada"
    elemento.click()
    # elemento "Recurso"
    elemento = navegador.encontrar("css selector", Localizadores.recurso.value)
    assert elemento != None, "Elemento 'Recurso' não encontrado"
    elemento.click()
    # aguardar o Aplicativo Oracle
    Windows.aguardar(
        lambda: "edge" not in Windows.titulo_janela_focada().lower() and Localizadores.texto_aplicativo_oracle.value.lower() in Windows.titulo_janela_focada().lower(),
        f"O Aplicativo Oracle não foi inicializado corretamente"
    )
    Logger.informar(f"Aplicativo Oracle aberto. Título '{ Windows.titulo_janela_focada() }'")
    # TODO implementar a procura pelo erro do "../screenshots/erro_java_tm_blocked.png"
    # informar os passos necessários para resolver o erro

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

    # # TODO - nome de recurso e departamento aleatorio para não dar conflito
    # from uuid import uuid4
    # for recurso in recursos:
    #     recurso.Geral.recurso = uuid4().__str__().split("-")[0]
    # for departamento in departamentos:
    #     departamento.Departamento.departamento = uuid4().__str__().split("-")[0]
    
    try:
        # abrir navegador no modo Internet Explorer
        with Navegador() as navegador:
            # maximizar janela do navegador
            # efetuar login no `SITE_EBS`
            janelaNavegador = Windows.janela_focada()
            janelaNavegador.maximizar()
            efetuar_login(navegador)

            # abrir o aplicativo oracle
            abrir_aplicativo_oracle_recurso(navegador)
            # abrir a organização Código 'ACN'
            # maximizar a janela
            abrir_organizacao_acn()
            # realizar o preenchimento de cada recurso
            for recurso in recursos:
                preenchidoSemErro = preencher_recurso(recurso)
                if preenchidoSemErro: Windows.atalho(["ctrl", "s"])
                # limpar para o próximo recurso
                Windows.atalho(["f6"])
            # fechar o Aplicativo Oracle sutilmente
            fechar_aplicativo_oracle()

            # focar a janela do navegador
            janelaNavegador.focar()
            sleep(5)

            # abrir o aplicativo oracle e maximiza-lo
            abrir_aplicativo_oracle_departamento(navegador)
            # abrir a organização Código 'ACN'
            # maximizar a janela
            abrir_organizacao_acn()
            # realizar o preenchimento de cada departamento
            for departamento in departamentos:
                preenchidoSemErro = preencher_departamento(departamento)
                if preenchidoSemErro: Windows.atalho(["ctrl", "s"])
                # limpar para o próximo departamento
                Windows.atalho(["f6"])
            # fechar o Aplicativo Oracle sutilmente
            fechar_aplicativo_oracle()

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
    Logger.informar("### Iniciado execução do fluxo ###")
    main()
    Logger.informar("### Finalizado execução com sucesso ###")