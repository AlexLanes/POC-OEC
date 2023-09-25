# std
from dataclasses import dataclass
# interno
from src.util import *
from src.logger import *
# externo
from pandas import read_excel

class Tipos:
    @dataclass
    class Geral:
        recurso: str
        descricao: str
        tipo: str
        tipoEncargo: str
        udm: str
    @dataclass
    class ProcessamentoExterno:
        processamentoExterno: bool
        item: str
    @dataclass
    class Faturamento:
        custeado: bool
        contaAbsorcao: str
        contaVariacao: str
        taxaPadrao: bool
    @dataclass
    class Taxa:
        tipoCusto: str
        descricao3: str
        custoUsuarioRecurso: str
    @dataclass
    class Departamento:
        departamento: str
        descricao: str
        categoria_de_custo: str
        local: str
    @dataclass
    class Recurso:
        recurso: str
        descricao2: str
        disponibilizar_horas: bool
        udm: str
        unidades: str

class Recurso:
    def __init__(self, r: tuple) -> None:
        """Criar um objeto para armazenar os atributos de um `Recurso`"""
        self.Geral = Tipos.Geral(r.recurso, r.descricao, r.tipo, r.tipo_de_encargo, r.udm)
        self.ProcessamentoExterno = Tipos.ProcessamentoExterno(r.processamento_externo, r.item)
        self.Faturamento = Tipos.Faturamento(r.custeado, r.conta_absorcao, r.conta_variacao, r.taxa_padrao)
        self.Taxas = [ Tipos.Taxa(r.tipo_de_custo, r.descricao3, r.custo_unitario_do_recurso) ] if  r.taxa_padrao else []

    def __eq__(self, other) -> bool:
        """Comparador de `Recurso`"""
        return self.Geral.__dict__ == other.Geral.__dict__\
            and self.ProcessamentoExterno.__dict__ == other.ProcessamentoExterno.__dict__\
            and self.Faturamento.__dict__ == other.Faturamento.__dict__

class Departamento:
    def __init__(self, d: tuple) -> None:
        """Criar um objeto para armazenar os atributos de um `Departamento`"""
        self.Departamento = Tipos.Departamento(d.departamento, d.descricao, d.categoria_de_custo, d.local)
        self.Recursos = [ Tipos.Recurso(d.recurso, d.descricao2, d.disponibilizar_horas, d.udm, d.unidades) ]

    def __eq__(self, other) -> bool:
        """Comparador de `Departamento`"""
        return self.Departamento.__dict__ == other.Departamento.__dict__
    
def parse_recursos(caminhoAbsolutoExcel: str) -> list[Recurso]:
    """Realiza o parse da planilha recursos do excel, com base no caminho absoluto, para um formato amigável"""
    NOME_PLANILHA = "RECURSOS"
    COLUNAS_TIPOS_ESPERADOS = {
        'recurso': 'object', 
        'descricao': 'object', 
        'tipo': 'object', 
        'tipo_de_encargo': 'object', 
        'udm': 'object', 
        'processamento_externo': 'bool', 
        'item': 'object', 
        'custeado': 'bool', 
        'conta_absorcao': 'object', 
        'conta_variacao': 'object', 
        'taxa_padrao': 'bool', 
        'tipo_de_custo': 'object', 
        'descricao3': 'object', 
        'custo_unitario_do_recurso': 'object'
    }

    Logger.informar(f"Iniciado o parse da planilha '{NOME_PLANILHA}'")
    try:
        # ler e criar dataframe 
        # tratar nome das colunas
        df = read_excel(caminhoAbsolutoExcel, sheet_name=NOME_PLANILHA, false_values=["N"], true_values=["Y"], keep_default_na=False, header=1)
        df.columns = df.columns.to_series().apply(normalizar)

        # validação das colunas do dataframe
        colunasDf = list( mapear_dtypes(df).keys() )
        colunasEsperadas = list( COLUNAS_TIPOS_ESPERADOS.keys() )
        assert sorted(colunasEsperadas) == sorted(colunasDf), f"Nome(s) de coluna inesperado \n\tesperado {colunasEsperadas} \n\trecebido {colunasDf}"
        
        # order by, de recurso até taxa_padrao, para as linhas iguais ficarem sequenciais
        # replace "NULL" por ""
        df.sort_values( by=colunasEsperadas[0:11], inplace=True )
        df.replace( to_replace={ "NULL": "" }, inplace=True )
        
        # conversão para str dos campos possivelmente numéricos
        df["item"] = df["item"].astype(str)
        df["custo_unitario_do_recurso"] = df["custo_unitario_do_recurso"].astype(str)
        
        # validação dos tipos do dataframe
        tiposDf = list( mapear_dtypes(df).values() )
        tiposEsperados = list( COLUNAS_TIPOS_ESPERADOS.values() )
        assert sorted(tiposEsperados) == sorted(tiposDf), f"Tipo(s) de coluna inesperado \n\tesperado {tiposEsperados} \n\trecebido {tiposDf}"

        # criar a lista de recursos
        # agregar os recursos iguais
        recursos: list[Recurso] = []
        for recurso in df.itertuples(index=False):
            recurso = Recurso(recurso)
            if len(recursos) == 0: del df # salvar memoria
            if len(recursos) > 0 and recurso == recursos[-1]:
                recursos[-1].Taxas.extend(recurso.Taxas)
            else: recursos.append(recurso)
        
        Logger.informar(f"Finalizado o parse da planilha '{NOME_PLANILHA}'")
        return recursos

    except FileNotFoundError:
        Logger.erro(f"Excel não encontrado em '{caminhoAbsolutoExcel}'")
        exit(1)
    except AssertionError as erro:
        Logger.erro(f"Falha de validação da planilha '{NOME_PLANILHA}': {erro}")
        exit(1)
    except Exception as erro:
        Logger.erro(f"Erro inesperado na leitura da planilha '{NOME_PLANILHA}': {erro}")
        exit(1)
    
def parse_departamentos(caminhoAbsolutoExcel: str) -> list[Departamento]:
    """Realiza o parse da planilha departamentos do excel, com base no caminho absoluto, para um formato amigável"""
    NOME_PLANILHA = "DEPARTAMENTOS"
    COLUNAS_TIPOS_ESPERADOS = {
        'departamento': 'object', 
        'descricao': 'object', 
        'categoria_de_custo': 'object', 
        'local': 'object', 
        'recurso': 'object', 
        'descricao2': 'object', 
        'disponibilizar_horas': 'bool', 
        'udm': 'object', 
        'unidades': 'object'
    }
    
    Logger.informar(f"Iniciado o parse da planilha '{NOME_PLANILHA}'")
    try:
        # ler e criar dataframe 
        # tratar nome das colunas
        df = read_excel(caminhoAbsolutoExcel, sheet_name=NOME_PLANILHA, false_values=["N"], true_values=["Y"], keep_default_na=False, header=1)
        df.columns = df.columns.to_series().apply(normalizar)

        # validação das colunas do dataframe
        colunasDf = list( mapear_dtypes(df).keys() )
        colunasEsperadas = list( COLUNAS_TIPOS_ESPERADOS.keys() )
        assert sorted(colunasEsperadas) == sorted(colunasDf), f"Nome(s) de coluna inesperado \n\tesperado {colunasEsperadas} \n\trecebido {colunasDf}"

        # order by, de departamento até local, para as linhas iguais ficarem sequenciais
        df.sort_values( by=colunasEsperadas[0:4], inplace=True )

        # conversão para str dos campos possivelmente numéricos
        df["unidades"] = df["unidades"].astype(str)

        # validação dos tipos do dataframe
        tiposDf = list( mapear_dtypes(df).values() )
        tiposEsperados = list( COLUNAS_TIPOS_ESPERADOS.values() )
        assert sorted(tiposEsperados) == sorted(tiposDf), f"Tipo(s) de coluna inesperado \n\tesperado {tiposEsperados} \n\trecebido {tiposDf}"

        # criar a lista de departamentos
        # agregar os departamentos iguais
        departamentos: list[Departamento] = []
        for departamento in df.itertuples(index=False):
            departamento = Departamento(departamento)
            if len(departamentos) == 0: del df # salvar memoria
            if len(departamentos) > 0 and departamento == departamentos[-1]:
                departamentos[-1].Recursos.extend(departamento.Recursos)
            else: departamentos.append(departamento)
        
        Logger.informar(f"Finalizado o parse da planilha '{NOME_PLANILHA}'")
        return departamentos

    except FileNotFoundError:
        Logger.erro(f"Excel não encontrado em '{caminhoAbsolutoExcel}'")
        exit(1)
    except AssertionError as erro:
        Logger.erro(f"Falha de validação da planilha '{NOME_PLANILHA}': {erro}")
        exit(1)
    except Exception as erro:
        Logger.erro(f"Erro inesperado na leitura da planilha '{NOME_PLANILHA}': {erro}")
        exit(1)
    
__all__ = [
    "Recurso",
    "Departamento",
    "parse_recursos",
    "parse_departamentos"
]