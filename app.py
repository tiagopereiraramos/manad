import pandas as pd
from dataclasses import dataclass
from typing import List, Dict

@dataclass
class RegistroK300:
    cod_rubrica: str
    valor_rubrica: float
    cod_reg_trab: str
    dt_comp: str

@dataclass
class RegistroK150:
    cod_rubrica: str
    desc_rubrica: str

class MANADProcessor:
    """A class for processing MANAD (Manual Normativo de Arquivos Digitais) files.
    This class handles the loading, parsing and processing of MANAD files, specifically
    focusing on K150 (rubric descriptions) and K300 (rubric entries) records.
    Attributes:
        file_path (str): Path to the MANAD file to be processed
        k300_data (List[RegistroK300]): List containing all K300 records
        k150_data (Dict[str, str]): Dictionary mapping rubric codes to their descriptions
    Methods:
        load_data(): Loads data from the MANAD file and stores K150 and K300 records
        parse_line(line: str): Identifies and processes each line according to its record type
        parse_k300(line: str) -> RegistroK300: Processes K300 records (rubric entries)
        parse_k150(line: str) -> RegistroK150: Processes K150 records (rubric descriptions)
        process_data() -> Tuple[pd.DataFrame, pd.DataFrame]: Consolidates and processes the data
        gerar_relatorio_formatado(df_bruto, agrupado, descricoes_rubricas, arquivo_saida): 
            Generates a formatted Excel report with sorted rubrics and associated descriptions
    Example:
        processor = MANADProcessor("path/to/manad/file.txt")
        processor.load_data()
        df_bruto, agrupado = processor.process_data()
        processor.gerar_relatorio_formatado(df_bruto, agrupado, 
                                          processor.k150_data, "output.xlsx")
"""                                          
    def __init__(self, file_path: str):
        self.file_path = file_path
        self.k300_data: List[RegistroK300] = []
        self.k150_data: Dict[str, str] = {}

    def load_data(self):
        """
        Carrega os dados do arquivo MANAD e armazena os registros K150 (descrição das rubricas)
        e K300 (lançamentos por rubrica).
        """
        with open(self.file_path, 'r', encoding='ISO-8859-1') as file:
            for line in file:
                self.parse_line(line.strip())

    def parse_line(self, line: str):
        """
        Identifica o tipo de registro na linha e processa de acordo.
        """
        if line.startswith("K300"):
            self.k300_data.append(self.parse_k300(line))
        elif line.startswith("K150"):
            k150_record = self.parse_k150(line)
            self.k150_data[k150_record.cod_rubrica] = k150_record.desc_rubrica

    def parse_k300(self, line: str) -> RegistroK300:
        """
        Processa registros K300 (lançamentos por rubrica).
        """
        fields = line.split('|')
        return RegistroK300(
            cod_rubrica=fields[6],
            valor_rubrica=float(fields[7].replace(',', '.')),
            cod_reg_trab=fields[4],
            dt_comp=fields[5]
        )

    def parse_k150(self, line: str) -> RegistroK150:
        """
        Processa registros K150 (descrição das rubricas).
        """
        fields = line.split('|')
        return RegistroK150(
            cod_rubrica=fields[3],
            desc_rubrica=fields[4]
        )

    def process_data(self):
        """
        Consolida os dados, agrupando por rubrica e período, ordena pelo código da rubrica
        e pela data (dt_comp) de forma ascendente.
        """
        # Converter dados para DataFrame
        df_k300 = pd.DataFrame([vars(record) for record in self.k300_data])

        # Ajustar formato de datas e criar coluna para análise por mês/ano
        df_k300['dt_comp'] = pd.to_datetime(df_k300['dt_comp'], format='%m%Y', errors='coerce')
        df_k300['mes_ano'] = df_k300['dt_comp'].dt.to_period('M')

        # Ordenar os dados pela data (dt_comp) em ordem ascendente
        df_k300.sort_values(by='dt_comp', inplace=True)

        # Agrupamento por rubrica
        agrupado = df_k300.groupby(['mes_ano', 'cod_rubrica']).agg(
            soma_valor=('valor_rubrica', 'sum'),
            funcionarios=('cod_reg_trab', 'nunique')
        ).reset_index()

        # Ordenar pelo código da rubrica
        agrupado['cod_rubrica'] = agrupado['cod_rubrica'].astype(int)
        agrupado.sort_values(by=['cod_rubrica'], inplace=True)

        return df_k300, agrupado


    @staticmethod
    def gerar_relatorio_formatado(df_bruto, agrupado, descricoes_rubricas, arquivo_saida):
        """
        Gera um relatório formatado no Excel com as rubricas ordenadas e descrições associadas.
        """
        # Adicionar a descrição da rubrica ao agrupamento antes de renomear as colunas
        agrupado['Nome da Rubrica'] = agrupado['cod_rubrica'].astype(str).map(descricoes_rubricas)

        # Renomear e reorganizar as colunas para o formato esperado
        agrupado.rename(columns={
            'cod_rubrica': 'Rubrica',
            'soma_valor': 'Valor Calculado',
            'funcionarios': 'Nº Empregados/Contribuintes'
        }, inplace=True)

        # Adicionar uma coluna "Valor Informado" para comparação manual
        agrupado['Valor Informado'] = None  # Esta coluna pode ser preenchida manualmente para validar

        # Reordenar colunas
        relatorio_final = agrupado[[
            'mes_ano', 'Rubrica', 'Nome da Rubrica', 
            'Nº Empregados/Contribuintes', 'Valor Informado', 'Valor Calculado'
        ]]

        # Exportar para Excel
        relatorio_final.to_excel(arquivo_saida, index=False)
        print(f"Relatório gerado: {arquivo_saida}")

import os

if __name__ == "__main__":
    # Caminhos das pastas de entrada e saída
    pasta_entrada = r"C:\\Users\\Tiago Notebook\\Desktop\\Manad"
    pasta_saida = r"C:\\Users\\Tiago Notebook\\Documents\\Projetos\\MANAD\\retorno"

    # Verificar se a pasta de saída existe; caso contrário, criá-la
    os.makedirs(pasta_saida, exist_ok=True)

    # Listar todos os arquivos na pasta de entrada
    for arquivo in os.listdir(pasta_entrada):
        if arquivo.endswith(".TXT"):  # Processar apenas arquivos .TXT
            caminho_arquivo = os.path.join(pasta_entrada, arquivo)

            # Nome do relatório com base no nome do arquivo
            nome_base = os.path.splitext(arquivo)[0]  # Remove a extensão
            nome_relatorio = f"Rel_{nome_base}.xlsx"
            caminho_relatorio = os.path.join(pasta_saida, nome_relatorio)

            print(f"Processando arquivo: {caminho_arquivo}")
            print(f"Gerando relatório: {caminho_relatorio}")

            # Processar o arquivo MANAD
            processor = MANADProcessor(caminho_arquivo)
            processor.load_data()

            # Processar os dados
            df_bruto, agrupado = processor.process_data()

            # Gerar relatório formatado
            MANADProcessor.gerar_relatorio_formatado(
                df_bruto,
                agrupado,
                processor.k150_data,
                caminho_relatorio
            )
