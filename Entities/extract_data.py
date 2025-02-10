import os
import pandas as pd
import xlwings as xw
import re
import multiprocessing as mp
from xlwings.main import Sheet
from xlwings.main import Range
from xlwings.main import Book
from typing import Literal
from datetime import datetime
from dependencies.functions import Functions
from time import sleep
import traceback
from logInformativo import LogInformativo


valid_sheet = 'Sheet0'

def __find_ranged_lines(ws:Sheet, *, tipo:Literal['Aplicações', 'Resgates / Vencimentos'], firs_column_letter:str="A",last_column_letter:str) -> Range:
    """
    Localiza e retorna o intervalo de linhas que contém registros do tipo especificado.
    Parâmetros:
      - ws: Planilha (`Sheet`) onde serão pesquisadas as linhas.
      - tipo: Uma string que define se o trecho é de "Aplicações" ou "Resgates / Vencimentos".
      - firs_column_letter: Letra da primeira coluna a analisar (padrão "A").
      - last_column_letter: Letra da última coluna a analisar.
    Retorno:
      - `Range` correspondente às linhas localizadas.
    """
    result = {}
    achou_primeiro = False
    for num in range(1,ws.used_range.last_cell.row):
        cell = ws.range(f'{firs_column_letter}{num}')

        if not achou_primeiro:
            if cell.value == tipo:
                result['start'] =  num + 1
                achou_primeiro = True
        else:
            if cell.value == 'Total':
                result['end'] =  num - 1
                break
    return ws.range(f'{firs_column_letter}{result["start"]}:{last_column_letter}{result["end"]}')

def __find_line(ws:Sheet, *, value:Literal['Dt. Aplicação', 'Empresa/CNPJ', 'Agência/conta'], firs_column_letter:str="A", last_column_letter:str) -> Range:
    """
    Encontra uma linha que contenha o texto especificado.
    Parâmetros:
      - ws: Planilha (`Sheet`) onde será feita a busca.
      - value: Texto exato a ser procurado.
      - firs_column_letter e last_column_letter: Definem o intervalo de colunas na planilha.
    Retorno:
      - Retorna o `Range` da linha encontrada ou `None` se não existir.
    """
    for num in range(1,ws.used_range.last_cell.row):
        cell = ws.range(f'{firs_column_letter}{num}')
        #print(cell.value)
        if cell.value:
            if value in cell.value:
                return ws.range(f'{firs_column_letter}{num}:{last_column_letter}{num}')
    return None

def verify_file(file_path:str) -> bool:
    """
    Verifica se o arquivo existe e se é do tipo .xls.
    Parâmetros:
      - file_path: Caminho completo do arquivo.
    Retorno:
      - Retorna o próprio caminho se for válido, caso contrário lança exceções.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Arquivo não encontrado no caminho")
    if not file_path.endswith('.xls'):
        raise ValueError(f"O arquivo não é um arquivo xls válido")
    return file_path

def __get_agencia_conta(ws: Sheet) -> dict:
    """
    Extrai os dados de agência e conta de uma linha específica da planilha.
    Parâmetros:
      - ws: Planilha (`Sheet`) de onde serão extraídos os dados.
    Retorno:
      - Dicionário com as chaves "agencia" e "conta" ou mensagens de erro se não encontradas.
    """
    result = {}
    text = __find_line(ws, value='Agência/conta', last_column_letter='A').value
    
    if (agencia:=re.search(r'(\d{4})', text)):
        result['agencia'] = agencia.group()
    else:
        result['agencia'] = "Agencia não encontrada"
        
    if (conta:=re.search(r'(\d+-\d)', text)):
        result['conta'] = conta.group()
    else:
        result['conta'] = "Conta não encontrada"

    return result

def __get_empresa_cnpj(ws:Sheet) -> dict:
    """
    Obtém informações sobre a empresa e seu CNPJ.
    Parâmetros:
      - ws: Planilha (`Sheet`) de onde serão extraídos os dados.
    Retorno:
      - Dicionário com as chaves "empresa" e "cnpj" ou mensagens de erro se não encontradas.
    """
    result = {}
    text = __find_line(ws, value='Empresa/CNPJ', last_column_letter='A').value
    
    if (empresa:=re.search(r'(?<=[:])[\w\d\D ]+(?=[|])', text)):
        result['empresa'] = empresa.group().strip()
    else:
        result['empresa'] = "Empresa não encontrada"
        
    if (cnpj:=re.search(r'(?<=[|])[\w\d\D ]+', text)):
        result['cnpj'] = cnpj.group().strip()   
    else:
        result['cnpj'] = "CNPJ não encontrado"      
        
    return result

def corrigir_linhas_dados(dados):
    """
    Ajusta a estrutura dos dados para garantir que seja retornada uma lista de listas.
    Parâmetros:
      - dados: Lista que pode conter uma única linha ou múltiplas linhas.
    Retorno:
      - Lista de listas contendo os dados, garantindo consistência na estrutura.
    """
    if isinstance(dados[0], list):
        return dados
    return [dados]

def get_dados(ws, *, tipo:Literal['Aplicações', 'Resgates'], periodo:datetime) -> pd.DataFrame:
    """
    Retorna um DataFrame contendo dados de Aplicações ou Resgates. 
    Parâmetros:
      - ws: Planilha (`Sheet`) a ser analisada.
      - tipo: Define qual tipo de registro será coletado (Aplicações ou Resgates).
      - periodo: Data usada para identificação no DataFrame.
    Retorno:
      - DataFrame com colunas padronizadas incluindo informações de conta e empresa.
    """
    data = None
    if tipo == 'Aplicações':

        
        data = [__find_line(ws, value='Dt. Aplicação', last_column_letter="K").value] + corrigir_linhas_dados(__find_ranged_lines(ws, tipo='Aplicações', last_column_letter="K").value)
    elif tipo == 'Resgates':
        data = [__find_line(ws, value='Dt. Aplicação', last_column_letter="K").value] + corrigir_linhas_dados(__find_ranged_lines(ws, tipo='Resgates / Vencimentos', last_column_letter="K").value)
            
    agencia_conta:dict = __get_agencia_conta(ws)
    empresa_cnpj:dict = __get_empresa_cnpj(ws)

    #import pdb; pdb.set_trace()
    if data:
        df = pd.DataFrame(data)
    else:
        return pd.DataFrame()
    
    df.columns = df.iloc[0]
    df = df[1:]
    df['Tipo'] = tipo
    df['Período'] = periodo.strftime("%d/%m/%Y")
    df['Agência'] = agencia_conta['agencia']
    df['Conta'] = agencia_conta['conta']
    df['CPF/CNPJ'] = empresa_cnpj['cnpj']
    df['Nome'] = empresa_cnpj['empresa']
    df['Certificado'] = ""
    df['Vlr da Renda'] = ""
    df['Valor de IOF'] = ""
    df['Valor de IRRF'] = ""
    
    df.rename(columns={
        'Dt. Aplicação': 'Data de Emissão',
        'Dt. Vencto': 'Data de Vencto',
        'Taxa (%)': 'Taxa/ PCT',
        'Vlr Princ. (R$)': 'Valor Principal',
        'Renda Total(R$)': 'Valor da Renda',
        'Vlr. IOF (R$)': 'Valor de IOF(*)',
        'Vlr. IRRF (R$)': 'Valor de IRRF(*)',
        'Vlr. Bruto (R$)': 'Valor de Resgate',
        'Dt. Resgate / Carência': 'Data de Pagto',
        'Vlr Líquido(R$)': 'Valor do Crédito',
        'Renda Bruta Per': 'Renda no Mês',
    }, inplace=True)
    return df

class ExtractData:
    @staticmethod
    def get_dataframe(*, file_path:str, periodo:datetime) -> pd.DataFrame:
        """
        Função principal para carregar e consolidar dados de Aplicações e Resgates de uma planilha.
        Parâmetros:
          - file_path: Caminho do arquivo xls a ser processado.
          - periodo: Data para rotulação em cada linha.
        Retorno:
          - DataFrame unificado, contendo todas as colunas definidas para análise posterior.
        """
        try:
            app = xw.App(visible=False)
            app.display_alerts = False
            app.screen_updating = False
            
            wb:Book = xw.Book(file_path, update_links=False, read_only=True)

            if not valid_sheet in wb.sheet_names:
                raise ValueError(f"Sheet não encontrada no arquivo")
            ws:Sheet = wb.sheets[valid_sheet]
            
            df = pd.DataFrame()
            df_aplic = get_dados(ws, tipo='Aplicações', periodo=periodo)
            df_resg = get_dados(ws, tipo='Resgates', periodo=periodo)
            
            if not 'Aplicações' in df_aplic.iloc[0,0]:
                df = pd.concat([df, df_aplic])
            
            if not 'Resgates / Vencimentos' in df_resg.iloc[0,0]:
                df = pd.concat([df, df_resg])
            
            #import pdb; pdb.set_trace()
            #df = pd.concat([df, get_dados(ws, tipo='Resgates', periodo=periodo)])
            
            if df.empty:
                return pd.DataFrame()
                
            df = df[[
                'Período',
                'Agência',
                'Conta',
                'CPF/CNPJ',
                'Nome',
                'Tipo',
                'Certificado',
                'Data de Emissão',
                'Data de Vencto',
                'Taxa/ PCT',
                'Valor Principal',
                'Valor da Renda',
                'Valor de IOF(*)',
                'Valor de IRRF(*)',
                'Valor de Resgate',
                'Data de Pagto',
                'Vlr da Renda',
                'Valor de IOF',
                'Valor de IRRF',
                'Valor do Crédito',
                'Renda no Mês'
                ]]
            
        
        finally:
            try:
                wb.close()
            except:
                pass
            try:
                app.kill()
            except:
                pass
            
            sleep(1)
            Functions.fechar_excel(file_path)

        return df
    
    @staticmethod
    def mp_get_dataframe(queue:mp.Queue, file_path:str, periodo:datetime):
        """
        Processa o arquivo utilizando multiprocessing e insere o DataFrame resultante na fila.
        Em caso de exceção, tenta até 5 vezes e registra os logs de erro.
        Parâmetros:
          - queue: Objeto Queue do módulo multiprocessing para armazenar o DataFrame.
          - file_path: Caminho do arquivo xls a ser processado.
          - periodo: Data utilizada para rotulação nas linhas do DataFrame.
        Retorno:
          - Não retorna valor diretamente; o DataFrame é inserido na queue.
        """
        for _ in range(5):
            try:
                return queue.put(ExtractData.get_dataframe(file_path=file_path, periodo=periodo))
            except Exception as e:
                print(f"[{_+1}/5]Erro no arquivo {os.path.basename(file_path)}: {e}")
                with open(datetime.now().strftime(f"logs/%Y%m%d%H%M%S{os.path.basename(file_path)}") + '.txt', 'w') as f:
                    f.write(traceback.format_exc())
                if _ == 4:
                    print(f"Final {os.path.basename(file_path)}")
                    return queue.put(pd.DataFrame())

   
if __name__ == "__main__":
    df1 = ExtractData.get_dataframe(file_path=r'C:\Users\renan.oliveira\Downloads\x\1101050008 - SPE AXIS - PORTO FINO - 12.2024 - CDB DI OK.XLS', periodo=datetime.now())    
    df2 = ExtractData.get_dataframe(file_path=r'C:\Users\renan.oliveira\Downloads\x\1101050008 - SPE AXIS - PORTO FINO - 12.2024 - CDB OK.XLS', periodo=datetime.now())    
    
    import pdb; pdb.set_trace()
    df = pd.concat([df1, df2], ignore_index=True)
    
    df.to_excel('output.xlsx', index=False)
