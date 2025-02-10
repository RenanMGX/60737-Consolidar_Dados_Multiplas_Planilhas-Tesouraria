from Entities.dependencies.arguments import Arguments
from Entities.dependencies.functions import P
from datetime import datetime
import os

class Execute:
    """
    Classe responsável por executar o processo de extração e consolidação dos dados.
    Procura arquivos com extensão .xls na pasta 'Files', processa-os e consolida os dados em um único 
    arquivo Excel, que é salvo na pasta 'ReturnFiles'. Além disso, remove os arquivos processados.
    """
    files_path: str = os.path.join(os.getcwd(), 'Files')
    if not os.path.exists(files_path):
        os.makedirs(files_path)
        
    return_file_path = os.path.join(os.getcwd(), 'ReturnFiles')
    if not os.path.exists(return_file_path):
        os.makedirs(return_file_path)
        
    @staticmethod
    def start():
        """
        Inicia o processo de consolidação dos arquivos.
        Percorre os arquivos .xls na pasta 'Files', utiliza ExtractData para extrair os dados e consolida
        os DataFrames resultantes. Ao final, salva o DataFrame unificado em um arquivo Excel na pasta 'ReturnFiles'.
        """
        from Entities.extract_data import ExtractData, pd
        
        if not os.listdir(Execute.files_path):
            print(P("Nenhum arquivo encontrado", color='red'))
            return
        
        for _file in os.listdir(Execute.return_file_path):
            os.unlink(os.path.join(Execute.return_file_path, _file))
        
        df = pd.DataFrame()
        
        for file in os.listdir(Execute.files_path):
            file_path = os.path.join(Execute.files_path, file)
            
            if os.path.isfile(file_path):
                if file_path.lower().endswith('.xls'):
                    print(P(f"{file} Iniciado", color='blue'))
                    df_temp = ExtractData.get_dataframe(file_path=file_path, periodo=datetime.now())
                    
                    os.unlink(file_path)
                    
                    if df_temp.empty:
                        print(P(f"{file} Vazio", color='yellow'))
                        continue
                    
                    df = pd.concat([df, df_temp], ignore_index=True)
                    print(P(f"{file} Finalizado", color='green'))
                    del df_temp
        
        target_path = os.path.join(Execute.return_file_path, datetime.now().strftime('%Y%m%d%H%M%S_output.xlsx'))          
        df.to_excel(target_path, index=False)
        
        for _file in os.listdir(Execute.files_path):
            os.unlink(os.path.join(Execute.files_path, _file))
        
if __name__ == "__main__":
    Arguments({
        'start': Execute.start
    })
