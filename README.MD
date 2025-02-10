<!-- filepath: /r:/60737 - Consolidar Dados Multiplas Planilhas - Tesouraria - Jeniffer/README.MD -->
# Consolidação de Dados em Múltiplas Planilhas

Este projeto consolida dados de arquivos `.xls` para um único arquivo Excel.

## Estrutura de Arquivos

- **main.py**  
  - Localiza arquivos na pasta `Files` e processa cada um através da função `ExtractData.get_dataframe`.  
  - Concatena os resultados em um DataFrame único e salva na pasta `ReturnFiles`.  
  - Remove os arquivos após o processamento.

- **Entities/extract_data.py**  
  - Carrega, via `xlwings`, a planilha desejada e busca dados de linhas específicas (Aplicações e Resgates).  
  - Constrói um DataFrame padronizado para cada arquivo (inserindo colunas como Agência, Conta, CNPJ etc.).  
  - Lida com exceções e fecha a instância do Excel.

## Uso
1. Coloque arquivos `.xls` na pasta `Files`.  
2. Execute o script `main.py`.  
3. Aguarde a geração do arquivo unificado em `ReturnFiles`.