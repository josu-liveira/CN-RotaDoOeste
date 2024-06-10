# Linear Fichas - Monitoração 2024
Este repositório possui scripts Python com uma interface gráfica amigável para análise de fichas em formato Excel. Ele extrai os dados de cada ficha e os organiza em um DataFrame do Pandas. Depois, combina os DataFrames de várias fichas em um único DataFrame e salva o resultado em um arquivo Excel.

# Interface

<div style="display: flex; flex-direction: row; align-items: center;">
  <img src="https://github.com/josu-liveira/cnro-fichas/assets/167824520/cdd1e5b9-a7a6-407a-9896-e0fde8af00b5" alt="programa" style="width: 500px;"/>
  <img src="https://github.com/josu-liveira/cnro-fichas/assets/167824520/adac9536-eb99-47b9-a126-fb48c773a71e" alt="image" style="width: 500px;"/>
</div>



## Funcionalidades

- Lê dados de várias fichas de drenagem e terraplenos em formato Excel.
- Organiza os dados em um DataFrame do Pandas.
- Combina os DataFrames de várias fichas em um único DataFrame.
- Salva o DataFrame final em um arquivo Excel.

## Bibliotecas necessárias

- Python 3.x
- pandas
- openpyxl
- customtkinter
- xlrd

## Guia de instalação

1. Baixe o zip deste repositório

2. Navegue até o diretório do projeto

3. Com base no diretório em que você extraiu os arquivos, crie um arquivo `executar.cmd`

5. Edite o arquivo `executar.cmd` adicionando `python` + `C:\Users\seuusuário\pasta\arquivo.py` e salve. Certifique-se de setar o endereço correto de seu script.

    ```
    python C:\Users\seuusuário\pasta\arquivo.py
    ```

6. (OPCIONAL) Crie um atalho do arquivo `executar.cmd` em seu Desktop.
   
7. Por fim, execute o atalho/arquivo.

## Licença

Este projeto está licenciado sob a licença [Open Source](https://opensource.org/licenses/MIT).
