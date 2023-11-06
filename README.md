# API_Gerador_CCIR

## Introdução:
Este é um projeto em Python com integração de Excel para criar automaticamente vários CCIR (Certificados de Cadastro de Imóvel Rural).

## Tecnologias utilizadas:
* [Python](https://www.python.org/): Linguagem de programação
* tkinter: Uma biblioteca padrão de Python para criar interfaces gráficas de usuário.
* customtkinter: Uma biblioteca personalizada baseada em Tkinter para criar interfaces gráficas de usuário (GUI).
* selenium: Uma biblioteca usada para automatizar ações em navegadores da web, como Chrome.
* openpyxl: Uma biblioteca usada para ler e escrever em planilhas do Excel.
* os: Uma biblioteca padrão para interagir com o sistema operacional, que é usada para operações de sistema de arquivos.
* time: Uma biblioteca padrão para trabalhar com medidas de tempo.

## Como executar:
### **1. Instale `Python` na sua máquina, por meio [deste link](https://www.python.org/)**

### **2. Faça um clone [desse repositório](https://github.com/TalesPequeno/API_Gerador_CCIR.git) na sua máquina:**

* Crie uma pasta no seu computador para esse programa, recomendo colocar o nome **API_Gerador_CCIR**
* Abra o `git bash` ou `terminal` dentro dessa pasta
* Copie a [URL](https://github.com/TalesPequeno/API_Gerador_CCIR.git) do repositório
* Digite `git clone <URL copiada>` e pressione `enter`

### **3. Instale as bibliotecas necessárias pelo terminal, dentro dessa pasta criada:**

* customtkinter: `pip install customtkinter`
* selenium: `pip install selenium`
* openpyxl: `pip install openpyxl`

### **4. Preencha os dados no main.py conforme seu computador:**

```py 
    # Altera a coordenada do Sou Humano do seu navegador (Google Chrome):
    coordenada_x = 200
    coordenada_y = 700
```
```py
    # Altera o caminho de onde é salvo o download
    caminho_download = "C:/Users/Tales/Downloads"
```
### **5. Execute o programa pelo terminal:**
* Digite `python gerar.py`

### **6. Execute o programa pelo terminal:**
* 1- Seleciona o caminho do excel com os dados dos clientes, com as seguintes especificações:
* Coluna A= Código do Imóvel Rural
* Coluna B= Estado
* Coluna C= Município
* Coluna E= CPF do Declarante
* Coluna G= Nome do Declarante
* Segue uma imagem de exemplo:
<img src="https://uploaddeimagens.com.br/images/004/656/533/original/Exemplo_excel.png?1699216681">

* As Colunas D e F no momento não estão sendo utilizados, mas no futuro irei atualizar a API para verificar as duas colunas.

* 2- Coloca no campo "Linha de Início" qual linha começa os dados do clientes no excel.
* 3- Coloca no campo "Linha de Término" qual linha termina os dados do clientes no excel.

* 4- Clicar em "Gerar" e aguardar o processo.
