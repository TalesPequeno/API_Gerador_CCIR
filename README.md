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

* Linha 136: coordenada_x = 200 --> Coloque a coordenada X do seu navegador
* Linha 137: coordenada_y = 700 --> Coloque a coordenada Y do seu navegador
* Linha 150: caminho_download = "C:/Users/Tales/Downloads" --> Coloque o caminho de onde é salvo o arquivo

### **5. Execute o programa pelo terminal:**
* Digite `python gerar.py`

### **6. Execute o programa pelo terminal:**
* Seleciona o caminho do excel com os dados dos clientes, com as seguintes especificações:
Coluna A= Código do Imóvel Rural
Coluna B= Estado
Coluna C= Município
Coluna E= CPF do Declarante
Coluna G= Nome do Declarante
Segue uma imagem de exemplo:
<img src="https://uploaddeimagens.com.br/images/004/656/533/original/Exemplo_excel.png?1699216681">

