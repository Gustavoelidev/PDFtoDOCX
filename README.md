# Conversor de PDF para DOCX usando Python e Adobe Acrobat

Este é um script Python que automatiza a conversão de arquivos PDF para DOCX utilizando o Adobe Acrobat. Ele percorre um diretório de entrada, localiza arquivos PDF correspondentes a um padrão especificado, e os converte para DOCX utilizando as capacidades do Acrobat.

## Pré-requisitos

Antes de executar este script, certifique-se de ter os seguintes requisitos instalados:

`pip install -r requirements.txt` 

Certifique-se de estar usando um ambiente virtual Python para gerenciar suas dependências.

## Configuração

1.  **Instalação do Python**: Se você não tiver o Python instalado, faça o download e instale a partir do [site oficial do Python](https://www.python.org/).
    
2.  **Ambiente Virtual**: É recomendável usar um ambiente virtual para isolar suas dependências Python. Você pode criar um ambiente virtual usando o `venv`:
    
    `python -m venv .venv` 
    
    Em seguida, ative o ambiente virtual:
    
    -   No Windows:

        `.venv\Scripts\activate` 
        
    -   No macOS e Linux:
                
        `source .venv/bin/activate` 
        
3.  **Instalação das Dependências**: Após configurar seu ambiente virtual, instale as dependências listadas no arquivo `requirements.txt`:
    
    
    `pip install -r requirements.txt` 
    

## Uso

Para usar o script, execute-o fornecendo os seguintes argumentos de linha de comando:

`python app.py <caminho_para_pasta_origem> <padrão_para_arquivos_pdf> <caminho_para_pasta_destino> <extensão_para_arquivos_docx>` 

-   `<caminho_para_pasta_origem>`: Caminho absoluto para o diretório que contém os arquivos PDF que você deseja converter.
-   `<padrão_para_arquivos_pdf>`: Padrão de nome de arquivo PDF que você deseja corresponder (por exemplo, `*.pdf`).
-   `<caminho_para_pasta_destino>`: Caminho absoluto para o diretório onde os arquivos DOCX convertidos serão salvos.
-   `<extensão_para_arquivos_docx>`: Extensão para os arquivos DOCX de saída (por exemplo, `.docx`).

### Exemplo

`python app.py 'C:\caminho\para\pasta\com\pdfs' '*.pdf' 'C:\caminho\para\pasta\de\saida' '.docx'` 

## Funcionalidades

-   Utiliza o Adobe Acrobat via `win32com` para abrir, converter e salvar arquivos PDF como DOCX.
-   Usa `scandir` para percorrer eficientemente o diretório de entrada em busca de arquivos PDF correspondentes ao padrão especificado.

## Contribuição

Contribuições são bem-vindas! Sinta-se à vontade para enviar pull requests ou abrir issues se tiver sugestões de melhorias ou correções.
