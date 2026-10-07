# Python de bolso — documentação

Documentação estática de Python e bibliotecas relacionadas, com guias organizados em arquivos JSON separados e cópias JavaScript para execução direta no navegador.

## Abrir no computador

Dê dois cliques em `index.html`. A documentação é carregada por `content.js`, então também funciona quando aberta diretamente como arquivo (`file://`).

## Abrir por servidor local (opcional)

No PowerShell, na pasta do projeto, execute:

    py -m http.server 8000 --directory outputs

Se o comando py não existir, use:

    python -m http.server 8000 --directory outputs

Depois abra http://localhost:8000. Não há dependências externas nem etapa de instalação.

## Arquivos

- index.html: página e estrutura do site.
- styles.css: apresentação responsiva e temas claro/escuro.
- app.js: navegação, destaques, busca, cópia de código e navegação anterior/próxima.
- content.json / content.js: conteúdo da cola geral de Python, com cópia JavaScript para abrir sem servidor.
- tkinter.json / tkinter.js: referência da biblioteca padrão Tkinter.
- json.json / json.js, random.json / random.js e sys.json / sys.js: guias dos módulos da biblioteca padrão.
- customtkinter.json / customtkinter.js: referência da biblioteca externa CustomTkinter.
- requests.json / requests.js: guia da biblioteca externa Requests para HTTP e APIs.
- win32com_client.json / win32com_client.js: guia do cliente COM do pywin32 para automação no Windows.
- pandas_openpyxl.json / pandas_openpyxl.js: tutorial original sobre listas de dicionários, DataFrames, seleção/renomeação de colunas, Excel e acabamento com openpyxl.
- app.js: navegação, busca, cópia de exemplos e alternância entre bibliotecas.

Use o seletor **Bibliotecas** na barra lateral: os módulos incluídos com Python ficam no grupo **Biblioteca padrão do Python**; pacotes externos ficam em **Bibliotecas externas**. Para consultar cada arquivo também sem servidor local, mantenha o JSON e seu JavaScript correspondente sincronizados. Os arquivos JavaScript expõem `window.DOC_CONTENT`, `window.TKINTER_CONTENT`, `window.CUSTOMTKINTER_CONTENT`, `window.JSON_CONTENT`, `window.RANDOM_CONTENT`, `window.SYS_CONTENT`, `window.REQUESTS_CONTENT`, `window.WIN32COM_CLIENT_CONTENT` e `window.PANDAS_OPENPYXL_CONTENT`.

Requests e pywin32 são pacotes externos e precisam ser instalados nos projetos Python que os utilizarem. `win32com.client` requer Windows e o aplicativo COM correspondente instalado.

Os tutoriais de Tkinter, CustomTkinter, json, random, sys, Requests e win32com.client foram reescritos para ensinar o raciocínio e o uso dos argumentos com exemplos próprios; eles não dependem do Word citado anteriormente.
