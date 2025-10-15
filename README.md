# 🧾 Automatizador de Padronização de Notas Fiscais e Boletos

Este projeto é uma ferramenta com interface gráfica (GUI) desenvolvida em **Python** para automatizar o processamento de **Notas Fiscais (NFs)** e **Boletos**, com base em uma planilha de controle no formato Excel. Ideal para empresas que lidam com grande volume de documentos fiscais.

---

## 📦 Funcionalidades

- ✅ Leitura de uma planilha Excel (.xlsx) com dados de pedidos e NFs.
- 📂 Seleção de pasta contendo os PDFs de Notas e Boletos.
- 🔍 Extração automática do número da nota a partir do conteúdo do PDF.
- ✏️ Atualização da planilha com número da NF e status de processamento.
- 🗃️ Renomeia e move arquivos para uma pasta organizada por mês.
- 🧾 Log detalhado do processamento na interface.
- 🖥️ Interface gráfica intuitiva feita com `tkinter`.

---

## 🖼️ Interface do Usuário

O fluxo de uso é simples:

1. Selecione a **planilha de controle Excel**.
2. Selecione a **pasta contendo as subpastas** `Notas/` e `Boletos/`.
3. Clique em **“Iniciar Processamento”**.
4. Acompanhe o progresso na barra e no log.

---

## 📂 Estrutura Esperada

### Exemplo de pasta de entrada:

````
📁 Documentos
 ┣ 📁 Notas
 ┃ ┗ 📄 123456.pdf
 ┣ 📁 Boletos
 ┃ ┗ 📄 123456.pdf
 ┗ 📄 controle.xlsx
````

A planilha deve conter uma aba com colunas como `PEDIDO`, `NF`, e `DANFE` (ou `Nº DANFE`). A aplicação identifica automaticamente essas colunas.

## 📤 Saída

Os arquivos PDF renomeados serão salvos em uma nova pasta com o nome do mês atual (ex: 10 - Outubro), no mesmo diretório da pasta de entrada:

````
📁 10 - Outubro
 ┣ 📄 10001 - Transmissora XYZ - Danfe.pdf
 ┣ 📄 10001 - Transmissora XYZ - Boleto.pdf
````
## ⚙️ Tecnologias Utilizadas

- #### Python 3.8+
- `tkinter` — Interface gráfica
- `openpyxl` — Manipulação de planilhas Excel
- `PyMuPDF (fitz)` — Leitura de PDFs
- `threading` — Execução paralela
- `re`, `os`, `shutil`, `datetime` — Utilitários padrão do Python

## ▶️ Como Executar

1. #### Clone o repositório:
```sh
git clone https://github.com/seu-usuario/automatizador-nfs-boletos.git
cd automatizador-nfs-boletos
```
2. #### Crie um ambiente virtual (opcional, mas recomendado):
```sh
python -m venv venv
source venv/bin/activate  # ou venv\Scripts\activate no Windows
```
3. #### Instale as dependências:
```sh
pip install -r requirements.txt
```
4. #### Execute o programa:
```sh
python process_nf.py
```
## 🐞 Possíveis Problemas

- Erro ao abrir planilha: Certifique-se de que ela não esteja aberta em outro programa.
- Erro ao processar PDF: Verifique se o arquivo não está corrompido e se o texto é extraível (alguns PDFs são imagens).

## ⚠️ Observações Importantes

- Arquivos já processados não são sobrescritos.
- Se um nome de arquivo já existir, um contador será adicionado automaticamente.
- A aplicação lida apenas com arquivos `.pdf`.
- Caso o PDF seja escaneado (sem texto selecionável), a extração do número da nota poderá falhar.

## 📌 To-Do / Melhorias Futuras
- Adicionar suporte a leitura OCR para PDFs escaneados.
- Exportar log para arquivo `.txt.`
- Adicionar suporte a múltiplas abas ou layouts personalizados de planilha.
