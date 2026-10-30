"""
Configuração de exemplo.

Como usar:
    1. Copie este arquivo para config.py  (cp config.example.py config.py)
    2. Troque todos os valores "COLE_AQUI" pelos seus
    3. NUNCA publique o config.py: ele já está no .gitignore

Dica: lembre do "r" antes das aspas nos caminhos do Windows: r"C:\\pasta\\arquivo".
"""

# ==================== POWER BI ====================

# Link do relatório no Power BI Service (copie da barra de endereços do navegador).
# Contém o ID do workspace (groups/...) e o ID do relatório (reports/...).
REPORT_URL = "https://app.powerbi.com/groups/COLE_AQUI_ID_DO_WORKSPACE/reports/COLE_AQUI_ID_DO_RELATORIO"

# Navegador usado pelo Playwright.
# Deixe vazio ("") para usar o Chromium do Playwright (rode: playwright install chromium).
# Ou aponte para um navegador já instalado: r"C:\caminho\do\navegador.exe"
OPERA_PATH = r""

# Pasta onde os arquivos exportados são salvos (criada se não existir).
PASTA_EXPORTACOES = r"C:\COLE_AQUI\pasta_de_exportacoes"

# Lista de extrações: (nome do indicador pessoal, arquivo de saída, texto-âncora).
#   - indicador pessoal: o "bookmark" salvo no relatório, que já aplica os filtros certos
#   - texto-âncora: um texto que só existe DENTRO da tabela a exportar (ex.: um cabeçalho
#     de coluna exclusivo dela). É usado para achar o botão "..." da tabela certa.
EXTRACOES = [
    ("Indicador A - Extração", "tabela_a.xlsx", "Texto exclusivo da tabela A"),
    ("Indicador B - Extração", "tabela_b.xlsx", "Texto exclusivo da tabela B"),
]

# ==================== ETAPAS (liga/desliga) ====================

# Etapa 1: abrir o Power BI e exportar os arquivos
EXTRAIR_DO_POWER_BI = True

# Etapa 2: colar os arquivos exportados na planilha .xlsm
# (para testar só a colagem, deixe a etapa 1 em False: usa os arquivos que já estão na pasta)
COLAR_NA_PLANILHA = True

# ==================== PLANILHA DE DESTINO ====================

# Caminho da planilha .xlsm. COMECE APONTANDO PARA UMA CÓPIA, não para a de produção.
PLANILHA_DESTINO = r"C:\COLE_AQUI\sua_planilha.xlsm"

# Faz uma cópia de segurança da planilha antes de colar
FAZER_BACKUP = True

# Macros da planilha para rodar DEPOIS de colar tudo, na ordem da lista.
# Deixe vazia ([]) para não rodar nenhuma.
# Exemplos: "NomeDaMacro" (macro de módulo) ou "Planilha1.NomeDaMacro" (macro de aba)
MACROS_DEPOIS_DE_COLAR = []

# Qual arquivo exportado vai para qual aba da planilha.
# O nome da aba precisa ser IGUAL ao da planilha, inclusive acento e maiúscula.
# Não é preciso dizer quantas colunas colar: o script usa as colunas que vieram no
# arquivo e só mexe nelas (colunas de fórmula ficam intactas).
ABAS_PARA_COLAR = [
    ("tabela_a.xlsx", "Aba A"),
    ("tabela_b.xlsx", "Aba B"),
]

# ==================== LEITURA DO ARQUIVO EXPORTADO ====================

# Texto com que começa a linha de cabeçalho no arquivo exportado (ignora acento e
# maiúscula). Acima dela o Power BI coloca uma linha de "Filtros aplicados".
PRIMEIRA_COLUNA_DO_CABECALHO = "Equipamento"

# Alguns cabeçalhos vêm do Power BI com outro nome. Aqui você diz "esse nome do arquivo
# equivale a esse nome da tabela da planilha" (acento e maiúscula não importam).
# Exemplo:  "'Tabela'[Coluna (Label)]": "Coluna"
CABECALHOS_EQUIVALENTES = {}
