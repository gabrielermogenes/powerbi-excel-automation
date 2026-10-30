# Power BI → Excel Automation

> **EN summary:** Python automation that logs into a Power BI report (manual login, MFA supported), applies saved bookmarks, exports the visual tables with Playwright, and pastes the data into an Excel `.xlsm` workbook through xlwings — keeping macros, charts and formulas intact. It validates columns before writing, creates a backup first, and never saves the workbook if something fails. Windows + Excel required.

Automação em Python que **extrai tabelas de um relatório do Power BI e atualiza uma planilha Excel (`.xlsm`)** sem nenhum trabalho manual de filtrar, exportar, abrir e colar.

## O problema

Em uma rotina diária de indicadores industriais, a atualização da planilha de acompanhamento exigia:

1. abrir o relatório no Power BI Service;
2. aplicar a sequência correta de filtros para cada visão;
3. exportar cada tabela;
4. abrir os arquivos e colar cada um na aba certa da planilha;
5. rodar as macros de organização e enviar o resultado por e-mail.

Era repetitivo, demorado e fácil de errar (filtro esquecido, aba errada, coluna deslocada).

## A solução

```
┌─────────────┐   Playwright    ┌──────────────┐   openpyxl    ┌───────────────┐   xlwings   ┌──────────────┐
│  Power BI   │ ──────────────► │  Arquivos    │ ────────────► │  Validação de │ ──────────► │  Planilha    │
│  (bookmarks)│  exporta tabelas│  .xlsx locais│  lê e confere │  colunas      │  cola dados │  .xlsm       │
└─────────────┘                 └──────────────┘               └───────────────┘  + macros   └──────────────┘
```

- **Etapa 1 (`src/powerbi_export.py`)**: abre o relatório, você faz o login (com MFA, se pedir) e o script segue sozinho: aplica cada indicador pessoal (*bookmark*), localiza a tabela correta e exporta os dados.
- **Etapa 2 (`src/excel_paste.py`)**: lê cada arquivo exportado, confere se as colunas batem com a tabela de destino e cola os dados usando o próprio Excel, preservando macros, gráficos e fórmulas. No final, pode rodar macros da planilha.

Cada etapa pode ser ligada ou desligada no `config.py`.

## Resultado

Na rotina em que foi criada, o processo completo (extração, atualização da planilha e envio por e-mail) levava **cerca de 45 minutos** manualmente e passou a **cerca de 2 minutos** (medição de uma rodada).

> O envio do e-mail é feito por um fluxo do **Power Automate**, separado e **fora deste repositório**. Este projeto cobre a extração e a atualização da planilha.

## Recursos

- **Login manual**: nenhuma senha, token ou sessão é lida, salva ou versionada.
- **Seleção do visual correto**: cada visual do Power BI fica em seu próprio `<visual-container>`; o botão "..." é achado dentro do container que tem um texto exclusivo da tabela desejada.
- **Validação antes de escrever**: o script só cola se os cabeçalhos do arquivo baterem com uma tabela da aba. Caso contrário, a aba é pulada e nada é alterado.
- **Backup automático** da planilha antes de qualquer escrita.
- **Segurança na escrita**: se houver erro no meio da colagem, a planilha **não é salva** e o original permanece intacto.
- **Abas independentes**: se uma extração falhar, as outras continuam, e a aba afetada não recebe dado antigo.
- **Aviso de arquivo velho**: informa quando o arquivo exportado não é de hoje.
- **Macros opcionais** depois da colagem, na ordem configurada.
- **Diagnóstico**: se a tabela não é encontrada, imprime o que viu e salva um print para investigação.

## Requisitos

- **Windows com Microsoft Excel instalado** (a etapa 2 usa o Excel via xlwings).
- Python 3.10 ou superior.
- Acesso ao relatório do Power BI Service, com os indicadores pessoais (*bookmarks*) já criados.
- Interface do Power BI **em português**: o script clica em textos como "Exportar dados" e "Dados resumidos".

## Instalação

```bash
git clone https://github.com/gabrielermogenes/powerbi-excel-automation.git
cd powerbi-excel-automation

python -m venv .venv
.venv\Scripts\activate

pip install -r requirements.txt
playwright install chromium   # só se não for usar um navegador já instalado
```

## Configuração

```bash
copy config.example.py config.py
```

Abra o `config.py` e preencha:

| Variável | O que informar |
|---|---|
| `REPORT_URL` | Link do relatório no Power BI Service |
| `OPERA_PATH` | Caminho de um navegador já instalado, ou vazio para usar o Chromium do Playwright |
| `PASTA_EXPORTACOES` | Pasta onde os arquivos exportados são salvos |
| `EXTRACOES` | Lista de `(indicador pessoal, arquivo de saída, texto-âncora da tabela)` |
| `PLANILHA_DESTINO` | Caminho da planilha `.xlsm` (comece apontando para uma **cópia**) |
| `ABAS_PARA_COLAR` | Qual arquivo exportado vai para qual aba |
| `MACROS_DEPOIS_DE_COLAR` | Macros a rodar depois da colagem (opcional) |
| `PRIMEIRA_COLUNA_DO_CABECALHO` | Texto com que começa a linha de cabeçalho no arquivo exportado |
| `CABECALHOS_EQUIVALENTES` | Mapa de nomes de colunas do Power BI para os nomes da planilha |

> O `config.py` está no `.gitignore` e **não deve ser publicado**: ele contém caminhos e links do seu ambiente.

## Como usar

```bash
python main.py
```

1. O navegador abre o relatório. Faça o login manualmente.
2. Quando o relatório carregar, o script continua sozinho.
3. Os arquivos são exportados para `PASTA_EXPORTACOES` e colados na planilha.
4. Leia o resumo no terminal e aperte ENTER para fechar.

Para testar só a colagem (sem logar no Power BI), use `EXTRAIR_DO_POWER_BI = False`: o script usa os arquivos que já estão na pasta.

## Como escolher o texto-âncora e os seletores

O texto-âncora precisa existir **somente dentro da tabela que você quer exportar** (por exemplo, o cabeçalho de uma coluna exclusiva). Se o script não achar a tabela, ele imprime um diagnóstico.

Para descobrir seletores de um novo relatório, use o inspetor do Playwright (`page.pause()`) e o modo de seleção de elementos. Os identificadores usados hoje são `bookmarks-btn`, `visual-more-options-btn`, `pbimenu-item.Exportar dados` e `export-btn`.

## Estrutura

```
powerbi-excel-automation/
├── main.py                  # ponto de entrada
├── config.example.py        # modelo de configuração (copie para config.py)
├── requirements.txt
├── src/
│   ├── powerbi_export.py    # etapa 1: Power BI -> arquivos
│   └── excel_paste.py       # etapa 2: arquivos -> planilha .xlsm
└── README.md
```

## Segurança e boas práticas

- Nunca coloque senhas ou tokens no código. Aqui o login é manual.
- Nunca versione planilhas, exportações ou prints: podem conter dados reais (o `.gitignore` já bloqueia esses tipos).
- Teste sempre em uma **cópia** da planilha antes de usar a de produção.
- Use somente em relatórios aos quais você tem acesso autorizado e respeitando as políticas da sua organização.

## Limitações

- Funciona somente no Windows com Excel instalado.
- Depende da interface do Power BI Service (textos em português e `data-testid`). Se a Microsoft mudar a interface, os seletores podem precisar de ajuste.
- Os intervalos de espera (`time.sleep`) são fixos; em conexões lentas pode ser preciso aumentá-los.
- As macros da planilha não fazem parte do repositório.

## Licença

MIT. Veja o arquivo [LICENSE](LICENSE).

## Autor

**Gabriel Ermogenes de Paula** — [LinkedIn](https://www.linkedin.com/in/devgabrielermogenes) · [GitHub](https://github.com/gabrielermogenes)
