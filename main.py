"""
Ponto de entrada: extrai os dados do Power BI (etapa 1) e cola na planilha
.xlsm (etapa 2). Cada etapa pode ser ligada/desligada em config.py.

Uso:
    python main.py
"""

import os

try:
    import config
except ModuleNotFoundError:
    raise SystemExit(
        "Arquivo config.py não encontrado.\n"
        "Copie config.example.py para config.py e preencha os valores."
    )

from src.excel_paste import colar_na_planilha
from src.powerbi_export import extrair_do_power_bi

PLACEHOLDER = "COLE_AQUI"


def validar_configuracao():
    """Avisa cedo se algum valor de exemplo ainda não foi preenchido."""
    pendentes = []
    if PLACEHOLDER in config.PASTA_EXPORTACOES:
        pendentes.append("PASTA_EXPORTACOES")
    if config.EXTRAIR_DO_POWER_BI and PLACEHOLDER in config.REPORT_URL:
        pendentes.append("REPORT_URL")
    if config.COLAR_NA_PLANILHA and PLACEHOLDER in config.PLANILHA_DESTINO:
        pendentes.append("PLANILHA_DESTINO")
    if pendentes:
        raise SystemExit("Preencha no config.py: " + ", ".join(pendentes))


def main():
    validar_configuracao()
    os.makedirs(config.PASTA_EXPORTACOES, exist_ok=True)

    # None = a etapa 1 não rodou agora (a etapa 2 usa os arquivos que já estão na pasta)
    exportados = None
    if config.EXTRAIR_DO_POWER_BI:
        exportados = extrair_do_power_bi()

    if config.COLAR_NA_PLANILHA:
        colar_na_planilha(exportados)

    # Sempre espera aqui antes de fechar, mesmo com erro, para dar tempo de
    # ler a mensagem no terminal.
    input("\nConcluído (com ou sem erro acima). Aperte ENTER para fechar esta janela...")


if __name__ == "__main__":
    main()
