"""
Etapa 1 - Extração dos dados do Power BI com Playwright.

Para cada indicador pessoal (bookmark) configurado, aplica o indicador no
relatório, localiza a tabela desejada e exporta os dados para uma pasta local.
O login no Power BI é sempre manual (inclui MFA): nenhuma credencial é
armazenada ou lida pelo script.
"""

import os
import time

from playwright.sync_api import sync_playwright

import config


def aplicar_indicador(page, nome_indicador):
    # Fecha qualquer overlay/menu que tenha ficado aberto da etapa anterior
    if page.locator(".cdk-overlay-backdrop").count() > 0:
        page.keyboard.press("Escape")
        time.sleep(1)

    page.get_by_test_id("bookmarks-btn").click()
    time.sleep(1)
    page.get_by_text(nome_indicador, exact=False).first.click()
    time.sleep(1)
    page.keyboard.press("Escape")
    time.sleep(3)  # espera os filtros da página aplicarem de vez


def _primeiro_visivel(locator):
    """Entre todas as cópias do texto na página, devolve a primeira que está
    visível (ou None). O texto pode existir mais de uma vez no HTML, com
    cópias escondidas, então o .first sozinho pode pegar a errada."""
    for i in range(locator.count()):
        item = locator.nth(i)
        try:
            if item.is_visible():
                return item
        except Exception:
            continue
    return None


def esperar_ancora(page, texto_ancora):
    """
    Espera o texto-âncora ficar visível. Algumas tabelas ficam mais abaixo
    na página e só aparecem depois de rolar. Como não há um único elemento
    responsável pela rolagem no Power BI, o script localiza todas as áreas
    roláveis da página e rola cada uma aos poucos, conferindo a cada passo
    se o texto já apareceu. Se não encontrar, imprime um diagnóstico e salva
    um print para facilitar a investigação.
    """
    locator = page.get_by_text(texto_ancora, exact=False)
    locator.first.wait_for(state="attached", timeout=20000)

    js_rolar = """
    (passo) => {
        const rolaveis = Array.from(document.querySelectorAll('*')).filter(e =>
            e.scrollHeight > e.clientHeight + 50 &&
            ['auto', 'scroll'].includes(getComputedStyle(e).overflowY)
        );
        rolaveis.forEach(e => { e.scrollTop += passo; });
        return rolaveis.map(e => e.tagName.toLowerCase() + '.' + String(e.className).split(' ')[0]);
    }
    """

    rolaveis_encontradas = []
    for _ in range(15):
        if _primeiro_visivel(locator) is not None:
            return
        rolaveis_encontradas = page.evaluate(js_rolar, 500)
        time.sleep(0.7)

    if _primeiro_visivel(locator) is not None:
        return

    # Não encontrou: imprime um diagnóstico para facilitar a investigação
    print(f"\n[DIAGNÓSTICO] Texto '{texto_ancora}' não ficou visível.")
    print(f"[DIAGNÓSTICO] Cópias do texto no HTML: {locator.count()}")
    for i in range(locator.count()):
        item = locator.nth(i)
        print(f"   cópia {i}: visível={item.is_visible()} caixa={item.bounding_box()}")
    print(f"[DIAGNÓSTICO] Áreas roláveis encontradas: {rolaveis_encontradas}")
    try:
        caminho_print = os.path.join(config.PASTA_EXPORTACOES, "erro_ancora.png")
        page.screenshot(path=caminho_print)
        print(f"[DIAGNÓSTICO] Print salvo em: {caminho_print}")
    except Exception:
        pass
    raise Exception(f"Texto-âncora '{texto_ancora}' não ficou visível na página.")


def clicar_botao_da_tabela(page, test_id, texto_ancora):
    """
    Cada visual da página (tabela, título, gráficos) fica dentro da sua
    própria tag <visual-container>, e o botão "..." (test_id) existe em
    TODOS eles, com o mesmo nome/aria-label. Por isso, em vez de procurar o
    botão na página inteira, primeiro localizamos o <visual-container> que
    contém um texto exclusivo da tabela desejada (um cabeçalho de coluna que
    não apareça em nenhum filtro/slicer da página) e buscamos o botão só
    DENTRO dele.
    """
    esperar_ancora(page, texto_ancora)

    container = page.locator("visual-container", has_text=texto_ancora).first

    # O Power BI coloca uma cópia ESCONDIDA do título dentro de cada visual
    # (um <p role="tooltip" aria-hidden="true">, só pra leitores de tela).
    # O .first pegava essa cópia invisível e o hover travava. Por isso
    # passamos o mouse só na cópia que está de fato visível.
    referencia = _primeiro_visivel(container.get_by_text(texto_ancora, exact=False))
    if referencia is None:
        raise Exception(
            f"Dentro do visual, nenhuma cópia visível do texto '{texto_ancora}'."
        )
    referencia.hover()
    time.sleep(1)
    container.get_by_test_id(test_id).click()


def exportar_dados(page, caminho_saida, texto_ancora):
    clicar_botao_da_tabela(page, "visual-more-options-btn", texto_ancora)
    time.sleep(1)

    page.get_by_test_id("pbimenu-item.Exportar dados").click()
    time.sleep(1.5)  # espera a caixa de diálogo de exportação abrir de vez

    # Escolhe a opção "Dados resumidos" -- clicamos pelo texto porque o ID
    # desse radio button (ex: #pbi-radio-button-1) muda toda vez que o
    # diálogo abre (não é fixo). Usamos exact=True porque "Dados resumidos"
    # também aparece dentro do texto explicativo abaixo da opção -- sem o
    # exact=True, o Playwright encontra as duas ocorrências e trava.
    page.get_by_text("Dados resumidos", exact=True).click()
    time.sleep(1)

    # Clica em exportar e captura o arquivo assim que o download começa
    with page.expect_download() as download_info:
        page.get_by_test_id("export-btn").click()
    download = download_info.value

    download.save_as(caminho_saida)
    print(f"OK -> arquivo salvo em: {caminho_saida}")

    # Fecha qualquer aviso/diálogo que tenha ficado aberto após a exportação
    page.keyboard.press("Escape")
    time.sleep(1)


def extrair_do_power_bi():
    """
    Etapa 1. Abre o Power BI, aplica cada indicador e exporta.
    Devolve a lista de arquivos exportados com sucesso nesta rodada.
    """
    exportados = []

    try:
        with sync_playwright() as p:
            if config.OPERA_PATH:
                browser = p.chromium.launch(executable_path=config.OPERA_PATH, headless=False)
            else:
                browser = p.chromium.launch(headless=False)

            context = browser.new_context(accept_downloads=True)
            page = context.new_page()

            print(f"Abrindo o relatório: {config.REPORT_URL}")
            page.goto(config.REPORT_URL)

            print(
                "\n>>> Faça o login manualmente (com MFA se pedir).\n"
                ">>> O script continua sozinho assim que o relatório carregar...\n"
            )
            page.wait_for_selector('[data-testid="bookmarks-btn"]', timeout=180000)
            print(">>> Relatório carregado! Continuando automaticamente...\n")

            for nome_indicador, nome_arquivo, texto_ancora in config.EXTRACOES:
                print(f"\n--- Extraindo: {nome_indicador} ---")
                try:
                    aplicar_indicador(page, nome_indicador)
                    caminho_saida = os.path.join(config.PASTA_EXPORTACOES, nome_arquivo)
                    exportar_dados(page, caminho_saida, texto_ancora)
                    exportados.append(nome_arquivo)
                except Exception as e:
                    print(f"ERRO na extração de '{nome_indicador}': {e}")
                finally:
                    # Garante que nenhum menu/overlay fique aberto antes do
                    # próximo indicador, mesmo se a extração acima tiver
                    # falhado (senão o overlay bloqueia o próximo clique)
                    page.keyboard.press("Escape")
                    time.sleep(1)

            browser.close()
    except Exception as e:
        print(f"\nERRO GERAL: {e}")

    return exportados
