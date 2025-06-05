import asyncio
from playwright.async_api import async_playwright
import openpyxl

estados = {
    "Acre": "ac", "Alagoas": "al", "Amapá": "ap", "Amazonas": "am", "Bahia": "ba", "Ceará": "ce",
    "Distrito Federal": "df", "Espírito Santo": "es", "Goiás": "go", "Maranhão": "ma",
    "Mato Grosso": "mt", "Mato Grosso do Sul": "ms", "Minas Gerais": "mg", "Pará": "pa",
    "Paraíba": "pb", "Paraná": "pr", "Pernambuco": "pe", "Piauí": "pi", "Rio de Janeiro": "rj",
    "Rio Grande do Norte": "rn", "Rio Grande do Sul": "rs", "Rondônia": "ro", "Roraima": "rr",
    "Santa Catarina": "sc", "São Paulo": "sp", "Sergipe": "se", "Tocantins": "to"
}

dados_finais = []

async def extrair_dado_por_texto(page, aba, texto_alvo):
    try:
        linha = page.locator(f"tr:has(th.lista__titulo:has-text('{aba}'))")
        await linha.scroll_into_view_if_needed()
        seta = linha.locator("th.lista__seta")
        await seta.click(force=True)
        await page.wait_for_timeout(2000)

        indicadores = page.locator("tr.lista__indicador")
        total = await indicadores.count()

        for i in range(total):
            row = indicadores.nth(i)
            nome = await row.locator("td.lista__nome").text_content()
            if texto_alvo.lower() in nome.lower():
                valor = await row.locator("td.lista__valor").text_content()
                return valor.strip()
        return "Indicador não encontrado"
    except Exception as e:
        return f"Erro ao coletar: {str(e)}"

async def extrair_dados_estado(page, estado, sigla):
    print(f"Coletando dados de: {estado}")
    url = f"https://cidades.ibge.gov.br/brasil/{sigla}/panorama"
    await page.goto(url, timeout=60000)
    await page.wait_for_timeout(3000)

    resultado = {"Estado": estado}

    try:
        bloco = page.locator("div.indicador__valor:below(h2:has-text('População'))")
        await bloco.first.wait_for(timeout=8000)
        texto = await bloco.first.inner_text()
        resultado["População"] = texto.strip()
    except Exception as e:
        resultado["População"] = f"Erro ao coletar: {str(e)}"

    resultado["Educação"] = await extrair_dado_por_texto(page, "Educação", "IDEB – Anos iniciais do ensino fundamental")
    resultado["Trabalho e Rendimento"] = await extrair_dado_por_texto(page, "Trabalho e Rendimento", "Rendimento nominal mensal domiciliar")
    resultado["Economia"] = await extrair_dado_por_texto(page, "Economia", "Total de receitas brutas realizadas")
    resultado["Meio Ambiente"] = await extrair_dado_por_texto(page, "Meio Ambiente", "")
    resultado["Território"] = await extrair_dado_por_texto(page, "Território", "Área da unidade territorial")

    print(f"Dados coletados com sucesso para: {estado}")
    return resultado

async def run_coleta():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page()

        for estado, sigla in estados.items():
            try:
                dados = await extrair_dados_estado(page, estado, sigla)
                dados_finais.append(dados)
            except Exception as e:
                print(f"Erro em {estado}: {e}")

        await browser.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Dados IBGE"
    ws.append([
        "Estado", "População", "Educação", "Trabalho e Rendimento",
        "Economia", "Meio Ambiente", "Território"
    ])

    for estado in dados_finais:
        ws.append([
            estado["Estado"], estado["População"], estado["Educação"],
            estado["Trabalho e Rendimento"], estado["Economia"],
            estado["Meio Ambiente"], estado["Território"]
        ])

    wb.save("resultado_ibge_final_detalhado.xlsx")
    print("Planilha salva: resultado_ibge_final_detalhado.xlsx")

if __name__ == "__main__":
    asyncio.run(run_coleta())
