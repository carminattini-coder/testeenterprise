import openpyxl
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, GradientFill
)
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Pipeline Enterprise PagBrasil"

# ── Paleta de cores ──────────────────────────────────────────────────────────
VERDE_ESCURO  = "1A5276"   # cabeçalho
VERDE_MEDIO   = "1E8449"   # linha de grupo
VERDE_CLARO   = "D5F5E3"   # zebra par
BRANCO        = "FFFFFF"
AMARELO       = "FFF9C4"   # A verificar
CINZA_BORDA   = "B2BABB"
AZUL_LINK     = "1F618D"

# ── Colunas ──────────────────────────────────────────────────────────────────
headers = [
    "Nome da Empresa",
    "URL da Loja",
    "Plataforma",
    "Métodos de Pagamento Ativos",
    "Carteiras Digitais",
    "Parceiro(s) de Pagamento",
    "Score Enterprise (estimado)",
    "Prioridade",
    "LinkedIn – Responsável E-commerce",
    "LinkedIn – Responsável Payments",
    "LinkedIn – CFO",
    "Observações",
]

# ── Dados ────────────────────────────────────────────────────────────────────
rows = [
    # ── VTEX ─────────────────────────────────────────────────────────────────
    [
        "Lojas Renner S.A.",
        "lojasrenner.com.br",
        "VTEX",
        "Cartão de crédito, Boleto, PIX, Parcelamento",
        "A verificar",
        "Braspag (Cielo Group)",
        "18/20",
        "Alta",
        "https://www.linkedin.com/in/lucasneves (Lucas Neves – Head of E-commerce)",
        "A verificar",
        "Daniel Martins dos Santos – CFO | https://www.linkedin.com/company/lojas-renner",
        "TPV digital R$ 595,8 mi/trimestre. Case de sucesso VTEX.",
    ],
    [
        "Grupo Boticário (Boticário + Eudora + Beleza na Web)",
        "boticario.com.br / belezanaweb.com.br",
        "VTEX IO",
        "Cartão de crédito, Boleto, PIX, Parcelamento",
        "A verificar",
        "Pagar.me",
        "20/20",
        "Alta",
        "A verificar",
        "A verificar",
        "A verificar (empresa fechada/familiar)",
        "GMV total R$ 35,7 bi em 2024 (+19%). Principal grupo de beleza do Brasil.",
    ],
    [
        "Grupo DPSP (Drogaria SP + Pacheco)",
        "drogariasaopaulo.com.br / pacheco.com.br",
        "VTEX",
        "Cartão de crédito, Boleto, PIX, Parcelamento",
        "Sim",
        "Adyen",
        "20/20",
        "Alta",
        "Leandro Rocha – Diretor de Canais Digitais | https://www.linkedin.com/company/grupodpsp",
        "A verificar",
        "A verificar",
        "Canal digital R$ 2,9 bi (20,7% de R$ 14 bi total 2024). Case confirmado Adyen.",
    ],
    [
        "Grupo SBF – Centauro + Nike Brasil",
        "centauro.com.br / nike.com.br",
        "VTEX",
        "Cartão de crédito, Boleto, PIX, Parcelamento, Apple Pay, Google Pay",
        "Sim",
        "A confirmar (possível Adyen)",
        "20/20",
        "Alta",
        "Gustavo Furtado – CEO (ex-CDO) | https://www.linkedin.com/in/gustavofurtado",
        "A verificar",
        "José Luís Magalhães Salazar – CFO | https://www.linkedin.com/company/grupo-sbf",
        "Digital R$ 675,5 mi em 2024 (21,7% da receita bruta). Apple Pay e Google Pay confirmados.",
    ],
    [
        "Grupo AZZAS 2154 (Hering + Reserva + Farm + Animale + Arezzo + Schutz)",
        "hering.com.br / reserva.ink / farmrio.com / animale.com.br / arezzo.com.br",
        "VTEX IO",
        "Cartão de crédito, Boleto, PIX, Parcelamento",
        "A verificar",
        "A confirmar",
        "20/20",
        "Alta",
        "A verificar (squad VTEX própria desde 2018)",
        "A verificar",
        "Eric Alencar – CFO (posse set/2025) | https://www.linkedin.com/company/azzas2154",
        "28+ marcas. Reserva +30% de vendas com VTEX (dez/2024). Hering +30% ticket médio.",
    ],
    [
        "Vulcabras (Mizuno + Under Armour + Olympikus)",
        "mizuno.com.br / underarmour.com.br / olympikus.com.br",
        "VTEX",
        "Cartão de crédito, Boleto, PIX, Parcelamento",
        "A verificar",
        "A confirmar",
        "17/20",
        "Alta",
        "A verificar",
        "A verificar",
        "A verificar | https://www.linkedin.com/company/vulcabras",
        "Crescimento 120% no digital em 2023. 3 marcas esportivas globais em VTEX.",
    ],
    [
        "Decathlon Brasil",
        "decathlon.com.br",
        "VTEX",
        "Cartão de crédito, Boleto, PIX, Parcelamento",
        "A verificar",
        "A confirmar",
        "18/20",
        "Alta",
        "Rennan Lenner Pinheiro – Executive Director, Digital Transformation | https://www.linkedin.com/in/rennan-lenner-pinheiro",
        "A verificar",
        "A verificar | https://www.linkedin.com/company/decathlon-brazil",
        "Omnichannel com VTEX inStore desde 2021. +5% market share digital após VTEX.",
    ],
    [
        "Whirlpool Brasil (Brastemp + Consul + KitchenAid)",
        "brastemp.com.br / consul.com.br",
        "VTEX",
        "Cartão de crédito, Boleto, PIX, Parcelamento",
        "A verificar",
        "NuPay / Braspag (a confirmar)",
        "18/20",
        "Alta",
        "A verificar | https://www.linkedin.com/company/whirlpool-brazil",
        "A verificar",
        "A verificar",
        "+360% receita Consumer Week 2021. 10+ anos de parceria VTEX.",
    ],
    [
        "C&A Brasil",
        "cea.com.br",
        "VTEX",
        "Cartão de crédito, Boleto, PIX, Parcelamento, C&A Pay",
        "Sim",
        "Adyen",
        "20/20",
        "Alta",
        "A verificar | https://www.linkedin.com/company/c-a-modas-ltda",
        "Filipe Matzembacher (Matz) – Diretor C&A Pay | https://www.linkedin.com/in/filipematz",
        "A verificar",
        "Cliente pioneiro Adyen+VTEX apresentado no VTEX Day. Acima de R$ 1 bi digital estimado.",
    ],
    [
        "Leroy Merlin Brasil",
        "leroymerlin.com.br",
        "VTEX",
        "Cartão de crédito, Boleto, PIX, Parcelamento",
        "A verificar",
        "Pagar.me",
        "18/20",
        "Alta",
        "Sergio Henrique Ferraz – Diretor de Vendas Digitais | https://www.linkedin.com/in/sergioferraz",
        "A verificar",
        "A verificar | https://www.linkedin.com/company/leroy-merlin-brasil",
        "Investiu R$ 300 mi em tecnologia. Parceria Pagar.me confirmada em entrevista.",
    ],
    [
        "MadeiraMadeira",
        "madeiramadeira.com.br",
        "VTEX",
        "Cartão de crédito, Boleto, PIX, Parcelamento",
        "Sim",
        "Adyen",
        "20/20",
        "Alta",
        "Robson Privado – CGO (Chief Growth Officer) | https://www.linkedin.com/company/madeiramadeira",
        "A verificar",
        "Marcelo Scandian – CFO (Co-fundador) | https://www.linkedin.com/in/marcelo-scandian",
        "Maior e-commerce de casa da América Latina. Case Adyen Unified Commerce confirmado.",
    ],
    [
        "Tok&Stok",
        "tokstok.com.br",
        "VTEX",
        "Cartão de crédito, Boleto, PIX, Parcelamento",
        "Sim",
        "Adyen",
        "17/20",
        "Alta",
        "A verificar | https://www.linkedin.com/company/tok-stok",
        "A verificar",
        "A verificar",
        "Case ACCT Global em VTEX CMS. Citada como cliente Adyen+VTEX.",
    ],
    [
        "Riachuelo (RCHLO – Grupo Guararapes)",
        "riachuelo.com.br",
        "VTEX",
        "Cartão de crédito, Boleto, PIX, Parcelamento, Cartão Riachuelo",
        "A verificar",
        "Midway Financeira (própria)",
        "18/20",
        "Media-Alta",
        "A verificar | https://www.linkedin.com/company/riachuelo",
        "A verificar",
        "A verificar",
        "RCHLO Marketplace integrado VTEX. 300+ lojas físicas. Financeira própria Midway.",
    ],
    [
        "Panvel Farmácias",
        "panvel.com",
        "VTEX",
        "Cartão de crédito, Boleto, PIX, Parcelamento",
        "A verificar",
        "A confirmar",
        "16/20",
        "Media-Alta",
        "Tiago Souto – Gerente Executivo de E-commerce e Digital | https://www.linkedin.com/in/tiago-souto",
        "A verificar",
        "A verificar | https://www.linkedin.com/company/panvel-farmácias",
        "Líder em retail media farmacêutico. Marketplace lançado com VTEX.",
    ],
    [
        "Electrolux Brasil",
        "electrolux.com.br",
        "VTEX",
        "Cartão de crédito, Boleto, PIX, Parcelamento",
        "Sim",
        "Adyen",
        "18/20",
        "Alta",
        "Adriana Orsi – D2C Sr. Manager | https://www.linkedin.com/company/electrolux-brasil",
        "Carmen Pedroso – Supervisora de Meios de Pagamento e Fraude LATAM | https://www.linkedin.com/in/carmen-pedroso-72a9ba170",
        "A verificar",
        "D2C + lojas de assistência técnica e outlet via VTEX Sales App. Case Adyen confirmado.",
    ],
    [
        "Vivara (VIVA3)",
        "vivara.com.br",
        "Plataforma própria / Thoughtworks",
        "Cartão de crédito, Boleto, PIX, Parcelamento",
        "A verificar",
        "A confirmar",
        "17/20",
        "Media-Alta",
        "Marina Canto / Viviane Castro – Diretoras | https://www.linkedin.com/company/vivara",
        "A verificar",
        "A verificar (empresa listada B3: VIVA3)",
        "14,4% das vendas online em 2024. +50% conversão após redesenho com Thoughtworks.",
    ],
    [
        "RD Saúde / Raia Drogasil (RADL3)",
        "drogaraia.com.br / drogasil.com.br",
        "VTEX",
        "Cartão de crédito, Boleto, PIX, Parcelamento",
        "A verificar",
        "A confirmar",
        "20/20",
        "Alta",
        "A verificar | https://www.linkedin.com/company/rd-saude",
        "A verificar",
        "A verificar (empresa listada B3: RADL3)",
        "15,2% de receita digital. Líder nacional farmacêutico. RD Marketplace no VTEX Day.",
    ],
    [
        "Natura (Natura &Co)",
        "natura.com.br",
        "VTEX",
        "Cartão de crédito, Boleto, PIX, Parcelamento",
        "A verificar",
        "A confirmar",
        "20/20",
        "Alta",
        "A verificar | https://www.linkedin.com/company/natura-cosmeticos",
        "A verificar",
        "A verificar (empresa listada B3: NTCO3)",
        "Grupo com Avon, The Body Shop e Aesop. Histórico VTEX + cliente Raccoon/VTEX.",
    ],
    [
        "ASICS Brasil",
        "asics.com/br",
        "VTEX",
        "Cartão de crédito, Boleto, PIX, Parcelamento",
        "Sim",
        "Adyen",
        "17/20",
        "Alta",
        "A verificar | https://www.linkedin.com/company/asics",
        "A verificar",
        "A verificar",
        "Citada no case Adyen de clientes VTEX Brasil. Comércio unificado on/off.",
    ],
    [
        "Polishop",
        "polishop.com.br",
        "VTEX",
        "Cartão de crédito, Boleto, PIX, Parcelamento",
        "A verificar",
        "A confirmar",
        "15/20",
        "Media",
        "A verificar | https://www.linkedin.com/company/polishop",
        "A verificar",
        "A verificar",
        "Implementação nativa SaaS VTEX. 260 lojas físicas + e-commerce. João Appolinário palestrante VTEX Day 2024.",
    ],
    # ── SHOPIFY PLUS ─────────────────────────────────────────────────────────
    [
        "Havaianas (Alpargatas – ALPA3)",
        "havaianas.com",
        "Shopify Plus",
        "Cartão de crédito, Boleto, PIX, Parcelamento",
        "A verificar",
        "A confirmar (possível PagBrasil / Mercado Pago)",
        "18/20",
        "Alta",
        "A verificar | https://www.linkedin.com/company/alpargatas",
        "A verificar",
        "A verificar (empresa listada B3: ALPA3)",
        "Confirmada pela Shopify Brasil como cliente Plus. Marca global com D2C forte.",
    ],
    [
        "Sallve",
        "sallve.com.br",
        "Shopify Plus",
        "Cartão de crédito, Boleto, PIX, Parcelamento",
        "A verificar",
        "A confirmar (possível PagBrasil / Mercado Pago)",
        "15/20",
        "Media-Alta",
        "A verificar | https://www.linkedin.com/company/sallve",
        "A verificar",
        "A verificar",
        "D2C de cosméticos premium. Confirmada Shopify Plus. Estimado R$ 60-100 mi/ano.",
    ],
    [
        "Grand Cru (vinhos premium)",
        "grandcru.com.br",
        "Shopify Plus",
        "Cartão de crédito, Boleto",
        "A verificar",
        "MOIP (a confirmar parceiro atual)",
        "15/20",
        "Media",
        "A verificar | https://www.linkedin.com/company/grand-cru",
        "A verificar",
        "A verificar",
        "Case Shopify Brasil confirmado. E-commerce cresceu de 4-5% para relevante do faturamento.",
    ],
    [
        "Coffee++ (cafés especiais)",
        "coffeemais.com.br",
        "Shopify Plus",
        "Cartão de crédito, Boleto, PIX, Parcelamento, Assinatura",
        "A verificar",
        "A confirmar",
        "14/20",
        "Media",
        "A verificar | https://www.linkedin.com/company/coffeemais",
        "A verificar",
        "A verificar",
        "Confirmada Shopify Plus. Clube de assinaturas >500 membros em 2 meses no início.",
    ],
    [
        "Dr. Jones (beleza/cuidados masculinos)",
        "drjones.com.br",
        "Shopify Plus",
        "Cartão de crédito, Boleto, PIX, Parcelamento, Assinatura",
        "A verificar",
        "A confirmar",
        "15/20",
        "Media-Alta",
        "A verificar | https://www.linkedin.com/company/dr-jones",
        "A verificar",
        "A verificar",
        "Confirmada Shopify Plus. D2C premium masculino. Estimado acima de R$ 60 mi/ano.",
    ],
    [
        "Farm Rio – operação EUA",
        "farmrio.com",
        "Shopify (EUA) / VTEX (Brasil)",
        "Cartão de crédito (EUA), PIX, Boleto (Brasil)",
        "A verificar",
        "A confirmar",
        "17/20",
        "Media-Alta",
        "A verificar | https://www.linkedin.com/company/farmrio",
        "A verificar",
        "Eric Alencar – CFO AZZAS 2154 | https://www.linkedin.com/company/azzas2154",
        "Expansão América Latina com VTEX. Operação USA via Shopify (a confirmar).",
    ],
    # ── PLATAFORMA A CONFIRMAR ────────────────────────────────────────────────
    [
        "Netshoes / Zattini (Magazine Luiza)",
        "netshoes.com.br / zattini.com.br",
        "VTEX (marketplace integrado)",
        "Cartão de crédito, Boleto, PIX, Parcelamento",
        "A verificar",
        "A confirmar (Magalu Pay possível)",
        "17/20",
        "Media-Alta",
        "Pedro Bertolacini – Director, Marketplace | https://www.linkedin.com/in/pedro-bertolacini",
        "A verificar",
        "Consolidado no Magalu | https://www.linkedin.com/company/magazineluiza",
        "54 mi visitantes únicos/mês. Maior e-commerce esportivo América Latina.",
    ],
    [
        "Dafiti Brasil (GFG Group)",
        "dafiti.com.br",
        "Plataforma própria + conector VTEX",
        "Cartão de crédito, Boleto, PIX, Parcelamento",
        "A verificar",
        "A confirmar",
        "17/20",
        "Media",
        "A verificar | https://www.linkedin.com/company/dafiti-com",
        "A verificar",
        "A verificar",
        "7,7 mi clientes ativos. Receita estimada acima de R$ 3 bi. Integra sellers via VTEX nativo.",
    ],
]

# ── Estilos ───────────────────────────────────────────────────────────────────
def make_fill(hex_color):
    return PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")

def make_border():
    thin = Side(style="thin", color=CINZA_BORDA)
    return Border(left=thin, right=thin, top=thin, bottom=thin)

header_font  = Font(name="Calibri", bold=True, color=BRANCO, size=11)
cell_font    = Font(name="Calibri", size=10)
link_font    = Font(name="Calibri", size=10, color=AZUL_LINK, underline="single")
warn_fill    = make_fill(AMARELO)
header_fill  = make_fill(VERDE_ESCURO)
alt_fill     = make_fill(VERDE_CLARO)
border       = make_border()

# ── Cabeçalho ────────────────────────────────────────────────────────────────
for col_idx, header in enumerate(headers, start=1):
    cell = ws.cell(row=1, column=col_idx, value=header)
    cell.fill   = header_fill
    cell.font   = header_font
    cell.border = border
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

ws.row_dimensions[1].height = 40

# ── Dados ────────────────────────────────────────────────────────────────────
for row_idx, row_data in enumerate(rows, start=2):
    fill = alt_fill if row_idx % 2 == 0 else make_fill(BRANCO)
    for col_idx, value in enumerate(row_data, start=1):
        cell = ws.cell(row=row_idx, column=col_idx, value=value)
        cell.border    = border
        cell.alignment = Alignment(vertical="top", wrap_text=True)

        # Destacar "A verificar" / "A confirmar" em amarelo
        if "A verificar" in str(value) or "A confirmar" in str(value):
            cell.fill = warn_fill
            cell.font = Font(name="Calibri", size=10, color="7D6608")
        # Links em azul
        elif "linkedin.com" in str(value).lower() or "http" in str(value).lower():
            cell.fill = fill
            cell.font = link_font
        else:
            cell.fill = fill
            cell.font = cell_font

        # Coluna Prioridade colorida
        if col_idx == 8:
            prioridade = str(value)
            if prioridade == "Alta":
                cell.fill = make_fill("1E8449")
                cell.font = Font(name="Calibri", size=10, bold=True, color=BRANCO)
            elif prioridade == "Media-Alta":
                cell.fill = make_fill("F39C12")
                cell.font = Font(name="Calibri", size=10, bold=True, color=BRANCO)
            elif prioridade == "Media":
                cell.fill = make_fill("2E86C1")
                cell.font = Font(name="Calibri", size=10, bold=True, color=BRANCO)

    ws.row_dimensions[row_idx].height = 55

# ── Larguras das colunas ──────────────────────────────────────────────────────
col_widths = [32, 40, 20, 45, 14, 35, 22, 12, 55, 55, 55, 55]
for i, width in enumerate(col_widths, start=1):
    ws.column_dimensions[get_column_letter(i)].width = width

# ── Freeze pane ───────────────────────────────────────────────────────────────
ws.freeze_panes = "A2"

# ── Filtro automático ─────────────────────────────────────────────────────────
ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"

# ── Aba de legenda ────────────────────────────────────────────────────────────
ws2 = wb.create_sheet("Legenda & Metodologia")
legenda_rows = [
    ["LEGENDA E METODOLOGIA", ""],
    ["", ""],
    ["Score Enterprise (estimado)", "Baseado nos 4 critérios do documento Enterprise 2026 PagBrasil:"],
    ["", "• Volume Financeiro Relevante (0–5 pts): 0 = até R$30M/ano | 3 = R$30M–60M/ano | 5 = acima de R$60M/ano"],
    ["", "• Estrutura Organizacional Consolidada (0–5 pts): 0 = decisão centralizada | 3 = 2-3 áreas | 5 = comitê completo"],
    ["", "• Capacidade Técnica Interna (0–5 pts): 0 = sem time técnico | 3 = 1-2 devs | 5 = squad estruturada/CTO"],
    ["", "• Relevância Estratégica no Segmento (0–5 pts): 0 = pouca relevância | 3 = top 20 | 5 = top 10 nacional"],
    ["", ""],
    ["Classificação Final", "16–20 = Enterprise Estratégico | 12–15 = Enterprise | 8–11 = Mid-Market | <8 = SMB"],
    ["", ""],
    ["Prioridade", "Alta = Adyen confirmado (oportunidade direta de substituição) ou volume digital >R$500M"],
    ["", "Media-Alta = Volume digital R$60M–500M ou plataforma VTEX/Shopify com parceiro a confirmar"],
    ["", "Media = Volume digital estimado R$30M–60M ou plataforma a confirmar"],
    ["", ""],
    ["Campos 'A verificar'", "Requerem acesso ao checkout ao vivo, Wappalyzer/BuiltWith, ou pesquisa direta no LinkedIn"],
    ["", ""],
    ["Fontes Principais", "VTEX Case Studies | Adyen Customer Stories | Relatórios IR (B3) | LinkedIn | ABComm | E-bit Nielsen"],
    ["", ""],
    ["Data de geração", "Maio 2026"],
    ["Responsável", "PagBrasil – Frente Enterprise"],
]
for r_idx, (col_a, col_b) in enumerate(legenda_rows, start=1):
    c1 = ws2.cell(row=r_idx, column=1, value=col_a)
    c2 = ws2.cell(row=r_idx, column=2, value=col_b)
    c2.alignment = Alignment(wrap_text=True)
    if r_idx == 1:
        c1.font = Font(name="Calibri", bold=True, size=14, color=VERDE_ESCURO)
    elif col_a and col_a != "":
        c1.font = Font(name="Calibri", bold=True, size=11)
        c1.fill = make_fill(VERDE_CLARO)
ws2.column_dimensions["A"].width = 30
ws2.column_dimensions["B"].width = 100

# ── Salvar ────────────────────────────────────────────────────────────────────
output_path = "/home/user/testeenterprise/Pipeline_Enterprise_PagBrasil_2026.xlsx"
wb.save(output_path)
print(f"Arquivo salvo em: {output_path}")
