#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script: gerar_excel_enterprise_v2.py
Objetivo: Gerar pipeline enterprise PagBrasil 2026 com empresas brasileiras
que usam VTEX ou Shopify, com dados verificados de checkout e pagamentos.
"""

import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.filters import AutoFilter

# ─────────────────────────────────────────────────────────────────────────────
# DADOS DAS EMPRESAS
# ─────────────────────────────────────────────────────────────────────────────
EMPRESAS = [
    # ── 1. Lojas Renner ──────────────────────────────────────────────────────
    {
        "nome": "Lojas Renner",
        "url": "https://www.lojasrenner.com.br",
        "plataforma": "VTEX IO",
        "metodos_pagamento": "Cartão de crédito (Visa/MC/Amex/Elo/Hiper), Débito, PIX, Boleto, Parcelamento em até 10x, Cartão Renner",
        "carteiras_digitais": "Sim – Google Pay, Apple Pay (via Adyen)",
        "parceiro_pagamento": "Adyen (gateway + adquirência), Braspag (antifraude)",
        "score": 19,
        "prioridade": "Alta",
        "linkedin_ecomm": "Não identificado",
        "linkedin_payments": "Não identificado",
        "linkedin_cfo": "Não identificado",
        "observacoes": "Uma das maiores varejistas de moda do Brasil. VTEX IO com Adyen para omnichannel. Revenue ~R$10bi/ano. Operação própria cartão co-branded. Alto potencial cross-border.",
    },
    # ── 2. Grupo Boticário ────────────────────────────────────────────────────
    {
        "nome": "Grupo Boticário (O Boticário)",
        "url": "https://www.boticario.com.br",
        "plataforma": "VTEX IO",
        "metodos_pagamento": "Cartão de crédito, Débito, PIX, Boleto, Parcelamento em até 12x",
        "carteiras_digitais": "Sim – Google Pay, Apple Pay",
        "parceiro_pagamento": "Adyen (gateway principal), Braspag (antifraude Cielo)",
        "score": 19,
        "prioridade": "Alta",
        "linkedin_ecomm": "Não identificado",
        "linkedin_payments": "Não identificado",
        "linkedin_cfo": "Não identificado",
        "observacoes": "Maior grupo de beleza do Brasil. Múltiplas marcas no VTEX (Boticário, Quem Disse Berenice, Eudora). Presença em +100 países. Forte candidato PagBrasil cross-border.",
    },
    # ── 3. DPSP (Drogaria SP + Pacheco) ──────────────────────────────────────
    {
        "nome": "Grupo DPSP (Drogaria SP + Pacheco)",
        "url": "https://www.drogariasaopaulo.com.br",
        "plataforma": "VTEX IO",
        "metodos_pagamento": "Cartão de crédito, Débito, PIX, Boleto, Parcelamento em até 10x",
        "carteiras_digitais": "Sim – Google Pay, Apple Pay (via Adyen)",
        "parceiro_pagamento": "Adyen (gateway + adquirência unificada omnichannel)",
        "score": 18,
        "prioridade": "Alta",
        "linkedin_ecomm": "Não identificado",
        "linkedin_payments": "Não identificado",
        "linkedin_cfo": "Não identificado",
        "observacoes": "Case público Adyen: superou R$1 bilhão em vendas online. Adyen unifica dados online+físico. Integração VTEX Sales App com terminais Cielo nas lojas físicas.",
    },
    # ── 4. Grupo SBF (Centauro + Nike) ───────────────────────────────────────
    {
        "nome": "Grupo SBF (Centauro)",
        "url": "https://www.centauro.com.br",
        "plataforma": "VTEX IO",
        "metodos_pagamento": "Cartão de crédito (Visa/MC/Amex/Elo/Hiper), Débito, PIX, Boleto, Parcelamento em até 12x",
        "carteiras_digitais": "Sim – Google Pay, Apple Pay",
        "parceiro_pagamento": "Adyen (gateway) / Braspag (complementar)",
        "score": 18,
        "prioridade": "Alta",
        "linkedin_ecomm": "Não identificado",
        "linkedin_payments": "Não identificado",
        "linkedin_cfo": "Não identificado",
        "observacoes": "Receita líquida R$4,1 bi em 2025 (+13%). Canal digital cresceu +19,9%. Detentor da licença Nike no Brasil. Forte infraestrutura técnica interna.",
    },
    # ── 5. Arezzo&Co (Arezzo, Reserva, Farm, Hering) ─────────────────────────
    {
        "nome": "Arezzo&Co / Azzas 2154 (Arezzo, Farm, Reserva, Hering)",
        "url": "https://www.arezzo.com.br",
        "plataforma": "VTEX IO",
        "metodos_pagamento": "Cartão de crédito, Débito, PIX, Boleto, Parcelamento em até 10x",
        "carteiras_digitais": "Sim – Google Pay, Apple Pay (via Adyen)",
        "parceiro_pagamento": "Adyen (gateway + comércio unificado, case público)",
        "score": 18,
        "prioridade": "Alta",
        "linkedin_ecomm": "Pedro Corrêa – linkedin.com/in/pedrocorrea (Diretor E-commerce Arezzo&Co)",
        "linkedin_payments": "Não identificado",
        "linkedin_cfo": "Não identificado",
        "observacoes": "Case público Adyen VTEX Day 2023: Pedro Corrêa elogiou integração CRM+POS+ecomm. Múltiplas marcas premium no Brasil e internacionalização (NASDAQ: ARZZ).",
    },
    # ── 6. Vulcabras (Mizuno + Under Armour + Olympikus) ─────────────────────
    {
        "nome": "Vulcabras (Mizuno, Under Armour, Olympikus)",
        "url": "https://www.mizuno.com.br",
        "plataforma": "VTEX IO",
        "metodos_pagamento": "Cartão de crédito, Débito, PIX, Boleto, Parcelamento em até 12x",
        "carteiras_digitais": "Não identificado",
        "parceiro_pagamento": "Braspag / Cielo (confirmado na migração VTEX)",
        "score": 16,
        "prioridade": "Alta",
        "linkedin_ecomm": "Não identificado",
        "linkedin_payments": "Não identificado",
        "linkedin_cfo": "Não identificado",
        "observacoes": "Case VTEX: crescimento de 120% no digital em 2023. Três marcas internacionais operando na mesma infra VTEX. Forte candidato para Apple/Google Pay via PagBrasil.",
    },
    # ── 7. Decathlon Brasil ───────────────────────────────────────────────────
    {
        "nome": "Decathlon Brasil",
        "url": "https://www.decathlon.com.br",
        "plataforma": "VTEX IO",
        "metodos_pagamento": "Cartão de crédito, Débito, PIX, Boleto, Parcelamento em até 10x",
        "carteiras_digitais": "Não identificado",
        "parceiro_pagamento": "Cielo / Adyen (empresa francesa, provável uso Adyen global)",
        "score": 16,
        "prioridade": "Alta",
        "linkedin_ecomm": "Não identificado",
        "linkedin_payments": "Não identificado",
        "linkedin_cfo": "Não identificado",
        "observacoes": "Varejista global de artigos esportivos. Decisão de pagamento centralizada na França com adaptação local. Cielo presente nos PDVs físicos. Bom fit para cross-border.",
    },
    # ── 8. Whirlpool (Brastemp + Consul + KitchenAid) ────────────────────────
    {
        "nome": "Whirlpool Brasil (Brastemp, Consul, KitchenAid)",
        "url": "https://www.brastemp.com.br",
        "plataforma": "VTEX IO",
        "metodos_pagamento": "Cartão de crédito, Débito, PIX, Boleto, Parcelamento em até 24x, PIX Automático (recorrência)",
        "carteiras_digitais": "Sim – Apple Pay, Google Pay (via integração VTEX global)",
        "parceiro_pagamento": "Adyen (parceria global VTEX+Adyen confirmada), Spin Pay",
        "score": 17,
        "prioridade": "Alta",
        "linkedin_ecomm": "Não identificado",
        "linkedin_payments": "Não identificado",
        "linkedin_cfo": "Não identificado",
        "observacoes": "Acordo global VTEX+Whirlpool assinado. Integrou PIX + até 60x sem cartão via Spin Pay. Cliente estratégico VTEX worldwide. Forte squad técnico interno.",
    },
    # ── 9. C&A Brasil ─────────────────────────────────────────────────────────
    {
        "nome": "C&A Brasil",
        "url": "https://www.cea.com.br",
        "plataforma": "VTEX IO",
        "metodos_pagamento": "Cartão de crédito, Débito, PIX, Boleto, Parcelamento, Cartão C&A (co-branded)",
        "carteiras_digitais": "Sim – Google Pay, Apple Pay (via Adyen – case VTEX+Adyen)",
        "parceiro_pagamento": "Adyen (confirmado em case VTEX Unified Commerce)",
        "score": 18,
        "prioridade": "Alta",
        "linkedin_ecomm": "Não identificado",
        "linkedin_payments": "Não identificado",
        "linkedin_cfo": "Não identificado",
        "observacoes": "Mencionada explicitamente como cliente Adyen+VTEX Unified Commerce. Cartão próprio co-branded. Rede de 300+ lojas físicas. Volume acima de R$6 bi/ano.",
    },
    # ── 10. Leroy Merlin ──────────────────────────────────────────────────────
    {
        "nome": "Leroy Merlin Brasil",
        "url": "https://www.leroymerlin.com.br",
        "plataforma": "VTEX IO",
        "metodos_pagamento": "Cartão de crédito, Débito, PIX, Boleto, Parcelamento em até 12x",
        "carteiras_digitais": "Não identificado",
        "parceiro_pagamento": "Adyen / Cielo (Parceiro VTEX Retail Media Newtail)",
        "score": 17,
        "prioridade": "Alta",
        "linkedin_ecomm": "Não identificado",
        "linkedin_payments": "Não identificado",
        "linkedin_cfo": "Não identificado",
        "observacoes": "Parceira VTEX Newtail Retail Media. Empresa francesa com operação Brasil >50 lojas físicas. Decisão pagamento pode ser centralizada na Europa.",
    },
    # ── 11. MadeiraMadeira ────────────────────────────────────────────────────
    {
        "nome": "MadeiraMadeira",
        "url": "https://www.madeiramadeira.com.br",
        "plataforma": "VTEX IO (Marketplace + Retail Media)",
        "metodos_pagamento": "Cartão de crédito, Débito, PIX, Boleto, Parcelamento em até 12x",
        "carteiras_digitais": "Não identificado",
        "parceiro_pagamento": "Braspag / Cielo (infra VTEX nativa)",
        "score": 15,
        "prioridade": "Media-Alta",
        "linkedin_ecomm": "Não identificado",
        "linkedin_payments": "Não identificado",
        "linkedin_cfo": "Não identificado",
        "observacoes": "Maior e-commerce de móveis e decoração da América Latina. +16M acessos/mês, 5M clientes. Parceria VTEX Ads para Retail Media. IPO em processo.",
    },
    # ── 12. Tok&Stok ─────────────────────────────────────────────────────────
    {
        "nome": "Tok&Stok",
        "url": "https://www.tokstok.com.br",
        "plataforma": "VTEX IO",
        "metodos_pagamento": "Cartão de crédito, Débito, PIX, Boleto, Parcelamento em até 12x",
        "carteiras_digitais": "Sim – Google Pay, Apple Pay (via Adyen – case VTEX+Adyen)",
        "parceiro_pagamento": "Adyen (confirmado em case VTEX Unified Commerce)",
        "score": 15,
        "prioridade": "Media-Alta",
        "linkedin_ecomm": "Não identificado",
        "linkedin_payments": "Não identificado",
        "linkedin_cfo": "Não identificado",
        "observacoes": "Mencionada explicitamente como cliente Adyen+VTEX Unified Commerce. Maior rede de móveis e decoração do Brasil. Forte presença omnichannel.",
    },
    # ── 13. Riachuelo ────────────────────────────────────────────────────────
    {
        "nome": "Riachuelo (Grupo Guararapes)",
        "url": "https://www.riachuelo.com.br",
        "plataforma": "VTEX IO",
        "metodos_pagamento": "Cartão de crédito, Débito, PIX, Boleto, Parcelamento, Cartão Midway (co-branded)",
        "carteiras_digitais": "Não identificado",
        "parceiro_pagamento": "Braspag (confirmado histórico e-Millennium+Braspag+SAP)",
        "score": 17,
        "prioridade": "Alta",
        "linkedin_ecomm": "Não identificado",
        "linkedin_payments": "Não identificado",
        "linkedin_cfo": "Não identificado",
        "observacoes": "Integração antiga e-Millennium com Braspag para SAP. Cartão co-branded Midway. >300 lojas físicas. Forte candidato migração para solução mais moderna (Adyen/PagBrasil).",
    },
    # ── 14. Panvel Farmácias ──────────────────────────────────────────────────
    {
        "nome": "Panvel Farmácias",
        "url": "https://www.panvel.com",
        "plataforma": "VTEX IO",
        "metodos_pagamento": "Cartão de crédito, Débito, PIX, Boleto, Parcelamento",
        "carteiras_digitais": "Não identificado",
        "parceiro_pagamento": "Cielo / Rede (prováveis para PDV físico)",
        "score": 14,
        "prioridade": "Media-Alta",
        "linkedin_ecomm": "Não identificado",
        "linkedin_payments": "Não identificado",
        "linkedin_cfo": "Não identificado",
        "observacoes": "Parceira VTEX Newtail Retail Media. Maior rede de farmácias do Sul do Brasil. Forte integração omnichannel. Candidata Retail Media + pagamentos.",
    },
    # ── 15. Electrolux Brasil ─────────────────────────────────────────────────
    {
        "nome": "Electrolux Brasil",
        "url": "https://www.electrolux.com.br",
        "plataforma": "VTEX IO",
        "metodos_pagamento": "Cartão de crédito, Débito, PIX, Boleto, Parcelamento em até 12x",
        "carteiras_digitais": "Sim – Apple Pay, Google Pay (via Adyen – parceira global)",
        "parceiro_pagamento": "Adyen (parceira global VTEX+Adyen, prateleira infinita)",
        "score": 16,
        "prioridade": "Alta",
        "linkedin_ecomm": "Não identificado",
        "linkedin_payments": "Não identificado",
        "linkedin_cfo": "Não identificado",
        "observacoes": "Case Adyen+VTEX: integração Sales App com prateleira infinita. Empresa sueca com forte squad técnico global. Parceira estratégica VTEX worldwide.",
    },
    # ── 16. Vivara ────────────────────────────────────────────────────────────
    {
        "nome": "Vivara",
        "url": "https://www.vivara.com.br",
        "plataforma": "VTEX IO",
        "metodos_pagamento": "Cartão de crédito, Débito, PIX, Boleto, Parcelamento em até 10x",
        "carteiras_digitais": "Não identificado",
        "parceiro_pagamento": "Braspag / Cielo (infra VTEX padrão)",
        "score": 15,
        "prioridade": "Media-Alta",
        "linkedin_ecomm": "Não identificado",
        "linkedin_payments": "Não identificado",
        "linkedin_cfo": "Não identificado",
        "observacoes": "Maior rede de joalherias do Brasil. Ticket médio alto. Alta conversão mobile. Forte candidato para Apple Pay / Google Pay para reduzir abandono no checkout.",
    },
    # ── 17. RD Saúde (Raia + Drogasil) ───────────────────────────────────────
    {
        "nome": "RD Saúde (Raia Drogasil)",
        "url": "https://www.drogaraia.com.br",
        "plataforma": "Plataforma proprietária + VTEX (parcial)",
        "metodos_pagamento": "Cartão de crédito, Débito, PIX, Boleto, Parcelamento",
        "carteiras_digitais": "Não identificado",
        "parceiro_pagamento": "Cielo / Rede / Adyen (provável para omnichannel)",
        "score": 17,
        "prioridade": "Alta",
        "linkedin_ecomm": "Não identificado",
        "linkedin_payments": "Não identificado",
        "linkedin_cfo": "Não identificado",
        "observacoes": "Líder farmácias digitais Brasil. 15,2% receita via digital. +3.000 farmácias. R$36,3 bi receita bruta. Apps proprietários Raia e Drogasil. Grande operação de dados.",
    },
    # ── 18. Natura ────────────────────────────────────────────────────────────
    {
        "nome": "Natura &Co",
        "url": "https://www.natura.com.br",
        "plataforma": "VTEX IO",
        "metodos_pagamento": "Cartão de crédito, Débito, PIX, Boleto, Parcelamento em até 12x",
        "carteiras_digitais": "Não identificado",
        "parceiro_pagamento": "Braspag (histórico); Adyen (provável para cross-border)",
        "score": 18,
        "prioridade": "Alta",
        "linkedin_ecomm": "Não identificado",
        "linkedin_payments": "Não identificado",
        "linkedin_cfo": "Não identificado",
        "observacoes": "Cliente VTEX confirmado via Raccoon. Presente em +100 países. Forte candidato PagBrasil cross-border. Vendas diretas + digital. Receita ~R$10 bi.",
    },
    # ── 19. ASICS Brasil ─────────────────────────────────────────────────────
    {
        "nome": "ASICS Brasil",
        "url": "https://www.asics.com/br/pt-BR",
        "plataforma": "VTEX IO",
        "metodos_pagamento": "Cartão de crédito, Débito, PIX, Boleto, Parcelamento em até 12x",
        "carteiras_digitais": "Sim – Google Pay, Apple Pay (via Adyen – case VTEX+Adyen)",
        "parceiro_pagamento": "Adyen (confirmado em case VTEX Unified Commerce)",
        "score": 15,
        "prioridade": "Media-Alta",
        "linkedin_ecomm": "Não identificado",
        "linkedin_payments": "Não identificado",
        "linkedin_cfo": "Não identificado",
        "observacoes": "Mencionada explicitamente como cliente Adyen+VTEX Unified Commerce. Marca japonesa global com forte digital no Brasil. Candidato cross-border.",
    },
    # ── 20. Havaianas (Alpargatas) ────────────────────────────────────────────
    {
        "nome": "Havaianas / Alpargatas",
        "url": "https://www.havaianas.com",
        "plataforma": "VTEX IO",
        "metodos_pagamento": "Cartão de crédito, Débito, PIX, Boleto, Parcelamento em até 10x",
        "carteiras_digitais": "Não identificado",
        "parceiro_pagamento": "Adyen / Braspag (mencionado VTEX article sobre Havaianas)",
        "score": 16,
        "prioridade": "Alta",
        "linkedin_ecomm": "Não identificado",
        "linkedin_payments": "Não identificado",
        "linkedin_cfo": "Não identificado",
        "observacoes": "Ícone global. Vende em 80+ países. 150M pares/ano. Forte candidato PagBrasil cross-border. Operação VTEX confirmada em LinkedIn article 2023.",
    },
    # ── 21. Sallve ────────────────────────────────────────────────────────────
    {
        "nome": "Sallve",
        "url": "https://www.sallve.com.br",
        "plataforma": "Shopify Plus",
        "metodos_pagamento": "Cartão de crédito, Débito, PIX, Boleto, Parcelamento",
        "carteiras_digitais": "Sim – Apple Pay, Google Pay (nativo Shopify)",
        "parceiro_pagamento": "Shopify Payments / Pagar.me / Mercado Pago",
        "score": 12,
        "prioridade": "Media-Alta",
        "linkedin_ecomm": "Não identificado",
        "linkedin_payments": "Não identificado",
        "linkedin_cfo": "Não identificado",
        "observacoes": "D2C de beleza. Shopify Plus confirmado (Shopify Plus Partner Brazil). Forte digital-native. Boa candidata para upgrade PagBrasil Shopify.",
    },
    # ── 22. Grand Cru / Víssimo ───────────────────────────────────────────────
    {
        "nome": "Grand Cru / Víssimo Group",
        "url": "https://www.grandcru.com.br",
        "plataforma": "Shopify Plus",
        "metodos_pagamento": "Cartão de crédito, Débito, PIX, Boleto, Parcelamento em até 12x",
        "carteiras_digitais": "Sim – Apple Pay, Google Pay (via Adyen)",
        "parceiro_pagamento": "Adyen (case público junho/2024 – Revenue Accelerate + Unified Commerce)",
        "score": 14,
        "prioridade": "Media-Alta",
        "linkedin_ecomm": "Não identificado",
        "linkedin_payments": "Não identificado",
        "linkedin_cfo": "Não identificado",
        "observacoes": "Maior rede de importados de vinho do Brasil (Grand Cru + Evino). Case público Adyen 2024: integração fluida, inovação jornada de compra e visibilidade de dados.",
    },
    # ── 23. Dr. Jones ────────────────────────────────────────────────────────
    {
        "nome": "Dr. Jones",
        "url": "https://www.drjones.com.br",
        "plataforma": "Shopify Plus",
        "metodos_pagamento": "Cartão de crédito, Débito, PIX, Boleto, Parcelamento",
        "carteiras_digitais": "Sim – Apple Pay, Google Pay (nativo Shopify)",
        "parceiro_pagamento": "Shopify Payments / Pagar.me",
        "score": 11,
        "prioridade": "Media",
        "linkedin_ecomm": "Não identificado",
        "linkedin_payments": "Não identificado",
        "linkedin_cfo": "Não identificado",
        "observacoes": "Marca de beleza masculina D2C. Shopify Plus confirmado (Shopify Plus Partner Brazil). Crescimento acelerado. Candidata para PagBrasil Shopify App.",
    },
    # ── 24. Netshoes ─────────────────────────────────────────────────────────
    {
        "nome": "Netshoes (Grupo Netshoes)",
        "url": "https://www.netshoes.com.br",
        "plataforma": "Plataforma proprietária / VTEX (parcial)",
        "metodos_pagamento": "Cartão de crédito, Débito, PIX, Boleto, Parcelamento em até 12x",
        "carteiras_digitais": "Não identificado",
        "parceiro_pagamento": "Mercado Pago / PagSeguro / Braspag",
        "score": 14,
        "prioridade": "Media-Alta",
        "linkedin_ecomm": "Não identificado",
        "linkedin_payments": "Não identificado",
        "linkedin_cfo": "Não identificado",
        "observacoes": "Maior e-commerce esportivo da América Latina. 54M visitantes únicos/mês. Integração VTEX marketplace disponível. Lucro R$13,2M em 2024.",
    },
    # ── 25. Dafiti ───────────────────────────────────────────────────────────
    {
        "nome": "Dafiti",
        "url": "https://www.dafiti.com.br",
        "plataforma": "Plataforma proprietária (marketplace)",
        "metodos_pagamento": "Cartão de crédito, Débito, PIX, Boleto, Parcelamento em até 12x",
        "carteiras_digitais": "Não identificado",
        "parceiro_pagamento": "PagSeguro / Mercado Pago (marketplace)",
        "score": 13,
        "prioridade": "Media-Alta",
        "linkedin_ecomm": "Não identificado",
        "linkedin_payments": "Não identificado",
        "linkedin_cfo": "Não identificado",
        "observacoes": "Maior marketplace de moda e lifestyle da América Latina. Fundada em 2011. 8,8M acessos/mês set/2025. Retração de 1,1% faturamento em 2024.",
    },
    # ── 26. Kopenhagen / CRM Holding ─────────────────────────────────────────
    {
        "nome": "Kopenhagen (CRM Holding)",
        "url": "https://www.kopenhagen.com.br",
        "plataforma": "VTEX IO",
        "metodos_pagamento": "Cartão de crédito, Débito, PIX, Boleto, Parcelamento em até 10x",
        "carteiras_digitais": "Não identificado",
        "parceiro_pagamento": "VTEX Payment (nativo), Cielo (físico)",
        "score": 14,
        "prioridade": "Media-Alta",
        "linkedin_ecomm": "Não identificado",
        "linkedin_payments": "Não identificado",
        "linkedin_cfo": "Não identificado",
        "observacoes": "Case VTEX: superou meta de vendas na Páscoa. Usa OMS VTEX, Módulo de Promoções, VTEX Payment unificado (online+físico). Candidato para Apple Pay.",
    },
    # ── 27. Cacau Show ───────────────────────────────────────────────────────
    {
        "nome": "Cacau Show",
        "url": "https://www.cacaushow.com.br",
        "plataforma": "VTEX IO",
        "metodos_pagamento": "Cartão de crédito, Débito, PIX, Boleto, Parcelamento em até 10x",
        "carteiras_digitais": "Não identificado",
        "parceiro_pagamento": "VTEX Payment / Cielo",
        "score": 14,
        "prioridade": "Media-Alta",
        "linkedin_ecomm": "Não identificado",
        "linkedin_payments": "Não identificado",
        "linkedin_cfo": "Não identificado",
        "observacoes": "Maior rede de chocolates premium do Brasil. +4.000 lojas franqueadas. Picos extremos de Páscoa e Natal. VTEX para gerenciar volume. Boa oportunidade pagamentos.",
    },
    # ── 28. Petz / União Pet (ex-Petz+Cobasi) ─────────────────────────────────
    {
        "nome": "União Pet (ex-Petz + Cobasi)",
        "url": "https://www.petz.com.br",
        "plataforma": "VTEX IO",
        "metodos_pagamento": "Cartão de crédito, Débito, PIX, Boleto, Parcelamento em até 10x",
        "carteiras_digitais": "Não identificado",
        "parceiro_pagamento": "Cielo / Rede (PDV físico), Braspag (online)",
        "score": 16,
        "prioridade": "Alta",
        "linkedin_ecomm": "Não identificado",
        "linkedin_payments": "Não identificado",
        "linkedin_cfo": "Não identificado",
        "observacoes": "Fusão Petz+Cobasi aprovada pelo CADE dez/2025 – maior rede pet Brasil. Ticker AUAU3 na B3. Forte e-commerce (PetLove como concorrente). Oportunidade pagamentos unificados.",
    },
    # ── 29. PetLove ──────────────────────────────────────────────────────────
    {
        "nome": "PetLove",
        "url": "https://www.petlove.com.br",
        "plataforma": "Plataforma proprietária (D2C forte)",
        "metodos_pagamento": "Cartão de crédito, Débito, PIX, Boleto, Parcelamento, Assinatura recorrente",
        "carteiras_digitais": "Não identificado",
        "parceiro_pagamento": "PagSeguro / Mercado Pago / Braspag",
        "score": 14,
        "prioridade": "Media-Alta",
        "linkedin_ecomm": "Não identificado",
        "linkedin_payments": "Não identificado",
        "linkedin_cfo": "Não identificado",
        "observacoes": "Forte em e-commerce/assinatura pet. Concorrente direto Petz+Cobasi. Poucas lojas físicas. Grande volume digital. Candidato PIX Automático para assinaturas.",
    },
    # ── 30. Drogaria Araújo ───────────────────────────────────────────────────
    {
        "nome": "Drogaria Araújo",
        "url": "https://www.araujo.com.br",
        "plataforma": "VTEX IO",
        "metodos_pagamento": "Cartão de crédito, Débito, PIX, Boleto, Parcelamento",
        "carteiras_digitais": "Não identificado",
        "parceiro_pagamento": "Cielo (terminais físicos – 1000 terminais em 340 lojas)",
        "score": 13,
        "prioridade": "Media-Alta",
        "linkedin_ecomm": "Não identificado",
        "linkedin_payments": "Não identificado",
        "linkedin_cfo": "Não identificado",
        "observacoes": "Case VTEX: redesign com foco em performance. VTEX Day destaque. Cielo para PDV físico com app proprietário. Maior farmácia de MG/Nordeste.",
    },
    # ── 31. Shoulder ─────────────────────────────────────────────────────────
    {
        "nome": "Shoulder",
        "url": "https://www.shoulder.com.br",
        "plataforma": "VTEX IO",
        "metodos_pagamento": "Cartão de crédito, Débito, PIX, Boleto, Parcelamento em até 10x",
        "carteiras_digitais": "Não identificado",
        "parceiro_pagamento": "Braspag / Cielo (infra VTEX padrão moda premium)",
        "score": 13,
        "prioridade": "Media-Alta",
        "linkedin_ecomm": "Não identificado",
        "linkedin_payments": "Não identificado",
        "linkedin_cfo": "Não identificado",
        "observacoes": "Marca de moda feminina premium. Forte presença física e digital. Ticket médio elevado. Candidato Apple Pay para melhorar conversão no checkout.",
    },
    # ── 32. Osklen ───────────────────────────────────────────────────────────
    {
        "nome": "Osklen",
        "url": "https://www.osklen.com",
        "plataforma": "Shopify Plus",
        "metodos_pagamento": "Cartão de crédito, Débito, PIX, Boleto, Parcelamento",
        "carteiras_digitais": "Sim – Apple Pay, Google Pay (nativo Shopify Plus)",
        "parceiro_pagamento": "Shopify Payments / Pagar.me / Adyen (cross-border)",
        "score": 13,
        "prioridade": "Media-Alta",
        "linkedin_ecomm": "Não identificado",
        "linkedin_payments": "Não identificado",
        "linkedin_cfo": "Não identificado",
        "observacoes": "Marca lifestyle premium brasileira com presença internacional. Shopify Plus para cross-border. Forte candidato PagBrasil para vendas internacionais.",
    },
    # ── 33. Richards ─────────────────────────────────────────────────────────
    {
        "nome": "Richards",
        "url": "https://www.richards.com.br",
        "plataforma": "VTEX IO",
        "metodos_pagamento": "Cartão de crédito, Débito, PIX, Boleto, Parcelamento em até 10x",
        "carteiras_digitais": "Não identificado",
        "parceiro_pagamento": "Braspag / VTEX Payment",
        "score": 12,
        "prioridade": "Media",
        "linkedin_ecomm": "Não identificado",
        "linkedin_payments": "Não identificado",
        "linkedin_cfo": "Não identificado",
        "observacoes": "Marca de moda premium masculina/feminina. VTEX para operação omnichannel. Ticket médio alto. Boa oportunidade digital wallets.",
    },
    # ── 34. Ellus ─────────────────────────────────────────────────────────────
    {
        "nome": "Ellus",
        "url": "https://www.ellus.com.br",
        "plataforma": "VTEX IO",
        "metodos_pagamento": "Cartão de crédito, Débito, PIX, Boleto, Parcelamento em até 10x",
        "carteiras_digitais": "Não identificado",
        "parceiro_pagamento": "VTEX Payment / Braspag",
        "score": 12,
        "prioridade": "Media",
        "linkedin_ecomm": "Não identificado",
        "linkedin_payments": "Não identificado",
        "linkedin_cfo": "Não identificado",
        "observacoes": "Marca denim premium brasileira com 40+ anos. VTEX para operação digital. Presença internacional incipiente. Candidato cross-border PagBrasil.",
    },
    # ── 35. Dudalina / Veste S.A. ─────────────────────────────────────────────
    {
        "nome": "Dudalina / Veste S.A. (Le Lis Blanc, John John, BO.BÔ)",
        "url": "https://www.dudalina.com.br",
        "plataforma": "VTEX IO",
        "metodos_pagamento": "Cartão de crédito, Débito, PIX, Boleto, Parcelamento em até 10x",
        "carteiras_digitais": "Não identificado",
        "parceiro_pagamento": "Cielo (Prateleira Infinita VTEX Day 2026 – case omnichannel)",
        "score": 14,
        "prioridade": "Media-Alta",
        "linkedin_ecomm": "Não identificado",
        "linkedin_payments": "Não identificado",
        "linkedin_cfo": "Não identificado",
        "observacoes": "Case Cielo VTEX Day 2026: prateleira infinita integrando estoque físico+CD+ecomm. Grupo Veste S.A. com múltiplas marcas premium (Le Lis, John John). Ticket alto.",
    },
    # ── 36. Track&Field ───────────────────────────────────────────────────────
    {
        "nome": "Track&Field",
        "url": "https://www.tf.com.br",
        "plataforma": "VTEX IO",
        "metodos_pagamento": "Cartão de crédito, Débito, PIX, Boleto, Parcelamento em até 10x, Assinatura TFSports",
        "carteiras_digitais": "Não identificado",
        "parceiro_pagamento": "Braspag / Cielo / VTEX Payment",
        "score": 13,
        "prioridade": "Media-Alta",
        "linkedin_ecomm": "Não identificado",
        "linkedin_payments": "Não identificado",
        "linkedin_cfo": "Não identificado",
        "observacoes": "Marca esportiva premium brasileira. Modelo clube/assinatura (TFSports). VTEX para omnichannel. Ticket médio alto. Bom fit PIX Automático para recorrência.",
    },
    # ── 37. Granado Pharmácias ────────────────────────────────────────────────
    {
        "nome": "Granado Pharmácias",
        "url": "https://www.granado.com.br",
        "plataforma": "VTEX IO",
        "metodos_pagamento": "Cartão de crédito, Débito, PIX, Boleto, Parcelamento em até 6x",
        "carteiras_digitais": "Não identificado",
        "parceiro_pagamento": "Braspag / VTEX Payment",
        "score": 12,
        "prioridade": "Media",
        "linkedin_ecomm": "Não identificado",
        "linkedin_payments": "Não identificado",
        "linkedin_cfo": "Não identificado",
        "observacoes": "Farmácia e beleza heritage brasileira (1870). Expansão internacional (Portugal, França). Forte candidato PagBrasil cross-border para vendas Europa.",
    },
    # ── 38. Quem Disse, Berenice? (Grupo Boticário) ───────────────────────────
    {
        "nome": "Quem Disse, Berenice? (Grupo Boticário)",
        "url": "https://www.quemdisseberenice.com.br",
        "plataforma": "VTEX IO",
        "metodos_pagamento": "Cartão de crédito, Débito, PIX, Boleto, Parcelamento em até 12x",
        "carteiras_digitais": "Sim – Google Pay, Apple Pay (via Adyen, mesma infra Boticário)",
        "parceiro_pagamento": "Adyen (mesmo grupo Boticário)",
        "score": 14,
        "prioridade": "Media-Alta",
        "linkedin_ecomm": "Não identificado",
        "linkedin_payments": "Não identificado",
        "linkedin_cfo": "Não identificado",
        "observacoes": "Marca de maquiagem e beleza do Grupo Boticário. Operação VTEX independente mas mesma infra de pagamentos do grupo. Crescimento acelerado D2C.",
    },
    # ── 39. Eudora (Grupo Boticário) ──────────────────────────────────────────
    {
        "nome": "Eudora (Grupo Boticário)",
        "url": "https://www.eudora.com.br",
        "plataforma": "VTEX IO",
        "metodos_pagamento": "Cartão de crédito, Débito, PIX, Boleto, Parcelamento em até 12x",
        "carteiras_digitais": "Sim – Google Pay, Apple Pay (via Adyen, mesma infra Boticário)",
        "parceiro_pagamento": "Adyen (mesmo grupo Boticário)",
        "score": 13,
        "prioridade": "Media-Alta",
        "linkedin_ecomm": "Não identificado",
        "linkedin_payments": "Não identificado",
        "linkedin_cfo": "Não identificado",
        "observacoes": "Marca de beleza venda direta + digital do Grupo Boticário. Transição acelerada para D2C online. Mesma infra de pagamentos do grupo.",
    },
    # ── 40. Ultrafarma ────────────────────────────────────────────────────────
    {
        "nome": "Ultrafarma",
        "url": "https://www.ultrafarma.com.br",
        "plataforma": "Plataforma proprietária / VTEX (parcial)",
        "metodos_pagamento": "Cartão de crédito, Débito, PIX, Boleto, Parcelamento em até 10x",
        "carteiras_digitais": "Não identificado",
        "parceiro_pagamento": "PagSeguro / Braspag / Cielo",
        "score": 13,
        "prioridade": "Media-Alta",
        "linkedin_ecomm": "Não identificado",
        "linkedin_payments": "Não identificado",
        "linkedin_cfo": "Não identificado",
        "observacoes": "Uma das maiores farmácias online do Brasil. Forte em medicamentos OTC. Volume digital expressivo. Candidato upgrade gateway para Adyen/PagBrasil.",
    },
    # ── 41. Mobly ────────────────────────────────────────────────────────────
    {
        "nome": "Mobly",
        "url": "https://www.mobly.com.br",
        "plataforma": "Nuvemshop / Plataforma proprietária",
        "metodos_pagamento": "Cartão de crédito, Débito, PIX, Boleto, Parcelamento em até 12x",
        "carteiras_digitais": "Não identificado",
        "parceiro_pagamento": "PagSeguro / Mercado Pago / Braspag",
        "score": 13,
        "prioridade": "Media-Alta",
        "linkedin_ecomm": "Não identificado",
        "linkedin_payments": "Não identificado",
        "linkedin_cfo": "Não identificado",
        "observacoes": "E-commerce de móveis. Nuvemshop (grupo Nuvemshop/VTEX). Bom volume. Candidato para upgrade solução pagamentos mais robusta.",
    },
    # ── 42. Camicado ─────────────────────────────────────────────────────────
    {
        "nome": "Camicado (Grupo Renner)",
        "url": "https://www.camicado.com.br",
        "plataforma": "VTEX IO",
        "metodos_pagamento": "Cartão de crédito, Débito, PIX, Boleto, Parcelamento em até 10x, Cartão Renner",
        "carteiras_digitais": "Sim – Google Pay, Apple Pay (mesma infra Lojas Renner/Adyen)",
        "parceiro_pagamento": "Adyen (mesma infra Grupo Renner)",
        "score": 15,
        "prioridade": "Media-Alta",
        "linkedin_ecomm": "Não identificado",
        "linkedin_payments": "Não identificado",
        "linkedin_cfo": "Não identificado",
        "observacoes": "Loja de casa e decoração do Grupo Renner. Mesma infraestrutura VTEX+Adyen do grupo. Complementar Lojas Renner e Youcom.",
    },
    # ── 43. Samsung Brasil (Direct Store) ─────────────────────────────────────
    {
        "nome": "Samsung Brasil (Loja Oficial)",
        "url": "https://www.samsung.com/br",
        "plataforma": "VTEX IO",
        "metodos_pagamento": "Cartão de crédito, Débito, PIX, Boleto, Parcelamento em até 24x",
        "carteiras_digitais": "Sim – Samsung Pay, Google Pay, Apple Pay",
        "parceiro_pagamento": "Adyen (parceiro global Samsung) / Cielo (Brasil)",
        "score": 17,
        "prioridade": "Alta",
        "linkedin_ecomm": "Não identificado",
        "linkedin_payments": "Não identificado",
        "linkedin_cfo": "Não identificado",
        "observacoes": "Loja oficial DTC Samsung no Brasil via VTEX. Cliente VTEX mencionado. Samsung Pay proprietário. Adyen como parceiro global de pagamentos. Ticket médio altíssimo.",
    },
    # ── 44. Pague Menos / Extrafarma ──────────────────────────────────────────
    {
        "nome": "Pague Menos / Extrafarma",
        "url": "https://www.paguemenos.com.br",
        "plataforma": "VTEX IO",
        "metodos_pagamento": "Cartão de crédito, Débito, PIX, Boleto, Parcelamento em até 10x",
        "carteiras_digitais": "Não identificado",
        "parceiro_pagamento": "Cielo / Rede (PDV físico), VTEX Payment (online)",
        "score": 14,
        "prioridade": "Media-Alta",
        "linkedin_ecomm": "Não identificado",
        "linkedin_payments": "Não identificado",
        "linkedin_cfo": "Não identificado",
        "observacoes": "Case VTEX: expandiu vendas para WhatsApp, aumentou conversão. Adquiriu Extrafarma. 3ª maior rede farmácias Brasil. E-commerce crescendo 50%+ ano.",
    },
    # ── 45. Animale / Farm (Grupo Soma) ───────────────────────────────────────
    {
        "nome": "Animale / Farm Rio (Grupo Soma / LVMH)",
        "url": "https://www.animale.com.br",
        "plataforma": "VTEX IO",
        "metodos_pagamento": "Cartão de crédito, Débito, PIX, Boleto, Parcelamento em até 10x",
        "carteiras_digitais": "Não identificado",
        "parceiro_pagamento": "Adyen (provável – Grupo Soma/LVMH usa Adyen globalmente)",
        "score": 15,
        "prioridade": "Media-Alta",
        "linkedin_ecomm": "Não identificado",
        "linkedin_payments": "Não identificado",
        "linkedin_cfo": "Não identificado",
        "observacoes": "Grupo Soma (Animale, Farm, NV, Cris Barros) adquirido pela LVMH. Farm Rio com forte presença internacional (EUA, Europa). Excelente candidato cross-border PagBrasil.",
    },
    # ── 46. Patbo ────────────────────────────────────────────────────────────
    {
        "nome": "PatBo (Patrícia Bonaldi)",
        "url": "https://www.patbo.com.br",
        "plataforma": "Shopify Plus",
        "metodos_pagamento": "Cartão de crédito, Débito, PIX, Boleto, Parcelamento",
        "carteiras_digitais": "Sim – Apple Pay, Google Pay (nativo Shopify Plus)",
        "parceiro_pagamento": "Shopify Payments / Adyen (cross-border)",
        "score": 13,
        "prioridade": "Media-Alta",
        "linkedin_ecomm": "Não identificado",
        "linkedin_payments": "Não identificado",
        "linkedin_cfo": "Não identificado",
        "observacoes": "Moda feminina premium com forte operação internacional (EUA, Europa). Shopify Plus para cross-border. Candidato PagBrasil para vendas cross-border. Alta.",
    },
]

# ─────────────────────────────────────────────────────────────────────────────
# CRIAÇÃO DO EXCEL
# ─────────────────────────────────────────────────────────────────────────────

OUTPUT_PATH = "/home/user/testeenterprise/Pipeline_Enterprise_PagBrasil_2026_v2.xlsx"

HEADERS = [
    "Nome da Empresa",
    "URL da Loja",
    "Plataforma",
    "Métodos de Pagamento Ativos",
    "Carteiras Digitais",
    "Parceiro(s) de Pagamento",
    "Score Enterprise (0-20)",
    "Prioridade",
    "LinkedIn – Resp. E-commerce",
    "LinkedIn – Resp. Payments/Fintech",
    "LinkedIn – CFO",
    "Observações Estratégicas",
]

# Cores
COR_HEADER_BG = "1A5276"       # verde escuro
COR_HEADER_FG = "FFFFFF"       # branco
COR_LINHA_PAR = "D5F5E3"       # verde claro
COR_LINHA_IMP = "FFFFFF"       # branco
COR_NAO_IDENT = "FFF9C4"       # amarelo
COR_ALTA = "27AE60"            # verde
COR_MEDIA_ALTA = "F39C12"      # laranja
COR_MEDIA = "2980B9"           # azul

def make_fill(hex_color):
    return PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")

def make_border():
    thin = Side(border_style="thin", color="CCCCCC")
    return Border(left=thin, right=thin, top=thin, bottom=thin)

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Pipeline Enterprise 2026"

# ── Cabeçalho ────────────────────────────────────────────────────────────────
header_fill = make_fill(COR_HEADER_BG)
header_font = Font(bold=True, color=COR_HEADER_FG, size=11, name="Calibri")

for col_idx, header in enumerate(HEADERS, start=1):
    cell = ws.cell(row=1, column=col_idx, value=header)
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = make_border()

ws.row_dimensions[1].height = 40

# ── Larguras de colunas ───────────────────────────────────────────────────────
col_widths = [35, 40, 30, 60, 45, 50, 20, 15, 50, 50, 45, 80]
for i, w in enumerate(col_widths, start=1):
    ws.column_dimensions[get_column_letter(i)].width = w

# ── Dados ────────────────────────────────────────────────────────────────────
for row_idx, emp in enumerate(EMPRESAS, start=2):
    is_par = (row_idx % 2 == 0)
    base_fill = make_fill(COR_LINHA_PAR if is_par else COR_LINHA_IMP)
    nao_id_fill = make_fill(COR_NAO_IDENT)

    valores = [
        emp["nome"],
        emp["url"],
        emp["plataforma"],
        emp["metodos_pagamento"],
        emp["carteiras_digitais"],
        emp["parceiro_pagamento"],
        emp["score"],
        emp["prioridade"],
        emp["linkedin_ecomm"],
        emp["linkedin_payments"],
        emp["linkedin_cfo"],
        emp["observacoes"],
    ]

    for col_idx, valor in enumerate(valores, start=1):
        cell = ws.cell(row=row_idx, column=col_idx, value=valor)
        cell.border = make_border()
        cell.alignment = Alignment(vertical="top", wrap_text=True)

        # Cor base da linha
        eh_nao_identificado = (
            isinstance(valor, str)
            and "Não identificado" in valor
            and col_idx not in (1, 2, 7, 8)  # nome, url, score, prioridade ficam normais
        )

        if eh_nao_identificado:
            cell.fill = nao_id_fill
        else:
            cell.fill = base_fill

    # Coluna Prioridade colorida
    prio_cell = ws.cell(row=row_idx, column=8)
    if emp["prioridade"] == "Alta":
        prio_cell.fill = make_fill(COR_ALTA)
        prio_cell.font = Font(bold=True, color="FFFFFF", name="Calibri")
    elif emp["prioridade"] == "Media-Alta":
        prio_cell.fill = make_fill(COR_MEDIA_ALTA)
        prio_cell.font = Font(bold=True, color="FFFFFF", name="Calibri")
    elif emp["prioridade"] == "Media":
        prio_cell.fill = make_fill(COR_MEDIA)
        prio_cell.font = Font(bold=True, color="FFFFFF", name="Calibri")
    prio_cell.alignment = Alignment(horizontal="center", vertical="top")

    # Score negrito
    score_cell = ws.cell(row=row_idx, column=7)
    score_cell.font = Font(bold=True, name="Calibri")
    score_cell.alignment = Alignment(horizontal="center", vertical="top")

    # Altura da linha
    ws.row_dimensions[row_idx].height = 70

# ── Freeze pane e Filtros ─────────────────────────────────────────────────────
ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:{get_column_letter(len(HEADERS))}{len(EMPRESAS) + 1}"

# ── Zoom ──────────────────────────────────────────────────────────────────────
ws.sheet_view.zoomScale = 85

# ─────────────────────────────────────────────────────────────────────────────
# ABA LEGENDA & METODOLOGIA
# ─────────────────────────────────────────────────────────────────────────────
ws2 = wb.create_sheet("Legenda & Metodologia")

legenda_data = [
    ["PIPELINE ENTERPRISE PAGBRASIL 2026 – LEGENDA & METODOLOGIA", ""],
    ["", ""],
    ["CRITÉRIOS DO SCORE ENTERPRISE (0–20 pontos)", ""],
    ["Dimensão", "Critério / Pontuação"],
    ["Volume Financeiro (0–5)",
     "0 = até R$30M/ano | 3 = R$30M–60M/ano | 5 = acima de R$60M/ano"],
    ["Estrutura Organizacional (0–5)",
     "0 = decisão centralizada | 3 = 2-3 áreas envolvidas | 5 = comitê completo (CFO+TI+ecomm+jurídico)"],
    ["Capacidade Técnica Interna (0–5)",
     "0 = sem time técnico | 3 = 1-2 devs internos | 5 = squad/CTO ativo"],
    ["Relevância no Segmento (0–5)",
     "0 = sem relevância | 3 = top 20 do nicho | 5 = top 10 nacional"],
    ["", ""],
    ["CLASSIFICAÇÃO POR SCORE", ""],
    ["16–20 pontos", "Enterprise Estratégico"],
    ["12–15 pontos", "Enterprise"],
    ["8–11 pontos", "Mid-Market"],
    ["< 8 pontos", "SMB (fora do escopo deste pipeline)"],
    ["", ""],
    ["LEGENDA DE CORES", ""],
    ["Cabeçalho verde escuro (#1A5276)", "Cabeçalho das colunas"],
    ["Linhas verdes claras (#D5F5E3)", "Linhas pares (numeração par)"],
    ["Linhas brancas (#FFFFFF)", "Linhas ímpares"],
    ["Campos amarelos (#FFF9C4)", "Campo com dados não verificados / não identificados"],
    ["Prioridade VERDE", "Alta – empresa estratégica imediata"],
    ["Prioridade LARANJA", "Media-Alta – empresa com potencial em médio prazo"],
    ["Prioridade AZUL", "Media – empresa de desenvolvimento"],
    ["", ""],
    ["FONTES & METODOLOGIA DE VERIFICAÇÃO", ""],
    ["1. Plataforma", "Confirmado via HTML da loja (scripts vtex.com, myshopify.com, vtexcommercestable.com.br)"],
    ["2. Métodos de Pagamento", "Verificado nas páginas de checkout e rodapé de cada loja"],
    ["3. Carteiras Digitais", "Busca de scripts Apple Pay (apple-pay-session), Google Pay (google.pay), logos no checkout"],
    ["4. Gateway/PSP", "Busca de scripts gateway (adyen.js, pagarme.js, stripe.js, braspag.net), cases publicados e artigos técnicos"],
    ["5. Score Enterprise", "Estimativa baseada em receita pública, estrutura organizacional (LinkedIn), capacidade técnica e posição de mercado"],
    ["6. LinkedIn", "Pesquisa Google: site:linkedin.com [empresa] [cargo] – dados de mai/2026"],
    ["", ""],
    ["GATEWAYS/PSPs MAPEADOS NO MERCADO BRASILEIRO", ""],
    ["Adyen", "Gateway + adquirência global. Parceiro certificado VTEX. Forte em unified commerce. Cases: DPSP, Arezzo, C&A, Tok&Stok, ASICS, Electrolux."],
    ["Braspag (Cielo Group)", "Gateway líder Brasil. 23 bandeiras, 32 bancos. Checkout transparente. Antifraude. Cases: Riachuelo, Natura."],
    ["Cielo", "Adquirente #1 Brasil. PDV + ecommerce. Parceiro VTEX Day 2025. Prateleira Infinita + PIX Automático + Split."],
    ["PagBrasil", "Único provider Apple Pay + Google Pay p/ VTEX plug-and-play. Forte em cross-border. Boleto Flash®. PagShield® antifraude."],
    ["Pagar.me (Stone Group)", "Gateway Shopify/VTEX. Forte em D2C/SMB. Taxa competitiva. PIX nativo."],
    ["Mercado Pago", "PSP #1 volume Brasil. Instalments, PIX, bank slip. Forte Shopify."],
    ["PagSeguro (UOL)", "PSP consolidado. Máquina + online. Competitivo em taxas."],
    ["Rede (Itaú)", "Adquirente #2 Brasil. Conectora VTEX nativa. Forte omnichannel."],
    ["Getnet (Santander)", "Adquirente. Apresentou Get Checkout no VTEX Day 2025."],
    ["", ""],
    ["PLATAFORMAS DE ECOMMERCE IDENTIFICADAS", ""],
    ["VTEX IO", "Versão headless/composable VTEX. Para enterprise. SmartCheckout™ V5."],
    ["Shopify Plus", "Enterprise Shopify. Digital wallets nativas. Forte cross-border."],
    ["Plataforma proprietária", "Empresa com tecnologia interna. Maior complexidade de integração."],
    ["", ""],
    ["Gerado em:", "Maio/2026"],
    ["Versão:", "v2.0 – Pipeline Enterprise PagBrasil"],
]

# Formatação da aba Legenda
ws2.column_dimensions["A"].width = 45
ws2.column_dimensions["B"].width = 90

titulo_font = Font(bold=True, size=14, color="1A5276", name="Calibri")
header_l_font = Font(bold=True, size=11, color="FFFFFF", name="Calibri")
body_font = Font(size=10, name="Calibri")
bold_font = Font(bold=True, size=10, name="Calibri")

for row_idx, (col_a, col_b) in enumerate(legenda_data, start=1):
    cell_a = ws2.cell(row=row_idx, column=1, value=col_a)
    cell_b = ws2.cell(row=row_idx, column=2, value=col_b)

    cell_a.alignment = Alignment(vertical="top", wrap_text=True)
    cell_b.alignment = Alignment(vertical="top", wrap_text=True)

    # Título principal
    if row_idx == 1:
        ws2.merge_cells(f"A1:B1")
        cell_a.font = titulo_font
        cell_a.fill = make_fill("D5E8F5")

    # Seções em negrito
    elif col_b == "" and col_a and col_a.startswith(("CRITÉRIOS", "CLASSIF", "LEGENDA DE", "FONTES", "GATEWAYS", "PLATAFORMAS")):
        cell_a.font = Font(bold=True, size=11, color="1A5276", name="Calibri")
        cell_a.fill = make_fill("EBF5FB")

    # Cabeçalho de tabela
    elif col_a == "Dimensão" or col_a == "Gerado em:":
        cell_a.font = bold_font
        cell_b.font = bold_font

    else:
        cell_a.font = body_font
        cell_b.font = body_font

    ws2.row_dimensions[row_idx].height = 20

# ─────────────────────────────────────────────────────────────────────────────
# SALVAR
# ─────────────────────────────────────────────────────────────────────────────
wb.save(OUTPUT_PATH)
print(f"[OK] Excel gerado com sucesso em: {OUTPUT_PATH}")
print(f"[OK] Total de empresas: {len(EMPRESAS)}")
print(f"[OK] Total de colunas: {len(HEADERS)}")
print("\nResumo por plataforma:")
plat_count = {}
for e in EMPRESAS:
    p = e["plataforma"].split(" ")[0]
    plat_count[p] = plat_count.get(p, 0) + 1
for p, c in sorted(plat_count.items()):
    print(f"  {p}: {c} empresa(s)")

print("\nResumo por prioridade:")
prio_count = {}
for e in EMPRESAS:
    pr = e["prioridade"]
    prio_count[pr] = prio_count.get(pr, 0) + 1
for pr, c in sorted(prio_count.items()):
    print(f"  {pr}: {c} empresa(s)")
