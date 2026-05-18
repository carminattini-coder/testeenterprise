#!/usr/bin/env python3
"""
Pipeline Enterprise PagBrasil 2026 v3
Gerado a partir de pesquisa profunda realizada em maio de 2026.
"""

import openpyxl
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.filters import AutoFilter
from datetime import datetime

# ─────────────────────────────────────────────
# DADOS DAS EMPRESAS
# ─────────────────────────────────────────────
EMPRESAS = [
    # MODA E VESTUÁRIO PREMIUM
    {
        "nome": "Grupo Soma (Farm Rio, Animale, NV, A.Brand)",
        "url": "https://www.farmrio.com.br",
        "plataforma": "VTEX IO",
        "segmento": "Moda Premium",
        "volume": "R$ 2,0 bi+ /ano (digital 60% receita)",
        "metodos": "Cartão crédito/débito, PIX, Boleto, Parcelamento",
        "carteiras": "Não confirmado",
        "gateway": "A verificar (VTEX nativo + possível Adyen)",
        "checkout": "Transparente",
        "score": 18,
        "prioridade": "Alta",
        "oportunidade": "Grupo de moda premium com múltiplas marcas na VTEX. Digital representa 60% da receita. Sem Apple Pay/Google Pay confirmado — público jovem premium é caso ideal. Possível consolidação de gateway cross-border (Farm Opera nos EUA e Latam). PIX Parcelado para aumentar ticket médio.",
        "linkedin_ecomm": "Alisson Calgaroto (Dir. Tecnologia SOMA) — linkedin.com/in/alissoncalgaroto",
        "linkedin_pay": "A verificar",
        "linkedin_cfo": "A verificar",
        "fonte": "VTEX Blog, AWS Case, Wicomm Case 2025, VTEX Day 2024"
    },
    {
        "nome": "Arezzo&Co (Arezzo, Schutz, Fiever, Carol Bassi)",
        "url": "https://www.arezzo.com.br",
        "plataforma": "VTEX IO",
        "segmento": "Calçados/Moda Premium",
        "volume": "R$ 3,0 bi+ /ano (GMV total grupo)",
        "metodos": "Cartão crédito/débito, PIX, Boleto, Parcelamento",
        "carteiras": "Não identificado",
        "gateway": "Adyen (confirmado pelo Dir. E-commerce)",
        "checkout": "Transparente",
        "score": 19,
        "prioridade": "Alta",
        "oportunidade": "Diretor de e-commerce Pedro Corrêa confirmou uso da Adyen (unified commerce). Grupo com múltiplas marcas premium — Apple Pay/Google Pay não confirmado no checkout. Oportunidade de PIX Parcelado e carteiras digitais. Cross-border (Arezzo USA).",
        "linkedin_ecomm": "Pedro Corrêa de Souza (Dir. E-commerce) — linkedin.com/in/pedrocorreadesouza",
        "linkedin_pay": "A verificar",
        "linkedin_cfo": "A verificar",
        "fonte": "Adyen Case (confirmado), VTEX Blog, E-Commerce Brasil 2024"
    },
    {
        "nome": "Lojas Renner S.A.",
        "url": "https://www.lojarenner.com.br",
        "plataforma": "VTEX IO",
        "segmento": "Moda/Varejo de Moda",
        "volume": "R$ 4,0 bi+ /ano digital estimado",
        "metodos": "Cartão crédito/débito, Cartão Renner, PIX, Boleto, Parcelamento",
        "carteiras": "A verificar",
        "gateway": "A verificar (operação própria + possível Adyen)",
        "checkout": "Transparente",
        "score": 18,
        "prioridade": "Alta",
        "oportunidade": "Uma das maiores varejistas de moda do Brasil. VTEX confirmado. Sem evidência clara de Apple Pay/Google Pay — oportunidade de carteiras digitais. Público feminino jovem premium. Head of E-commerce identificado no LinkedIn.",
        "linkedin_ecomm": "Lucas Neves (Head of E-commerce) — br.linkedin.com/in/lucasneves",
        "linkedin_pay": "A verificar",
        "linkedin_cfo": "A verificar",
        "fonte": "VTEX cases, E-Commerce Brasil, LinkedIn"
    },
    {
        "nome": "Hering (Cia. Hering)",
        "url": "https://www.hering.com.br",
        "plataforma": "VTEX",
        "segmento": "Moda/Vestuário",
        "volume": "R$ 500 M–1 bi /ano digital estimado",
        "metodos": "Cartão crédito/débito, PIX, Boleto, Parcelamento",
        "carteiras": "Não confirmado",
        "gateway": "A verificar",
        "checkout": "Transparente",
        "score": 15,
        "prioridade": "Media-Alta",
        "oportunidade": "Marca centenária com forte e-commerce VTEX. Omnichannel com QR code e geolocalização. Sem Apple Pay confirmado — oportunidade de carteiras digitais. Público jovem que compra via mobile.",
        "linkedin_ecomm": "A verificar",
        "linkedin_pay": "A verificar",
        "linkedin_cfo": "A verificar",
        "fonte": "VTEX cases, agenciaeplus.com.br 2025"
    },
    {
        "nome": "C&A Brasil",
        "url": "https://www.cea.com.br",
        "plataforma": "VTEX",
        "segmento": "Moda/Varejo de Moda",
        "volume": "R$ 2,0 bi+ /ano digital estimado",
        "metodos": "Cartão crédito/débito, Cartão C&A, PIX, Boleto",
        "carteiras": "A verificar",
        "gateway": "Adyen (confirmado — triplicou taxa de autorização digital)",
        "checkout": "Transparente",
        "score": 17,
        "prioridade": "Alta",
        "oportunidade": "C&A usa Adyen como gateway principal (confirmado). Reduziu chargebacks 35% e triplicou autorização digital. Oportunidade de complementar com Apple Pay/Google Pay via PagBrasil no VTEX. PIX Parcelado para ticket médio elevado.",
        "linkedin_ecomm": "A verificar",
        "linkedin_pay": "A verificar",
        "linkedin_cfo": "A verificar",
        "fonte": "Adyen Centro de Conhecimento (confirmado), VTEX Raccoon Case"
    },
    {
        "nome": "Costume (moda premium masculina)",
        "url": "https://www.costume.com.br",
        "plataforma": "VTEX",
        "segmento": "Moda Premium Masculina",
        "volume": "R$ 100–300 M /ano estimado",
        "metodos": "Cartão crédito/débito, PIX, Boleto, Parcelamento",
        "carteiras": "Não confirmado",
        "gateway": "A verificar",
        "checkout": "Transparente",
        "score": 13,
        "prioridade": "Media-Alta",
        "oportunidade": "Loja VTEX confirmada. Moda premium masculina com ticket médio alto. Apple Pay/Google Pay ausente — público de alta renda com iPhone premium. Oportunidade de PIX Parcelado para pedidos de alto valor.",
        "linkedin_ecomm": "A verificar",
        "linkedin_pay": "A verificar",
        "linkedin_cfo": "A verificar",
        "fonte": "E-Commerce Brasil, VTEX Blog (lançamento loja)"
    },
    {
        "nome": "Brooksfield / Brooksfield Donna",
        "url": "https://www.brooksfield.com.br",
        "plataforma": "VTEX",
        "segmento": "Moda Premium",
        "volume": "R$ 100–200 M /ano estimado",
        "metodos": "Cartão crédito/débito, PIX, Boleto, Parcelamento",
        "carteiras": "Não confirmado",
        "gateway": "A verificar",
        "checkout": "Transparente",
        "score": 13,
        "prioridade": "Media-Alta",
        "oportunidade": "Case VTEX: 25% crescimento pedidos e 50% receita em 3 meses. Moda premium executiva — público de alto poder aquisitivo com iPhone. Apple Pay ausente — forte oportunidade mobile premium. PIX Parcelado para alto ticket.",
        "linkedin_ecomm": "A verificar",
        "linkedin_pay": "A verificar",
        "linkedin_cfo": "A verificar",
        "fonte": "VTEX Blog case Brooksfield 2024, agenciaeplus.com.br"
    },

    # CALÇADOS E ESPORTES
    {
        "nome": "Centauro (SBF Group)",
        "url": "https://www.centauro.com.br",
        "plataforma": "VTEX",
        "segmento": "Esportes/Calçados",
        "volume": "R$ 2,0 bi+ /ano digital estimado",
        "metodos": "Cartão crédito/débito, PIX, Boleto, Parcelamento",
        "carteiras": "Não confirmado",
        "gateway": "A verificar",
        "checkout": "Transparente",
        "score": 17,
        "prioridade": "Alta",
        "oportunidade": "Maior loja de esportes do Brasil. VTEX confirmado. Público jovem esportivo — heavy user de iPhone. Apple Pay/Google Pay ausente = grande perda mobile. Gerente E-commerce identificado no LinkedIn. Oportunidade de PIX Parcelado para equipamentos caros.",
        "linkedin_ecomm": "Gustavo Holtz (E-commerce Sr. Manager) — br.linkedin.com/in/gustavo-holtz",
        "linkedin_pay": "A verificar",
        "linkedin_cfo": "A verificar",
        "fonte": "VTEX cases, LinkedIn, agenciaeplus.com.br 2025"
    },
    {
        "nome": "Netshoes (Grupo Netshoes)",
        "url": "https://www.netshoes.com.br",
        "plataforma": "VTEX (marketplace integração)",
        "segmento": "Esportes/Calçados Online",
        "volume": "R$ 1,0 bi+ /ano estimado",
        "metodos": "Cartão crédito/débito, PIX, Boleto, Parcelamento",
        "carteiras": "A verificar",
        "gateway": "A verificar (MercadoPago provável)",
        "checkout": "Transparente",
        "score": 15,
        "prioridade": "Media-Alta",
        "oportunidade": "Pioneiro do e-commerce de esportes no Brasil. 54 milhões visitantes únicos/mês. Integração com VTEX marketplace. Se usar MercadoPago/PagSeguro como gateway principal = forte oportunidade migração para solução premium + Apple Pay.",
        "linkedin_ecomm": "A verificar",
        "linkedin_pay": "A verificar",
        "linkedin_cfo": "A verificar",
        "fonte": "VTEX Help Center, E-Commerce Brasil, Baguete"
    },
    {
        "nome": "Decathlon Brasil",
        "url": "https://www.decathlon.com.br",
        "plataforma": "VTEX",
        "segmento": "Esportes/Equipamentos",
        "volume": "R$ 1,0 bi+ /ano digital estimado",
        "metodos": "Cartão crédito/débito, PIX, Boleto, Parcelamento",
        "carteiras": "Não confirmado",
        "gateway": "A verificar (Adyen provável — marca global)",
        "checkout": "Transparente",
        "score": 17,
        "prioridade": "Alta",
        "oportunidade": "Marca global francesa com forte operação VTEX no Brasil (inStore confirmado). Como marca internacional, Adyen é o gateway natural. Oportunidade de confirmar gateway e oferecer Apple Pay/Google Pay + PIX Parcelado. Cross-border como diferencial.",
        "linkedin_ecomm": "A verificar",
        "linkedin_pay": "A verificar",
        "linkedin_cfo": "A verificar",
        "fonte": "VTEX inStore case, E-Commerce Brasil 2024"
    },

    # BELEZA E COSMÉTICOS
    {
        "nome": "Grupo Boticário (O Boticário, Eudora, Quem Disse Berenice, Beleza na Web)",
        "url": "https://www.boticario.com.br",
        "plataforma": "VTEX",
        "segmento": "Beleza/Cosméticos",
        "volume": "R$ 35,7 bi GMV total (2024) / digital relevante",
        "metodos": "Cartão crédito/débito, PIX, Boleto, Parcelamento",
        "carteiras": "A verificar",
        "gateway": "A verificar (múltiplos possíveis)",
        "checkout": "Transparente",
        "score": 19,
        "prioridade": "Alta",
        "oportunidade": "Maior grupo de beleza do Brasil. GMV R$ 35,7 bi em 2024. Múltiplas marcas em VTEX. Público feminino jovem premium com iPhone — Apple Pay/Google Pay = forte oportunidade. PIX Parcelado para compras de cesta completa de beleza. Possível unificação de gateway.",
        "linkedin_ecomm": "A verificar",
        "linkedin_pay": "A verificar",
        "linkedin_cfo": "A verificar",
        "fonte": "Grupo Boticário press release 2024, VTEX Baguete, E-Commerce Brasil"
    },
    {
        "nome": "Natura &Co Brasil",
        "url": "https://www.natura.com.br",
        "plataforma": "VTEX (confirmado via Raccoon)",
        "segmento": "Beleza/Cosméticos Sustentáveis",
        "volume": "R$ 5,0 bi+ /ano digital estimado",
        "metodos": "Cartão crédito/débito, PIX, Boleto, Parcelamento",
        "carteiras": "A verificar",
        "gateway": "A verificar",
        "checkout": "Transparente",
        "score": 18,
        "prioridade": "Alta",
        "oportunidade": "Marca icônica brasileira com forte e-commerce VTEX (case Raccoon). Presença internacional (Avon, The Body Shop). Cross-border potencial. Público consciente e digital — Apple Pay/Google Pay alinhado ao posicionamento premium. PIX Parcelado.",
        "linkedin_ecomm": "A verificar",
        "linkedin_pay": "A verificar",
        "linkedin_cfo": "A verificar",
        "fonte": "E-Commerce Brasil (Raccoon case), VTEX Brasil"
    },
    {
        "nome": "Sephora Brasil",
        "url": "https://www.sephora.com.br",
        "plataforma": "A verificar (possível VTEX ou proprietária LVMH)",
        "segmento": "Beleza/Luxo",
        "volume": "R$ 500 M–1 bi /ano estimado",
        "metodos": "Cartão crédito/débito, PIX, Boleto",
        "carteiras": "A verificar",
        "gateway": "A verificar (Adyen provável — LVMH global)",
        "checkout": "A verificar",
        "score": 16,
        "prioridade": "Alta",
        "oportunidade": "Marca de luxo global (LVMH) com loja .com.br. Público AB1 com forte uso de iPhone — Apple Pay é esperado mas não confirmado. Cross-border: se usar gateway global Adyen, PagBrasil pode complementar para métodos locais (PIX, Boleto Flash). Alto ticket médio.",
        "linkedin_ecomm": "A verificar",
        "linkedin_pay": "A verificar",
        "linkedin_cfo": "A verificar",
        "fonte": "Mercado geral, LVMH cross-border Brasil"
    },
    {
        "nome": "WePink (influencer-led beauty)",
        "url": "https://www.wepink.com.br",
        "plataforma": "VTEX",
        "segmento": "Beleza/D2C",
        "volume": "R$ 100–300 M /ano estimado",
        "metodos": "Cartão crédito/débito, PIX, Boleto, Parcelamento",
        "carteiras": "Não confirmado",
        "gateway": "A verificar",
        "checkout": "Transparente",
        "score": 14,
        "prioridade": "Media-Alta",
        "oportunidade": "Fenômeno de vendas com transmissões ao vivo (live commerce). VTEX pela escalabilidade. Público jovem feminino 100% mobile. Apple Pay/Google Pay ausente = perda em lives. PIX 1-Click seria diferencial crítico para conversão em live commerce.",
        "linkedin_ecomm": "A verificar",
        "linkedin_pay": "A verificar",
        "linkedin_cfo": "A verificar",
        "fonte": "VTEX Blog (case WePink), agenciaeplus.com.br 2025"
    },

    # FARMÁCIAS E SAÚDE
    {
        "nome": "RaiaDrogasil (Droga Raia + Drogasil + RD Marketplace)",
        "url": "https://www.drogaraia.com.br",
        "plataforma": "VTEX (RD Marketplace + Linx Commerce histórico)",
        "segmento": "Farmácia/Saúde",
        "volume": "R$ 5,0 bi+ /ano digital (125k pedidos/dia)",
        "metodos": "Cartão crédito/débito, PIX, Boleto, Parcelamento",
        "carteiras": "A verificar",
        "gateway": "A verificar (múltiplos)",
        "checkout": "Transparente",
        "score": 18,
        "prioridade": "Alta",
        "oportunidade": "Maior rede de farmácias do Brasil. 125 mil pedidos/dia digital. Marketplace em VTEX (RD Marketplace). Presença no VTEX Day. Se gateway for PagSeguro/MercadoPago = forte oportunidade de migração. Apple Pay para compras rápidas de saúde via mobile.",
        "linkedin_ecomm": "A verificar",
        "linkedin_pay": "A verificar",
        "linkedin_cfo": "A verificar",
        "fonte": "VTEX Blog, Lojistard.com.br, Brazil Journal, Abrafarma"
    },
    {
        "nome": "Ultrafarma",
        "url": "https://www.ultrafarma.com.br",
        "plataforma": "A verificar (provável VTEX ou proprietária)",
        "segmento": "Farmácia/Medicamentos Online",
        "volume": "R$ 500 M–1 bi /ano estimado",
        "metodos": "Cartão crédito/débito, PIX, Boleto, Parcelamento",
        "carteiras": "A verificar",
        "gateway": "A verificar",
        "checkout": "A verificar",
        "score": 14,
        "prioridade": "Media-Alta",
        "oportunidade": "Uma das farmácias online mais respeitadas do Brasil (top pesquisa Abrafarma). Volume alto em medicamentos e saúde. Se checkout for redirecionado = forte oportunidade de checkout transparente. PIX com 1-Click aumentaria conversão em recompras recorrentes.",
        "linkedin_ecomm": "A verificar",
        "linkedin_pay": "A verificar",
        "linkedin_cfo": "A verificar",
        "fonte": "Abrafarma pesquisa, site Ultrafarma"
    },
    {
        "nome": "Panvel Farmácias",
        "url": "https://www.panvel.com",
        "plataforma": "A verificar",
        "segmento": "Farmácia/Beleza",
        "volume": "R$ 300–600 M /ano estimado",
        "metodos": "Cartão crédito/débito, PIX, Boleto",
        "carteiras": "A verificar",
        "gateway": "A verificar",
        "checkout": "A verificar",
        "score": 13,
        "prioridade": "Media-Alta",
        "oportunidade": "Grande rede de farmácias no Sul do Brasil com forte digital. Retail media em crescimento (Abradilan). Sul tem alta penetração de iPhone. Apple Pay/Google Pay seriam diferencial competitivo. PIX Parcelado para compras de alto valor.",
        "linkedin_ecomm": "A verificar",
        "linkedin_pay": "A verificar",
        "linkedin_cfo": "A verificar",
        "fonte": "Abradilan, Abrafarma"
    },

    # CASA, MÓVEIS E DECORAÇÃO
    {
        "nome": "MadeiraMadeira",
        "url": "https://www.madeiramadeira.com.br",
        "plataforma": "VTEX (marketplace + própria)",
        "segmento": "Casa/Móveis/Decoração",
        "volume": "R$ 2,0 bi+ /ano estimado (maior móveis LAm)",
        "metodos": "Cartão crédito/débito, PIX, Boleto, Parcelamento",
        "carteiras": "Não confirmado",
        "gateway": "A verificar",
        "checkout": "Transparente",
        "score": 17,
        "prioridade": "Alta",
        "oportunidade": "Maior e-commerce de móveis e decoração da América Latina. 16 mi acessos/mês. VTEX confirmado (partner portal + marketplace). Alto ticket médio (móveis). PIX Parcelado é diferencial crítico. Apple Pay/Google Pay ausente — oportunidade no mobile. CFOs identificados no LinkedIn.",
        "linkedin_ecomm": "A verificar",
        "linkedin_pay": "Rodrigo Fonsatti — linkedin.com/in/rodrigofonsatti",
        "linkedin_cfo": "Daniel Scandian (fundador/CEO) — br.linkedin.com/in/danielscandian",
        "fonte": "VTEX Partner Portal, VTEX Help Center, E-Commerce Brasil"
    },
    {
        "nome": "Tok&Stok",
        "url": "https://www.tokstok.com.br",
        "plataforma": "VTEX (CMS confirmado — ACCT Global)",
        "segmento": "Casa/Móveis/Decoração",
        "volume": "R$ 500 M–1 bi /ano estimado",
        "metodos": "Cartão crédito/débito, PIX, Boleto, Parcelamento",
        "carteiras": "Não confirmado",
        "gateway": "A verificar",
        "checkout": "Transparente",
        "score": 15,
        "prioridade": "Media-Alta",
        "oportunidade": "Referência em móveis e decoração premium. VTEX CMS confirmado. Crescimento 23,9% visitas online. Alto ticket médio. PIX Parcelado para móveis caros. Apple Pay/Google Pay ausente — público AB1 com iPhone.",
        "linkedin_ecomm": "A verificar",
        "linkedin_pay": "A verificar",
        "linkedin_cfo": "A verificar",
        "fonte": "ACCT Global case (VTEX CMS), E-Commerce Brasil casa & decoração"
    },
    {
        "nome": "Leroy Merlin Brasil",
        "url": "https://www.leroymerlin.com.br",
        "plataforma": "VTEX (confirmado — Raccoon case + VTEX Day 2024)",
        "segmento": "Casa/Construção/Decoração",
        "volume": "R$ 2,0 bi+ /ano digital estimado",
        "metodos": "Cartão crédito/débito, PIX, Boleto, Parcelamento",
        "carteiras": "A verificar",
        "gateway": "Adyen (provável — grupo Adeo global)",
        "checkout": "Transparente",
        "score": 17,
        "prioridade": "Alta",
        "oportunidade": "Líder em Search Share do segmento casa (21%). VTEX confirmado. Grupo global Adeo — gateway pode ser Adyen globalmente mas não confirmado localmente. Marketplace ativo. PIX Parcelado para obras e reformas (alto ticket). Oportunidade de Apple Pay.",
        "linkedin_ecomm": "A verificar",
        "linkedin_pay": "A verificar",
        "linkedin_cfo": "A verificar",
        "fonte": "E-Commerce Brasil, Raccoon case, VTEX Day 2024"
    },
    {
        "nome": "Mobly",
        "url": "https://www.mobly.com.br",
        "plataforma": "A verificar (própria/VTEX)",
        "segmento": "Casa/Móveis Online",
        "volume": "R$ 500 M–1 bi /ano estimado",
        "metodos": "Cartão crédito/débito, PIX, Boleto, Parcelamento",
        "carteiras": "Não confirmado",
        "gateway": "A verificar",
        "checkout": "A verificar",
        "score": 14,
        "prioridade": "Media-Alta",
        "oportunidade": "3º lugar em e-commerce de casa & móveis. Volume relevante. Alto ticket médio (móveis). PIX Parcelado para pedidos acima de R$500. Se redirecionamento no checkout = oportunidade de checkout transparente. Apple Pay ausente.",
        "linkedin_ecomm": "A verificar",
        "linkedin_pay": "A verificar",
        "linkedin_cfo": "A verificar",
        "fonte": "E-Commerce Brasil casa & decoração ranking"
    },

    # ELETRÔNICOS E TECNOLOGIA
    {
        "nome": "Samsung Brasil",
        "url": "https://www.samsung.com/br",
        "plataforma": "VTEX headless (app nativo React Native + VTEX APIs)",
        "segmento": "Eletrônicos/Tecnologia",
        "volume": "R$ 3,0 bi+ /ano digital estimado",
        "metodos": "Cartão crédito/débito, PIX, Boleto, Parcelamento, Samsung Pay",
        "carteiras": "Samsung Pay (confirmado). Apple Pay/Google Pay: A verificar",
        "gateway": "Adyen (provável — Samsung global) + VTEX nativo",
        "checkout": "Transparente",
        "score": 18,
        "prioridade": "Alta",
        "oportunidade": "App nativo com VTEX headless — case de referência. Samsung Pay confirmado. Adyen provável globalmente. Oportunidade de PIX Parcelado para eletrodomésticos de alto valor (TV, geladeira). Complementar gateway para maximizar conversão iOS (Apple Pay).",
        "linkedin_ecomm": "A verificar",
        "linkedin_pay": "A verificar",
        "linkedin_cfo": "A verificar",
        "fonte": "VTEX Blog headless case, Quality Digital case, VTEX Day 2024"
    },
    {
        "nome": "Electrolux Brasil",
        "url": "https://www.electrolux.com.br",
        "plataforma": "VTEX (confirmado — VTEX Day 2024, 7 painéis)",
        "segmento": "Eletrodomésticos/Tecnologia",
        "volume": "R$ 1,0 bi+ /ano digital estimado",
        "metodos": "Cartão crédito/débito, PIX, Boleto, Parcelamento",
        "carteiras": "Não confirmado",
        "gateway": "A verificar",
        "checkout": "Transparente",
        "score": 16,
        "prioridade": "Alta",
        "oportunidade": "Marca global com VTEX confirmado. D2C em digitalização (Gerente Sênior D2C no VTEX Day). Alto ticket médio. PIX Parcelado para eletrodomésticos premium. Apple Pay/Google Pay ausente. Possível cross-border para unificar gateway global.",
        "linkedin_ecomm": "A verificar",
        "linkedin_pay": "A verificar",
        "linkedin_cfo": "A verificar",
        "fonte": "VTEX Day 2024 (Electrolux em 7 painéis), VTEX cases"
    },
    {
        "nome": "Whirlpool Brasil (Brastemp, Consul, KitchenAid)",
        "url": "https://www.brastemp.com.br",
        "plataforma": "VTEX (confirmado — VP Vendas no VTEX Day 2024)",
        "segmento": "Eletrodomésticos/B2B+B2C",
        "volume": "R$ 1,0 bi+ /ano digital estimado",
        "metodos": "Cartão crédito/débito, PIX, Boleto, Parcelamento",
        "carteiras": "Não confirmado",
        "gateway": "A verificar",
        "checkout": "Transparente",
        "score": 16,
        "prioridade": "Alta",
        "oportunidade": "Whirlpool confirmou VTEX. Múltiplas marcas (Brastemp, Consul, KitchenAid). B2B em destaque no VTEX Day 2024. Oportunidade de solução completa: gateway + PIX Parcelado + Apple Pay para D2C premium KitchenAid. Cross-border potencial.",
        "linkedin_ecomm": "A verificar",
        "linkedin_pay": "A verificar",
        "linkedin_cfo": "A verificar",
        "fonte": "VTEX Day 2024 (Eduardo Akira VP Vendas), VTEX cases"
    },

    # ALIMENTAÇÃO GOURMET E BEBIDAS PREMIUM
    {
        "nome": "Wine.com.br (maior e-comm de vinhos do Brasil)",
        "url": "https://www.wine.com.br",
        "plataforma": "A verificar (provável VTEX ou proprietária)",
        "segmento": "Bebidas Premium/Assinatura",
        "volume": "R$ 300–600 M /ano estimado",
        "metodos": "Cartão crédito/débito, PIX, Boleto, Parcelamento",
        "carteiras": "Não confirmado",
        "gateway": "A verificar",
        "checkout": "A verificar",
        "score": 14,
        "prioridade": "Media-Alta",
        "oportunidade": "Maior e-commerce de vinhos + clube de assinatura do Brasil. VTEX Day 2024 teve sommelier pop-up. Público AB1 premium — Apple Pay/Google Pay alinhados. PIX Parcelado para cestas de alto valor. Modelo de assinatura = recorrência de pagamento.",
        "linkedin_ecomm": "A verificar",
        "linkedin_pay": "A verificar",
        "linkedin_cfo": "A verificar",
        "fonte": "Wine.com.br institucional, VTEX Day 2024 sommelier"
    },
    {
        "nome": "Grand Cru / Evino (Grupo Víssimo)",
        "url": "https://www.grandcru.com.br",
        "plataforma": "A verificar",
        "segmento": "Bebidas Premium/Luxo",
        "volume": "R$ 200–400 M /ano estimado",
        "metodos": "Cartão crédito/débito, PIX, Boleto, Parcelamento",
        "carteiras": "Não confirmado",
        "gateway": "Adyen (confirmado — Grupo Víssimo usa Revenue Accelerate + unified commerce Adyen)",
        "checkout": "A verificar",
        "score": 15,
        "prioridade": "Media-Alta",
        "oportunidade": "Maior grupo de vinhos importados do Brasil. Adyen CONFIRMADO (Revenue Accelerate). Público AB1 extremo. Apple Pay/Google Pay seria natural para este público. Oportunidade de complementar Adyen com solução PagBrasil para métodos locais (PIX 1-Click, Boleto Flash).",
        "linkedin_ecomm": "A verificar",
        "linkedin_pay": "A verificar",
        "linkedin_cfo": "A verificar",
        "fonte": "Adyen Centro de Conhecimento (confirmado Grupo Víssimo)"
    },

    # PET PREMIUM
    {
        "nome": "Petlove",
        "url": "https://www.petlove.com.br",
        "plataforma": "Proprietária (e-commerce próprio desde 1999)",
        "segmento": "Pet/Assinatura",
        "volume": "R$ 700 M–1,5 bi /ano estimado",
        "metodos": "Cartão crédito/débito, PIX, Boleto, Parcelamento",
        "carteiras": "A verificar",
        "gateway": "A verificar",
        "checkout": "Transparente",
        "score": 16,
        "prioridade": "Alta",
        "oportunidade": "Maior e-commerce pet do Brasil. Ticket médio R$220. Modelo assinatura. Tecnologia própria com CTO ativo. Apple Pay/Google Pay ausente provável. Recorrência = oportunidade de gateway premium + PIX Parcelado. Squad técnico próprio facilita integração.",
        "linkedin_ecomm": "A verificar",
        "linkedin_pay": "A verificar",
        "linkedin_cfo": "Alexander Borges (CTO) — linkedin.com/in/alexander-borges-0970909",
        "fonte": "Petlove.com.br, E-Commerce Brasil pet 2024, Neotrust Q1 2024"
    },
    {
        "nome": "Cobasi (agora Grupo PetCenter Cobasi)",
        "url": "https://www.cobasi.com.br",
        "plataforma": "VTEX (confirmado desde 2015 — case omnichannel)",
        "segmento": "Pet/Varejo Especializado",
        "volume": "R$ 1,5 bi+ /ano estimado (fusão Petz)",
        "metodos": "Cartão crédito/débito, PIX, Boleto, Assinatura, Parcelamento",
        "carteiras": "Não confirmado",
        "gateway": "A verificar",
        "checkout": "Transparente",
        "score": 16,
        "prioridade": "Alta",
        "oportunidade": "VTEX desde 2015. Assinatura = 40% das compras online. Fusão com Petz (R$7bi combinados). Apple Pay/Google Pay ausente = oportunidade pós-fusão para modernizar checkout. PIX Parcelado para rações premium. Estrutura omnichannel consolida volume.",
        "linkedin_ecomm": "A verificar",
        "linkedin_pay": "A verificar",
        "linkedin_cfo": "A verificar",
        "fonte": "VTEX Blog omnichannel case Cobasi, E-Commerce Brasil pet fusão"
    },

    # JOIAS E LUXO
    {
        "nome": "Vivara",
        "url": "https://www.vivara.com.br",
        "plataforma": "VTEX (vtex.vivara.com.br confirmado)",
        "segmento": "Joias/Luxo",
        "volume": "R$ 1,0 bi+ /ano estimado (275+ lojas + digital 40% receita)",
        "metodos": "Cartão crédito/débito, PIX (Adyen), Boleto, Parcelamento",
        "carteiras": "A verificar (Apple Pay não confirmado publicamente)",
        "gateway": "Adyen (CONFIRMADO — parceria desde 2014, PIX integrado, omnichannel)",
        "checkout": "Transparente",
        "score": 18,
        "prioridade": "Alta",
        "oportunidade": "Vivara usa Adyen (CONFIRMADO). Digital = 40% receita. Omnichannel 140% mais rápido com Adyen. VTEX confirmado. Já tem PIX via Adyen. Oportunidade: Apple Pay/Google Pay e PIX Parcelado (alto ticket joias). Complementar Adyen com PagBrasil para métodos exclusivos.",
        "linkedin_ecomm": "Pedro Rondon — br.linkedin.com/in/pedrorondon",
        "linkedin_pay": "Caroline Pinheiro de Souza (Gerente Executiva) — br.linkedin.com/in/caroline-pinheiro-de-souza-3430034a",
        "linkedin_cfo": "Leonardo Bichara — br.linkedin.com/in/leonardobichara",
        "fonte": "Adyen press release, VTEX (vtex.vivara.com.br), SEGS, E-Commerce Brasil"
    },
    {
        "nome": "H.Stern",
        "url": "https://www.hstern.com.br",
        "plataforma": "A verificar (provável VTEX ou SAP hybris)",
        "segmento": "Joias/Luxo Internacional",
        "volume": "R$ 500 M–1 bi /ano estimado",
        "metodos": "Cartão crédito/débito, PIX, Boleto, Parcelamento",
        "carteiras": "Não confirmado",
        "gateway": "A verificar",
        "checkout": "A verificar",
        "score": 16,
        "prioridade": "Alta",
        "oportunidade": "Joalheria brasileira de prestígio internacional (lojas em NY, Paris, etc.). Cross-border nativo. Público ultra premium — Apple Pay é esperado mas não confirmado. Checkout possivelmente legado. Alto ticket = PIX Parcelado diferencial. Oportunidade cross-border premium.",
        "linkedin_ecomm": "A verificar",
        "linkedin_pay": "A verificar",
        "linkedin_cfo": "A verificar",
        "fonte": "H.Stern site, mercado geral joias Brasil"
    },

    # CROSS-BORDER / MARCAS INTERNACIONAIS
    {
        "nome": "Nike Brasil (Nike Direct)",
        "url": "https://www.nike.com.br",
        "plataforma": "Proprietária (Nike Direct Technology)",
        "segmento": "Esportes/Calçados Premium Cross-border",
        "volume": "R$ 1,0 bi+ /ano estimado",
        "metodos": "Cartão crédito/débito, PIX, Boleto, Parcelamento",
        "carteiras": "A verificar (Apple Pay provável — Nike global usa)",
        "gateway": "Adyen (provável — Nike global usa Adyen)",
        "checkout": "Transparente",
        "score": 17,
        "prioridade": "Alta",
        "oportunidade": "Nike Direct com tecnologia própria. Gateway global provavelmente Adyen. Oportunidade: garantir PIX Parcelado local com sub-adquirente (PagBrasil). Apple Pay/Google Pay: confirmar status no .com.br. Cross-border diferencial.",
        "linkedin_ecomm": "A verificar",
        "linkedin_pay": "A verificar",
        "linkedin_cfo": "A verificar",
        "fonte": "Nike.com.br, mercado cross-border brasil 2024"
    },
    {
        "nome": "Adidas Brasil",
        "url": "https://www.adidas.com.br",
        "plataforma": "VTEX (provável) ou SAP Commerce",
        "segmento": "Esportes/Moda Premium Cross-border",
        "volume": "R$ 800 M–1,5 bi /ano estimado",
        "metodos": "Cartão crédito/débito, PIX, Boleto, Parcelamento",
        "carteiras": "A verificar",
        "gateway": "Adyen (provável — Adidas global usa Adyen)",
        "checkout": "Transparente",
        "score": 16,
        "prioridade": "Alta",
        "oportunidade": "Marca global com forte digital no Brasil. Gateway global = Adyen. Oportunidade PagBrasil: PIX Parcelado, Boleto Flash e Apple Pay para público jovem premium brasileiro. Cross-border: unificar experiência de pagamento LatAm.",
        "linkedin_ecomm": "A verificar",
        "linkedin_pay": "A verificar",
        "linkedin_cfo": "A verificar",
        "fonte": "Mercado cross-border, comparativo plataformas brasil"
    },
    {
        "nome": "L'Oréal Brasil (Lancôme, L'Oréal Paris, Kiehl's, NYX)",
        "url": "https://www.loreal.com.br",
        "plataforma": "VTEX (confirmado — VTEX client list)",
        "segmento": "Beleza Premium Cross-border",
        "volume": "R$ 1,0 bi+ /ano digital estimado",
        "metodos": "Cartão crédito/débito, PIX, Boleto, Parcelamento",
        "carteiras": "Não confirmado",
        "gateway": "Adyen (provável — L'Oréal global usa Adyen)",
        "checkout": "Transparente",
        "score": 17,
        "prioridade": "Alta",
        "oportunidade": "L'Oréal Group em VTEX confirmado. Múltiplas marcas premium. Público feminino AB1. Apple Pay/Google Pay não confirmado no .com.br — oportunidade com PagBrasil. PIX Parcelado para sets premium. Consolidação de gateway cross-border.",
        "linkedin_ecomm": "A verificar",
        "linkedin_pay": "A verificar",
        "linkedin_cfo": "A verificar",
        "fonte": "VTEX client list (L'Oréal), VTEX Day casos internacionais"
    },
    {
        "nome": "Havaianas / Alpargatas Brasil",
        "url": "https://www.havaianas.com.br",
        "plataforma": "Shopify Plus (confirmado — Shopify blog BR)",
        "segmento": "Calçados/Lifestyle Cross-border",
        "volume": "R$ 500 M–1 bi /ano estimado",
        "metodos": "Cartão crédito/débito, PIX, Boleto, Parcelamento",
        "carteiras": "Não confirmado",
        "gateway": "A verificar (Shopify Payments ou terceiro)",
        "checkout": "Transparente",
        "score": 15,
        "prioridade": "Media-Alta",
        "oportunidade": "Marca icônica brasileira em Shopify Plus. Cross-border forte (EUA, Europa). Shopify Plus = oportunidade para PagBrasil como sub-adquirente para métodos locais (PIX, Boleto Flash). Apple Pay/Google Pay provavelmente disponível via Shopify Payments mas sem PIX Parcelado local.",
        "linkedin_ecomm": "A verificar",
        "linkedin_pay": "A verificar",
        "linkedin_cfo": "A verificar",
        "fonte": "Shopify BR blog (Havaianas mencionado), mercado cross-border"
    },
    {
        "nome": "Levi's Brasil",
        "url": "https://www.levi.com.br",
        "plataforma": "Shopify Plus (provável — Levi's global usa Shopify)",
        "segmento": "Moda/Denim Premium Cross-border",
        "volume": "R$ 200–500 M /ano estimado",
        "metodos": "Cartão crédito/débito, PIX, Boleto",
        "carteiras": "A verificar",
        "gateway": "A verificar",
        "checkout": "A verificar",
        "score": 14,
        "prioridade": "Media-Alta",
        "oportunidade": "Marca global de denim premium. Shopify Plus global = desafio para métodos locais (PIX, Boleto). Oportunidade PagBrasil: integração Shopify para oferecer PIX Parcelado e Boleto Flash. Apple Pay provavelmente disponível mas sem parcelamento.",
        "linkedin_ecomm": "A verificar",
        "linkedin_pay": "A verificar",
        "linkedin_cfo": "A verificar",
        "fonte": "Mercado cross-border, Shopify global"
    },

    # OUTROS SEGMENTOS RELEVANTES
    {
        "nome": "Amaro (moda feminina D2C)",
        "url": "https://www.amaro.com",
        "plataforma": "VTEX (confirmado — case omnichannel AR/IA)",
        "segmento": "Moda Feminina Premium D2C",
        "volume": "R$ 200–500 M /ano estimado",
        "metodos": "Cartão crédito/débito, PIX, Boleto, Parcelamento",
        "carteiras": "Não confirmado",
        "gateway": "A verificar",
        "checkout": "Transparente",
        "score": 14,
        "prioridade": "Media-Alta",
        "oportunidade": "Pioneira D2C com realidade aumentada e IA na VTEX. Público jovem feminino digital-native. Apple Pay/Google Pay seria natural. PIX 1-Click para conversão mobile. Marca valoriza tecnologia = receptiva a soluções premium.",
        "linkedin_ecomm": "A verificar",
        "linkedin_pay": "A verificar",
        "linkedin_cfo": "A verificar",
        "fonte": "VTEX cases (Amaro AR/IA), agenciaeplus.com.br 2025"
    },
    {
        "nome": "Drogaria São Paulo (Grupo DPSP)",
        "url": "https://www.drogariasaopaulo.com.br",
        "plataforma": "VTEX (confirmado — agenciaeplus case)",
        "segmento": "Farmácia/Saúde",
        "volume": "R$ 500 M–1 bi /ano estimado",
        "metodos": "Cartão crédito/débito, PIX, Boleto, Parcelamento",
        "carteiras": "A verificar",
        "gateway": "A verificar",
        "checkout": "Transparente",
        "score": 15,
        "prioridade": "Media-Alta",
        "oportunidade": "VTEX confirmado (implantação via agenciaeplus). Grupo DPSP com Pacheco+Drogaria SP. Volume digital expressivo. Apple Pay ausente para compras rápidas de saúde. PIX Parcelado para dermocosméticos premium. Oportunidade de migração se gateway for básico.",
        "linkedin_ecomm": "A verificar",
        "linkedin_pay": "A verificar",
        "linkedin_cfo": "A verificar",
        "fonte": "agenciaeplus.com.br case, Baguete"
    },
    {
        "nome": "Carrefour Brasil (Digital)",
        "url": "https://www.carrefour.com.br",
        "plataforma": "VTEX (confirmado — case 2020 e VTEX Day)",
        "segmento": "Supermercado/Varejo",
        "volume": "R$ 5,0 bi+ /ano digital estimado",
        "metodos": "Cartão crédito/débito, PIX, Boleto, Parcelamento, Cartão Carrefour",
        "carteiras": "A verificar",
        "gateway": "A verificar (múltiplos)",
        "checkout": "Transparente",
        "score": 17,
        "prioridade": "Alta",
        "oportunidade": "Carrefour VTEX confirmado (case histórico + VTEX Day 2024). Volume massivo. Marketplace ativo. Apple Pay/Google Pay para compras no supermercado via mobile. PIX como método de alta preferência para alimentação. Mercado alimentar cresce no digital.",
        "linkedin_ecomm": "A verificar",
        "linkedin_pay": "A verificar",
        "linkedin_cfo": "A verificar",
        "fonte": "Baguete (Carrefour VTEX 2020), VTEX Day, E-Commerce Brasil"
    },
    {
        "nome": "Granado Pharmácias",
        "url": "https://www.granado.com.br",
        "plataforma": "A verificar",
        "segmento": "Farmácia/Beleza Premium Heritage",
        "volume": "R$ 100–300 M /ano estimado",
        "metodos": "Cartão crédito/débito, PIX, Boleto, Parcelamento",
        "carteiras": "Não confirmado",
        "gateway": "A verificar",
        "checkout": "A verificar",
        "score": 12,
        "prioridade": "Media",
        "oportunidade": "Marca heritage premium brasileira + expansão internacional (Portugal, EUA). Cross-border natural. Público AB1 que valoriza qualidade. Apple Pay para experiência premium consistente. PIX 1-Click para aumentar conversão. Checkout possivelmente desatualizado.",
        "linkedin_ecomm": "A verificar",
        "linkedin_pay": "A verificar",
        "linkedin_cfo": "A verificar",
        "fonte": "Site Granado, mercado beleza premium Brasil"
    },
    {
        "nome": "Imaginarium / Loja do Mecanismo",
        "url": "https://www.imaginarium.com.br",
        "plataforma": "VTEX (provável)",
        "segmento": "Presentes/Lifestyle Premium",
        "volume": "R$ 100–200 M /ano estimado",
        "metodos": "Cartão crédito/débito, PIX, Boleto, Parcelamento",
        "carteiras": "Não confirmado",
        "gateway": "A verificar",
        "checkout": "A verificar",
        "score": 10,
        "prioridade": "Media",
        "oportunidade": "Marca de presentes premium com e-commerce relevante. Público jovem AB. Apple Pay para compras de presentes via mobile. PIX para conversão rápida em datas comemorativas. Verificar plataforma atual.",
        "linkedin_ecomm": "A verificar",
        "linkedin_pay": "A verificar",
        "linkedin_cfo": "A verificar",
        "fonte": "Mercado presentes, conhecimento geral"
    },
    {
        "nome": "DIA Supermercado Brasil",
        "url": "https://www.dia.com.br",
        "plataforma": "VTEX (CONFIRMADO — case 100% aumento conversão 2020)",
        "segmento": "Supermercado/Alimentação",
        "volume": "R$ 300–600 M /ano digital estimado",
        "metodos": "Cartão crédito/débito, PIX, Boleto, Parcelamento",
        "carteiras": "Não confirmado",
        "gateway": "A verificar",
        "checkout": "Transparente",
        "score": 14,
        "prioridade": "Media-Alta",
        "oportunidade": "VTEX confirmado com OMS + pick and pack. 100% aumento conversão. 50% mais clientes online. Apple Pay/Google Pay ausente = perda em compras rápidas de supermercado. PIX natural para alimentação. Se gateway for básico = oportunidade de upgrade.",
        "linkedin_ecomm": "A verificar",
        "linkedin_pay": "A verificar",
        "linkedin_cfo": "A verificar",
        "fonte": "VTEX Blog case DIA Supermercado, agenciaeplus.com.br"
    },
    {
        "nome": "Pague Menos / Extrafarma",
        "url": "https://www.paguemenos.com.br",
        "plataforma": "A verificar",
        "segmento": "Farmácia/Saúde",
        "volume": "R$ 500 M–1 bi /ano digital estimado",
        "metodos": "Cartão crédito/débito, PIX, Boleto, Parcelamento",
        "carteiras": "A verificar",
        "gateway": "A verificar",
        "checkout": "A verificar",
        "score": 13,
        "prioridade": "Media-Alta",
        "oportunidade": "Grande rede de farmácias com operação digital crescente. Mencionada em estudos Abrafarma sobre e-commerces de farmácias respeitados. Oportunidade de checkout premium + PIX Parcelado para dermocosméticos. Apple Pay para compras rápidas.",
        "linkedin_ecomm": "A verificar",
        "linkedin_pay": "A verificar",
        "linkedin_cfo": "A verificar",
        "fonte": "Abrafarma pesquisa, VTEX Day (mencionada)"
    },
    {
        "nome": "Abanu (cosméticos veganos D2C)",
        "url": "https://www.abanu.com.br",
        "plataforma": "Shopify (confirmado — case Shopify BR 2024)",
        "segmento": "Beleza/Cosméticos Veganos D2C",
        "volume": "R$ 30–80 M /ano estimado",
        "metodos": "Cartão crédito/débito, PIX, Boleto, Parcelamento",
        "carteiras": "Não confirmado",
        "gateway": "Shopify Payments (provável)",
        "checkout": "Transparente",
        "score": 9,
        "prioridade": "Media",
        "oportunidade": "Case Shopify Brasil 2024 — crescimento constante mês a mês. Público vegano premium jovem digital. Apple Pay natural para este público. PIX Parcelado para kits de beleza. Menor volume mas crescimento acelerado — Mid-Market com potencial enterprise.",
        "linkedin_ecomm": "A verificar",
        "linkedin_pay": "A verificar",
        "linkedin_cfo": "A verificar",
        "fonte": "Shopify Brasil blog case 2024"
    },
    {
        "nome": "LK Sneakers",
        "url": "https://www.lksneakers.com.br",
        "plataforma": "Shopify (confirmado — case Shopify BR 2024)",
        "segmento": "Calçados/Sneakers Premium",
        "volume": "R$ 30–100 M /ano estimado (15x crescimento em 2024)",
        "metodos": "Cartão crédito/débito, PIX, Boleto, Parcelamento",
        "carteiras": "Não confirmado",
        "gateway": "Shopify Payments (provável)",
        "checkout": "Transparente",
        "score": 9,
        "prioridade": "Media",
        "oportunidade": "15x crescimento em 2024 via Shopify. Sneakers premium com público jovem AB. Apple Pay esperado para este perfil. PIX Parcelado para sneakers premium (R$800–2.000). Crescimento acelerado indica necessidade de gateway mais robusto.",
        "linkedin_ecomm": "A verificar",
        "linkedin_pay": "A verificar",
        "linkedin_cfo": "A verificar",
        "fonte": "Shopify Brasil blog case 2024"
    },
]

# ─────────────────────────────────────────────
# CORES E ESTILOS
# ─────────────────────────────────────────────
COR_HEADER_FUNDO = "1A5276"
COR_HEADER_FONTE = "FFFFFF"
COR_PAR = "D5F5E3"
COR_IMPAR = "FFFFFF"
COR_AVERIGUAR_FUNDO = "FFF9C4"
COR_AVERIGUAR_FONTE = "7D6608"
COR_OPORTUNIDADE = "EBF5FB"
COR_ALTA = "1E8449"
COR_MEDIA_ALTA = "F39C12"
COR_MEDIA = "2E86C1"

PRIORIDADE_CORES = {
    "Alta": (COR_ALTA, "FFFFFF"),
    "Media-Alta": (COR_MEDIA_ALTA, "FFFFFF"),
    "Media": (COR_MEDIA, "FFFFFF"),
}

COLUNAS = [
    ("Nome da Empresa", 30),
    ("URL da Loja", 38),
    ("Plataforma", 18),
    ("Segmento", 18),
    ("Volume Digital Estimado", 22),
    ("Métodos de Pagamento", 45),
    ("Carteiras Digitais (Apple Pay / Google Pay)", 28),
    ("Gateway / Parceiro de Pagamento", 30),
    ("Checkout", 16),
    ("Score Enterprise (0–20)", 12),
    ("Prioridade", 14),
    ("Oportunidade PagBrasil", 60),
    ("LinkedIn – E-commerce", 55),
    ("LinkedIn – Payments", 55),
    ("LinkedIn – CFO / Financeiro", 55),
    ("Fonte / Observações", 45),
]

NOMES_COLUNAS = [c[0] for c in COLUNAS]
LARGURAS_COLUNAS = [c[1] for c in COLUNAS]

AVERIGUAR_KEYWORDS = [
    "a verificar", "não confirmado", "provável", "possível",
    "estimado", "estimada"
]


def contem_averiguar(texto: str) -> bool:
    t = texto.lower()
    return any(kw in t for kw in AVERIGUAR_KEYWORDS)


def fill(hex_color: str) -> PatternFill:
    return PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")


def thin_border() -> Border:
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)


def criar_excel():
    wb = openpyxl.Workbook()

    # ── ABA 1: PIPELINE ────────────────────────────────────────────────────
    ws = wb.active
    ws.title = "Pipeline Enterprise"
    ws.freeze_panes = "A2"

    # Cabeçalho
    header_fill = fill(COR_HEADER_FUNDO)
    header_font = Font(name="Arial", bold=True, color=COR_HEADER_FONTE, size=11)
    header_align = Alignment(horizontal="center", vertical="center",
                              wrap_text=True)

    for col_idx, (nome_col, _) in enumerate(COLUNAS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=nome_col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_align
        cell.border = thin_border()

    ws.row_dimensions[1].height = 45

    # Linhas de dados
    for row_idx, empresa in enumerate(EMPRESAS, start=2):
        is_par = (row_idx % 2 == 0)
        cor_base = COR_PAR if is_par else COR_IMPAR

        valores = [
            empresa["nome"],
            empresa["url"],
            empresa["plataforma"],
            empresa["segmento"],
            empresa["volume"],
            empresa["metodos"],
            empresa["carteiras"],
            empresa["gateway"],
            empresa["checkout"],
            empresa["score"],
            empresa["prioridade"],
            empresa["oportunidade"],
            empresa["linkedin_ecomm"],
            empresa["linkedin_pay"],
            empresa["linkedin_cfo"],
            empresa["fonte"],
        ]

        for col_idx, valor in enumerate(valores, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=valor)
            cell.border = thin_border()

            col_nome = NOMES_COLUNAS[col_idx - 1]
            txt = str(valor) if valor is not None else ""

            # Alinhamento
            if col_nome in ("Score Enterprise (0–20)",):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_nome == "Oportunidade PagBrasil":
                cell.alignment = Alignment(horizontal="left", vertical="top",
                                           wrap_text=True)
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center",
                                           wrap_text=True)

            # Cor de fundo
            if col_nome == "Prioridade":
                cores = PRIORIDADE_CORES.get(txt, (cor_base, "000000"))
                cell.fill = fill(cores[0])
                cell.font = Font(name="Arial", color=cores[1], bold=True, size=10)
            elif col_nome == "Oportunidade PagBrasil":
                cell.fill = fill(COR_OPORTUNIDADE)
                cell.font = Font(name="Arial", size=10)
            elif contem_averiguar(txt):
                cell.fill = fill(COR_AVERIGUAR_FUNDO)
                cell.font = Font(name="Arial", color=COR_AVERIGUAR_FONTE, size=10,
                                 italic=True)
            else:
                cell.fill = fill(cor_base)
                cell.font = Font(name="Arial", size=10)

        ws.row_dimensions[row_idx].height = 60

    # Larguras de colunas
    for col_idx, largura in enumerate(LARGURAS_COLUNAS, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = largura

    # Filtro automático
    ultima_col = get_column_letter(len(COLUNAS))
    ultima_linha = len(EMPRESAS) + 1
    ws.auto_filter.ref = f"A1:{ultima_col}{ultima_linha}"

    # ── ABA 2: LEGENDA & METODOLOGIA ──────────────────────────────────────
    ws2 = wb.create_sheet("Legenda & Metodologia")

    legenda_header = Font(name="Arial", bold=True, color="FFFFFF", size=12)
    legenda_titulo_fill = fill("1A5276")
    legenda_subtit_fill = fill("2874A6")
    legenda_subtit_font = Font(name="Arial", bold=True, size=11)

    linhas_legenda = [
        ("LEGENDA & METODOLOGIA — PIPELINE ENTERPRISE PAGBRASIL 2026 v3", "titulo"),
        (f"Data de geração: {datetime.now().strftime('%d/%m/%Y %H:%M')}", "info"),
        ("Pesquisa realizada: maio de 2026 | Fontes: VTEX Blog, Adyen, Shopify, E-Commerce Brasil, LinkedIn, Webshoppers", "info"),
        ("", ""),
        ("CRITÉRIOS DE SCORE ENTERPRISE (0–20 pontos)", "subtitulo"),
        ("Volume Financeiro (0–5 pontos):", "campo"),
        ("  0 pts = Até R$30M/ano de volume digital estimado", "detalhe"),
        ("  3 pts = R$30M–R$60M/ano de volume digital estimado", "detalhe"),
        ("  5 pts = Acima de R$60M/ano de volume digital estimado", "detalhe"),
        ("Estrutura Organizacional (0–5 pontos):", "campo"),
        ("  0 pts = Decisão centralizada em 1 pessoa", "detalhe"),
        ("  3 pts = 2–3 áreas envolvidas (ecomm + TI ou financeiro)", "detalhe"),
        ("  5 pts = Comitê completo: CFO + TI + E-commerce + Jurídico", "detalhe"),
        ("Capacidade Técnica (0–5 pontos):", "campo"),
        ("  0 pts = Sem time interno de tech", "detalhe"),
        ("  3 pts = 1–2 desenvolvedores internos", "detalhe"),
        ("  5 pts = Squad/CTO ativo com autonomia técnica", "detalhe"),
        ("Relevância no Segmento (0–5 pontos):", "campo"),
        ("  0 pts = Sem relevância de mercado", "detalhe"),
        ("  3 pts = Top 20 do nicho", "detalhe"),
        ("  5 pts = Top 10 nacional no segmento", "detalhe"),
        ("", ""),
        ("CLASSIFICAÇÃO POR SCORE:", "subtitulo"),
        ("  16–20 = Enterprise Estratégico (Alta Prioridade)", "detalhe"),
        ("  12–15 = Enterprise (Media-Alta Prioridade)", "detalhe"),
        ("  8–11  = Mid-Market (Media Prioridade)", "detalhe"),
        ("", ""),
        ("CORES DO PIPELINE:", "subtitulo"),
        ("  VERDE ESCURO (#1E8449) = Alta Prioridade (Score 16+)", "detalhe"),
        ("  LARANJA (#F39C12)     = Media-Alta Prioridade (Score 12–15)", "detalhe"),
        ("  AZUL (#2E86C1)        = Media Prioridade (Score 8–11)", "detalhe"),
        ("  AZUL CLARO (#EBF5FB)  = Coluna Oportunidade PagBrasil (destaque)", "detalhe"),
        ("  AMARELO (#FFF9C4)     = Campo 'A verificar' / 'Não confirmado' (ação necessária)", "detalhe"),
        ("  VERDE CLARO (#D5F5E3) = Linhas pares (legibilidade)", "detalhe"),
        ("", ""),
        ("FONTES DE DADOS:", "subtitulo"),
        ("• VTEX Blog (vtex.com/pt-br/blog e vtex.com/pt-br/casos-de-clientes)", "fonte_item"),
        ("• Adyen Centro de Conhecimento (adyen.com/pt_BR)", "fonte_item"),
        ("• E-Commerce Brasil (ecommercebrasil.com.br)", "fonte_item"),
        ("• LinkedIn (br.linkedin.com)", "fonte_item"),
        ("• Shopify Brasil (shopify.com/br)", "fonte_item"),
        ("• Abrafarma, Abradilan, ABComm", "fonte_item"),
        ("• Webshoppers 51ª Edição (Conversion)", "fonte_item"),
        ("• VTEX Day 2024 (cobertura jornalística)", "fonte_item"),
        ("• Comunidade E-Commerce (comunidadeecommerce.com)", "fonte_item"),
        ("• agenciaeplus.com.br (15 lojas VTEX + 23 cases)", "fonte_item"),
        ("• AWS Partner Success Stories", "fonte_item"),
        ("", ""),
        ("METODOLOGIA DE VERIFICAÇÃO DE PLATAFORMA:", "subtitulo"),
        ("Tentativa de WebFetch direto nas lojas para identificar scripts VTEX/Shopify.", "info"),
        ("Muitos sites retornaram HTTP 403 (bloqueio de crawlers).", "info"),
        ("Nestes casos, plataforma foi inferida via: cases publicados, VTEX Partner Portal,", "info"),
        ("subdomínios (vtex.vivara.com.br), anúncios de lançamento e fontes secundárias.", "info"),
        ("Campos 'A verificar' requerem análise manual do HTML da loja.", "info"),
        ("", ""),
        ("OBSERVAÇÕES IMPORTANTES PARA O TIME COMERCIAL:", "subtitulo"),
        ("1. Empresas com Adyen já como gateway NÃO são concorrência — são oportunidade COMPLEMENTAR.", "detalhe"),
        ("   PagBrasil complementa Adyen com: PIX Parcelado, Boleto Flash e Apple Pay plug-and-play.", "detalhe"),
        ("2. Empresas sem gateway identificado = oportunidade de prospecção direta e urgente.", "detalhe"),
        ("3. Apple Pay via PagBrasil é a ÚNICA solução plug-and-play sem dev para VTEX (agosto 2025).", "detalhe"),
        ("4. Foco em público jovem premium = Apple Pay; foco em ticket alto = PIX Parcelado.", "detalhe"),
        ("5. Cross-border (marcas internacionais com .com.br): oportunidade de unificar gateway.", "detalhe"),
    ]

    row = 1
    for texto, tipo in linhas_legenda:
        cell = ws2.cell(row=row, column=1, value=texto)
        if tipo == "titulo":
            cell.fill = legenda_titulo_fill
            cell.font = legenda_header
            cell.alignment = Alignment(horizontal="left", vertical="center")
        elif tipo == "subtitulo":
            cell.fill = legenda_subtit_fill
            cell.font = legenda_subtit_font
        elif tipo == "campo":
            cell.font = Font(name="Arial", bold=True, size=10)
        elif tipo == "detalhe":
            cell.font = Font(name="Arial", size=10)
        elif tipo == "fonte_item":
            cell.font = Font(name="Arial", size=10, color="1A5276")
        elif tipo == "info":
            cell.font = Font(name="Arial", size=10, italic=True, color="555555")

        row += 1

    ws2.column_dimensions["A"].width = 100
    ws2.sheet_view.showGridLines = False

    # Salvar
    output_path = "/home/user/testeenterprise/Pipeline_Enterprise_PagBrasil_2026_v3.xlsx"
    wb.save(output_path)
    print(f"\n✅ Excel gerado com sucesso: {output_path}")
    print(f"   Total de empresas: {len(EMPRESAS)}")

    # Estatísticas
    alta = sum(1 for e in EMPRESAS if e["prioridade"] == "Alta")
    media_alta = sum(1 for e in EMPRESAS if e["prioridade"] == "Media-Alta")
    media = sum(1 for e in EMPRESAS if e["prioridade"] == "Media")
    vtex = sum(1 for e in EMPRESAS if "VTEX" in e["plataforma"])
    shopify = sum(1 for e in EMPRESAS if "Shopify" in e["plataforma"])
    prop = sum(1 for e in EMPRESAS if "Proprietária" in e["plataforma"] or "proprietária" in e["plataforma"])
    averiguar = sum(1 for e in EMPRESAS if "A verificar" in e["plataforma"])

    print(f"\n📊 PRIORIDADES:")
    print(f"   Alta:       {alta} empresas")
    print(f"   Media-Alta: {media_alta} empresas")
    print(f"   Media:      {media} empresas")

    print(f"\n🖥️ PLATAFORMAS:")
    print(f"   VTEX:        {vtex} empresas")
    print(f"   Shopify:     {shopify} empresas")
    print(f"   Proprietária:{prop} empresa(s)")
    print(f"   A verificar: {averiguar} empresa(s)")

    print(f"\n🏆 TOP OPORTUNIDADES APPLE PAY / CARTEIRAS DIGITAIS:")
    sem_apple = [e for e in EMPRESAS
                 if e["prioridade"] == "Alta"
                 and ("não confirmado" in e["carteiras"].lower()
                      or "a verificar" in e["carteiras"].lower())]
    for e in sem_apple:
        print(f"   - {e['nome']} ({e['segmento']}, Score {e['score']})")

    print(f"\n⚠️ GATEWAYS BÁSICOS (MERCADOPAGO/PAGSEGURO) IDENTIFICADOS:")
    gw_basico = [e for e in EMPRESAS
                 if "mercadopago" in e["gateway"].lower()
                 or "pagseguro" in e["gateway"].lower()]
    for e in gw_basico:
        print(f"   - {e['nome']} — Gateway: {e['gateway']}")
    if not gw_basico:
        print("   (Nenhum confirmado — campo A verificar em muitas empresas)")
        print("   Empresas a verificar: Netshoes, Ultrafarma, Mobly, Pague Menos")

    return output_path


if __name__ == "__main__":
    criar_excel()
