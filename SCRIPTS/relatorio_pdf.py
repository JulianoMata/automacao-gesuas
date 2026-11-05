# -*- coding: utf-8 -*-
"""
========================================================================================
SISTEMA DE AUTOMAÇÃO DE ATENDIMENTOS E ATUALIZAÇÃO - GESUAS
MÓDULO: RELATÓRIO OFICIAL DE AUDITORIA EM PDF (relatorio_pdf.py)
VERSÃO: 2.1 (Outubro/2026)
AUTORIA: Juliano França da Mata
FORMAÇÃO: Bacharel em Ciência de Dados e Inteligência Artificial | Tecnólogo em Gestão de TI
CARGO: Agente Administrativo / Operador do Cadastro Único - CRAS Sul (Campo Belo/MG)
DESCRIÇÃO: Constrói dinamicamente o documento PDF paisagem via ReportLab contendo
            cabeçalho executivo, legenda cromática compacta no topo, totalizadores
            e tabela de auditoria com estilos tipográficos específicos (Visita em
            itálico tachado e Unipessoal em negrito).
========================================================================================
"""

import os
from datetime import datetime
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from SCRIPTS.config import PASTA_PDF, AUTOR_NOME, AUTOR_CARGO, AUTOR_QUALIF


def formatar_observacao_estilizada(obs_texto, estilo_paragrafo):
    """Aplica formatação HTML para o ReportLab:
    - Visita: itálico tachado (<i><strike>Visita</strike></i>)
    - Unipessoal: negrito (<b>Unipessoal</b>)"""
    if not obs_texto:
        return ""

    txt = str(obs_texto).strip()

    # 1. Combinação Unipessoal - Visita
    if "Unipessoal - Visita" in txt:
        txt = txt.replace("Unipessoal - Visita", "<b>Unipessoal</b> - <i><strike>Visita</strike></i>")
    # 2. Unipessoal - CRAS
    elif "Unipessoal - CRAS" in txt:
        txt = txt.replace("Unipessoal - CRAS", "<b>Unipessoal - CRAS</b>")
    # 3. Qualquer menção isolada a Unipessoal
    elif "Unipessoal" in txt:
        txt = txt.replace("Unipessoal", "<b>Unipessoal</b>")

    # 4. Visita isolada (para famílias que não são unipessoais)
    if "Visita" in txt and "<strike>" not in txt:
        txt = txt.replace("Visita", "<i><strike>Visita</strike></i>")

    return Paragraph(txt, estilo_paragrafo)


def gerar_relatorio_pdf(relatorio_dados, nome_aba):
    if not relatorio_dados:
        print("[PDF] Nenhum registro processado para emissão de relatório.")
        return

    timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    caminho_pdf = os.path.join(PASTA_PDF, f"Relatorio_GESUAS_{nome_aba}_{timestamp}.pdf")

    doc = SimpleDocTemplate(
        caminho_pdf,
        pagesize=landscape(letter),
        rightMargin=25,
        leftMargin=25,
        topMargin=25,
        bottomMargin=25
    )
    styles = getSampleStyleSheet()

    subtitulo_style = ParagraphStyle(
        'SubtituloDoc',
        parent=styles['Normal'],
        fontSize=8.5,
        textColor=colors.HexColor("#222222"),
        spaceAfter=6,
        leading=12
    )

    obs_cell_style = ParagraphStyle(
        'ObsCellStyle',
        parent=styles['Normal'],
        fontSize=7.5,
        leading=9.5,
        textColor=colors.HexColor("#111111")
    )

    elementos = []

    # Totalizadores por Ação
    total_atualizacoes = sum(1 for r in relatorio_dados if r.get("TipoAcao") == "ATUALIZAÇÃO")
    total_inclusoes = sum(1 for r in relatorio_dados if r.get("TipoAcao") == "INCLUSÃO")
    total_consultas = sum(1 for r in relatorio_dados if r.get("TipoAcao") == "CONSULTA")
    total_exclusoes = sum(1 for r in relatorio_dados if r.get("TipoAcao") == "EXCLUSÃO")
    total_geral = len(relatorio_dados)

    titulo = Paragraph(
        "<b>Relatório Operacional de Atendimentos CadÚnico</b> — CRAS Sul (Campo Belo/MG)",
        styles['Title']
    )

    subtitulo = Paragraph(
        f"<b>Responsável Técnico:</b> {AUTOR_NOME} ({AUTOR_QUALIF}) — {AUTOR_CARGO}<br/>"
        f"<b>Emissão:</b> {datetime.now().strftime('%d/%m/%Y às %H:%M:%S')} &nbsp;|&nbsp; <b>Mês de Referência:</b> {nome_aba}<br/>"
        f"<b>Resumo Geral:</b> Atualizações: <b>{total_atualizacoes}</b> &nbsp;|&nbsp; "
        f"Inclusões: <b>{total_inclusoes}</b> &nbsp;|&nbsp; "
        f"Consultas: <b>{total_consultas}</b> &nbsp;|&nbsp; "
        f"Exclusões: <b>{total_exclusoes}</b> &nbsp;|&nbsp; "
        f"<b>Total de Atendimentos:</b> <b>{total_geral}</b>",
        subtitulo_style
    )

    elementos.append(titulo)
    elementos.append(subtitulo)
    elementos.append(Spacer(1, 4))

    # -------------------------------------------------------------
    # LEGENDA DISCRETA (NO INÍCIO, LOGO ANTES DA TABELA)
    # -------------------------------------------------------------
    legenda_dados = [[
        "Legenda:",
        "Atualização",
        "Consulta",
        "Inclusão",
        "Exclusão",
        "Intervenção / Erro"
    ]]

    t_legenda = Table(legenda_dados, colWidths=[55, 110, 100, 100, 100, 130])
    t_legenda.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 7),
        ('ALIGN', (0, 0), (0, 0), 'LEFT'),
        ('ALIGN', (1, 0), (-1, -1), 'CENTER'),
        ('TEXTCOLOR', (0, 0), (0, 0), colors.HexColor("#555555")),
        ('TEXTCOLOR', (1, 0), (-1, -1), colors.HexColor("#222222")),
        ('BACKGROUND', (1, 0), (1, 0), colors.HexColor("#D4EDDA")),  # Verde
        ('BACKGROUND', (2, 0), (2, 0), colors.HexColor("#FFF2CC")),  # Amarelo
        ('BACKGROUND', (3, 0), (3, 0), colors.HexColor("#F8D7DA")),  # Vermelho
        ('BACKGROUND', (4, 0), (4, 0), colors.HexColor("#D1ECF1")),  # Azul
        ('BACKGROUND', (5, 0), (5, 0), colors.HexColor("#FFE0B2")),  # Laranja
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ('GRID', (1, 0), (-1, -1), 0.5, colors.HexColor("#CCCCCC")),
    ]))
    elementos.append(t_legenda)
    elementos.append(Spacer(1, 6))

    # -------------------------------------------------------------
    # TABELA PRINCIPAL (7 COLUNAS = 740 PT <= 742 PT ÚTIL)
    # -------------------------------------------------------------
    cabecalho: list[str | Paragraph] = ["ID", "Usuário(a)", "CPF", "Data / Hora", "Ação", "Status GESUAS", "Observação"]
    dados_tabela: list[list[str | Paragraph]] = [cabecalho]

    estilo_tabela = [
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#0078D7")),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 7.5),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 4),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.lightgrey),
    ]

    for idx, r in enumerate(relatorio_dados, start=1):
        st = str(r.get("Status", "")).upper()
        tipo_acao = str(r.get("TipoAcao", "ATUALIZAÇÃO")).upper()
        obs_raw = str(r.get("Obs", "")).strip()

        # Normalização dos Status
        if "CONSULTA" in st:
            status_exibicao = "REGISTRADO"
        elif "SUCESSO" in st:
            status_exibicao = "SUCESSO"
        elif "CADASTRADO" in st or "DUPLICIDADE" in st:
            status_exibicao = "JÁ CADASTRADO"
        elif "ERRO" in st or "NÃO ENCONTRADO" in st:
            status_exibicao = "ERRO"
        elif "DESLIGADA" in st:
            status_exibicao = "DESLIGADA"
        elif "EXCLUS" in st:
            status_exibicao = "PENDENTE"
        else:
            status_exibicao = st[:12]

        celula_obs = formatar_observacao_estilizada(obs_raw, obs_cell_style)

        dados_tabela.append([
            str(r.get("ID", "")),
            str(r.get("Nome", ""))[:24],
            str(r.get("CPF", "")),
            f"{r.get('Data', '')} {r.get('Hora', '')}",
            tipo_acao,
            status_exibicao,
            celula_obs
        ])

        # Colorização condicional das linhas
        if "ERRO" in st or "NÃO ENCONTRADO" in st or "DESLIGADA" in st:
            estilo_tabela.append(('BACKGROUND', (0, idx), (-1, idx), colors.HexColor("#FFE0B2")))  # Laranja
        elif tipo_acao == "ATUALIZAÇÃO":
            estilo_tabela.append(('BACKGROUND', (0, idx), (-1, idx), colors.HexColor("#D4EDDA")))  # Verde
        elif tipo_acao == "CONSULTA":
            estilo_tabela.append(('BACKGROUND', (0, idx), (-1, idx), colors.HexColor("#FFF2CC")))  # Amarelo
        elif tipo_acao == "INCLUSÃO":
            estilo_tabela.append(('BACKGROUND', (0, idx), (-1, idx), colors.HexColor("#F8D7DA")))  # Vermelho
        elif tipo_acao == "EXCLUSÃO":
            estilo_tabela.append(('BACKGROUND', (0, idx), (-1, idx), colors.HexColor("#D1ECF1")))  # Azul

    t = Table(dados_tabela, colWidths=[30, 175, 80, 90, 80, 85, 200])
    t.setStyle(TableStyle(estilo_tabela))
    elementos.append(t)
    doc.build(elementos)
    print(f"\n[OK] Relatório PDF gerado com sucesso em:\n{caminho_pdf}")
    