# -*- coding: utf-8 -*-
"""
========================================================================================
SISTEMA DE AUTOMAÇÃO DE ATENDIMENTOS E ATUALIZAÇÃO - GESUAS
MÓDULO: PONTO DE ENTRADA E ORQUESTRADOR PRINCIPAL (main.py)
VERSÃO: 2.1 (Outubro/2026)
AUTORIA: Juliano França da Mata
FORMAÇÃO: Bacharel em Ciência de Dados e Inteligência Artificial | Tecnólogo em Gestão de TI
CARGO: Agente Administrativo / Operador do Cadastro Único - CRAS Sul (Campo Belo/MG)
DESCRIÇÃO: Script principal de orquestração em lote. Realiza a leitura em modo somente
            leitura das planilhas em DADOS_BRUTOS, gerencia seleção de ano/mês,
            captura interrupção via ESC, coordena as chamadas ao bot do GESUAS
            e assegura a consolidação completa de atendimentos em relatórios PDF.
========================================================================================
"""

import os
import re
import sys
import time
import unicodedata
from datetime import datetime
import pandas as pd
from playwright.sync_api import sync_playwright

# Módulo nativo do Windows para captura de teclado não-bloqueante
try:
    import msvcrt
except ImportError:
    msvcrt = None

from SCRIPTS.config import (
    PORTA_DEBUG, MESES_PT, PASTA_DADOS_BRUTOS,
    listar_planilhas_disponiveis, exibir_cabecalho_terminal
)
from SCRIPTS.notificacoes import alerta_intervencao_manual
from SCRIPTS.relatorio_pdf import gerar_relatorio_pdf
from SCRIPTS.gesuas_bot import (
    iniciar_edge, conectar_pagina_gesuas, buscar_municipe,
    verificar_pessoa_desligada_ou_acessar_familia, checar_duplicidade_atendimento,
    registrar_atendimento_e_encerrar
)


def verificar_tecla_esc():
    """Verifica se a tecla ESC foi pressionada sem travar o loop."""
    if msvcrt and msvcrt.kbhit():
        tecla = msvcrt.getch()
        if tecla == b'\x1b':  # Código ASCII do ESC
            return True
    return False


def normalizar_texto(texto):
    if not texto or str(texto).lower() in ['nan', 'none', 'nat', '']:
        return ""
    return unicodedata.normalize('NFKD', str(texto)).encode('ASCII', 'ignore').decode('utf-8').strip().upper()


def classificar_modalidade_obs(obs_raw):
    """Normaliza o texto da coluna OBS para capturar variações de Unipessoal e Visita."""
    txt = normalizar_texto(obs_raw)
    if not txt:
        return "Família (Atend. CRAS)"

    eh_unipessoal = bool(re.search(r'\bUNI\b|UNIPESSOAL|UNI\s*PESSOAL', txt))
    eh_visita = bool(re.search(r'\bVD\b|VISIT|DOMICILI|ENTREVISTA', txt))

    if eh_unipessoal and eh_visita:
        return "Unipessoal (Visita Domiciliar)"
    elif eh_unipessoal:
        return "Unipessoal (CRAS)"
    elif eh_visita:
        return "Família (Visita Domiciliar)"
    else:
        return "Família (Atend. CRAS)"


def classificar_tipo_acao(servico_raw):
    """Classifica com precisão o Tipo de Ação a partir da coluna SERVIÇO."""
    txt = normalizar_texto(servico_raw)
    if "INCLUSA" in txt:
        return "INCLUSÃO"
    elif "CONSULTA" in txt:
        return "CONSULTA"
    elif "EXCLUSA" in txt:
        return "EXCLUSÃO"
    else:
        return "ATUALIZAÇÃO"


def selecionar_planilha():
    """Permite selecionar dinamicamente a planilha anual dentro de DADOS_BRUTOS."""
    planilhas = listar_planilhas_disponiveis()
    if not planilhas:
        alerta_intervencao_manual(
            "Ficheiro Não Encontrado",
            f"Nenhuma planilha .xlsx foi localizada na pasta:\n{PASTA_DADOS_BRUTOS}"
        )
        return None

    if len(planilhas) == 1:
        return planilhas[0]

    print("\n" + "=" * 65)
    print(" SELEÇÃO DA PLANILHA DE TRABALHO (ANO) ")
    print("=" * 65)
    for idx, p in enumerate(planilhas, start=1):
        nome_arq = os.path.basename(p)
        print(f"  [{idx}] {nome_arq}")
    print("  [0] Sair do Sistema")
    print("=" * 65)

    while True:
        escolha = input(f">> Escolha a planilha (0 a {len(planilhas)}): ").strip()
        if escolha == "0":
            return "SAIR"
        if escolha.isdigit() and 1 <= int(escolha) <= len(planilhas):
            return planilhas[int(escolha) - 1]
        print("Opção inválida. Tente novamente.")


def selecionar_aba_mes(caminho_planilha):
    """Lê as abas disponíveis na planilha selecionada e solicita a escolha."""
    excel_obj = pd.ExcelFile(caminho_planilha)
    abas = excel_obj.sheet_names
    mes_atual_pt = MESES_PT.get(datetime.now().month, "").upper()

    print("\n" + "=" * 65)
    print(f" SELEÇÃO DO MÊS — [{os.path.basename(caminho_planilha)}] ")
    print("=" * 65)
    for idx, a in enumerate(abas, start=1):
        marcador = " (Mês Atual)" if str(a).strip().upper() == mes_atual_pt else ""
        print(f"  [{idx}] {a}{marcador}")
    print("  [V] Voltar para seleção de Planilha")
    print("  [0] Sair do Sistema")
    print("=" * 65)

    while True:
        escolha = input(f">> Digite o número do mês ou prima ENTER para a primeira aba: ").strip().upper()
        if escolha == "0":
            return "SAIR"
        if escolha == "V":
            return "VOLTAR"
        if escolha == "":
            return abas[0]
        if escolha.isdigit() and 1 <= int(escolha) <= len(abas):
            return abas[int(escolha) - 1]
        for a in abas:
            if str(a).strip().upper() == escolha:
                return a
        print("Opção inválida. Tente novamente.")


def processar_lote(caminho_planilha, aba_selecionada):
    """Executa o processamento automatizado da aba do mês."""
    print(f"\n[A CARREGAR] Lendo aba: {aba_selecionada}...")
    df = pd.read_excel(caminho_planilha, sheet_name=aba_selecionada)

    if not iniciar_edge():
        return

    print("\n" + "-" * 70)
    print("O Microsoft Edge foi iniciado na porta CDP.")
    print("1. Certifique-se de estar conectado ao GESUAS com sua conta.")
    print("2. Deixe o navegador no Painel principal.")
    print("[DICA]: Pressione [ESC] no terminal durante a execução para interromper.")
    print("-" * 70)
    input(">> Pressione ENTER no terminal assim que estiver no Painel do GESUAS...")

    relatorio = []

    try:
        with sync_playwright() as p:
            try:
                browser = p.chromium.connect_over_cdp(f"http://localhost:{PORTA_DEBUG}")
                page = conectar_pagina_gesuas(browser)
            except Exception as e:
                alerta_intervencao_manual(
                    "Falha de Comunicação",
                    f"Não foi possível estabelecer conexão com o Edge.\nErro: {e}"
                )
                return

            total_linhas = len(df)
            for index, row in df.iterrows():
                # Verificação de interrupção com a tecla ESC
                if verificar_tecla_esc():
                    print("\n\n" + "!" * 80)
                    print(" [ESC DETECTADO] EXECUÇÃO PAUSADA PELO OPERADOR!")
                    print("!" * 80)
                    opcao_parada = input(">> Deseja encerrar este lote agora e gerar relatório do que foi feito? [S/N]: ").strip().upper()
                    if opcao_parada == "S":
                        print("Encerrando lote antecipadamente...")
                        break
                    else:
                        print("Continuando processamento...")

                servico_raw = row.get('SERVICO', '')
                tipo_acao = classificar_tipo_acao(servico_raw)
                status_gesuas = str(row.get('GESUAS', '')).strip()

                nome = str(row.get('NOME', '')).strip()
                cpf = str(row.get('CPF', '')).strip()
                obs_linha = row.get('OBS', '')
                modalidade = classificar_modalidade_obs(obs_linha)

                val_data = row.get('DATA', '')
                val_hora = row.get('HORAS', '')

                # Validação de preenchimento de Data/Hora
                if pd.isna(val_data) or str(val_data).strip().lower() in ['nat', 'nan', 'none', '']:
                    print(f"   [PULANDO] Linha ID {row.get('ID')} ({nome}): Sem DATA na planilha.")
                    continue

                if pd.isna(val_hora) or str(val_hora).strip().lower() in ['nat', 'nan', 'none', '']:
                    print(f"   [PULANDO] Linha ID {row.get('ID')} ({nome}): Sem HORA na planilha.")
                    continue

                data_atend = val_data.strftime('%d/%m/%Y') if (isinstance(val_data, datetime) and not pd.isna(val_data)) else str(val_data).strip()
                hora_atend = val_hora.strftime('%H:%M') if (isinstance(val_hora, datetime) and not pd.isna(val_hora)) else str(val_hora).strip()
                if len(hora_atend) > 5:
                    hora_atend = hora_atend[:5]

                data_hora_str = f"{data_atend} {hora_atend}"

                # CENÁRIO 1: REGISTROS COM GESUAS 'OK' NA PLANILHA
                if status_gesuas != "" and status_gesuas.lower() not in ["nan", "none"]:
                    print(f"\n[{index + 1}/{total_linhas}] Linha ID {row.get('ID')}: {nome} | [JÁ CADASTRADO - PLANILHA GESUAS '{status_gesuas}']")
                    relatorio.append({
                        "ID": row.get('ID'),
                        "Nome": nome,
                        "CPF": cpf,
                        "Data": data_atend,
                        "Hora": hora_atend,
                        "TipoAcao": tipo_acao,
                        "Modalidade": modalidade,
                        "Status": "JÁ CADASTRADO",
                        "Obs": f"{tipo_acao} já cadastrada (Planilha: {status_gesuas})"
                    })
                    continue

                # CENÁRIO 2: REGISTRO DE 'CONSULTA'
                if tipo_acao == "CONSULTA":
                    print(f"\n[{index + 1}/{total_linhas}] Linha ID {row.get('ID')}: {nome} | [CONSULTA]")
                    relatorio.append({
                        "ID": row.get('ID'),
                        "Nome": nome,
                        "CPF": cpf,
                        "Data": data_atend,
                        "Hora": hora_atend,
                        "TipoAcao": tipo_acao,
                        "Modalidade": modalidade,
                        "Status": "CONSULTA REGISTRADA",
                        "Obs": "Atendimento de consulta registrado para fins de gestão"
                    })
                    continue

                # CENÁRIO 3: REGISTRO DE 'EXCLUSÃO'
                if tipo_acao == "EXCLUSÃO":
                    print(f"\n[{index + 1}/{total_linhas}] Linha ID {row.get('ID')}: {nome} | [EXCLUSÃO] (Requer Ação Manual)")
                    relatorio.append({
                        "ID": row.get('ID'),
                        "Nome": nome,
                        "CPF": cpf,
                        "Data": data_atend,
                        "Hora": hora_atend,
                        "TipoAcao": tipo_acao,
                        "Modalidade": modalidade,
                        "Status": "EXCLUSÃO PENDENTE",
                        "Obs": "Atendimento de exclusão registrado - Tratar no GESUAS"
                    })
                    continue

                # CENÁRIO 4: INCLUSÕES E ATUALIZAÇÕES PENDENTES -> EXECUTA NO EDGE
                print(f"\n[{index + 1}/{total_linhas}] Linha ID {row.get('ID')}: {nome} | Ação: [{tipo_acao}] | Modalidade: [{modalidade}]")

                try:
                    # 1 e 2: Busca por CPF / Nome
                    encontrou = buscar_municipe(page, nome, cpf, tipo_acao)
                    if not encontrou:
                        alerta_intervencao_manual(
                            "Munícipe Não Encontrado",
                            f"Atenção: Não localizado no GESUAS!\n\n"
                            f"Nome: {nome}\nCPF: {cpf}\nAção: {tipo_acao}\n\n"
                            "Verifique a situação no Edge.\n"
                            "Ao clicar em OK, o script avançará para o próximo registro."
                        )
                        relatorio.append({
                            "ID": row.get('ID'),
                            "Nome": nome,
                            "CPF": cpf,
                            "Data": data_atend,
                            "Hora": hora_atend,
                            "TipoAcao": tipo_acao,
                            "Modalidade": modalidade,
                            "Status": "NÃO ENCONTRADO",
                            "Obs": f"Não localizado por CPF nem Nome ({tipo_acao})"
                        })
                        continue

                    # 3: Checagem de Pessoa Desligada e acesso à Família
                    status_familia = verificar_pessoa_desligada_ou_acessar_familia(page, nome, tipo_acao)
                    if status_familia == "DESLIGADA":
                        relatorio.append({
                            "ID": row.get('ID'),
                            "Nome": nome,
                            "CPF": cpf,
                            "Data": data_atend,
                            "Hora": hora_atend,
                            "TipoAcao": tipo_acao,
                            "Modalidade": modalidade,
                            "Status": "PESSOA DESLIGADA",
                            "Obs": "Pessoa desligada da família - Tratar manualmente"
                        })
                        continue

                    # 4: Conferência obrigatória em 'Editar um Atendimento'
                    duplicado = checar_duplicidade_atendimento(page, nome, tipo_acao, data_hora_str, data_atend, hora_atend)
                    if duplicado:
                        relatorio.append({
                            "ID": row.get('ID'),
                            "Nome": nome,
                            "CPF": cpf,
                            "Data": data_atend,
                            "Hora": hora_atend,
                            "TipoAcao": tipo_acao,
                            "Modalidade": modalidade,
                            "Status": "JÁ CADASTRADO",
                            "Obs": f"{tipo_acao} já registrada em {data_hora_str}"
                        })
                        continue

                    # 5 e 6: Iniciar e encerrar atendimento
                    obs_adicional = registrar_atendimento_e_encerrar(page, data_hora_str, nome, tipo_acao)

                    relatorio.append({
                        "ID": row.get('ID'),
                        "Nome": nome,
                        "CPF": cpf,
                        "Data": data_atend,
                        "Hora": hora_atend,
                        "TipoAcao": tipo_acao,
                        "Modalidade": modalidade,
                        "Status": "SUCESSO",
                        "Obs": f"{tipo_acao} registrada com êxito{obs_adicional}"
                    })
                    print(f"   [SUCESSO] {tipo_acao} registrada com êxito para {nome}!{obs_adicional}")

                except Exception as e:
                    alerta_intervencao_manual(
                        "Exceção na Operação",
                        f"Falha ao processar {nome} ({tipo_acao}):\n{e}\n\nClique em OK para continuar."
                    )
                    relatorio.append({
                        "ID": row.get('ID'),
                        "Nome": nome,
                        "CPF": cpf,
                        "Data": data_atend,
                        "Hora": hora_atend,
                        "TipoAcao": tipo_acao,
                        "Modalidade": modalidade,
                        "Status": "ERRO",
                        "Obs": str(e)[:30]
                    })
                    continue

            browser.close()

    finally:
        if relatorio:
            gerar_relatorio_pdf(relatorio, aba_selecionada)
        print("\n=== LOTE CONCLUÍDO ===")


def executar():
    """Loop principal de navegação e execução contínua."""
    exibir_cabecalho_terminal()

    while True:
        caminho_planilha = selecionar_planilha()
        if not caminho_planilha or caminho_planilha == "SAIR":
            print("\nEncerrando a Automação GESUAS. Até logo!")
            break

        while True:
            aba_selecionada = selecionar_aba_mes(caminho_planilha)
            if not aba_selecionada or aba_selecionada == "SAIR":
                print("\nEncerrando a Automação GESUAS. Até logo!")
                return
            if aba_selecionada == "VOLTAR":
                break

            processar_lote(caminho_planilha, aba_selecionada)

            print("\n" + "-" * 65)
            opcao = input(f">> Deseja processar outra aba de {os.path.basename(caminho_planilha)}? [S]im / [N]ão (trocar de planilha): ").strip().upper()
            if opcao != "S":
                break


if __name__ == "__main__":
    executar()
    