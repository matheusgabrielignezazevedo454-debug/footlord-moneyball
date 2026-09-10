#!/usr/bin/env python3
"""Processador FM26 Moneyball com saída XLSX limpa e compatível com Excel Mobile."""

from __future__ import annotations

import argparse
import math
import re
import shutil
import sqlite3
import sys
import tempfile
import zipfile
from bisect import bisect_left, bisect_right
from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# Constantes removidas; agora calculadas dinamicamente via get_date_info()

COLORS = {
    "title": "0B3D2E",
    "header": "0E4A36",
    "bad": "F4CCCC",
    "medium": "FFF2CC",
    "good": "D9EAD3",
    "excellent": "CFE2F3",
    "neutral": "FFFFFF",
    "subtitle": "EAF3EE",
}

# Papéis de linha (sem GK) — cada um vira coluna própria
OUTFIELD_ROLES = ["CD", "LB", "RB", "CDM", "CM", "CAM", "LM", "RM", "LW", "RW", "ST"]

PLAYERS_HEADERS = [
    "Rank", "Jogador", "Posição",
    "Role CD", "Role LB", "Role RB", "Role CDM", "Role CM", "Role CAM", "Role LM", "Role RM", "Role LW", "Role RW", "Role ST",
    "Idade", "Time", "Liga", "Divisão", "País",
    "Overall do Footlord (0-100)", "Overall início da temporada", "Variação overall na temporada", "Overall início da carreira", "Variação overall desde início da carreira",
    "Jogos no ano", "Cartões amarelos", "Cartões vermelhos", "Cartões totais", "Amarelos/partida", "Vermelhos/partida",
    "Gols no ano", "Gols sem pênalti", "Pênaltis batidos", "Gols de pênalti", "Pênaltis perdidos (total)", "Pênaltis perdidos (fora)", "Taxa conversão pênalti", "xGP (pênaltis esperados)", "Gols pênalti vs xGP",
    "Assistências no ano", "G+A no ano", "Gols/partida", "Assistências/partida", "G+A/partida",
    "Gols decisivos", "Gols salvadores", "Assistências decisivas", "Assistências salvadoras", "Hat-tricks",
    "Entradas do banco", "Gols como reserva", "Gols decisivos como reserva", "Gols salvadores como reserva",
    "Assistências como reserva", "Assistências decisivas como reserva", "Assistências salvadoras como reserva",
    "MVPs como reserva",
    "Média de avaliação", "Minutos estimados", "Minutos/partida", "Minutos/gol", "Minutos/assistência",
    "MVP do time", "Pior nota do time", "MVP %", "Pior nota %", "Clean sheets",
    "Valor de mercado em €", "Salário em €", "Contrato (anos)", "Status", "Empréstimo", "Lista de transferências",
    "PAC", "SHO", "PAS", "DEF", "PHY", "MEN", "GKP", "ID interno",
]


GK_HEADERS = [
    "Rank", "Goleiro", "Time", "Liga", "Partidas GK", "Minutos", "Clean sheets", "Clean sheets/partida", "xG adversário", "Gols sofridos", "Gols/partida GK",
    "Gols evitados vs xG", "Gols evitados/partida", "Gols evitados/90", "Taxa de prevenção", "Gols/xG", 
    "Pênaltis enfrentados", "Pênaltis defendidos", "Taxa defesa pênalti", "xGP (pênaltis esperados)", "Pênaltis defendidos vs xGP",
    "Classificação", "Elegível (mín. 5 jogos)", "Overall do Footlord", "ID interno",
]

CAREER_HEADERS = [
    "Id", "Nome do jogador", "Posição",
    "Role CD", "Role LB", "Role RB", "Role CDM", "Role CM", "Role CAM", "Role LM", "Role RM", "Role LW", "Role RW", "Role ST",
    "Idade", "Nacionalidade", "Situação",
    "Overall início do save", "Maior overall da carreira", "Overall atual",
    "Jogos na carreira", "Jogos da seleção", "Jogos média por temporada",
    "Gols na carreira", "Gols/partida", "Gols/temporada", "Gols na seleção",
    "Assistências na carreira", "Assistências/partida", "Assistências/temporada", "Assistências na seleção",
    "G+A na carreira", "G+A/temporada", "G+A/partida",
    "Média avaliação carreira", "Média avaliação seleção",
    "Clean sheets na carreira", "Clean sheets/temporada",
    "Valor de mercado atual €",
    # TOP 11 / Time do ano / Seleção do torneio = h_rewards tipo 5
    "Time do ano — ligas", "Time do ano — copas nacionais",
    "Seleção do torneio — continental", "Seleção do torneio — Copa do Mundo",
    "Vezes no time do ano (total)",
    # MVP de PARTIDA = h_rewards tipo 6 (NÃO é time do ano)
    "MVP partida — ligas", "MVP partida — copas nacionais",
    "MVP partida — continental", "MVP partida — Copa do Mundo", "MVP partida — total",
    "Artilheiro — ligas", "Artilheiro — copas nacionais", "Artilheiro — continental", "Artilheiro — Copa do Mundo",
    "Maior assistência — ligas", "Maior assistência — copas nacionais", "Maior assistência — continental", "Maior assistência — Copa do Mundo",
    "Melhor goleiro — ligas", "Melhor goleiro — copas nacionais", "Melhor goleiro — continental", "Melhor goleiro — Copa do Mundo",
    "Títulos na carreira", "Títulos de liga", "Títulos de copa nacional", "Títulos de supercopa nacional",
    "Títulos de supercopa continental", "Títulos de competição continental", "Copas do Mundo",
    "Chuteira de ouro", "Luva de ouro", "Golden Boy", "Bola de ouro",
]

TIMES_HEADERS = [
    "Rank overall", "Time", "Liga", "Competição continental", "Jogadores no time", "Média de idade",
    "Partidas no ano", "Vitórias no ano", "Empates no ano", "Derrotas no ano", "Cartões amarelos no ano", "Cartões vermelhos no ano", "Cartões totais no ano", "Amarelos/partida", "Vermelhos/partida", "Partidas antes 1/jan", "Vitórias antes", "Empates antes", "Derrotas antes",
    "Partidas depois 1/jan", "Vitórias depois", "Empates depois", "Derrotas depois", "Maior sequência vitórias antes", "Maior sequência derrotas antes",
    "Maior sequência vitórias depois", "Maior sequência derrotas depois",
    "Maior sequência invencibilidade da temporada", "Maior sequência sem vencer da temporada",
    "Gols marcados no ano", "Gols sofridos no ano", "Gols marcados/partida",
    "Gols sofridos/partida", "Chutes totais", "Chutes/partida", "Chutes no alvo", "% no alvo", "Conversão de chutes",
    "xG criado na temporada", "xG sofrido na temporada", "Média de xG criado por partida", "Média de xG sofrido por partida",
    "Gols marcados − xG criado", "xG sofrido − gols sofridos", "xG criado − xG sofrido", "Saldo de gols − saldo xG",
    "Gols marcados antes", "Gols sofridos antes", "Gols marcados/partida antes", "Gols sofridos/partida antes",
    "Gols marcados depois", "Gols sofridos depois", "Gols marcados/partida depois", "Gols sofridos/partida depois", "Jogos sem marcar",
    "Jogos sem sofrer gols", "Sem marcar antes", "Sem sofrer gols antes", "Sem marcar depois", "Sem sofrer gols depois", "Maior placar aplicado no ano",
    "Para quem — maior placar ano", "Pior placar sofrido no ano", "Para quem — pior placar ano", "Maior placar aplicado antes", "Para quem — maior placar antes",
    "Pior placar sofrido antes", "Para quem — pior placar antes", "Maior placar aplicado depois", "Para quem — maior placar depois", "Pior placar sofrido depois",
    "Para quem — pior placar depois", "Overall início da temporada", "Overall meio da temporada", "Overall atual / final da temporada", "Expectativa — posição por overall inicial",
    "Posição na liga no save", "Diferença expectativa/posição", "Média overall adversário — vitória", "Média overall adversário — empate",
    "Média overall adversário — derrota", "Valor do time — elenco em €", "Média salário dos jogadores em €", "Saldo bancário em €",
    "Gastos janela janeiro em €", "Vendas janela janeiro em €", "Overall médio compras janeiro", "Overall médio vendas janeiro",
    "Gastos janela agosto em €", "Vendas janela agosto em €", "Overall médio compras agosto", "Overall médio vendas agosto",
]

HEADER_TRANSLATIONS_EN = {
    "Maior sequência invencibilidade da temporada": "Longest unbeaten streak (season)",
    "Maior sequência sem vencer da temporada": "Longest winless streak (season)",
    "xG criado na temporada": "xG created (season)",
    "xG sofrido na temporada": "xG conceded (season)",
    "Média de xG criado por partida": "xG created per match",
    "Média de xG sofrido por partida": "xG conceded per match",
    "Gols marcados − xG criado": "Goals scored − xG created",
    "xG sofrido − gols sofridos": "xG conceded − goals conceded",
    "xG criado − xG sofrido": "xG created − xG conceded",
    "Saldo de gols − saldo xG": "Goal difference − xG difference",

    "Gols decisivos": "Winning goals",
    "Gols salvadores": "Equalizer goals",
    "Assistências decisivas": "Winning assists",
    "Assistências salvadoras": "Equalizer assists",
    "Hat-tricks": "Hat-tricks",
    "Entradas do banco": "Sub appearances",
    "Gols como reserva": "Goals as sub",
    "Gols decisivos como reserva": "Winning goals as sub",
    "Gols salvadores como reserva": "Equalizer goals as sub",
    "Assistências como reserva": "Assists as sub",
    "Assistências decisivas como reserva": "Winning assists as sub",
    "Assistências salvadoras como reserva": "Equalizer assists as sub",
    "MVPs como reserva": "MVPs as sub",
    "Overall início da temporada": "Overall season start",
    "Variação overall na temporada": "Overall change this season",
    "Overall início da carreira": "Overall career start",
    "Variação overall desde início da carreira": "Overall change since career start",
    "Cartões amarelos": "Yellow cards",
    "Cartões vermelhos": "Red cards",
    "Cartões totais": "Total cards",
    "Amarelos/partida": "Yellows/match",
    "Vermelhos/partida": "Reds/match",
    "Gols sem pênalti": "Non-penalty goals",
    "Pênaltis batidos": "Penalties taken",
    "Gols de pênalti": "Penalty goals",
    "Pênaltis perdidos (total)": "Penalties missed (total)",
    "Pênaltis perdidos (fora)": "Penalties missed (off target)",
    "Taxa conversão pênalti": "Penalty conversion rate",
    "xGP (pênaltis esperados)": "xGP (expected penalties)",
    "Gols pênalti vs xGP": "Penalty goals vs xGP",
    "MVP do time": "Team MVP",
    "Pior nota do time": "Team worst rating",
    "Pior nota %": "Worst rating %",
    "Valor de mercado em €": "Market value €",
    "Salário em €": "Wage €",
    "Contrato (anos)": "Contract (years)",
    "Empréstimo": "Loan",
    "Lista de transferências": "Transfer listed",
    "ID interno": "Internal ID",
    "Divisão": "Division",
    "País": "Country",
    "Rank": "Rank", "Quality Index (0-100)": "Quality Index (0-100)", "Quality Index": "Quality Index", "Score Moneyball (0-100)": "Moneyball Score (0-100)", "Score Moneyball": "Moneyball Score", "MVP %": "MVP %", "Clean sheets": "Clean sheets", "Clean sheets/partida": "Clean sheets/match", "Status": "Status", "PAC": "PAC", "SHO": "SHO", "PAS": "PAS", "DEF": "DEF", "PHY": "PHY", "MEN": "MEN", "GKP": "GKP", "Gols/xG": "Goals/xG",
    "Jogador": "Player", "Posição": "Position", "Role CD": "Role CD", "Role LB": "Role LB", "Role RB": "Role RB", "Role CDM": "Role CDM", "Role CM": "Role CM", "Role CAM": "Role CAM", "Role LM": "Role LM", "Role RM": "Role RM", "Role LW": "Role LW", "Role RW": "Role RW", "Role ST": "Role ST", "Overall do Footlord (0-100)": "Footlord Overall (0-100)", "Overall do Footlord": "Footlord Overall", "Idade": "Age", "Time": "Team", "Liga": "League", "Divisão": "Division", "País": "Country",
    "Jogos no ano": "Matches in year", "Gols no ano": "Goals in year", "Assistências no ano": "Assists in year", "G+A no ano": "G+A in year", "Gols/partida": "Goals/match", "Assistências/partida": "Assists/match", "G+A/partida": "G+A/match", "Média de avaliação": "Average rating", "Minutos estimados": "Estimated minutes", "Partidas com minutos": "Matches with minutes", "Minutos/partida": "Minutes/match", "Minutos/gol": "Minutes/goal", "Minutos/assistência": "Minutes/assist", "MVP do time": "Team MVP", "Pior nota do time": "Team worst rating", "Partidas com rating": "Matches with rating", "Pior nota %": "Worst rating %", "Partidas GK": "GK matches", "xG adversário GK": "Opponent xG GK", "Gols sofridos GK": "Goals conceded GK", "Gols evitados vs xG": "Goals prevented vs xG", "xG/partida GK": "xG/match GK", "Gols/partida GK": "Goals/match GK", "Gols/xG GK": "Goals/xG GK", "Valor de mercado em €": "Market value in €", "Salário em €": "Salary in €", "Contrato (anos)": "Contract (years)", "Empréstimo": "Loan", "Lista de transferências": "Transfer listed", "ID interno": "Internal ID",
    "Goleiro": "Goalkeeper", "Minutos": "Minutes", "xG adversário": "Opponent xG", "Gols sofridos": "Goals conceded", "Gols evitados/partida": "Goals prevented/match", "Gols evitados/90": "Goals prevented/90", "Taxa de prevenção": "Prevention rate", "Classificação": "Rating", "Elegível (mín. 5 jogos)": "Eligible (min. 5 matches)",
    "Rank cartões": "Cards rank", "Cartões amarelos": "Yellow cards", "Cartões vermelhos": "Red cards", "Cartões totais": "Total cards", "Amarelos/partida": "Yellow cards/match", "Vermelhos/partida": "Red cards/match", "Cartões amarelos no ano": "Yellow cards in year", "Cartões vermelhos no ano": "Red cards in year", "Cartões totais no ano": "Total cards in year", "Overall início da temporada": "Overall at season start", "Overall atual": "Current overall", "Variação overall na temporada": "Overall change in season", "Overall início da carreira": "Overall at career start", "Variação overall desde início da carreira": "Overall change since career start",
    "Rank overall": "Overall rank", "Competição continental": "Continental competition", "Jogadores no time": "Players in team", "Média de idade": "Average age", "Partidas no ano": "Matches in year", "Vitórias no ano": "Wins in year", "Empates no ano": "Draws in year", "Derrotas no ano": "Losses in year", "Partidas antes 1/jan": "Matches before Jan 1", "Vitórias antes": "Wins before", "Empates antes": "Draws before", "Derrotas antes": "Losses before", "Partidas depois 1/jan": "Matches after Jan 1", "Vitórias depois": "Wins after", "Empates depois": "Draws after", "Derrotas depois": "Losses after", "Maior sequência vitórias antes": "Longest win streak before", "Maior sequência derrotas antes": "Longest loss streak before", "Maior sequência vitórias depois": "Longest win streak after", "Maior sequência derrotas depois": "Longest loss streak after", "Gols marcados no ano": "Goals scored in year", "Gols sofridos no ano": "Goals conceded in year", "Gols marcados/partida": "Goals scored/match", "Gols sofridos/partida": "Goals conceded/match", "Chutes totais": "Total shots", "Chutes/partida": "Shots/match", "Chutes no alvo": "Shots on target", "% no alvo": "% on target", "Conversão de chutes": "Shot conversion", "Gols marcados antes": "Goals scored before", "Gols sofridos antes": "Goals conceded before", "Gols marcados/partida antes": "Goals scored/match before", "Gols sofridos/partida antes": "Goals conceded/match before", "Gols marcados depois": "Goals scored after", "Gols sofridos depois": "Goals conceded after", "Gols marcados/partida depois": "Goals scored/match after", "Gols sofridos/partida depois": "Goals conceded/match after", "Jogos sem marcar": "Matches without scoring", "Jogos sem sofrer gols": "Clean sheets", "Sem marcar antes": "Without scoring before", "Sem sofrer gols antes": "Clean sheets before", "Sem marcar depois": "Without scoring after", "Sem sofrer gols depois": "Clean sheets after", "Maior placar aplicado no ano": "Biggest win in year", "Para quem — maior placar ano": "Opponent — biggest win in year", "Pior placar sofrido no ano": "Worst loss in year", "Para quem — pior placar ano": "Opponent — worst loss in year", "Maior placar aplicado antes": "Biggest win before", "Para quem — maior placar antes": "Opponent — biggest win before", "Pior placar sofrido antes": "Worst loss before", "Para quem — pior placar antes": "Opponent — worst loss before", "Maior placar aplicado depois": "Biggest win after", "Para quem — maior placar depois": "Opponent — biggest win after", "Pior placar sofrido depois": "Worst loss after", "Para quem — pior placar depois": "Opponent — worst loss after", "Overall início da temporada": "Overall at season start", "Overall meio da temporada": "Overall mid-season", "Overall atual / final da temporada": "Overall current / season end", "Overall início do ano": "Overall at season start", "Overall fechamento agosto": "Overall mid-season", "Overall final do ano": "Overall current / season end", "Expectativa — posição por overall inicial": "Expected position by starting overall", "Posição na liga no save": "League position in save", "Diferença expectativa/posição": "Expectation/position difference", "Média overall adversário — vitória": "Average opponent overall — win", "Média overall adversário — empate": "Average opponent overall — draw", "Média overall adversário — derrota": "Average opponent overall — loss", "Valor do time — elenco em €": "Team value — squad in €", "Média salário dos jogadores em €": "Average player salary in €", "Saldo bancário em €": "Bank balance in €", "Gastos janela janeiro em €": "January window spending in €", "Vendas janela janeiro em €": "January window sales in €", "Overall médio compras janeiro": "Average overall — January signings", "Overall médio vendas janeiro": "Average overall — January sales", "Gastos janela agosto em €": "August window spending in €", "Vendas janela agosto em €": "August window sales in €", "Overall médio compras agosto": "Average overall — August signings", "Overall médio vendas agosto": "Average overall — August sales",
    "Gols sem pênalti": "Non-penalty goals", "Pênaltis batidos": "Penalties taken", "Gols de pênalti": "Penalty goals", "Pênaltis perdidos (total)": "Penalties missed (total)", "Pênaltis perdidos (fora)": "Penalties missed (out)", "Taxa conversão pênalti": "Penalty conversion rate", "xGP (pênaltis esperados)": "xGP (expected penalties)", "Gols pênalti vs xGP": "Penalty goals vs xGP",
    "Pênaltis enfrentados": "Penalties faced", "Pênaltis defendidos": "Penalties saved", "Taxa defesa pênalti": "Penalty save rate", "Pênaltis defendidos vs xGP": "Penalties saved vs xGP",
}


def localized_headers(headers, language):
    return [HEADER_TRANSLATIONS_EN.get(header, header) for header in headers] if language == "en" else headers

NUM_FORMAT_INT = '#,##0;[Red]-#,##0'
NUM_FORMAT_DEC1 = '0.0;[Red]-0.0'
NUM_FORMAT_DEC2 = '0.00;[Red]-0.00'
NUM_FORMAT_DEC3 = '0.000;[Red]-0.000'
NUM_FORMAT_PCT = '0.0%;[Red]-0.0%'

THIN_BLACK = Side(style="thin", color="000000")
BORDER = Border(left=THIN_BLACK, right=THIN_BLACK, top=THIN_BLACK, bottom=THIN_BLACK)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=False)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=False)
CENTER_WRAP = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT_WRAP = Alignment(horizontal="left", vertical="center", wrap_text=True)
TITLE_FONT = Font(name="Aptos Display", size=14, bold=True, color="FFFFFF")
HEADER_FONT = Font(name="Aptos", size=9, bold=True, color="FFFFFF")
DATA_FONT = Font(name="Aptos", size=9, color="111827")
FILLS = {name: PatternFill("solid", fgColor=value) for name, value in COLORS.items()}


def emit(percent: int, message: str) -> None:
    print(f"PROGRESS|{percent}|{message}", flush=True)


def sf(value, default=0.0):
    try:
        return default if value is None or value == "" else float(value)
    except (TypeError, ValueError):
        return default


def si(value, default=0):
    try:
        return default if value is None or value == "" else int(float(value))
    except (TypeError, ValueError):
        return default


def clip(value, low=0.0, high=100.0):
    return max(low, min(high, sf(value)))


def parse_ids(raw):
    return [si(value) for value in str(raw or "").split("_") if str(value).strip()]


def parse_values(raw):
    values = []
    for value in str(raw or "").split("_"):
        try:
            values.append(float(value))
        except (TypeError, ValueError):
            pass
    return values


def parse_count_blocks(raw):
    total = 0.0
    for item in str(raw or "").split(","):
        if ":" not in item:
            continue
        try:
            total += float(item.rsplit(":", 1)[1])
        except (TypeError, ValueError):
            pass
    return total


def parse_rating_blocks(raw):
    values = []
    for item in str(raw or "").split(","):
        if ":" not in item:
            continue
        for value in item.rsplit(":", 1)[1].split("|"):
            try:
                values.append(float(value))
            except (TypeError, ValueError):
                pass
    return values


def parse_match_ratings(raw):
    result = {}
    for item in str(raw or "").split(";"):
        if "|" not in item:
            continue
        player_id, rating = item.split("|", 1)
        try:
            result[int(player_id)] = float(rating)
        except (TypeError, ValueError):
            pass
    return result


def id_set(value):
    return {int(part) for part in str(value or "").split("_") if part.isdigit()}


def rating_ids(value):
    ids = set()
    for part in str(value or "").split(";"):
        if "|" in part:
            try:
                ids.add(int(part.split("|", 1)[0]))
            except (TypeError, ValueError):
                pass
    return ids


def parse_role_levels(raw):
    """Parse 'GK:10.00,CD:1.78,...' -> dict role -> float."""
    result = {}
    for item in str(raw or "").split(","):
        if ":" not in item:
            continue
        role, value = item.split(":", 1)
        try:
            result[role.strip()] = float(value)
        except (TypeError, ValueError):
            pass
    return result


def parse_events(raw):
    """Eventos de partida. Minuto pode passar de 119 (pênaltis de desempate).
    Formato: minuto|codigo|player_id separados por ; (às vezes ,).
    Códigos: 0=gol, 1=assistência, 2=amarelo, 3=vermelho, 5=entrou, 6=saiu, 7=pênalti gol, 8=pênalti perdido.
    """
    events = []
    for item in str(raw or "").replace(",", ";").split(";"):
        parts = item.split("|")
        if len(parts) != 3:
            continue
        try:
            events.append((max(0, int(float(parts[0]))), int(float(parts[1])), int(float(parts[2]))))
        except (TypeError, ValueError):
            pass
    return events


def analyze_match_goals(events_1, events_2, goals_1, goals_2, lineup_1, lineup_2, entered_1, entered_2):
    """Gols/assistências decisivos e salvadores + hat-tricks + ações como reserva.

    Decisivo: gol de empate → vantagem E o time venceu a partida.
    Salvador: gol de desvantagem → empate.
    Assistência herda a classificação do gol no mesmo minuto.
    Reserva: jogador em entered e fora do lineup titular.
    """
    all_ev = []
    for m, c, p in events_1:
        all_ev.append((m, c, p, 1))
    for m, c, p in events_2:
        all_ev.append((m, c, p, 2))
    all_ev.sort(key=lambda x: (x[0], 0 if x[1] == 0 else 1 if x[1] == 1 else 2))

    lineup = {1: set(lineup_1 or []), 2: set(lineup_2 or [])}
    # Reserva = jogador que ENTROU via evento código 5 (não usar só o campo entered_*,
    # pois em alguns saves ele se sobrepõe ao lineup).
    sub_on = {1: set(), 2: set()}
    entry_min = {}
    for m, c, p, side in all_ev:
        if c == 5:
            sub_on[side].add(p)
            entry_min[p] = m

    def is_sub(pid, side):
        return pid in sub_on[side]

    score = {1: 0, 2: 0}
    goals_by_player = defaultdict(int)
    goal_log = []  # (pid, side, kind, is_sub) kind: goal|decisive|savior|assist|assist_decisive|assist_savior
    last_goal = None

    for m, c, p, side in all_ev:
        if c == 0 and m < 120:
            opp = 3 - side
            before_own, before_opp = score[side], score[opp]
            score[side] += 1
            goals_by_player[p] += 1
            sub = is_sub(p, side)
            if before_own == before_opp:
                kind = "decisive"
            elif before_own < before_opp:
                kind = "savior"
            else:
                kind = "goal"
            last_goal = (p, side, kind, sub, m)
            goal_log.append(last_goal)
        elif c == 1 and m < 120 and last_goal and last_goal[4] == m and last_goal[1] == side:
            sub = is_sub(p, side)
            gk = last_goal[2]
            ak = "assist_decisive" if gk == "decisive" else ("assist_savior" if gk == "savior" else "assist")
            goal_log.append((p, side, ak, sub, m))

    g1, g2 = int(goals_1 or 0), int(goals_2 or 0)
    winner = 1 if g1 > g2 else (2 if g2 > g1 else 0)

    out = defaultdict(lambda: {
        "dec_goals": 0, "sav_goals": 0, "dec_ast": 0, "sav_ast": 0, "hat": 0,
        "sub_apps": 0, "sub_goals": 0, "sub_dec_goals": 0, "sub_sav_goals": 0,
        "sub_ast": 0, "sub_dec_ast": 0, "sub_sav_ast": 0, "sub_mvp": 0,
    })

    for pid, side, kind, sub, m in goal_log:
        if kind == "decisive" and side == winner:
            out[pid]["dec_goals"] += 1
            if sub:
                out[pid]["sub_dec_goals"] += 1
                out[pid]["sub_goals"] += 1
        elif kind == "savior":
            out[pid]["sav_goals"] += 1
            if sub:
                out[pid]["sub_sav_goals"] += 1
                out[pid]["sub_goals"] += 1
        elif kind == "goal" and sub:
            out[pid]["sub_goals"] += 1
        elif kind == "assist_decisive" and side == winner:
            out[pid]["dec_ast"] += 1
            if sub:
                out[pid]["sub_dec_ast"] += 1
                out[pid]["sub_ast"] += 1
        elif kind == "assist_savior":
            out[pid]["sav_ast"] += 1
            if sub:
                out[pid]["sub_sav_ast"] += 1
                out[pid]["sub_ast"] += 1
        elif kind == "assist" and sub:
            out[pid]["sub_ast"] += 1

    for pid, cnt in goals_by_player.items():
        if cnt >= 3:
            out[pid]["hat"] += 1

    for side in (1, 2):
        for pid in sub_on[side]:
            out[pid]["sub_apps"] += 1

    return out



def estimate_minutes(lineup, entered, subbed, events):
    minutes = {pid: 90.0 for pid in lineup}
    for pid in entered:
        minutes[pid] = 45.0
    for minute, code, pid in events:
        if code == 6 and pid in minutes:
            minutes[pid] = float(minute)
        elif code == 5:
            minutes[pid] = 90.0 - float(minute)
    return minutes


def percentile(values, value):
    if not values:
        return 50.0
    sorted_values = sorted(values)
    idx = bisect_left(sorted_values, value)
    return clip(idx / len(sorted_values) * 100)


def get_date_info(cur):
    gen = dict(cur.execute("SELECT * FROM general LIMIT 1").fetchone())
    current_day = si(gen["date"])
    starting_year = si(gen.get("starting_year"), 2025)
    # 1º de julho do ano inicial é o dia 0
    start_date = datetime(starting_year, 7, 1)
    current_date = start_date + timedelta(days=current_day)
    
    # 1º de janeiro do ano civil ATUAL (em relação ao current_day)
    jan_1st = datetime(current_date.year, 1, 1)
    # Se hoje é antes de julho, o 1º de janeiro relevante é o deste ano.
    # Se hoje é julho ou depois, o 1º de janeiro relevante é o do ano que vem?
    # O usuário quer dividir a temporada. Temporada europeia: Jul-Jun.
    # Se estamos em Fev/2030, o "antes de 1/jan" é Jul/2029-Dez/2029.
    # Se estamos em Out/2030, o "antes de 1/jan" é Jul/2030-Dez/2030 (vazio) e "depois" é Jan/2031+ (vazio).
    # Regra: O 1º de janeiro que divide a temporada atual.
    # Se mês >= 7: o 1º de jan é o do ano seguinte.
    # Se mês < 7: o 1º de jan é o deste ano.
    target_jan = datetime(current_date.year if current_date.month < 7 else current_date.year + 1, 1, 1)
    jan_cutoff = (target_jan - start_date).days

    # Início da TEMPORADA atual (1º de julho): não usar ano civil inteiro
    if current_date.month >= 7:
        season_july = datetime(current_date.year, 7, 1)
    else:
        season_july = datetime(current_date.year - 1, 7, 1)
    season_start_day = max(0, (season_july - start_date).days)
    
    managed_id = si(gen.get("team_id"), -1)
    managed_name = "Clube"
    row = cur.execute("SELECT id, name FROM teams2 WHERE id=?", (managed_id,)).fetchone()
    if row:
        managed_id, managed_name = si(row[0]), (row[1] or "Clube")
    else:
        # team_id fantasma: tenta pelo manager_id do general
        mid = si(gen.get("manager_id"), -1)
        row = cur.execute(
            "SELECT id, name FROM teams2 WHERE manager_id=? AND COALESCE(is_national_team,0)=0 LIMIT 1",
            (mid,),
        ).fetchone() if mid > 0 else None
        if row:
            managed_id, managed_name = si(row[0]), (row[1] or "Clube")
        else:
            managed_id = -1
            managed_name = "Clube"
    return {
        "current_day": current_day,
        "current_date": current_date,
        "start_date": start_date,
        "jan_cutoff": jan_cutoff,
        "season_start_day": season_start_day,
        "managed_id": managed_id,
        "managed_name": managed_name,
    }

def build_team_rows(cur, date_info, players, played_matches, team_cards):
    teams = {row["id"]: dict(row) for row in cur.execute("SELECT * FROM teams2").fetchall()}
    leagues = {row["id"]: dict(row) for row in cur.execute("SELECT * FROM leagues2").fetchall()}
    champs = {row["id"]: dict(row) for row in cur.execute("SELECT * FROM champs2").fetchall()}
    int_cups = list(cur.execute("SELECT id, continent, name FROM int_cups2").fetchall())
    # mapa IC_{CONT}_{id} → nome
    ic_name = {}
    for row in int_cups:
        ic_name[(str(row["continent"]).upper(), si(row["id"]))] = row["name"]

    def team_ov(tid):
        t = teams.get(tid) or {}
        return sf(t.get("value")) or 50.0

    def parse_ch_points(team_row):
        """Pontos na liga atual a partir de s_wins/s_draws (só blocos CH_)."""
        champ_id = si(team_row.get("champ_id"), -1)
        if champ_id <= 0:
            return 0
        key = f"CH_{champ_id}"
        wins = draws = 0.0
        for raw, target in ((team_row.get("s_wins"), "w"), (team_row.get("s_draws"), "d")):
            for item in str(raw or "").split(","):
                if ":" not in item:
                    continue
                left, right = item.rsplit(":", 1)
                if left.strip().upper().startswith(key) or left.strip().upper() == key:
                    try:
                        val = float(right)
                    except (TypeError, ValueError):
                        val = 0.0
                    if target == "w":
                        wins += val
                    else:
                        draws += val
        return int(3 * wins + draws)

    def continental_name(team_row):
        """Melhor competição continental da temporada via s_wins (IC_XX_n)."""
        best = None  # (priority, name) — menor id = mais importante
        for raw in (team_row.get("s_wins"), team_row.get("s_played"), team_row.get("s_gs")):
            for item in str(raw or "").split(","):
                left = item.split(":")[0].strip().upper()
                if not left.startswith("IC_"):
                    continue
                # IC_EU_0 / IC_SA_1 / ICK_SA_0
                parts = left.split("_")
                if len(parts) < 3:
                    continue
                cont, cid = parts[1], parts[2]
                try:
                    cid_i = int(cid)
                except ValueError:
                    continue
                name = ic_name.get((cont, cid_i))
                if not name:
                    name = f"{cont} #{cid_i}"
                pri = cid_i  # 0 = principal (Libertadores/UCL)
                if best is None or pri < best[0]:
                    best = (pri, name)
        return best[1] if best else "Nenhuma"

    for team_id, team in teams.items():
        champ_id = si(team.get("champ_id"), -1)
        champ = champs.get(champ_id, {})
        league_id = si(champ.get("league_id"), -1)
        league = leagues.get(league_id, {})
        team.update({
            "league": league.get("name") or champ.get("name") or "Sem liga",
            "league_division": league.get("division"),
            "champ_id": champ_id,
            "champ_name": champ.get("name") or "",
            "continent_comp": continental_name(team),
            "league_points": parse_ch_points(team),
            "all": {"games": 0, "wins": 0, "draws": 0, "losses": 0, "goals_for": 0, "goals_against": 0, "xg_for": 0.0, "xg_against": 0.0, "shots": 0, "on_target": 0, "clean_sheets": 0, "failed_to_score": 0, "streak_w": 0, "streak_l": 0, "cur_w": 0, "cur_l": 0, "streak_unb": 0, "streak_wl": 0, "cur_unb": 0, "cur_wl": 0, "opp_ov_w": [], "opp_ov_d": [], "opp_ov_l": [], "biggest_win": (0, ""), "worst_loss": (0, "")},
            "before": {"games": 0, "wins": 0, "draws": 0, "losses": 0, "goals_for": 0, "goals_against": 0, "clean_sheets": 0, "failed_to_score": 0, "streak_w": 0, "streak_l": 0, "cur_w": 0, "cur_l": 0, "biggest_win": (0, ""), "worst_loss": (0, "")},
            "after": {"games": 0, "wins": 0, "draws": 0, "losses": 0, "goals_for": 0, "goals_against": 0, "clean_sheets": 0, "failed_to_score": 0, "streak_w": 0, "streak_l": 0, "cur_w": 0, "cur_l": 0, "biggest_win": (0, ""), "worst_loss": (0, "")},
            # Elenco principal: exclui jogadores da base (academy=1)
            "players": [p for p in players.values()
                        if si(p.get("team_id"), -1) == team_id and si(p.get("academy"), 0) != 1],
            "yellow_cards": team_cards.get(team_id, {}).get("yellow", 0),
            "red_cards": team_cards.get(team_id, {}).get("red", 0),
        })

    # Ordenar por data para sequências corretas
    for raw in sorted(played_matches, key=lambda m: (si(m["date"]), si(m["id"]) if "id" in m.keys() else 0)):
        match = dict(raw)
        t1, t2 = si(match["team_1_id"]), si(match["team_2_id"])
        date = si(match["date"])
        for tid, side in [(t1, 1), (t2, 2)]:
            if tid not in teams:
                continue
            t, other = teams[tid], 3 - side
            og = si(match.get(f"goals_{other}"))
            my_g = si(match.get(f"goals_{side}"))
            my_xg = sf(match.get(f"x_goals_{side}"))
            opp_xg = sf(match.get(f"x_goals_{other}"))
            my_s = si(match.get(f"shots_on_{side}")) + si(match.get(f"shots_off_{side}"))
            my_ot = si(match.get(f"shots_on_{side}"))
            period = t["before"] if date < date_info["jan_cutoff"] else t["after"]
            for d in (t["all"], period):
                d["games"] += 1
                d["goals_for"] += my_g
                d["goals_against"] += og
                d["clean_sheets"] += 1 if og == 0 else 0
                d["failed_to_score"] += 1 if my_g == 0 else 0
                if d is t["all"]:
                    d["shots"] += my_s
                    d["on_target"] += my_ot
                    d["xg_for"] += my_xg
                    d["xg_against"] += opp_xg
                if my_g > og:
                    d["wins"] += 1
                    d["cur_w"] += 1
                    d["cur_l"] = 0
                    d["streak_w"] = max(d["streak_w"], d["cur_w"])
                    if d is t["all"]:
                        d["opp_ov_w"].append(team_ov(si(match.get(f"team_{other}_id"))))
                        # invencível continua; sem vencer zera
                        d["cur_unb"] += 1
                        d["streak_unb"] = max(d["streak_unb"], d["cur_unb"])
                        d["cur_wl"] = 0
                    margin = my_g - og
                    if margin > d["biggest_win"][0]:
                        d["biggest_win"] = (margin, teams.get(si(match.get(f"team_{other}_id")), {}).get("name", "?"))
                elif my_g == og:
                    d["draws"] += 1
                    d["cur_w"] = d["cur_l"] = 0
                    if d is t["all"]:
                        d["opp_ov_d"].append(team_ov(si(match.get(f"team_{other}_id"))))
                        d["cur_unb"] += 1
                        d["streak_unb"] = max(d["streak_unb"], d["cur_unb"])
                        d["cur_wl"] += 1
                        d["streak_wl"] = max(d["streak_wl"], d["cur_wl"])
                else:
                    d["losses"] += 1
                    d["cur_l"] += 1
                    d["cur_w"] = 0
                    d["streak_l"] = max(d["streak_l"], d["cur_l"])
                    if d is t["all"]:
                        d["opp_ov_l"].append(team_ov(si(match.get(f"team_{other}_id"))))
                        d["cur_unb"] = 0
                        d["cur_wl"] += 1
                        d["streak_wl"] = max(d["streak_wl"], d["cur_wl"])
                    margin = og - my_g
                    if margin > d["worst_loss"][0]:
                        d["worst_loss"] = (margin, teams.get(si(match.get(f"team_{other}_id")), {}).get("name", "?"))

    # Posição na liga: ranking por pontos CH_ dentro do mesmo champ_id
    by_champ = defaultdict(list)
    for tid, t in teams.items():
        cid = t.get("champ_id", -1)
        if cid and cid > 0:
            by_champ[cid].append((tid, t.get("league_points", 0), team_ov(tid)))
    league_pos = {}
    for cid, lst in by_champ.items():
        lst.sort(key=lambda x: (-x[1], -x[2]))  # pontos desc, overall desempate
        for pos, (tid, pts, _) in enumerate(lst, 1):
            league_pos[tid] = pos

    # Janelas de transferência (historic_transfers2)
    # Jan: 15/dez → 5/fev | Agosto: 1/jun → 10/set (relativo ao ano civil da data do save)
    start_date = date_info["start_date"]
    current_date = date_info["current_date"]
    current_day = date_info["current_day"]
    season_start = date_info.get("season_start_day", 0)

    def day_of(year, month, day):
        return max(0, (datetime(year, month, day) - start_date).days)

    y = current_date.year
    # janela janeiro da temporada atual (se estamos após jul, jan seguinte; se antes jul, jan deste ano)
    if current_date.month >= 7:
        jan_lo, jan_hi = day_of(y, 12, 15), day_of(y + 1, 2, 5)
        aug_lo, aug_hi = day_of(y, 6, 1), day_of(y, 9, 10)
    else:
        jan_lo, jan_hi = day_of(y - 1, 12, 15), day_of(y, 2, 5)
        aug_lo, aug_hi = day_of(y - 1, 6, 1), day_of(y - 1, 9, 10)

    # gastos/vendas por time
    tw = defaultdict(lambda: {
        "jan_buy": 0, "jan_sell": 0, "jan_buy_ov": [], "jan_sell_ov": [],
        "aug_buy": 0, "aug_sell": 0, "aug_buy_ov": [], "aug_sell_ov": [],
    })
    # overall do jogador no momento ≈ temp_value atual (melhor que nada)
    player_ov = {si(p["id"]): sf(p.get("overall") or p.get("quality") or p.get("temp_value")) for p in players.values()}

    try:
        for row in cur.execute(
            "SELECT player_id, buyer_id, seller_id, transfer_date, transfer_price, transfer_type "
            "FROM historic_transfers2 WHERE transfer_date>=? AND transfer_date<=? AND transfer_price>=0",
            (season_start, current_day),
        ):
            d = si(row["transfer_date"])
            price = si(row["transfer_price"])
            buyer, seller = si(row["buyer_id"]), si(row["seller_id"])
            pov = player_ov.get(si(row["player_id"]), 0.0)
            # type 0 = compra permanente típica; ignora type altos sem preço se quiser
            if jan_lo <= d <= jan_hi:
                if buyer > 0:
                    tw[buyer]["jan_buy"] += price
                    if pov:
                        tw[buyer]["jan_buy_ov"].append(pov)
                if seller > 0:
                    tw[seller]["jan_sell"] += price
                    if pov:
                        tw[seller]["jan_sell_ov"].append(pov)
            elif aug_lo <= d <= aug_hi:
                if buyer > 0:
                    tw[buyer]["aug_buy"] += price
                    if pov:
                        tw[buyer]["aug_buy_ov"].append(pov)
                if seller > 0:
                    tw[seller]["aug_sell"] += price
                    if pov:
                        tw[seller]["aug_sell_ov"].append(pov)
    except Exception:
        pass

    rows = []
    for tid, t in teams.items():
        if not t["all"]["games"] and not t["players"]:
            continue
        all_d, bef, aft = t["all"], t["before"], t["after"]
        # Overall da TEMPORADA (não do ano civil):
        # team_in_time_value = snapshots da temporada atual (início → atual)
        # h_team_in_time_value = histórico longo (fallback)
        hist = parse_values(t.get("team_in_time_value"))
        if len(hist) < 2:
            hist_h = parse_values(t.get("h_team_in_time_value"))
            if hist_h:
                # usa a cauda da temporada (~últimos N pontos) se série curta vazia
                hist = hist_h[-max(6, len(hist_h)//4):] if len(hist_h) > 6 else hist_h
        # início da temporada = 1º snapshot da série da temporada
        # meio = ponto central (aprox. virada do ano / janela de janeiro)
        # final/atual = último snapshot (NÃO usar teams2.value — escala diferente)
        if hist:
            ov_init = round(hist[0], 2)
            ov_aug = round(hist[len(hist) // 2], 2) if len(hist) > 2 else round(hist[-1], 2)
            ov_final = round(hist[-1], 2)
        else:
            ov_init = ov_aug = ov_final = round(sf(t.get("value")) or 0, 2)
        roster = t["players"]
        avg_sal = (
            sum(max(0, si(p.get("salary"))) for p in roster) / len(roster)
            if roster else 0
        )
        n_teams_champ = max(1, len(by_champ.get(t.get("champ_id"), [])))
        # expectativa: rank por overall inicial entre times do mesmo campeonato
        pos = league_pos.get(tid)
        # expectativa por overall: ordenar times do champ por ov_init
        peers = by_champ.get(t.get("champ_id"), [])
        if peers:
            # usar value atual como proxy de overall inicial de ranking esperado
            ranked = sorted(peers, key=lambda x: -x[2])
            exp = next((i for i, (tt, _, _) in enumerate(ranked, 1) if tt == tid), None)
        else:
            exp = None
        diff = (exp - pos) if (exp is not None and pos is not None) else None

        def avg(lst):
            return round(sum(lst) / len(lst), 1) if lst else None

        tr = tw[tid]
        rows.append({
            "name": t.get("name") or f"ID {tid}",
            "league": t.get("league") or "Sem liga",
            "continent": t.get("continent_comp") or "Nenhuma",
            "players_count": len(roster),
            "avg_age": (sum(sf(p.get("age")) for p in roster) / len(roster)) if roster else 0,
            "all": all_d,
            "before": bef,
            "after": aft,
            "yellow_cards": t["yellow_cards"],
            "red_cards": t["red_cards"],
            "ov_init": ov_init,
            "ov_aug": ov_aug,
            "ov_final": ov_final,
            "league_pos": pos,
            "expected_pos": exp,
            "pos_diff": diff,
            "opp_ov_w": avg(all_d["opp_ov_w"]),
            "opp_ov_d": avg(all_d["opp_ov_d"]),
            "opp_ov_l": avg(all_d["opp_ov_l"]),
            "value": si(t.get("season_mk_value")) or sum(max(0, si(p.get("market_value"))) for p in roster),
            "salary": round(avg_sal),
            "bank": si(t.get("bank_balance")),
            "jan_buy": tr["jan_buy"] or None,
            "jan_sell": tr["jan_sell"] or None,
            "jan_buy_ov": avg(tr["jan_buy_ov"]),
            "jan_sell_ov": avg(tr["jan_sell_ov"]),
            "aug_buy": tr["aug_buy"] or None,
            "aug_sell": tr["aug_sell"] or None,
            "aug_buy_ov": avg(tr["aug_buy_ov"]),
            "aug_sell_ov": avg(tr["aug_sell_ov"]),
        })
    return sorted(rows, key=lambda r: (-(1 if r["all"]["games"] else 0), -r["ov_final"]))



def parse_csv_floats(raw):
    values = []
    for item in str(raw or "").split(","):
        item = item.strip()
        if not item:
            continue
        try:
            values.append(float(item))
        except (TypeError, ValueError):
            values.append(0.0)
    return values


def parse_reward_items(raw):
    """h_rewards: 'tipo|COMP:qtd,...' → [(tipo, comp, qtd), ...]
    Tipos (UI Footlord — aba Conquistas):
      0 = artilheiro (global → Chuteira de ouro)
      1 = maior assistência
      2 = global → Bola de Ouro
      3 = Golden Boy (global) / jovem do ano (competição)
      4 = melhor goleiro (global → Luva de ouro)
      5 = Time do ano / Seleção do torneio (TOP 11)
      6 = MVP de PARTIDA (não é time do ano)
    """
    items = []
    for part in str(raw or "").split(","):
        part = part.strip()
        if not part or "|" not in part:
            continue
        left, right = part.split("|", 1)
        if ":" in right:
            comp, cnt = right.rsplit(":", 1)
        else:
            comp, cnt = right, "1"
        try:
            items.append((int(left), comp.strip(), int(float(cnt))))
        except (TypeError, ValueError):
            continue
    return items


def reward_bucket(comp: str) -> str:
    c = (comp or "").upper()
    if c == "GLOBAL" or c.startswith("GLOBAL"):
        return "global"
    if c.startswith("WCQ"):
        return "wcq"
    if c.startswith("WC"):
        return "wc"
    if c.startswith("IC"):
        return "continental"
    if c.startswith("NC"):
        return "national_cup"
    if c.startswith("CH"):
        return "league"
    return "other"


def build_career_rows(cur, current_day: int, starting_year: int = 2025):
    """Carreira completa: ativos (players2) + aposentados (players_retired2)."""
    teams = {row["id"]: row["name"] for row in cur.execute("SELECT id, name FROM teams2")}

    # Títulos / campanhas por jogador (historic_results2, id_type=2)
    # result_type 0 pos=1 → título de liga
    # result_type 1/7 pos=0 → copa nacional (vitória)
    # result_type 2/6 pos=0 → competição continental / supercopa continental
    titles = defaultdict(lambda: {
        "league": 0, "national_cup": 0, "super_national": 0,
        "super_continental": 0, "continental": 0, "world_cup": 0, "total": 0,
    })
    # result_type (id_type=2, pos=0/1 = título):
    #   0 pos=1 → liga
    #   7 pos=0 → copa nacional (vencedor); rt1 é campanha/colocação, não misturar
    #   3 ou 4 pos=0 → supercopa nacional (Community Shield, Supercoppa, etc.)
    #   2 pos=0 → competição continental (UCL/UEL/…)
    #   6 pos=0 → supercopa continental (UEFA Super Cup, etc.)
    #   11 pos=0 → Copa do Mundo
    for row in cur.execute(
        "SELECT id, result_type, specific_id, pos FROM historic_results2 WHERE id_type=2"
    ):
        pid, rt, sid, pos = si(row[0]), si(row[1]), si(row[2]), si(row[3])
        t = titles[pid]
        # Títulos (pos=0 vitória, pos=1 às vezes 2º; liga usa pos=1 campeão):
        #   rt0 pos=1 → liga
        #   rt7 pos=0 → copa nacional
        #   rt1 pos=0 → supercopa nacional (separado da copa)
        #   rt3/4 pos=0 → supercopa nacional (outros países)
        #   rt2 pos=0 → competição continental (UCL/UEL)
        #   rt6 pos=0 → supercopa continental
        #   rt11 pos=0 → Copa do Mundo
        if rt == 0 and pos == 1:
            t["league"] += 1
        elif rt == 7 and pos == 0:
            t["national_cup"] += 1
        elif rt == 1 and pos == 0:
            t["super_national"] += 1
        elif rt in (3, 4) and pos == 0:
            t["super_national"] += 1
        elif rt == 2 and pos == 0:
            t["continental"] += 1
        elif rt == 6 and pos == 0:
            t["super_continental"] += 1
        elif rt == 11 and pos == 0:
            t["world_cup"] += 1
        t["total"] = (
            t["league"] + t["national_cup"] + t["super_national"]
            + t["super_continental"] + t["continental"] + t["world_cup"]
        )

    def season_count(dates, played):
        """Aproxima temporadas com jogos (segmentos com played>0)."""
        return max(1, sum(1 for p in played if p > 0))

    def one_player(row, status: str):
        played = parse_csv_floats(row["h_played"] if "h_played" in row.keys() else "")
        goals = parse_csv_floats(row["h_goals"] if "h_goals" in row.keys() else "")
        assists = parse_csv_floats(row["h_assists"] if "h_assists" in row.keys() else "")
        rating = parse_csv_floats(row["h_rating"] if "h_rating" in row.keys() else "")
        team_ids = parse_csv_floats(row["h_team_id"] if "h_team_id" in row.keys() else "")
        dates = parse_csv_floats(row["h_date"] if "h_date" in row.keys() else "")
        cleans = parse_csv_floats(row["h_clean_sheets"] if "h_clean_sheets" in row.keys() else "")
        is_nat = parse_csv_floats(row["h_is_national"] if "h_is_national" in row.keys() else "")

        n = max(len(played), len(goals), len(assists), len(rating), len(team_ids), len(dates), len(cleans), len(is_nat), 1)

        def pad(a):
            a = list(a)
            return a + [0.0] * (n - len(a))

        played, goals, assists, rating, team_ids, dates, cleans, is_nat = map(
            pad, [played, goals, assists, rating, team_ids, dates, cleans, is_nat]
        )

        # Seleção: team_id == -1 com jogos.
        # h_is_national às vezes vem com comprimento diferente da série (desalinha e marca clube como seleção).
        # Padrão estável no save: segmentos de seleção usam team_id=-1; livre sem jogos fica 0.
        nat_games = nat_goals = nat_assists = 0.0
        nat_rating_sum = nat_rating_w = 0.0
        club_games = club_goals = club_assists = 0.0
        rating_sum = rating_w = 0.0
        for i in range(n):
            p, g, a, r = played[i], goals[i], assists[i], rating[i]
            national = int(team_ids[i]) == -1 and p > 0
            if national:
                nat_games += p
                nat_goals += g
                nat_assists += a
                if p > 0 and r > 0:
                    nat_rating_sum += r * p
                    nat_rating_w += p
            else:
                club_games += p
                club_goals += g
                club_assists += a
            if p > 0 and r > 0:
                rating_sum += r * p
                rating_w += p

        total_games = sum(played)
        total_goals = sum(goals)
        total_assists = sum(assists)
        total_ga = total_goals + total_assists
        total_clean = sum(cleans)
        seasons = season_count(dates, played)

        # Overall history
        ov_hist = parse_values(row["value_in_time"] if "value_in_time" in row.keys() else "")
        ov_start = round(ov_hist[0], 1) if ov_hist else None
        ov_peak = round(max(ov_hist), 1) if ov_hist else None
        if status == "Ativo":
            ov_now = round(sf(row["temp_value"] if "temp_value" in row.keys() else None) or (ov_hist[-1] if ov_hist else 0), 1)
            if row["role"] == "GK":
                ov_now = round(sf(row["GKP"] if "GKP" in row.keys() else ov_now) or ov_now, 1)
            mk = si(row["temp_mk_value_modded"] if "temp_mk_value_modded" in row.keys() else 0) or si(row["temp_mk_value"] if "temp_mk_value" in row.keys() else 0)
        else:
            ov_now = round(ov_hist[-1], 1) if ov_hist else None
            mk = None  # aposentado

        # Idade
        bday = sf(row["birth_date"] if "birth_date" in row.keys() else 0)
        age = round((current_day - bday) / 365.25, 1) if bday else None

        levels = parse_role_levels(row["roles_level"] if "roles_level" in row.keys() else "")
        role_cols = [round(float(levels.get(role, 0.0)), 1) for role in OUTFIELD_ROLES]

        # Rewards
        rw = parse_reward_items(row["h_rewards"] if "h_rewards" in row.keys() else "")
        totm = {
            "toy_league": 0, "toy_cup": 0, "sel_cont": 0, "sel_wc": 0, "toy_total": 0,
            "mvp_league": 0, "mvp_cup": 0, "mvp_cont": 0, "mvp_wc": 0, "mvp_total": 0,
            "art_league": 0, "art_cup": 0, "art_cont": 0, "art_wc": 0,
            "ast_league": 0, "ast_cup": 0, "ast_cont": 0, "ast_wc": 0,
            "gk_league": 0, "gk_cup": 0, "gk_cont": 0, "gk_wc": 0,
            "golden_boot": 0, "golden_glove": 0, "golden_boy": 0, "ballon": 0,
        }
        # UI Footlord (aba Conquistas):
        #   tipo 6 = MVP de PARTIDA (LaLiga x27, UCL x12…) — NÃO é time do ano
        #   tipo 5 = Time do ano / Seleção do torneio (TOP 11)
        #   tipo 3 global = Golden Boy; tipo 3 em competição = Jovem do ano
        #   tipo 2 global = Bola de Ouro
        #   tipo 0 = Artilheiro; tipo 1 = Maior assistência; tipo 4 = Melhor goleiro
        for typ, comp, cnt in rw:
            b = reward_bucket(comp)
            if typ == 6:  # MVP de partida
                totm["mvp_total"] += cnt
                if b == "league":
                    totm["mvp_league"] += cnt
                elif b == "national_cup":
                    totm["mvp_cup"] += cnt
                elif b == "continental":
                    totm["mvp_cont"] += cnt
                elif b in ("wc", "wcq"):
                    totm["mvp_wc"] += cnt
            elif typ == 5:  # Time do ano / Seleção do torneio
                totm["toy_total"] += cnt
                if b == "league":
                    totm["toy_league"] += cnt
                elif b == "national_cup":
                    totm["toy_cup"] += cnt
                elif b == "continental":
                    totm["sel_cont"] += cnt
                elif b == "wc":
                    totm["sel_wc"] += cnt
                # global tipo 5 = time do ano mundial → só no total
            elif typ == 0:
                if b == "global":
                    totm["golden_boot"] += cnt
                elif b == "league":
                    totm["art_league"] += cnt
                elif b == "national_cup":
                    totm["art_cup"] += cnt
                elif b == "continental":
                    totm["art_cont"] += cnt
                elif b == "wc":
                    totm["art_wc"] += cnt
            elif typ == 1:
                if b == "league":
                    totm["ast_league"] += cnt
                elif b == "national_cup":
                    totm["ast_cup"] += cnt
                elif b == "continental":
                    totm["ast_cont"] += cnt
                elif b == "wc":
                    totm["ast_wc"] += cnt
            elif typ == 4:
                if b == "global":
                    totm["golden_glove"] += cnt
                elif b == "league":
                    totm["gk_league"] += cnt
                elif b == "national_cup":
                    totm["gk_cup"] += cnt
                elif b == "continental":
                    totm["gk_cont"] += cnt
                elif b == "wc":
                    totm["gk_wc"] += cnt
            elif typ == 3:
                if b == "global":
                    totm["golden_boy"] += cnt
            elif typ == 2:
                if b == "global":
                    totm["ballon"] += cnt

        ttl = titles.get(si(row["id"]), {
            "league": 0, "national_cup": 0, "super_national": 0,
            "super_continental": 0, "continental": 0, "world_cup": 0, "total": 0,
        })
        first = str(row["name"] if "name" in row.keys() else "" or "").strip()
        last = str(row["surname"] if "surname" in row.keys() else "" or "").strip()
        name = " ".join(p for p in (first, last) if p) or f"ID {row['id']}"

        def r3(v):
            return None if v is None else round(float(v), 3)

        def r1(v):
            return None if v is None else round(float(v), 1)

        return [
            si(row["id"]),
            name,
            row["role"] if "role" in row.keys() else "",
            *role_cols,
            age,
            row["nationality"] if "nationality" in row.keys() else "",
            status,
            ov_start,
            ov_peak,
            ov_now,
            int(total_games),
            int(nat_games),
            r1(total_games / seasons) if seasons else None,
            int(total_goals),
            r3(total_goals / total_games) if total_games else None,
            r1(total_goals / seasons) if seasons else None,
            int(nat_goals),
            int(total_assists),
            r3(total_assists / total_games) if total_games else None,
            r1(total_assists / seasons) if seasons else None,
            int(nat_assists),
            int(total_ga),
            r1(total_ga / seasons) if seasons else None,
            r3(total_ga / total_games) if total_games else None,
            r1(rating_sum / rating_w) if rating_w else None,
            r1(nat_rating_sum / nat_rating_w) if nat_rating_w else None,
            int(total_clean),
            r1(total_clean / seasons) if seasons else None,
            mk,
            totm["toy_league"],
            totm["toy_cup"],
            totm["sel_cont"],
            totm["sel_wc"],
            totm["toy_total"],
            totm["mvp_league"],
            totm["mvp_cup"],
            totm["mvp_cont"],
            totm["mvp_wc"],
            totm["mvp_total"],
            totm["art_league"],
            totm["art_cup"],
            totm["art_cont"],
            totm["art_wc"],
            totm["ast_league"],
            totm["ast_cup"],
            totm["ast_cont"],
            totm["ast_wc"],
            totm["gk_league"],
            totm["gk_cup"],
            totm["gk_cont"],
            totm["gk_wc"],
            ttl["total"],
            ttl["league"],
            ttl["national_cup"],
            ttl["super_national"],
            ttl["super_continental"],
            ttl["continental"],
            ttl["world_cup"],
            totm["golden_boot"],
            totm["golden_glove"],
            totm["golden_boy"],
            totm["ballon"],
        ]


    rows = []
    for row in cur.execute("SELECT * FROM players2"):
        d = dict(row)
        # pular quem não tem nenhum histórico
        if not str(d.get("h_played") or "").strip() and not str(d.get("h_rewards") or "").strip():
            continue
        # se h_played só zeros
        played = parse_csv_floats(d.get("h_played"))
        if sum(played) <= 0 and not str(d.get("h_rewards") or "").strip():
            continue
        rows.append(one_player(d, "Ativo"))

    for row in cur.execute("SELECT * FROM players_retired2"):
        d = dict(row)
        played = parse_csv_floats(d.get("h_played"))
        if sum(played) <= 0 and not str(d.get("h_rewards") or "").strip():
            continue
        rows.append(one_player(d, "Aposentado"))

    # Ordenar: overall atual desc, depois gols carreira
    def sort_key(r):
        ov = r[19] if isinstance(r[19], (int, float)) else -1  # overall atual index
        goals = r[23] if isinstance(r[23], (int, float)) else 0
        return (-(ov or 0), -(goals or 0))

    # indices: 0 id, 1 name, ... 16 situação, 17 ov_start, 18 ov_peak, 19 ov_now, 20 games, 23 goals
    rows.sort(key=lambda r: (
        -(r[19] if isinstance(r[19], (int, float)) else -1),
        -(r[23] if isinstance(r[23], (int, float)) else 0),
    ))
    return rows


def extract_data(save_path: Path):
    save_path = Path(save_path)
    with tempfile.TemporaryDirectory() as td:
        if save_path.suffix.lower() == ".db":
            db_path = save_path
        else:
            with zipfile.ZipFile(save_path) as z:
                name = next(
                    (n for n in z.namelist() if re.fullmatch(r"save_.*\.db", Path(n).name)),
                    None,
                )
                if name is None:
                    name = next(n for n in z.namelist() if n.lower().endswith(".db") and not Path(n).name.startswith("temp_"))
                db_path = Path(z.extract(name, td))
        conn = sqlite3.connect(db_path)
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        emit(10, "Lendo configurações do save")
        date_info = get_date_info(cur)
        current_day = date_info["current_day"]
        managed_id = date_info["managed_id"]
        managed_name = date_info["managed_name"]

        # Nome do arquivo costuma ser "Time_-_data.fl" — útil quando team_id do general é fantasma
        if managed_name in ("", "Clube") or managed_id <= 0:
            stem = Path(save_path).stem  # ex: Londrina_-_30_Jun_2030
            guess = stem.split("_-_")[0].replace("_", " ").strip()
            if guess:
                row = cur.execute(
                    "SELECT id, name FROM teams2 WHERE name=? COLLATE NOCASE AND COALESCE(is_national_team,0)=0 LIMIT 1",
                    (guess,),
                ).fetchone()
                if not row:
                    row = cur.execute(
                        "SELECT id, name FROM teams2 WHERE name LIKE ? AND COALESCE(is_national_team,0)=0 LIMIT 1",
                        (f"%{guess}%",),
                    ).fetchone()
                if row:
                    managed_id, managed_name = si(row[0]), row[1]
                    date_info["managed_id"] = managed_id
                    date_info["managed_name"] = managed_name
        
        leagues = {row["id"]: dict(row) for row in cur.execute("SELECT * FROM leagues2").fetchall()}
        champs = {row["id"]: dict(row) for row in cur.execute("SELECT * FROM champs2").fetchall()}
        teams_meta = {row["id"]: dict(row) for row in cur.execute("SELECT id, name, champ_id, nation_id FROM teams2").fetchall()}
        teams = {tid: meta["name"] for tid, meta in teams_meta.items()}

        def resolve_league(team_id, player_champ_id=0):
            """Resolve liga/divisão via time.champ_id → champs → leagues (player.champ_id costuma ser 0)."""
            meta = teams_meta.get(si(team_id), {})
            champ_id = si(meta.get("champ_id"), 0) or si(player_champ_id, 0)
            champ = champs.get(champ_id, {})
            league_id = si(champ.get("league_id"), -1)
            league = leagues.get(league_id, {})
            name = league.get("name") or champ.get("name") or "Sem liga"
            division = league.get("division")
            return name, division

        # Se clube gerenciado ainda genérico, tenta pelo nome mais frequente no save path não disponível aqui;
        # reforço: se managed_id inválido, usa time do manager se aparecer depois.
        emit(25, "Lendo jogadores")
        players = {}
        for row in cur.execute("SELECT * FROM players2").fetchall():
            player = dict(row)
            tid = si(player.get("team_id"), -1)
            player["team_name"] = teams.get(tid, "Sem clube")
            league_name, league_div = resolve_league(tid, player.get("champ_id"))
            player["league_name"] = league_name
            player["league_division"] = league_div
            player["role"] = player.get("role") or "N/A"
            player["group"] = "GK" if player["role"] == "GK" else "DEF" if player["role"] in ("LB", "RB", "CB", "LWB", "RWB", "CD") else "MID" if player["role"] in ("CM", "LM", "RM", "CDM", "CAM") else "FWD"
            player["quality"] = sf(player.get("GKP")) if player["role"] == "GK" else sf(player.get("temp_value"))
            player["overall"] = player["quality"]
            
            # s_matches costuma vir vazio; o volume da temporada está em s_played (CH_6:38,WC_0:9,...)
            player["games"] = parse_count_blocks(player.get("s_played")) or parse_count_blocks(player.get("s_matches"))
            player["goals"] = parse_count_blocks(player.get("s_goals"))
            player["assists"] = parse_count_blocks(player.get("s_assists"))
            player["raw_clean"] = parse_count_blocks(player.get("s_clean_sheets"))
            player["raw_conceded"] = parse_count_blocks(player.get("s_g_conceded"))
            
            # Avaliação da temporada (s_ratings) com fallback carreira (h_rating)
            season_ratings = parse_rating_blocks(player.get("s_ratings"))
            career_ratings = parse_rating_blocks(player.get("h_rating"))
            ratings_list = season_ratings or career_ratings
            player["raw_rating"] = sum(ratings_list) / len(ratings_list) if ratings_list else 0.0
            player["rating_count"] = len(ratings_list)
            
            levels = parse_role_levels(player.get("roles_level"))
            player["role_scores"] = {role: round(float(levels.get(role, 0.0)), 1) for role in OUTFIELD_ROLES}
            player["season_overall_history"] = parse_values(player.get("s_value_in_time"))
            player["career_overall_history"] = parse_values(player.get("value_in_time"))
            player["yellow_cards"] = parse_count_blocks(player.get("s_y_cards"))
            player["red_cards"] = parse_count_blocks(player.get("s_r_cards"))
            player["market_value"] = si(player.get("temp_mk_value_modded")) or si(player.get("temp_mk_value"))
            player["salary"] = si(player.get("salary"))
            
            # Correção de Idade: birth_date é dias desde o início? Não, costuma ser dias relativos.
            # No save auditado: birth_date ~ -8707. 8707 / 365.25 = 23.8 anos.
            # Idade = (current_day - birth_date) / 365.25
            bday = sf(player.get("birth_date"))
            player["age"] = round((current_day - bday) / 365.25, 1)
            
            exp = si(player.get("contract_expiration"))
            player["contract_years"] = max(0.0, (exp - current_day) / 365.0) if exp else None
            player.update({
                "mvp": 0, "worst": 0, "rating_matches": 0, "rating_sum": 0.0, "apps": 0, "minutes": 0.0, 
                "clean_sheets": 0.0, "gk_apps": 0, "gk_minutes": 0.0, "gk_clean": 0.0, "gk_xg": 0.0, "gk_goals": 0.0,
                "pen_taken": 0, "pen_scored": 0, "pen_missed_total": 0, "pen_missed_out": 0, "pen_missed_saved": 0, "xgp_sum": 0.0,
                "gk_pen_faced": 0, "gk_pen_saved": 0, "gk_xgp_sum": 0.0,
                "dec_goals": 0, "sav_goals": 0, "dec_ast": 0, "sav_ast": 0, "hat_tricks": 0,
                "sub_apps": 0, "sub_goals": 0, "sub_dec_goals": 0, "sub_sav_goals": 0,
                "sub_ast": 0, "sub_dec_ast": 0, "sub_sav_ast": 0, "sub_mvp": 0,
            })
            players[si(player["id"])] = player
        emit(38, "Calculando minutos, ratings e xG (temporada atual)")
        season_start = date_info.get("season_start_day", 0)
        all_matches = cur.execute(
            "SELECT * FROM matches2 WHERE date>=? AND date<=? ORDER BY date,id",
            (season_start, current_day),
        ).fetchall()
        played_matches = [
            match for match in all_matches
            if si(match["state"], -1) == 2 and match["goals_1"] is not None and match["goals_2"] is not None
        ]
        team_cards = defaultdict(lambda: {"yellow": 0, "red": 0})
        for raw in played_matches:
            match = dict(raw)
            for side in (1, 2):
                team_id = si(match.get(f"team_{side}_id"), -1)
                if team_id > 0:
                    team_cards[team_id]["yellow"] += max(0, si(match.get(f"yellows_{side}")))
                    team_cards[team_id]["red"] += max(0, si(match.get(f"reds_{side}")))
        club_summary = {"matches": 0, "shots": 0, "goals": 0, "xga": 0.0, "ga": 0}
        for raw in played_matches:
            match = dict(raw)
            sides = []
            for side in (1, 2):
                other = 3 - side
                lineup, entered, subbed = parse_ids(match.get(f"lineup_{side}")), parse_ids(match.get(f"entered_{side}")), parse_ids(match.get(f"subbed_{side}"))
                events = parse_events(match.get(f"events_{side}"))
                other_events = parse_events(match.get(f"events_{other}"))
                minutes = estimate_minutes(lineup, entered, subbed, events)
                ratings = parse_match_ratings(match.get(f"ratings_{side}"))
                if ratings:
                    high, low = max(ratings.values()), min(ratings.values())
                    sub_on_ids = {pid for (m, code, pid) in events if code == 5}
                    for player_id, rating in ratings.items():
                        if player_id in players:
                            players[player_id]["rating_matches"] += 1
                            players[player_id]["rating_sum"] += rating
                            is_mvp = rating == high
                            players[player_id]["mvp"] += is_mvp
                            players[player_id]["worst"] += rating == low
                            if is_mvp and player_id in sub_on_ids:
                                players[player_id]["sub_mvp"] += 1
                sides.append({
                    "side": side, "team": si(match.get(f"team_{side}_id")), "minutes": minutes,
                    "ratings": ratings, "goals": si(match.get(f"goals_{side}")),
                    "xg": sf(match.get(f"x_goals_{side}")), "events": events, "other_events": other_events,
                    "lineup": lineup, "entered": entered,
                })
            
            # Gols decisivos/salvadores, hat-tricks, reservas (uma vez por partida)
            try:
                s1 = next(s for s in sides if s["side"] == 1)
                s2 = next(s for s in sides if s["side"] == 2)
                analyzed = analyze_match_goals(
                    s1["events"], s2["events"],
                    s1["goals"], s2["goals"],
                    s1["lineup"], s2["lineup"],
                    s1["entered"], s2["entered"],
                )
                for pid, st in analyzed.items():
                    if pid not in players:
                        continue
                    p = players[pid]
                    p["dec_goals"] += st["dec_goals"]
                    p["sav_goals"] += st["sav_goals"]
                    p["dec_ast"] += st["dec_ast"]
                    p["sav_ast"] += st["sav_ast"]
                    p["hat_tricks"] += st["hat"]
                    p["sub_apps"] += st["sub_apps"]
                    p["sub_goals"] += st["sub_goals"]
                    p["sub_dec_goals"] += st["sub_dec_goals"]
                    p["sub_sav_goals"] += st["sub_sav_goals"]
                    p["sub_ast"] += st["sub_ast"]
                    p["sub_dec_ast"] += st["sub_dec_ast"]
                    p["sub_sav_ast"] += st["sub_sav_ast"]
            except Exception:
                pass

            # Pênaltis
            for side_idx, side in enumerate(sides):
                other = sides[1 - side_idx]
                # Códigos 7/8 = pênalti convertido/perdido; inclui desempate (minuto > 119)
                game_pens = [e for e in side["events"] if e[1] in (7, 8)]
                if not game_pens: continue

                other_side_id = other["side"]
                other_lineup = id_set(match.get(f"lineup_{other_side_id}"))
                other_events = other.get("events", [])

                def get_active_gk(minute):
                    # Identifica quem era o goleiro em campo no minuto dado
                    # 1. Quem começou o jogo como GK?
                    initial_gk = next((pid for pid in other_lineup if pid in players and players[pid]["role"] == "GK"), None)
                    # 2. Houve substituição envolvendo GK antes/nesse minuto?
                    # Código 6 = sai, Código 5 = entra
                    current_gk = initial_gk
                    gk_subs = sorted([e for e in other_events if e[1] in (5, 6) and e[2] in players and players[e[2]]["role"] == "GK"])
                    for m, code, pid in gk_subs:
                        if m > minute: break
                        if code == 6: current_gk = None
                        elif code == 5: current_gk = pid
                    # 3. O goleiro atual foi expulso antes/nesse minuto?
                    # Código 3 = vermelho direto, Código 4 = segundo amarelo
                    if current_gk:
                        reds = [e for e in other_events if e[1] in (3, 4) and e[2] == current_gk and e[0] <= minute]
                        if reds: current_gk = None
                    return current_gk

                for minute, code, pid in game_pens:
                    if pid not in players: continue
                    gk_id = get_active_gk(minute)
                    
                    b_ov = players[pid].get("SHO", 50.0)
                    g_ov = players[gk_id].get("GKP", 50.0) if gk_id else 50.0
                    diff = g_ov - b_ov
                    xgp = 0.750 - (diff * 0.015)
                    xgp = max(0.01, min(0.99, xgp))
                    
                    players[pid]["pen_taken"] += 1
                    saved = any(e[1] == 9 and e[0] == minute for e in other.get("events", []))
                    if code == 7:
                        players[pid]["pen_scored"] += 1
                        # Reconciliar: se o save não contou o gol (bug do save), nós contamos
                        # E garantimos que 'gols sem pênalti' não fique negativo
                        players[pid]["goals"] = max(players[pid]["goals"], players[pid]["pen_scored"])
                    else:
                        players[pid]["pen_missed_total"] += 1
                        if saved:
                            players[pid]["pen_missed_saved"] += 1
                        else:
                            players[pid]["pen_missed_out"] += 1
                    players[pid]["xgp_sum"] += xgp
                    
                    if gk_id:
                        if code == 7 or saved:
                            players[gk_id]["gk_pen_faced"] += 1
                            players[gk_id]["gk_xgp_sum"] += xgp
                            if saved:
                                players[gk_id]["gk_pen_saved"] += 1

            for side in sides:
                other = sides[1] if side["side"] == 1 else sides[0]
                player_ids = set(side["minutes"]) | set(side["ratings"])
                total_gk_minutes = sum(side["minutes"].get(player_id, 0.0) for player_id in player_ids if player_id in players and players[player_id]["role"] == "GK")
                for player_id in player_ids:
                    if player_id not in players:
                        continue
                    minutes = side["minutes"].get(player_id, 45.0 if player_id in side["ratings"] else 0.0)
                    if minutes <= 0:
                        continue
                    player = players[player_id]
                    player["apps"] += 1
                    player["minutes"] += minutes
                    if other["goals"] == 0 and player["group"] == "DEF":
                        player["clean_sheets"] += 1
                    if player["role"] == "GK":
                        player["gk_apps"] += 1
                        player["gk_minutes"] += minutes
                        player["gk_clean"] += other["goals"] == 0
                        share = minutes / total_gk_minutes if total_gk_minutes else 1.0
                        player["gk_xg"] += other["xg"] * share
                        player["gk_goals"] += other["goals"] * share
            if si(match.get("team_1_id")) == managed_id or si(match.get("team_2_id")) == managed_id:
                side = 1 if si(match.get("team_1_id")) == managed_id else 2
                other = 3 - side
                club_summary["matches"] += 1
                club_summary["shots"] += si(match.get(f"shots_on_{side}")) + si(match.get(f"shots_off_{side}"))
                club_summary["goals"] += si(match.get(f"goals_{side}"))
                club_summary["xga"] += sf(match.get(f"x_goals_{other}"))
                club_summary["ga"] += si(match.get(f"goals_{other}"))
        emit(56, "Calculando overall Footlord")
        for player in players.values():
            if player["minutes"] <= 0 and player["games"]:
                player["minutes"], player["apps"] = player["games"] * 45.0, int(player["games"])
            if player["rating_matches"] == 0 and player["rating_count"]:
                player["rating_matches"], player["rating_sum"] = player["rating_count"], player["raw_rating"] * player["rating_count"]
            player["rating"] = player["rating_sum"] / player["rating_matches"] if player["rating_matches"] else player["raw_rating"]
            player["ga"] = player["goals"] + player["assists"]
            player["goals_pg"] = player["goals"] / player["games"] if player["games"] else 0.0
            player["assists_pg"] = player["assists"] / player["games"] if player["games"] else 0.0
            player["ga_pg"] = player["ga"] / player["games"] if player["games"] else 0.0
            player["minutes_goal"] = player["minutes"] / player["goals"] if player["goals"] else None
            player["minutes_assist"] = player["minutes"] / player["assists"] if player["assists"] else None
            player["mvp_pct"] = player["mvp"] / player["rating_matches"] if player["rating_matches"] else 0.0
            player["worst_pct"] = player["worst"] / player["rating_matches"] if player["rating_matches"] else 0.0
            if player["group"] == "DEF":
                player["clean_sheets"] = max(player["clean_sheets"], player["raw_clean"])
            # Sem partidas no DB (início de temporada / só fixtures futuras): usar s_*
            if player["apps"] <= 0 and player["games"] > 0:
                player["apps"] = int(player["games"])
                if player["minutes"] <= 0:
                    player["minutes"] = player["apps"] * 90.0
            if not player["rating"] and player["raw_rating"]:
                player["rating"] = player["raw_rating"]
                player["rating_matches"] = max(player["rating_count"], player["apps"], 1)
            if player["role"] == "GK" and not player["gk_apps"] and player["games"]:
                player["gk_apps"] = int(player["games"])
                player["gk_minutes"] = player["minutes"] if player["minutes"] else player["gk_apps"] * 90.0
                player["gk_clean"] = player["raw_clean"]
                player["gk_goals"] = player["raw_conceded"]
            if player["role"] == "GK":
                player["output"] = (player["gk_xg"] - player["gk_goals"]) / (player["gk_minutes"] / 90) if player["gk_minutes"] else 0.0
            elif player["group"] == "DEF":
                player["output"] = 0.65 * (player["clean_sheets"] / player["games"] if player["games"] else 0) + 0.35 * min(1, player["ga_pg"])
            else:
                player["output"] = min(1, player["ga_pg"])
            player["rating_score"] = clip((player["rating"] - 5.0) / 3 * 100) if player["rating"] else 0
            player["efficiency"] = player["quality"] / (1 + max(0, player["market_value"]) / 1_000_000 + 0.5 * max(0, player["salary"]) / 1_000_000)
        group_outputs = defaultdict(list)
        for player in players.values():
            if player["games"] or player["role"] == "GK":
                group_outputs[player["group"]].append(player["output"])
        efficiencies = [player["efficiency"] for player in players.values()]
        for player in players.values():
            output_score, cost_score = percentile(group_outputs[player["group"]], player["output"]), percentile(efficiencies, player["efficiency"])
            age_val = sf(player.get("age"), 34)
            age_score = clip((34 - age_val) / 16 * 100)
            player["score"] = clip(0.50 * player["quality"] + 0.20 * player["rating_score"] + 0.15 * output_score + 0.10 * cost_score + 0.05 * age_score)
        ordered_players = sorted(players.values(), key=lambda player: (-player["overall"], player["market_value"], player["id"]))
        goalkeepers = sorted(
            [player for player in players.values() if player["role"] == "GK" and (player["gk_apps"] or player["overall"] >= 60)],
            key=lambda player: (-player["overall"], -player["gk_apps"], player["id"]),
        )
        team_rows = build_team_rows(cur, date_info, players, played_matches, team_cards)
        emit(90, "Montando aba Carreira (ativos + aposentados)")
        gen_row = dict(cur.execute("SELECT * FROM general LIMIT 1").fetchone())
        career_rows = build_career_rows(cur, current_day, starting_year=si(gen_row.get("starting_year"), 2025))
        conn.close()
        return {
            "current_day": current_day,
            "club": managed_name,
            "players": ordered_players,
            "gks": goalkeepers,
            "teams": team_rows,
            "club_summary": club_summary,
            "played_matches": len(played_matches),
            "career": career_rows,
        }


def player_values(players):
    rows = []
    for rank, player in enumerate(players, 1):
        name = " ".join(part for part in (str(player.get("name") or "").strip(), str(player.get("surname") or "").strip()) if part) or f"ID {player['id']}"
        current = round(player["overall"], 1)
        season_start = player["season_overall_history"][0] if player["season_overall_history"] else current
        career_start = player["career_overall_history"][0] if player["career_overall_history"] else season_start
        # Usar apps (reconstruído das partidas) se games (bruto do save) estiver zerado
        games_val = player["apps"] if player["apps"] > 0 else player["games"]
        games, yellow, red = int(round(games_val)), int(round(player["yellow_cards"])), int(round(player["red_cards"]))
        
        pen_taken = player["pen_taken"]
        pen_scored = player["pen_scored"]
        pen_missed_total = player["pen_missed_total"]
        pen_missed_out = player["pen_missed_out"]
        xgp = player["xgp_sum"]
        pen_rate = pen_scored / pen_taken if pen_taken else None
        pen_vs_xgp = pen_scored - xgp if pen_taken else None
        
        role_cols = [player.get("role_scores", {}).get(role, 0.0) for role in OUTFIELD_ROLES]
        # Decimais no máximo milésimo (3 casas)
        def r3(v):
            return None if v is None else round(float(v), 3)
        def r1(v):
            return None if v is None else round(float(v), 1)
        rows.append([
            rank, name, player["role"], *role_cols, player["age"], player.get("team_name") or "Sem clube", player.get("league_name") or "Sem liga", player.get("league_division"), player.get("nationality") or player.get("nation_name") or "",
            r1(current), r1(season_start), r1(current - season_start), r1(career_start), r1(current - career_start), games, yellow, red, yellow + red, r3(yellow / games) if games else None, r3(red / games) if games else None,
            int(round(player["goals"])), int(round(player["goals"] - pen_scored)), int(pen_taken), int(pen_scored), int(pen_missed_total), int(pen_missed_out), r3(pen_rate), r3(xgp), r3(pen_vs_xgp),
            int(round(player["assists"])), int(round(player["ga"])),
            r3(player["goals"] / games) if games else 0.0,
            r3(player["assists"] / games) if games else 0.0,
            r3(player["ga"] / games) if games else 0.0,
            int(player.get("dec_goals") or 0), int(player.get("sav_goals") or 0),
            int(player.get("dec_ast") or 0), int(player.get("sav_ast") or 0),
            int(player.get("hat_tricks") or 0),
            int(player.get("sub_apps") or 0), int(player.get("sub_goals") or 0),
            int(player.get("sub_dec_goals") or 0), int(player.get("sub_sav_goals") or 0),
            int(player.get("sub_ast") or 0), int(player.get("sub_dec_ast") or 0),
            int(player.get("sub_sav_ast") or 0), int(player.get("sub_mvp") or 0),
            r1(player["rating"]) if player["rating"] else None,
            int(round(player["minutes"])) if player.get("minutes") else 0,
            r1(player["minutes"] / player["apps"]) if player["apps"] else None,
            r1(player["minutes_goal"]) if player.get("minutes_goal") else None,
            r1(player["minutes_assist"]) if player.get("minutes_assist") else None,
            int(player.get("mvp") or 0), int(player.get("worst") or 0),
            r3(player["mvp_pct"]), r3(player["worst_pct"]),
            int(round(player["clean_sheets"])) if player["group"] == "DEF" else None,
            int(player["market_value"]) if player.get("market_value") is not None else None,
            int(player["salary"]) if player.get("salary") is not None else None,
            r1(player["contract_years"]) if player.get("contract_years") is not None else None,
            "Sem clube" if si(player.get("team_id"), -1) <= 0 else ("Em empréstimo" if si(player.get("loan_status")) else "Contrato"),
            "Sim" if si(player.get("loan_status")) else "Não",
            "Sim" if si(player.get("transfer_status")) == 1 else "Não",
            r1(sf(player.get("PAC"))), r1(sf(player.get("SHO"))), r1(sf(player.get("PAS"))),
            r1(sf(player.get("DEF"))), r1(sf(player.get("PHY"))), r1(sf(player.get("MEN"))),
            r1(sf(player.get("GKP"))), player["id"],
        ])
    return rows


def goalkeeper_values(goalkeepers):
    rows = []
    for rank, player in enumerate(goalkeepers, 1):
        xg, conceded, apps, minutes = player["gk_xg"], player["gk_goals"], player["gk_apps"], player["gk_minutes"]
        prevented = xg - conceded
        rate = prevented / xg if xg else None
        
        pen_faced = player["gk_pen_faced"]
        pen_saved = player["gk_pen_saved"]
        xgp = player["gk_xgp_sum"]
        pen_save_rate = pen_saved / pen_faced if pen_faced else None
        pen_vs_xgp = pen_saved - (pen_faced - xgp) if pen_faced else None # xGP é gol esperado, logo defesa esperada é faced - xGP
        
        classification = "Amostra curta" if apps < 5 else "Sem xG" if rate is None else "Excelente" if rate >= .15 else "Bom" if rate >= .05 else "Neutro" if rate >= -.05 else "Crítico"
        name = " ".join(part for part in (str(player.get("name") or "").strip(), str(player.get("surname") or "").strip()) if part) or f"ID {player['id']}"
        def r3(v):
            return None if v is None else round(float(v), 3)
        def r1(v):
            return None if v is None else round(float(v), 1)
        rows.append([
            rank, name, player.get("team_name") or "Sem clube", player.get("league_name") or "Sem liga", apps, round(minutes), round(player["gk_clean"]), r3(player["gk_clean"] / apps) if apps else None, r3(xg), r3(conceded), r3(conceded / apps) if apps else None, r3(prevented), r3(prevented / apps) if apps else None, r3(prevented / (minutes / 90)) if minutes else None, r3(rate), r3(conceded / xg) if xg else None,
            pen_faced, pen_saved, r3(pen_save_rate), r3(xgp), r3(pen_vs_xgp),
            classification, "Sim" if apps >= 5 else "Não", r1(player["overall"]), player["id"]
        ])
    return rows


def percentiles_for_rows(rows, column_indices):
    result = {}
    for column in column_indices:
        values = sorted(float(row[column]) for row in rows if column < len(row) and isinstance(row[column], (int, float)) and math.isfinite(float(row[column])))
        result[column] = values
    return result


def color_name(value, values, inverse=False):
    if value is None or not isinstance(value, (int, float)) or not values:
        return "neutral"
    lower, medium, high = values[max(0, int((len(values) - 1) * .20))], values[max(0, int((len(values) - 1) * .50))], values[max(0, int((len(values) - 1) * .80))]
    if inverse:
        return "excellent" if value <= lower else "good" if value <= medium else "medium" if value <= high else "bad"
    return "bad" if value <= lower else "medium" if value <= medium else "good" if value <= high else "excellent"


def set_cell_style(cell, fill="neutral", number_format=None, alignment=CENTER, font=DATA_FONT):
    cell.font, cell.fill, cell.border, cell.alignment = font, FILLS[fill], BORDER, alignment
    if number_format:
        cell.number_format = number_format

def write_sheet(wb, title, subtitle, headers, rows, number_formats, color_columns=None, inverse_columns=None):
    ws = wb.create_sheet(title)
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A5"
    # Título e Subtítulo
    ws.append([title.upper()])
    set_cell_style(ws.cell(row=1, column=1), fill="title", font=TITLE_FONT, alignment=LEFT_WRAP)
    ws.append([subtitle])
    set_cell_style(ws.cell(row=2, column=1), fill="subtitle", font=HEADER_FONT, alignment=LEFT_WRAP)
    
    # Legenda
    legend = ["LEGENDA:", "RUIM", "MÉDIO", "BOM", "EXCELENTE"]
    legend_fills = ["neutral", "bad", "medium", "good", "excellent"]
    ws.append(legend)
    for i, f in enumerate(legend_fills, 1):
        set_cell_style(ws.cell(row=3, column=i), fill=f)
    
    # Cabeçalhos (wrap só no header)
    ws.append(headers)
    for i in range(1, len(headers) + 1):
        set_cell_style(ws.cell(row=4, column=i), fill="header", font=HEADER_FONT, alignment=CENTER_WRAP)
        
    # Larguras — fileiras finas: colunas Role estreitas, dados sem wrap
    for i, h in enumerate(headers, 1):
        h_str = str(h or "").strip()
        h_low = h_str.casefold()
        if h_str.startswith("Role ") or h_str.startswith("ROLE "):
            w = 7
        elif any(x in h_str for x in ("Jogador", "Player", "Goleiro", "Goalkeeper")):
            w = 22
        elif any(x in h_str for x in ("Time", "Team")):
            w = 18
        elif any(x in h_str for x in ("Liga", "League", "Adversário", "Opponent", "Para quem")):
            w = 18
        elif "€" in h_str:
            w = 14
        elif any(x in h_str for x in ("Overall", "overall", "xGP")):
            w = 12
        else:
            w = 11
        ws.column_dimensions[get_column_letter(i)].width = w

    # Dados (wrap_text=False via CENTER → fileiras finas)
    percentiles = percentiles_for_rows(rows, color_columns or [])
    for row_idx, row_data in enumerate(rows, 5):
        ws.append(row_data)
        for col_idx, val in enumerate(row_data, 1):
            fill = "neutral"
            if color_columns and (col_idx - 1) in color_columns:
                fill = color_name(val, percentiles.get(col_idx - 1), inverse=(inverse_columns and (col_idx - 1) in inverse_columns))
            set_cell_style(ws.cell(row=row_idx, column=col_idx), fill=fill, number_format=number_formats.get(col_idx - 1), alignment=CENTER)
            
    last = max(4, 4 + len(rows))
    ws.auto_filter.ref = f"A4:{get_column_letter(len(headers))}{last}"


def create_workbook(data, output_path, language="pt"):
    wb = Workbook()
    if "Sheet" in wb.sheetnames:
        wb.remove(wb["Sheet"])
    
    # Players 2
    # Índices após expandir Role CD..ST (11 colunas no lugar de 1):
    # 0 Rank, 1 Nome, 2 Pos, 3-13 Roles, 14 Idade, 15 Time, 16 Liga, 17 Div, 18 País,
    # 19 Overall, 20 ov_temp, 21 var_temp, 22 ov_car, 23 var_car,
    # 24 jogos, 25 Y, 26 R, 27 tot, 28 Y/p, 29 R/p,
    # 30 gols, 31 gols_sem_pen, 32 pen_bat, 33 gols_pen, 34 pen_perd, 35 pen_fora, 36 taxa_pen, 37 xgp, 38 pen_vs_xgp,
    # 39 ast, 40 GA, 41 G/p, 42 A/p, 43 GA/p, 44 média, 45 min,
    # 46 min/p, 47 min/gol, 48 min/ast, 49 MVP, 50 pior, 51 MVP%, 52 pior%, 53 clean,
    # 54 mk, 55 sal, 56 contr, 57 status, 58 loan, 59 list, 60-66 attrs, 67 id
    p_rows = player_values(data["players"])
    # Índices alinhados a PLAYERS_HEADERS (81 cols)
    # Inteiros (contagens): jogos, cartões, gols, pens, assists, impacto, sub, MVP, clean, €, ID
    # Decimais: roles, overalls, rates, rating, min/p, attrs
    p_fmts = {
        **{i: NUM_FORMAT_DEC1 for i in range(3, 14)},  # roles 0-10
        14: NUM_FORMAT_DEC1,  # idade
        **{i: NUM_FORMAT_DEC1 for i in range(19, 24)},  # overalls
        24: "0",  # jogos
        25: "0", 26: "0", 27: "0",  # cartões
        28: NUM_FORMAT_DEC3, 29: NUM_FORMAT_DEC3,  # cards/match
        30: "0", 31: "0", 32: "0", 33: "0", 34: "0", 35: "0",  # gols/pens counts
        36: NUM_FORMAT_PCT, 37: NUM_FORMAT_DEC3, 38: NUM_FORMAT_DEC3,
        39: "0", 40: "0",  # assists, G+A
        41: NUM_FORMAT_DEC3, 42: NUM_FORMAT_DEC3, 43: NUM_FORMAT_DEC3,  # rates
        **{i: "0" for i in range(44, 57)},  # impacto + reserva (inteiros)
        57: NUM_FORMAT_DEC1,  # rating
        58: "0",  # minutos
        59: NUM_FORMAT_DEC1, 60: NUM_FORMAT_DEC1, 61: NUM_FORMAT_DEC1,  # min rates
        62: "0", 63: "0",  # mvp / pior counts
        64: NUM_FORMAT_PCT, 65: NUM_FORMAT_PCT,
        66: "0",  # clean sheets
        67: NUM_FORMAT_INT, 68: NUM_FORMAT_INT,  # €
        69: NUM_FORMAT_DEC1,  # contrato anos
        **{i: NUM_FORMAT_DEC1 for i in range(73, 80)},  # PAC..GKP
        80: "0",  # id
    }
    # Colorir só o que compara desempenho (não nomes/IDs/status)
    p_colors = (
        list(range(3, 14))  # roles
        + [14, 19, 20, 21, 22, 23]  # idade + overalls
        + [24]  # jogos
        + list(range(25, 30))  # cartões
        + list(range(30, 44))  # gols/pens/assists/rates
        + list(range(44, 57))  # impacto + reserva
        + [57, 58, 59, 60, 61, 62, 63, 64, 65, 66]  # rating/min/mvp/clean
        + [67, 68]  # valor/salário
        + list(range(73, 80))  # atributos
    )
    # Invertido: menor é melhor
    p_inv = [
        14,  # idade (mais novo melhor no moneyball)
        25, 26, 27, 28, 29,  # cartões
        34, 35,  # pênaltis perdidos
        60, 61,  # min/gol e min/assist (menos = mais eficiente)
        63, 65,  # pior nota e %
    ]
    write_sheet(wb, "Players 2", f"Relatório de Jogadores - {data['club']} ({data['current_day']} dias)", localized_headers(PLAYERS_HEADERS, language), p_rows, p_fmts, p_colors, p_inv)
    
    # Goleiros xG
    g_rows = goalkeeper_values(data["gks"])
    g_fmts = {
        5: NUM_FORMAT_INT, 6: '0', 7: NUM_FORMAT_DEC3, 8: NUM_FORMAT_DEC3, 9: NUM_FORMAT_DEC3, 10: NUM_FORMAT_DEC3,
        11: NUM_FORMAT_DEC3, 12: NUM_FORMAT_DEC3, 13: NUM_FORMAT_DEC3, 14: NUM_FORMAT_PCT, 15: NUM_FORMAT_DEC3,
        16: '0', 17: '0', 18: NUM_FORMAT_PCT, 19: NUM_FORMAT_DEC3, 20: NUM_FORMAT_DEC3, 23: NUM_FORMAT_DEC1, 24: '0',
    }
    # Comparação: clean/p, evitados, taxa, gols/xG, taxa defesa pen, vs xGP, overall
    g_colors = [4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 23]
    g_inv = [9, 10, 15]  # gols sofridos, gols/partida, gols/xG (menor melhor)
    write_sheet(wb, "Goleiros xG", f"Desempenho de Goleiros - {data['club']}", localized_headers(GK_HEADERS, language), g_rows, g_fmts, g_colors, g_inv)
    
    # Times
    t_rows = []
    for rank, t in enumerate(data["teams"], 1):
        all_d, bef, aft = t["all"], t["before"], t["after"]
        t_rows.append([
            rank, t["name"], t["league"], t["continent"], t["players_count"], round(t["avg_age"], 1),
            all_d["games"], all_d["wins"], all_d["draws"], all_d["losses"], t["yellow_cards"], t["red_cards"], t["yellow_cards"] + t["red_cards"], round(t["yellow_cards"] / all_d["games"], 3) if all_d["games"] else None, round(t["red_cards"] / all_d["games"], 3) if all_d["games"] else None,
            bef["games"], bef["wins"], bef["draws"], bef["losses"],
            aft["games"], aft["wins"], aft["draws"], aft["losses"],
            bef["streak_w"], bef["streak_l"], aft["streak_w"], aft["streak_l"],
            all_d.get("streak_unb", 0), all_d.get("streak_wl", 0),
            all_d["goals_for"], all_d["goals_against"], round(all_d["goals_for"] / all_d["games"], 3) if all_d["games"] else None,
            round(all_d["goals_against"] / all_d["games"], 3) if all_d["games"] else None,
            all_d["shots"], round(all_d["shots"] / all_d["games"], 3) if all_d["games"] else None, all_d["on_target"], round(all_d["on_target"] / all_d["shots"], 3) if all_d["shots"] else None, round(all_d["goals_for"] / all_d["on_target"], 3) if all_d["on_target"] else None,
            # xG temporada
            round(all_d.get("xg_for", 0.0), 3),
            round(all_d.get("xg_against", 0.0), 3),
            round(all_d["xg_for"] / all_d["games"], 3) if all_d["games"] else None,
            round(all_d["xg_against"] / all_d["games"], 3) if all_d["games"] else None,
            round(all_d["goals_for"] - all_d.get("xg_for", 0.0), 3),
            round(all_d.get("xg_against", 0.0) - all_d["goals_against"], 3),
            round(all_d.get("xg_for", 0.0) - all_d.get("xg_against", 0.0), 3),
            round((all_d["goals_for"] - all_d["goals_against"]) - (all_d.get("xg_for", 0.0) - all_d.get("xg_against", 0.0)), 3),
            bef["goals_for"], bef["goals_against"], round(bef["goals_for"] / bef["games"], 3) if bef["games"] else None, round(bef["goals_against"] / bef["games"], 3) if bef["games"] else None,
            aft["goals_for"], aft["goals_against"], round(aft["goals_for"] / aft["games"], 3) if aft["games"] else None, round(aft["goals_against"] / aft["games"], 3) if aft["games"] else None,
            all_d["failed_to_score"], all_d["clean_sheets"], bef["failed_to_score"], bef["clean_sheets"], aft["failed_to_score"], aft["clean_sheets"],
            all_d["biggest_win"][0], all_d["biggest_win"][1], all_d["worst_loss"][0], all_d["worst_loss"][1],
            bef["biggest_win"][0], bef["biggest_win"][1], bef["worst_loss"][0], bef["worst_loss"][1],
            aft["biggest_win"][0], aft["biggest_win"][1], aft["worst_loss"][0], aft["worst_loss"][1],
            t["ov_init"], t["ov_aug"], t["ov_final"],
            t.get("expected_pos"), t.get("league_pos"), t.get("pos_diff"),
            t.get("opp_ov_w"), t.get("opp_ov_d"), t.get("opp_ov_l"),
            t["value"], t["salary"], t["bank"],
            t.get("jan_buy"), t.get("jan_sell"), t.get("jan_buy_ov"), t.get("jan_sell_ov"),
            t.get("aug_buy"), t.get("aug_sell"), t.get("aug_buy_ov"), t.get("aug_sell_ov"),
        ])
    t_fmts = {
        5: NUM_FORMAT_DEC1, 13: NUM_FORMAT_DEC3, 14: NUM_FORMAT_DEC3,
        27: "0", 28: "0",  # invencibilidade / sem vencer
        31: NUM_FORMAT_DEC3, 32: NUM_FORMAT_DEC3,  # gols/partida
        34: NUM_FORMAT_DEC3, 36: NUM_FORMAT_PCT, 37: NUM_FORMAT_PCT,  # chutes
        38: NUM_FORMAT_DEC3, 39: NUM_FORMAT_DEC3, 40: NUM_FORMAT_DEC3, 41: NUM_FORMAT_DEC3,  # xG totals/avg
        42: NUM_FORMAT_DEC3, 43: NUM_FORMAT_DEC3, 44: NUM_FORMAT_DEC3, 45: NUM_FORMAT_DEC3,  # xG diffs
        48: NUM_FORMAT_DEC3, 49: NUM_FORMAT_DEC3, 52: NUM_FORMAT_DEC3, 53: NUM_FORMAT_DEC3,
        72: NUM_FORMAT_DEC1, 73: NUM_FORMAT_DEC1, 74: NUM_FORMAT_DEC1,  # overalls
        75: "0", 76: "0", 77: "0",  # expectativa / posição / diff
        78: NUM_FORMAT_DEC1, 79: NUM_FORMAT_DEC1, 80: NUM_FORMAT_DEC1,  # opp overall
        81: NUM_FORMAT_INT, 82: NUM_FORMAT_INT, 83: NUM_FORMAT_INT,  # value/salary/bank
        84: NUM_FORMAT_INT, 85: NUM_FORMAT_INT, 86: NUM_FORMAT_DEC1, 87: NUM_FORMAT_DEC1,  # jan
        88: NUM_FORMAT_INT, 89: NUM_FORMAT_INT, 90: NUM_FORMAT_DEC1, 91: NUM_FORMAT_DEC1,  # aug
    }
    # Comparativos de desempenho/disciplina/eficiência
    t_colors = list(range(5, 46)) + list(range(72, 92))
    t_inv = [9, 11, 12, 13, 14, 18, 22, 24, 26, 28, 30, 37, 39, 41, 43, 44, 46, 48]
    write_sheet(wb, "Times", f"Relatório de Equipes - {data['club']}", localized_headers(TIMES_HEADERS, language), t_rows, t_fmts, t_colors, t_inv)
    
    # Carreira (ativos + aposentados)
    c_rows = data.get("career") or []
    c_fmts = {
        14: NUM_FORMAT_DEC1,  # idade
        17: NUM_FORMAT_DEC1, 18: NUM_FORMAT_DEC1, 19: NUM_FORMAT_DEC1,
        22: NUM_FORMAT_DEC1,
        24: NUM_FORMAT_DEC3, 25: NUM_FORMAT_DEC1,
        30: NUM_FORMAT_DEC3, 31: NUM_FORMAT_DEC1,
        34: NUM_FORMAT_DEC1, 35: NUM_FORMAT_DEC3,
        37: NUM_FORMAT_DEC1, 38: NUM_FORMAT_DEC1,
        41: NUM_FORMAT_DEC1,
        42: NUM_FORMAT_INT, 43: NUM_FORMAT_INT,
    }
    # colorir overalls, gols, G+A, ratings, títulos, prêmios
    c_colors = [17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 29, 30, 31, 32, 33, 34, 35, 37, 38, 39, 40, 41, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71]
    c_inv = [14]  # idade: menor pode ser melhor no contexto young talent — na verdade neutro; deixar só overalls positivos
    write_sheet(
        wb,
        "Carreira",
        f"Carreira completa (ativos + aposentados) — {data['club']} ({data['current_day']} dias)",
        localized_headers(CAREER_HEADERS, language),
        c_rows,
        c_fmts,
        c_colors,
        [],
    )

    # Resumo e Dicionário
    write_sheet(wb, "Resumo", "Resumo do Processamento", ["Indicador", "Valor"], [
        ["Clube Gerenciado", data["club"]],
        ["Dia Interno", data["current_day"]],
        ["Jogadores Processados", len(data["players"])],
        ["Goleiros Elegíveis", len(data["gks"])],
        ["Times Analisados", len(data["teams"])],
        ["Partidas Concluídas", data["played_matches"]],
        ["Linhas na aba Carreira", len(c_rows)],
    ], {})
    write_sheet(wb, "Dicionário", "Significado das Colunas", ["Campo", "Leitura"], [
        ["xGP", "Gols esperados de pênalti (base 0,75 ajustado por overall SHO vs GKP)"],
        ["Gols pênalti vs xGP", "Diferença entre gols reais e esperados (positivo = acima da média)"],
        ["Taxa conversão", "Percentual de acerto nas cobranças (inclui pênaltis de desempate, minuto > 119)"],
        ["Pênaltis (desempate)", "Eventos com código 7/8 após o 119' entram nas mesmas colunas de pênalti"],
        ["Role CD…ST", "Nota 0–10 daquele papel no jogador (0 se não tiver); GK não tem coluna própria aqui"],
        ["Cores", "Percentis na coluna: vermelho ruim → azul excelente (colunas invertidas: menor é melhor)"],
        ["Carreira — fonte", "players2 (ativos) + players_retired2 (aposentados); séries h_played/h_goals/h_assists/h_rating/h_team_id"],
        ["Carreira — seleção", "Apenas segmentos com h_is_national=1 (team_id=-1 sozinho = livre, não conta)"],
        ["Carreira — prêmios", "tipo 6=MVP de PARTIDA | tipo 5=Time do ano/Seleção do torneio (TOP11) | tipo 3 global=Golden Boy | tipo 2 global=Bola de Ouro | 0=artilheiro | 1=assistência | 4=melhor GK"],
        ["MVP vs Time do ano", "NÃO confundir: MVP (tipo6) é por partida; Time do ano/Seleção (tipo5) é prêmio de temporada/torneio"],
        ["Títulos", "rt0 pos1=liga | rt7 pos0=copa nacional | rt1/3/4 pos0=supercopa nacional | rt2 pos0=continental | rt6 pos0=supercopa continental | rt11=Copa do Mundo"],
        ["Jogos no ano", "s_played (s_matches costuma vir vazio). Se matches2 só tem fixtures futuras, não há xG/MVP de partida"],
        ["Carreira — títulos", "historic_results2 (id_type=2): liga pos=1; copa/continental pos=0 conforme result_type"],
        ["Hat-tricks / decisivos / reservas", "Calculados pelos eventos da partida na TEMPORADA atual (código 0=gol, 1=assist, 5=entrou do banco)"],
    ], {})
    
    wb.save(output_path)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--save")
    parser.add_argument("--template")
    parser.add_argument("--output", required=True)
    parser.add_argument("--language", default="pt")
    parser.add_argument("--template-only", action="store_true")
    args = parser.parse_args()
    try:
        if args.template_only:
            # Gerar apenas o modelo limpo (headers e legendas)
            data = {
                "current_day": 0, "club": "Modelo", "players": [], "gks": [], "teams": [],
                "club_summary": {"matches": 0, "shots": 0, "goals": 0, "xga": 0.0, "ga": 0},
                "played_matches": 0
            }
            create_workbook(data, args.output, args.language)
        else:
            if not args.save:
                print("ERRO: --save é obrigatório quando não usar --template-only", file=sys.stderr)
                sys.exit(1)
            data = extract_data(Path(args.save))
            create_workbook(data, args.output, args.language)
            emit(100, "Relatório concluído com sucesso")
    except Exception as e:
        print(f"ERRO: {e}", file=sys.stderr)
        sys.exit(1)


def _running_in_streamlit() -> bool:
    """Detecta se o arquivo está sendo executado via `streamlit run`."""
    import os
    # Variáveis que o Streamlit define no processo
    if os.environ.get("STREAMLIT_SERVER_PORT") or os.environ.get("STREAMLIT_RUNTIME_ENV"):
        return True
    if any("streamlit" in str(arg).lower() for arg in sys.argv):
        return True
    try:
        from streamlit.runtime.scriptrunner import get_script_run_ctx
        return get_script_run_ctx() is not None
    except Exception:
        return False


def run_streamlit_app() -> None:
    """Interface web — não altera a lógica de processamento."""
    import streamlit as st

    st.set_page_config(
        page_title="FM26 Moneyball",
        page_icon="⚽",
        layout="centered",
        initial_sidebar_state="expanded",
    )

    st.title("⚽ FM26 Moneyball")
    st.caption("Envie um save Footlord / FM26 (.fl ou .db) e baixe a planilha Moneyball.")

    with st.sidebar:
        st.header("Opções")
        language = st.radio(
            "Idioma da planilha",
            options=("pt", "en"),
            format_func=lambda x: "Português" if x == "pt" else "English",
            index=0,
        )
        st.markdown("---")
        st.markdown(
            "Abas geradas:\n"
            "- **Players 2**\n"
            "- **Goleiros xG**\n"
            "- **Times**\n"
            "- **Resumo**\n"
            "- **Dicionário**"
        )
        st.markdown("---")
        st.caption("CLI: `python app.py --save arquivo.fl --output saida.xlsx`")
        if st.button("Limpar resultado em memória"):
            for key in ("xlsx_bytes", "xlsx_name", "xlsx_info", "template_bytes"):
                st.session_state.pop(key, None)
            st.rerun()

    uploaded = st.file_uploader(
        "Save FM26 / Footlord",
        type=["fl", "db"],
        help="Arquivo .fl (ZIP do save) ou .db extraído",
    )

    col1, col2 = st.columns(2)
    with col1:
        generate = st.button(
            "Gerar planilha",
            type="primary",
            use_container_width=True,
            disabled=uploaded is None,
        )
    with col2:
        template_only = st.button("Só modelo vazio", use_container_width=True)

    if template_only:
        with tempfile.TemporaryDirectory(prefix="fm26-st-") as temporary_dir:
            out = Path(temporary_dir) / "Moneyball_modelo_limpo.xlsx"
            try:
                with st.spinner("Criando modelo limpo..."):
                    data = {
                        "current_day": 0,
                        "club": "Modelo",
                        "players": [],
                        "gks": [],
                        "teams": [],
                        "club_summary": {"matches": 0, "shots": 0, "goals": 0, "xga": 0.0, "ga": 0},
                        "played_matches": 0,
                    }
                    create_workbook(data, out, language)
                st.session_state["template_bytes"] = out.read_bytes()
                st.session_state["xlsx_bytes"] = None
                st.session_state["xlsx_name"] = None
                st.session_state["xlsx_info"] = None
            except Exception as error:
                st.error(f"Erro: {error}")

    if generate and uploaded is not None:
        progress = st.progress(0, text="Iniciando...")
        status = st.empty()

        def ui_emit(percent: int, message: str) -> None:
            progress.progress(min(100, max(0, int(percent))) / 100.0, text=message)
            status.info(message)

        global emit
        original_emit = emit
        emit = ui_emit  # type: ignore[assignment]

        try:
            with tempfile.TemporaryDirectory(prefix="fm26-st-") as temporary_dir:
                safe_name = Path(uploaded.name).name.replace(" ", "_")
                save_path = Path(temporary_dir) / safe_name
                save_path.write_bytes(uploaded.getvalue())
                out_name = f"Moneyball_{Path(safe_name).stem}.xlsx"
                output_path = Path(temporary_dir) / out_name

                ui_emit(5, "Lendo o save...")
                data = extract_data(save_path)
                ui_emit(80, "Montando a planilha...")
                create_workbook(data, output_path, language)
                ui_emit(100, "Concluído")

                xlsx_bytes = output_path.read_bytes()
                st.session_state["xlsx_bytes"] = xlsx_bytes
                st.session_state["xlsx_name"] = out_name
                st.session_state["xlsx_info"] = {
                    "club": data.get("club", "Save"),
                    "day": data.get("current_day", "?"),
                    "players": len(data.get("players", [])),
                    "gks": len(data.get("gks", [])),
                    "teams": len(data.get("teams", [])),
                    "size_mb": round(len(xlsx_bytes) / (1024 * 1024), 2),
                }
                st.session_state.pop("template_bytes", None)
                status.empty()
                progress.empty()
        except Exception as error:
            st.error(f"Erro ao processar o save: {error}")
            st.session_state.pop("xlsx_bytes", None)
            st.session_state.pop("xlsx_name", None)
            st.session_state.pop("xlsx_info", None)
        finally:
            emit = original_emit  # type: ignore[assignment]

    # Download fora do if generate → sobrevive ao clique (session_state)
    if st.session_state.get("xlsx_bytes"):
        info = st.session_state.get("xlsx_info") or {}
        st.success(
            f"**{info.get('club', 'Save')}** — data interna **{info.get('day', '?')}** · "
            f"{info.get('players', '?')} jogadores · {info.get('gks', '?')} goleiros · "
            f"{info.get('teams', '?')} times · ~{info.get('size_mb', '?')} MB"
        )
        st.download_button(
            label="⬇️ Baixar planilha Moneyball (.xlsx)",
            data=st.session_state["xlsx_bytes"],
            file_name=st.session_state.get("xlsx_name") or "Moneyball.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="dl_moneyball",
        )

    if st.session_state.get("template_bytes"):
        st.success("Modelo limpo criado.")
        st.download_button(
            label="⬇️ Baixar modelo XLSX",
            data=st.session_state["template_bytes"],
            file_name="Moneyball_modelo_limpo.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="dl_template",
        )


if _running_in_streamlit():
    run_streamlit_app()
elif __name__ == "__main__":
    main()
