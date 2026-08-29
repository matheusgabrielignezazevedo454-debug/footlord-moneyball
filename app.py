#!/usr/bin/env python3
"""Processador FM26 Moneyball com saída XLSX limpa e compatível com Excel Mobile."""

from __future__ import annotations

import argparse
import math
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


JAN_CUTOFF = 184  # 01/01/2026; a data interna inicia em 01/07/2025.
JAN_WINDOW = (184, 243)
AUG_WINDOW = (365, 426)

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

PLAYERS_HEADERS = [
    "Rank", "Jogador", "Posição", "Roles secundários (0-10)", "Idade", "Time", "Liga", "Divisão", "País",
    "Overall do Footlord (0-100)", "Overall início da temporada", "Variação overall na temporada", "Overall início da carreira", "Variação overall desde início da carreira",
    "Jogos no ano", "Cartões amarelos", "Cartões vermelhos", "Cartões totais", "Amarelos/partida", "Vermelhos/partida", "Gols no ano", "Assistências no ano", "G+A no ano",
    "Gols/partida", "Assistências/partida", "G+A/partida", "Média de avaliação", "Minutos estimados",
    "Minutos/partida", "Minutos/gol", "Minutos/assistência", "MVP do time", "Pior nota do time",
    "MVP %", "Pior nota %", "Clean sheets", "Valor de mercado em €", "Salário em €",
    "Contrato (anos)", "Status", "Empréstimo", "Lista de transferências", "PAC", "SHO", "PAS", "DEF", "PHY", "MEN", "GKP", "ID interno",
]

GK_HEADERS = [
    "Rank", "Goleiro", "Time", "Liga", "Partidas GK", "Minutos", "Clean sheets", "Clean sheets/partida", "xG adversário", "Gols sofridos", "Gols/partida GK",
    "Gols evitados vs xG", "Gols evitados/partida", "Gols evitados/90", "Taxa de prevenção", "Gols/xG", "Classificação",
    "Elegível (mín. 5 jogos)", "Overall do Footlord", "ID interno",
]

PLAYER_CARDS_HEADERS = [
    "Rank cartões", "Jogador", "Posição", "Time", "Liga", "Partidas no ano", "Cartões amarelos", "Cartões vermelhos",
    "Cartões totais", "Amarelos/partida", "Vermelhos/partida", "ID interno",
]

TEAM_CARDS_HEADERS = [
    "Rank cartões", "Time", "Liga", "Partidas no ano", "Cartões amarelos", "Cartões vermelhos", "Cartões totais",
    "Amarelos/partida", "Vermelhos/partida",
]

OVERALL_EVOLUTION_HEADERS = [
    "Rank overall", "Jogador", "Posição", "Time", "Liga", "Overall início da temporada", "Overall atual",
    "Variação overall na temporada", "Overall início da carreira", "Variação overall desde início da carreira", "ID interno",
]

TIMES_HEADERS = [
    "Rank overall", "Time", "Liga", "Competição continental", "Jogadores no time", "Média de idade",
    "Partidas no ano", "Vitórias no ano", "Empates no ano", "Derrotas no ano", "Cartões amarelos no ano", "Cartões vermelhos no ano", "Cartões totais no ano", "Amarelos/partida", "Vermelhos/partida", "Partidas antes 1/jan", "Vitórias antes", "Empates antes", "Derrotas antes",
    "Partidas depois 1/jan", "Vitórias depois", "Empates depois", "Derrotas depois", "Maior sequência vitórias antes", "Maior sequência derrotas antes",
    "Maior sequência vitórias depois", "Maior sequência derrotas depois", "Gols marcados no ano", "Gols sofridos no ano", "Gols marcados/partida",
    "Gols sofridos/partida", "Chutes totais", "Chutes/partida", "Chutes no alvo", "% no alvo", "Conversão de chutes", "Gols marcados antes", "Gols sofridos antes", "Gols marcados/partida antes", "Gols sofridos/partida antes",
    "Gols marcados depois", "Gols sofridos depois", "Gols marcados/partida depois", "Gols sofridos/partida depois", "Jogos sem marcar",
    "Jogos sem sofrer gols", "Sem marcar antes", "Sem sofrer gols antes", "Sem marcar depois", "Sem sofrer gols depois", "Maior placar aplicado no ano",
    "Para quem — maior placar ano", "Pior placar sofrido no ano", "Para quem — pior placar ano", "Maior placar aplicado antes", "Para quem — maior placar antes",
    "Pior placar sofrido antes", "Para quem — pior placar antes", "Maior placar aplicado depois", "Para quem — maior placar depois", "Pior placar sofrido depois",
    "Para quem — pior placar depois", "Overall início do ano", "Overall fechamento agosto", "Overall final do ano", "Expectativa — posição por overall inicial",
    "Posição na liga no save", "Diferença expectativa/posição", "Média overall adversário — vitória", "Média overall adversário — empate",
    "Média overall adversário — derrota", "Valor do time — elenco em €", "Média salário dos jogadores em €", "Saldo bancário em €",
    "Gastos janela janeiro em €", "Vendas janela janeiro em €", "Overall médio compras janeiro", "Overall médio vendas janeiro",
    "Gastos janela agosto em €", "Vendas janela agosto em €", "Overall médio compras agosto", "Overall médio vendas agosto",
]

HEADER_TRANSLATIONS_EN = {
    "Rank": "Rank", "Quality Index (0-100)": "Quality Index (0-100)", "Quality Index": "Quality Index", "Score Moneyball (0-100)": "Moneyball Score (0-100)", "Score Moneyball": "Moneyball Score", "MVP %": "MVP %", "Clean sheets": "Clean sheets", "Clean sheets/partida": "Clean sheets/match", "Status": "Status", "PAC": "PAC", "SHO": "SHO", "PAS": "PAS", "DEF": "DEF", "PHY": "PHY", "MEN": "MEN", "GKP": "GKP", "Gols/xG": "Goals/xG",
    "Jogador": "Player", "Posição": "Position", "Roles secundários (0-10)": "Secondary roles (0-10)", "Overall do Footlord (0-100)": "Footlord Overall (0-100)", "Overall do Footlord": "Footlord Overall", "Idade": "Age", "Time": "Team", "Liga": "League", "Divisão": "Division", "País": "Country",
    "Jogos no ano": "Matches in year", "Gols no ano": "Goals in year", "Assistências no ano": "Assists in year", "G+A no ano": "G+A in year", "Gols/partida": "Goals/match", "Assistências/partida": "Assists/match", "G+A/partida": "G+A/match", "Média de avaliação": "Average rating", "Minutos estimados": "Estimated minutes", "Partidas com minutos": "Matches with minutes", "Minutos/partida": "Minutes/match", "Minutos/gol": "Minutes/goal", "Minutos/assistência": "Minutes/assist", "MVP do time": "Team MVP", "Pior nota do time": "Team worst rating", "Partidas com rating": "Matches with rating", "Pior nota %": "Worst rating %", "Partidas GK": "GK matches", "xG adversário GK": "Opponent xG GK", "Gols sofridos GK": "Goals conceded GK", "Gols evitados vs xG": "Goals prevented vs xG", "xG/partida GK": "xG/match GK", "Gols/partida GK": "Goals/match GK", "Gols/xG GK": "Goals/xG GK", "Valor de mercado em €": "Market value in €", "Salário em €": "Salary in €", "Contrato (anos)": "Contract (years)", "Empréstimo": "Loan", "Lista de transferências": "Transfer listed", "ID interno": "Internal ID",
    "Goleiro": "Goalkeeper", "Minutos": "Minutes", "xG adversário": "Opponent xG", "Gols sofridos": "Goals conceded", "Gols evitados/partida": "Goals prevented/match", "Gols evitados/90": "Goals prevented/90", "Taxa de prevenção": "Prevention rate", "Classificação": "Rating", "Elegível (mín. 5 jogos)": "Eligible (min. 5 matches)",
    "Rank cartões": "Cards rank", "Cartões amarelos": "Yellow cards", "Cartões vermelhos": "Red cards", "Cartões totais": "Total cards", "Amarelos/partida": "Yellow cards/match", "Vermelhos/partida": "Red cards/match", "Cartões amarelos no ano": "Yellow cards in year", "Cartões vermelhos no ano": "Red cards in year", "Cartões totais no ano": "Total cards in year", "Overall início da temporada": "Overall at season start", "Overall atual": "Current overall", "Variação overall na temporada": "Overall change in season", "Overall início da carreira": "Overall at career start", "Variação overall desde início da carreira": "Overall change since career start",
    "Rank overall": "Overall rank", "Competição continental": "Continental competition", "Jogadores no time": "Players in team", "Média de idade": "Average age", "Partidas no ano": "Matches in year", "Vitórias no ano": "Wins in year", "Empates no ano": "Draws in year", "Derrotas no ano": "Losses in year", "Partidas antes 1/jan": "Matches before Jan 1", "Vitórias antes": "Wins before", "Empates antes": "Draws before", "Derrotas antes": "Losses before", "Partidas depois 1/jan": "Matches after Jan 1", "Vitórias depois": "Wins after", "Empates depois": "Draws after", "Derrotas depois": "Losses after", "Maior sequência vitórias antes": "Longest win streak before", "Maior sequência derrotas antes": "Longest loss streak before", "Maior sequência vitórias depois": "Longest win streak after", "Maior sequência derrotas depois": "Longest loss streak after", "Gols marcados no ano": "Goals scored in year", "Gols sofridos no ano": "Goals conceded in year", "Gols marcados/partida": "Goals scored/match", "Gols sofridos/partida": "Goals conceded/match", "Chutes totais": "Total shots", "Chutes/partida": "Shots/match", "Chutes no alvo": "Shots on target", "% no alvo": "% on target", "Conversão de chutes": "Shot conversion", "Gols marcados antes": "Goals scored before", "Gols sofridos antes": "Goals conceded before", "Gols marcados/partida antes": "Goals scored/match before", "Gols sofridos/partida antes": "Goals conceded/match before", "Gols marcados depois": "Goals scored after", "Gols sofridos depois": "Goals conceded after", "Gols marcados/partida depois": "Goals scored/match after", "Gols sofridos/partida depois": "Goals conceded/match after", "Jogos sem marcar": "Matches without scoring", "Jogos sem sofrer gols": "Clean sheets", "Sem marcar antes": "Without scoring before", "Sem sofrer gols antes": "Clean sheets before", "Sem marcar depois": "Without scoring after", "Sem sofrer gols depois": "Clean sheets after", "Maior placar aplicado no ano": "Biggest win in year", "Para quem — maior placar ano": "Opponent — biggest win in year", "Pior placar sofrido no ano": "Worst loss in year", "Para quem — pior placar ano": "Opponent — worst loss in year", "Maior placar aplicado antes": "Biggest win before", "Para quem — maior placar antes": "Opponent — biggest win before", "Pior placar sofrido antes": "Worst loss before", "Para quem — pior placar antes": "Opponent — worst loss before", "Maior placar aplicado depois": "Biggest win after", "Para quem — maior placar depois": "Opponent — biggest win after", "Pior placar sofrido depois": "Worst loss after", "Para quem — pior placar depois": "Opponent — worst loss after", "Overall início do ano": "Overall at start of year", "Overall fechamento agosto": "Overall at August deadline", "Overall final do ano": "Overall at end of year", "Expectativa — posição por overall inicial": "Expected position by starting overall", "Posição na liga no save": "League position in save", "Diferença expectativa/posição": "Expectation/position difference", "Média overall adversário — vitória": "Average opponent overall — win", "Média overall adversário — empate": "Average opponent overall — draw", "Média overall adversário — derrota": "Average opponent overall — loss", "Valor do time — elenco em €": "Team value — squad in €", "Média salário dos jogadores em €": "Average player salary in €", "Saldo bancário em €": "Bank balance in €", "Gastos janela janeiro em €": "January window spending in €", "Vendas janela janeiro em €": "January window sales in €", "Overall médio compras janeiro": "Average overall — January signings", "Overall médio vendas janeiro": "Average overall — January sales", "Gastos janela agosto em €": "August window spending in €", "Vendas janela agosto em €": "August window sales in €", "Overall médio compras agosto": "Average overall — August signings", "Overall médio vendas agosto": "Average overall — August sales",
    "Indicador": "Metric", "Valor": "Value", "Interpretação": "Interpretation", "Campo": "Field", "Fonte": "Source", "Cálculo": "Calculation", "Leitura": "Reading",
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
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
TITLE_FONT = Font(name="Aptos Display", size=14, bold=True, color="FFFFFF")
HEADER_FONT = Font(name="Aptos", size=10, bold=True, color="FFFFFF")
DATA_FONT = Font(name="Aptos", size=10, color="111827")
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


def parse_events(raw):
    events = []
    for item in str(raw or "").split(";"):
        parts = item.split("|")
        if len(parts) != 3:
            continue
        try:
            events.append((max(0, min(90, int(parts[0]))), int(parts[1]), int(parts[2])))
        except (TypeError, ValueError):
            pass
    return events


def parse_role_levels(raw):
    result = {}
    for item in str(raw or "").split(","):
        if ":" not in item:
            continue
        role, value = item.split(":", 1)
        result[role] = sf(value)
    return result


def role_group(role):
    if role == "GK":
        return "GK"
    if role in {"CD", "CB", "LB", "RB", "CDM", "DM", "LWB", "RWB"}:
        return "DEF"
    if role in {"CM", "LM", "RM", "CAM", "AM"}:
        return "MID"
    return "ATT"


def quality_index(player, role):
    values = {key: sf(player.get(key)) for key in ("PAC", "SHO", "PAS", "DEF", "PHY", "MEN", "GKP")}
    weights = {
        "GK": {"GKP": 0.65, "PAS": 0.10, "DEF": 0.10, "PHY": 0.10, "MEN": 0.05},
        "DEF": {"DEF": 0.30, "PAS": 0.20, "PAC": 0.15, "PHY": 0.15, "MEN": 0.10, "SHO": 0.10},
        "MID": {"PAS": 0.20, "SHO": 0.20, "PAC": 0.15, "DEF": 0.15, "PHY": 0.15, "MEN": 0.15},
        "ATT": {"SHO": 0.30, "PAC": 0.22, "PAS": 0.18, "PHY": 0.15, "MEN": 0.15},
    }[role_group(role)]
    return clip(sum(values[key] * weight for key, weight in weights.items()))


def estimate_minutes(lineup_ids, entered_ids, subbed_ids, events):
    lineup, entered, subbed = set(lineup_ids), set(entered_ids), set(subbed_ids)
    appearances = lineup | entered | subbed
    entry, exit_ = {}, {}
    for minute, event_type, player_id in events:
        if event_type == 5:
            entry[player_id] = min(entry.get(player_id, minute), minute)
            appearances.add(player_id)
        elif event_type == 6:
            exit_[player_id] = min(exit_.get(player_id, minute), minute)
            appearances.add(player_id)
    result = {}
    for player_id in appearances:
        start = 0 if player_id in lineup else entry.get(player_id, 45)
        finish = exit_.get(player_id, 90)
        minutes = max(0, min(90, finish) - min(90, start))
        if minutes:
            result[player_id] = minutes
    return result


def average(values):
    clean = [sf(value) for value in values if value is not None and math.isfinite(sf(value))]
    return sum(clean) / len(clean) if clean else None


def percentile(values, value):
    ordered = sorted(value_ for value_ in values if value_ is not None and math.isfinite(sf(value_)))
    if not ordered or value is None:
        return 50.0
    if len(ordered) == 1:
        return 100.0
    less = bisect_left(ordered, value)
    equal = bisect_right(ordered, value) - less
    return 100.0 * (less + 0.5 * equal) / (len(ordered) - 1)


def team_overall_at(team, day, current_day):
    values = parse_values(team.get("team_in_time_value"))
    if not values:
        return sf(team.get("value"), None)
    if len(values) == 1 or current_day <= 0:
        return values[-1]
    ratio = max(0.0, min(1.0, sf(day) / current_day))
    position = ratio * (len(values) - 1)
    lower = int(position)
    upper = min(len(values) - 1, lower + 1)
    fraction = position - lower
    return values[lower] * (1 - fraction) + values[upper] * fraction


def new_stat():
    return {"games": 0, "wins": 0, "draws": 0, "losses": 0, "gf": 0, "ga": 0, "shots": 0, "shots_on": 0, "no_score": 0, "clean": 0, "results": [], "opp": {"W": [], "D": [], "L": []}}


def add_stat(stat, day, own_goals, opponent_goals, opponent, opponent_overall, shots_total=0, shots_on=0):
    own_goals, opponent_goals = si(own_goals), si(opponent_goals)
    result = "W" if own_goals > opponent_goals else "D" if own_goals == opponent_goals else "L"
    stat["games"] += 1
    stat["wins"] += result == "W"
    stat["draws"] += result == "D"
    stat["losses"] += result == "L"
    stat["gf"] += own_goals
    stat["ga"] += opponent_goals
    stat["shots"] += si(shots_total)
    stat["shots_on"] += si(shots_on)
    stat["no_score"] += own_goals == 0
    stat["clean"] += opponent_goals == 0
    stat["results"].append((si(day), result, own_goals, opponent_goals, opponent))
    if opponent_overall is not None:
        stat["opp"][result].append(opponent_overall)


def max_streak(results, wanted):
    best = current = 0
    for _, result, *_ in sorted(results, key=lambda item: item[0]):
        if result == wanted:
            current += 1
            best = max(best, current)
        else:
            current = 0
    return best


def extreme_score(results, best=True):
    if best:
        candidates = [item for item in results if item[2] > item[3]]
        key = lambda item: (item[2] - item[3], item[2], -item[0])
    else:
        candidates = [item for item in results if item[3] > item[2]]
        key = lambda item: (item[3] - item[2], item[3], -item[0])
    if not candidates:
        return None, None
    item = max(candidates, key=key)
    return f"{item[2]}–{item[3]}", item[4]


def period_values(stat):
    games = stat["games"]
    return {
        "games": games, "wins": stat["wins"], "draws": stat["draws"], "losses": stat["losses"],
        "win_streak": max_streak(stat["results"], "W"), "loss_streak": max_streak(stat["results"], "L"),
        "gf": stat["gf"], "ga": stat["ga"], "gf_pg": stat["gf"] / games if games else None, "ga_pg": stat["ga"] / games if games else None,
        "shots": stat["shots"], "shots_pg": stat["shots"] / games if games else None, "shots_on": stat["shots_on"], "on_target_pct": stat["shots_on"] / stat["shots"] if stat["shots"] else None, "conversion": stat["gf"] / stat["shots"] if stat["shots"] else None,
        "no_score": stat["no_score"], "clean": stat["clean"], "best_score": extreme_score(stat["results"], True), "worst_score": extreme_score(stat["results"], False),
    }


def transfer_metrics(cur, players, current_day):
    result = defaultdict(lambda: {"jan": {"spend": 0, "sales": 0, "buys": [], "sells": []}, "aug": {"spend": 0, "sales": 0, "buys": [], "sells": []}})
    tables = {row[0] for row in cur.execute("SELECT name FROM sqlite_master WHERE type='table'")}
    if "historic_transfers2" not in tables:
        return result
    for raw in cur.execute("SELECT * FROM historic_transfers2 WHERE transfer_date<=?", (current_day,)):
        row, day = dict(raw), si(raw["transfer_date"])
        window = "jan" if JAN_WINDOW[0] <= day <= JAN_WINDOW[1] else "aug" if AUG_WINDOW[0] <= day <= AUG_WINDOW[1] else None
        if not window:
            continue
        buyer, seller, price, move_type = si(row.get("buyer_id"), -1), si(row.get("seller_id"), -1), max(0, si(row.get("transfer_price"))), si(row.get("transfer_type"), -1)
        player = players.get(si(row.get("player_id")))
        overall = sf(player.get("temp_value"), None) if player else None
        if buyer > 0 and move_type in {0, 1, 2}:
            result[buyer][window]["spend"] += price if move_type in {0, 1} else 0
            if overall is not None:
                result[buyer][window]["buys"].append(overall)
        if seller > 0 and move_type in {0, 1} and price:
            result[seller][window]["sales"] += price
            if overall is not None:
                result[seller][window]["sells"].append(overall)
    return result


def build_team_rows(cur, current_day, players, played_matches, cards_by_team):
    raw_teams = cur.execute("""
        SELECT t.*, ch.name AS champ_name, ch.league_id AS league_id, l.name AS league_name
        FROM teams2 t LEFT JOIN champs2 ch ON ch.id=t.champ_id LEFT JOIN leagues2 l ON l.id=ch.league_id
        WHERE COALESCE(t.is_national_team,0)=0 AND t.id>0 AND COALESCE(t.champ_id,0)>0 ORDER BY t.name
    """).fetchall()
    teams = {si(row["id"]): dict(row) for row in raw_teams}
    continental = defaultdict(list)
    tables = {row[0] for row in cur.execute("SELECT name FROM sqlite_master WHERE type='table'")}
    if "int_cups2" in tables:
        for raw in cur.execute("SELECT name, teams_id FROM int_cups2"):
            for team_id in parse_ids(raw["teams_id"]):
                continental[team_id].append(raw["name"])
    stats = defaultdict(lambda: {"all": new_stat(), "before": new_stat(), "after": new_stat()})
    league_table = defaultdict(lambda: defaultdict(lambda: {"pts": 0, "gd": 0, "gf": 0}))
    for raw in played_matches:
        match, day = dict(raw), si(raw["date"])
        period = "before" if day < JAN_CUTOFF else "after"
        for side, other in ((1, 2), (2, 1)):
            team_id, opponent_id = si(match.get(f"team_{side}_id"), -1), si(match.get(f"team_{other}_id"), -1)
            if team_id not in teams:
                continue
            opponent = teams.get(opponent_id, {}).get("name") or f"ID {opponent_id}"
            opponent_overall = team_overall_at(teams[opponent_id], day, current_day) if opponent_id in teams else None
            shots_on = si(match.get(f"shots_on_{side}"))
            shots_total = shots_on + si(match.get(f"shots_off_{side}"))
            add_stat(stats[team_id]["all"], day, match.get(f"goals_{side}"), match.get(f"goals_{other}"), opponent, opponent_overall, shots_total, shots_on)
            add_stat(stats[team_id][period], day, match.get(f"goals_{side}"), match.get(f"goals_{other}"), opponent, opponent_overall, shots_total, shots_on)
        competition = str(match.get("competition") or "")
        if competition.startswith("league_"):
            for team_id, goals_for, goals_against in ((si(match.get("team_1_id"), -1), si(match.get("goals_1")), si(match.get("goals_2"))), (si(match.get("team_2_id"), -1), si(match.get("goals_2")), si(match.get("goals_1")))):
                if team_id not in teams:
                    continue
                standing = league_table[competition][team_id]
                standing["gf"] += goals_for
                standing["gd"] += goals_for - goals_against
                standing["pts"] += 3 if goals_for > goals_against else 1 if goals_for == goals_against else 0
    league_positions = {}
    for standings in league_table.values():
        for position, (team_id, _) in enumerate(sorted(standings.items(), key=lambda item: (-item[1]["pts"], -item[1]["gd"], -item[1]["gf"], item[0])), 1):
            league_positions[team_id] = position
    expected_positions = {}
    by_league = defaultdict(list)
    for team_id, team in teams.items():
        if team.get("league_id") is not None:
            by_league[si(team["league_id"])].append((team_id, team))
    for entries in by_league.values():
        for position, (team_id, _) in enumerate(sorted(entries, key=lambda item: (-sf(team_overall_at(item[1], JAN_CUTOFF, current_day), -1), item[0])), 1):
            expected_positions[team_id] = position
    transfers = transfer_metrics(cur, players, current_day)
    rows = []
    for team_id, team in teams.items():
        roster_ids = list(dict.fromkeys(parse_ids(team.get("players"))))
        roster = [players[player_id] for player_id in roster_ids if player_id in players]
        periods = {key: period_values(stats[team_id][key]) for key in ("all", "before", "after")}
        transfer = transfers[team_id]
        expected, actual = expected_positions.get(team_id), league_positions.get(team_id)
        rows.append({
            "rank": 0, "name": team.get("name") or f"ID {team_id}", "league": team.get("league_name") or team.get("champ_name") or "Sem liga",
            "continental": " / ".join(continental.get(team_id, [])) or "Não", "players": len(roster), "age": average([p.get("age") for p in roster]),
            "market": si(team.get("season_mk_value")) or sum(max(0, si(p.get("market_value"))) for p in roster), "salary": average([p.get("salary") for p in roster]),
            "balance": si(team.get("bank_balance")), "start": team_overall_at(team, JAN_CUTOFF, current_day),
            "aug": team_overall_at(team, AUG_WINDOW[1], current_day), "final": team_overall_at(team, current_day, current_day), "expected": expected, "actual": actual,
            "diff": actual - expected if actual and expected else None, "all": periods["all"], "before": periods["before"], "after": periods["after"],
            "opp_w": average(stats[team_id]["all"]["opp"]["W"]), "opp_d": average(stats[team_id]["all"]["opp"]["D"]), "opp_l": average(stats[team_id]["all"]["opp"]["L"]), "jan": transfer["jan"], "aug_t": transfer["aug"],
            "yellow_cards": cards_by_team[team_id]["yellow"], "red_cards": cards_by_team[team_id]["red"],
        })
    rows.sort(key=lambda row: (-sf(row["final"], -1), row["name"]))
    for rank, row in enumerate(rows, 1):
        row["rank"] = rank
    return rows


def team_values(row):
    all_, before, after = row["all"], row["before"], row["after"]
    def score(period, kind): return period[kind][0] if period[kind][0] else None
    def opponent(period, kind): return period[kind][1] if period[kind][1] else None
    jan, aug = row["jan"], row["aug_t"]
    return [
        row["rank"], row["name"], row["league"], row["continental"], row["players"], round(row["age"], 1) if row["age"] is not None else None,
        all_["games"], all_["wins"], all_["draws"], all_["losses"], row["yellow_cards"], row["red_cards"], row["yellow_cards"] + row["red_cards"], round(row["yellow_cards"] / all_["games"], 3) if all_["games"] else None, round(row["red_cards"] / all_["games"], 3) if all_["games"] else None, before["games"], before["wins"], before["draws"], before["losses"], after["games"], after["wins"], after["draws"], after["losses"],
        before["win_streak"], before["loss_streak"], after["win_streak"], after["loss_streak"], all_["gf"], all_["ga"], all_["gf_pg"], all_["ga_pg"], all_["shots"], all_["shots_pg"], all_["shots_on"], all_["on_target_pct"], all_["conversion"], before["gf"], before["ga"], before["gf_pg"], before["ga_pg"], after["gf"], after["ga"], after["gf_pg"], after["ga_pg"],
        all_["no_score"], all_["clean"], before["no_score"], before["clean"], after["no_score"], after["clean"], score(all_, "best_score"), opponent(all_, "best_score"), score(all_, "worst_score"), opponent(all_, "worst_score"),
        score(before, "best_score"), opponent(before, "best_score"), score(before, "worst_score"), opponent(before, "worst_score"), score(after, "best_score"), opponent(after, "best_score"), score(after, "worst_score"), opponent(after, "worst_score"),
        row["start"], row["aug"], row["final"], row["expected"], row["actual"], row["diff"], row["opp_w"], row["opp_d"], row["opp_l"], row["market"], row["salary"], row["balance"],
        jan["spend"], jan["sales"], average(jan["buys"]), average(jan["sells"]), aug["spend"], aug["sales"], average(aug["buys"]), average(aug["sells"]),
    ]


def build_data(db_path):
    emit(18, "Lendo dados do save FM26")
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    table_names = {row[0] for row in cur.execute("SELECT name FROM sqlite_master WHERE type='table'")}
    required = {"general", "players2", "teams2", "matches2", "champs2", "leagues2", "nations2"}
    missing = sorted(required - table_names)
    if missing:
        raise ValueError("Save incompatível: tabelas ausentes: " + ", ".join(missing))
    general_row = cur.execute("SELECT * FROM general LIMIT 1").fetchone()
    if not general_row:
        raise ValueError("Save sem informações gerais")
    general, current_day = dict(general_row), si(general_row["date"])
    team_rows = cur.execute("SELECT * FROM teams2").fetchall()
    teams = {si(row["id"]): dict(row) for row in team_rows}
    managed_id = next((si(general.get(key), -1) for key in ("user_id", "user_team_id", "club_id", "team_id") if si(general.get(key), -1) in teams), next(iter(teams), -1))
    managed_name = teams.get(managed_id, {}).get("name") or "Clube do save"
    player_query = """
        SELECT p.*, t.name AS team_name, ch.name AS champ_name, ch.league_id AS league_id, l.name AS league_name, l.division AS league_division, n.name AS nation_name
        FROM players2 p LEFT JOIN teams2 t ON t.id=p.team_id LEFT JOIN champs2 ch ON ch.id=t.champ_id LEFT JOIN leagues2 l ON l.id=ch.league_id LEFT JOIN nations2 n ON n.id=p.nation_id ORDER BY p.id
    """
    players = {}
    for raw in cur.execute(player_query):
        player = dict(raw)
        player["role"] = str(player.get("role") or "—").strip() or "—"
        player["group"] = role_group(player["role"])
        player["role_levels"] = parse_role_levels(player.get("roles_level"))
        player["secondary_roles"] = ", ".join(f"{role}:{value:.1f}" for role, value in player["role_levels"].items() if role != player["role"] and value > 0) or "—"
        player["games"] = parse_count_blocks(player.get("s_played"))
        player["goals"] = parse_count_blocks(player.get("s_goals"))
        player["assists"] = parse_count_blocks(player.get("s_assists"))
        player["raw_conceded"] = parse_count_blocks(player.get("s_g_conceded"))
        player["raw_clean"] = parse_count_blocks(player.get("s_clean_sheets"))
        ratings = parse_rating_blocks(player.get("s_ratings"))
        player["rating_count"] = len(ratings)
        player["raw_rating"] = sum(ratings) / len(ratings) if ratings else 0.0
        player["age"] = max(0, int((current_day - si(player.get("birth_date"))) / 365.25)) if player.get("birth_date") is not None else None
        expiry = si(player.get("contract_expiration"), -1)
        player["contract_years"] = max(0.0, (expiry - current_day) / 365.25) if expiry >= 0 else None
        player["fit"] = clip(player["role_levels"].get(player["role"], 0), 0, 10)
        player["quality"] = quality_index(player, player["role"])
        raw_overall = sf(player.get("temp_value"), None)
        goalkeeper_overall = sf(player.get("GKP"), None)
        player["overall"] = clip(goalkeeper_overall) if player["role"] == "GK" and goalkeeper_overall is not None and goalkeeper_overall > 0 else clip(raw_overall) if raw_overall is not None and raw_overall > 0 else player["quality"]
        player["season_overall_history"] = parse_values(player.get("s_value_in_time"))
        player["career_overall_history"] = parse_values(player.get("value_in_time"))
        player["yellow_cards"] = parse_count_blocks(player.get("s_y_cards"))
        player["red_cards"] = parse_count_blocks(player.get("s_r_cards"))
        player["market_value"] = si(player.get("temp_mk_value_modded")) or si(player.get("temp_mk_value"))
        player["salary"] = si(player.get("salary"))
        player.update({"mvp": 0, "worst": 0, "rating_matches": 0, "rating_sum": 0.0, "apps": 0, "minutes": 0.0, "clean_sheets": 0.0, "gk_apps": 0, "gk_minutes": 0.0, "gk_clean": 0.0, "gk_xg": 0.0, "gk_goals": 0.0})
        players[si(player["id"])] = player
    emit(38, "Calculando minutos, ratings e xG")
    all_matches = cur.execute("SELECT * FROM matches2 WHERE date<=? ORDER BY date,id", (current_day,)).fetchall()
    played_matches = [match for match in all_matches if si(match["state"], -1) == 2 and match["goals_1"] is not None and match["goals_2"] is not None]
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
            lineup, entered, subbed = parse_ids(match.get(f"lineup_{side}")), parse_ids(match.get(f"entered_{side}")), parse_ids(match.get(f"subbed_{side}"))
            minutes = estimate_minutes(lineup, entered, subbed, parse_events(match.get(f"events_{side}")))
            ratings = parse_match_ratings(match.get(f"ratings_{side}"))
            if ratings:
                high, low = max(ratings.values()), min(ratings.values())
                for player_id, rating in ratings.items():
                    if player_id in players:
                        players[player_id]["rating_matches"] += 1
                        players[player_id]["rating_sum"] += rating
                        players[player_id]["mvp"] += rating == high
                        players[player_id]["worst"] += rating == low
            sides.append({"side": side, "team": si(match.get(f"team_{side}_id")), "minutes": minutes, "ratings": ratings, "goals": si(match.get(f"goals_{side}")), "xg": sf(match.get(f"x_goals_{side}"))})
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
        if player["role"] == "GK" and not player["gk_apps"] and player["games"]:
            player["gk_apps"], player["gk_minutes"], player["gk_clean"], player["gk_goals"] = int(player["games"]), player["minutes"], player["raw_clean"], player["raw_conceded"]
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
        age_score = clip((34 - sf(player["age"], 34)) / 16 * 100)
        player["score"] = clip(0.50 * player["quality"] + 0.20 * player["rating_score"] + 0.15 * output_score + 0.10 * cost_score + 0.05 * age_score)
    ordered_players = sorted(players.values(), key=lambda player: (-player["overall"], player["market_value"], player["id"]))
    goalkeepers = sorted([player for player in players.values() if player["role"] == "GK" and player["gk_apps"]], key=lambda player: (-player["overall"], -player["gk_apps"], player["id"]))
    team_rows = build_team_rows(cur, current_day, players, played_matches, team_cards)
    conn.close()
    return {"current_day": current_day, "club": managed_name, "players": ordered_players, "gks": goalkeepers, "teams": team_rows, "club_summary": club_summary, "played_matches": len(played_matches)}


def player_values(players):
    rows = []
    for rank, player in enumerate(players, 1):
        name = " ".join(part for part in (str(player.get("name") or "").strip(), str(player.get("surname") or "").strip()) if part) or f"ID {player['id']}"
        gk_xg = player["gk_xg"] if player["role"] == "GK" and player["gk_apps"] else None
        gk_ga = player["gk_goals"] if player["role"] == "GK" and player["gk_apps"] else None
        prevented = gk_xg - gk_ga if gk_xg is not None and gk_ga is not None else None
        clean = round(player["clean_sheets"]) if player["group"] == "DEF" else None
        current = round(player["overall"], 1)
        season_start = player["season_overall_history"][0] if player["season_overall_history"] else current
        career_start = player["career_overall_history"][0] if player["career_overall_history"] else season_start
        games, yellow, red = round(player["games"]), round(player["yellow_cards"]), round(player["red_cards"])
        rows.append([
            rank, name, player["role"], player["secondary_roles"], player["age"], player.get("team_name") or "Sem clube", player.get("league_name") or "Sem liga", player.get("league_division"), player.get("nationality") or player.get("nation_name") or "",
            current, round(season_start, 1), round(current - season_start, 1), round(career_start, 1), round(current - career_start, 1), games, yellow, red, yellow + red, round(yellow / games, 3) if games else None, round(red / games, 3) if games else None, round(player["goals"]), round(player["assists"]), round(player["ga"]), round(player["goals_pg"], 3), round(player["assists_pg"], 3), round(player["ga_pg"], 3), round(player["rating"], 2) if player["rating"] else None,
            round(player["minutes"]), round(player["minutes"] / player["apps"], 1) if player["apps"] else None, round(player["minutes_goal"], 1) if player["minutes_goal"] else None, round(player["minutes_assist"], 1) if player["minutes_assist"] else None,
            player["mvp"], player["worst"], player["mvp_pct"], player["worst_pct"], clean,
            player["market_value"], player["salary"], round(player["contract_years"], 1) if player["contract_years"] is not None else None, "Sem clube" if si(player.get("team_id"), -1) <= 0 else ("Em empréstimo" if si(player.get("loan_status")) else "Contrato"),
            "Sim" if si(player.get("loan_status")) else "Não", "Sim" if si(player.get("transfer_status")) == 1 else "Não", round(sf(player.get("PAC")), 1), round(sf(player.get("SHO")), 1), round(sf(player.get("PAS")), 1), round(sf(player.get("DEF")), 1), round(sf(player.get("PHY")), 1), round(sf(player.get("MEN")), 1), round(sf(player.get("GKP")), 1), player["id"],
        ])
    return rows


def goalkeeper_values(goalkeepers):
    rows = []
    for rank, player in enumerate(goalkeepers, 1):
        xg, conceded, apps, minutes = player["gk_xg"], player["gk_goals"], player["gk_apps"], player["gk_minutes"]
        prevented = xg - conceded
        rate = prevented / xg if xg else None
        classification = "Amostra curta" if apps < 5 else "Sem xG" if rate is None else "Excelente" if rate >= .15 else "Bom" if rate >= .05 else "Neutro" if rate >= -.05 else "Crítico"
        name = " ".join(part for part in (str(player.get("name") or "").strip(), str(player.get("surname") or "").strip()) if part) or f"ID {player['id']}"
        rows.append([rank, name, player.get("team_name") or "Sem clube", player.get("league_name") or "Sem liga", apps, round(minutes), round(player["gk_clean"]), round(player["gk_clean"] / apps, 3) if apps else None, round(xg, 2), round(conceded, 2), round(conceded / apps, 2) if apps else None, round(prevented, 2), round(prevented / apps, 2) if apps else None, round(prevented / (minutes / 90), 2) if minutes else None, rate, round(conceded / xg, 3) if xg else None, classification, "Sim" if apps >= 5 else "Não", round(player["overall"], 1), player["id"]])
    return rows


def player_card_values(players):
    ordered = sorted(players, key=lambda player: (-(player["yellow_cards"] + player["red_cards"]), -player["red_cards"], player["id"]))
    rows = []
    for rank, player in enumerate(ordered, 1):
        games = round(player["games"])
        yellow, red = round(player["yellow_cards"]), round(player["red_cards"])
        name = " ".join(part for part in (str(player.get("name") or "").strip(), str(player.get("surname") or "").strip()) if part) or f"ID {player['id']}"
        rows.append([rank, name, player["role"], player.get("team_name") or "Sem clube", player.get("league_name") or "Sem liga", games, yellow, red, yellow + red, round(yellow / games, 3) if games else None, round(red / games, 3) if games else None, player["id"]])
    return rows


def team_card_values(teams):
    ordered = sorted(teams, key=lambda team: (-(team["yellow_cards"] + team["red_cards"]), -team["red_cards"], team["name"]))
    rows = []
    for rank, team in enumerate(ordered, 1):
        games, yellow, red = team["all"]["games"], team["yellow_cards"], team["red_cards"]
        rows.append([rank, team["name"], team["league"], games, yellow, red, yellow + red, round(yellow / games, 3) if games else None, round(red / games, 3) if games else None])
    return rows


def overall_evolution_values(players):
    rows = []
    ordered = sorted(players, key=lambda player: (-player["overall"], player["id"]))
    for rank, player in enumerate(ordered, 1):
        current = round(player["overall"], 1)
        season_start = player["season_overall_history"][0] if player["season_overall_history"] else current
        career_start = player["career_overall_history"][0] if player["career_overall_history"] else season_start
        name = " ".join(part for part in (str(player.get("name") or "").strip(), str(player.get("surname") or "").strip()) if part) or f"ID {player['id']}"
        rows.append([rank, name, player["role"], player.get("team_name") or "Sem clube", player.get("league_name") or "Sem liga", round(season_start, 1), current, round(current - season_start, 1), round(career_start, 1), round(current - career_start, 1), player["id"]])
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


def styled_cell(ws, value, fill="neutral", number_format=None, alignment=CENTER, font=DATA_FONT):
    cell = WriteOnlyCell(ws, value=value)
    cell.font, cell.fill, cell.border, cell.alignment = font, FILLS[fill], BORDER, alignment
    if number_format:
        cell.number_format = number_format
    return cell


def write_sheet(wb, title, subtitle, headers, rows, number_formats, color_columns=None, inverse_columns=None, widths=None, language="pt"):
    ws = wb.create_sheet(title)
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A5"
    custom_widths = widths or {}
    output_headers = localized_headers(headers, language)
    for index, header in enumerate(headers, 1):
        header_text = header.casefold()
        base = 14
        if header_text in ("jogador", "goleiro"):
            base = 32
        elif header_text == "time":
            base = 30
        elif header_text in ("liga", "competição continental"):
            base = 26
        elif "para quem" in header_text:
            base = 30
        elif any(word in header_text for word in ("valor", "salário", "folha", "saldo", "gastos", "vendas")):
            base = 26
        elif "overall" in header_text:
            base = 20
        ws.column_dimensions[get_column_letter(index)].width = custom_widths.get(index - 1, base)
    ws.append([styled_cell(ws, title, "title", alignment=LEFT, font=TITLE_FONT)] + [styled_cell(ws, None, "title", alignment=LEFT, font=TITLE_FONT) for _ in output_headers[1:]])
    ws.append([styled_cell(ws, subtitle, "subtitle", alignment=LEFT)] + [styled_cell(ws, None, "subtitle", alignment=LEFT) for _ in output_headers[1:]])
    legend = [("Poor", "bad"), ("Average", "medium"), ("Good", "good"), ("Excellent", "excellent")] if language == "en" else [("Ruim", "bad"), ("Mediano", "medium"), ("Bom", "good"), ("Excelente", "excellent")]
    ws.append([styled_cell(ws, legend[index][0] if index < 4 else None, legend[index][1] if index < 4 else "subtitle", alignment=CENTER, font=Font(name="Aptos", size=10, bold=True)) for index in range(len(output_headers))])
    ws.append([styled_cell(ws, header, "header", alignment=CENTER, font=HEADER_FONT) for header in output_headers])
    color_columns = color_columns or []
    inverse_columns = inverse_columns or []
    distribution = percentiles_for_rows(rows, color_columns)
    for row in rows:
        cells = []
        for index, value in enumerate(row):
            fill = color_name(value, distribution.get(index), index in inverse_columns) if index in color_columns else "neutral"
            fmt = number_formats.get(index)
            cells.append(styled_cell(ws, value, fill, fmt, CENTER))
        ws.append(cells)
    last_row = 4 + len(rows)
    ws.auto_filter.ref = f"A4:{get_column_letter(len(output_headers))}{max(4, last_row)}"
    return ws


def make_dictionary_rows():
    return [
        ["Campo", "Fonte", "Cálculo", "Leitura"],
        ["Recorte antes/depois de 1º de janeiro", "general.date e matches2.date", "Data interna 184 é 01/01/2026", "Todas as métricas temporais usam este corte"],
        ["Expectativa — posição por overall", "teams2.team_in_time_value", "Ranking do overall de 1º de janeiro entre times da mesma liga", "Quanto menor a posição, maior a expectativa"],
        ["Overall do Footlord", "Atributos do jogador", "Média ponderada para jogadores de linha; atributo GKP para goleiros", "Referência técnica para ordenar jogadores e goleiros"],
        ["Cartões de jogadores", "players2.s_y_cards e players2.s_r_cards", "Soma dos cartões acumulados por competição", "Disciplinar: menos cartões recebe melhor classificação visual"],
        ["Cartões de times", "matches2.yellows e matches2.reds", "Soma dos cartões de cada lado em partidas concluídas", "Disciplinar: menos cartões recebe melhor classificação visual"],
        ["Evolução de overall", "players2.s_value_in_time e players2.value_in_time", "Overall atual menos o primeiro valor registrado na temporada e na carreira", "Positivo indica aumento; negativo indica queda"],
        ["Chutes e conversão", "matches2.shots_on e matches2.shots_off", "Chutes no alvo + fora; gols/chutes", "Volume, precisão e eficiência ofensiva do time"],
        ["Gols evitados vs xG", "matches2.x_goals e gols", "xG adversário menos gols sofridos", "Valor positivo indica desempenho acima do esperado"],
        ["Classificação de cores", "Valores de cada coluna", "Percentis relativos entre as linhas da aba", "Vermelho: ruim; amarelo: mediano; verde: bom; azul: excelente"],
        ["Valores financeiros", "Campos monetários do save", "Número inteiro sem símbolo nas células", "Unidade indicada somente no cabeçalho"],
        ["Ordenação mobile", "AutoFilter padrão", "Intervalo somente de dados, sem células mescladas", "Compatível com ordenação crescente/decrescente"],
    ]


def create_workbook(data, output, language="pt"):
    label = f"{data['club']} — data interna {data['current_day']}"
    player_rows, gk_rows, team_rows = player_values(data["players"]), goalkeeper_values(data["gks"]), [team_values(row) for row in data["teams"]]
    wb = Workbook(write_only=True)
    player_formats = {0: '0', 4: '0', 9: NUM_FORMAT_DEC1, 10: NUM_FORMAT_DEC1, 11: NUM_FORMAT_DEC1, 12: NUM_FORMAT_DEC1, 13: NUM_FORMAT_DEC1, 14: '0', 15: '0', 16: '0', 17: '0', 18: NUM_FORMAT_DEC3, 19: NUM_FORMAT_DEC3, 20: '0', 21: '0', 22: '0', 23: NUM_FORMAT_DEC3, 24: NUM_FORMAT_DEC3, 25: NUM_FORMAT_DEC3, 26: NUM_FORMAT_DEC2, 27: NUM_FORMAT_INT, 28: NUM_FORMAT_DEC1, 29: NUM_FORMAT_DEC1, 30: NUM_FORMAT_DEC1, 31: '0', 32: '0', 33: NUM_FORMAT_PCT, 34: NUM_FORMAT_PCT, 35: '0', 36: NUM_FORMAT_INT, 37: NUM_FORMAT_INT, 38: NUM_FORMAT_DEC1, 42: NUM_FORMAT_DEC1, 43: NUM_FORMAT_DEC1, 44: NUM_FORMAT_DEC1, 45: NUM_FORMAT_DEC1, 46: NUM_FORMAT_DEC1, 47: NUM_FORMAT_DEC1, 48: NUM_FORMAT_DEC1, 49: '0'}
    gk_formats = {0: '0', 4: '0', 5: NUM_FORMAT_INT, 6: '0', 7: NUM_FORMAT_DEC3, 8: NUM_FORMAT_DEC2, 9: NUM_FORMAT_DEC2, 10: NUM_FORMAT_DEC2, 11: NUM_FORMAT_DEC2, 12: NUM_FORMAT_DEC2, 13: NUM_FORMAT_DEC2, 14: NUM_FORMAT_PCT, 15: NUM_FORMAT_DEC3, 18: NUM_FORMAT_DEC1, 19: '0'}
    team_formats = {index: '0' for index in range(len(TIMES_HEADERS))}
    for index, header in enumerate(TIMES_HEADERS):
        if header in {"Média de idade", "Gols marcados/partida", "Gols sofridos/partida", "Chutes/partida", "Overall início do ano", "Overall fechamento agosto", "Overall final do ano", "Média overall adversário — vitória", "Média overall adversário — empate", "Média overall adversário — derrota", "Overall médio compras janeiro", "Overall médio vendas janeiro", "Overall médio compras agosto", "Overall médio vendas agosto"}:
            team_formats[index] = NUM_FORMAT_DEC2
        elif header in {"% no alvo", "Conversão de chutes"}:
            team_formats[index] = NUM_FORMAT_PCT
        elif header in {"Amarelos/partida", "Vermelhos/partida"}:
            team_formats[index] = NUM_FORMAT_DEC3
        elif any(word in header for word in ("Valor", "salário", "Saldo", "Gastos", "Vendas")):
            team_formats[index] = NUM_FORMAT_INT
    player_color_cols = [4, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 42, 43, 44, 45, 46, 47, 48]
    write_sheet(wb, "Players 2", f"Base completa de jogadores do save — {label}", PLAYERS_HEADERS, player_rows, player_formats, player_color_cols, inverse_columns=[4, 15, 16, 17, 18, 19, 29, 30, 32, 34, 37], language=language)
    write_sheet(wb, "Goleiros xG", f"Goleiros com aparições até a data do save — {label}", GK_HEADERS, gk_rows, gk_formats, [4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 18], inverse_columns=[9, 10, 15], language=language)
    summary_headers = ["Indicador", "Valor", "Interpretação"]
    matches = data["club_summary"]["matches"]
    summary_rows = [
        ["Clube controlado", data["club"], "Identificado no save"], ["Data interna do save", data["current_day"], "Calendário FM26 iniciado em 01/07/2025"], ["Jogadores na base", len(player_rows), "Todos os jogadores de players2"], ["Goleiros com aparição", len(gk_rows), "Goleiros com minutos ou aparição no save"], ["Times analisados", len(team_rows), "Times de clubes com competição definida"], ["Partidas realizadas", data["played_matches"], "Apenas state=2"], ["Partidas do clube", matches, "Partidas concluídas até a data do save"], ["Chutes do clube", data["club_summary"]["shots"], "Chutes no alvo + fora"], ["Chutes por partida", data["club_summary"]["shots"] / matches if matches else None, "Indicador agregado"], ["Gols do clube", data["club_summary"]["goals"], "Total em partidas concluídas"], ["xG sofrido pelo clube", data["club_summary"]["xga"], "Soma de xG adversário"], ["Gols sofridos pelo clube", data["club_summary"]["ga"], "Total em partidas concluídas"], ["Gols evitados vs xG", data["club_summary"]["xga"] - data["club_summary"]["ga"], "xG sofrido menos gols sofridos"],
    ]
    write_sheet(wb, "Resumo", f"Indicadores agregados do save — {label}", summary_headers, summary_rows, {1: NUM_FORMAT_DEC2}, [7, 8, 9, 10, 11, 12], language=language)
    dict_rows = make_dictionary_rows()
    write_sheet(wb, "Dicionário", "Metodologia e origem dos campos calculados", dict_rows[0], dict_rows[1:], {}, [], widths={0: 32, 1: 34, 2: 44, 3: 44}, language=language)
    team_color_cols = [index for index, header in enumerate(TIMES_HEADERS) if header not in {"Rank overall", "Time", "Liga", "Competição continental", "Maior placar aplicado no ano", "Para quem — maior placar ano", "Pior placar sofrido no ano", "Para quem — pior placar ano", "Maior placar aplicado antes", "Para quem — maior placar antes", "Pior placar sofrido antes", "Para quem — pior placar antes", "Maior placar aplicado depois", "Para quem — maior placar depois", "Pior placar sofrido depois", "Para quem — pior placar depois"}]
    team_inverse_cols = [TIMES_HEADERS.index(header) for header in ("Cartões amarelos no ano", "Cartões vermelhos no ano", "Cartões totais no ano", "Amarelos/partida", "Vermelhos/partida", "Derrotas no ano", "Derrotas antes", "Derrotas depois", "Maior sequência derrotas antes", "Maior sequência derrotas depois", "Gols sofridos no ano", "Gols sofridos/partida", "Gols sofridos antes", "Gols sofridos/partida antes", "Gols sofridos depois", "Gols sofridos/partida depois", "Jogos sem marcar", "Sem marcar antes", "Sem marcar depois")]
    write_sheet(wb, "Times", f"Análise de clubes, com corte em 1º de janeiro (data interna {JAN_CUTOFF}) — {label}", TIMES_HEADERS, team_rows, team_formats, team_color_cols, team_inverse_cols, language=language)
    wb.save(output)


def create_template(output, language="pt"):
    data = {"club": "Modelo limpo", "current_day": 0, "players": [], "gks": [], "teams": [], "club_summary": {"matches": 0, "shots": 0, "goals": 0, "xga": 0.0, "ga": 0}, "played_matches": 0}
    create_workbook(data, output, language)


def extract_db(save_path, temporary_dir):
    if save_path.suffix.lower() == ".db":
        return save_path
    if save_path.suffix.lower() != ".fl":
        raise ValueError("Envie um save FM26 nos formatos .fl ou .db")
    try:
        with zipfile.ZipFile(save_path) as archive:
            candidates = [info for info in archive.infolist() if not info.is_dir() and (info.filename.lower().endswith(".db") or info.filename.lower().endswith("database"))]
            if not candidates:
                raise ValueError("Não foi encontrado um banco SQLite dentro do arquivo .fl")
            candidate = max(candidates, key=lambda info: info.file_size)
            output = Path(temporary_dir) / "save.db"
            with archive.open(candidate) as source, output.open("wb") as destination:
                shutil.copyfileobj(source, destination)
            return output
    except zipfile.BadZipFile as error:
        raise ValueError("O arquivo .fl não pôde ser aberto como save FM26") from error


def validate_template(template_path):
    if not template_path:
        return
    try:
        workbook = load_workbook(template_path, read_only=True, data_only=False)
        required = {"Players 2", "Goleiros xG", "Resumo", "Dicionário", "Times"}
        missing = required - set(workbook.sheetnames)
        workbook.close()
        if missing:
            raise ValueError("O modelo não contém as abas obrigatórias: " + ", ".join(sorted(missing)))
    except ValueError:
        raise
    except Exception as error:
        raise ValueError("Não foi possível abrir o modelo XLSX") from error


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--save")
    parser.add_argument("--template")
    parser.add_argument("--output", required=True)
    parser.add_argument("--template-only", action="store_true")
    parser.add_argument("--language", choices=("pt", "en"), default="pt")
    args = parser.parse_args()
    output = Path(args.output)
    output.parent.mkdir(parents=True, exist_ok=True)
    if args.template_only:
        emit(20, "Criando modelo XLSX limpo")
        create_template(output, args.language)
        emit(100, "Modelo limpo criado")
        return
    if not args.save:
        raise ValueError("O save FM26 é obrigatório")
    validate_template(Path(args.template)) if args.template else None
    with tempfile.TemporaryDirectory(prefix="fm26-") as temporary_dir:
        emit(5, "Preparando o save FM26")
        db_path = extract_db(Path(args.save), temporary_dir)
        data = build_data(db_path)
        emit(76, "Montando as abas e a análise de times")
        create_workbook(data, output, args.language)
        emit(100, "Planilha Moneyball concluída")


def _running_in_streamlit() -> bool:
    try:
        from streamlit.runtime.scriptrunner import get_script_run_ctx
        return get_script_run_ctx() is not None
    except Exception:
        return False


def run_streamlit_app() -> None:
    """Interface web Streamlit — não altera a lógica de processamento."""
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
            "- **Resumo**\n"
            "- **Dicionário**\n"
            "- **Times**"
        )
        st.markdown("---")
        st.caption("CLI continua disponível: `python app.py --save arquivo.fl --output saida.xlsx`")
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
        generate = st.button("Gerar planilha", type="primary", use_container_width=True, disabled=uploaded is None)
    with col2:
        template_only = st.button("Só modelo vazio", use_container_width=True)

    # --- Modelo vazio ---
    if template_only:
        with tempfile.TemporaryDirectory(prefix="fm26-st-") as temporary_dir:
            out = Path(temporary_dir) / "Moneyball_modelo_limpo.xlsx"
            try:
                with st.spinner("Criando modelo limpo..."):
                    create_template(out, language)
                st.session_state["template_bytes"] = out.read_bytes()
                st.session_state["xlsx_bytes"] = None  # evita misturar com save
                st.session_state["xlsx_name"] = None
                st.session_state["xlsx_info"] = None
            except Exception as error:
                st.error(f"Erro: {error}")

    # --- Processar save ---
    if generate and uploaded is not None:
        progress = st.progress(0, text="Iniciando...")
        status = st.empty()

        def ui_emit(percent: int, message: str) -> None:
            progress.progress(min(100, max(0, percent)) / 100.0, text=message)
            status.info(message)

        # Redireciona emit() só durante este processamento
        global emit
        original_emit = emit
        emit = ui_emit  # type: ignore[assignment]

        try:
            with tempfile.TemporaryDirectory(prefix="fm26-st-") as temporary_dir:
                # Nome seguro (evita path estranho do upload)
                safe_name = Path(uploaded.name).name.replace(" ", "_")
                save_path = Path(temporary_dir) / safe_name
                save_path.write_bytes(uploaded.getvalue())
                out_name = f"Moneyball_{Path(safe_name).stem}.xlsx"
                output_path = Path(temporary_dir) / out_name

                ui_emit(5, "Preparando o save FM26")
                db_path = extract_db(save_path, temporary_dir)
                data = build_data(db_path)
                ui_emit(76, "Montando as abas e a análise de times")
                create_workbook(data, output_path, language)
                ui_emit(100, "Planilha Moneyball concluída")

                # Guarda em session_state para sobreviver ao rerun do botão de download
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

    # --- Botões de download (fora do if generate: sobrevivem ao clique) ---
    if st.session_state.get("xlsx_bytes"):
        info = st.session_state.get("xlsx_info") or {}
        st.success(
            f"**{info.get('club', 'Save')}** — data interna **{info.get('day', '?')}** · "
            f"{info.get('players', '?')} jogadores · {info.get('gks', '?')} goleiros · "
            f"{info.get('teams', '?')} times · arquivo ~{info.get('size_mb', '?')} MB"
        )
        st.download_button(
            label="⬇️ Baixar planilha Moneyball (.xlsx)",
            data=st.session_state["xlsx_bytes"],
            file_name=st.session_state.get("xlsx_name") or "Moneyball.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="dl_moneyball",
        )
        st.caption("Se o download não iniciar, use o botão de novo ou clique com o botão direito → salvar link.")

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
    try:
        main()
    except Exception as error:
        print(f"ERROR|{error}", file=sys.stderr, flush=True)
        sys.exit(1)
