import streamlit as st
import tempfile
import sqlite3
import zipfile
import os
import sys
from datetime import datetime, timedelta
from collections import defaultdict
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

# ==============================================================================
# 1. FUNÇÕES DE PROCESSAMENTO (SEU CÓDIGO ORIGINAL)
# ==============================================================================

COLOR_TITLE = "0B3D2E"
COLOR_SUBTITLE = "EAF3EE"
COLOR_HEADER = "0E4A36"
COLOR_RUIM = "F4CCCC"
COLOR_MEDIANO = "FFF2CC"
COLOR_BOM = "D9EAD3"
COLOR_EXCELENTE = "CFE2F3"

thin_border = Border(
    left=Side(style="thin", color="000000"),
    right=Side(style="thin", color="000000"),
    top=Side(style="thin", color="000000"),
    bottom=Side(style="thin", color="000000"),
)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left = Alignment(horizontal="left", vertical="center", wrap_text=True)

def parse_stat(s):
    if not s: return {}, 0
    d, total = {}, 0
    for part in s.split(","):
        if ":" in part:
            k, v = part.split(":", 1)
            try:
                val = float(v) if "." in v else int(v)
                d[k] = val
                total += val
            except ValueError: pass
    return d, total

def parse_ratings_str(s):
    if not s: return [], 0.0
    ratings = []
    for part in s.split(","):
        if ":" in part:
            _, vals = part.split(":", 1)
            for v in vals.split("|"):
                try: ratings.append(float(v))
                except ValueError: pass
    avg = sum(ratings) / len(ratings) if ratings else 0.0
    return ratings, avg

def parse_match_ratings(s):
    if not s: return []
    out = []
    for part in s.split(";"):
        if "|" in part:
            pid_s, rat_s = part.split("|", 1)
            try: out.append((int(pid_s), float(rat_s)))
            except ValueError: pass
    return out

def get_age(birth_date, current_date):
    if birth_date is None: return None
    start = datetime(2025, 7, 1)
    birth = start + timedelta(days=birth_date)
    current = start + timedelta(days=current_date)
    return (current - birth).days // 365

def contract_years(exp, current_date):
    if exp is None or exp <= current_date: return 0
    return round((exp - current_date) / 365.0, 1)

def color_quartile(v, q1, q2, q3, higher_better=True):
    if higher_better:
        if v >= q3: return COLOR_EXCELENTE
        if v >= q2: return COLOR_BOM
        if v >= q1: return COLOR_MEDIANO
        return COLOR_RUIM
    else:
        if v <= q1: return COLOR_EXCELENTE
        if v <= q2: return COLOR_BOM
        if v <= q3: return COLOR_MEDIANO
        return COLOR_RUIM

def style_header(cell, kind="header"):
    cell.alignment = center
    cell.border = thin_border
    if kind == "title":
        cell.fill = PatternFill(patternType="solid", fgColor=COLOR_TITLE)
        cell.font = Font(bold=True, color="FFFFFF", size=14)
    elif kind == "subtitle":
        cell.fill = PatternFill(patternType="solid", fgColor=COLOR_SUBTITLE)
        cell.font = Font(size=10)
    elif kind == "header":
        cell.fill = PatternFill(patternType="solid", fgColor=COLOR_HEADER)
        cell.font = Font(bold=True, color="FFFFFF", size=9)

def extract_db_from_fl(fl_path, tmpdir):
    with zipfile.ZipFile(fl_path, "r") as z:
        dbs = [n for n in z.namelist() if n.endswith(".db") and not n.startswith("temp_")]
        if not dbs: dbs = [n for n in z.namelist() if n.endswith(".db")]
        if not dbs: raise RuntimeError("Nenhum .db encontrado dentro do .fl")
        dbs.sort(key=lambda x: (0 if x.startswith("save_") else 1, x))
        target = dbs[0]
        z.extract(target, tmpdir)
        return os.path.join(tmpdir, target)

def build_moneyball(db_path, output_path, club_hint="NAC Breda"):
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    cur.execute("SELECT * FROM general LIMIT 1")
    general = dict(cur.fetchone())
    current_date = general["date"]

    cur.execute("SELECT id, name FROM teams2 WHERE name LIKE ?", (f"%{club_hint}%",))
    rows = cur.fetchall()
    user_team_id = rows[0]["id"] if rows else 2947
    user_team_name = rows[0]["name"] if rows else club_hint

    nations = {r["id"]: r["name"] for r in cur.execute("SELECT id, name FROM nations2")}
    leagues = {r["id"]: dict(r) for r in cur.execute("SELECT id, name, division, nation_id FROM leagues2")}
    champs = {r["id"]: dict(r) for r in cur.execute("SELECT id, name, league_id FROM champs2")}

    teams = {}
    for r in cur.execute("SELECT id, name, champ_id, nation_id, bank_balance, is_national_team FROM teams2 WHERE is_national_team=0"):
        t = dict(r)
        lid = champs.get(t["champ_id"], {}).get("league_id")
        t["league_id"] = lid
        t["league_name"] = leagues.get(lid, {}).get("name", "") if lid else ""
        t["division"] = leagues.get(lid, {}).get("division", "") if lid else ""
        t["country"] = nations.get(t["nation_id"], "")
        teams[t["id"]] = t

    players_raw = [
        dict(r) for r in cur.execute(
            """SELECT id, team_id, nation_id, nationality, name, surname, role, roles_level,
                      birth_date, PAC, SHO, PAS, DEF, PHY, MEN, GKP, salary, contract_expiration,
                      temp_value, temp_mk_value, status, transfer_status, loan_status, on_loan,
                      s_matches, s_played, s_goals, s_assists, s_g_conceded, s_clean_sheets, s_ratings
               FROM players2 WHERE retired=0"""
        )
    ]

    players = []
    for p in players_raw:
        apps_str = p["s_played"] or p["s_matches"] or ""
        _, apps = parse_stat(apps_str)
        _, goals = parse_stat(p["s_goals"] or "")
        _, assists = parse_stat(p["s_assists"] or "")
        _, clean_sheets = parse_stat(p["s_clean_sheets"] or "")
        _, conceded = parse_stat(p["s_g_conceded"] or "")
        ratings_list, avg_rating = parse_ratings_str(p["s_ratings"] or "")

        sec_roles = ""
        if p["roles_level"]:
            main = p["role"] or ""
            parts = []
            for rp in p["roles_level"].split(","):
                if ":" in rp:
                    rname, rval = rp.split(":", 1)
                    try:
                        rv = float(rval)
                        if rname != main and rv > 0: parts.append(f"{rname}:{rv:.1f}")
                    except ValueError: pass
            sec_roles = ", ".join(parts[:5])

        age = get_age(p["birth_date"], current_date)
        overall = round(p["temp_value"] or 0, 1)
        team = teams.get(p["team_id"], {})
        n_ratings = len(ratings_list)
        minutos = n_ratings * 85 if n_ratings else (apps * 85 if apps else 0)
        ga = goals + assists

        players.append({
            "id": p["id"], "name": f"{p['name'] or ''} {p['surname'] or ''}".strip(),
            "role": p["role"] or "", "sec_roles": sec_roles, "age": age,
            "team": team.get("name", ""), "team_id": p["team_id"],
            "league": team.get("league_name", ""), "division": team.get("division", ""),
            "country": team.get("country", "") or p.get("nationality") or nations.get(p["nation_id"], ""),
            "overall": overall, "apps": apps, "goals": goals, "assists": assists, "ga": ga,
            "g_p": round(goals / apps, 3) if apps else 0, "a_p": round(assists / apps, 3) if apps else 0,
            "ga_p": round(ga / apps, 3) if apps else 0, "avg_rating": round(avg_rating, 2) if avg_rating else None,
            "minutos": minutos, "min_p": round(minutos / apps, 1) if apps else 0,
            "min_g": round(minutos / goals, 1) if goals else None,
            "min_a": round(minutos / assists, 1) if assists else None,
            "clean_sheets": clean_sheets, "mk_value": p["temp_mk_value"] or 0,
            "salary": p["salary"] or 0, "contract": contract_years(p["contract_expiration"], current_date),
            "status": p["status"] or "", "emprestimo": "Sim" if (p["on_loan"] or p["loan_status"]) else "Não",
            "transfer_list": "Sim" if (p["transfer_status"] or 0) > 0 else "Não",
            "PAC": round(p["PAC"] or 0, 1), "SHO": round(p["SHO"] or 0, 1), "PAS": round(p["PAS"] or 0, 1),
            "DEF": round(p["DEF"] or 0, 1), "PHY": round(p["PHY"] or 0, 1), "MEN": round(p["MEN"] or 0, 1),
            "GKP": round(p["GKP"] or 0, 1), "conceded": conceded,
        })

    players = [p for p in players if p["apps"] > 0 or (p["overall"] or 0) >= 70]
    players.sort(key=lambda x: (-(x["overall"] or 0), -x["apps"]))
    for i, p in enumerate(players): p["rank"] = i + 1

    mvp_count = defaultdict(int)
    pior_count = defaultdict(int)
    matches_with_rating = defaultdict(int)

    for row in cur.execute("SELECT ratings_1, ratings_2 FROM matches2 WHERE state=2"):
        for ratings in (parse_match_ratings(row[0]), parse_match_ratings(row[1])):
            if len(ratings) < 2: continue
            for pid, _ in ratings: matches_with_rating[pid] += 1
            max_r = max(r for _, r in ratings)
            min_r = min(r for _, r in ratings)
            for pid, r in ratings:
                if r == max_r: mvp_count[pid] += 1
                if r == min_r: pior_count[pid] += 1

    for p in players:
        pid = p["id"]
        p["mvp"] = mvp_count.get(pid, 0)
        p["pior"] = pior_count.get(pid, 0)
        n = matches_with_rating.get(pid, 0)
        p["mvp_pct"] = round(100.0 * p["mvp"] / n, 1) if n else None
        p["pior_pct"] = round(100.0 * p["pior"] / n, 1) if n else None

    matches = [
        dict(r) for r in cur.execute(
            """SELECT date, team_1_id, team_2_id, goals_1, goals_2, x_goals_1, x_goals_2,
                      shots_on_1, shots_off_1, shots_on_2, shots_off_2, state
               FROM matches2 WHERE state=2"""
        )
    ]

    team_xg_against = defaultdict(list)
    for m in matches:
        if m.get("x_goals_1") is not None: team_xg_against[m["team_2_id"]].append((m["x_goals_1"], m.get("goals_1") or 0))
        if m.get("x_goals_2") is not None: team_xg_against[m["team_1_id"]].append((m["x_goals_2"], m.get("goals_1") or 0))
    team_avg_xg = {tid: sum(x[0] for x in lst) / len(lst) for tid, lst in team_xg_against.items() if lst}

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Players 2"
    headers = [
        "Rank", "Jogador", "Posição", "Roles secundários (0-10)", "Idade", "Time", "Liga", "Divisão", "País",
        "Overall do Footlord (0-100)", "Jogos no ano", "Gols no ano", "Assistências no ano", "G+A no ano",
        "Gols/partida", "Assistências/partida", "G+A/partida", "Média de avaliação", "Minutos estimados",
        "Minutos/partida", "Minutos/gol", "Minutos/assistência", "MVP do time (vezes)", "Pior nota do time (vezes)",
        "MVP %", "Pior nota %", "Clean sheets", "Valor de mercado em €", "Salário em €", "Contrato (anos)",
        "Status", "Empréstimo", "Lista de transferências", "PAC", "SHO", "PAS", "DEF", "PHY", "MEN", "GKP", "ID interno",
    ]
    n_cols = len(headers)

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    style_header(ws.cell(1, 1, "Players 2"), "title")
    for c in range(2, n_cols + 1): style_header(ws.cell(1, c), "title")

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n_cols)
    style_header(ws.cell(2, 1, f"Base completa — {user_team_name} — data interna {current_date}"), "subtitle")
    for c in range(2, n_cols + 1): style_header(ws.cell(2, c), "subtitle")

    for i, (txt, col) in enumerate([("Ruim", COLOR_RUIM), ("Mediano", COLOR_MEDIANO), ("Bom", COLOR_BOM), ("Excelente", COLOR_EXCELENTE)], 1):
        cell = ws.cell(3, i, txt)
        cell.fill = PatternFill(patternType="solid", fgColor=col)
        cell.border = thin_border
        cell.alignment = center
        cell.font = Font(size=9, bold=True)
    for c in range(5, n_cols + 1):
        ws.cell(3, c).fill = PatternFill(patternType="solid", fgColor=COLOR_SUBTITLE)
        ws.cell(3, c).border = thin_border

    for col, h in enumerate(headers, 1): style_header(ws.cell(4, col, h), "header")

    higher_better = {
        9: True, 10: True, 11: True, 12: True, 13: True, 14: True, 15: True, 16: True, 17: True,
        18: True, 19: True, 20: False, 21: False, 22: True, 23: False, 24: True, 25: False, 26: True,
        27: True, 28: False, 29: True, 33: True, 34: True, 35: True, 36: True, 37: True, 38: True, 39: True,
    }

    col_vals = defaultdict(list)
    for p in players:
        row_vals = [
            p["rank"], p["name"], p["role"], p["sec_roles"], p["age"], p["team"], p["league"], p["division"], p["country"],
            p["overall"], p["apps"], p["goals"], p["assists"], p["ga"], p["g_p"], p["a_p"], p["ga_p"], p["avg_rating"],
            p["minutos"], p["min_p"], p["min_g"], p["min_a"], p["mvp"], p["pior"], p["mvp_pct"], p["pior_pct"],
            p["clean_sheets"], p["mk_value"], p["salary"], p["contract"], p["status"], p["emprestimo"], p["transfer_list"],
            p["PAC"], p["SHO"], p["PAS"], p["DEF"], p["PHY"], p["MEN"], p["GKP"], p["id"],
        ]
        for i, v in enumerate(row_vals):
            if i in higher_better and isinstance(v, (int, float)): col_vals[i].append(v)

    quartiles = {}
    for i, vals in col_vals.items():
        s = sorted(vals)
        n = len(s)
        if n: quartiles[i] = (s[n // 4], s[n // 2], s[3 * n // 4])

    for row_idx, p in enumerate(players):
        r = row_idx + 5
        vals = [
            p["rank"], p["name"], p["role"], p["sec_roles"], p["age"], p["team"], p["league"], p["division"], p["country"],
            p["overall"], p["apps"], p["goals"], p["assists"], p["ga"], p["g_p"], p["a_p"], p["ga_p"], p["avg_rating"],
            p["minutos"], p["min_p"], p["min_g"], p["min_a"], p["mvp"], p["pior"], p["mvp_pct"], p["pior_pct"],
            p["clean_sheets"], p["mk_value"], p["salary"], p["contract"], p["status"], p["emprestimo"], p["transfer_list"],
            p["PAC"], p["SHO"], p["PAS"], p["DEF"], p["PHY"], p["MEN"], p["GKP"], p["id"],
        ]
        for col, v in enumerate(vals, 1):
            cell = ws.cell(r, col, v if v is not None else "")
            cell.border = thin_border
            cell.alignment = center
            cell.font = Font(size=8)
            idx = col - 1
            if idx in quartiles and isinstance(v, (int, float)):
                q1, q2, q3 = quartiles[idx]
                color = color_quartile(v, q1, q2, q3, higher_better[idx])
                cell.fill = PatternFill(patternType="solid", fgColor=color)

    widths = [6, 18, 8, 22, 6, 16, 14, 8, 12, 10, 8, 8, 10, 8, 9, 11, 9, 10, 10, 10, 9, 11, 10, 12, 7, 8, 8, 12, 10, 9, 10, 8, 10, 6, 6, 6, 6, 6, 6, 6, 8]
    for i, w in enumerate(widths, 1): ws.column_dimensions[get_column_letter(i)].width = w
    ws.auto_filter.ref = f"A4:{get_column_letter(n_cols)}{4 + len(players)}"
    ws.freeze_panes = "A5"

    wb.save(output_path)
    conn.close()
    return output_path

# ==============================================================================
# 2. INTERFACE STREAMLIT (PÁGINA WEB LOCAL)
# ==============================================================================

st.set_page_config(page_title="Footlord Moneyball", page_icon="⚽", layout="centered")

st.title("⚽ Footlord Moneyball Generator")
st.write("Suba o arquivo do seu save (`.fl` ou `.db`) para gerar a planilha de análises estilo Moneyball.")

st.markdown("---")

# Formulário da Interface
uploaded_file = st.file_uploader("Selecione o arquivo do save (.fl)", type=["fl", "db"])
club_name = st.text_input("Nome do seu Clube no jogo", value="NAC Breda")

st.markdown("---")

if uploaded_file is not None:
    if st.button("🚀 Gerar Planilha Excel", type="primary", use_container_width=True):
        with st.spinner("Lendo arquivo do save e processando estatísticas..."):
            try:
                # Criando pasta temporária isolada para processamento
                with tempfile.TemporaryDirectory() as tmpdir:
                    save_path = Path(tmpdir) / uploaded_file.name
                    
                    # Salva o arquivo enviado no disco temporário
                    with open(save_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())
                    
                    # Extrai o DB se for .fl
                    if save_path.suffix.lower() == ".fl":
                        db_path = extract_db_from_fl(str(save_path), tmpdir)
                    else:
                        db_path = str(save_path)
                    
                    # Caminho final da planilha gerada
                    output_excel_path = Path(tmpdir) / f"Moneyball_{save_path.stem}.xlsx"
                    
                    # Roda o gerador
                    build_moneyball(db_path, str(output_excel_path), club_hint=club_name)
                    
                    # Lê o arquivo Excel para o botão de download
                    with open(output_excel_path, "rb") as excel_file:
                        excel_bytes = excel_file.read()

                st.success("Planilha gerada com sucesso! Clique no botão abaixo para baixar.")
                
                # Botão de Download do Excel
                st.download_button(
                    label="📥 Baixar Planilha (.xlsx)",
                    data=excel_bytes,
                    file_name=f"Moneyball_{Path(uploaded_file.name).stem}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
            except Exception as e:
                st.error(f"Ocorreu um erro ao processar o save: {str(e)}")
else:
    st.info("Aguardando o envio do arquivo do save para liberar o botão de geração.")
