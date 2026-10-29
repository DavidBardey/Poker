from datetime import datetime
from pathlib import Path
from flask import Flask, render_template, request, jsonify, send_file, redirect, url_for, session
from openpyxl import Workbook

app = Flask(__name__)
app.secret_key = "poker-salomon-secret"

BASE_DIR = Path(__file__).resolve().parent
EXPORTS_DIR = BASE_DIR / "exports"
EXPORTS_DIR.mkdir(exist_ok=True)

LANGS = ["fr", "es", "en"]

TEXTS = {
    "fr": {
        "title": "Poker Salomon",
        "subtitle": "Gestion des blinds, recaves, kills et gains",
        "start": "Commencer",
        "setup": "Configuration de la partie",
        "players": "Nombre joueurs",
        "duration": "Durée (min)",
        "buyin": "Cave",
        "buyin_money": "Prix d’une cave (argent)",
        "buyin_chips": "Jetons reçus par cave",
        "max_rebuys": "Recaves max",
        "paid_places": "Places payées",
        "winner_pct": "Pourcentage 1er",
        "second_pct": "Pourcentage 2e",
        "names": "Noms des joueurs",
        "calculate": "Calculer les blinds",
        "launch": "Lancer l'horloge",
        "back": "Retour",
        "clock": "PARTIE EN COURS",
        "level": "Niveau",
        "next": "Niveau suivant",
        "alive": "Joueurs en vie",
        "out": "Éliminés",
        "pause": "Pause",
        "resume": "Reprendre",
        "kill": "Kill / Pause",
        "previous": "Niveau précédent",
        "next_level": "Niveau suivant",
        "end_excel": "Fin de partie + Excel",
        "killed_player": "Joueur killé",
        "killer": "Joueur qui a fait le kill",
        "outcome": "Résultat",
        "rebuy": "Le joueur recave",
        "eliminated": "Le joueur est éliminé",
        "validate": "Valider",
        "cancel": "Annuler",
        "winner": "Félicitations",
        "pot_total": "Pot total",
        "request_error": "La requête a échoué. Réessayez ou annulez.",
        "invalid_outcome": "Résultat de kill invalide.",
        "menu": "Menu",
        "money_help": "Utilisé pour le pot et les gains, recaves comprises.",
        "chips_help": "Utilisé pour calculer les blinds, indépendamment du prix payé.",
        "setup_invalid": "Vérifiez les nombres saisis et leurs limites.",
        "setup_payout": "Les pourcentages des places payées doivent être positifs et leur somme doit être inférieure à 100 % avant la dernière place.",
        "setup_duplicate": "Chaque joueur doit avoir un nom différent.",
        "prize_pool": "Répartition des gains",
    },
    "es": {
        "title": "Poker Salomon",
        "subtitle": "Gestión de ciegas, recompras, eliminaciones y premios",
        "start": "Empezar",
        "setup": "Configuración de la partida",
        "players": "Número jugadores",
        "duration": "Duración (min)",
        "buyin": "Entrada",
        "buyin_money": "Precio de una entrada (dinero)",
        "buyin_chips": "Fichas recibidas por entrada",
        "max_rebuys": "Recompras máx.",
        "paid_places": "Puestos pagados",
        "winner_pct": "Porcentaje 1º",
        "second_pct": "Porcentaje 2º",
        "names": "Nombres de jugadores",
        "calculate": "Calcular ciegas",
        "launch": "Lanzar reloj",
        "back": "Volver",
        "clock": "PARTIDA EN CURSO",
        "level": "Nivel",
        "next": "Nivel siguiente",
        "alive": "Jugadores vivos",
        "out": "Eliminados",
        "pause": "Pausa",
        "resume": "Reanudar",
        "kill": "Eliminar / Pausa",
        "previous": "Nivel anterior",
        "next_level": "Nivel siguiente",
        "end_excel": "Fin de partida + Excel",
        "killed_player": "Jugador eliminado",
        "killer": "Jugador que eliminó",
        "outcome": "Resultado",
        "rebuy": "El jugador recompra",
        "eliminated": "El jugador queda eliminado",
        "validate": "Validar",
        "cancel": "Cancelar",
        "winner": "Felicitaciones",
        "pot_total": "Pozo total",
        "request_error": "La solicitud falló. Reintenta o cancela.",
        "invalid_outcome": "Resultado de eliminación inválido.",
        "menu": "Menú",
        "money_help": "Se usa para el pozo y los premios, incluidas las recompras.",
        "chips_help": "Se usa para calcular las ciegas, independientemente del precio pagado.",
        "setup_invalid": "Revisa los números introducidos y sus límites.",
        "setup_payout": "Los porcentajes pagados deben ser positivos y sumar menos del 100 % antes del último puesto.",
        "setup_duplicate": "Cada jugador debe tener un nombre distinto.",
        "prize_pool": "Reparto de premios",
    },
    "en": {
        "title": "Poker Salomon",
        "subtitle": "Blinds, rebuys, kills and payouts manager",
        "start": "Start",
        "setup": "Game setup",
        "players": "Players",
        "duration": "Duration (min)",
        "buyin": "Buy-in",
        "buyin_money": "Price per buy-in (money)",
        "buyin_chips": "Chips received per buy-in",
        "max_rebuys": "Max rebuys",
        "paid_places": "Paid places",
        "winner_pct": "1st place %",
        "second_pct": "2nd place %",
        "names": "Player names",
        "calculate": "Calculate blinds",
        "launch": "Launch clock",
        "back": "Back",
        "clock": "GAME IN PROGRESS",
        "level": "Level",
        "next": "Next level",
        "alive": "Players alive",
        "out": "Eliminated",
        "pause": "Pause",
        "resume": "Resume",
        "kill": "Kill / Pause",
        "previous": "Previous level",
        "next_level": "Next level",
        "end_excel": "End game + Excel",
        "killed_player": "Killed player",
        "killer": "Killer",
        "outcome": "Outcome",
        "rebuy": "Player rebuys",
        "eliminated": "Player is eliminated",
        "validate": "Validate",
        "cancel": "Cancel",
        "winner": "Congratulations",
        "pot_total": "Total pot",
        "request_error": "The request failed. Retry or cancel.",
        "invalid_outcome": "Invalid kill outcome.",
        "menu": "Menu",
        "money_help": "Used for the pot and payouts, including rebuys.",
        "chips_help": "Used to calculate blinds, independently of the price paid.",
        "setup_invalid": "Check the numbers entered and their limits.",
        "setup_payout": "Paid percentages must be positive and total less than 100% before the last paid place.",
        "setup_duplicate": "Each player must have a different name.",
        "prize_pool": "Prize distribution",
    },
}

STATE = {
    "configured": False,
    "players": [],
    "active_players": [],
    "eliminated_players": [],
    "events": [],
    "schedule": [],
    "detailed": [],
    "minutes_per_level": 15,
    "current_level": 1,
    "buyin_money": 50000,
    "buyin_chips": 1200,
    "max_rebuys": 1,
    "paid_places": 3,
    "winner_pct": 60,
    "second_pct": 30,
}


def lang():
    return session.get("lang", "fr")


def t():
    return TEXTS[lang()]


def clean_name(name):
    return str(name).strip()


def find_matching_player(name, players):
    cleaned = clean_name(name)
    for player in players:
        if clean_name(player) == cleaned:
            return player
    return None


def recommended_levels(n, total_chips):
    base = 8
    if n >= 6:
        base += 2
    if n >= 9:
        base += 1
    if total_chips >= 20000:
        base += 1
    if total_chips >= 40000:
        base += 1
    return min(base, 15)


def build_schedule(starting_stack, levels):
    allowed = [
        10, 15, 20, 25, 30, 40, 50, 75,
        100, 150, 200, 300, 400, 500,
        600, 800, 1000, 1200, 1500, 2000,
        2500, 3000, 4000, 5000, 8000, 10000
    ]

    def closest(x):
        return min(allowed, key=lambda v: abs(v - x))

    sb = max(10, starting_stack // 100)
    result = []
    used = set()

    for _ in range(levels):
        small = closest(sb)
        text = f"{small} / {small * 2}"
        if text not in used:
            result.append(text)
            used.add(text)
        sb *= 1.5

    while len(result) < levels:
        last = int(result[-1].split("/")[0].strip())
        small = closest(last * 1.5)
        text = f"{small} / {small * 2}"
        if text not in used:
            result.append(text)
            used.add(text)
        else:
            small = closest(last * 2)
            text = f"{small} / {small * 2}"
            if text not in used:
                result.append(text)
                used.add(text)
            else:
                # The predefined steps end at 10,000. Always make progress
                # instead of retrying the same capped blind forever.
                small = last * 2
                text = f"{small} / {small * 2}"
                result.append(text)
                used.add(text)

    return result[:levels]


def format_minutes(m):
    return f"{m // 60:02d}:{m % 60:02d}"


def compute_blinds(nb_players, duration, buyin_chips, max_rebuys):
    total_theoretical_chips = nb_players * buyin_chips * (1 + max_rebuys)
    levels = recommended_levels(nb_players, total_theoretical_chips)
    minutes_per_level = max(5, duration // levels)
    schedule = build_schedule(buyin_chips, levels)

    detailed = []
    start = 0

    for i, blind in enumerate(schedule, start=1):
        end = start + minutes_per_level
        detailed.append({
            "level": i,
            "blind": blind,
            "range": f"{format_minutes(start)} - {format_minutes(end)}"
        })
        start = end

    return schedule, detailed, minutes_per_level


def record_event(action, killed="", killer="", recave=""):
    STATE["events"].append({
        "date": datetime.now().strftime("%Y-%m-%d"),
        "time": datetime.now().strftime("%H:%M:%S"),
        "level": STATE["current_level"],
        "action": action,
        "killed": clean_name(killed),
        "killer": clean_name(killer),
        "rebuy": recave,
    })


def actual_rebuys():
    return sum(
        1 for e in STATE["events"]
        if e["action"] == "Kill" and e["rebuy"] == "Oui"
    )


def rebuys_for_player(player):
    player = clean_name(player)
    return sum(
        1 for e in STATE["events"]
        if e["action"] == "Kill"
        and clean_name(e["killed"]) == player
        and e["rebuy"] == "Oui"
    )


def total_pot():
    return STATE["buyin_money"] * (len(STATE["players"]) + actual_rebuys())


def ranking():
    eliminated = STATE["eliminated_players"]
    active = STATE["active_players"]

    result = []

    if len(active) == 1:
        result.append((1, active[0]))
        start_rank = 2
    else:
        start_rank = len(active) + 1

    eliminated_reversed = list(reversed(eliminated))

    for i, p in enumerate(eliminated_reversed, start=start_rank):
        result.append((i, p))

    return result


def payout_percentages():
    paid = STATE["paid_places"]
    first = STATE["winner_pct"]
    second = STATE["second_pct"]

    if paid == 1:
        return [100]

    if paid == 2:
        return [first, 100 - first]

    third = 100 - first - second
    return [first, second, third]


@app.route("/set-lang/<code>")
def set_lang(code):
    if code in LANGS:
        session["lang"] = code
    return redirect(request.referrer or url_for("index"))


@app.route("/api/lang/<code>", methods=["POST"])
def api_lang(code):
    if code not in LANGS:
        return jsonify({"error": "Unknown language"}), 400
    session["lang"] = code
    return jsonify(TEXTS[code])


@app.route("/")
def index():
    return render_template("index.html", t=t(), lang=lang())


@app.route("/setup", methods=["GET", "POST"])
def setup():
    if request.method == "POST":
        try:
            def number(key, minimum, maximum=None):
                value = int(request.form.get(key, ""))
                if value < minimum or (maximum is not None and value > maximum):
                    raise ValueError
                return value

            n = number("players", 2, 20)
            duration = number("duration", 15)
            buyin_money = number("buyin_money", 1)
            buyin_chips = number("buyin_chips", 1)
            max_rebuys = number("max_rebuys", 0)
            paid_places = number("paid_places", 1, 3)
            # Hidden, unused percentages must not prevent starting a game.
            winner_pct = number("winner_pct", 1, 100) if paid_places > 1 else 100
            second_pct = number("second_pct", 1, 99) if paid_places == 3 else 0
        except (ValueError, OverflowError):
            return render_template("setup.html", t=t(), lang=lang(),
                                   error=t()["setup_invalid"]), 400

        if paid_places > 1 and winner_pct + second_pct >= 100:
            return render_template("setup.html", t=t(), lang=lang(),
                                   error=t()["setup_payout"]), 400

        players = []
        seen = set()
        for i in range(1, n + 1):
            name = clean_name(request.form.get(f"player_{i}", "")) or f"Player {i}"
            if name.casefold() in seen:
                return render_template("setup.html", t=t(), lang=lang(),
                                       error=t()["setup_duplicate"]), 400
            seen.add(name.casefold())
            players.append(name)

        schedule, detailed, mpl = compute_blinds(
            nb_players=n,
            duration=duration,
            buyin_chips=buyin_chips,
            max_rebuys=max_rebuys
        )

        STATE.update({
            "configured": True,
            "players": players,
            "active_players": players.copy(),
            "eliminated_players": [],
            "events": [],
            "schedule": schedule,
            "detailed": detailed,
            "minutes_per_level": mpl,
            "current_level": 1,
            "buyin_money": buyin_money,
            "buyin_chips": buyin_chips,
            "max_rebuys": max_rebuys,
            "paid_places": paid_places,
            "winner_pct": winner_pct,
            "second_pct": second_pct,
        })

        record_event("Début")
        return render_template("summary.html", t=t(), lang=lang(), state=STATE)

    return render_template("setup.html", t=t(), lang=lang())


@app.route("/clock")
def clock():
    if not STATE["configured"]:
        return redirect(url_for("setup"))
    return render_template(
        "clock.html",
        t=t(),
        lang=lang(),
        state=STATE,
        texts=TEXTS[lang()]
    )


@app.route("/api/level", methods=["POST"])
def api_level():
    data = request.get_json() or {}
    direction = data.get("direction", "")

    current = int(STATE.get("current_level", 1))
    max_level = max(1, len(STATE.get("schedule", [])))

    if direction == "next":
        STATE["current_level"] = min(current + 1, max_level)
    elif direction == "previous":
        STATE["current_level"] = max(current - 1, 1)

    if STATE["current_level"] < 1:
        STATE["current_level"] = 1

    return jsonify({
        "ok": True,
        "state": STATE
    })


@app.route("/api/kill", methods=["POST"])
def api_kill():
    data = request.get_json() or {}

    killed_raw = data.get("killed", "")
    killer_raw = data.get("killer", "")
    outcome = clean_name(data.get("outcome", ""))

    killed = find_matching_player(killed_raw, STATE["active_players"])
    killer = find_matching_player(killer_raw, STATE["active_players"])

    active_names = [clean_name(p) for p in STATE["active_players"]]

    if not killed:
        return jsonify({
            "ok": False,
            "error": f"{clean_name(killed_raw)} n'est pas actif. Joueurs actifs : {', '.join(active_names)}"
        })

    if not killer:
        return jsonify({
            "ok": False,
            "error": f"{clean_name(killer_raw)} n'est pas actif. Joueurs actifs : {', '.join(active_names)}"
        })

    if clean_name(killed) == clean_name(killer):
        return jsonify({
            "ok": False,
            "error": "auto-kill impossible"
        })

    if outcome not in ("rebuy", "out"):
        return jsonify({"ok": False, "error": t()["invalid_outcome"]}), 400

    if outcome == "rebuy":
        if rebuys_for_player(killed) >= STATE["max_rebuys"]:
            return jsonify({
                "ok": False,
                "error": "rebuy limit reached"
            })

        record_event("Kill", killed, killer, "Oui")

    else:
        record_event("Kill", killed, killer, "Non")

        matched_killed = find_matching_player(killed, STATE["active_players"])
        if matched_killed in STATE["active_players"]:
            STATE["active_players"].remove(matched_killed)

        if killed not in STATE["eliminated_players"]:
            STATE["eliminated_players"].append(killed)

        record_event("Élimination", killed, killer, "Non")

    return jsonify({
        "ok": True,
        "state": STATE
    })


@app.route("/api/final")
def api_final():
    ranks = ranking()
    pot = total_pot()
    percents = payout_percentages()

    payouts = []

    for rank, player in ranks:
        if rank <= len(percents):
            pct = percents[rank - 1]
            payouts.append({
                "rank": rank,
                "player": player,
                "percent": pct,
                "amount": round(pot * pct / 100),
            })

    return jsonify({
        "winner": STATE["active_players"][0] if len(STATE["active_players"]) == 1 else "",
        "pot": pot,
        "payouts": payouts,
        "buyin_money": STATE["buyin_money"],
        "buyin_chips": STATE["buyin_chips"],
        "actual_rebuys": actual_rebuys(),
    })


@app.route("/export")
def export():
    wb = Workbook()

    ws = wb.active
    ws.title = "Journal"
    ws.append([
        "Date",
        "Heure",
        "Niveau",
        "Action",
        "Joueur killé",
        "Killeur",
        "Recave"
    ])

    for e in STATE["events"]:
        ws.append([
            e["date"],
            e["time"],
            e["level"],
            e["action"],
            e["killed"],
            e["killer"],
            e["rebuy"]
        ])

    ws2 = wb.create_sheet("Classement")
    ws2.append(["Rang", "Joueur"])

    for r, p in ranking():
        ws2.append([r, p])

    pot = total_pot()
    percents = payout_percentages()
    ranks = ranking()

    ws3 = wb.create_sheet("Gains")
    ws3.append(["Pot total", pot])
    ws3.append(["Cave argent", STATE["buyin_money"]])
    ws3.append(["Cave jetons", STATE["buyin_chips"]])
    ws3.append(["Nombre de joueurs", len(STATE["players"])])
    ws3.append(["Nombre de recaves réelles", actual_rebuys()])
    ws3.append([])
    ws3.append(["Rang", "Joueur", "Pourcentage", "Gain"])

    for rank, player in ranks:
        if rank <= len(percents):
            pct = percents[rank - 1]
            ws3.append([
                rank,
                player,
                pct,
                round(pot * pct / 100)
            ])

    ws4 = wb.create_sheet("Résumé joueurs")
    ws4.append([
        "Joueur",
        "Kills réalisés",
        "Kills subis",
        "Recaves",
        "Statut"
    ])

    winner = STATE["active_players"][0] if len(STATE["active_players"]) == 1 else ""

    for p in STATE["players"]:
        kills_done = sum(
            1 for e in STATE["events"]
            if clean_name(e["killer"]) == clean_name(p) and e["action"] == "Kill"
        )

        kills_taken = sum(
            1 for e in STATE["events"]
            if clean_name(e["killed"]) == clean_name(p) and e["action"] == "Kill"
        )

        reb = rebuys_for_player(p)

        if winner and clean_name(winner) == clean_name(p):
            status = "Vainqueur"
        elif p in STATE["eliminated_players"]:
            status = "Éliminé"
        else:
            status = "En jeu"

        ws4.append([
            p,
            kills_done,
            kills_taken,
            reb,
            status
        ])

    ws5 = wb.create_sheet("Paramètres")
    ws5.append(["Paramètre", "Valeur"])
    ws5.append(["Cave argent", STATE["buyin_money"]])
    ws5.append(["Cave jetons", STATE["buyin_chips"]])
    ws5.append(["Recaves max", STATE["max_rebuys"]])
    ws5.append(["Places payées", STATE["paid_places"]])
    ws5.append(["Pourcentage 1er", STATE["winner_pct"]])
    ws5.append(["Pourcentage 2e", STATE["second_pct"]])
    ws5.append(["Recaves réelles", actual_rebuys()])
    ws5.append(["Pot total", total_pot()])

    file = EXPORTS_DIR / f"poker_summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    wb.save(file)

    return send_file(file, as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True)
