from __future__ import annotations

import io
import os
import secrets
import sqlite3
from datetime import datetime
from functools import wraps
from pathlib import Path

from flask import (
    Flask,
    abort,
    flash,
    g,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)
from openpyxl import Workbook, load_workbook
from werkzeug.security import check_password_hash, generate_password_hash


BASE_DIR = Path(__file__).resolve().parent
DEFAULT_EXCEL = BASE_DIR / "DATABASE GLOBALE TPS ELETTRONICA.xlsx"


SCHEMA = """
CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT NOT NULL UNIQUE COLLATE NOCASE,
    full_name TEXT NOT NULL,
    password_hash TEXT NOT NULL,
    role TEXT NOT NULL CHECK (role IN ('admin', 'operatore')),
    active INTEGER NOT NULL DEFAULT 1 CHECK (active IN (0, 1)),
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS components (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    code TEXT NOT NULL UNIQUE COLLATE NOCASE,
    category TEXT,
    value TEXT,
    description TEXT NOT NULL,
    laboratory TEXT,
    cabinet TEXT,
    door TEXT,
    shelf TEXT,
    drawer TEXT,
    quantity INTEGER NOT NULL DEFAULT 0 CHECK (quantity >= 0),
    reorder_threshold INTEGER NOT NULL DEFAULT 0 CHECK (reorder_threshold >= 0),
    target_quantity INTEGER NOT NULL DEFAULT 0 CHECK (target_quantity >= 0),
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS movements (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    component_id INTEGER NOT NULL REFERENCES components(id),
    movement_type TEXT NOT NULL CHECK (
        movement_type IN ('prelievo', 'restituzione', 'carico', 'rettifica_carico', 'rettifica_scarico')
    ),
    quantity INTEGER NOT NULL CHECK (quantity > 0),
    delta INTEGER NOT NULL CHECK (delta != 0),
    balance_after INTEGER NOT NULL CHECK (balance_after >= 0),
    recipient TEXT,
    note TEXT,
    user_id INTEGER NOT NULL REFERENCES users(id),
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX IF NOT EXISTS idx_components_code ON components(code);
CREATE INDEX IF NOT EXISTS idx_components_category ON components(category);
CREATE INDEX IF NOT EXISTS idx_movements_component ON movements(component_id);
CREATE INDEX IF NOT EXISTS idx_movements_created ON movements(created_at DESC);

CREATE TRIGGER IF NOT EXISTS movements_no_update
BEFORE UPDATE ON movements
BEGIN
    SELECT RAISE(ABORT, 'Lo storico movimenti non può essere modificato');
END;

CREATE TRIGGER IF NOT EXISTS movements_no_delete
BEFORE DELETE ON movements
BEGIN
    SELECT RAISE(ABORT, 'Lo storico movimenti non può essere cancellato');
END;
"""


def create_app(test_config: dict | None = None) -> Flask:
    app = Flask(__name__, instance_relative_config=True)
    app.config.from_mapping(
        SECRET_KEY=os.environ.get("MAGAZZINO_SECRET_KEY", secrets.token_hex(32)),
        DATABASE=str(Path(app.instance_path) / "magazzino.db"),
        EXCEL_SOURCE=str(DEFAULT_EXCEL),
        CSRF_ENABLED=True,
    )
    if test_config:
        app.config.update(test_config)

    Path(app.instance_path).mkdir(parents=True, exist_ok=True)

    app.teardown_appcontext(close_db)
    app.before_request(load_current_user)
    app.before_request(ensure_csrf)
    app.context_processor(inject_helpers)

    register_routes(app)

    with app.app_context():
        initialize_database()

    return app


def get_db() -> sqlite3.Connection:
    if "db" not in g:
        g.db = sqlite3.connect(current_app_config("DATABASE"), timeout=10)
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA foreign_keys = ON")
    return g.db


def current_app_config(name: str):
    from flask import current_app

    return current_app.config[name]


def close_db(_error=None) -> None:
    db = g.pop("db", None)
    if db is not None:
        db.close()


def initialize_database() -> None:
    db = get_db()
    db.executescript(SCHEMA)
    count = db.execute("SELECT COUNT(*) FROM components").fetchone()[0]
    if count == 0:
        source = Path(current_app_config("EXCEL_SOURCE"))
        if source.exists():
            import_excel(source, db)
    db.commit()


def clean_cell(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def non_negative_int(value, field_name: str) -> int:
    try:
        parsed = int(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{field_name}: inserire un numero intero") from exc
    if parsed < 0:
        raise ValueError(f"{field_name}: il valore non può essere negativo")
    return parsed


def positive_int(value, field_name: str = "Quantità") -> int:
    parsed = non_negative_int(value, field_name)
    if parsed == 0:
        raise ValueError(f"{field_name}: il valore deve essere maggiore di zero")
    return parsed


def import_excel(source: Path, db: sqlite3.Connection) -> int:
    workbook = load_workbook(source, data_only=True, read_only=True)
    sheet = workbook["COMPONENTI"]
    rows = []
    for values in sheet.iter_rows(min_row=3, values_only=True):
        code = clean_cell(values[0])
        if not code:
            continue
        quantity = non_negative_int(values[9] or 0, f"Quantità di {code}")
        threshold = non_negative_int(values[10] or 0, f"Soglia di {code}")
        rows.append(
            (
                code,
                clean_cell(values[1]),
                clean_cell(values[2]),
                clean_cell(values[3]) or code,
                clean_cell(values[4]),
                clean_cell(values[5]),
                clean_cell(values[6]),
                clean_cell(values[7]),
                clean_cell(values[8]),
                quantity,
                threshold,
                quantity,
            )
        )
    db.executemany(
        """
        INSERT INTO components (
            code, category, value, description, laboratory, cabinet, door,
            shelf, drawer, quantity, reorder_threshold, target_quantity
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        rows,
    )
    return len(rows)


def load_current_user() -> None:
    user_id = session.get("user_id")
    g.user = None
    if user_id is not None:
        g.user = get_db().execute(
            "SELECT * FROM users WHERE id = ? AND active = 1", (user_id,)
        ).fetchone()
        if g.user is None:
            session.clear()


def ensure_csrf():
    if "csrf_token" not in session:
        session["csrf_token"] = secrets.token_urlsafe(32)
    if request.method == "POST" and current_app_config("CSRF_ENABLED"):
        supplied = request.form.get("csrf_token", "")
        if not secrets.compare_digest(supplied, session["csrf_token"]):
            abort(400, "Token di sicurezza non valido. Ricaricare la pagina.")


def inject_helpers() -> dict:
    return {
        "csrf_token": session.get("csrf_token", ""),
        "stock_status": stock_status,
    }


def stock_status(component) -> tuple[str, str]:
    quantity = component["quantity"]
    threshold = component["reorder_threshold"]
    target = component["target_quantity"]
    if quantity == 0:
        return "Esaurito", "danger"
    if threshold > 0 and quantity <= threshold:
        return "Da riordinare", "danger"
    if threshold > 0:
        warning_band = max(1, round(max(target - threshold, threshold) * 0.2))
        if quantity <= threshold + warning_band:
            return "Vicino alla soglia", "warning"
    return "Disponibile", "success"


def login_required(view):
    @wraps(view)
    def wrapped(**kwargs):
        if g.user is None:
            return redirect(url_for("login", next=request.full_path))
        return view(**kwargs)

    return wrapped


def admin_required(view):
    @wraps(view)
    def wrapped(**kwargs):
        if g.user is None:
            return redirect(url_for("login"))
        if g.user["role"] != "admin":
            abort(403)
        return view(**kwargs)

    return wrapped


def first_user_missing() -> bool:
    return get_db().execute("SELECT NOT EXISTS(SELECT 1 FROM users)").fetchone()[0] == 1


def register_routes(app: Flask) -> None:
    @app.route("/setup", methods=("GET", "POST"))
    def setup():
        if not first_user_missing():
            return redirect(url_for("login"))
        if request.method == "POST":
            full_name = request.form.get("full_name", "").strip()
            username = request.form.get("username", "").strip()
            password = request.form.get("password", "")
            if not full_name or not username or len(password) < 8:
                flash("Compilare tutti i campi; la password deve avere almeno 8 caratteri.", "error")
            else:
                db = get_db()
                db.execute(
                    "INSERT INTO users (username, full_name, password_hash, role) VALUES (?, ?, ?, 'admin')",
                    (username, full_name, generate_password_hash(password)),
                )
                db.commit()
                flash("Amministratore creato. Ora puoi accedere.", "success")
                return redirect(url_for("login"))
        return render_template("setup.html")

    @app.route("/login", methods=("GET", "POST"))
    def login():
        if first_user_missing():
            return redirect(url_for("setup"))
        if request.method == "POST":
            user = get_db().execute(
                "SELECT * FROM users WHERE username = ? AND active = 1",
                (request.form.get("username", "").strip(),),
            ).fetchone()
            if user is None or not check_password_hash(user["password_hash"], request.form.get("password", "")):
                flash("Credenziali non valide.", "error")
            else:
                session.clear()
                session["user_id"] = user["id"]
                session["csrf_token"] = secrets.token_urlsafe(32)
                destination = request.args.get("next", "")
                if not destination.startswith("/") or destination.startswith("//"):
                    destination = url_for("dashboard")
                return redirect(destination)
        return render_template("login.html")

    @app.post("/logout")
    @login_required
    def logout():
        session.clear()
        return redirect(url_for("login"))

    @app.get("/")
    @login_required
    def dashboard():
        db = get_db()
        stats = db.execute(
            """
            SELECT COUNT(*) total,
                   SUM(quantity = 0) exhausted,
                   SUM(reorder_threshold > 0 AND quantity <= reorder_threshold) below_threshold,
                   SUM(reorder_threshold > 0
                       AND quantity > reorder_threshold
                       AND quantity <= reorder_threshold + MAX(
                           1,
                           ROUND(MAX(target_quantity - reorder_threshold, reorder_threshold) * 0.2)
                       )) approaching,
                   SUM(target_quantity = 0) missing_target
            FROM components
            """
        ).fetchone()
        recent = db.execute(
            """
            SELECT m.*, c.code, c.description, u.full_name AS operator_name
            FROM movements m
            JOIN components c ON c.id = m.component_id
            JOIN users u ON u.id = m.user_id
            ORDER BY m.id DESC LIMIT 10
            """
        ).fetchall()
        attention = db.execute(
            """
            SELECT * FROM components
            WHERE quantity = 0
               OR (reorder_threshold > 0 AND quantity <= reorder_threshold)
               OR (reorder_threshold > 0
                   AND quantity > reorder_threshold
                   AND quantity <= reorder_threshold + MAX(
                       1,
                       ROUND(MAX(target_quantity - reorder_threshold, reorder_threshold) * 0.2)
                   ))
            ORDER BY (quantity = 0) DESC, quantity - reorder_threshold, code
            LIMIT 12
            """
        ).fetchall()
        return render_template("dashboard.html", stats=stats, recent=recent, attention=attention)

    @app.get("/componenti")
    @login_required
    def components():
        db = get_db()
        query = request.args.get("q", "").strip()
        category = request.args.get("categoria", "").strip()
        status = request.args.get("stato", "").strip()
        clauses, params = [], []
        if query:
            clauses.append("(code LIKE ? OR description LIKE ? OR value LIKE ?)")
            needle = f"%{query}%"
            params.extend((needle, needle, needle))
        if category:
            clauses.append("category = ?")
            params.append(category)
        if status == "esaurito":
            clauses.append("quantity = 0")
        elif status == "riordino":
            clauses.append("target_quantity > quantity AND quantity <= reorder_threshold")
        elif status == "vicino":
            clauses.append(
                """reorder_threshold > 0
                AND quantity > reorder_threshold
                AND quantity <= reorder_threshold + MAX(
                    1,
                    ROUND(MAX(target_quantity - reorder_threshold, reorder_threshold) * 0.2)
                )"""
            )
        elif status == "obiettivo":
            clauses.append("target_quantity = 0")
        sql = "SELECT * FROM components"
        if clauses:
            sql += " WHERE " + " AND ".join(clauses)
        sql += " ORDER BY code LIMIT 1000"
        items = db.execute(sql, params).fetchall()
        categories = db.execute(
            "SELECT DISTINCT category FROM components WHERE category <> '' ORDER BY category"
        ).fetchall()
        return render_template("components.html", components=items, categories=categories)

    @app.get("/componenti/<int:component_id>")
    @login_required
    def component_detail(component_id: int):
        db = get_db()
        component = db.execute("SELECT * FROM components WHERE id = ?", (component_id,)).fetchone()
        if component is None:
            abort(404)
        movements = db.execute(
            """
            SELECT m.*, u.full_name AS operator_name FROM movements m
            JOIN users u ON u.id = m.user_id
            WHERE m.component_id = ? ORDER BY m.id DESC LIMIT 100
            """,
            (component_id,),
        ).fetchall()
        return render_template("component_detail.html", component=component, movements=movements)

    @app.route("/componenti/<int:component_id>/modifica", methods=("GET", "POST"))
    @admin_required
    def component_edit(component_id: int):
        db = get_db()
        component = db.execute("SELECT * FROM components WHERE id = ?", (component_id,)).fetchone()
        if component is None:
            abort(404)
        if request.method == "POST":
            try:
                threshold = non_negative_int(request.form.get("reorder_threshold"), "Soglia")
                target = non_negative_int(request.form.get("target_quantity"), "Quantità obiettivo")
                if target and threshold > target:
                    raise ValueError("La soglia non può superare la quantità obiettivo")
                description = request.form.get("description", "").strip()
                if not description:
                    raise ValueError("La descrizione è obbligatoria")
                db.execute(
                    """
                    UPDATE components SET description = ?, category = ?, value = ?, laboratory = ?,
                        cabinet = ?, door = ?, shelf = ?, drawer = ?, reorder_threshold = ?,
                        target_quantity = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?
                    """,
                    (
                        description,
                        request.form.get("category", "").strip(),
                        request.form.get("value", "").strip(),
                        request.form.get("laboratory", "").strip(),
                        request.form.get("cabinet", "").strip(),
                        request.form.get("door", "").strip(),
                        request.form.get("shelf", "").strip(),
                        request.form.get("drawer", "").strip(),
                        threshold,
                        target,
                        component_id,
                    ),
                )
                db.commit()
                flash("Dati del componente aggiornati.", "success")
                return redirect(url_for("component_detail", component_id=component_id))
            except ValueError as error:
                flash(str(error), "error")
        return render_template("component_form.html", component=component)

    @app.route("/movimenti/nuovo", methods=("GET", "POST"))
    @login_required
    def movement_new():
        db = get_db()
        selected_id = request.args.get("componente", type=int)
        if request.method == "POST":
            selected_id = request.form.get("component_id", type=int)
            movement_type = request.form.get("movement_type", "")
            recipient = request.form.get("recipient", "").strip()
            note = request.form.get("note", "").strip()
            try:
                if movement_type not in {
                    "prelievo", "restituzione", "carico", "rettifica_carico", "rettifica_scarico"
                }:
                    raise ValueError("Tipo di movimento non valido")
                quantity = positive_int(request.form.get("quantity"))
                if movement_type == "prelievo" and not recipient:
                    raise ValueError("Indicare chi preleva il materiale")
                db.execute("BEGIN IMMEDIATE")
                component = db.execute(
                    "SELECT * FROM components WHERE id = ?", (selected_id,)
                ).fetchone()
                if component is None:
                    raise ValueError("Componente non trovato")
                outgoing = movement_type in {"prelievo", "rettifica_scarico"}
                delta = -quantity if outgoing else quantity
                balance = component["quantity"] + delta
                if balance < 0:
                    raise ValueError(
                        f"Giacenza insufficiente: disponibili {component['quantity']} pezzi"
                    )
                db.execute(
                    "UPDATE components SET quantity = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?",
                    (balance, selected_id),
                )
                db.execute(
                    """
                    INSERT INTO movements (
                        component_id, movement_type, quantity, delta, balance_after,
                        recipient, note, user_id
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (selected_id, movement_type, quantity, delta, balance, recipient, note, g.user["id"]),
                )
                db.commit()
                flash(f"Movimento registrato. Nuova giacenza: {balance}.", "success")
                return redirect(url_for("component_detail", component_id=selected_id))
            except ValueError as error:
                db.rollback()
                flash(str(error), "error")
        components_list = db.execute(
            "SELECT id, code, description, quantity FROM components ORDER BY code"
        ).fetchall()
        return render_template(
            "movement_form.html", components=components_list, selected_id=selected_id
        )

    @app.get("/movimenti")
    @login_required
    def movements():
        rows = get_db().execute(
            """
            SELECT m.*, c.code, c.description, u.full_name AS operator_name
            FROM movements m JOIN components c ON c.id = m.component_id
            JOIN users u ON u.id = m.user_id
            ORDER BY m.id DESC LIMIT 1000
            """
        ).fetchall()
        return render_template("movements.html", movements=rows)

    @app.get("/riordino")
    @login_required
    def reorder_report():
        return render_template(
            "reorder.html", items=get_reorder_items(), approaching=get_approaching_items()
        )

    @app.get("/riordino.xlsx")
    @login_required
    def reorder_excel():
        items = get_reorder_items()
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Distinta riordino"
        sheet.append(
            ["Codice", "Descrizione", "Categoria", "Giacenza", "Soglia", "Obiettivo", "Da ordinare", "Ubicazione"]
        )
        for item in items:
            target = item["target_quantity"] if item["target_quantity"] > 0 else "DA DEFINIRE"
            to_order = item["to_order"] if item["to_order"] is not None else "DA DEFINIRE"
            sheet.append(
                [
                    item["code"], item["description"], item["category"], item["quantity"],
                    item["reorder_threshold"], target, to_order,
                    location_text(item),
                ]
            )
        for column in sheet.columns:
            sheet.column_dimensions[column[0].column_letter].width = min(
                55, max(len(str(cell.value or "")) for cell in column) + 2
            )
        output = io.BytesIO()
        workbook.save(output)
        output.seek(0)
        stamp = datetime.now().strftime("%Y-%m-%d")
        return send_file(
            output,
            as_attachment=True,
            download_name=f"distinta-riordino-{stamp}.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    @app.get("/riordino.pdf")
    @login_required
    def reorder_pdf():
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib.styles import getSampleStyleSheet
        from reportlab.lib.units import mm
        from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

        items = get_reorder_items()
        output = io.BytesIO()
        doc = SimpleDocTemplate(
            output, pagesize=landscape(A4), leftMargin=10 * mm, rightMargin=10 * mm,
            topMargin=10 * mm, bottomMargin=10 * mm,
        )
        styles = getSampleStyleSheet()
        story = [
            Paragraph("ITIS Cannizzaro Colleferro — Distinta materiali da ordinare", styles["Title"]),
            Paragraph(datetime.now().strftime("Generata il %d/%m/%Y alle %H:%M"), styles["Normal"]),
            Spacer(1, 5 * mm),
        ]
        data = [["Codice", "Descrizione", "Giac.", "Soglia", "Obiettivo", "Ordine"]]
        for item in items:
            target = item["target_quantity"] if item["target_quantity"] > 0 else "Da definire"
            to_order = item["to_order"] if item["to_order"] is not None else "Da definire"
            data.append(
                [item["code"], Paragraph(item["description"], styles["BodyText"]), item["quantity"],
                 item["reorder_threshold"], target, to_order]
            )
        if not items:
            data.append(["—", "Nessun componente da ordinare", "", "", "", ""])
        table = Table(data, colWidths=[28 * mm, 130 * mm, 20 * mm, 22 * mm, 24 * mm, 22 * mm], repeatRows=1)
        table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#173d6b")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#eef3f8")]),
                    ("ALIGN", (2, 1), (-1, -1), "RIGHT"),
                    ("FONTSIZE", (0, 0), (-1, -1), 8),
                ]
            )
        )
        story.append(table)
        doc.build(story)
        output.seek(0)
        stamp = datetime.now().strftime("%Y-%m-%d")
        return send_file(output, as_attachment=True, download_name=f"distinta-riordino-{stamp}.pdf", mimetype="application/pdf")

    @app.route("/utenti", methods=("GET", "POST"))
    @admin_required
    def users():
        db = get_db()
        if request.method == "POST":
            username = request.form.get("username", "").strip()
            full_name = request.form.get("full_name", "").strip()
            password = request.form.get("password", "")
            role = request.form.get("role", "operatore")
            try:
                if not username or not full_name or len(password) < 8:
                    raise ValueError("Compilare tutti i campi; password minima di 8 caratteri")
                if role not in {"admin", "operatore"}:
                    raise ValueError("Ruolo non valido")
                db.execute(
                    "INSERT INTO users (username, full_name, password_hash, role) VALUES (?, ?, ?, ?)",
                    (username, full_name, generate_password_hash(password), role),
                )
                db.commit()
                flash("Utente creato.", "success")
                return redirect(url_for("users"))
            except sqlite3.IntegrityError:
                flash("Nome utente già utilizzato.", "error")
            except ValueError as error:
                flash(str(error), "error")
        rows = db.execute("SELECT * FROM users ORDER BY active DESC, full_name").fetchall()
        return render_template("users.html", users=rows)

    @app.post("/utenti/<int:user_id>/stato")
    @admin_required
    def user_toggle(user_id: int):
        if user_id == g.user["id"]:
            flash("Non puoi disattivare il tuo account.", "error")
        else:
            db = get_db()
            db.execute("UPDATE users SET active = NOT active WHERE id = ?", (user_id,))
            db.commit()
            flash("Stato utente aggiornato.", "success")
        return redirect(url_for("users"))


def get_reorder_items():
    return get_db().execute(
        """
        SELECT *,
               CASE
                   WHEN target_quantity > quantity THEN target_quantity - quantity
                   ELSE NULL
               END AS to_order
        FROM components
        WHERE quantity = 0
           OR (target_quantity > quantity AND quantity <= reorder_threshold)
        ORDER BY (target_quantity = 0) DESC, category, code
        """
    ).fetchall()


def get_approaching_items():
    return get_db().execute(
        """
        SELECT *, quantity - reorder_threshold AS units_to_threshold
        FROM components
        WHERE reorder_threshold > 0
          AND quantity > reorder_threshold
          AND quantity <= reorder_threshold + MAX(
              1,
              ROUND(MAX(target_quantity - reorder_threshold, reorder_threshold) * 0.2)
          )
        ORDER BY units_to_threshold, category, code
        """
    ).fetchall()


def location_text(component) -> str:
    parts = [
        component["laboratory"],
        f"Arm. {component['cabinet']}" if component["cabinet"] else "",
        f"Anta {component['door']}" if component["door"] else "",
        f"Scaff. {component['shelf']}" if component["shelf"] else "",
        f"Cass. {component['drawer']}" if component["drawer"] else "",
    ]
    return " · ".join(part for part in parts if part)


app = create_app()


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "3000")), debug=False)
