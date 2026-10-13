import sqlite3
import tempfile
import unittest
from io import BytesIO
from pathlib import Path

from openpyxl import load_workbook
from werkzeug.security import generate_password_hash

from app import create_app, get_db, migrate_reorder_lines


class WarehouseTest(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.db_path = str(Path(self.temp.name) / "test.db")
        self.app = create_app(
            {
                "TESTING": True,
                "DATABASE": self.db_path,
                "EXCEL_SOURCE": str(Path(self.temp.name) / "missing.xlsx"),
                "CSRF_ENABLED": False,
                "SECRET_KEY": "test",
            }
        )
        with self.app.app_context():
            db = get_db()
            db.execute(
                "INSERT INTO users (username, full_name, password_hash, role) VALUES (?, ?, ?, ?)",
                ("admin", "Admin Test", generate_password_hash("password-test"), "admin"),
            )
            db.execute(
                """
                INSERT INTO components
                    (code, description, quantity, reorder_threshold, target_quantity)
                VALUES ('TPS0001', 'Resistore test', 10, 3, 10)
                """
            )
            db.commit()
        self.client = self.app.test_client()
        self.client.post("/login", data={"username": "admin", "password": "password-test"})

    def tearDown(self):
        self.temp.cleanup()

    def test_withdrawal_updates_stock_and_audit(self):
        response = self.client.post(
            "/movimenti/nuovo",
            data={
                "component_id": "1",
                "movement_type": "prelievo",
                "quantity": "4",
                "recipient": "Prof. Rossi",
                "note": "Classe 4A",
            },
        )
        self.assertEqual(response.status_code, 302)
        with self.app.app_context():
            db = get_db()
            self.assertEqual(db.execute("SELECT quantity FROM components").fetchone()[0], 6)
            movement = db.execute("SELECT * FROM movements").fetchone()
            self.assertEqual(movement["delta"], -4)
            self.assertEqual(movement["recipient"], "Prof. Rossi")

    def test_stock_cannot_become_negative(self):
        response = self.client.post(
            "/movimenti/nuovo",
            data={
                "component_id": "1",
                "movement_type": "prelievo",
                "quantity": "11",
                "recipient": "Prof. Rossi",
            },
            follow_redirects=True,
        )
        self.assertIn("Giacenza insufficiente".encode(), response.data)
        with self.app.app_context():
            self.assertEqual(get_db().execute("SELECT quantity FROM components").fetchone()[0], 10)

    def test_reorder_report_uses_target_minus_current_stock(self):
        self.client.post(
            "/movimenti/nuovo",
            data={
                "component_id": "1",
                "movement_type": "prelievo",
                "quantity": "7",
                "recipient": "Lab TPS",
            },
        )
        response = self.client.get("/riordino")
        self.assertIn(b"TPS0001", response.data)
        self.assertIn(b'name="order_quantity_1"', response.data)
        self.assertIn(b'value="7"', response.data)

    def test_report_warns_before_reaching_threshold(self):
        self.client.post(
            "/movimenti/nuovo",
            data={
                "component_id": "1",
                "movement_type": "prelievo",
                "quantity": "6",
                "recipient": "Lab TPS",
            },
        )
        response = self.client.get("/riordino")
        self.assertIn("articoli vicini alla soglia".encode(), response.data)
        self.assertIn(b"TPS0001", response.data)

    def test_zero_stock_without_target_appears_in_reorder_report(self):
        with self.app.app_context():
            db = get_db()
            db.execute(
                """
                INSERT INTO components
                    (code, description, quantity, reorder_threshold, target_quantity)
                VALUES ('TPS0002', 'Componente esaurito', 0, 0, 0)
                """
            )
            db.execute(
                """
                INSERT INTO components
                    (code, description, quantity, reorder_threshold, target_quantity)
                VALUES ('TPS0003', 'Obiettivo ancora da definire', 2, 3, 0)
                """
            )
            db.commit()
        response = self.client.get("/riordino")
        self.assertIn(b"TPS0002", response.data)
        self.assertIn(b"TPS0003", response.data)
        self.assertIn(b"Da definire", response.data)

        excel = self.client.get("/riordino.xlsx")
        self.assertEqual(excel.status_code, 200)
        workbook = load_workbook(BytesIO(excel.data), data_only=True)
        rows = list(workbook["Distinta da completare"].iter_rows(values_only=True))
        zero_item = next(row for row in rows if row[0] == "TPS0002")
        self.assertEqual(zero_item[5], "DA DEFINIRE")
        self.assertEqual(zero_item[6], "DA DEFINIRE")

    def test_bulk_order_quantity_applies_equally_to_selected_components(self):
        with self.app.app_context():
            db = get_db()
            db.execute(
                """INSERT INTO components
                   (code, description, quantity, reorder_threshold, target_quantity)
                   VALUES ('TPS0002', 'Esaurito', 0, 0, 0)"""
            )
            db.execute(
                """INSERT INTO components
                   (code, description, quantity, reorder_threshold, target_quantity)
                   VALUES ('TPS0003', 'Sotto soglia', 2, 3, 5)"""
            )
            db.commit()

        response = self.client.post(
            "/riordino/quantita",
            data={
                "action": "bulk",
                "bulk_order_quantity": "4",
                "selected_ids": ["2", "3"],
            },
            follow_redirects=True,
        )
        self.assertIn("Aggiunti al modulo 2 componenti".encode(), response.data)
        with self.app.app_context():
            db = get_db()
            rows = db.execute(
                "SELECT code, quantity, target_quantity FROM components ORDER BY code"
            ).fetchall()
            self.assertEqual(rows[0]["target_quantity"], 10)
            self.assertEqual(rows[1]["target_quantity"], 0)
            self.assertEqual(rows[2]["target_quantity"], 5)
            order_lines = db.execute(
                """
                SELECT c.code, r.desired_quantity
                FROM reorder_lines r
                JOIN components c ON c.id = r.component_id
                ORDER BY c.code
                """
            ).fetchall()
            self.assertEqual(
                [(row["code"], row["desired_quantity"]) for row in order_lines],
                [("TPS0002", 4), ("TPS0003", 4)],
            )

        pending = self.client.get("/riordino")
        self.assertNotIn(b"TPS0002", pending.data)
        self.assertNotIn(b"TPS0003", pending.data)

        order_form = self.client.get("/modulo-riordino")
        self.assertIn(b"TPS0002", order_form.data)
        self.assertIn(b"TPS0003", order_form.data)
        self.assertIn(b"Pezzi totali", order_form.data)

        excel = self.client.get("/modulo-riordino.xlsx")
        workbook = load_workbook(BytesIO(excel.data), data_only=True)
        exported = {
            row[0]: row[4]
            for row in workbook["Modulo riordino"].iter_rows(min_row=2, values_only=True)
        }
        self.assertEqual(exported["TPS0002"], 4)
        self.assertEqual(exported["TPS0003"], 4)

        pending_excel = self.client.get("/riordino.xlsx")
        pending_workbook = load_workbook(BytesIO(pending_excel.data), data_only=True)
        pending_codes = {
            row[0]
            for row in pending_workbook["Distinta da completare"].iter_rows(
                min_row=2, values_only=True
            )
        }
        self.assertNotIn("TPS0002", pending_codes)
        self.assertNotIn("TPS0003", pending_codes)

    def test_individual_order_quantity_can_differ_from_threshold_and_return_to_pending(self):
        with self.app.app_context():
            db = get_db()
            db.execute(
                """INSERT INTO components
                   (code, description, quantity, reorder_threshold, target_quantity)
                   VALUES ('TPS0002', 'Esaurito', 0, 6, 0)"""
            )
            db.commit()

        added = self.client.post(
            "/riordino/quantita",
            data={"action": "single:2", "order_quantity_2": "5"},
            follow_redirects=True,
        )
        self.assertIn("TPS0002 aggiunto al modulo con quantit".encode(), added.data)
        with self.app.app_context():
            row = get_db().execute(
                """
                SELECT c.target_quantity, r.desired_quantity
                FROM components c
                JOIN reorder_lines r ON r.component_id = c.id
                WHERE c.code = 'TPS0002'
                """
            ).fetchone()
            self.assertEqual(row["target_quantity"], 0)
            self.assertEqual(row["desired_quantity"], 5)

        updated = self.client.post(
            "/modulo-riordino/2/quantita",
            data={"desired_quantity": "8"},
            follow_redirects=True,
        )
        self.assertIn("modulo aggiornata".encode(), updated.data)
        with self.app.app_context():
            self.assertEqual(
                get_db().execute(
                    "SELECT desired_quantity FROM reorder_lines WHERE component_id = 2"
                ).fetchone()[0],
                8,
            )

        removed = self.client.post(
            "/modulo-riordino/2/rimuovi", follow_redirects=True
        )
        self.assertIn("riportato nella distinta".encode(), removed.data)
        with self.app.app_context():
            db = get_db()
            self.assertIsNone(
                db.execute(
                    "SELECT 1 FROM reorder_lines WHERE component_id = 2"
                ).fetchone()
            )
            self.assertEqual(
                db.execute(
                    "SELECT target_quantity FROM components WHERE id = 2"
                ).fetchone()[0],
                0,
            )
        self.assertIn(b"TPS0002", self.client.get("/riordino").data)

    def test_operator_can_view_but_cannot_modify_order_form(self):
        with self.app.app_context():
            db = get_db()
            cursor = db.execute(
                "INSERT INTO users (username, full_name, password_hash, role) VALUES (?, ?, ?, ?)",
                ("operatore", "Operatore Test", generate_password_hash("password-test"), "operatore"),
            )
            operator_id = cursor.lastrowid
            component = db.execute(
                """INSERT INTO components
                   (code, description, quantity, reorder_threshold, target_quantity)
                   VALUES ('TPS0002', 'Esaurito', 0, 0, 0)"""
            )
            component_id = component.lastrowid
            db.execute(
                "INSERT INTO reorder_lines (component_id, desired_quantity) VALUES (?, 6)",
                (component_id,),
            )
            db.commit()
        with self.client.session_transaction() as session_data:
            session_data["user_id"] = operator_id

        page = self.client.get("/modulo-riordino")
        self.assertEqual(page.status_code, 200)
        self.assertIn(b"TPS0002", page.data)
        self.assertNotIn(b'name="desired_quantity"', page.data)
        self.assertNotIn("Rimetti in distinta".encode(), page.data)
        self.assertEqual(self.client.get("/modulo-riordino.xlsx").status_code, 200)
        self.assertEqual(self.client.get("/modulo-riordino.pdf").status_code, 200)

        add_response = self.client.post(
            "/riordino/quantita",
            data={"action": "single:1", "order_quantity_1": "5"},
        )
        update_response = self.client.post(
            f"/modulo-riordino/{component_id}/quantita",
            data={"desired_quantity": "9"},
        )
        remove_response = self.client.post(
            f"/modulo-riordino/{component_id}/rimuovi"
        )
        self.assertEqual(add_response.status_code, 403)
        self.assertEqual(update_response.status_code, 403)
        self.assertEqual(remove_response.status_code, 403)
        with self.app.app_context():
            self.assertEqual(
                get_db().execute(
                    "SELECT desired_quantity FROM reorder_lines WHERE component_id = ?",
                    (component_id,),
                ).fetchone()[0],
                6,
            )

    def test_legacy_target_migration_is_idempotent_and_preserves_existing_lines(self):
        with self.app.app_context():
            db = get_db()
            db.execute(
                "DELETE FROM app_settings WHERE key = 'reorder_lines_migrated_from_targets_v1'"
            )
            migrated_component = db.execute(
                """INSERT INTO components
                   (code, description, quantity, reorder_threshold, target_quantity)
                   VALUES ('TPS0002', 'Legacy da migrare', 0, 0, 12)"""
            ).lastrowid
            existing_component = db.execute(
                """INSERT INTO components
                   (code, description, quantity, reorder_threshold, target_quantity)
                   VALUES ('TPS0003', 'Riga esistente', 0, 0, 9)"""
            ).lastrowid
            positive_component = db.execute(
                """INSERT INTO components
                   (code, description, quantity, reorder_threshold, target_quantity)
                   VALUES ('TPS0004', 'Legacy sotto soglia', 2, 3, 10)"""
            ).lastrowid
            natural_depletion = db.execute(
                """INSERT INTO components
                   (code, description, quantity, reorder_threshold, target_quantity,
                    created_at, updated_at)
                   VALUES ('TPS0005', 'Esaurimento naturale', 0, 0, 10,
                           '2026-01-01 08:00:00', '2026-01-02 08:00:00')"""
            ).lastrowid
            confirmed_after_movement = db.execute(
                """INSERT INTO components
                   (code, description, quantity, reorder_threshold, target_quantity,
                    created_at, updated_at)
                   VALUES ('TPS0006', 'Conferma dopo prelievo', 2, 3, 6,
                           '2026-01-01 08:00:00', '2026-01-03 08:00:00')"""
            ).lastrowid
            db.executemany(
                """INSERT INTO movements
                   (component_id, movement_type, quantity, delta, balance_after,
                    recipient, user_id, created_at)
                   VALUES (?, 'prelievo', ?, ?, ?, 'Laboratorio', 1, ?)""",
                [
                    (natural_depletion, 10, -10, 0, "2026-01-02 08:00:00"),
                    (confirmed_after_movement, 8, -8, 2, "2026-01-02 08:00:00"),
                ],
            )
            db.execute(
                "INSERT INTO reorder_lines (component_id, desired_quantity) VALUES (?, 4)",
                (existing_component,),
            )

            migrate_reorder_lines(db)
            rows = db.execute(
                """
                SELECT c.code, r.desired_quantity
                FROM reorder_lines r
                JOIN components c ON c.id = r.component_id
                ORDER BY c.code
                """
            ).fetchall()
            self.assertEqual(
                [(row["code"], row["desired_quantity"]) for row in rows],
                [
                    ("TPS0002", 12),
                    ("TPS0003", 4),
                    ("TPS0004", 8),
                    ("TPS0006", 4),
                ],
            )
            self.assertIsNone(
                db.execute(
                    "SELECT 1 FROM reorder_lines WHERE component_id = ?",
                    (natural_depletion,),
                ).fetchone()
            )
            self.assertEqual(
                db.execute(
                    "SELECT desired_quantity FROM reorder_lines WHERE component_id = ?",
                    (positive_component,),
                ).fetchone()[0],
                8,
            )

            db.execute(
                "UPDATE components SET target_quantity = 20 WHERE id = ?",
                (migrated_component,),
            )
            migrate_reorder_lines(db)
            rows_after_second_run = db.execute(
                "SELECT component_id, desired_quantity FROM reorder_lines ORDER BY component_id"
            ).fetchall()
            self.assertEqual(len(rows_after_second_run), 4)
            self.assertEqual(
                db.execute(
                    "SELECT desired_quantity FROM reorder_lines WHERE component_id = ?",
                    (migrated_component,),
                ).fetchone()[0],
                12,
            )
            self.assertIsNotNone(
                db.execute(
                    """SELECT 1 FROM app_settings
                       WHERE key = 'reorder_lines_migrated_from_targets_v1'"""
                ).fetchone()
            )

    def test_movement_history_is_immutable_in_database(self):
        self.client.post(
            "/movimenti/nuovo",
            data={
                "component_id": "1",
                "movement_type": "carico",
                "quantity": "1",
                "recipient": "",
            },
        )
        with self.app.app_context():
            with self.assertRaises(sqlite3.IntegrityError):
                get_db().execute("DELETE FROM movements")

    def test_exports_are_downloadable(self):
        self.client.post(
            "/movimenti/nuovo",
            data={
                "component_id": "1",
                "movement_type": "prelievo",
                "quantity": "7",
                "recipient": "Lab TPS",
            },
        )
        excel = self.client.get("/riordino.xlsx")
        pdf = self.client.get("/riordino.pdf")
        self.assertEqual(excel.status_code, 200)
        self.assertTrue(excel.data.startswith(b"PK"))
        self.assertEqual(pdf.status_code, 200)
        self.assertTrue(pdf.data.startswith(b"%PDF"))


if __name__ == "__main__":
    unittest.main()
