import sqlite3
import tempfile
import unittest
from io import BytesIO
from pathlib import Path

from openpyxl import load_workbook
from werkzeug.security import generate_password_hash

from app import create_app, get_db


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
        rows = list(workbook["Distinta riordino"].iter_rows(values_only=True))
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
        self.assertIn("per 2 componenti".encode(), response.data)
        with self.app.app_context():
            rows = get_db().execute(
                "SELECT code, quantity, target_quantity FROM components ORDER BY code"
            ).fetchall()
            self.assertEqual(rows[0]["target_quantity"], 10)
            self.assertEqual(rows[1]["target_quantity"] - rows[1]["quantity"], 4)
            self.assertEqual(rows[2]["target_quantity"] - rows[2]["quantity"], 4)

        excel = self.client.get("/riordino.xlsx")
        workbook = load_workbook(BytesIO(excel.data), data_only=True)
        exported = {
            row[0]: row[6]
            for row in workbook["Distinta riordino"].iter_rows(min_row=2, values_only=True)
        }
        self.assertEqual(exported["TPS0002"], 4)
        self.assertEqual(exported["TPS0003"], 4)

    def test_individual_order_quantity_and_threshold_validation(self):
        with self.app.app_context():
            db = get_db()
            db.execute(
                """INSERT INTO components
                   (code, description, quantity, reorder_threshold, target_quantity)
                   VALUES ('TPS0002', 'Esaurito', 0, 6, 0)"""
            )
            db.commit()

        rejected = self.client.post(
            "/riordino/quantita",
            data={"action": "single:2", "order_quantity_2": "5"},
            follow_redirects=True,
        )
        self.assertIn("ordinare almeno 6 pezzi".encode(), rejected.data)

        accepted = self.client.post(
            "/riordino/quantita",
            data={"action": "single:2", "order_quantity_2": "8"},
            follow_redirects=True,
        )
        self.assertIn("TPS0002 impostata a 8".encode(), accepted.data)
        with self.app.app_context():
            self.assertEqual(
                get_db().execute(
                    "SELECT target_quantity FROM components WHERE code = 'TPS0002'"
                ).fetchone()[0],
                8,
            )

    def test_operator_cannot_change_order_quantities(self):
        with self.app.app_context():
            db = get_db()
            cursor = db.execute(
                "INSERT INTO users (username, full_name, password_hash, role) VALUES (?, ?, ?, ?)",
                ("operatore", "Operatore Test", generate_password_hash("password-test"), "operatore"),
            )
            operator_id = cursor.lastrowid
            db.commit()
        with self.client.session_transaction() as session_data:
            session_data["user_id"] = operator_id
        response = self.client.post(
            "/riordino/quantita",
            data={"action": "single:1", "order_quantity_1": "5"},
        )
        self.assertEqual(response.status_code, 403)

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
