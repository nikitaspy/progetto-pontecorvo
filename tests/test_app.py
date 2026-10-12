import sqlite3
import tempfile
import unittest
from pathlib import Path

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
        self.assertIn(b">7<", response.data)

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
