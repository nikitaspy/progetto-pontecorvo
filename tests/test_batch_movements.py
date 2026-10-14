import tempfile
import unittest
from pathlib import Path

from werkzeug.datastructures import MultiDict
from werkzeug.security import generate_password_hash

from app import create_app, get_db


class BatchMovementTest(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.app = create_app(
            {
                "TESTING": True,
                "DATABASE": str(Path(self.temp.name) / "batch-test.db"),
                "EXCEL_SOURCE": str(Path(self.temp.name) / "missing.xlsx"),
                "CSRF_ENABLED": False,
                "SECRET_KEY": "batch-test",
            }
        )
        with self.app.app_context():
            db = get_db()
            self.user_id = db.execute(
                """INSERT INTO users (username, full_name, password_hash, role)
                   VALUES (?, ?, ?, 'operatore')""",
                (
                    "docente",
                    "Prof.ssa Ada Bianchi",
                    generate_password_hash("password-test"),
                ),
            ).lastrowid
            db.executemany(
                """INSERT INTO components
                   (code, description, quantity, reorder_threshold, target_quantity)
                   VALUES (?, ?, ?, 0, ?)""",
                [
                    ("TPS0001", "Resistore", 10, 10),
                    ("TPS0002", "Condensatore", 4, 4),
                    ("TPS0003", "LED", 0, 0),
                ],
            )
            db.commit()
        self.client = self.app.test_client()
        with self.client.session_transaction() as session_data:
            session_data["user_id"] = self.user_id
            session_data["auth_version"] = 1

    def tearDown(self):
        self.temp.cleanup()

    def payload(self, operation_id, component_ids, quantities, **overrides):
        data = MultiDict(
            [
                ("operation_id", operation_id),
                ("movement_type", "prelievo"),
                ("destination", "Laboratorio TPS"),
                ("note", "Classe 4A"),
            ]
        )
        for component_id in component_ids:
            data.add("component_id", str(component_id))
        for quantity in quantities:
            data.add("quantity", str(quantity))
        for key, value in overrides.items():
            data.setlist(key, value if isinstance(value, list) else [value])
        return data

    def test_operator_withdraws_repeated_and_different_components_in_one_batch(self):
        page = self.client.get("/movimenti/multipli")
        self.assertEqual(page.status_code, 200)
        self.assertIn("Prof.ssa Ada Bianchi".encode(), page.data)
        self.assertIn("Aggiungi componente".encode(), page.data)

        response = self.client.post(
            "/movimenti/multipli",
            data=self.payload(
                "batch-multiple-components-0001",
                [1, 1, 2],
                [2, 3, 4],
            ),
        )
        self.assertEqual(response.status_code, 302)

        with self.app.app_context():
            db = get_db()
            quantities = {
                row["code"]: row["quantity"]
                for row in db.execute("SELECT code, quantity FROM components")
            }
            self.assertEqual(quantities["TPS0001"], 5)
            self.assertEqual(quantities["TPS0002"], 0)
            movements = db.execute(
                "SELECT * FROM movements ORDER BY id"
            ).fetchall()
            self.assertEqual(len(movements), 3)
            self.assertEqual([row["balance_after"] for row in movements], [8, 5, 0])
            self.assertEqual({row["batch_id"] for row in movements}, {1})
            self.assertEqual({row["user_id"] for row in movements}, {self.user_id})
            self.assertEqual(
                {row["recipient"] for row in movements}, {"Prof.ssa Ada Bianchi"}
            )
            self.assertEqual(
                {row["note"] for row in movements},
                {"Destinazione: Laboratorio TPS — Classe 4A"},
            )

    def test_batch_return_adds_multiple_quantities(self):
        response = self.client.post(
            "/movimenti/multipli",
            data=self.payload(
                "batch-return-components-0001",
                [1, 3],
                [2, 5],
                movement_type="restituzione",
            ),
        )
        self.assertEqual(response.status_code, 302)
        with self.app.app_context():
            db = get_db()
            quantities = {
                row["code"]: row["quantity"]
                for row in db.execute("SELECT code, quantity FROM components")
            }
            self.assertEqual(quantities["TPS0001"], 12)
            self.assertEqual(quantities["TPS0003"], 5)
            movements = db.execute("SELECT * FROM movements ORDER BY id").fetchall()
            self.assertEqual([row["delta"] for row in movements], [2, 5])
            self.assertEqual({row["movement_type"] for row in movements}, {"restituzione"})

    def test_insufficient_stock_rolls_back_every_line(self):
        response = self.client.post(
            "/movimenti/multipli",
            data=self.payload(
                "batch-atomic-insufficient-0001",
                [1, 2],
                [3, 5],
            ),
        )
        self.assertEqual(response.status_code, 200)
        self.assertIn("Giacenza insufficiente per TPS0002".encode(), response.data)
        with self.app.app_context():
            db = get_db()
            self.assertEqual(
                [row[0] for row in db.execute("SELECT quantity FROM components ORDER BY id")],
                [10, 4, 0],
            )
            self.assertEqual(db.execute("SELECT COUNT(*) FROM movements").fetchone()[0], 0)
            self.assertEqual(
                db.execute("SELECT COUNT(*) FROM movement_batches").fetchone()[0], 0
            )

    def test_invalid_row_writes_nothing(self):
        response = self.client.post(
            "/movimenti/multipli",
            data=self.payload(
                "batch-invalid-line-number-0001",
                [1, ""],
                [2, 1],
            ),
        )
        self.assertEqual(response.status_code, 200)
        self.assertIn("Riga 2".encode(), response.data)
        with self.app.app_context():
            db = get_db()
            self.assertEqual(db.execute("SELECT SUM(quantity) FROM components").fetchone()[0], 14)
            self.assertEqual(db.execute("SELECT COUNT(*) FROM movements").fetchone()[0], 0)

    def test_repeated_submit_is_idempotent_for_whole_batch(self):
        payload = self.payload(
            "batch-idempotent-operation-0001",
            [1, 2],
            [3, 2],
        )
        first = self.client.post("/movimenti/multipli", data=payload)
        second = self.client.post("/movimenti/multipli", data=payload)
        self.assertEqual(first.status_code, 302)
        self.assertEqual(second.status_code, 302)

        with self.app.app_context():
            db = get_db()
            self.assertEqual(
                [row[0] for row in db.execute("SELECT quantity FROM components ORDER BY id")],
                [7, 2, 0],
            )
            self.assertEqual(db.execute("SELECT COUNT(*) FROM movements").fetchone()[0], 2)
            self.assertEqual(
                db.execute("SELECT COUNT(*) FROM movement_batches").fetchone()[0], 1
            )

    def test_login_is_required_before_batch_can_change_stock(self):
        with self.client.session_transaction() as session_data:
            session_data.clear()
        response = self.client.post(
            "/movimenti/multipli",
            data=self.payload("batch-anonymous-attempt-0001", [1], [2]),
        )
        self.assertEqual(response.status_code, 302)
        self.assertIn("/login", response.headers["Location"])
        with self.app.app_context():
            db = get_db()
            self.assertEqual(db.execute("SELECT quantity FROM components WHERE id = 1").fetchone()[0], 10)
            self.assertEqual(db.execute("SELECT COUNT(*) FROM movements").fetchone()[0], 0)


if __name__ == "__main__":
    unittest.main()
