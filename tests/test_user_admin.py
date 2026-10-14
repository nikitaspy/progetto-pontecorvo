import tempfile
import unittest
from pathlib import Path

from werkzeug.security import check_password_hash, generate_password_hash

from app import create_app, get_db


class UserAdministrationTest(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.app = create_app(
            {
                "TESTING": True,
                "DATABASE": str(Path(self.temp.name) / "test.db"),
                "EXCEL_SOURCE": str(Path(self.temp.name) / "missing.xlsx"),
                "CSRF_ENABLED": False,
                "SECRET_KEY": "test",
            }
        )
        with self.app.app_context():
            db = get_db()
            admin = db.execute(
                """INSERT INTO users (username, full_name, password_hash, role)
                   VALUES (?, ?, ?, 'admin')""",
                ("admin", "Admin Test", generate_password_hash("password-test")),
            )
            component = db.execute(
                """INSERT INTO components
                       (code, description, quantity, reorder_threshold, target_quantity)
                   VALUES ('TPS0001', 'Resistore test', 10, 3, 10)"""
            )
            db.commit()
            self.admin_id = admin.lastrowid
            self.component_id = component.lastrowid
        self.client = self.app.test_client()
        self.client.post(
            "/login", data={"username": "admin", "password": "password-test"}
        )

    def tearDown(self):
        self.temp.cleanup()

    def create_user(self, username="docente", role="operatore") -> int:
        with self.app.app_context():
            db = get_db()
            cursor = db.execute(
                """INSERT INTO users (username, full_name, password_hash, role)
                   VALUES (?, ?, ?, ?)""",
                (
                    username,
                    "Prof.ssa Maria Rossi",
                    generate_password_hash("password-test"),
                    role,
                ),
            )
            db.commit()
            return cursor.lastrowid

    def use_account(self, user_id: int) -> None:
        with self.client.session_transaction() as session_data:
            session_data.clear()
            session_data["user_id"] = user_id
            session_data["auth_version"] = 1

    def test_admin_can_edit_identity_role_password_and_status(self):
        user_id = self.create_user()

        response = self.client.post(
            f"/utenti/{user_id}/modifica",
            data={
                "full_name": "Prof. Mario Verdi",
                "username": "mverdi",
                "password": "nuova-password",
                "role": "admin",
                # An omitted checkbox represents a disabled account.
            },
        )

        self.assertEqual(response.status_code, 302)
        with self.app.app_context():
            user = get_db().execute(
                "SELECT * FROM users WHERE id = ?", (user_id,)
            ).fetchone()
            self.assertEqual(user["full_name"], "Prof. Mario Verdi")
            self.assertEqual(user["username"], "mverdi")
            self.assertEqual(user["role"], "admin")
            self.assertEqual(user["active"], 0)
            self.assertTrue(check_password_hash(user["password_hash"], "nuova-password"))

    def test_current_admin_cannot_be_demoted_deactivated_or_deleted(self):
        edit = self.client.post(
            f"/utenti/{self.admin_id}/modifica",
            data={
                "full_name": "Admin Test",
                "username": "admin",
                "role": "operatore",
            },
            follow_redirects=True,
        )
        self.assertEqual(edit.status_code, 200)
        self.assertIn("Non puoi rimuovere i privilegi".encode(), edit.data)

        delete = self.client.post(f"/utenti/{self.admin_id}/elimina")
        self.assertEqual(delete.status_code, 302)
        with self.app.app_context():
            user = get_db().execute(
                "SELECT role, active FROM users WHERE id = ?", (self.admin_id,)
            ).fetchone()
            self.assertEqual(user["role"], "admin")
            self.assertEqual(user["active"], 1)

    def test_delete_removes_user_without_movements(self):
        user_id = self.create_user()
        with self.app.app_context():
            db = get_db()
            db.execute(
                """INSERT INTO reorder_lines (component_id, desired_quantity, set_by)
                   VALUES (?, 5, ?)""",
                (self.component_id, user_id),
            )
            db.commit()

        response = self.client.post(f"/utenti/{user_id}/elimina")

        self.assertEqual(response.status_code, 302)
        with self.app.app_context():
            db = get_db()
            self.assertIsNone(
                db.execute("SELECT id FROM users WHERE id = ?", (user_id,)).fetchone()
            )
            self.assertIsNone(
                db.execute(
                    "SELECT set_by FROM reorder_lines WHERE component_id = ?",
                    (self.component_id,),
                ).fetchone()[0]
            )

    def test_delete_deactivates_user_with_movements_and_preserves_audit(self):
        user_id = self.create_user()
        with self.app.app_context():
            db = get_db()
            db.execute(
                "UPDATE components SET quantity = 9 WHERE id = ?", (self.component_id,)
            )
            movement = db.execute(
                """INSERT INTO movements
                       (component_id, movement_type, quantity, delta, balance_after,
                        recipient, user_id)
                   VALUES (?, 'prelievo', 1, -1, 9, 'Prof.ssa Maria Rossi', ?)""",
                (self.component_id, user_id),
            )
            db.commit()
            movement_id = movement.lastrowid

        response = self.client.post(
            f"/utenti/{user_id}/elimina", follow_redirects=True
        )

        self.assertEqual(response.status_code, 200)
        self.assertIn("lo storico \u00e8 stato conservato".encode(), response.data)
        with self.app.app_context():
            db = get_db()
            user = db.execute(
                "SELECT active FROM users WHERE id = ?", (user_id,)
            ).fetchone()
            self.assertEqual(user["active"], 0)
            self.assertIsNotNone(
                db.execute(
                    "SELECT id FROM movements WHERE id = ? AND user_id = ?",
                    (movement_id, user_id),
                ).fetchone()
            )

    def test_operator_cannot_manage_users_or_privileges(self):
        user_id = self.create_user()
        self.use_account(user_id)

        requests = (
            self.client.get("/utenti"),
            self.client.post(
                "/utenti",
                data={
                    "full_name": "Altro Utente",
                    "username": "altro",
                    "password": "password-test",
                    "role": "admin",
                },
            ),
            self.client.get(f"/utenti/{user_id}/modifica"),
            self.client.post(f"/utenti/{user_id}/stato"),
            self.client.post(f"/utenti/{user_id}/elimina"),
        )
        self.assertTrue(all(response.status_code == 403 for response in requests))

        with self.app.app_context():
            db = get_db()
            self.assertEqual(
                db.execute("SELECT COUNT(*) FROM users WHERE username = 'altro'").fetchone()[0],
                0,
            )
            user = db.execute(
                "SELECT role, active FROM users WHERE id = ?", (user_id,)
            ).fetchone()
            self.assertEqual(user["role"], "operatore")
            self.assertEqual(user["active"], 1)


if __name__ == "__main__":
    unittest.main()
