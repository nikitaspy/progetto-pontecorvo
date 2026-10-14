import html
import re
import tempfile
import unittest
from pathlib import Path
from urllib.parse import urlsplit

from werkzeug.security import generate_password_hash

from app import create_app, get_db


HIDDEN_INPUT_PATTERN = re.compile(
    rb'<input\b[^>]*\bname=["\'](?P<name>[^"\']+)["\'][^>]*>',
    re.IGNORECASE,
)
VALUE_PATTERN = re.compile(rb'\bvalue=["\'](?P<value>[^"\']*)["\']', re.IGNORECASE)


def hidden_input(response, name: str) -> str:
    """Estrae un campo hidden senza dipendere dall'ordine degli attributi HTML."""
    expected = name.encode("utf-8")
    for match in HIDDEN_INPUT_PATTERN.finditer(response.data):
        if match.group("name") != expected:
            continue
        value = VALUE_PATTERN.search(match.group(0))
        if value is None:
            raise AssertionError(f"Il campo {name!r} non contiene un valore")
        return html.unescape(value.group("value").decode("utf-8"))
    raise AssertionError(f"Campo {name!r} non trovato nella risposta")


class SecurityRegressionTest(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.app = create_app(
            {
                "TESTING": True,
                "DATABASE": str(Path(self.temp.name) / "security-test.db"),
                "EXCEL_SOURCE": str(Path(self.temp.name) / "missing.xlsx"),
                "CSRF_ENABLED": True,
                "SECRET_KEY": "security-regression-test",
            }
        )
        with self.app.app_context():
            db = get_db()
            self.admin_id = db.execute(
                """INSERT INTO users (username, full_name, password_hash, role)
                   VALUES (?, ?, ?, 'admin')""",
                (
                    "admin",
                    "Amministratore Test",
                    generate_password_hash("password-test"),
                ),
            ).lastrowid
            self.alice_id = db.execute(
                """INSERT INTO users (username, full_name, password_hash, role)
                   VALUES (?, ?, ?, 'operatore')""",
                (
                    "alice",
                    "Prof.ssa Alice Originale",
                    generate_password_hash("password-test"),
                ),
            ).lastrowid
            self.bob_id = db.execute(
                """INSERT INTO users (username, full_name, password_hash, role)
                   VALUES (?, ?, ?, 'operatore')""",
                (
                    "bob",
                    "Prof. Bob Verdi",
                    generate_password_hash("password-test"),
                ),
            ).lastrowid
            db.executemany(
                """INSERT INTO components
                       (code, description, quantity, reorder_threshold, target_quantity)
                   VALUES (?, ?, ?, 0, ?)""",
                [
                    ("TPS0001", "Resistore di sicurezza", 10, 10),
                    ("TPS0002", "Condensatore di sicurezza", 7, 7),
                ],
            )
            db.commit()

    def tearDown(self):
        self.temp.cleanup()

    def login(self, client, username="alice", password="password-test", next_url=None):
        query = {"next": next_url} if next_url is not None else None
        page = client.get("/login", query_string=query)
        self.assertEqual(page.status_code, 200)
        response = client.post(
            "/login",
            query_string=query,
            data={
                "username": username,
                "password": password,
                "csrf_token": hidden_input(page, "csrf_token"),
            },
        )
        return response

    def quick_form(self, client, code="TPS0001"):
        page = client.get(f"/q/{code}")
        self.assertEqual(page.status_code, 200)
        return page, hidden_input(page, "csrf_token"), hidden_input(
            page, "operation_id"
        )

    def post_quick(
        self,
        client,
        *,
        code="TPS0001",
        movement_type="prelievo",
        quantity="1",
        operation_id=None,
        follow_redirects=False,
    ):
        _page, csrf_token, generated_operation_id = self.quick_form(client, code)
        return client.post(
            f"/q/{code}",
            data={
                "csrf_token": csrf_token,
                "operation_id": operation_id or generated_operation_id,
                "movement_type": movement_type,
                "quantity": quantity,
                "destination": "Laboratorio TPS",
                "note": "Test di regressione",
            },
            follow_redirects=follow_redirects,
        )

    def quantities(self):
        with self.app.app_context():
            return {
                row["code"]: row["quantity"]
                for row in get_db().execute(
                    "SELECT code, quantity FROM components ORDER BY code"
                )
            }

    def test_real_csrf_login_and_quick_withdrawal_and_anonymous_rejection(self):
        client = self.app.test_client()
        login = self.login(client, next_url="/q/TPS0001")
        self.assertEqual(login.status_code, 302)
        self.assertEqual(urlsplit(login.headers["Location"]).path, "/q/TPS0001")

        withdrawal = self.post_quick(client, quantity="2")
        self.assertEqual(withdrawal.status_code, 302)
        self.assertEqual(self.quantities()["TPS0001"], 8)

        anonymous = self.app.test_client()
        public_page = anonymous.get("/q/TPS0001")
        self.assertEqual(public_page.status_code, 200)
        rejected = anonymous.post(
            "/q/TPS0001",
            data={
                "operation_id": "anonymous-without-csrf-token-0001",
                "movement_type": "prelievo",
                "quantity": "3",
            },
        )
        self.assertEqual(rejected.status_code, 400)
        self.assertEqual(self.quantities()["TPS0001"], 8)
        with self.app.app_context():
            self.assertEqual(get_db().execute("SELECT COUNT(*) FROM movements").fetchone()[0], 1)

    def test_login_next_rejects_external_protocol_relative_and_backslash_urls(self):
        unsafe_destinations = (
            "https://evil.example/phishing",
            "//evil.example/phishing",
            "/\\evil.example/phishing",
            "\\evil.example/phishing",
        )
        for destination in unsafe_destinations:
            with self.subTest(destination=destination):
                client = self.app.test_client()
                response = self.login(client, next_url=destination)
                self.assertEqual(response.status_code, 302)
                self.assertEqual(
                    response.headers["Location"],
                    "/",
                    f"Il parametro next non sicuro è stato accettato: {destination!r}",
                )

    def test_qr_component_page_prevents_caching_and_indexing(self):
        anonymous = self.app.test_client()
        public_response = anonymous.get("/q/TPS0001")

        authenticated = self.app.test_client()
        self.assertEqual(self.login(authenticated).status_code, 302)
        private_response = authenticated.get("/q/TPS0001")

        for response in (public_response, private_response):
            with self.subTest(authenticated=response is private_response):
                self.assertEqual(response.status_code, 200)
                self.assertIn("no-store", response.headers.get("Cache-Control", "").lower())
                self.assertIn("noindex", response.headers.get("X-Robots-Tag", "").lower())

    def test_operation_id_reuse_with_different_data_component_or_user_is_rejected(self):
        operation_id = "security-idempotency-token-0001"
        alice = self.app.test_client()
        self.assertEqual(self.login(alice).status_code, 302)
        first = self.post_quick(
            alice,
            quantity="2",
            operation_id=operation_id,
        )
        self.assertEqual(first.status_code, 302)
        self.assertEqual(self.quantities(), {"TPS0001": 8, "TPS0002": 7})

        conflicts = []
        conflicts.append(
            self.post_quick(
                alice,
                quantity="3",
                operation_id=operation_id,
                follow_redirects=True,
            )
        )
        conflicts.append(
            self.post_quick(
                alice,
                code="TPS0002",
                quantity="2",
                operation_id=operation_id,
                follow_redirects=True,
            )
        )

        bob = self.app.test_client()
        self.assertEqual(self.login(bob, username="bob").status_code, 302)
        conflicts.append(
            self.post_quick(
                bob,
                quantity="2",
                operation_id=operation_id,
                follow_redirects=True,
            )
        )

        for response in conflicts:
            with self.subTest(response=response):
                self.assertEqual(response.status_code, 200)
                page = response.get_data(as_text=True).lower()
                self.assertNotIn("operazione già registrata", page)
                self.assertTrue(
                    "dati diversi" in page
                    or "già utilizzato" in page
                    or "conflitto" in page,
                    "Il riuso incompatibile deve essere mostrato come errore",
                )

        self.assertEqual(self.quantities(), {"TPS0001": 8, "TPS0002": 7})
        with self.app.app_context():
            db = get_db()
            self.assertEqual(db.execute("SELECT COUNT(*) FROM movements").fetchone()[0], 1)
            movement = db.execute("SELECT * FROM movements").fetchone()
            self.assertEqual(movement["component_id"], 1)
            self.assertEqual(movement["quantity"], 2)
            self.assertEqual(movement["user_id"], self.alice_id)

    def test_history_keeps_operator_name_snapshot_after_user_is_renamed(self):
        client = self.app.test_client()
        self.assertEqual(self.login(client).status_code, 302)
        form = client.get("/movimenti/nuovo")
        self.assertEqual(form.status_code, 200)
        recorded = client.post(
            "/movimenti/nuovo",
            data={
                "csrf_token": hidden_input(form, "csrf_token"),
                "operation_id": hidden_input(form, "operation_id"),
                "component_id": "1",
                "movement_type": "prelievo",
                "quantity": "1",
                "recipient": "Classe destinataria 4A",
                "note": "Nome operatore separato dal destinatario",
            },
        )
        self.assertEqual(recorded.status_code, 302)

        with self.app.app_context():
            db = get_db()
            db.execute(
                "UPDATE users SET full_name = ? WHERE id = ?",
                ("Prof.ssa Alice Rinominata", self.alice_id),
            )
            db.commit()

        history = client.get("/movimenti")
        self.assertEqual(history.status_code, 200)
        self.assertIn(
            b"<td>Prof.ssa Alice Originale</td>",
            history.data,
            "Lo storico deve mostrare il nome dell'operatore al momento del movimento",
        )

    def test_password_change_revokes_an_existing_session_when_versioned(self):
        with self.app.app_context():
            columns = {
                row["name"] for row in get_db().execute("PRAGMA table_info(users)")
            }
        if "auth_version" not in columns:
            self.skipTest("La versione delle credenziali non è implementata")

        old_session = self.app.test_client()
        self.assertEqual(self.login(old_session).status_code, 302)
        self.assertEqual(old_session.get("/").status_code, 200)

        admin = self.app.test_client()
        self.assertEqual(self.login(admin, username="admin").status_code, 302)
        edit_form = admin.get(f"/utenti/{self.alice_id}/modifica")
        self.assertEqual(edit_form.status_code, 200)
        changed = admin.post(
            f"/utenti/{self.alice_id}/modifica",
            data={
                "csrf_token": hidden_input(edit_form, "csrf_token"),
                "full_name": "Prof.ssa Alice Originale",
                "username": "alice",
                "password": "password-nuova",
                "role": "operatore",
                "active": "1",
            },
        )
        self.assertEqual(changed.status_code, 302)

        revoked = old_session.get("/")
        self.assertEqual(revoked.status_code, 302)
        self.assertEqual(urlsplit(revoked.headers["Location"]).path, "/login")


if __name__ == "__main__":
    unittest.main()
