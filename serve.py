import os

from waitress import serve

from app import app


if __name__ == "__main__":
    serve(
        app,
        host=os.environ.get("MAGAZZINO_HOST", "0.0.0.0"),
        port=int(os.environ.get("PORT", "3000")),
        threads=int(os.environ.get("MAGAZZINO_THREADS", "8")),
    )
