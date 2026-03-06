"""
run_app.py — Point d'entrée PyInstaller pour TCO Automator.

En mode frozen (EXE), sys.executable = TCO_Automator.exe, donc subprocess
ne peut pas lancer Streamlit. On utilise stcli.main() directement.
En mode développement, subprocess est utilisé normalement.
"""

import os
import sys

import dotenv  # noqa: F401 — Force l'inclusion dans le bundle PyInstaller


def main() -> None:
    import streamlit.web.cli as stcli  # noqa: PLC0415

    if getattr(sys, "frozen", False):
        # Mode EXE PyInstaller : app.py est dans sys._MEIPASS
        app_path = os.path.join(sys._MEIPASS, "app.py")
    else:
        # Mode développement normal
        here = os.path.dirname(os.path.abspath(__file__))
        app_path = os.path.join(here, "app.py")

    sys.argv = [
        "streamlit",
        "run",
        app_path,
        "--global.developmentMode=false",
        "--server.headless=false",
    ]

    try:
        sys.exit(stcli.main())
    except SystemExit:
        raise
    except KeyboardInterrupt:
        sys.exit(0)


if __name__ == "__main__":
    main()
