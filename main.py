"""
main.py — Conversor Excel → CNAB BB
"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def check_dependencies():
    missing = []
    for pkg in ["PyQt5", "pandas", "openpyxl"]:
        try:
            __import__(pkg)
        except ImportError:
            missing.append(pkg)
    if missing:
        print(f"Instale as dependências: pip install {' '.join(missing)}")
        sys.exit(1)


def main():
    check_dependencies()
    from gui.app import run
    run()


if __name__ == "__main__":
    main()
