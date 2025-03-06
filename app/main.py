import os
import sys

from flask import Flask

project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), "../"))
sys.path.append(project_root)

from app import create_app

app = create_app()


if __name__ == "__main__":
    app.run(port=5000, debug=True)
