"""Build dist/조회서생성/ (exe + templates). Input values are not bundled; the app saves 입력값.json next to the exe.

Usage: .venv\\Scripts\\python build.py
"""
import shutil
import subprocess
from pathlib import Path

HERE = Path(__file__).resolve().parent
NAME = "조회서생성"

subprocess.run([str(HERE / ".venv/Scripts/pyinstaller.exe"), "--noconfirm", "--windowed", "--name", NAME,
                "--distpath", str(HERE / "dist"), "--workpath", str(HERE / "build"),
                "--specpath", str(HERE / "build"), str(HERE / "app.py")], check=True)
dst = HERE / "dist" / NAME
shutil.copytree(HERE / "templates", dst / "templates", dirs_exist_ok=True)
print(f"완료: {dst}")
