"""Prepare the Windows virtual environment used by run.bat."""

import errno
import hashlib
import msvcrt
from pathlib import Path
import subprocess
import sys
import time


def prepare(root: Path) -> None:
    venv_dir = root / ".venv"
    python = venv_dir / "Scripts" / "python.exe"
    requirements = root / "requirements.txt"
    stamp = venv_dir / "requirements.sha256"

    venv_dir.mkdir(exist_ok=True)
    with (venv_dir / "run-setup.lock").open("a+b") as lock:
        # Windows releases the lock when the file closes, including on process exit.
        while True:
            try:
                msvcrt.locking(lock.fileno(), msvcrt.LK_NBLCK, 1)
                break
            except OSError as exc:
                if exc.errno != errno.EACCES:
                    raise
                time.sleep(0.1)

        digest = hashlib.sha256(requirements.read_bytes()).digest()
        if python.is_file() and stamp.is_file() and stamp.read_bytes() == digest:
            return

        # A partial install can change the environment even when pip fails.
        stamp.unlink(missing_ok=True)
        pip_entry = venv_dir / "Lib" / "site-packages" / "pip" / "__main__.py"
        if not python.is_file() or not pip_entry.is_file():
            print("[setup] .venv를 준비합니다...", flush=True)
            subprocess.run([sys.executable, "-m", "venv", venv_dir], check=True)

        print("[setup] 의존성 설치 중...", flush=True)
        subprocess.run([python, "-m", "pip", "install", "-q", "--upgrade", "pip"])
        subprocess.run([python, "-m", "pip", "install", "-q", "-r", requirements], check=True)
        if hashlib.sha256(requirements.read_bytes()).digest() != digest:
            raise RuntimeError("설치 중 requirements.txt가 변경되었습니다. 다시 실행하세요.")
        stamp.write_bytes(digest)


if __name__ == "__main__":
    try:
        prepare(Path(__file__).resolve().parents[1])
    except (OSError, RuntimeError, subprocess.CalledProcessError) as exc:
        print(f"[setup] {exc}", file=sys.stderr)
        sys.exit(1)
