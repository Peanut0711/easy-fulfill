"""브라우저 프로필(쿠팡 WING·우체국 포털)을 한 번에 한 작업만 쓰도록 하는 프로세스 간 잠금.

같은 Chromium 프로필을 두 프로세스가 동시에 열면 뒤의 실행이 알기 어려운 오류로 실패하거나,
세션 삭제가 사용 중인 파일 때문에 일부만 지워진다. 프로필 폴더 옆의 잠금 파일을 OS 파일 잠금으로
잡고, 잡은 프로세스가 끝나면(정상·비정상 모두) OS가 잠금을 푼다.
"""

from __future__ import annotations

import os
from pathlib import Path


class ProfileBusyError(RuntimeError):
    """다른 작업이 같은 브라우저 프로필을 쓰고 있다."""


_held: dict[str, object] = {}


def _lock_path(profile_dir: Path) -> Path:
    # 세션 삭제(프로필 폴더 삭제)와 무관하도록 프로필 폴더 밖, 같은 위치에 둔다.
    return profile_dir.parent / f"{profile_dir.name}.lock"


def hold_profile_lock(profile_dir: Path, label: str) -> None:
    """이 프로세스가 끝날 때까지 프로필을 단독으로 쓴다. 다른 프로세스가 쓰는 중이면 ProfileBusyError."""
    path = _lock_path(Path(profile_dir))
    key = str(path.resolve())
    if key in _held:
        return  # 같은 프로세스 안에서 다시 여는 경우(로그인 확인 뒤 재실행 등)
    path.parent.mkdir(parents=True, exist_ok=True)
    handle = open(path, "a+b")
    try:
        if os.name == "nt":
            import msvcrt
            handle.seek(0)
            msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
        else:
            import fcntl
            fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
    except OSError:
        handle.close()
        raise ProfileBusyError(
            f"{label} 브라우저를 다른 작업이 사용 중입니다. "
            "열려 있는 작업(브라우저 창)을 마치거나 닫은 뒤 다시 실행해 주세요.") from None
    _held[key] = handle


def release_profile_lock(profile_dir: Path) -> None:
    """이 프로세스가 잡은 프로필 잠금을 푼다(프로세스 종료 시에는 OS가 자동으로 푼다)."""
    handle = _held.pop(str(_lock_path(Path(profile_dir)).resolve()), None)
    if handle is not None:
        handle.close()
