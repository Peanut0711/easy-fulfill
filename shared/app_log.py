"""프로그램 로그: 콘솔(cmd)에 찍히는 줄과 로그 전용 줄을 메모리에 모아 「도움말 › 프로그램 로그」 창에 보인다.

- install(): sys.stdout·sys.stderr를 감싸 print 출력을 콘솔에 그대로 쓰면서 로그에도 남긴다.
- log(text): 콘솔에는 쓰지 않고 로그에만 남긴다(시작 단계처럼 평소엔 볼 필요 없는 줄).
- 여러 스레드에서 써도 되도록 잠금을 쓰고, Qt에 의존하지 않는다(창은 entries_since로 새 줄을 가져간다).
"""

import sys
import threading
import time
from collections import deque

MAX_ENTRIES = 3000

_lock = threading.Lock()
_entries = deque(maxlen=MAX_ENTRIES)  # (번호, 시각 "HH:MM:SS", 글자, 오류 출력 여부)
_next_seq = 1
_pending = {}  # (스레드 id, 오류 출력 여부) → 아직 줄바꿈이 오지 않은 글자


def _append_locked(text, is_error):
    global _next_seq
    _entries.append((_next_seq, time.strftime("%H:%M:%S"), text, is_error))
    _next_seq += 1


def log(text):
    """콘솔에는 쓰지 않고 로그 창에만 한 줄(여러 줄이면 줄마다) 남긴다."""
    with _lock:
        for line in str(text).splitlines() or [""]:
            _append_locked(line, False)


def _record(chunk, is_error):
    if not chunk:
        return
    key = (threading.get_ident(), is_error)
    with _lock:
        text = _pending.pop(key, "") + chunk
        *lines, rest = text.split("\n")
        for line in lines:
            _append_locked(line.rstrip("\r"), is_error)
        if rest:
            _pending[key] = rest


def entries_since(seq):
    """번호가 seq보다 큰 줄을 [(번호, 시각, 글자, 오류 여부), …]로 돌려준다."""
    with _lock:
        return [entry for entry in _entries if entry[0] > seq]


def clear():
    with _lock:
        _entries.clear()
        _pending.clear()


class _TeeStream:
    """원래 출력(없으면 생략)에 쓰고 로그에도 남긴다. 그 밖의 속성은 원래 출력에 맡긴다."""

    def __init__(self, original, is_error):
        self._original = original
        self._is_error = is_error

    def write(self, chunk):
        _record(chunk, self._is_error)
        if self._original is None:
            return len(chunk)
        try:
            return self._original.write(chunk)
        except Exception:
            return len(chunk)

    def flush(self):
        if self._original is not None:
            try:
                self._original.flush()
            except Exception:
                pass

    def __getattr__(self, name):
        if self._original is None:
            raise AttributeError(name)
        return getattr(self._original, name)


def install():
    """print 출력을 로그에도 남기도록 sys.stdout·sys.stderr를 감싼다. 여러 번 불러도 한 번만 감싼다."""
    if not isinstance(sys.stdout, _TeeStream):
        sys.stdout = _TeeStream(sys.stdout, False)
    if not isinstance(sys.stderr, _TeeStream):
        sys.stderr = _TeeStream(sys.stderr, True)
