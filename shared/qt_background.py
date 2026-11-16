"""느린 작업(네트워크·엑셀 읽기)을 작업 스레드에서 실행하는 동안 Qt 화면이 멈추지 않게 하는 도우미.

대화상자 사이에 끼어 있는 동기 흐름(확인창 → 조회 → 확인창 …)을 통째로 비동기로 바꾸지 않고,
느린 호출 한 줄만 작업 스레드로 보낸다. 기다리는 동안 화면 다시 그리기·타이머는 처리하지만
사용자 입력은 받지 않아 같은 버튼을 다시 누르는 재진입은 생기지 않는다.
"""

from __future__ import annotations

import threading

from PySide6.QtCore import QCoreApplication, QEventLoop, QTimer

POLL_INTERVAL_MS = 15


def run_in_background(func, *args, **kwargs):
    """func(*args, **kwargs)를 작업 스레드에서 실행하고 결과를 돌려준다(예외는 그대로 다시 발생).

    Qt 애플리케이션이 없거나 이미 작업 스레드에서 불렸으면 그냥 바로 실행한다.
    """
    app = QCoreApplication.instance()
    if app is None or threading.current_thread() is not threading.main_thread():
        return func(*args, **kwargs)
    outcome = {}
    done = threading.Event()

    def target():
        try:
            outcome["value"] = func(*args, **kwargs)
        except BaseException as error:  # 호출한 쪽에서 원래 예외로 처리하도록 넘긴다.
            outcome["error"] = error
        finally:
            done.set()

    worker = threading.Thread(target=target, name=f"background-{getattr(func, '__name__', 'task')}", daemon=True)
    worker.start()
    loop = QEventLoop()
    timer = QTimer()
    timer.setInterval(POLL_INTERVAL_MS)
    timer.timeout.connect(lambda: done.is_set() and loop.quit())
    timer.start()
    if not done.is_set():
        loop.exec(QEventLoop.ProcessEventsFlag.ExcludeUserInputEvents)
    timer.stop()
    worker.join()
    if "error" in outcome:
        raise outcome["error"]
    return outcome.get("value")
