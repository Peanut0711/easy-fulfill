"""첫 화면에 필요 없는 무거운 모듈을 처음 쓰는 시점까지 늦게 불러오는 도우미.

pandas·NumPy처럼 import에 수백 ms가 드는 모듈을 모듈 상단에서 바로 읽으면 앱 첫 화면이
그만큼 늦게 뜬다. ``pd = deferred_module("pandas")``로 두면 ``pd.read_excel``처럼 속성을
처음 쓸 때 실제 모듈을 불러오고, 이후에는 같은 모듈 객체를 그대로 쓴다.
"""

from __future__ import annotations

import importlib
import threading
from types import ModuleType


class DeferredModule:
    """속성에 처음 접근할 때 실제 모듈을 import하는 대리 객체."""

    __slots__ = ("_name", "_module")

    def __init__(self, name: str):
        self._name = name
        self._module: ModuleType | None = None

    def _load(self) -> ModuleType:
        module = self._module
        if module is None:
            # import 잠금이 여러 스레드의 동시 첫 사용을 한 번의 import로 정리한다.
            module = importlib.import_module(self._name)
            self._module = module
        return module

    def __getattr__(self, attr: str):
        return getattr(self._load(), attr)

    def __dir__(self):
        return dir(self._load())

    def __repr__(self) -> str:
        state = "loaded" if self._module is not None else "deferred"
        return f"<DeferredModule {self._name!r} ({state})>"


def deferred_module(name: str) -> DeferredModule:
    """처음 사용할 때 import되는 모듈 대리 객체를 만든다."""
    return DeferredModule(name)


def preload_modules_in_background(*names: str) -> threading.Thread:
    """화면 표시 뒤 첫 작업 대기를 줄이도록 모듈을 백그라운드 스레드에서 미리 import한다."""

    def load():
        for name in names:
            try:
                importlib.import_module(name)
            except Exception as error:  # 미리 읽기 실패는 첫 사용 시 다시 시도된다.
                print(f"! {name} 미리 불러오기 실패: {error}")

    thread = threading.Thread(target=load, name="preload-" + "-".join(names), daemon=True)
    thread.start()
    return thread
