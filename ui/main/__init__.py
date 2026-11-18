"""메인 창(MainWindow)을 영역별 믹스인으로 나눈 모듈들(docs/ui-redesign-plan.md 1-2).

믹스인은 ``easy-fulfill.py``에 남아 있는 모듈 전역(작업 스레드 클래스, ``gspread``, 상수 등)을
``host(self).이름``으로 실행할 때마다 찾는다. 옮기기 전과 똑같이 그 모듈의 현재 값을 쓰므로,
테스트가 ``patch.object(app_module, "이름", …)``로 바꾼 값도 그대로 따른다.
"""


class _HostNamespace:
    """모듈 전역 dict를 속성으로 읽는 얇은 창구."""

    __slots__ = ("_globals",)

    def __init__(self, module_globals):
        object.__setattr__(self, "_globals", module_globals)

    def __getattr__(self, name):
        try:
            return self._globals[name]
        except KeyError:
            raise AttributeError(f"easy-fulfill.py 전역에 없는 이름: {name}") from None


_registered_globals = None


def register_host(module_globals):
    """``easy-fulfill.py``가 불러와질 때 자기 전역 dict를 등록한다.

    MainWindow 없이 메서드만 떼어 만든 가짜 창(테스트)에서 쓰는 예비 경로다. 같은 프로세스에서
    ``easy-fulfill.py``를 여러 번 불러오면 마지막 것이 등록된다.
    """
    global _registered_globals
    _registered_globals = module_globals


def host(obj):
    """``obj``(메인 창)를 정의한 ``easy-fulfill.py`` 모듈의 전역을 속성처럼 읽는 객체."""
    for cls in type(obj).__mro__:
        if cls.__name__ == "MainWindow":
            init = vars(cls).get("__init__")
            if init is not None:
                return _HostNamespace(init.__globals__)
    if _registered_globals is None:
        raise LookupError("easy-fulfill.py가 아직 불러와지지 않았습니다.")
    return _HostNamespace(_registered_globals)
