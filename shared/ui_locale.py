"""Qt 기본 문구(확인 창의 Yes/No/OK/Cancel, 입력칸 오른쪽 클릭 메뉴 등)를 한국어로 보이게 한다.

PySide6에 들어 있는 ``qtbase_ko.qm``(Qt 기본 번역)을 불러온다. 파일이 없으면 영어 그대로 둔다(동작에는
영향 없음). ``shared.ui_style.apply_app_style``이 테마와 함께 부르므로 ``main()``과 측정·캡처 도구가 같다.
"""

TRANSLATION_FILES = ("qtbase_ko",)


def install_korean_translations(app) -> list:
    """번역을 앱에 설치하고 설치한 파일 이름 목록을 돌려준다. 번역기는 앱이 살아 있는 동안 붙잡아 둔다."""
    from PySide6.QtCore import QLibraryInfo, QTranslator

    directory = QLibraryInfo.path(QLibraryInfo.LibraryPath.TranslationsPath)
    installed = getattr(app, "_easy_fulfill_translators", None)
    if installed is None:
        installed = []
        app._easy_fulfill_translators = installed
    names = []
    for name in TRANSLATION_FILES:
        if any(t.objectName() == name for t in installed):
            names.append(name)
            continue
        translator = QTranslator(app)
        translator.setObjectName(name)
        if translator.load(name, directory):
            app.installTranslator(translator)
            installed.append(translator)
            names.append(name)
    return names
