import sys

from PySide6.QtWidgets import QApplication

from app.ui.main_window import WordCompareApp


def main() -> int:
    import multiprocessing

    multiprocessing.freeze_support()

    try:
        from ctypes import windll
        windll.ole32.CoInitialize(None)
    except ImportError:
        pass

    app = QApplication(sys.argv)
    window = WordCompareApp()
    window.show()
    return app.exec()


if __name__ == '__main__':
    sys.exit(main())
