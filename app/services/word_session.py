try:
    import win32com.client as win32
except ImportError:
    win32 = None


class WordSession:
    def __init__(self):
        self.word_app = None

    def __enter__(self):
        if win32 is None:
            raise RuntimeError("pywin32 또는 Microsoft Word 자동화 환경을 찾을 수 없습니다.")

        self.word_app = win32.gencache.EnsureDispatch("Word.Application")
        self.word_app.Visible = False
        self.word_app.DisplayAlerts = 0
        self.word_app.ScreenUpdating = False
        return self.word_app

    def __exit__(self, exc_type, exc, tb):
        if not self.word_app:
            return

        try:
            self.word_app.ScreenUpdating = True
        except Exception:
            pass

        self.word_app.Quit(SaveChanges=False)

    @staticmethod
    def ensure_hidden(word_app) -> None:
        try:
            if word_app.Visible:
                word_app.Visible = False
        except Exception:
            pass
