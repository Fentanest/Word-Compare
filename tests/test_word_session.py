import unittest

from app.services.word_session import WordSession


class _FailingWordApp:
    def __init__(self):
        self.screen_updating = False
        self.quit_called = False

    @property
    def ScreenUpdating(self):
        return self.screen_updating

    @ScreenUpdating.setter
    def ScreenUpdating(self, value):
        self.screen_updating = value

    def Quit(self, SaveChanges=False):
        self.quit_called = True
        raise RuntimeError("COM disconnected during shutdown")


class WordSessionTests(unittest.TestCase):
    def test_exit_swallows_shutdown_disconnect(self):
        session = WordSession()
        app = _FailingWordApp()
        session.word_app = app

        session.__exit__(None, None, None)

        self.assertTrue(app.quit_called)


if __name__ == "__main__":
    unittest.main()
