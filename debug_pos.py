import os

from tools.debug_pos import debug_positions


if __name__ == "__main__":
    if os.path.exists("before.docx"):
        debug_positions("before.docx")
