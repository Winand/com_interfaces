import ctypes
from tkinter import Tk

from com_interfaces.interfaces.itaskbarlist3 import TBPF, ITaskBarList3

if __name__ == '__main__':
    root = Tk()
    root.title("Test window")

    # don't group (called before `update_idletasks`)
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID('winand.test.window')
    root.update_idletasks()  # https://stackoverflow.com/a/29159152

    top_level_hwnd = int(root.wm_frame(), 16)
    bar = ITaskBarList3()
    bar.SetProgressValue(top_level_hwnd, 50, 100)
    bar.SetProgressState(top_level_hwnd, TBPF.PAUSED)

    root.mainloop()
