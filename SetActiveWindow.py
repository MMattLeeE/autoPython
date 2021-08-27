import win32com.client
from comtypes.client import GetActiveObject
import win32.win32gui as gui
import re
import time

class WindowMgr:
    """Encapsulates some calls to the winapi for window management"""

    def __init__ (self):
        """Constructor"""
        self._handle = None

    def _window_enum_callback(self, hwnd, wildcard):
        """Pass to win32gui.EnumWindows() to check all the opened windows"""
        if re.match(wildcard, str(gui.GetWindowText(hwnd))) is not None:
            self._handle = hwnd

    def find_window_wildcard(self, wildcard):
        """find a window whose title matches the wildcard regex"""
        self._handle = None
        gui.EnumWindows(self._window_enum_callback, wildcard)

    def set_foreground(self):
        """put the window in the foreground"""
        gui.SetForegroundWindow(self._handle)

w = WindowMgr()
w.find_window_wildcard(".*VACO.*")
w.set_foreground()


app = win32com.client.GetObject(Class='Excel.Application')
if app:
    print('has instance')
