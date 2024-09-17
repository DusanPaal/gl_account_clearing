"""
The 'biaSAP.py' module uses win32com (package: pywin32) to connect
to SAP GUI scripting engine and login/logout to/from the defined
SAP system.

Version history:
1.0.20210311 - Initial version.
1.0.20220615 - Refactored and simplified code, added/updated docstrings.
"""

from logging import Logger, getLogger
from os.path import isfile
import subprocess
import win32com.client
from win32com.client import CDispatch
import win32ui

SYS_P25 = "OG ERP: P25 Productive SSO"
SYS_Q25 = "OG ERP: Q25 Quality Assurance SSO"

_logger: Logger = getLogger("master")

def _window_exists(name: str) -> bool:
    """Checks wheter SAP GUI process is running."""

    try:
        win32ui.FindWindow(None, name)
    except win32ui.error:
        return False
    else:
        return True

def login(gui_exe_path: str, sys_name: str) -> CDispatch:
    """
    Logs into SAP GUI from the SAP Logon window and returns a system connection session.

    Params:
        gui_exe_path: Path to the client SAP GUI executable.
        sys_name: Name of the SAP system to log in.

    Returns: An initialized SAP GuiSession object.
    """

    assert isfile(gui_exe_path), f"SAP GUI executable not found at path: {gui_exe_path}!"
    assert sys_name in (SYS_P25, SYS_Q25), "Invalid system name! Login to SAP using one of the dedicated constants!"

    _logger.info("Logging to SAP ...")
    _logger.debug(f" Params: system: {sys_name}; executable path: {gui_exe_path}")

    if not _window_exists("SAP Logon 750"):

        try:
            _logger.debug(" Starting SAP GUI process ...")
            proc = subprocess.Popen(gui_exe_path)
        except Exception as exc:
            _logger.critical(f" Could not start SAP GUI application. Reason: {exc}")
            return None

        try:
            proc.communicate(timeout = 8)
        except Exception:
            _logger.debug(" Attempting to communicate with SAP GUI timed out.")

    try:
        _logger.debug(" Connecting to SAP GUI ...")
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
    except Exception as exc:
        _logger.critical(f" Could not bind SapGuiAuto reference. Reason: {exc}")
        return None

    try:
        _logger.debug(" Connecting to scripting engine ...")
        app = SapGuiAuto.GetScriptingEngine
    except Exception as exc:
        _logger.critical(f" Could not connect to SAP scripting engine. Reason: {exc}")
        return None

    try:
        if app.Children.Count == 0:
            _logger.debug(" Opening connection to SAP backend process ...")
            Conn = app.OpenConnection(sys_name, True)
        else:
            _logger.debug(" An open connection to SAP backend process found.")
            Conn = app.Children(0)
    except Exception as exc:
        _logger.critical(f" Could not create connection to SAP. Reason: {exc}")
        return None

    try:
        _logger.debug(" Creating SAP session ...")
        Sess = Conn.Children(0)
    except Exception as exc:
        _logger.critical(f" Could not create new session. Reason: {exc}")
        return None

    _logger.debug("Successfully logged in.")

    return Sess

def logout(sess: CDispatch):
    """
    Logs out from an active SAP GUI system.

    Params:
        sess: An initialized SAP GuiSession object.

    Returns: None.
    """

    assert sess is not None, "Trying to log out from an unitialized SAP GUI!"
    assert type(sess) is CDispatch and sess.type == "GuiSession", "Argument 'sess' has invalid type!"

    _logger.info("Logging out from SAP GUI application ...")

    try:
        sess.findById("wnd[0]/tbar[0]/okcd").text = "/nex"
        sess.findById("wnd[0]").sendVKey(0)
    except Exception as exc:
        _logger.error(f" Unable to close logout from SAP GUI! Reason: {exc}")

    return
