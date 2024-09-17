# pylint: disable = C0103, W0603, W0703, W1203

"""
Description:
The 'biaFBL3N.py' module automates the standard SAP GUI FBL3N transaction in order
to load and export accounting data located on GL accounts into a plain text file.

Version history:
1.0.20210720 - initial version
1.0.20210614 - Refactored and simplified code. Copying to clipboard now mediated
               by 'pyperclip' library instead of previously used 'tkinter' module.
"""

from datetime import datetime
from logging import Logger, getLogger
from os.path import exists
from pyperclip import copy as copy_to_clipboard
from win32com.client import CDispatch

_sess = None
_main_wnd = None
_stat_bar = None

_logger: Logger = getLogger("master")

# keyboard to SAP virtual keys mapping
_vkeys = {
    "Enter":       0,
    "F3":          3,
    "F8":          8,
    "F9":          9,
    "CtrlS":       11,
    "F12":         12,
    "ShiftF4":     16,
    "ShiftF12":    24,
    "CtrlF1":      25,
    "CtrlF8":      32,
    "CtrlShiftF6": 42
}

def _is_popup_dialog() -> bool:
    """
    Checks if the active window is a popup dialog window.
    """

    is_pupup = (_sess.ActiveWindow.type == "GuiModalWindow")

    return is_pupup

def _close_popup_dialog(confirm: bool):
    """
    Closes a pop-up window by confirming or declining the dialog message.

    Params:
        confirm: True if the pop-up dialog prompt should be confirmed, False if declined.

    Returns: None.
    """

    if _sess.ActiveWindow.text == "Information":
        if confirm:
            _main_wnd.SendVKey(_vkeys["Enter"]) # press yes
        else:
            _main_wnd.SendVKey(_vkeys["F12"])   # press no/cancel
        return

    btn_caption = "Yes" if confirm else "No"

    for child in _sess.ActiveWindow.Children:
        for grandchild in child.Children:
            if grandchild.Type == "GuiButton" and btn_caption == grandchild.text.strip():
                grandchild.Press()
                return

def _set_company_code(val: str):
    """
    Enters company code into the 'Company code' field
    located on the main transaction window.
    """

    if _main_wnd.findAllByName("SD_BUKRS-LOW", "GuiCTextField").count > 0:
        _main_wnd.findByName("SD_BUKRS-LOW", "GuiCTextField").text = val
    elif _main_wnd.findAllByName("SO_WLBUK-LOW", "GuiCTextField").count > 0:
        _main_wnd.findByName("SO_WLBUK-LOW", "GuiCTextField").text = val

def _set_layout(val: str):
    """
    Enters layout name into the 'Layout' field
    located on the main transaction window.
    """

    _main_wnd.findByName("PA_VARI", "GuiCTextField").text = val

def _set_accounts(accs: list):

    # remap vals to str since accounts
    # may be passed in as ints
    vals = list(map(str, accs))

    # open selection table for company codes
    _main_wnd.findByName("%_SD_SAKNR_%_APP_%-VALU_PUSH", "GuiButton").press()

    _main_wnd.SendVKey(_vkeys["ShiftF4"])   # clear any previous values
    copy_to_clipboard("\r\n".join(vals))    # copy accounts to clipboard
    _main_wnd.SendVKey(_vkeys["ShiftF12"])  # confirm selection
    copy_to_clipboard("")                   # clear the clipboard
    _main_wnd.SendVKey(_vkeys["F8"])        # confirm

def _set_item_selection_date():
    """
    Selects 'Open items' selection option and
    enters key date value to the corresponding field.
    """

    exp_date = datetime.date(datetime.now()).strftime("%d.%m.%Y")
    _logger.debug(f"Export date: {exp_date}")

    _main_wnd.findByName("X_OPSEL", "GuiRadioButton").select()
    _main_wnd.FindByName("SO_BUDAT-LOW", "GuiCTextField").text = exp_date

def _toggle_worklist(activate: bool):
    """
    Activates or deactivates the 'Use worklist' option
    in the transaction main search mask.
    """

    used = _main_wnd.FindAllByName("PA_WLSAK", "GuiCTextField").Count > 0

    if (activate or used) and not (activate and used):
        _main_wnd.SendVKey(_vkeys["CtrlF1"])

def _set_export_params(path: str, name: str, enc: str = "4120"):
    """
    Enters folder path, file name and encoding of the file to which the
    exported data will be written.
    """

    _sess.FindById("wnd[1]").FindByName("DY_PATH", "GuiCTextField").text = path
    _sess.FindById("wnd[1]").FindByName("DY_FILENAME", "GuiCTextField").text = name
    _sess.FindById("wnd[1]").FindByName("DY_FILE_ENCODING", "GuiCTextField").text = enc

def _select_data_format(idx: int):
    """
    Selects data export format from the export options dialog
    based on the option index on the list.
    """

    option_wnd = _sess.FindById("wnd[1]")
    option_wnd.FindAllByName("SPOPLI-SELFLAG", "GuiRadioButton")(idx).Select()

def initialize(sess: CDispatch) -> bool:
    """
    Initializes module private fields and starts the transaction.

    Params:
        sess: An initialized SAP GuiSession object reference.

    Returns: True if transaction starts successfully, False if not.
    """

    assert isinstance(sess, CDispatch), "Argument 'sess' has incorrect type!"
    assert sess.type == "GuiSession", "Argument 'sess' has incorrect type!"

    global _sess
    global _main_wnd
    global _stat_bar

    _logger.info("Starting FBL3N ...")

    # set SAP object references
    _sess = sess
    _main_wnd = _sess.findById("wnd[0]")
    _stat_bar = _main_wnd.findById("sbar")

    try:
        _sess.StartTransaction("FBL3N")
    except Exception as exc:
        _logger.critical(f"Failed to start the transaction. Reason: {exc}")
        return False

    return True

def release():
    """
    Closes running transaction and releases module private resources.

    Params: None.

    Returns: None.
    """

    global _sess
    global _main_wnd
    global _stat_bar

    assert _sess is not None, "Trying to release an uninitialized module!"

    _logger.info("Closing FBL3N ...")

    try:
        _sess.EndTransaction()
    except Exception as exc:
        _logger.error(f"Failed to close the transaction. Reason: {exc}")
    else:
        if _is_popup_dialog():
            _close_popup_dialog(confirm = True)

    _sess = None
    _stat_bar = None
    _main_wnd = None

def export(file_path: str, cmp_cd: str, gl_accs: list, layout: str):
    """
    Exports open items data from GL accounts into a plain text file. If export fails, a RuntimeError exception
    will be raised. If no open items are found on account(s), a RuntimeWarning exception will be raised.

    Params:
        file_path: Path to the exported text file.
        cmp_cds: Company codes for which the data export will be performed.
        gl_accs: GL accounts for which the data export will be performed.
        exp_date: Date for which open accounting items will be loaded.
        layout: Name of FBL3N layout defining data fields to export.

    Returns: None.
    """

    folder_path = file_path[0:file_path.rfind("\\")]
    file_name = file_path[file_path.rfind("\\") + 1:]

    assert _sess is not None and _main_wnd is not None, "Transaction not initialized! Use the start() method to run the transaction first."
    assert exists(folder_path), "Destination folder not found!"
    assert type(cmp_cd) is str, "Argument 'cmp_cd' has incorrect type!"
    assert len(cmp_cd) == 4, "Invalid company code! Value shoud be a 4-digit string (e.g. '0075')"
    assert all([type(acc) in (str, int) and len(str(acc)) == 8 for acc in gl_accs]), "Invalid GL account(s)!"
    assert type(layout) is str, "Argument 'layout' has incorrect type!"

    _logger.info(f"Exporting FBL3N data for company code '{cmp_cd}' ...")

    _toggle_worklist(activate = False)
    _set_company_code(cmp_cd)
    _set_layout(layout)
    _set_accounts(gl_accs)
    _set_item_selection_date()
    _main_wnd.SendVKey(_vkeys["F8"]) # load item list

    try: # SAP crash can be caught only after next statement following item loading
        msg = _stat_bar.Text
    except Exception as exc:
        raise RuntimeError(
            "Data export failed. "
            f"Reason: Connection to SAP lost. Details: {exc}"
        ) from exc

    if "No items selected" in msg:
        raise RuntimeWarning(
            f"Could not export data for company code '{cmp_cd}'. "
            "Reason: No items found for the used selection criteria."
        )

    if "items displayed" not in msg:
        raise RuntimeError(f"Data export failed. Reason: {msg}")

    _main_wnd.SendVKey(_vkeys["CtrlF8"])        # open layout mgmt dialog
    _main_wnd.SendVKey(_vkeys["CtrlShiftF6"])   # toggle technical names
    _main_wnd.SendVKey(_vkeys["Enter"])         # Confirm Layout Changes
    _main_wnd.SendVKey(_vkeys["F9"])            # open local data file export dialog
    _select_data_format(0)                      # set plain text data export format
    _main_wnd.SendVKey(_vkeys["Enter"])         # confirm
    _set_export_params(folder_path, file_name)  # enter data export file name and folder path
    _main_wnd.SendVKey(_vkeys["CtrlS"])         # replace an exiting file
    _main_wnd.SendVKey(_vkeys["F3"])            # Load main mask

    _logger.debug(f"Data exported to file: {file_path}")
