# pylint: disable=C0123, C0103, C0301, C0302, E0401, E0611, W0603, W1203, W0703

"""The module automates the clearing of open items posted on GL accpunts."""

from collections import namedtuple
from datetime import date, datetime, timedelta
from logging import getLogger
import time
import numpy as np
from win32com.client import CDispatch

_sess = None
_main_wnd = None
_stat_bar = None

_logger = getLogger("master")

_TYPE_ERROR = "E"
_TYPE_WARNING = "W"

# SAP virtual keys mapping
_vkeys = {
    "Enter": 0,
    "F2": 2,
    "F6": 6,
    "F12": 12,
    "ShiftF2": 14,
    "ShiftF4": 16,
    "PageDown": 82
}

Record = namedtuple(
    "Record", [
    "DC_Amounts",
    "Document_Numbers",
    "Document_Types",
    "Document_Dates",
    "Posting_Dates",
    "Unique_Assignments",
    "Unique_References",
    "Unique_Document_Numbers",
    "All_Assignments",
    "Texts",
    "Trading_Partners",
    "Indexes"
])

def _parse_amount(num: str) -> float:
    """
    Converts amount in the SAP
    string format to a float literal.
    """

    parsed = num.strip()
    parsed = parsed.replace(".", "")
    parsed = parsed.replace(",", ".")

    if num.endswith("-"):
        parsed = parsed.replace("-", "")
        parsed = "".join(["-", parsed])

    return float(parsed)

def _is_popup_dialog() -> bool:
    """Checks if the active window is a popup dialog window."""
    return _sess.ActiveWindow.type == "GuiModalWindow"

def _close_popup_dialog(confirm: bool):
    """Closes a pop-up window by confirming or declining the dialog message."""

    if _sess.ActiveWindow.text == "Information":
        if confirm:
            _main_wnd.SendVKey(_vkeys["Enter"]) # press yes
        else:
            _main_wnd.SendVKey(_vkeys["F12"])   # press no/cancel
        return

    btn_caption = "Yes" if confirm else "No"

    for child in _sess.ActiveWindow.children:
        for grandchild in child.children:
            if grandchild.Type == "GuiButton" and btn_caption == grandchild.text.strip():
                grandchild.Press()
                return

def _set_selection_method(name: str):
    """
    Chooses the open items selection
    criteria based on the method name provided.
    """

    options = _main_wnd.findAllWyName("RF05A-XPOS1", "GuiRadioButton")

    for opt in options:
        if opt.text == name:
            opt.select()
            break

def _get_current_date() -> date:
    """Returns current date."""
    return datetime.now().date()

def _end_of_month(day: date) -> date:
    """Calculates last day of the month for a given day."""

    assert type(day) is date, "Argument 'day' has incorrect type!"

    next_mon = day.replace(day=28) + timedelta(days=4)
    first_day_next_mon = next_mon - timedelta(days=next_mon.day)

    return first_day_next_mon

def _start_of_month(day: date) -> date:
    """Calculates first day of the month for a given day."""
    assert type(day) is date, "Argument 'day' has incorrect type!"
    return day.replace(day = 1)

def _get_month_ultimo(day: date, off_days: list) -> date:
    """Calculates ultimo date for the month of a given day."""

    ultimo = _end_of_month(day)

    while not np.is_busday(ultimo, holidays = off_days):
        ultimo -= timedelta(1)

    return ultimo

def _get_month_uplusone(day: date, off_days: list) -> date:
    """Calculates ultimo plus one date for the month of a given day."""

    upone = _start_of_month(day)

    while not np.is_busday(upone, holidays = off_days):
        upone += timedelta(1)

    return upone

def _get_prev_ultimo(uplusone: date, off_days: list) -> date:
    """Calculates ultimo date corresponding to a given ultimo plus one day."""

    ultimo = uplusone - timedelta(1)

    while not np.is_busday(ultimo, holidays = off_days):
        ultimo -= timedelta(1)

    return ultimo

def _get_actual_off_days(day: date, off_days: list) -> list:
    """Returns a list of company's calculated out of office days."""

    actual = []

    for item in off_days:
        curr_day = date(day.year, item.month, item.day)
        actual.append(curr_day)

    return actual

def _calc_clearing_date(day: date, off_days: list) -> date:
    """Returns a calculated clearing date for items to post."""

    actual_off_days = _get_actual_off_days(day, off_days)
    uplusone = _get_month_uplusone(day, actual_off_days)
    ultimo = _get_month_ultimo(day, actual_off_days)

    if ultimo < day:
        clr_date = ultimo
    elif day <= uplusone:
        clr_date = _get_prev_ultimo(uplusone, actual_off_days)
    else:
        clr_date = day

    return clr_date

def _calc_clearing_period(curr_date: date, clr_date: date) -> int:
    """Calculates posting period based on currently used period."""

    if clr_date == curr_date:
        period = datetime.now().month
    elif datetime.now().month == 1:
        period = 12
    else:
        period = datetime.now().month - 1

    return period

def _get_gui_table(usr_area) -> CDispatch:
    """
    Returns a reference to the 'GuiTableControl'
    object containing list of loaded open items.
    """
    return usr_area.findByName("SAPDF05XTC_6103", "GuiTableControl")

def _get_field_indices(item_tbl: CDispatch) -> dict:
    """
    Returns a dict that maps field technical
    names to layout indexes of the item table.
    """

    mapper = {}

    for idx, col in enumerate(item_tbl.Columns):
        mapper.update({col.name: idx})

    return mapper

def _select_items(usr_area: object, crits: Record) -> CDispatch:
    """
    Scrolls down the list of all loaded open items
    and selects only relevant positions to clear that
    meet the defined criteria.
    """

    amnts = crits.DC_Amounts
    item_tbl = _get_gui_table(usr_area)
    visble_row_count = item_tbl.VisibleRowCount
    loaded_item_count = item_tbl.VerticalScrollbar.Maximum + 1
    cleared_count = len(amnts)

    # select appropriate layout
    fld_indexer = _get_field_indices(item_tbl)

    assign_idx = fld_indexer["RFOPS_DK-ZUONR"]
    docnum_idx = fld_indexer["RFOPS_DK-BELNR"]
    doctype_idx = fld_indexer["RFOPS_DK-BLART"]
    pstdate_idx = fld_indexer["RFOPS_DK-BUDAT"]
    docdate_idx = fld_indexer["RFOPS_DK-BLDAT"]
    partner_idx = fld_indexer["RFOPS_DK-VBUND"]
    text_idx = fld_indexer["RFOPS_DK-SGTXT"]
    amount_idx = fld_indexer["DF05B-PSBET"]

    usr_area.findByName("ICON_SELECT_ALL", "GuiButton").press()

    if loaded_item_count - cleared_count > cleared_count:
        activated = False
        usr_area.findByName("IC_Z-", "GuiButton").press()
    else:
        usr_area.findByName("IC_Z+", "GuiButton").press()
        activated = True

    # set again GuiTableControl object reference as
    # the table got changed by selecting all the items
    item_tbl = _get_gui_table(usr_area)

    selected = []

    # select relevant items from the list
    for row_idx in range(0, loaded_item_count):

        # get the index of the current visible row
        visible_row_idx = row_idx % visble_row_count

        # scroll down on large list to unhide rows
        if visible_row_idx == 0 and row_idx > 0:
            item_tbl.VerticalScrollbar.position = row_idx
            item_tbl = _get_gui_table(usr_area)

        item_assign = item_tbl.GetCell(visible_row_idx, assign_idx).text
        item_docnum = item_tbl.GetCell(visible_row_idx, docnum_idx).text
        item_doctype = item_tbl.GetCell(visible_row_idx, doctype_idx).text
        item_pstdate = item_tbl.GetCell(visible_row_idx, pstdate_idx).text
        item_docdate = item_tbl.GetCell(visible_row_idx, docdate_idx).text
        item_partner = item_tbl.GetCell(visible_row_idx, partner_idx).text
        item_text = item_tbl.GetCell(visible_row_idx, text_idx).text
        item_amount = item_tbl.GetCell(visible_row_idx, amount_idx).text

        # convert to numbers where needed
        item_amount = _parse_amount(item_amount)

        # clean text field string
        item_text = item_text.strip().replace("\"", "")
        item_assign = item_assign.replace("#", "")

        # find the corresponding amount in the seletion critera, then check
        # the current list item against the selection criteria. Deselect the
        # item if no matching criteria was found.
        idx = 0
        matched = False

        while item_amount in amnts[idx:]:

            # get the index of amount found
            if idx > len(amnts) - 1:
                break

            idx = amnts.index(item_amount, idx)

            if idx in selected:
                idx += 1
                continue

            # compare found params with data at the same index in the criteria
            are_equal = []

            are_equal.append(item_pstdate == crits.Posting_Dates[idx])
            are_equal.append(item_docdate == crits.Document_Dates[idx])
            are_equal.append(item_text == crits.Texts[idx])
            are_equal.append(item_assign == crits.All_Assignments[idx])
            are_equal.append(item_docnum == crits.Document_Numbers[idx])
            are_equal.append(item_amount == crits.DC_Amounts[idx])
            are_equal.append(item_partner == crits.Trading_Partners[idx])
            are_equal.append(item_doctype == crits.Document_Types[idx])

            # flag item indexes that met the criteria as processed
            # and select the item in the list for posting
            if all(are_equal):
                matched = True

            if matched and not activated:

                selected.append(idx)
                item_tbl.GetCell(row_idx % visble_row_count, amount_idx).SetFocus()
                _main_wnd.SendVKey(_vkeys["F2"])

                # set again GuiTableControl object reference as
                # the table got changed by double-clicking the cell
                item_tbl = _get_gui_table(usr_area)

                break

            idx += 1

        if activated and not matched:

            item_tbl.GetCell(row_idx % visble_row_count, amount_idx).SetFocus()
            _main_wnd.SendVKey(_vkeys["F2"])

            # set again GuiTableControl object reference as
            # the table got changed by double-clicking the cell
            item_tbl = _get_gui_table(usr_area)

    # perform final balance check, return the posting toolbar
    # if zero, otherwise return to main mask and raise exception
    balance_fld = usr_area.FindByName("RF05A-DIFFB", "GuiTextField")

    if _parse_amount(balance_fld.text) != 0:
        _main_wnd.SendVKey(_vkeys["F12"])
        _main_wnd.SendVKey(_vkeys["F12"])
        _close_popup_dialog(True)
        raise RuntimeError("Non-zero final balance!")

    post_btn = _main_wnd.FindById("tbar[0]/btn[11]")

    return post_btn

def start(sess: CDispatch) -> CDispatch:
    """
    Starts the transaction.

    Params:
    -------
    sess:
        An initialized SAP GuiSession object reference.

    Returns:
    --------
    An initialized transaction search mask reference.
    """
    global _sess
    global _main_wnd
    global _stat_bar

    # set object references
    _sess = sess
    _main_wnd = _sess.findById("wnd[0]")
    _stat_bar = _main_wnd.findById("sbar")

    try:
        _sess.StartTransaction("F-03")
    except Exception as exc:
        _logger.critical(f"Failed to start the transaction. Reason: {exc}")
        return False

    return True

def close():
    """Closes running transaction."""

    global _sess
    global _main_wnd
    global _stat_bar

    _logger.info("Closing F-03 ...")

    try:
        _sess.EndTransaction()
    except Exception as exc:
        _logger.error(f"Failed to close the transaction. Reason: {exc}")
    else:
        if _is_popup_dialog():
            _close_popup_dialog(confirm = True)

    _sess = None
    _main_wnd = None
    _stat_bar = None

def clear_items(
        accs: list, cmp_cd: str, curr: str,
        holidays: list, crits: Record,
        assigns: list = None, refs: list = None,
        doc_nums: list = None
    ) -> CDispatch:
    """
    Clears open items posted on a GL account.

    Params:
        accs: GL accounts with open items to clear.
        cmp_cd: Company code for which ipen items will be cleared.
        curr: Posting currency of the clearing document.
        holidays: List of out of office days according to the company's fiscal year calendar.
        assigns: Assignment values based on which open items will be selected from a GL account.
        refs: Reference values based on which open items will be selected from a GL account.
        doc_nums: Document numbers based on which open items will be selected from a GL account.

    Returns: A reference to the F-03 GuiUserArea object.
    """

    curr_date = _get_current_date()
    calc_date = _calc_clearing_date(curr_date, holidays)
    calc_period = _calc_clearing_period(curr_date, calc_date)

    clr_date = calc_date.strftime("%d.%m.%Y")
    clr_period = str(calc_period)

    for idx, acc in enumerate(accs):

        if idx == 0:

            # set clearing document header data
            _main_wnd.findByName("BKPF-BUDAT", "GuiCTextField").text = clr_date     # Clearing Date
            _main_wnd.findByName("BKPF-MONAT", "GuiTextField").text = clr_period    # Period
            _main_wnd.findByName("BKPF-BUKRS", "GuiCTextField").text = cmp_cd       # Company Code
            _main_wnd.findByName("BKPF-WAERS", "GuiCTextField").text = curr.upper() # Currency

        _main_wnd.findByName("RF05A-AGKON", "GuiCTextField").text = acc             # Account to load items from

        # selection by assignment
        if assigns is not None:

            identifiers = assigns

            if idx != 0:
                _set_selection_method("Assignment")
            else:

                _main_wnd.findAllByName("RF05A-XPOS1", "GuiRadioButton")[11].select()
                _main_wnd.SendVKey(_vkeys["Enter"])

                if _stat_bar.MessageType == _TYPE_ERROR:
                    msg = _stat_bar.Text
                    _main_wnd.SendVKey(_vkeys["F12"])
                    _close_popup_dialog(confirm = True)
                    raise RuntimeError(msg)

                while _stat_bar.MessageType == _TYPE_WARNING:
                    _main_wnd.SendVKey(_vkeys["Enter"])

                _main_wnd.SendVKey(_vkeys["PageDown"])
                _sess.findById("wnd[1]").findAllByName("RF05A-XPOS1", "GuiRadioButton")[4].select()

        elif refs is not None:

            identifiers = refs

            if idx == 0:
                _main_wnd.findAllByName("RF05A-XPOS1", "GuiRadioButton")[3].select()
            else:
                _set_selection_method("Reference")

        elif doc_nums is not None:

            identifiers = doc_nums

            if idx == 0:
                _main_wnd.findAllByName("RF05A-XPOS1", "GuiRadioButton")[1].select()
            else:
                _set_selection_method("Document Number")

        else:
            _main_wnd.findAllByName("RF05A-XPOS1", "GuiRadioButton")[0].select()

        if _stat_bar.MessageType in (_TYPE_WARNING, _TYPE_ERROR):
            msg = _stat_bar.Text
            _main_wnd.SendVKey(_vkeys["F12"])
            _close_popup_dialog(confirm = True)
            raise RuntimeError(msg)

        # confirm to open F03 submask
        # for entering item identificators
        _main_wnd.SendVKey(_vkeys["Enter"])

        if _stat_bar.MessageType == _TYPE_ERROR:

            msg = _stat_bar.Text
            _main_wnd.SendVKey(_vkeys["F12"])
            _close_popup_dialog(confirm = True)

            if "no authorization" in msg:
                _logger.error(f"Loading failed. Reason: {msg}\n")
                raise PermissionError(msg)

            _logger.error(f"Loading failed. Reason: {msg}\n")
            raise RuntimeError(msg)

        while _stat_bar.MessageType == _TYPE_WARNING:
            _main_wnd.SendVKey(_vkeys["Enter"])

        usr_area = _main_wnd.findByName("usr", "GuiUserArea")

        if assigns is None and refs is None and doc_nums is None: # do we need this here???
            return usr_area

        # insert item assignments that identify items located on the processed GL account
        row_count = _main_wnd.FindByName(":SAPMF05A:0731", "GuiSimpleContainer").LoopRowCount

        for idx, item in enumerate(identifiers):

            _main_wnd.FindAllByName("RF05A-SEL01", "GuiTextField")[idx % row_count].text = item

            if idx % row_count == row_count - 1:
                _main_wnd.SendVKey(_vkeys["Enter"])

        # click the "Process open items" button
        _main_wnd.SendVKey(_vkeys["ShiftF4"])

        if _stat_bar.MessageType in (_TYPE_WARNING, _TYPE_ERROR):
            msg = _stat_bar.Text
            _main_wnd.SendVKey(_vkeys["F12"])
            _main_wnd.SendVKey(_vkeys["F12"])
            _close_popup_dialog(True)
            _logger.error(f"Clearing failed. Reason: {msg}\n")
            raise RuntimeError(msg)

        if idx < len(accs) - 1:
            # expected further loading from another account
            _main_wnd.SendVKey(_vkeys["ShiftF2"])
            _main_wnd.SendVKey(_vkeys["F6"])
            return None

    post_btn = _select_items(usr_area, crits)

    # confirm posting
    post_btn.press()

    # just warnings, press enter to continue
    while _stat_bar.MessageType == _TYPE_WARNING:
        _main_wnd.SendVKey(_vkeys["Enter"])

    # pause code execution for a couple of secs
    # to allow SAP unlock the DB for changes
    time.sleep(5)

    if _stat_bar.MessageType == _TYPE_ERROR:
        msg = _stat_bar.text
        _main_wnd.SendVKey(_vkeys["F12"])
        _close_popup_dialog(confirm = True)
        _logger.error(f"Posting failed. Reason: {msg}")
        raise RuntimeError(msg)

    # get the posting number from the status bar text
    tokens = _stat_bar.text.split()
    pst_num = int(tokens[1])


    return pst_num
