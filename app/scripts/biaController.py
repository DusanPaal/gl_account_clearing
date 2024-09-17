# pylint: disable = C0103, R0911, R1711, W0603, W0703, W1203

"""
The 'biaController.py' module represents the main communication
channel that manages data and control flow between the connected
highly specialized modules.

Version history:
1.0.20210819 - Initial version.
1.0.20220615 - Refactored and simplified code
             - Added/updated docstrings.
"""

from datetime import datetime
from glob import glob
import logging
from logging import config, getLogger
from os import remove, mkdir
from os.path import exists, isfile, join,split
from shutil import move
import sys

from win32com.client import CDispatch
import yaml

import scripts.biaDates as dat
import scripts.biaF03 as f03
import scripts.biaFBL3N as fbl3n
import scripts.biaMail as mail
import scripts.biaProcessor as proc
import scripts.biaReport as rep
import scripts.biaSAP as sap

_logger = getLogger("master")
_ent_states = {}

def _set_entity_state(cocd: str, state: str, val: object):
    """
    Stores a state value for a given entity (company code).

    Params:
        cmp_cd: Company code representing the entity for which a state value will be stored.
        state: State of an entity (e. g. 'exported', 'cleared', ...)
        val: Value to store.

    Returns: None.
    """

    assert (cocd.isnumeric() and len(cocd) == 4) or cocd == "499L", "Argument 'cocd' has incorrect value!"
    assert cocd in _ent_states, f"Company code '{cocd}' not in states!"
    assert state in _ent_states[cocd], f"Invalid state '{state}'!"

    _ent_states[cocd][state] = val

    return

def _get_entity_state(cmp_cd: str, state: str) -> object:
    """
    Returns a state value stored for a given entity (company code).

    Params:
        cmp_cd: Company code representing the entity for which a state value will be stored.
        state: State of an entity.

    Returns: State value for a given entity.
    """

    assert (cmp_cd.isnumeric() and len(cmp_cd) == 4) or cmp_cd == "499L", "Argument 'cmp_cd' has incorrect value!"
    assert cmp_cd in _ent_states, f"Company code '{cmp_cd}' not in states!"
    assert state in _ent_states[cmp_cd], f"Invalid state '{state}'!"

    val = _ent_states[cmp_cd][state]

    return val

def init_logger(app_name: str, app_ver: str) -> bool:
    """
    Initializes application global logger, then creates a new or
    clears the content of a previous log file. Lastly, log header
    containing application name and version is printed to the file.

    Parameters:
        app_name: Name of the application.
        app_ver: Version of the application.

    Returns: True if initialization succeeds, False if it fails.
    """

    # read logging configuration file
    log_cfg_path = join(sys.path[0], "logging.yaml")
    logpath = join(sys.path[0], "log.log")

    try:
        with open(log_cfg_path, 'r', encoding = "utf-8") as file:
            content = file.read()
    except Exception as exc:
        print(exc)
        return False

    log_cfg = yaml.safe_load(content)
    config.dictConfig(log_cfg)

    _logger.setLevel(logging.INFO)

    prev_file_handler = _logger.handlers.pop(1)
    new_file_handler = logging.FileHandler(logpath)
    new_file_handler.setFormatter(prev_file_handler.formatter)
    _logger.addHandler(new_file_handler)

    try: # create an empty / clear previous log file
        with open(logpath, 'w', encoding = "utf-8"):
            pass
    except Exception as exc:
        print(exc)
        return False

    # write log header
    curr_date = datetime.now().strftime("%d-%b-%Y")
    _logger.info(f"{app_name} ver. {app_ver}")
    _logger.info(f"Log date: {curr_date}\n")

    return True

def load_app_config() -> dict:
    """
    Reads and parses a file containing
    application runtime configuration params.

    Params: None.

    Returns: Application runtime configuration params.
    """

    _logger.info("Loading application configuration ...")
    file_path = join(sys.path[0], "appconfig.yaml")

    with open(file_path, 'r', encoding = "utf-8") as file:
        content = file.read()

    repl = content.replace("$app_dir$", sys.path[0])
    cfg = yaml.safe_load(repl)

    return cfg

def load_clearing_rules(file_path: str) -> dict:
    """
    Reads and parses file that contains entity-specific
    rules for GL accounting data procesing.

    Params:
        file_path: Path to the file containing processing rules.

    Returns: A dict object mapping company codes to processing rules.
    """

    _logger.info("Loading clearig rules ...")

    with open(file_path, 'r', encoding = "utf-8") as file:
        content = file.read()

    rules = yaml.safe_load(content)
    active_rules = {}

    for cocd in rules:

        if not rules[cocd]["active"]:
            _logger.warning(f"Company code '{cocd}' excluded from "
            "clearing according to settings in 'rules.yaml'.")
            continue

        s_gl_accs = rules[cocd]["accounts"]
        gl_accs = [acc for acc in s_gl_accs if s_gl_accs[acc]["active"]]

        if len(gl_accs) == 0:
            _logger.warning(f" Company code '{cocd}' excluded from clearing. "
            "Reason: Company code is active but contains no active accounts "
            "according to settings in 'rules.yaml'.")
            continue

        _ent_states[cocd] = {
            "exported": False,
            "cleared": False,
            "no_open_items": False
        }

        active_rules.update({cocd: rules[cocd]})

    if len(active_rules) == 0:
        _logger.warning("No active company code found!")
        return None

    return active_rules

def initialize_sap(cfg: dict) -> CDispatch:
    """
    Manages application connection to SAP GUI
    scripting engine.

    Params:
        cfg: Application configuration data.

    Returns: An initialized GuiSession object.
    """

    sess = sap.login(cfg["gui_exe_path"], sap.SYS_P25)

    return sess

def export_fbl3n_data(data_cfg: dict, sap_cfg: dict, rules: dict, sess: CDispatch) -> bool:
    """
    Manages data export from GL accounts into a local data file.

    Params:
        data_cfg: Configuration parameters for application data processing.
        sap_cfg: Configuration parameters for SAP GUI system and transactions.
        rules: Acconting data processig params mapped to company codes.
        sess: An intialized SAP GuiSession object.

    Returns: True if data export succeeds, False if it fails.
    """

    if not fbl3n.initialize(sess):
        return False

    layout = sap_cfg["fbl3n_layout"]
    exp_dir = data_cfg["export_dir"]
    exported = False

    for cocd in rules:

        all_gl_accs = rules[cocd]["accounts"]
        active_gl_accs = [acc for acc in all_gl_accs if all_gl_accs[acc]["active"]]

        country = rules[cocd]["country"]
        exp_name = data_cfg["fbl3n_data_export_name"].replace("$company_code$", cocd).replace("$country$", country)
        exp_path = join(exp_dir, exp_name)

        if isfile(exp_path):
            _logger.warning(f"Data export for company code '{cocd}' skipped. "
            "Reason: Data already exported into a file in the previous run.")
            _set_entity_state(cocd, "exported", True)
            exported |= True
            continue

        try:
            fbl3n.export(exp_path, cocd, active_gl_accs, layout)
        except RuntimeError as rt_err:
            _set_entity_state(cocd, "exported", False)
            _logger.error(rt_err)
        except RuntimeWarning as rt_wng:
            _set_entity_state(cocd, "no_open_items", True)
            _logger.warning(rt_wng)
        else:
            _set_entity_state(cocd, "exported", True)
            exported |= True

    fbl3n.release()

    return exported

def process_fbl3n_data(data_cfg: dict, rules: dict) -> tuple:
    """
    Manages reading, parsing and evaluation of data that was # preformulovat a spresnit
    previously exported from FBL3N into a local file.

    Params:
        data_cfg: Configuration parameters for application data processing.
        rules: Acconting data processig params mapped to company codes.

    Returns: A tuple of clearing input data and evaluated accounting items.
    """

    output = {}
    match_count = 0
    exp_dir = data_cfg["export_dir"]
    all_items = {}

    for cocd in rules:

        if _get_entity_state(cocd, "no_open_items"):
            _logger.warning(f"FBL3N data for company code '{cocd}' will not be processed. "
            "Reason: No open items found on any of the entered GL accounts.")
            continue

        if not _get_entity_state(cocd, "exported"):
            _logger.warning(f"FBL3N data for company code '{cocd}' will not be processed. "
            "Reason: Accounting data export from FBL3N failed.")
            continue

        _logger.info(f"Processing FBL3N data for company code '{cocd}' ...")
        country = rules[cocd]["country"]
        exp_name = data_cfg["fbl3n_data_export_name"].replace("$company_code$", cocd).replace("$country$", country)
        exp_path = join(exp_dir, exp_name)

        converted = proc.convert_fbl3n_data(exp_path, cocd)

        # if any error occurs during conversion
        if converted is None:
            return (None, None)

        matched = proc.find_matches(converted, rules, cocd)

        # if any error occurs during matching
        if matched is None:
            return (None, None)

        all_items.update({cocd: matched})

        if not matched["Match"].any():
            _logger.info(" No matches were found.")
            continue

        _logger.info(" Generating clearing input ...")
        result = proc.generate_clearing_input(matched, cocd)
        clr_input = result if len(result) > 0 else None
        match_count += matched[~matched["Excluded"] & matched["Match"]].shape[0]
        output.update({cocd: clr_input})

    # items per entire company code
    _logger.info(f" Total items to clear found: {match_count}.")

    return (output, all_items)

def clear_open_items(matches: dict, all_items: dict, clear_cfg: dict, sess: CDispatch) -> bool:
    """
    Manages clearing of open GL account items.

    Params:
        matches: dictionary mapping of account items matches to company code(s)
        all_items: A dict of company codes mapped to their evaluated accounting data.
        clear_cfg: list of Off-work days according to Ledvance fiscal year calender
        sess: An intialized SAP GuiSession object.

    Returns: True on success, False on failure.
    """

    if not f03.start(sess):
        return False

    holidays = clear_cfg["holidays"]

    for cocd, clr_input in matches.items():

        if clr_input is None:
            _logger.info(f"Skipping account clearing for company code '{cocd}'. "
            "Reason: No items matched.")
            continue

        _logger.info(f"Clearing open items for company code '{cocd}' ...")

        for acc in clr_input:

            for curr in clr_input[acc]:

                curr_items = clr_input[acc][curr]
                assigns = curr_items.Unique_Assignments
                refs = curr_items.Unique_References
                doc_nums = curr_items.Unique_Document_Numbers
                idx = curr_items.Indexes

                _logger.info(f" Processing account '{acc}' with currency '{curr}' ...")

                try:
                    pst_num = f03.clear_items([acc], cocd, curr, holidays, curr_items, assigns, refs, doc_nums)
                except PermissionError as p_err:
                    all_items[cocd].loc[idx, "Message"] = f"Clearing error: {p_err}"
                    continue
                except RuntimeError as rt_exc:
                    all_items[cocd].loc[idx, "Message"] = f"Clearing error: {rt_exc}"
                    continue

                _logger.info(f" Items posted under document number '{pst_num}'.")

                all_items[cocd].loc[idx, "Posting_Number"] = pst_num
                all_items[cocd].loc[idx, "Message"] = "Successfully cleared."

        _set_entity_state(cocd, "cleared", True)

    f03.close()

    return True

def create_reports(rep_cfg: dict, rules: dict, all_items: dict):
    """
    Creates excel reports from processed account clearing data.

    Args:
        rep_cfg: Report configuration parameters.
        rules: Acconting data processig params mapped to company codes.
        all_items: A dict of company codes mapped to their evaluated accounting data.

    Returns: None.
    """

    _logger.info("Creating user reports ...")

    for cocd in rules:

        if _get_entity_state(cocd, "no_open_items"):
            _logger.warning(f" Could not create report for company code '{cocd}'. "
            "Reason: There were no open items found for the given company code.")
            continue

        if not _get_entity_state(cocd, "exported"):
            _logger.warning(f" Could not create report for company code '{cocd}'. "
            "Reason: Export of accounting data from FBL3N failed.")
            continue

        country = rules[cocd]["country"]
        rep_name = rep_cfg["name"].replace("$company_code$", cocd).replace("$country$", country)
        rep_path = join(rep_cfg['local_dir'], rep_name)
        sht_name = rep_cfg["sheet_name"]

        if cocd in rep.field_order:
            fields = rep.field_order[cocd]
        else:
            fields = rep.field_order["other"]

        rep.create(all_items[cocd], fields, rep_path, sht_name)

    return

def upload_reports(rep_cfg: dict):
    """
    Creates a new subfolder in the destination folder (if this does not already exist)
    and moves the excel report file from a local source folder to the subfolder.

    Params:
        rep_cfg: Report configuration parameters.

    Returns: None.
    """

    _logger.info("Uploading reports ...")

    src_dir = rep_cfg["local_dir"]
    dst_dir = rep_cfg["net_dir"]
    dst_subdir = dat.get_date().strftime(rep_cfg["net_subdir_format"])

    dst_fldr_path = join(dst_dir, dst_subdir)
    report_paths = glob(join(src_dir, "*.xlsx"))

    if not exists(dst_fldr_path):
        try:
            mkdir(dst_fldr_path)
        except Exception as exc:
            _logger.exception(exc)
            return False

    for rep_path in report_paths:
        rep_name = split(rep_path)[1]
        _logger.debug(f"Uploading report '{rep_name}' from: {src_dir} to: {dst_fldr_path}")
        dst_file_path = join(dst_fldr_path, rep_name)

        try:
            move(rep_path, dst_file_path)
        except Exception as exc:
            _logger.exception(exc)
            return False

    return True

def notify_users(all_items: dict, rep_cfg: dict, notif_cfg: dict, active_cmp_cds: set):
    """
    Summarizes account clearing data into a HTML table, then creates user notifications
    containing the table and clearing report location if available.

    Params:
        all_items: A dict of company codes mapped to their evaluated accounting data.
        rep_cfg: Report configuration parameters.
        notif_cfg: Notification configuration parameters.
        active_cmp_cds: List of active (processed) company codes.

    Returns: None.
    """

    if not notif_cfg["send"]:
        _logger.warning("Sending of notifications to users is disabled in 'appconfig.yaml'.")
        return

    def read_template(file_path: str) -> str:
        with open(file_path, 'r', encoding = "utf-8") as txt_file:
            content = txt_file.read()
        return content

    _logger.info("Sending notifications to users ...")

    templates = notif_cfg["templates"]

    for usr in notif_cfg["users"]:

        name = usr["name"]
        surname = usr["surname"]
        full_name = " ".join([name, surname])
        usr_cmp_cds = usr["company_codes"]
        recip = usr["email"]

        if len(set(usr_cmp_cds).intersection(active_cmp_cds)) == 0:
            _logger.warning(f" Notification will not be sent to user '{full_name}'. "
            "Reason: None of the user's company codes has been processed.")
            continue

        if not usr["send"]:
            _logger.warning(f" Notification will not be sent to user '{full_name}'. "
            "Reason: User excluded from receiving notifications according to settings in 'appconfig.yaml'.")
            continue

        net_subdir = dat.get_date().strftime(rep_cfg["net_subdir_format"])
        net_rep_path = join(rep_cfg["net_dir"], net_subdir)
        notif_dir = notif_cfg["notification_dir"]
        notif_name = notif_cfg["notification_name"].replace("$user_name$", name).replace("$user_surname$", surname)
        notif_path = join(notif_dir, notif_name)
        sender = notif_cfg["sender"]
        time_stamp = dat.get_date().strftime(notif_cfg["date_stamp_format"])
        subject = notif_cfg["subject"].replace("$date$", time_stamp)

        for cmp_cd in active_cmp_cds:
            if cmp_cd in usr_cmp_cds and _get_entity_state(cmp_cd, "no_open_items"):
                usr_cmp_cds.remove(cmp_cd)

        if len(usr_cmp_cds) == 0:
            notif_templ = read_template(templates["no_open_items"])
            notif = notif_templ.replace("$user$", name)
        else:
            notif_templ = read_template(templates["general"])
            summ = rep.summarize(all_items, usr_cmp_cds)
            notif = notif_templ.replace("<TR><TD>$tbl_rows$</TD></TR>", summ)
            notif = notif.replace("$user$", name).replace("$report_path$", net_rep_path)

        try:
            with open(notif_path, 'w', encoding = "utf-8") as n_file:
                n_file.write(notif)
        except Exception as exc:
            _logger.error(f" Failed writing notification to HTML file '{notif_path}'. Reason: {exc}")

        mail.send_message(sender, subject, notif, recip)

    return

def _clean_temp(dir_path: str):
    """
    Deletes application temporary files.
    """

    file_paths = glob(join(dir_path, "**", "*.*"), recursive = True)

    if len(file_paths) == 0:
        _logger.warning("No temporary files to delete found!")
        return

    _logger.info("Deleting temporary files ...")

    for file_path in file_paths:
        try:
            remove(file_path)
        except Exception as exc:
            _logger.exception(exc)
            return

    return

def cleanup(data_cfg: dict = None, sess: CDispatch = None):
    """
    Releases application-allocated resources. If 'data_cfg'
    argument is provided, then all temporary data created during
    runtime will be deleted. If 'sess' argument is provided, then
    the logout from SAP GUI will be performed.

    Params:
        data_cfg: Configuration parameters for application data processing.
        sess: An intialized SAP GuiSession object.

    Returns: None.
    """

    global _ent_states
    _ent_states = None

    # logout from running SAP
    if sess is not None:
        sap.logout(sess)

    # clean up any application temprary data
    if data_cfg is not None:
        _clean_temp(data_cfg["temp_dir"])
