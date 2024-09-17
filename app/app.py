# pylint: disable = C0103, R0911, W1203

"""
The 'app.py' module contains the main entry of the program
that controls application data and operation flow via
the high level wrapper-like 'biaController.py' module.
Application event logger is initiated upon module import.

Version history:
1.0.20210819: - Initial version.
1.1.20220616: - Refactored and simplified code.
              - Removed associative clearing of GL accounts.
              - Updated docstrings.
1.0.20220902: - Calculation of clearing day is now performed
                in a dynamic fashion independently from the actual year.
"""

import logging
import sys
from scripts import biaController as ctrlr

_logger = logging.getLogger("master")

def main() -> int:
    """
    Program entry point.

    Params: None.

    Returns: An integer representing program completion state.
    """

    APP_NAME = "GL Account Clearing"
    APP_VERSION = "1.1.20220616"

    if not ctrlr.init_logger(APP_NAME, APP_VERSION):
        return 1

    _logger.info("=== Initialization ===")
    cfg = ctrlr.load_app_config()
    rules = ctrlr.load_clearing_rules(cfg["clearing"]["rules_path"])

    if rules is None:
        _logger.info("No active entity found.")
        return 2

    sess = ctrlr.initialize_sap(cfg["sap"])

    if sess is None:
        _logger.critical("Failed to connect to SAP.")
        return 3

    _logger.info("=== Processing ===")
    if not ctrlr.export_fbl3n_data(cfg["data"], cfg["sap"], rules, sess):
        _logger.critical("Failed data export or no open items.")
        return 4

    matches, all_items = ctrlr.process_fbl3n_data(cfg["data"], rules)

    if matches is None:
        _logger.critical("Failed FBL3N data processing.")
        return 5

    if len(matches) > 0 and not ctrlr.clear_open_items(matches, all_items, cfg["clearing"], sess):
        _logger.critical("Failed account clearing.")
        return 6

    ctrlr.create_reports(cfg["reports"], rules, all_items)

    if not ctrlr.upload_reports(cfg["reports"]):
        return 7

    ctrlr.notify_users(all_items, cfg["reports"], cfg["notifications"], set(rules.keys()))

    _logger.info("=== Cleanup ===")
    ctrlr.cleanup(cfg["data"], sess)

    return 0

if __name__ == "__main__":
    rc = main()
    _logger.info(f"=== System shutdown with return code: {rc} ===")
    logging.shutdown()
    sys.exit(rc)
