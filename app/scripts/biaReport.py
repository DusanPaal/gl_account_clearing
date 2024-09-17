# pylint: disable = C0103, E0110

"""
The 'biaReport.py' module manages:
    - creating exel reports containing the account clearing output data,
    - creating user notification summaries containing clearing and accounting status.

Version history:
1.0.20210722 - Initial version.
1.0.20220615 - Refactored and simplified code.
             - Added/updated docstrings.
"""

from datetime import date, datetime
import pandas as pd
from pandas import DataFrame, ExcelWriter, Index, Series

field_order = {

    "499L": [
        "Account", "Currency", "DC_Amount", "Document_Date",
        "Document_Number", "Document_Type", "Posting_Date",
        "Assignment", "Reference", "Trading_Partner", "Text",
        "Deal_Number", "Match", "Posting_Number", "Message"
    ],

    "1052": [
        "Account", "Currency", "DC_Amount", "Value_Date",
        "Document_Number", "Document_Type", "Document_Date",
        "Posting_Date", "Assignment", "Reference", "Trading_Partner",
        "Text", "Deal_Number", "Match", "Posting_Number", "Message"

    ],

    "other": [
        "Account", "Currency", "DC_Amount", "Document_Number",
        "Document_Type", "Document_Date", "Posting_Date",
        "Assignment", "Reference", "Trading_Partner", "Text",
        "Match", "Posting_Number", "Message"
    ]

}

def _to_excel_serial(val: date) -> int:
    """
    Converts a date object into
    excel-compatible date integer
    (serial) format.
    """

    delta = val - datetime(1899, 12, 30).date()
    days = delta.days

    return days

def _replace_col_char(col_names: Index, char: str, repl: str) -> Index:
    """
    Replaces a given character in data column names with a string.
    """

    col_names = col_names.str.replace(char, repl, regex = False)

    return col_names

def _col_to_rng(data: DataFrame, first_col: str, last_col: str = None,
                row: int = -1, last_row: int = -1) -> str:
    """
    Generates excel data range notation (e.g. 'A1:D1', 'B2:G2').
    If 'last_col' is None, then only single-column range will be
    generated (e.g. 'A:A', 'B1:B1'). if 'row' is '-1', then the
    generated range will span all the column(s) rows (e.g. 'A:A', 'E:E').

    Params:
    ------
    data: Data for which colum names should be converted to a range.
    first_col: Name of the first column.
    last_col: Name of the last column.
    row: Index of the row for which the range will be generated.

    Returns:
    --------
    Excel data range notation.
    """

    if isinstance(first_col, str):
        first_col_idx = data.columns.get_loc(first_col)
    elif isinstance(first_col, int):
        first_col_idx = first_col
    else:
        assert False, "Argument 'first_col' has invalid type!"

    first_col_idx += 1
    prim_lett_idx = first_col_idx // 26
    sec_lett_idx = first_col_idx % 26

    lett_a = chr(ord('@') + prim_lett_idx) if prim_lett_idx != 0 else ""
    lett_b = chr(ord('@') + sec_lett_idx) if sec_lett_idx != 0 else ""
    lett = "".join([lett_a, lett_b])

    if last_col is None:
        last_lett = lett
    else:

        if isinstance(last_col, str):
            last_col_idx = data.columns.get_loc(last_col)
        elif isinstance(last_col, int):
            last_col_idx = last_col
        else:
            assert False, "Argument 'last_col' has invalid type!"

        last_col_idx += 1
        prim_lett_idx = last_col_idx // 26
        sec_lett_idx = last_col_idx % 26

        lett_a = chr(ord('@') + prim_lett_idx) if prim_lett_idx != 0 else ""
        lett_b = chr(ord('@') + sec_lett_idx) if sec_lett_idx != 0 else ""
        last_lett = "".join([lett_a, lett_b])

    if row == -1:
        rng = ":".join([lett, last_lett])
    elif first_col == last_col and row != -1 and last_row == -1:
        rng = f"{lett}{row}"
    elif first_col == last_col and row != -1 and last_row != -1:
        rng = ":".join([f"{lett}{row}", f"{lett}{last_row}"])
    elif first_col != last_col and row != -1 and last_row == -1:
        rng = ":".join([f"{lett}{row}", f"{last_lett}{row}"])
    elif first_col != last_col and row != -1 and last_row != -1:
        rng = ":".join([f"{lett}{row}", f"{last_lett}{last_row}"])
    else:
        assert False, "Undefined argument combination!"

    return rng

def _get_col_width(vals: Series, col_name: str):
    """
    Returns excel column width calculated as
    the maximum count of characters contained
    in the column name and column data strings.
    """

    lens = vals.astype("string").dropna().str.len()
    lens = list(lens)
    lens.append(len(str(col_name)))
    width = max(lens)

    if col_name != "Message":
        width += 6 # increase width for all fields except 'Message'

    return width

def summarize(data: dict, usr_cocd: list) -> str:
    """
    Summarizes data processing results for given user
    company codes into a HTML table across company codes,
    accounts and currencies.

    Params:
    -------
    data:
        A dict of all processed company codes (keys)
        with their accounting data (values).

    usr_cocd:
        User-owned company codes.

    Returns:
    --------
    Notification HTML template updated on processing summary.
    """

    tbl_rows = []

    for cmp_cd in usr_cocd:

        if cmp_cd not in data:
            continue

        for acc in data[cmp_cd]["Account"].unique():

            acc_subset = data[cmp_cd].query(f"Account == '{acc}'")
            acc_currs = acc_subset[acc_subset["Account"] == acc]["Currency"].unique()

            for curr in acc_currs:

                curr_items = acc_subset[(acc_subset["Account"] == acc) & (acc_subset["Currency"] == curr)]
                match_count = curr_items[curr_items["Match"]].shape[0]
                err_count = curr_items[curr_items["Message"].str.contains("error", regex=True)].shape[0]
                clear_count = match_count - err_count
                left_count = curr_items[~curr_items["Message"].str.contains("cleared") & (curr_items["Account"] == acc)].shape[0]
                note = "" # left intentionally empty

                row = f"""
                    <TR>
                        <TD style="BORDER: purple 2px solid; PADDING: 5px">{cmp_cd}</TD>
                        <TD style="BORDER: purple 2px solid; PADDING: 5px">{acc}</TD>
                        <TD style="BORDER: purple 2px solid; PADDING: 5px">{curr}</TD>
                        <TD style="BORDER: purple 2px solid; PADDING: 5px">{left_count}</TD>
                        <TD style="BORDER: purple 2px solid; PADDING: 5px">{clear_count}</TD>
                        <TD style="BORDER: purple 2px solid; PADDING: 5px">{err_count}</TD>
                        <TD style="BORDER: purple 2px solid; PADDING: 5px">{note}</TD>
                    </TR>
                    """

                tbl_rows.append(row)

    summ = "\n".join(tbl_rows)

    return summ

def create(data: DataFrame, fields: list, rep_path: str, sht_name: str):
    """
    Creates processing report in .xlsx file format from processed accounting data.

    Parmas:
        data: A DataFrame object containing accounting data that will be written to an Excel report.
        fields: List of field names that will be used for ordering of report columns.
        rep_path: Path to the output Excel report file.
        sht_name: Nname of the report sheet to which accounting data will be placed.

    Returns: None.
    """
    # perform final data type conversion for the purpose of correct
    mask = data["Trading_Partner"].str.isnumeric()
    data.loc[mask, "Trading_Partner"] = pd.to_numeric(data.loc[mask, "Trading_Partner"]).astype("uint16")
    data["Document_Number"] = pd.to_numeric(data["Document_Number"]).astype("uint64")

    # sort data on selected fields
    sorted_data = data.sort_values([
        "Account",
        "Currency",
        "DC_Amount_ABS",
        "Posting_Number"
    ])

    # convert account numbers to int where possible
    sorted_data["Account"] = pd.to_numeric(sorted_data["Account"], errors = "ignore")

    ordered = sorted_data.reindex(columns = fields)

    date_columns = (
        "Document_Date",
        "Posting_Date",
        "Value_Date"
    )

    for col in date_columns:

        if col not in ordered.columns:
            continue

        ordered[col] = ordered[col].apply(
            lambda x: _to_excel_serial(x) if not pd.isna(x) else x
        )

    with ExcelWriter(rep_path, engine = "xlsxwriter") as wrtr:

        # replace underscores in comuln names with spaces for clarity and
        # write data to excel. Then replace spaces back with undescores
        # for better manipulation with data
        ordered.columns = _replace_col_char(ordered.columns, "_", " ")
        ordered.to_excel(wrtr, index = False, sheet_name = sht_name)
        ordered.columns = _replace_col_char(ordered.columns, " ", "_")

        # get report data sheet
        report = wrtr.book # pylint: disable=E1101
        data_sht = wrtr.sheets[sht_name]

        # define sheet custom data formats
        date_fmt = report.add_format({"num_format": "dd.mm.yyyy", "align": "center"})
        money_fmt = report.add_format({"num_format": "#,##0.00", "align": "center"})
        general_fmt = report.add_format({"align": "center"})
        header_fmt = report.add_format({
            "bg_color": "black", "font_color": "white", "bold": True, "align": "center"
        })

        # apply formats to cleared items column data
        for idx, col in enumerate(ordered.columns):

            col_width = _get_col_width(ordered[col], col)

            # select column data format
            if col == "DC_Amount":
                col_fmt = money_fmt
            elif col in date_columns:
                col_fmt = date_fmt
            else:
                col_fmt = general_fmt

            # apply new column params
            data_sht.set_column(idx, idx, col_width, col_fmt)

        # apply formats to headers
        first_col, last_col = ordered.columns[0], ordered.columns[-1]

        HEADER_ROW_IDX = 1

        data_sht.conditional_format(
            _col_to_rng(ordered, first_col, last_col, HEADER_ROW_IDX),
            {"type": "no_errors", "format": header_fmt}
        )

        # freeze data header row and set autofiler on all fields
        data_sht.freeze_panes(HEADER_ROW_IDX, 0)
        data_sht.autofilter(_col_to_rng(ordered, first_col, last_col, row = HEADER_ROW_IDX))

    return
