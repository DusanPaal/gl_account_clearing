"""
The 'biaProcessor.py' module performs the following operations on accounting data:
    - conversion,
    - item matching,
    - generating clearing input.

Version history:
1.0.20210819 - Initial version.
1.0.20211011 - Data field 'Assigment' will be right-stripped to preserve any leading whitespaces.
               See comments in the data striping code section in the affected procedure 'convert_fbl3n_data()'
1.0.20220615 - Refactored and simplified code, added/updated docstrings.
"""

from io import StringIO
from logging import getLogger
import re
import pandas as pd
from pandas import DataFrame, Series
from scripts.biaF03 import Record

_logger = getLogger("master")

# clearing criteria keys used in accounting config
_criterias = {
    "A": "Assignment",
    "C": "Cummulative_Sum",
    "D": "Document_Number",
    "O": "Oldest_Assignment",
    "P": "Trading_Partner",
    "R": "Reference",
    "T": "Text",
    "X": "Deal_Number"
}

def _parse_amounts(vals: Series) -> Series:
    """
    Parses string amounts stored
    in the SAP amount format into
    floats.
    """

    parsed = vals.copy()
    parsed = parsed.str.replace(".", "", regex = False).str.replace(",", ".", regex = False)
    parsed = parsed.mask(parsed.str.endswith("-"), "-" + parsed.str.replace("-", ""))
    parsed = pd.to_numeric(parsed)

    return parsed

def _parse_dates(vals: Series) -> Series:
    """
    Parses string dates stored
    in the SAP format into
    'date' object type.
    """

    parsed = pd.to_datetime(vals, dayfirst = True, errors = "coerce")

    return parsed.dt.date

def convert_fbl3n_data(file_path: str, cocd: str) -> DataFrame:
    """
    Converts plain FBL3N text data into a panel dataset.

    Params:
        file_path: Path to the text file with FBL5N data to parse.
        cocd: Company code of the data to convert.

    Returns: A DataFrame object containing accounting data
             if data conversion succeeds, None if it fails.
    """

    _logger.info(" Converting FBL3N data ...")

    with open(file_path, 'r', encoding = "utf-8") as file:
        txt = file.read()

    # get all data lines containing accounting items
    matches = re.findall(r"^\|\s+\w{3}\s+\|\w+\s*\|.*\|$", txt, re.M)
    replaced = list(map(lambda x: x[1:-1], matches))
    preproc = "\n".join(replaced)

    # define data header names
    header = [
        "Currency",
        "Account",
        "DC_Amount",
        "Document_Number",
        "Document_Type",
        "Document_Date",
        "Posting_Date",
        "Assignment",
        "Reference",
        "Trading_Partner",
        "Text",
        "Value_Date"
    ]

    parsed = pd.read_csv(
        StringIO(preproc),
        sep = "|",
        names = header,
        dtype = "string",
        keep_default_na = False
    )

    assert not parsed.empty, "Parsing failed!"

    # trim string data except for the 'Assignment' field where only tailing whitespaces
    # may be stripped. Some items appear on accounts with entered leading whitespace chars,
    # striping the whole string will result in failure to load items from GL accounts when
    # assignment selection criteria is used in F-30
    for col in parsed.columns:
        if col == "Assignment":
            parsed[col] = parsed[col].str.rstrip()
        else:
            parsed[col] = parsed[col].str.strip()

    parsed["DC_Amount"] = _parse_amounts(parsed["DC_Amount"])
    parsed["Document_Date"] = _parse_dates(parsed["Document_Date"])
    parsed["Posting_Date"] = _parse_dates(parsed["Posting_Date"])
    parsed["Value_Date"] = _parse_dates(parsed["Value_Date"])

    # add new fields to data frame
    parsed = parsed.assign(
        DC_Amount_ABS = parsed["DC_Amount"].abs(),
        Deal_Number = pd.NA,
        Posting_Number = pd.NA,
        Match = False,
        Processed = False,
        Excluded = False,
        Message = ""
    )

    # extract deal number from text where applicable
    if cocd == "499L":
        parsed["Deal_Number"] = parsed["Text"].str.extract(r"(\d{13})$")
    elif cocd == "0073":
        parsed["Deal_Number"] = parsed["Text"].str.extract(r";(\d+)$")
        mask = parsed["Deal_Number"].notna()
        parsed.loc[mask, "Deal_Number"] = parsed.loc[mask, "Deal_Number"].astype("uint64")

    converted = parsed.copy()
    converted["Currency"] = converted["Currency"].astype("category")
    converted["Account"] = converted["Account"].astype("category") # some accs contain letters
    converted["Reference"] = converted["Reference"].astype("category")
    converted["Document_Type"] = converted["Document_Type"].astype("category")
    converted["Trading_Partner"] = converted["Trading_Partner"].astype("object")

    return converted

def _match_oldest_assign(data: DataFrame) -> DataFrame:
    """
    Matches items for a given GL account considering
    oldest assignment as the matching criteria.

    Params:
    -------
    data: Accounting data containing records (items) to match.

    Returns:
    --------
    A copy of the original DataFrame object with 'True' values
    in the 'Match' field indicating matched records, if any.
    """

    if data.empty:
        raise ValueError("Argument 'data' contains no records!")

    copied = data.copy()

    # match single +/- amounts
    grouped = copied.assign(
        Sum = copied.groupby(
            ["Currency", "DC_Amount_ABS", "Assignment"]
        )["DC_Amount"].transform("sum")
    )

    matched = (grouped["Sum"].round(2) == 0)
    grouped.loc[:, "Match"] = matched
    grouped.loc[:, "Processed"] = matched
    grouped.drop("Sum", axis = 1, inplace = True)

    # in the next step, match the remaining multi +/- amounts based on the oldest assignment
    currs = grouped[~grouped["Processed"]]["Currency"].unique()

    for curr in currs:

        curr_items = grouped[~grouped["Processed"] & (grouped["Currency"] == curr)]
        dupl_amnts = curr_items[["DC_Amount_ABS", "Assignment"]][~curr_items["Processed"] & curr_items["DC_Amount_ABS"].duplicated()]
        uniq_amnts = dupl_amnts.drop_duplicates(keep = "first")

        for amnt, assign in uniq_amnts.itertuples(index = False):

            posit_mask = ((curr_items["DC_Amount"] == amnt) & (curr_items["Assignment"] == assign))
            negat_mask = ((curr_items["DC_Amount"] == -amnt) & (curr_items["Assignment"] == assign))

            posit = curr_items[posit_mask]
            negat = curr_items[negat_mask]

            clr_count = min(posit.shape[0], negat.shape[0])

            # only negative or positive values exist, continue processing
            if clr_count == 0:
                grouped.loc[posit.index, "Processed"] = True
                grouped.loc[negat.index, "Processed"] = True
                continue

            oldest_posit = pd.to_datetime(curr_items.loc[posit_mask, "Document_Date"]).nsmallest(clr_count, keep="first")
            oldest_negat = pd.to_datetime(curr_items.loc[negat_mask, "Document_Date"]).nsmallest(clr_count, keep="first")

            grouped.loc[oldest_posit.index, "Match"] = True
            grouped.loc[oldest_posit.index, "Processed"] = True

            grouped.loc[oldest_negat.index, "Match"] = True
            grouped.loc[oldest_negat.index, "Processed"] = True

    return grouped

def _match_cumm_sum(data: DataFrame) -> DataFrame:
    """
    Matches items only where the cummulative
    sum of DC amounts equals zero.

    Params:
    -------
    acc_data: A GL account data containing records (items) to match.

    Returns:
    --------
    A copy of the original DataFrame object with 'True' values
    in the 'Match' field indicating matched records, if any.
    """

    if data.empty:
        raise ValueError("Argument 'data' contains no records!")

    result = data.copy()
    result.sort_values("Value_Date", inplace = True)
    result["Cummulative_Sum"] = result["DC_Amount"].cumsum().round(2)
    qry = result.query("Cummulative_Sum == 0.0")

    if qry.shape[0] == 0:
        return result

    last_zero_idx = qry.tail(1).index[0]
    result.loc[:last_zero_idx, "Match"] = True

    result.drop("Cummulative_Sum", axis = 1, inplace = True)

    return result

def _match_deal_number(data: DataFrame, cocd: str) -> DataFrame:
    """
    Matches items only where the sum of DC amounts for
    given deal numbers equals zero.

    Params:
    -------
    data: A GL account data containing records (items) to match.
    cocd: Company code of the accounting data.

    Returns:
    --------
    A copy of the original DataFrame object with 'True' values
    in the 'Match' field indicating matched records, if any.
    """

    if data.empty:
        raise ValueError("Argument 'data' contains no records!")

    result = data.copy()
    result = result.assign(
        Sum = result[result["Deal_Number"].notna()].groupby([
                "Currency", "Deal_Number"
            ])["DC_Amount"].transform("sum")
    )

    matched = (result["Sum"].round(2) == 0)
    result.loc[matched, "Match"] = True
    result.loc[matched, "Processed"] = True
    result.drop("Sum", axis = 1, inplace = True)

    if cocd == "499L":
        qry = result.query("~Deal_Number.str.startswith('60') and Match == True")
        result.loc[qry.index, "Message"] = "Excluded from clearing based on deal number criteria."
        result.loc[qry.index, "Excluded"] = True

    return result

def _match_amounts(acc_subset: DataFrame, crits: list, tr_ptr: str = []) -> DataFrame:
    """
    Matches items using account-specific matching criteria.
    If a trading partner if provided, only items that are
    associated with that trading partner will be considered
    for matching.

    Params:
    -------
    acc_subset: A GL account data containing records (items) to match.
    crits: Account-specific matching criteria.
    tr_ptr: Trading partner.

    Returns:
    -------
    A copy of the original DataFrame object with 'True' values
    in the 'Match' field indicating matched records, if any.
    """

    if acc_subset.empty:
        raise ValueError("Argument 'acc_subset' contains no records!")

    result = acc_subset.copy()

    if len(tr_ptr) == 0:
        subset = acc_subset.copy()
    else:
        subset = acc_subset[acc_subset["Trading_Partner"].isin(tr_ptr)].copy()

    # if the sum of currency amounts equals 0, match them
    subset = subset.assign(
        Sum = subset.groupby(["Currency"])["DC_Amount"].transform("sum")
    )

    matched = (subset["Sum"].round(2) == 0)
    subset.loc[matched, "Match"] = True
    subset.loc[matched, "Processed"] = True

    # entire acc done
    if subset.loc[subset.index, "Processed"].all():
        return subset

    # match +/- amounts where negative amount count equals to positive amount count
    subset["Sum"] = subset[~subset["Processed"]].groupby(["Currency", "DC_Amount_ABS"])["DC_Amount"].transform("sum")
    matched = (subset["Sum"].round(2) == 0)
    subset.loc[matched, "Match"] = True
    subset.loc[matched, "Processed"] = True

    # match the remaining multi +/- amounts using additional criteria
    for crit in crits:
        subset["Sum"] = subset[~subset["Processed"]].groupby(["Currency", crit])["DC_Amount"].transform("sum")
        matched = (subset["Sum"].round(2) == 0)
        subset.loc[matched, "Match"] = True
        subset.loc[matched, "Processed"] = True

    subset.drop("Sum", axis = 1, inplace = True)
    result.loc[subset.index, :] = subset

    return result

def _get_trading_partners(criteria: list) -> list:
    """
    Returns a list of trading prtners
    found in matching criteria.
    """

    partners = []

    for crit in criteria:
        if "P" in crit:
            partners = crit.split("_")[1:]

    return partners

def find_matches(data: DataFrame, rules: dict, cocd: str) -> DataFrame:
    """
    A wrapper for higly specialized procedures dedicated to account
    items matching based on specific criteria.

    Params:
    -------
    data: Data on which matching is performed.
    rules: Matching criteria specified for a given account.
    cocd: Company code of the data to process.

    Returns:
    --------
    A DataFrame object containing matched items, if any matches were
    found, otherwise the original DataFrame object.
    """

    _logger.info(" Searching data for items to clear ...")

    if data.empty:
        raise ValueError("Argument 'data' contains no records!")

    sorted_data = data.sort_values(["Account", "Currency", "DC_Amount_ABS", "Posting_Date"])
    gl_accs = rules[cocd]["accounts"]
    output = []

    # exclude any associated accounts, as these get processsed separately
    for acc in sorted_data["Account"].unique():

        if not gl_accs[acc]["active"]:
            continue

        acc_crits = gl_accs[acc]["criteria"]
        criteria = [_criterias[crit.split("_")[0]] for crit in acc_crits]
        acc_subset = sorted_data[sorted_data["Account"] == acc]

        if "Oldest_Assignment" in criteria:
            matched = _match_oldest_assign(acc_subset)
        elif "Cummulative_Sum" in criteria:
            matched = _match_cumm_sum(acc_subset)
        elif "Deal_Number" in criteria:
            matched = _match_deal_number(acc_subset, cocd)
        elif "Trading_Partner" in criteria:
            partners = _get_trading_partners(acc_crits)
            matched = _match_amounts(acc_subset, criteria, partners)
        else:
            matched = _match_amounts(acc_subset, criteria)

        output.append(matched)

    result = pd.concat(output)

    return result

def generate_clearing_input(data: DataFrame, cmp_cd: str) -> dict:
    """
    Identifies matched items per account and currency, company codes and clearing criteria,
    then creates input records from the corresponding accounting details for subsequent F-03
    clearing. Associated accounts are excluded from processing.

    Params:
        data: Accounting data containing matched items.
        cmp_cd: Company code of the data to process.

    Returns: A tuple of the recordlist with input records and the total number of matched items.
    """

    output = {}
    tot_matches = 0

    if not data["Match"].any():
        return output

    for acc in data["Account"].unique():

        subset = data[data["Account"] == acc]

        # check if here are any DC amounts matched
        if not subset["Match"].any():
            continue

        items_to_clear = subset[~subset["Excluded"] & subset["Match"]]
        clear_count = items_to_clear.shape[0]

        # no matches selected, go to next account
        if clear_count == 0:
            continue

        tot_matches += clear_count
        output.update({acc: dict()})

        DATE_FORMAT = "%d.%m.%Y"

        for curr in items_to_clear["Currency"].unique():

            # all matched items should have assignments
            # that are a non-empty strings, otherwise skip
            curr_mask = (items_to_clear["Currency"] == curr)
            assigns = items_to_clear.loc[curr_mask, "Assignment"].unique()
            refs = items_to_clear.loc[curr_mask, "Reference"].unique()
            docnums = items_to_clear.loc[curr_mask, "Document_Number"].unique()

            # performance optimizatio for Finland
            if cmp_cd == "0073" and acc == "24182000":
                assigns = None
                refs = None

            rec = Record(
                DC_Amounts = items_to_clear["DC_Amount"][curr_mask].tolist(),
                Document_Numbers = items_to_clear["Document_Number"][curr_mask].tolist(),
                Document_Types = items_to_clear["Document_Type"][curr_mask].tolist(),
                Document_Dates = list(map(lambda x: x.strftime(DATE_FORMAT), items_to_clear["Document_Date"][curr_mask])),
                Posting_Dates = list(map(lambda x: x.strftime(DATE_FORMAT), items_to_clear["Posting_Date"][curr_mask])),
                Unique_Assignments = None if assigns is None or "" in assigns else list(assigns),
                Unique_References = None if refs is None or "" in refs else list(refs),
                Unique_Document_Numbers = list(docnums),
                All_Assignments = items_to_clear["Assignment"][curr_mask].tolist(),
                Texts = items_to_clear["Text"][curr_mask].tolist(),
                Trading_Partners = items_to_clear["Trading_Partner"][curr_mask].tolist(),
                Indexes = items_to_clear[curr_mask].index
            )

            output[acc].update({curr: rec})

    return output
