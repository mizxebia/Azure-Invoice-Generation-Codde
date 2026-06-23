import json
import logging
import os
import time
import io
import re
import traceback
import zipfile
from datetime import datetime
from pathlib import Path
 
import azure.functions as func
import pandas as pd
import requests
from msal import ConfidentialClientApplication
 
 
CONFIG_PATH = Path(__file__).with_name("config.json")
 
GRAPH_BASE = "https://graph.microsoft.com/v1.0"
CLOSING_PDF_FILE_COLUMN = "cr109_closingticketdetailspdf"
 
ENVIRONMENT_SETTINGS = {
    "DEV": {
        "user_email_key": "user_email_dev",
        "user_email": "akambotdev1@akam.com",
        "invoice_template_path": (
            "New Sales RPA/DEV/Excel Sheets/Closing Forms deposits.xlsx"
        ),
        "base_folder": "New Sales RPA/DEV/BotShareDrive/InProgress",
    },
    "UAT": {
        "user_email_key": "user_email_uat",
        "user_email": "akambotuat2@akam.com",
        "invoice_template_path": (
            "New Sales RPA/UAT/Excel Sheets/Closing Forms deposits.xlsx"
        ),
        "base_folder": "New Sales RPA/UAT/BotShareDrive/InProgress",
    },
    "PROD": {
        "user_email_key": "user_email_prod",
        "user_email": "akambotnewsalesclosure@akam.com",
        "invoice_template_path": (
            "New Sales RPA/PROD/Excel Sheets/Closing Forms deposits.xlsx"
        ),
        "base_folder": "New Sales RPA/PROD/BotShareDrive/InProgress",
    },
}
 

PAYABLE_MAP = {
    716070000: "Building",
    716070001: "AKAM",
    716070002: "Other",
}

TRANSACTION_TYPE_DEAL_MAP = {
    396620000: "All Cash",
    396620001: "Financing",
    396620002: "Transfer",
    396620003: "Trust Transfer",
}

DUE_AT_CLOSING_MAP = {
    396620000: "Adjournment Fee",
    396620001: "Admin Fee",
    396620002: "Air Conditioning Fee",
    396620003: "AKAM Processing Fee",
    396620004: "Appliance Fee",
    396620005: "Application Fee",
    396620006: "Arrears",
    396620007: "Assessment",
    396620008: "Assignment of Share",
    396620009: "Background Check",
    396620010: "Building Admin Fee",
    396620011: "Cable Charges",
    396620012: "Capital Assessment Fee",
    396620013: "Carpet Deposit",
    396620014: "Change of Occupancy",
    396620015: "Closing Fee (Non-Refundable)",
    396620016: "Contribution Fee (Non-Refundable)",
    396620017: "Contribution Reserves",
    396620018: "COOP Prospectus",
    396620019: "COOP Questionnaire",
    396620020: "Credit Report / Check",
    396620021: "Electric Fee",
    396620022: "Elevator Fee",
    396620023: "Energy Charge",
    396620024: "Escrow Maintenance",
    396620025: "Estate Review Fee",
    396620026: "Expediting Fee",
    396620027: "Flip Tax",
    396620028: "Guarantee Fee",
    396620029: "Inspection",
    396620030: "Legal Fee",
    396620031: "Lost Stock & Lease",
    396620032: "Maintenance Fees",
    396620033: "Major/Minor Alteration Fee",
    396620034: "Messenger",
    396620035: "Meter Fee",
    396620036: "Mortgage Questionnaire",
    396620037: "Move In/Out Deposit",
    396620038: "Move In/Out Fee",
    396620039: "Other",
    396620040: "Over-Time Fee",
    396620041: "Parking",
    396620042: "POA Fee",
    396620043: "Processing Fee",
    396620044: "Purchaser Fee (Transfer Fee)",
    396620045: "Real Estate Tax",
    396620046: "Recognition Agreement",
    396620047: "Repair Charge",
    396620048: "Resident Manager Contribution",
    396620049: "Security Deposit",
    396620050: "Service Fee",
    396620051: "Stock Transfer Fee",
    396620052: "Storage Unit",
    396620053: "Sublet Deposit",
    396620054: "Sublet Fee",
    396620055: "Transfer Fee",
    396620056: "Utilities",
    396620057: "Waiver Fee",
    396620058: "Working Capital",
}

 
# =====================================
# CONFIG
# =====================================
 
def load_config():
 
    with open(
        CONFIG_PATH,
        "r",
        encoding="utf-8"
    ) as config_file:
 
        return json.load(config_file)
 
 
def normalize_environment(env_value):

    env = (
        env_value
        or "DEV"
    ).strip().upper()

    if env not in ENVIRONMENT_SETTINGS:

        valid_envs = ", ".join(
            ENVIRONMENT_SETTINGS
        )

        raise ValueError(
            f"Invalid environment '{env}'. "
            f"Use one of: {valid_envs}."
        )

    return env


def get_runtime_settings(config, env):

    storage = config["storage"]

    auth_config = config["auth"]

    dv = config["dataverse"]

    env_settings = ENVIRONMENT_SETTINGS[
        env
    ]

    dataverse = (
        auth_config
        .get("dataverse_by_env", {})
        .get(env, {})
    )

    if env != "DEV" and not dataverse:

        raise ValueError(
            f"Missing Dataverse config for "
            f"environment '{env}'. Add "
            f"auth.dataverse_by_env.{env} "
            "to config.json."
        )

    tables = (
        dv.get("tables_by_env", {}).get(env)
        or dv.get(f"tables_{env.lower()}")
        or dv["tables"]
    )

    return {
        "env": env,
        "user_email": storage.get(
            env_settings["user_email_key"],
            env_settings["user_email"],
        ),
        "invoice_template_path": (
            env_settings[
                "invoice_template_path"
            ]
        ),
        "base_folder": (
            env_settings["base_folder"]
        ),
        "dataverse_tables": tables,
        "dataverse_url": dataverse.get(
            "dataverse_url",
            auth_config["dataverse_url"],
        ),
        "dataverse_scope": dataverse.get(
            "dataverse_scope",
            auth_config["dataverse_scope"],
        ),
    }


def get_request_value(req, name, default=None):

    value = req.params.get(name)

    if value is not None:
        return value

    try:
        body = req.get_json()
    except ValueError:
        body = {}

    return body.get(name, default)


config = load_config()
 
auth = config["auth"]
 
 
# =====================================
# LOGGING
# =====================================
 
logging.basicConfig(
 
    level=logging.INFO,
 
    format=(
        "%(asctime)s | "
        "%(levelname)s | "
        "%(message)s"
    ),
)
 
logger = logging.getLogger(__name__)
 
 
# =====================================
# AUTH
# =====================================
 
def get_access_token(scope):
 
    try:
 
        authority_url = (
            "https://login.microsoftonline.com/"
            f"{auth['tenant_id']}"
        )
 
        app = (
            ConfidentialClientApplication(
 
                auth["client_id"],
 
                authority=authority_url,
 
                client_credential=(
                    auth["client_secret"]
                ),
            )
        )
 
        token_response = (
            app.acquire_token_for_client(
                scopes=[scope]
            )
        )
 
        if (
            "access_token"
            not in token_response
        ):
 
            raise Exception(token_response)
 
        logger.info(
            "Access token acquired"
        )
 
        return token_response[
            "access_token"
        ]
 
    except Exception as error:
 
        logger.error(
            traceback.format_exc()
        )
 
        logger.error(str(error))
 
        return None
 
 
# =====================================
# DATAVERSE
# =====================================
 
def fetch_table(
    dataverse_url,
    table_name,
    token,
    ticket_column,
    ticket_value,
):
 
    headers = {
 
        "Authorization":
            f"Bearer {token}",
 
        "Accept":
            "application/json",
    }
 
    url = (
        f"{dataverse_url}"
        f"/api/data/v9.2/"
        f"{table_name}"
        f"?$filter="
        f"{ticket_column} "
        f"eq '{ticket_value}'"
    )
 
    all_records = []
 
    while url:
 
        response = requests.get(
            url,
            headers=headers
        )
 
        if response.status_code != 200:
 
            logger.error(response.text)
 
            raise Exception(response.text)
 
        data = response.json()
 
        all_records.extend(
            data.get("value", [])
        )
 
        url = data.get(
            "@odata.nextLink"
        )
 
    logger.info(
        f"{table_name} fetched: "
        f"{len(all_records)} records"
    )
 
    return all_records
 
 
def validate_required_dataverse_data(
    env,
    ticket_value,
    closing_data,
    invoice_data,
    closing_table,
    invoice_table,
):

    errors = []

    if not closing_data:

        errors.append(
            "No Closing Ticket Details "
            "records found in table "
            f"'{closing_table}' for ticket "
            f"'{ticket_value}'."
        )

    if not invoice_data:

        errors.append(
            "No Invoice Details records "
            f"found in table '{invoice_table}' "
            f"for ticket '{ticket_value}'."
        )

    if errors:

        raise ValueError(
            f"Dataverse data missing for "
            f"environment '{env}'. "
            + " ".join(errors)
        )


def get_choice_label(mapping, value):

    if value in (
        None,
        ""
    ) or pd.isna(value):

        return ""

    try:

        normalized_value = int(value)

    except (
        TypeError,
        ValueError
    ):

        normalized_value = value

    return mapping.get(
        normalized_value,
        ""
    )


def get_row_id(row, table_name):

    primary_key = (
        table_name[:-2] + "id"
        if table_name.endswith("es")
        else f"{table_name}id"
    )

    if primary_key in row:

        return row[primary_key]

    for key, value in row.items():

        if (
            key.endswith("id")
            and not key.startswith("_")
        ):

            return value

    raise ValueError(
        f"Could not determine Dataverse "
        f"row id for table '{table_name}'."
    )


# =====================================
# GRAPH HELPERS
# =====================================
 
def _headers(
    token,
    extra=None
):
 
    headers = {
 
        "Authorization":
            f"Bearer {token}",
 
        "Content-Type":
            "application/json",
    }
 
    if extra:
 
        headers.update(extra)
 
    return headers
 
 
def _ensure_folder(
    token,
    user_email,
    folder_path,
):
 
    check_url = (
        f"{GRAPH_BASE}/users/"
        f"{user_email}"
        f"/drive/root:/"
        f"{folder_path}"
    )
 
    response = requests.get(
        check_url,
        headers=_headers(token)
    )
 
    if response.status_code == 200:
        return
 
    parts = folder_path.rsplit("/", 1)
 
    parent_path = (
        parts[0]
        if len(parts) == 2
        else ""
    )
 
    folder_name = parts[-1]
 
    if parent_path:
 
        create_url = (
            f"{GRAPH_BASE}/users/"
            f"{user_email}"
            f"/drive/root:/"
            f"{parent_path}"
            f":/children"
        )
 
    else:
 
        create_url = (
            f"{GRAPH_BASE}/users/"
            f"{user_email}"
            f"/drive/root/children"
        )
 
    body = {
 
        "name": folder_name,
 
        "folder": {},
 
        "@microsoft.graph.conflictBehavior":
            "replace",
    }
 
    response = requests.post(
        create_url,
        headers=_headers(token),
        json=body,
    )
 
    if response.status_code not in [
        200,
        201,
    ]:
 
        raise Exception(
            f"Failed to create folder "
            f"{folder_path}: "
            f"{response.text}"
        )
 
    logger.info(
        f"Folder created: "
        f"{folder_path}"
    )
 
 
def _get_file_metadata(
    token,
    user_email,
    file_path,
):

    url = (
        f"{GRAPH_BASE}/users/"
        f"{user_email}"
        f"/drive/root:/"
        f"{file_path}:"
    )

    response = requests.get(
        url,
        headers=_headers(token)
    )

    if response.status_code == 200:
        return response.json()

    if response.status_code == 404:
        return None

    raise Exception(
        f"Error checking file "
        f"'{file_path}': "
        f"{response.text}"
    )


def _move_and_rename_file(
    token,
    user_email,
    source_path,
    destination_folder_path,
    new_name,
):

    folder_url = (
        f"{GRAPH_BASE}/users/"
        f"{user_email}"
        f"/drive/root:/"
        f"{destination_folder_path}"
    )

    folder_response = requests.get(
        folder_url,
        headers=_headers(token)
    )

    if folder_response.status_code != 200:

        raise Exception(
            f"Cannot find destination "
            f"folder "
            f"'{destination_folder_path}': "
            f"{folder_response.text}"
        )

    folder_id = folder_response.json()["id"]

    move_url = (
        f"{GRAPH_BASE}/users/"
        f"{user_email}"
        f"/drive/root:/"
        f"{source_path}"
    )

    body = {

        "parentReference": {
            "id": folder_id
        },

        "name": new_name,
    }

    response = requests.patch(
        move_url,
        headers=_headers(token),
        json=body,
    )

    if response.status_code != 200:

        raise Exception(
            f"Failed to move/rename file: "
            f"{response.text}"
        )

    logger.info(
        f"Archived file: {new_name}"
    )


def archive_active_file_if_exists(
    token,
    user_email,
    active_folder,
    inactive_folder,
    output_file_name,
):

    active_file_path = (
        f"{active_folder}/"
        f"{output_file_name}"
    )

    existing = _get_file_metadata(
        token,
        user_email,
        active_file_path,
    )

    if existing is None:

        logger.info(
            "No existing file in "
            "Active folder"
        )

        return

    now = datetime.now()

    date_str = now.strftime(
        "%Y-%m-%d"
    )

    datetime_str = now.strftime(
        "%Y-%m-%d_%H-%M-%S"
    )

    name_parts = (
        output_file_name.rsplit(".", 1)
    )

    if len(name_parts) == 2:

        archived_name = (
            f"{name_parts[0]}_"
            f"{datetime_str}."
            f"{name_parts[1]}"
        )

    else:

        archived_name = (
            f"{output_file_name}_"
            f"{datetime_str}"
        )

    date_folder = (
        f"{inactive_folder}/"
        f"{date_str}"
    )

    _ensure_folder(
        token,
        user_email,
        date_folder,
    )

    _move_and_rename_file(
        token,
        user_email,
        active_file_path,
        date_folder,
        archived_name,
    )

    logger.info(
        f"Archived existing file to "
        f"{date_folder}/{archived_name}"
    )


def _list_folder_children(
    token,
    user_email,
    folder_path,
):
    """Return list of file items inside a OneDrive folder."""

    url = (
        f"{GRAPH_BASE}/users/"
        f"{user_email}"
        f"/drive/root:/"
        f"{folder_path}"
        f":/children"
    )

    all_items = []

    while url:

        response = requests.get(
            url,
            headers=_headers(token)
        )

        if response.status_code == 404:
            return []

        if response.status_code != 200:

            raise Exception(
                f"Failed to list folder "
                f"'{folder_path}': "
                f"{response.text}"
            )

        data = response.json()

        all_items.extend(
            data.get("value", [])
        )

        url = data.get(
            "@odata.nextLink"
        )

    return all_items


def archive_active_files_by_prefix(
    token,
    user_email,
    active_folder,
    inactive_folder,
    file_prefix,
):
    """Archive every file in active_folder whose name starts with file_prefix."""

    items = _list_folder_children(
        token,
        user_email,
        active_folder,
    )

    matching = [
        item for item in items
        if "folder" not in item
        and item.get("name", "").startswith(
            file_prefix
        )
    ]

    if not matching:

        logger.info(
            f"No files matching prefix "
            f"'{file_prefix}' in Active "
            f"- skipping"
        )

        return

    now = datetime.now()

    date_str = now.strftime(
        "%Y-%m-%d"
    )

    datetime_str = now.strftime(
        "%Y-%m-%d_%H-%M-%S"
    )

    date_folder = (
        f"{inactive_folder}/"
        f"{date_str}"
    )

    _ensure_folder(
        token,
        user_email,
        date_folder,
    )

    for item in matching:

        file_name = item["name"]

        active_file_path = (
            f"{active_folder}/"
            f"{file_name}"
        )

        name_parts = (
            file_name.rsplit(".", 1)
        )

        if len(name_parts) == 2:

            archived_name = (
                f"{name_parts[0]}_"
                f"{datetime_str}."
                f"{name_parts[1]}"
            )

        else:

            archived_name = (
                f"{file_name}_"
                f"{datetime_str}"
            )

        _move_and_rename_file(
            token,
            user_email,
            active_file_path,
            date_folder,
            archived_name,
        )

        logger.info(
            f"Archived to: "
            f"{date_folder}/"
            f"{archived_name}"
        )


def setup_ticket_folders(
    token,
    user_email,
    base_folder,
    ticket_id,
):
 
    ticket_folder = (
        f"{base_folder}/"
        f"{ticket_id}"
    )
 
    active_folder = (
        f"{ticket_folder}/Active"
    )
 
    inactive_folder = (
        f"{ticket_folder}/Inactive"
    )
 
    _ensure_folder(
        token,
        user_email,
        ticket_folder,
    )
 
    _ensure_folder(
        token,
        user_email,
        active_folder,
    )
 
    _ensure_folder(
        token,
        user_email,
        inactive_folder,
    )
 
    return (
        active_folder,
        inactive_folder
    )
 


# =====================================
# COPY TEMPLATE / EXCEL POPULATE / PDF
# =====================================

def download_template_from_onedrive(token, user_email, template_path, local_path):
    url = (
        f"{GRAPH_BASE}/users/{user_email}"
        f"/drive/root:/{template_path}:/content"
    )

    response = requests.get(url, headers=_headers(token))

    if response.status_code != 200:
        logger.error(response.text)
        raise Exception(f"Failed to download template: {response.text}")

    with open(local_path, "wb") as file:
        file.write(response.content)

    logger.info(f"Template downloaded: {local_path}")


def _cell_addr_to_row_col(addr):
    """Convert cell address like 'D1' to (row, col) 1-indexed."""
    match = re.match(r"([A-Z]+)(\d+)", addr.upper())
    col_str, row_str = match.group(1), match.group(2)
    col = 0
    for ch in col_str:
        col = col * 26 + (ord(ch) - ord("A") + 1)
    return int(row_str), col


def _escape_xml(value):
    """Escape special XML characters in a cell value."""
    s = str(value) if value is not None else ""
    s = s.replace("&", "&amp;")
    s = s.replace("<", "&lt;")
    s = s.replace(">", "&gt;")
    s = s.replace('"', "&quot;")
    s = s.replace("'", "&apos;")
    return s


def _set_cell_in_sheet_xml(sheet_xml: str, addr: str, value, style=None) -> str:
    """
    Patch a single cell in worksheet XML.

    Handles both regular <c ...>...</c> and self-closing <c ... /> cells.
    Replaces the cell with an inlineStr value.
    If the cell doesn't exist, inserts it into the correct <row>.
    style: optional Excel style index string (s= attribute) to apply.
    """
    escaped = _escape_xml(value)
    row_num, _ = _cell_addr_to_row_col(addr)
    style_attr = f' s="{style}"' if style is not None else ""

    # ── 1a. Replace self-closing cell: <c r="ADDR" ... /> ────────────────────
    self_closing_pattern = re.compile(
        r'<c\s[^>]*\br="' + re.escape(addr) + r'"[^>]*/>'
    )

    def replace_self_closing(m):
        tag = m.group(0)[:-2]  # remove '/>'
        tag = re.sub(r'\s*\bt="[^"]*"', "", tag)
        if style is not None:
            tag = re.sub(r'\s*\bs="[^"]*"', "", tag)
            tag = re.sub(r'(<c\b)', rf'\1 t="inlineStr"{style_attr}', tag, count=1)
        else:
            tag = re.sub(r'(<c\b)', r'\1 t="inlineStr"', tag, count=1)
        return f'{tag}><is><t>{escaped}</t></is></c>'

    new_xml, count = self_closing_pattern.subn(replace_self_closing, sheet_xml)
    if count:
        return new_xml

    # ── 1b. Replace regular open/close cell: <c r="ADDR" ...>...</c> ─────────
    cell_pattern = re.compile(
        r'<c\s[^>]*\br="' + re.escape(addr) + r'"[^>]*>.*?</c>',
        re.DOTALL,
    )

    def replace_cell(m):
        open_tag_match = re.match(r'<c[^>]*>', m.group(0))
        if not open_tag_match:
            return m.group(0)
        open_tag = open_tag_match.group(0)
        open_tag = re.sub(r'\s*\bt="[^"]*"', "", open_tag)
        if style is not None:
            open_tag = re.sub(r'\s*\bs="[^"]*"', "", open_tag)
            open_tag = re.sub(r'(<c\b)', rf'\1 t="inlineStr"{style_attr}', open_tag, count=1)
        else:
            open_tag = re.sub(r'(<c\b)', r'\1 t="inlineStr"', open_tag, count=1)
        return f'{open_tag}<is><t>{escaped}</t></is></c>'

    new_xml, count = cell_pattern.subn(replace_cell, sheet_xml)
    if count:
        return new_xml

    # ── 2. Cell not found — insert into existing row ─────────────────────────
    row_pattern = re.compile(
        r'(<row\b[^>]*\br="' + str(row_num) + r'"[^>]*>)(.*?)(</row>)',
        re.DOTALL,
    )

    def insert_into_row(m):
        new_cell = f'<c r="{addr}" t="inlineStr"{style_attr}><is><t>{escaped}</t></is></c>'
        return m.group(1) + m.group(2) + new_cell + m.group(3)

    new_xml, count = row_pattern.subn(insert_into_row, sheet_xml)
    if count:
        return new_xml

    # ── 3. Row not found — insert a new row before </sheetData> ──────────────
    new_row = (
        f'<row r="{row_num}">'
        f'<c r="{addr}" t="inlineStr"{style_attr}><is><t>{escaped}</t></is></c>'
        f'</row>'
    )
    return sheet_xml.replace("</sheetData>", new_row + "</sheetData>", 1)


def populate_excel_template_local(
    template_path,
    output_path,
    closing_ticket_df,
    invoice_df,
):
    """
    Patch the Excel template by editing the worksheet XML directly inside the
    xlsx zip — leaving all other parts (images, drawings, printer settings,
    custom XML, etc.) completely intact so OneDrive can convert it to PDF.
    """
    row = closing_ticket_df.iloc[0]

    purchase_price = row.get("cr109_saleprice", "")
    closing_date = row.get("cr7de_closingdate", "")
    seller_tcode = row.get("cr7de_sellertcode", "")
    property_address = row.get("cr7de_buildingaddress", "")
    unit = row.get("cr7de_unitnumber", "")
    seller1_name = row.get("cr7de_sellername", "")
    deal = get_choice_label(
        TRANSACTION_TYPE_DEAL_MAP,
        row.get("cr109_transactiontypedeal", ""),
    )
    buyer1_name = row.get("cr7de_buyername", "")
    shares = row.get("cr109_shares", "")
    closing_agent = row.get("cr7de_closingagentname", "")
    closing_agent_phone = row.get("cr7de_closingagentphone", "")
    closing_agent_email = row.get("cr7de_closingagentemail", "")
    closing_agent_title = row.get("cr7de_titlerole", "")
    notes = row.get("cr7de_notes", "")
    building_name = row.get("cr109_legalname", "")
    current_date = datetime.now().strftime("%m/%d/%Y")

    try:
        if closing_date:
            closing_date = pd.to_datetime(closing_date).strftime("%m/%d/%Y")
    except Exception:
        pass

    cell_updates = {
        "D1": current_date,
        "D2": purchase_price,
        "D3": deal,
        "D4": shares,
        "D5": closing_date,
        "D6": seller_tcode,
        "D7": property_address,
        "D8": unit,
        "B11": row.get("cr109_locationofclosing", ""),
        "C13": seller1_name,
        "C43": buyer1_name,
        "B86": closing_agent,
        "B87": closing_agent_email,
        "B88": closing_agent_phone,
        "B89": closing_agent_title,
        "B91": notes,
    }

    if "cr7de_paidby" not in invoice_df.columns:
        raise Exception(
            f"Column cr7de_paidby not found. "
            f"Available columns: {invoice_df.columns.tolist()}"
        )

    # Style 8 = wrapText + center + border (matches D column data rows)
    WRAP_STYLE = "8"

    seller_df = invoice_df[invoice_df["cr7de_paidby"] == 716070000]
    for idx, (_, inv_row) in enumerate(seller_df.iterrows()):
        r = 15 + idx
        payable_to = get_choice_label(PAYABLE_MAP, inv_row.get("cr7de_payableto", ""))
        if payable_to == "Building" and building_name:
            payable_to = building_name
        elif payable_to == "AKAM":
            payable_to = "AKAM Associates, Inc"
        elif payable_to == "Other":
            payable_to = inv_row.get("cr109_otherpayableto", "") or "Other"
        cell_updates[f"A{r}"] = inv_row.get("cr7de_chequenumber", "")
        cell_updates[f"B{r}"] = get_choice_label(DUE_AT_CLOSING_MAP, inv_row.get("cr109_dueatclosing", ""))
        cell_updates[f"C{r}"] = inv_row.get("cr7de_amount", "")
        cell_updates[f"D{r}"] = payable_to

    buyer_df = invoice_df[invoice_df["cr7de_paidby"] == 716070001]
    for idx, (_, inv_row) in enumerate(buyer_df.iterrows()):
        r = 45 + idx
        payable_to = get_choice_label(PAYABLE_MAP, inv_row.get("cr7de_payableto", ""))
        if payable_to == "Building" and building_name:
            payable_to = building_name
        elif payable_to == "AKAM":
            payable_to = "AKAM Associates, Inc"
        elif payable_to == "Other":
            payable_to = inv_row.get("cr109_otherpayableto", "") or "Other"
        cell_updates[f"A{r}"] = inv_row.get("cr7de_chequenumber", "")
        cell_updates[f"B{r}"] = get_choice_label(DUE_AT_CLOSING_MAP, inv_row.get("cr109_dueatclosing", ""))
        cell_updates[f"C{r}"] = inv_row.get("cr7de_amount", "")
        cell_updates[f"D{r}"] = payable_to

    TARGET_SHEET = "Closing Check Transmittal Form"
    SHEET_ENTRY = "xl/worksheets/sheet1.xml"
    WORKBOOK_ENTRY = "xl/workbook.xml"

    with zipfile.ZipFile(template_path, "r") as zin:
        entries = zin.namelist()
        sheet_xml = zin.read(SHEET_ENTRY).decode("utf-8")
        workbook_xml = zin.read(WORKBOOK_ENTRY).decode("utf-8")

        # Patch cell values — E column cells get wrap-text style (s=8)
        for addr, value in cell_updates.items():
            if addr.startswith("E"):
                sheet_xml = _set_cell_in_sheet_xml(sheet_xml, addr, value, style=WRAP_STYLE)
            else:
                sheet_xml = _set_cell_in_sheet_xml(sheet_xml, addr, value)

        # Hide all sheets except the target so PDF only contains one sheet
        def hide_non_target_sheets(wb_xml):
            def patch_sheet(m):
                name = m.group(1)
                rest = m.group(2)
                if name == TARGET_SHEET:
                    rest = re.sub(r'\s*state="[^"]*"', "", rest)
                    return f'<sheet name="{name}"{rest}/>'
                else:
                    rest = re.sub(r'\s*state="[^"]*"', "", rest)
                    return f'<sheet name="{name}"{rest} state="hidden"/>'
            return re.sub(r'<sheet name="([^"]+)"([^/]*)/>', patch_sheet, wb_xml)

        workbook_xml = hide_non_target_sheets(workbook_xml)

        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
            for entry in entries:
                if entry == SHEET_ENTRY:
                    zout.writestr(entry, sheet_xml.encode("utf-8"))
                elif entry == WORKBOOK_ENTRY:
                    zout.writestr(entry, workbook_xml.encode("utf-8"))
                else:
                    zout.writestr(entry, zin.read(entry))

    with open(output_path, "wb") as f:
        f.write(buf.getvalue())

    logger.info(f"Excel populated and saved: {output_path}")


def upload_excel_to_onedrive(token, user_email, folder_path, local_file_path, onedrive_file_name):
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": (
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    }

    url = (
        f"{GRAPH_BASE}/users/{user_email}"
        f"/drive/root:/{folder_path}/{onedrive_file_name}:/content"
    )

    with open(local_file_path, "rb") as file:
        file_content = file.read()

    response = requests.put(url, headers=headers, data=file_content)

    if response.status_code not in [200, 201]:
        logger.error(response.text)
        raise Exception(response.text)

    logger.info(f"Excel uploaded: {folder_path}/{onedrive_file_name}")


def upload_pdf_to_onedrive(token, user_email, folder_path, local_file_path, onedrive_file_name):
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/pdf",
    }

    url = (
        f"{GRAPH_BASE}/users/{user_email}"
        f"/drive/root:/{folder_path}/{onedrive_file_name}:/content"
    )

    with open(local_file_path, "rb") as file:
        file_content = file.read()

    response = requests.put(url, headers=headers, data=file_content)

    if response.status_code not in [200, 201]:
        logger.error(response.text)
        raise Exception(response.text)

    logger.info(f"PDF uploaded: {folder_path}/{onedrive_file_name}")


def convert_excel_to_pdf_onedrive(token, user_email, file_path, pdf_output_path):
    """Convert an uploaded Excel file to PDF via OneDrive Graph API."""
    url = (
        f"{GRAPH_BASE}/users/{user_email}"
        f"/drive/root:/{file_path}:/content?format=pdf"
    )

    response = requests.get(url, headers=_headers(token))

    logger.info(f"PDF conversion response: {response.status_code}")

    if response.status_code != 200:
        logger.error(response.text)
        raise Exception("PDF conversion failed")

    with open(pdf_output_path, "wb") as file:
        file.write(response.content)

    logger.info("PDF generated successfully")


 
def upload_file_to_dataverse_file_column(
    dataverse_url,
    table_name,
    row_id,
    token,
    file_column,
    local_file_path,
    file_name,
):

    url = (
        f"{dataverse_url}/api/data/v9.2/"
        f"{table_name}({row_id})/"
        f"{file_column}"
    )

    headers = {

        "Authorization":
            f"Bearer {token}",

        "OData-MaxVersion":
            "4.0",

        "OData-Version":
            "4.0",

        "If-None-Match":
            "null",

        "Accept":
            "application/json",

        "Content-Type":
            "application/octet-stream",

        "x-ms-file-name":
            file_name,
    }

    with open(
        local_file_path,
        "rb"
    ) as file:

        response = requests.patch(
            url,
            headers=headers,
            data=file.read(),
        )

    if response.status_code != 204:

        logger.error(response.text)

        raise Exception(
            f"Failed to upload PDF to "
            f"Dataverse file column "
            f"'{file_column}': "
            f"{response.text}"
        )

    logger.info(
        f"Uploaded PDF to Dataverse "
        f"column: {table_name}."
        f"{file_column}"
    )


# =====================================
# MAIN
# =====================================
def main(req: func.HttpRequest) -> func.HttpResponse:
 
    try:
 
        env = normalize_environment(
            get_request_value(
                req,
                "env",
                "DEV"
            )
        )
 
        runtime = get_runtime_settings(
            config,
            env,
        )
 
        ticket_value = get_request_value(
            req,
            "id",
        )
 
        if not ticket_value:
 
            ticket_value = config[
                "dataverse"
            ]["filter"]["ticket_id"]
 
        logger.info(
            f"Environment: {env}"
        )
 
        logger.info(
            f"Ticket Value: "
            f"{ticket_value}"
        )
 
        dv_token = get_access_token(
 
            runtime[
                "dataverse_scope"
            ]
        )
 
        graph_token = get_access_token(
 
            config["auth"][
                "graph_scope"
            ]
        )
 
        dv = config["dataverse"]
 
        dv_tables = runtime[
            "dataverse_tables"
        ]
 
        dataverse_url = runtime[
            "dataverse_url"
        ]
 
        ticket_column = (
            dv["columns"]["ticket_id"]
        )
 
        template_path = (
            runtime[
                "invoice_template_path"
            ]
        )
 
        user_email = runtime[
            "user_email"
        ]
 
        base_folder = runtime[
            "base_folder"
        ]
 
        output_excel_name = (
            f"{ticket_value}"
            f"_Closing_Form.xlsx"
        )

        pdf_output_name = (
            f"{ticket_value}"
            f"_Closing_Form.pdf"
        )

        local_template_path = os.path.join(
            "/tmp",
            f"template_{ticket_value}.xlsx",
        )

        local_output_path = os.path.join(
            "/tmp",
            f"output_{ticket_value}.xlsx",
        )

        local_pdf_path = os.path.join(
            "/tmp",
            pdf_output_name,
        )
 
        logger.info(
            f"Dataverse URL: "
            f"{dataverse_url}"
        )
 
        closing_table = dv_tables[
            "closing_ticket_details"
        ]
 
        invoice_table = dv_tables[
            "invoice_details"
        ]
 
        closing_data = fetch_table(
 
            dataverse_url,
 
            closing_table,
 
            dv_token,
 
            ticket_column,
 
            ticket_value,
        )
 
        invoice_data = fetch_table(
 
            dataverse_url,
 
            invoice_table,
 
            dv_token,
 
            ticket_column,
 
            ticket_value,
        )
 
        validate_required_dataverse_data(
 
            env,
 
            ticket_value,
 
            closing_data,
 
            invoice_data,
 
            closing_table,
 
            invoice_table,
        )
 
        df_closing = pd.DataFrame(
            closing_data
        )
 
        df_invoice = pd.DataFrame(
            invoice_data
        )

        closing_row_id = get_row_id(
            closing_data[0],
            closing_table,
        )
 
        logger.info(
            f"Closing DF Shape: "
            f"{df_closing.shape}"
        )
 
        logger.info(
            f"Invoice DF Shape: "
            f"{df_invoice.shape}"
        )
 
        (
            active_folder,
            inactive_folder
 
        ) = setup_ticket_folders(
 
            graph_token,
 
            user_email,
 
            base_folder,
 
            ticket_value,
        )
 
        archive_active_file_if_exists(

            graph_token,

            user_email,

            active_folder,

            inactive_folder,

            output_excel_name,
        )

        archive_active_files_by_prefix(

            graph_token,

            user_email,

            active_folder,

            inactive_folder,

            f"{ticket_value}_Condo_Invoice",
        )

        # Download template locally, populate with openpyxl, upload back
        download_template_from_onedrive(
            graph_token,
            user_email,
            template_path,
            local_template_path,
        )

        populate_excel_template_local(
            local_template_path,
            local_output_path,
            df_closing,
            df_invoice,
        )

        upload_excel_to_onedrive(
            graph_token,
            user_email,
            active_folder,
            local_output_path,
            output_excel_name,
        )

        # Wait for OneDrive to process the uploaded file before PDF conversion
        logger.info("Waiting for OneDrive to process the file...")
        time.sleep(5)

        onedrive_excel_path = (
            f"{active_folder}/{output_excel_name}"
        )

        convert_excel_to_pdf_onedrive(
            graph_token,
            user_email,
            onedrive_excel_path,
            local_pdf_path,
        )

        upload_pdf_to_onedrive(
            graph_token,
            user_email,
            active_folder,
            local_pdf_path,
            pdf_output_name,
        )

        upload_file_to_dataverse_file_column(

            dataverse_url,

            closing_table,

            closing_row_id,

            dv_token,

            CLOSING_PDF_FILE_COLUMN,

            local_pdf_path,

            pdf_output_name,
        )
 
        for path in [local_template_path, local_output_path, local_pdf_path]:
            if os.path.exists(path):
                os.remove(path)
 
        return func.HttpResponse(
 
            body=json.dumps({
 
                "status":
                    "SUCCESS",
 
                "ticket_id":
                    ticket_value,

                "environment":
                    env,

                "output_file":
                    output_excel_name,

                "pdf_file":
                    pdf_output_name,
            }),
 
            status_code=200,
 
            mimetype="application/json",
        )
 
    except Exception as e:
 
        logs = (
            f"\nError :{str(e)}"
        )
 
        tb = traceback.format_exc()
 
        logger.error(logs)
 
        logger.error(tb)
 
        return func.HttpResponse(
 
            body=json.dumps({
 
                "status":
                    "ERROR",
 
                "message":
                    str(e),
 
                "traceback":
                    tb
            }),
 
            status_code=500,
 
            mimetype="application/json",
        )
   
 
 
if __name__ == "__main__":
    main("Test")
