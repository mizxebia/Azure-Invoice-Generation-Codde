import json
import logging
import os
import argparse
import time
import traceback
from datetime import datetime
from pathlib import Path

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


def load_config():
    with open(CONFIG_PATH, "r", encoding="utf-8") as config_file:
        return json.load(config_file)


def parse_args():
    parser = argparse.ArgumentParser(
        description="Generate closing form invoice files."
    )
    parser.add_argument(
        "environment",
        nargs="?",
        help="Target environment: DEV, UAT, or PROD.",
    )
    parser.add_argument(
        "--env",
        dest="env",
        help="Target environment: DEV, UAT, or PROD.",
    )
    return parser.parse_args()


def normalize_environment(env_value):
    env = (env_value or "DEV").strip().upper()

    if env not in ENVIRONMENT_SETTINGS:
        valid_envs = ", ".join(ENVIRONMENT_SETTINGS)
        raise ValueError(f"Invalid environment '{env}'. Use one of: {valid_envs}.")

    return env


def get_runtime_settings(config, env):
    storage = config["storage"]
    auth_config = config["auth"]
    dv = config["dataverse"]
    env_settings = ENVIRONMENT_SETTINGS[env]
    dataverse = auth_config.get("dataverse_by_env", {}).get(env, {})

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
        "invoice_template_path": env_settings["invoice_template_path"],
        "base_folder": env_settings["base_folder"],
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


config = load_config()
auth = config["auth"]

logging.basicConfig(
    filename=config["logging"]["log_file"],
    level=getattr(logging, config["logging"]["log_level"]),
    format="%(asctime)s | %(levelname)s | %(message)s",
)
logger = logging.getLogger(__name__)


def get_access_token(scope):
    try:
        authority_url = f"https://login.microsoftonline.com/{auth['tenant_id']}"

        app = ConfidentialClientApplication(
            auth["client_id"],
            authority=authority_url,
            client_credential=auth["client_secret"],
        )

        token_response = app.acquire_token_for_client(scopes=[scope])

        if "access_token" not in token_response:
            raise Exception(token_response)

        return token_response["access_token"]

    except Exception as error:
        logger.error(traceback.format_exc())
        print(error)
        return None


def fetch_table(table_name, token, ticket_column, ticket_value):
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
    }

    url = (
        f"{auth['dataverse_url']}/api/data/v9.2/{table_name}"
        f"?$filter={ticket_column} eq '{ticket_value}'"
    )

    all_records = []

    while url:
        response = requests.get(url, headers=headers)

        if response.status_code != 200:
            raise Exception(response.text)

        data = response.json()
        all_records.extend(data.get("value", []))
        url = data.get("@odata.nextLink")

    return all_records


def fetch_table_from_dataverse_url(
    dataverse_url,
    table_name,
    token,
    ticket_column,
    ticket_value,
):
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
    }

    url = (
        f"{dataverse_url}/api/data/v9.2/{table_name}"
        f"?$filter={ticket_column} eq '{ticket_value}'"
    )

    all_records = []

    while url:
        response = requests.get(url, headers=headers)

        if response.status_code != 200:
            raise Exception(response.text)

        data = response.json()
        all_records.extend(data.get("value", []))
        url = data.get("@odata.nextLink")

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
            "No Closing Ticket Details records found "
            f"in table '{closing_table}' for ticket '{ticket_value}'."
        )

    if not invoice_data:
        errors.append(
            "No Invoice Details records found "
            f"in table '{invoice_table}' for ticket '{ticket_value}'."
        )

    if errors:
        message = (
            f"Dataverse data missing for environment '{env}'. "
            + " ".join(errors)
        )
        raise ValueError(message)


def get_choice_label(mapping, value):
    if value in (None, "") or pd.isna(value):
        return ""

    try:
        normalized_value = int(value)
    except (TypeError, ValueError):
        normalized_value = value

    return mapping.get(normalized_value, "")


def get_row_id(row, table_name):
    primary_key = table_name[:-2] + "id" if table_name.endswith("es") else f"{table_name}id"

    if primary_key in row:
        return row[primary_key]

    for key, value in row.items():
        if key.endswith("id") and not key.startswith("_"):
            return value

    raise ValueError(f"Could not determine Dataverse row id for table '{table_name}'.")


def _headers(token, extra=None):
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }

    if extra:
        headers.update(extra)

    return headers


def get_onedrive_file(token, user_email, file_path):
    url = (
        f"{GRAPH_BASE}/users/{user_email}"
        f"/drive/root:/{file_path}:/"
    )

    response = requests.get(url, headers={"Authorization": f"Bearer {token}"})

    if response.status_code != 200:
        raise Exception(response.text)

    return response.json()


def create_onedrive_folder(token, user_email, base_folder, ticket_id):
    parent_path = base_folder
    url = (
        f"{GRAPH_BASE}/users/{user_email}"
        f"/drive/root:/{parent_path}:/children"
    )
    body = {
        "name": ticket_id,
        "folder": {},
        "@microsoft.graph.conflictBehavior": "replace",
    }

    response = requests.post(url, headers=_headers(token), json=body)

    if response.status_code not in [200, 201]:
        raise Exception(response.text)

    print("Folder created successfully")


def upload_file_to_onedrive(
    token,
    user_email,
    base_folder,
    ticket_id,
    local_file_path,
    onedrive_file_name,
):
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": (
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    }

    onedrive_folder = f"{base_folder}/{ticket_id}"
    url = (
        f"{GRAPH_BASE}/users/{user_email}"
        f"/drive/root:/{onedrive_folder}/"
        f"{onedrive_file_name}:/content"
    )

    with open(local_file_path, "rb") as file:
        file_content = file.read()

    response = requests.put(url, headers=headers, data=file_content)

    if response.status_code not in [200, 201]:
        raise Exception(response.text)

    print("File uploaded successfully")


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
        f"{dataverse_url}/api/data/v9.2/{table_name}"
        f"({row_id})/{file_column}"
    )
    headers = {
        "Authorization": f"Bearer {token}",
        "OData-MaxVersion": "4.0",
        "OData-Version": "4.0",
        "If-None-Match": "null",
        "Accept": "application/json",
        "Content-Type": "application/octet-stream",
        "x-ms-file-name": file_name,
    }

    with open(local_file_path, "rb") as file:
        response = requests.patch(url, headers=headers, data=file.read())

    if response.status_code != 204:
        raise Exception(
            f"Failed to upload PDF to Dataverse file column "
            f"'{file_column}': {response.text}"
        )

    print(f"Uploaded PDF to Dataverse column: {table_name}.{file_column}")


def _ensure_folder(token, user_email, folder_path):
    check_url = (
        f"{GRAPH_BASE}/users/{user_email}"
        f"/drive/root:/{folder_path}"
    )

    response = requests.get(check_url, headers=_headers(token))

    if response.status_code == 200:
        return

    parts = folder_path.rsplit("/", 1)
    parent_path = parts[0] if len(parts) == 2 else ""
    folder_name = parts[-1]

    if parent_path:
        create_url = (
            f"{GRAPH_BASE}/users/{user_email}"
            f"/drive/root:/{parent_path}:/children"
        )
    else:
        create_url = f"{GRAPH_BASE}/users/{user_email}/drive/root/children"

    body = {
        "name": folder_name,
        "folder": {},
        "@microsoft.graph.conflictBehavior": "replace",
    }

    response = requests.post(create_url, headers=_headers(token), json=body)

    if response.status_code not in [200, 201]:
        raise Exception(f"Failed to create folder '{folder_path}': {response.text}")

    print(f"\nFolder created: {folder_path}")


def _get_file_metadata(token, user_email, file_path):
    url = (
        f"{GRAPH_BASE}/users/{user_email}"
        f"/drive/root:/{file_path}"
    )

    response = requests.get(url, headers=_headers(token))

    if response.status_code == 200:
        return response.json()

    if response.status_code == 404:
        return None

    raise Exception(f"Error checking file '{file_path}': {response.text}")


def _move_and_rename_file(
    token,
    user_email,
    source_path,
    destination_folder_path,
    new_name,
):
    folder_url = (
        f"{GRAPH_BASE}/users/{user_email}"
        f"/drive/root:/{destination_folder_path}"
    )
    folder_response = requests.get(folder_url, headers=_headers(token))

    if folder_response.status_code != 200:
        raise Exception(
            f"Cannot find destination folder '{destination_folder_path}': "
            f"{folder_response.text}"
        )

    folder_id = folder_response.json()["id"]
    move_url = (
        f"{GRAPH_BASE}/users/{user_email}"
        f"/drive/root:/{source_path}"
    )
    body = {
        "parentReference": {"id": folder_id},
        "name": new_name,
    }

    response = requests.patch(move_url, headers=_headers(token), json=body)

    if response.status_code != 200:
        raise Exception(f"Failed to move/rename file: {response.text}")

    print(f"\nArchived: {new_name}")


def _delete_file(token, user_email, file_path):
    url = (
        f"{GRAPH_BASE}/users/{user_email}"
        f"/drive/root:/{file_path}"
    )

    response = requests.delete(url, headers=_headers(token))

    if response.status_code not in [204, 404]:
        raise Exception(f"Failed to delete '{file_path}': {response.text}")


def setup_ticket_folders(token, user_email, base_folder, ticket_id):
    ticket_folder = f"{base_folder}/{ticket_id}"
    active_folder = f"{ticket_folder}/Active"
    inactive_folder = f"{ticket_folder}/Inactive"

    _ensure_folder(token, user_email, ticket_folder)
    _ensure_folder(token, user_email, active_folder)
    _ensure_folder(token, user_email, inactive_folder)

    return active_folder, inactive_folder


def archive_active_file_if_exists(
    token,
    user_email,
    active_folder,
    inactive_folder,
    output_file_name,
):
    active_file_path = f"{active_folder}/{output_file_name}"
    existing = _get_file_metadata(token, user_email, active_file_path)

    if existing is None:
        print("\nNo existing file in Active - skipping archive")
        return

    now = datetime.now()
    date_str = now.strftime("%Y-%m-%d")
    datetime_str = now.strftime("%Y-%m-%d_%H-%M-%S")
    name_parts = output_file_name.rsplit(".", 1)

    if len(name_parts) == 2:
        archived_name = f"{name_parts[0]}_{datetime_str}.{name_parts[1]}"
    else:
        archived_name = f"{output_file_name}_{datetime_str}"

    date_folder = f"{inactive_folder}/{date_str}"
    _ensure_folder(token, user_email, date_folder)
    _move_and_rename_file(
        token,
        user_email,
        active_file_path,
        date_folder,
        archived_name,
    )

    print(f"\nArchived to: Inactive/{date_str}/{archived_name}")


def _list_folder_children(token, user_email, folder_path):
    """Return list of file items inside a OneDrive folder."""
    url = (
        f"{GRAPH_BASE}/users/{user_email}"
        f"/drive/root:/{folder_path}:/children"
    )

    all_items = []

    while url:
        response = requests.get(url, headers=_headers(token))

        if response.status_code == 404:
            return []

        if response.status_code != 200:
            raise Exception(
                f"Failed to list folder '{folder_path}': {response.text}"
            )

        data = response.json()
        all_items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")

    return all_items


def archive_active_files_by_prefix(
    token,
    user_email,
    active_folder,
    inactive_folder,
    file_prefix,
):
    """Archive every file in active_folder whose name starts with file_prefix."""
    items = _list_folder_children(token, user_email, active_folder)

    matching = [
        item for item in items
        if "folder" not in item
        and item.get("name", "").startswith(file_prefix)
    ]

    if not matching:
        print(f"\nNo files matching prefix '{file_prefix}' in Active - skipping")
        return

    now = datetime.now()
    date_str = now.strftime("%Y-%m-%d")
    datetime_str = now.strftime("%Y-%m-%d_%H-%M-%S")
    date_folder = f"{inactive_folder}/{date_str}"
    _ensure_folder(token, user_email, date_folder)

    for item in matching:
        file_name = item["name"]
        active_file_path = f"{active_folder}/{file_name}"
        name_parts = file_name.rsplit(".", 1)

        if len(name_parts) == 2:
            archived_name = f"{name_parts[0]}_{datetime_str}.{name_parts[1]}"
        else:
            archived_name = f"{file_name}_{datetime_str}"

        _move_and_rename_file(
            token,
            user_email,
            active_file_path,
            date_folder,
            archived_name,
        )

        print(f"\nArchived to: Inactive/{date_str}/{archived_name}")


def copy_template_to_active(
    token,
    user_email,
    template_path,
    active_folder,
    output_file_name,
):
    copy_url = (
        f"{GRAPH_BASE}/users/{user_email}"
        f"/drive/root:/{template_path}:/copy"
    )
    body = {
        "parentReference": {"path": f"/drive/root:/{active_folder}"},
        "name": output_file_name,
    }

    response = requests.post(copy_url, headers=_headers(token), json=body)

    if response.status_code not in [200, 201, 202]:
        print(response.text)
        raise Exception("Template copy failed")

    monitor_url = response.headers.get("Location")

    if monitor_url:
        print("\nWaiting for copy to complete...")
        _poll_copy_operation(monitor_url)

    print("\nTemplate copied to Active/")
    return f"{active_folder}/{output_file_name}"


def _poll_copy_operation(monitor_url, max_retries=20):
    for _ in range(max_retries):
        response = requests.get(monitor_url)
        data = response.json()
        status = data.get("status", "")

        if status == "completed":
            return

        if status == "failed":
            raise Exception(f"Copy operation failed: {data}")

        time.sleep(2)

    raise Exception("Copy operation timed out")


def create_workbook_session(token, user_email, file_path, max_retries=10):
    url = (
        f"{GRAPH_BASE}/users/{user_email}"
        f"/drive/root:/{file_path}"
        f":/workbook/createSession"
    )

    response = None

    for attempt in range(max_retries):
        response = requests.post(
            url,
            headers=_headers(token),
            json={"persistChanges": True},
        )

        if response.status_code == 201:
            session_id = response.json()["id"]
            print("\nWorkbook session created")
            return session_id

        print(
            "\nWaiting for workbook availability "
            f"(Attempt {attempt + 1}/{max_retries})"
        )
        time.sleep(3)

    if response is not None:
        print(response.text)

    raise Exception("Failed to create workbook session")


def close_workbook_session(token, user_email, file_path, session_id):
    url = (
        f"{GRAPH_BASE}/users/{user_email}"
        f"/drive/root:/{file_path}"
        f":/workbook/closeSession"
    )

    requests.post(
        url,
        headers=_headers(token, {"workbook-session-id": session_id}),
    )

    print("\nWorkbook session closed")


def update_cell_range(
    token,
    user_email,
    file_path,
    session_id,
    sheet_name,
    cell_range,
    values,
):
    headers = _headers(token, {"workbook-session-id": session_id})
    encoded_sheet = requests.utils.quote(sheet_name)
    url = (
        f"{GRAPH_BASE}/users/{user_email}"
        f"/drive/root:/{file_path}"
        f":/workbook/worksheets/{encoded_sheet}"
        f"/range(address='{cell_range}')"
    )

    response = requests.patch(url, headers=headers, json={"values": values})

    if response.status_code != 200:
        print(response.text)
        raise Exception(f"Failed to update range {cell_range}")


def populate_excel_template(
    token,
    user_email,
    file_path,
    session_id,
    closing_ticket_df,
    invoice_df,
):
    sheet = "Closing Check Transmittal Form"
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
    current_date = datetime.now().strftime("%m/%d/%Y")

    try:
        if closing_date:
            closing_date = pd.to_datetime(closing_date).strftime("%m/%d/%Y")
    except Exception:
        pass

    def patch(cell_range, values):
        update_cell_range(
            token,
            user_email,
            file_path,
            session_id,
            sheet,
            cell_range,
            values,
        )

    patch("D1", [[current_date]])
    patch("D2", [[purchase_price]])
    patch("D3", [[deal]])
    patch("D4", [[shares]])
    patch("D5", [[closing_date]])
    patch("D6", [[seller_tcode]])
    patch("D7", [[property_address]])
    patch("D8", [[unit]])
    patch("C13", [[seller1_name]])
    patch("C43", [[buyer1_name]])
    patch("B86", [[closing_agent]])
    patch("B87", [[closing_agent_email]])
    patch("B88", [[closing_agent_phone]])
    patch("B89", [[closing_agent_title]])
    patch("B91", [[notes]])

    seller_df = invoice_df[invoice_df["cr7de_paidby"] == 716070000]

    for idx, (_, inv_row) in enumerate(seller_df.iterrows()):
        row_number = 15 + idx
        due_at_closing = get_choice_label(
            DUE_AT_CLOSING_MAP,
            inv_row.get("cr109_dueatclosing", ""),
        )
        payable_to = get_choice_label(
            PAYABLE_MAP,
            inv_row.get("cr7de_payableto", ""),
        )

        patch(
            f"A{row_number}:D{row_number}",
            [[
                inv_row.get("cr7de_chequenumber", ""),
                due_at_closing,
                inv_row.get("cr7de_amount", ""),
                payable_to,
            ]],
        )

    buyer_df = invoice_df[invoice_df["cr7de_paidby"] == 716070001]

    for idx, (_, inv_row) in enumerate(buyer_df.iterrows()):
        row_number = 45 + idx
        due_at_closing = get_choice_label(
            DUE_AT_CLOSING_MAP,
            inv_row.get("cr109_dueatclosing", ""),
        )
        payable_to = get_choice_label(
            PAYABLE_MAP,
            inv_row.get("cr7de_payableto", ""),
        )

        patch(
            f"A{row_number}:D{row_number}",
            [[
                inv_row.get("cr7de_chequenumber", ""),
                due_at_closing,
                inv_row.get("cr7de_amount", ""),
                payable_to,
            ]],
        )

    print("\nExcel populated successfully")


def _get_worksheets(token, user_email, file_path, session_id):
    headers = _headers(token, {"workbook-session-id": session_id})
    encoded_path = requests.utils.quote(file_path, safe="/:")
    url = (
        f"{GRAPH_BASE}/users/{user_email}"
        f"/drive/root:/{file_path}"
        f":/workbook/worksheets"
    )

    response = requests.get(url, headers=headers)

    if response.status_code != 200:
        raise Exception(f"Failed to list worksheets: {response.text}")

    return response.json().get("value", [])


def _set_worksheet_visibility(token, user_email, file_path, session_id, sheet_name, visibility):
    headers = _headers(token, {"workbook-session-id": session_id})
    encoded_sheet = requests.utils.quote(sheet_name)
    url = (
        f"{GRAPH_BASE}/users/{user_email}"
        f"/drive/root:/{file_path}"
        f":/workbook/worksheets/{encoded_sheet}"
    )

    response = requests.patch(url, headers=headers, json={"visibility": visibility})

    if response.status_code != 200:
        raise Exception(
            f"Failed to set visibility '{visibility}' on sheet '{sheet_name}': {response.text}"
        )


def convert_onedrive_file_to_pdf(token, user_email, file_path, pdf_output_path):
    target_sheet = "Closing Check Transmittal Form"

    # Open a session to hide non-target sheets before PDF conversion
    session_id = create_workbook_session(token, user_email, file_path)
    hidden_sheets = []

    try:
        worksheets = _get_worksheets(token, user_email, file_path, session_id)

        for ws in worksheets:
            name = ws.get("name", "")
            visibility = ws.get("visibility", "Visible")

            if name != target_sheet and visibility != "Hidden":
                _set_worksheet_visibility(
                    token, user_email, file_path, session_id, name, "Hidden"
                )
                hidden_sheets.append(name)
                print(f"  Hidden sheet: {name}")

    finally:
        close_workbook_session(token, user_email, file_path, session_id)

    # Convert to PDF (only the visible sheet will be included)
    url = (
        f"{GRAPH_BASE}/users/{user_email}"
        f"/drive/root:/{file_path}:/content"
        "?format=pdf"
    )

    response = requests.get(url, headers=_headers(token))
    print("\nPDF conversion response:", response.status_code)

    if response.status_code != 200:
        print(response.text)
        raise Exception("PDF conversion failed")

    with open(pdf_output_path, "wb") as file:
        file.write(response.content)

    print("\nPDF generated successfully")

    # Restore visibility of hidden sheets
    if hidden_sheets:
        session_id = create_workbook_session(token, user_email, file_path)

        try:
            for name in hidden_sheets:
                _set_worksheet_visibility(
                    token, user_email, file_path, session_id, name, "Visible"
                )
                print(f"  Restored sheet: {name}")
        finally:
            close_workbook_session(token, user_email, file_path, session_id)


def main():
    args = parse_args()
    env = normalize_environment(
        args.env
        or args.environment
        or os.getenv("APP_ENV")
        or os.getenv("ENV")
        or config.get("environment")
    )
    runtime = get_runtime_settings(config, env)

    dv_token = get_access_token(runtime["dataverse_scope"])
    graph_token = get_access_token(config["auth"]["graph_scope"])

    dv = config["dataverse"]
    dv_tables = runtime["dataverse_tables"]
    dataverse_url = runtime["dataverse_url"]

    ticket_column = dv["columns"]["ticket_id"]
    ticket_value = dv["filter"]["ticket_id"]
    template_path = runtime["invoice_template_path"]
    user_email = runtime["user_email"]
    base_folder = runtime["base_folder"]
    output_excel_name = f"{ticket_value}_Closing_Form.xlsx"
    pdf_output_path = f"{ticket_value}_Closing_Form.pdf"

    print(f"\nEnvironment: {env}")
    print(f"OneDrive user: {user_email}")
    print(f"Dataverse URL: {dataverse_url}")

    closing_table = dv_tables["closing_ticket_details"]
    invoice_table = dv_tables["invoice_details"]

    closing_data = fetch_table_from_dataverse_url(
        dataverse_url,
        closing_table,
        dv_token,
        ticket_column,
        ticket_value,
    )
    invoice_data = fetch_table_from_dataverse_url(
        dataverse_url,
        invoice_table,
        dv_token,
        ticket_column,
        ticket_value,
    )

    print(f"Closing Ticket Details records: {len(closing_data)}")
    print(f"Invoice Details records: {len(invoice_data)}")

    validate_required_dataverse_data(
        env,
        ticket_value,
        closing_data,
        invoice_data,
        closing_table,
        invoice_table,
    )

    df_closing = pd.DataFrame(closing_data)
    df_invoice = pd.DataFrame(invoice_data)
    closing_row_id = get_row_id(closing_data[0], closing_table)

    active_folder, inactive_folder = setup_ticket_folders(
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

    # Archive any existing Condo Invoice files (.docx, .pdf, or any extension)
    archive_active_files_by_prefix(
        graph_token,
        user_email,
        active_folder,
        inactive_folder,
        f"{ticket_value}_Condo_Invoice",
    )

    output_file_path = copy_template_to_active(
        graph_token,
        user_email,
        template_path,
        active_folder,
        output_excel_name,
    )

    session_id = create_workbook_session(graph_token, user_email, output_file_path)

    try:
        populate_excel_template(
            graph_token,
            user_email,
            output_file_path,
            session_id,
            df_closing,
            df_invoice,
        )
    finally:
        close_workbook_session(
            graph_token,
            user_email,
            output_file_path,
            session_id,
        )

    convert_onedrive_file_to_pdf(
        graph_token,
        user_email,
        output_file_path,
        pdf_output_path,
    )

    upload_file_to_onedrive(
        graph_token,
        user_email,
        base_folder,
        f"{ticket_value}/Active",
        pdf_output_path,
        pdf_output_path,
    )

    upload_file_to_dataverse_file_column(
        dataverse_url,
        closing_table,
        closing_row_id,
        dv_token,
        CLOSING_PDF_FILE_COLUMN,
        pdf_output_path,
        pdf_output_path,
    )

    if os.path.exists(pdf_output_path):
        os.remove(pdf_output_path)

    print("\nProcess completed")


if __name__ == "__main__":
    main()
