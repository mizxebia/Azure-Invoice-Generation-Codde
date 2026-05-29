import json
import logging
import os
import time
import traceback
from datetime import datetime
from pathlib import Path
 
import azure.functions as func
import pandas as pd
import requests
from msal import ConfidentialClientApplication
 
 
CONFIG_PATH = Path(__file__).with_name("config.json")
 
GRAPH_BASE = "https://graph.microsoft.com/v1.0"
 
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
# COPY TEMPLATE
# =====================================
 
def _poll_copy_operation(
    monitor_url,
    max_retries=20,
):
 
    for _ in range(max_retries):
 
        response = requests.get(
            monitor_url
        )
 
        data = response.json()
 
        status = data.get(
            "status",
            ""
        )
 
        if status == "completed":
            return
 
        if status == "failed":
 
            raise Exception(
                f"Copy failed: {data}"
            )
 
        time.sleep(2)
 
    raise Exception(
        "Copy operation timed out"
    )
 
 
def copy_template_to_active(
    token,
    user_email,
    template_path,
    active_folder,
    output_file_name,
):
 
    copy_url = (
        f"{GRAPH_BASE}/users/"
        f"{user_email}"
        f"/drive/root:/"
        f"{template_path}"
        f":/copy"
    )
 
    body = {
 
        "parentReference": {
 
            "path":
                f"/drive/root:/"
                f"{active_folder}"
        },
 
        "name":
            output_file_name,
    }
 
    response = requests.post(
        copy_url,
        headers=_headers(token),
        json=body,
    )
 
    if response.status_code not in [
        200,
        201,
        202,
    ]:
 
        logger.error(response.text)
 
        raise Exception(
        f"Template copy failed: "
        f"{response.status_code} - "
        f"{response.text}"
    )
 
    monitor_url = response.headers.get(
        "Location"
    )
 
    if monitor_url:
 
        logger.info(
            "Waiting for copy..."
        )
 
        _poll_copy_operation(
            monitor_url
        )
 
        logger.info(
            "Waiting for workbook..."
        )
 
        time.sleep(10)
 
    logger.info(
        "Template copied"
    )
 
    return (
        f"{active_folder}/"
        f"{output_file_name}"
    )
 
 
# =====================================
# WORKBOOK SESSION
# =====================================
 
def create_workbook_session(
    token,
    user_email,
    file_path,
    max_retries=20,
):
 
    url = (
        f"{GRAPH_BASE}/users/"
        f"{user_email}"
        f"/drive/root:/"
        f"{file_path}"
        f":/workbook/createSession"
    )
 
    response = None
 
    for attempt in range(max_retries):
 
        response = requests.post(
 
            url,
 
            headers=_headers(token),
 
            json={
                "persistChanges": True
            },
        )
 
        if response.status_code == 201:
 
            session_id = (
                response.json()["id"]
            )
 
            logger.info(
                "Workbook session created"
            )
 
            return session_id
 
        logger.info(
            f"Workbook retry "
            f"{attempt + 1}/"
            f"{max_retries}"
        )
 
        time.sleep(5)
 
    if response is not None:
 
        logger.error(response.text)
 
    raise Exception(
        f"Failed to create "
        f"workbook session: "
        f"{response.text}"
    )
 
 
def close_workbook_session(
    token,
    user_email,
    file_path,
    session_id,
):
 
    url = (
        f"{GRAPH_BASE}/users/"
        f"{user_email}"
        f"/drive/root:/"
        f"{file_path}"
        f":/workbook/closeSession"
    )
 
    requests.post(
 
        url,
 
        headers=_headers(
            token,
            {
                "workbook-session-id":
                    session_id
            },
        ),
    )
 
    logger.info(
        "Workbook session closed"
    )
 
 
# =====================================
# EXCEL
# =====================================
 
def update_cell_range(
    token,
    user_email,
    file_path,
    session_id,
    sheet_name,
    cell_range,
    values,
):
 
    headers = _headers(
 
        token,
 
        {
            "workbook-session-id":
                session_id
        },
    )
 
    encoded_sheet = (
        requests.utils.quote(
            sheet_name
        )
    )
 
    url = (
        f"{GRAPH_BASE}/users/"
        f"{user_email}"
        f"/drive/root:/"
        f"{file_path}"
        f":/workbook/worksheets/"
        f"{encoded_sheet}"
        f"/range(address='"
        f"{cell_range}')"
    )
 
    response = requests.patch(
 
        url,
 
        headers=headers,
 
        json={"values": values},
    )
 
    if response.status_code != 200:
 
        logger.error(response.text)
 
        raise Exception(
            f"Failed updating "
            f"{cell_range}"
        )
 
 
def populate_excel_template(
    token,
    user_email,
    file_path,
    session_id,
    closing_ticket_df,
    invoice_df,
):
 
    sheet = (
        "Closing Check "
        "Transmittal Form"
    )
 
    row = closing_ticket_df.iloc[0]

    purchase_price = row.get(
        "cr109_saleprice",
        ""
    )

    closing_date = row.get(
        "cr7de_closingdate",
        ""
    )

    seller_tcode = row.get(
        "cr7de_sellertcode",
        ""
    )

    property_address = row.get(
        "cr7de_buildingaddress",
        ""
    )

    unit = row.get(
        "cr7de_unitnumber",
        ""
    )

    seller1_name = row.get(
        "cr7de_sellername",
        ""
    )

    deal = get_choice_label(
        TRANSACTION_TYPE_DEAL_MAP,
        row.get(
            "cr109_transactiontypedeal",
            ""
        )
    )

    buyer1_name = row.get(
        "cr7de_buyername",
        ""
    )

    shares = row.get(
        "cr109_shares",
        ""
    )

    closing_agent = row.get(
        "cr7de_closingagentname",
        ""
    )

    closing_agent_phone = row.get(
        "cr7de_closingagentphone",
        ""
    )

    closing_agent_email = row.get(
        "cr7de_closingagentemail",
        ""
    )

    closing_agent_title = row.get(
        "cr7de_titlerole",
        ""
    )

    notes = row.get(
        "cr7de_notes",
        ""
    )

    current_date = datetime.now().strftime(
        "%m/%d/%Y"
    )

    try:

        if closing_date:

            closing_date = (
                pd.to_datetime(
                    closing_date
                )
                .strftime("%m/%d/%Y")
            )

    except Exception:

        pass
 
    def patch(
        cell_range,
        values
    ):
 
        update_cell_range(
            token,
            user_email,
            file_path,
            session_id,
            sheet,
            cell_range,
            values,
        )
 
    patch(
        "D1",
        [[
            current_date
        ]]
    )

    patch(
        "D2",
        [[
            purchase_price
        ]]
    )

    patch(
        "D3",
        [[
            deal
        ]]
    )

    patch(
        "D4",
        [[
            shares
        ]]
    )

    patch(
        "D5",
        [[
            closing_date
        ]]
    )

    patch(
        "D6",
        [[
            seller_tcode
        ]]
    )

    patch(
        "D7",
        [[
            property_address
        ]]
    )

    patch(
        "D8",
        [[
            unit
        ]]
    )

    patch(
        "C13",
        [[
            seller1_name
        ]]
    )

    patch(
        "C43",
        [[
            buyer1_name
        ]]
    )

    patch(
        "B86",
        [[
            closing_agent
        ]]
    )

    patch(
        "B87",
        [[
            closing_agent_email
        ]]
    )

    patch(
        "B88",
        [[
            closing_agent_phone
        ]]
    )

    patch(
        "B89",
        [[
            closing_agent_title
        ]]
    )

    patch(
        "B91",
        [[
            notes
        ]]
    )

    if (
        "cr7de_paidby"
        not in invoice_df.columns
    ):
 
        raise Exception(
            f"Column cr7de_paidby "
            f"not found. "
            f"Available columns: "
            f"{invoice_df.columns.tolist()}"
        )
 
    seller_df = invoice_df[
        invoice_df["cr7de_paidby"]
        == 716070000
    ]

    for idx, (_, inv_row) in enumerate(
        seller_df.iterrows()
    ):

        row_number = 15 + idx

        due_at_closing = get_choice_label(
            DUE_AT_CLOSING_MAP,
            inv_row.get(
                "cr109_dueatclosing",
                ""
            )
        )

        payable_to = get_choice_label(
            PAYABLE_MAP,
            inv_row.get(
                "cr7de_payableto",
                ""
            )
        )

        patch(
            f"A{row_number}:D{row_number}",
            [[
                inv_row.get(
                    "cr7de_chequenumber",
                    ""
                ),
                due_at_closing,
                inv_row.get(
                    "cr7de_amount",
                    ""
                ),
                payable_to,
            ]],
        )

    buyer_df = invoice_df[
        invoice_df["cr7de_paidby"]
        == 716070001
    ]

    for idx, (_, inv_row) in enumerate(
        buyer_df.iterrows()
    ):

        row_number = 45 + idx

        due_at_closing = get_choice_label(
            DUE_AT_CLOSING_MAP,
            inv_row.get(
                "cr109_dueatclosing",
                ""
            )
        )

        payable_to = get_choice_label(
            PAYABLE_MAP,
            inv_row.get(
                "cr7de_payableto",
                ""
            )
        )

        patch(
            f"A{row_number}:D{row_number}",
            [[
                inv_row.get(
                    "cr7de_chequenumber",
                    ""
                ),
                due_at_closing,
                inv_row.get(
                    "cr7de_amount",
                    ""
                ),
                payable_to,
            ]],
        )

    logger.info(
        "Excel populated successfully"
    )
 
 
# =====================================
# PDF
# =====================================
 
def convert_onedrive_file_to_pdf(
    token,
    user_email,
    file_path,
    pdf_output_path,
):
 
    url = (
        f"{GRAPH_BASE}/users/"
        f"{user_email}"
        f"/drive/root:/"
        f"{file_path}"
        f":/content?format=pdf"
    )
 
    response = requests.get(
 
        url,
 
        headers=_headers(token)
    )
 
    logger.info(
        f"PDF response: "
        f"{response.status_code}"
    )
 
    if response.status_code != 200:
 
        logger.error(response.text)
 
        raise Exception(
            "PDF conversion failed"
        )
 
    with open(
        pdf_output_path,
        "wb"
    ) as file:
 
        file.write(
            response.content
        )
 
    logger.info(
        "PDF generated"
    )
 
 
def upload_file_to_onedrive(
    token,
    user_email,
    base_folder,
    ticket_id,
    local_file_path,
    onedrive_file_name,
):
 
    headers = {
 
        "Authorization":
            f"Bearer {token}",
 
        "Content-Type":
            "application/pdf",
    }
 
    onedrive_folder = (
        f"{base_folder}/"
        f"{ticket_id}"
    )
 
    url = (
        f"{GRAPH_BASE}/users/"
        f"{user_email}"
        f"/drive/root:/"
        f"{onedrive_folder}/"
        f"{onedrive_file_name}"
        f":/content"
    )
 
    with open(
        local_file_path,
        "rb"
    ) as file:
 
        file_content = file.read()
 
    response = requests.put(
 
        url,
 
        headers=headers,
 
        data=file_content,
    )
 
    if response.status_code not in [
        200,
        201,
    ]:
 
        logger.error(response.text)
 
        raise Exception(
            response.text
        )
 
    logger.info(
        "PDF uploaded"
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
 
        pdf_output_path = os.path.join(
 
            "/tmp",
 
            f"{ticket_value}"
            f"_Closing_Form.pdf"
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

        output_file_path = (
            copy_template_to_active(
 
                graph_token,
 
                user_email,
 
                template_path,
 
                active_folder,
 
                output_excel_name,
            )
        )
 
        session_id = (
            create_workbook_session(
 
                graph_token,
 
                user_email,
 
                output_file_path,
            )
        )
 
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
 
            f"{ticket_value}"
            f"_Closing_Form.pdf",
        )
 
        if os.path.exists(
            pdf_output_path
        ):
 
            os.remove(
                pdf_output_path
            )
 
        return func.HttpResponse(
 
            body=json.dumps({
 
                "status":
                    "SUCCESS",
 
                "ticket_id":
                    ticket_value
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
