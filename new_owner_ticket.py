import argparse
import json
import logging
import os
import re
import time
import traceback
import zipfile
from datetime import datetime
from pathlib import Path
from xml.dom import minidom

import pandas as pd
import requests
from msal import ConfidentialClientApplication


CONFIG_PATH = Path(__file__).with_name("config.json")
GRAPH_BASE = "https://graph.microsoft.com/v1.0"
WORK_DIR = Path(__file__).with_name("_generated")
NEW_OWNER_PDF_FILE_COLUMN = "cr109_newownerticketpdf"

ENVIRONMENT_SETTINGS = {
    "DEV": {
        "user_email_key": "user_email_dev",
        "user_email": "akambotdev1@akam.com",
        "template_path": (
            "New Sales RPA/DEV/Excel Sheets/New Owner Ticket Submission.docx"
        ),
        "base_folder": "New Sales RPA/DEV/BotShareDrive/InProgress",
    },
    "UAT": {
        "user_email_key": "user_email_uat",
        "user_email": "akambotuat2@akam.com",
        "template_path": (
            "New Sales RPA/UAT/Excel Sheets/New Owner Ticket Submission.docx"
        ),
        "base_folder": "New Sales RPA/UAT/BotShareDrive/InProgress",
    },
    "PROD": {
        "user_email_key": "user_email_prod",
        "user_email": "akambotnewsalesclosure@akam.com",
        "template_path": (
            "New Sales RPA/PROD/Excel Sheets/New Owner Ticket Submission.docx"
        ),
        "base_folder": "New Sales RPA/PROD/BotShareDrive/InProgress",
    },
}

WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"


FIELD_MAP = {
    "Property ID": "cr109_nyccode",
    "Closing Date": "cr7de_closingdate",
    "Sales Price": "cr109_PurchasePrice",
    "Building Name": "cr109_buildingname",
    "Property Address": "cr7de_address",
    "Unit Number": "cr7de_Unit",
    "Seller T-Code": "cr7de_SellerTCode",
    "Seller Name": "cr7de_SellerName",
    "Seller Contact Number": "cr7de_SellerContactNumber",
    "Seller Contact Email": "cr7de_SellerContactEmail",
    "Seller SSN/EIN": "cr7de_SellerSSNEIN",
    "Secondary Seller Name": "cr109_Seller2Name",
    "Secondary Seller Contact Email": "",
    "Secondary Seller SSN/EIN": "cr109_Seller2SSNEIN",
    "Primary Owner Name ": "cr7de_NewPrimaryOwnerName",
    "Primary Owner TCODE": "cr109_primaryownertcode",
    "Primary Owner SSN/EIN": "cr7de_PrimaryOwnerSSNEIN",
    "Primary Owner Email": "cr7de_PrimaryOwnerEmail",
    "Secondary Owner Name": "cr7de_NewSecondaryOwnerName",
    "Secondary Owner SSN/EIN": "cr7de_SecondaryOwnerSSNEIN",
    "Secondary Owner Email": "cr7de_SecondaryOwnerEmail",
    "Alternate Mailing Address": "cr7de_AlternateMailingAddress",
    "Payments Applied to New Owner Account": "cr7de_PaymentAppliedtoSellerAccount",
    "Buyer 1 Occupancy": "cr109_Purchaser1Occupancy",
    "Buyer 2 Occupancy": "cr109_purchaser2occupancy",
    "Additional Occupant Name 1": "cr109_Additional_Occupants1Name",
    "Additional Occupant Name 2": "cr109_AdditionalOccupant2Name",
}


ADDRESS_FIELDS = {
    "Seller Forwarding Address": [
        "cr109_Seller1Address",
        "cr109_Seller1City",
        "cr109_Seller1State",
        "cr109_Seller1ZIP",
    ],
    "Secondary Seller Forwarding Address": [
        "cr109_Seller2Address",
        "cr109_Seller2City",
        "cr109_Seller2State",
        "cr109_Seller2ZIP",
    ],
}


PHONE_FIELDS = {
    "Primary Owner Phone Number": [
        ("Work", "cr109_PrimaryWorkPhoneNumber"),
        ("Cell", "cr7de_PrimaryPhoneNumber"),
        ("Home", "cr109_PrimaryHomePhoneNumber"),
    ],
    "Secondary Owner Phone Number": [
        ("Work", "cr109_SecondaryOwnerWorkPhoneNumber"),
        ("Cell", "cr7de_secondaryphonenumber"),
        ("Home", "cr109_SecondaryOwnerHomePhoneNumber"),
    ],
}


def load_config():
    with open(CONFIG_PATH, "r", encoding="utf-8") as config_file:
        return json.load(config_file)


def parse_args():
    parser = argparse.ArgumentParser(
        description="Generate the New Owner Ticket Submission document."
    )
    parser.add_argument("environment", nargs="?", help="DEV, UAT, or PROD.")
    parser.add_argument("--env", dest="env", help="DEV, UAT, or PROD.")
    parser.add_argument("--id", dest="ticket_id", help="Ticket ID to process.")
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
    env_settings = ENVIRONMENT_SETTINGS[env]
    dataverse = auth_config.get("dataverse_by_env", {}).get(env, {})

    if env != "DEV" and not dataverse:
        raise ValueError(
            f"Missing Dataverse config for environment '{env}'. "
            f"Add auth.dataverse_by_env.{env} to config.json."
        )

    return {
        "env": env,
        "user_email": storage.get(
            env_settings["user_email_key"],
            env_settings["user_email"],
        ),
        "template_path": env_settings["template_path"],
        "base_folder": env_settings["base_folder"],
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


def _headers(token, extra=None):
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }

    if extra:
        headers.update(extra)

    return headers


def fetch_ticket_row(dataverse_url, table_name, token, ticket_column, ticket_value):
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
    }
    url = (
        f"{dataverse_url}/api/data/v9.2/{table_name}"
        f"?$filter={ticket_column} eq '{ticket_value}'"
    )

    response = requests.get(url, headers=headers)

    if response.status_code != 200:
        raise Exception(response.text)

    records = response.json().get("value", [])

    if not records:
        raise ValueError(
            f"No New Owner Ticket Details record found in table "
            f"'{table_name}' for ticket '{ticket_value}'."
        )

    if len(records) > 1:
        logger.warning(
            "Multiple New Owner Ticket Details records found for %s. "
            "Using the first record.",
            ticket_value,
        )

    return records[0]


def get_row_id(row, table_name):
    primary_key = table_name[:-2] + "id" if table_name.endswith("es") else f"{table_name}id"

    if primary_key in row:
        return row[primary_key]

    for key, value in row.items():
        if key.endswith("id") and not key.startswith("_"):
            return value

    raise ValueError(f"Could not determine Dataverse row id for table '{table_name}'.")


def row_value(row, field_name):
    if not field_name:
        return ""

    if field_name in row:
        return row.get(field_name)

    lower_field_name = field_name.lower()

    for key, value in row.items():
        if key.lower() == lower_field_name:
            return value

    return ""


def format_value(value):
    if value in (None, "") or pd.isna(value):
        return ""

    if isinstance(value, str):
        if value.strip().upper() in {"N/A", "NA", "NONE", "NULL"}:
            return ""

        try:
            parsed_date = pd.to_datetime(value, errors="raise")
            if "T" in value or "-" in value:
                return parsed_date.strftime("%m/%d/%Y")
        except Exception:
            return value.strip()

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if float(value).is_integer():
            return str(int(value))
        return str(value)

    return str(value)


def format_phone_number(value):
    value = format_value(value)

    if not value:
        return ""

    digits = re.sub(r"\D", "", value)

    if len(digits) == 11 and digits.startswith("1"):
        digits = digits[1:]

    if len(digits) == 10:
        return f"({digits[:3]}) {digits[3:6]}-{digits[6:]}"

    return value


def compile_address(row, fields):
    address = format_value(row_value(row, fields[0]))
    city = format_value(row_value(row, fields[1]))
    state = format_value(row_value(row, fields[2]))
    zip_code = format_value(row_value(row, fields[3]))
    city_state_zip = " ".join(part for part in [state, zip_code] if part)

    if city and city_state_zip:
        city_state_zip = f"{city}, {city_state_zip}"
    elif city:
        city_state_zip = city

    return ", ".join(part for part in [address, city_state_zip] if part)


def compile_phone_numbers(row, phone_fields):
    values = []

    for _, field_name in phone_fields:
        phone = format_phone_number(row_value(row, field_name))

        if phone:
            values.append(phone)

    return ", ".join(values)


def build_content_control_values(row):
    values = {}

    for tag, field_name in FIELD_MAP.items():
        values[tag] = format_value(row_value(row, field_name))

    for tag, fields in ADDRESS_FIELDS.items():
        values[tag] = compile_address(row, fields)

    for tag, phone_fields in PHONE_FIELDS.items():
        values[tag] = compile_phone_numbers(row, phone_fields)

    # Seller 2 phone is intentionally blank for now per current requirement.
    values["Secondary Seller Contact Number"] = ""

    # This template field does not have a confirmed Dataverse mapping yet.
    values["Effective Month for Charge Migration to New Owner"] = ""

    return values


def populate_docx_template(template_path, output_path, values_by_tag):
    with zipfile.ZipFile(template_path, "r") as source_docx:
        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as output_docx:
            for item in source_docx.infolist():
                data = source_docx.read(item.filename)

                if item.filename.startswith("word/") and item.filename.endswith(".xml"):
                    try:
                        document = minidom.parseString(data)
                    except Exception:
                        output_docx.writestr(item, data)
                        continue

                    changed = False

                    for sdt in document.getElementsByTagNameNS(WORD_NS, "sdt"):
                        tag = get_sdt_tag(sdt)

                        if tag in values_by_tag:
                            set_sdt_text(sdt, values_by_tag[tag])
                            changed = True

                    if changed:
                        data = document.toxml(encoding="utf-8")

                output_docx.writestr(item, data)


def get_sdt_tag(sdt):
    tag_nodes = sdt.getElementsByTagNameNS(WORD_NS, "tag")

    if tag_nodes:
        return tag_nodes[0].getAttribute("w:val") or tag_nodes[0].getAttribute("val")

    alias_nodes = sdt.getElementsByTagNameNS(WORD_NS, "alias")

    if alias_nodes:
        return alias_nodes[0].getAttribute("w:val") or alias_nodes[0].getAttribute("val")

    return ""


def set_sdt_text(sdt, value):
    content_nodes = sdt.getElementsByTagNameNS(WORD_NS, "sdtContent")

    if not content_nodes:
        return

    text_nodes = content_nodes[0].getElementsByTagNameNS(WORD_NS, "t")

    if not text_nodes:
        return

    set_text_node_value(text_nodes[0], value)

    for text_node in text_nodes[1:]:
        set_text_node_value(text_node, "")


def set_text_node_value(text_node, value):
    while text_node.firstChild:
        text_node.removeChild(text_node.firstChild)

    text_node.setAttributeNS(XML_NS, "xml:space", "preserve")
    text_node.appendChild(text_node.ownerDocument.createTextNode(value))


def _ensure_folder(token, user_email, folder_path):
    check_url = (
        f"{GRAPH_BASE}/users/{user_email}"
        f"/drive/root:/{folder_path}"
    )
    response = requests.get(check_url, headers=_headers(token))

    if response.status_code == 200:
        print(f"Folder exists: {folder_path}")
        return

    parent_path, folder_name = folder_path.rsplit("/", 1)
    create_url = (
        f"{GRAPH_BASE}/users/{user_email}"
        f"/drive/root:/{parent_path}:/children"
    )
    body = {
        "name": folder_name,
        "folder": {},
        "@microsoft.graph.conflictBehavior": "replace",
    }
    response = requests.post(create_url, headers=_headers(token), json=body)

    if response.status_code not in [200, 201]:
        raise Exception(f"Failed to create folder '{folder_path}': {response.text}")

    print(f"Folder created: {folder_path}")


def _get_file_metadata(token, user_email, file_path):
    url = f"{GRAPH_BASE}/users/{user_email}/drive/root:/{file_path}"
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

    move_url = f"{GRAPH_BASE}/users/{user_email}/drive/root:/{source_path}"
    body = {
        "parentReference": {"id": folder_response.json()["id"]},
        "name": new_name,
    }
    response = requests.patch(move_url, headers=_headers(token), json=body)

    if response.status_code != 200:
        raise Exception(f"Failed to move/rename file: {response.text}")


def setup_ticket_folders(token, user_email, base_folder, ticket_id):
    ticket_folder = f"{base_folder}/{ticket_id}"
    active_folder = f"{ticket_folder}/Active"
    inactive_folder = f"{ticket_folder}/Inactive"

    _ensure_folder(token, user_email, ticket_folder)
    _ensure_folder(token, user_email, active_folder)
    _ensure_folder(token, user_email, inactive_folder)

    print(f"Active folder: {active_folder}")
    print(f"Inactive folder: {inactive_folder}")

    return active_folder, inactive_folder


def archive_file_if_exists(
    token,
    user_email,
    active_folder,
    inactive_folder,
    file_name,
):
    active_file_path = f"{active_folder}/{file_name}"
    existing = _get_file_metadata(token, user_email, active_file_path)

    if existing is None:
        return

    now = datetime.now()
    date_folder = f"{inactive_folder}/{now.strftime('%Y-%m-%d')}"
    timestamp = now.strftime("%Y-%m-%d_%H-%M-%S")
    name_parts = file_name.rsplit(".", 1)

    if len(name_parts) == 2:
        archived_name = f"{name_parts[0]}_{timestamp}.{name_parts[1]}"
    else:
        archived_name = f"{file_name}_{timestamp}"

    _ensure_folder(token, user_email, date_folder)
    _move_and_rename_file(
        token,
        user_email,
        active_file_path,
        date_folder,
        archived_name,
    )

    print(f"Archived existing file: {archived_name}")


def download_onedrive_file(token, user_email, file_path, local_file_path):
    url = (
        f"{GRAPH_BASE}/users/{user_email}"
        f"/drive/root:/{file_path}:/content"
    )
    response = requests.get(url, headers=_headers(token))

    if response.status_code != 200:
        raise Exception(f"Failed to download template: {response.text}")

    with open(local_file_path, "wb") as file:
        file.write(response.content)


def upload_file_to_onedrive(
    token,
    user_email,
    folder_path,
    local_file_path,
    onedrive_file_name,
    content_type,
):
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": content_type,
    }
    url = (
        f"{GRAPH_BASE}/users/{user_email}"
        f"/drive/root:/{folder_path}/{onedrive_file_name}:/content"
    )

    with open(local_file_path, "rb") as file:
        response = requests.put(url, headers=headers, data=file.read())

    if response.status_code not in [200, 201]:
        raise Exception(f"Failed to upload '{onedrive_file_name}': {response.text}")

    print(f"Uploaded: {folder_path}/{onedrive_file_name}")


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


def convert_onedrive_file_to_pdf(
    token,
    user_email,
    file_path,
    pdf_output_path,
    max_retries=6,
):
    url = (
        f"{GRAPH_BASE}/users/{user_email}"
        f"/drive/root:/{file_path}:/content"
        "?format=pdf"
    )

    response = None

    for attempt in range(max_retries):
        response = requests.get(url, headers=_headers(token))

        if response.status_code == 200:
            with open(pdf_output_path, "wb") as file:
                file.write(response.content)

            print("PDF generated successfully")
            return

        print(
            "Waiting for Word PDF conversion "
            f"(Attempt {attempt + 1}/{max_retries})"
        )
        time.sleep(5)

    raise Exception(f"PDF conversion failed: {response.text}")


def cleanup_local_files(*paths):
    for path in paths:
        if path and Path(path).exists():
            Path(path).unlink()


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

    dv = config["dataverse"]
    ticket_value = args.ticket_id or dv["filter"]["ticket_id"]
    ticket_column = dv["columns"]["ticket_id"]
    table_name = dv["tables"]["new_owner_ticket_details"]
    closing_table_name = dv["tables"]["closing_ticket_details"]

    dv_token = get_access_token(runtime["dataverse_scope"])
    graph_token = get_access_token(config["auth"]["graph_scope"])

    print(f"\nEnvironment: {env}")
    print(f"OneDrive user: {runtime['user_email']}")
    print(f"Dataverse URL: {runtime['dataverse_url']}")
    print(f"Ticket ID: {ticket_value}")

    row = fetch_ticket_row(
        runtime["dataverse_url"],
        table_name,
        dv_token,
        ticket_column,
        ticket_value,
    )
    closing_row = fetch_ticket_row(
        runtime["dataverse_url"],
        closing_table_name,
        dv_token,
        ticket_column,
        ticket_value,
    )
    closing_row_id = get_row_id(closing_row, closing_table_name)
    values_by_tag = build_content_control_values(row)

    active_folder, inactive_folder = setup_ticket_folders(
        graph_token,
        runtime["user_email"],
        runtime["base_folder"],
        ticket_value,
    )

    output_docx_name = f"{ticket_value}_New_Owner_Ticket_Submission.docx"
    output_pdf_name = f"{ticket_value}_New_Owner_Ticket_Submission.pdf"
    active_docx_path = f"{active_folder}/{output_docx_name}"

    archive_file_if_exists(
        graph_token,
        runtime["user_email"],
        active_folder,
        inactive_folder,
        output_docx_name,
    )
    archive_file_if_exists(
        graph_token,
        runtime["user_email"],
        active_folder,
        inactive_folder,
        output_pdf_name,
    )

    WORK_DIR.mkdir(exist_ok=True)
    local_template_path = WORK_DIR / f"{ticket_value}_template.docx"
    local_docx_path = WORK_DIR / output_docx_name
    local_pdf_path = WORK_DIR / output_pdf_name

    try:
        download_onedrive_file(
            graph_token,
            runtime["user_email"],
            runtime["template_path"],
            local_template_path,
        )
        populate_docx_template(
            local_template_path,
            local_docx_path,
            values_by_tag,
        )
        upload_file_to_onedrive(
            graph_token,
            runtime["user_email"],
            active_folder,
            local_docx_path,
            output_docx_name,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        convert_onedrive_file_to_pdf(
            graph_token,
            runtime["user_email"],
            active_docx_path,
            local_pdf_path,
        )
        upload_file_to_onedrive(
            graph_token,
            runtime["user_email"],
            active_folder,
            local_pdf_path,
            output_pdf_name,
            "application/pdf",
        )
        upload_file_to_dataverse_file_column(
            runtime["dataverse_url"],
            closing_table_name,
            closing_row_id,
            dv_token,
            NEW_OWNER_PDF_FILE_COLUMN,
            local_pdf_path,
            output_pdf_name,
        )
    finally:
        cleanup_local_files(local_template_path, local_docx_path, local_pdf_path)

    print("\nNew Owner Ticket Submission generated successfully")


if __name__ == "__main__":
    main()
