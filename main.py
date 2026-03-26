import asyncio
import logging
import sys
import argparse
import os
from datetime import datetime
from zoneinfo import ZoneInfo

import re

import fitz  # pymupdf

from automation_server_client import AutomationServer, Workqueue, WorkItemError, Credential, WorkItemStatus
from odk_tools.tracking import Tracker

from services.mail_service import MailService, extract_text_from_html, parse_email_data
from process.config import load_excel_mapping, get_excel_mapping
from kmd_nexus_client import NexusClientManager


tracker: Tracker = None
mail_service: MailService = None
nexus: NexusClientManager = None
regler: list = []
forløb: list = []
proces_navn = "Indgangen - hjælpemiddelansøgninger"


def parse_ansoegning(pdf_text: str, attachments: list) -> dict:
    cpr_match = re.search(r"CPR-nummer\n(.+)", pdf_text)
    cpr = cpr_match.group(1).strip() if cpr_match else None

    tlf_match = re.search(r"Telefonnummer\n(.+)", pdf_text)
    telefonnummer = tlf_match.group(1).strip() if tlf_match else None

    block_match = re.search(
        r"(Hvilken funktionsnedsættelse er årsag til ansøgningen\?.*?)Vedhæft eventuelt yderligere oplysninger",
        pdf_text, re.DOTALL
    )
    funktionsnedsaettelse_block = block_match.group(1).strip() if block_match else None

    hjaelpemidler_match = re.search(
        r"Hvilke hjælpemidler.*?afhjælpe dig i hverdagen\?\n(.*?)\n\s*\nVedhæft",
        pdf_text, re.DOTALL
    )
    hjaelpemidler = hjaelpemidler_match.group(1).strip() if hjaelpemidler_match else None

    antal_filer = len(attachments) if attachments else 0

    return {
        "cpr": cpr,
        "telefonnummer": telefonnummer,
        "funktionsnedsaettelse_block": funktionsnedsaettelse_block,
        "hjaelpemidler": hjaelpemidler,
        "antal_filer": antal_filer
    }


def match_regler(hjaelpemidler_text: str, regler: list[dict]) -> dict[str, list[str]]:
    """Match hjaelpemidler free text against keyword lists per paragraph column.
    Returns a dict mapping matched paragraph names to the keywords that were found."""
    if not hjaelpemidler_text:
        return {}

    text_lower = hjaelpemidler_text.lower()
    paragraphs = regler[0].keys() if regler else []
    matched = {}

    for paragraph in paragraphs:
        keywords = [row[paragraph].strip().lower() for row in regler if row.get(paragraph, "").strip()]
        found = [keyword for keyword in keywords if re.search(rf"\b{re.escape(keyword)}\b", text_lower)]
        if found:
            matched[paragraph] = found

    return matched

def søg_borger(cpr: str, telefonnummer: str = None) -> dict:
    """Søg efter borger i Nexus ved CPR-nummer. Returner borgerdata som dict."""
    if not cpr:
        return None

    try:
        findes_borger = nexus.borgere.søg_borgere(cpr=cpr)
        if not findes_borger:
            # Hvis borger ikke findes, så opret i nexus
            borger = nexus.borgere.opret_borger(cpr)
            # Opdater telefonnummer på borgeren
            prototype = nexus.nexus_client.get(borger["_links"]["self"]["href"])
            prototype["homeTelephone"] = telefonnummer
            nexus.nexus_client.put(borger["_links"]["update"]["href"], json=prototype)

        borger = nexus.borgere.hent_borger(cpr)
        return borger
    except Exception as e:
        logging.error(f"Error searching for borger with CPR {cpr}: {e}")
        return None

def opret_forløb(borger: dict) -> None:
    """Opret forløb i Nexus for given borger."""
    
    nexus.forløb.opret_forløb(borger, "Ældre og sundhedsfagligt grundforløb", "Sag SOFF: Afgørelse - Lov om social service")
    nexus.forløb.opret_forløb(borger, "Ældre og sundhedsfagligt grundforløb", "Sag SOFF: Hjælpemidler, forbrugsgoder, boligindretning m.v.")

def opret_skema(borger: dict, ansøgning: dict, matched_paragraffer: dict) -> None:
    """Opret skema i Nexus for given borger baseret på ansøgning og matchede paragraffer."""
    for matched_paragraph in matched_paragraffer:
        print({', '.join(matched_paragraffer[matched_paragraph])})
        # Formatter dato
        dt = datetime.now(ZoneInfo("Europe/Copenhagen")).replace(
            hour=0, minute=0, second=0, microsecond=0
        )
        dato = dt
        nexus.skemaer.opret_komplet_skema(
            borger=borger,
            skematype_navn="Henvendelse - Hjælpemidler, forbrugsgoder, boligindretning m.v.",
            grundforløb="Ældre og sundhedsfagligt grundforløb",
            forløb="Sag SOFF: Hjælpemidler, forbrugsgoder, boligindretning m.v.",
            handling_navn="Udfyldt",
            data = {
                "Henvendelse modtaget" : dato,
                "Ansvarlig myndighedsorganisation" : "Indgangen",
                "Kilde som henvendelses kommer fra" : "Borger",
                "Er borgeren indforstået med henvendelsen?" : "Ja",
                "Henvendelsesårsag" : (
                    f"Fundne følgende hjælpemidler: {', '.join(matched_paragraffer[matched_paragraph])}\n"
                    f"Fundet antal filer i mail: {ansøgning['antal_filer']}\n"
                    f"{datetime.now().date().strftime('%d-%m-%Y')} //Robotten Tyra\n"
                    f"{ansøgning['funktionsnedsaettelse_block']}"
                )
            }
        )


async def populate_queue(workqueue: Workqueue):
    logger = logging.getLogger(__name__)

    logger.info("Hello from populate workqueue!")
    mails = await mail_service.check_inbox_messages(limit=20)
    # Fjern alle mails der ikke er fra xflow eller hjælpemidler (#2 skal slettes før produktion)
    mails = [mail for mail in mails if mail["from_address"].lower() in {"xflow@odense.dk", "hjaelpemidler@odense.dk"}]

    for mail in mails:
        if "Ansøgning om hjælpemiddel, forbrugsgode eller boligindretning" not in mail["subject"]:
            continue

        if "RPA TEST" not in mail["body_preview"]:
            continue
    
        workqueue.add_item(
            data= {"id": mail["id"]},
            reference=mail["id"]
        )
        


async def process_workqueue(workqueue: Workqueue):
    logger = logging.getLogger(__name__)

    logger.info("Hello from process workqueue!")

    for item in workqueue:
        with item:
            data = item.data  # Item data deserialized from json as dict
 
            try:
                # Process the item here
                attachments = await mail_service.list_attachments("hjaelpemidler@odense.dk", data["id"])

                target_name = "Ansoegning_om_hjaelpemiddel_forbrugsgode_eller_boligindretning.pdf"
                pdf_attachment = next(
                    (a for a in attachments if a[0] == target_name), None
                )

                if pdf_attachment is None:
                    raise WorkItemError(f"Attachment '{target_name}' not found in message {data['id']}")

                _, pdf_path, _ = pdf_attachment
                with fitz.open(pdf_path) as pdf:
                    pdf_text = "\n".join(page.get_text() for page in pdf)

                ansoegning = parse_ansoegning(pdf_text, attachments)
                matched_paragraffer = match_regler(ansoegning["hjaelpemidler"], regler)
                logger.info(f"Parsed ansøgning: {ansoegning}")
                logger.info(f"Matched paragraffer: {matched_paragraffer}")

                # Søg efter borger i Nexus ved CPR-nummer. Hvis borger ikke findes, så opret i nexus
                # borger = søg_borger(ansoegning["cpr"], ansoegning["telefonnummer"])

                # Opret forløb
                # opret_forløb(borger)
                borger = nexus.borgere.hent_borger(os.environ.get("TEST_CPR"))  # TODO: Fjern test CPR og hent rigtigt fra søg_borger

                # Opret skema
                opret_skema(borger, ansoegning, matched_paragraffer)

            except Exception as e:
                logger.error(f"Error processing item: {data}. Error: {e}")
                item.fail(str(e))


async def main():
    global tracker, mail_service, nexus, regler, forløb

    ats = AutomationServer.from_environment()
    workqueue = ats.workqueue()

    # Initialize external systems for automation here..
    tracking_credential = Credential.get_credential("Odense SQL Server")
    roboa_credential = Credential.get_credential("RoboA") # bruges til at hente emails
    nexus_credential = Credential.get_credential("KMD Nexus - produktion")

    tracker = Tracker(
        username=tracking_credential.username,
        password=tracking_credential.password
    )

    nexus = NexusClientManager(
        client_id=nexus_credential.username,
        client_secret=nexus_credential.password,
        instance=nexus_credential.data["instance"],
        timeout=60,
    )

    # Parse command line arguments
    parser = argparse.ArgumentParser(description=proces_navn)
    parser.add_argument(
        "--excel-file",
        default=os.environ.get("EXCEL_MAPPING_PATH"),
        help="Path to the Excel file containing mapping data (default: ./Regelsæt.xlsx)",
    )
    parser.add_argument(
        "--queue",
        action="store_true",
        help="Populate the queue with test data and exit",
    )
    args = parser.parse_args()

    # Validate Excel files exists (skip validation for Windows paths on Linux)
    def is_windows_path(path: str) -> bool:
        """Check if path is a Windows path (has drive letter or UNC path)"""
        return (
            (len(path) > 1 and path[1] == ":")
            or path.startswith("\\\\")
            or path.startswith("//")
        )

    # Load excel mapping data once on startup (only if files exist on current system)
    if os.path.isfile(args.excel_file):
        load_excel_mapping(args.excel_file)
    elif not is_windows_path(args.excel_file):
        raise FileNotFoundError(f"Excel file not found: {args.excel_file}")

    regler = get_excel_mapping().get("Placeringer", [])
    forløb = get_excel_mapping().get("Forløb", [])

    # Initialize mail service (async)
    mail_service = MailService(roboa_credential)
    await mail_service.initialize()

    # Queue management
    if args.queue:
        workqueue.clear_workqueue(WorkItemStatus.NEW)
        await populate_queue(workqueue)
        return

    # Process workqueue
    await process_workqueue(workqueue)


if __name__ == "__main__":
    asyncio.run(main())
