import asyncio
import logging
import sys
import argparse
import os

from automation_server_client import AutomationServer, Workqueue, WorkItemError, Credential, WorkItemStatus
from odk_tools.tracking import Tracker

from services.mail_service import MailService, extract_text_from_html, parse_email_data
from process.config import load_excel_mapping, get_excel_mapping
from kmd_nexus_client import NexusClientManager


tracker: Tracker = None
mail_service: MailService = None
nexus: NexusClientManager = None
proces_navn = "Indgangen - hjælpemiddelansøgninger"



async def populate_queue(workqueue: Workqueue):
    logger = logging.getLogger(__name__)

    logger.info("Hello from populate workqueue!")
    mails = await mail_service.check_inbox_messages(limit=10)
    print(f"Found {len(mails)} emails in inbox")


async def process_workqueue(workqueue: Workqueue):
    logger = logging.getLogger(__name__)

    logger.info("Hello from process workqueue!")

    for item in workqueue:
        with item:
            data = item.data  # Item data deserialized from json as dict
 
            try:
                # Process the item here
                pass
            except WorkItemError as e:
                # A WorkItemError represents a soft error that indicates the item should be passed to manual processing or a business logic fault
                logger.error(f"Error processing item: {data}. Error: {e}")
                item.fail(str(e))


async def main():
    global tracker, mail_service, nexus

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
