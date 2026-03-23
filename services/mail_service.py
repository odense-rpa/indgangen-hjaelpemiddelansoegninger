"""
Mail service utilities for working with Microsoft Graph API.
"""
import base64
import logging
import re
import tempfile
import time
from typing import Any, Dict, List, Optional, Tuple

import aiofiles
from azure.identity import UsernamePasswordCredential
from bs4 import BeautifulSoup
from msgraph.graph_service_client import GraphServiceClient

from automation_server_client import Credential

logger = logging.getLogger(__name__)


def extract_text_from_html(html_content: str) -> str:
    """Extract plain text from HTML content."""
    soup = BeautifulSoup(html_content, 'html.parser')
    # Get text and clean up extra whitespace
    text = soup.get_text(separator='\n', strip=True)
    # Remove excessive blank lines
    lines = [line for line in text.split('\n') if line.strip()]
    return '\n'.join(lines)


def parse_email_data(text_content: str) -> dict:
    """
    Extract relevant data fields from email text.
    
    Returns:
        Dictionary with extracted fields
    """
    data = {}
        
    # Extract Indsendt dato
    dato_match = re.search(r'Indsendt dato:\s*([\d-]+)', text_content, re.IGNORECASE)
    if dato_match:
        data['indsendt_dato'] = dato_match.group(1)
    
    # Extract CPR-nr
    cpr_match = re.search(r'CPR-nr\.?:\s*(\d{10})', text_content, re.IGNORECASE)
    if cpr_match:
        data['cpr_nr'] = cpr_match.group(1)
    
    # Extract location (Hvor er barnet i hverdagen)
    location_match = re.search(r'Hvor er barnet i hverdagen:\s*(.+?)(?:\n|$)', text_content, re.IGNORECASE)
    if location_match:
        data['lokation'] = location_match.group(1).strip()

    # Extract navn
    navn_match = re.search(r'Navn:\s*(.+?)(?:\n|$)', text_content, re.IGNORECASE)
    if navn_match:
        data['navn'] = navn_match.group(1).strip()
    
    return data


class MailService:
    """Utility service for email operations using Microsoft Graph API."""

    def __init__(self, roboc_credential: Credential):
        """
        Initialize the mail service.

        Args:
            roboc_credential: Automation Server credential for email access
        """
        self.credential_obj = roboc_credential
        self.graph_client: Optional[GraphServiceClient] = None
        self.credential: Optional[UsernamePasswordCredential] = None

    async def initialize(self) -> None:
        """Initialize the Microsoft Graph client."""
        await self._initialize_graph_client()

    async def _initialize_graph_client(self) -> None:
        """Initialize Microsoft Graph client with username/password credential."""

        try:
            # Extract Graph API configuration from credential data
            credential_data = self.credential_obj.data if hasattr(self.credential_obj, 'data') else {}
            
            logger.debug(f"Credential data keys: {credential_data.keys()}")
            logger.debug(f"Username: {self.credential_obj.username}")
            logger.debug(f"Tenant ID: {credential_data.get('tenant_id', 'NOT SET')}")
            logger.debug(f"Client ID: {credential_data.get('client_id', 'NOT SET')}")
            
            # Create credential for delegated authentication
            username_with_domain = f"{self.credential_obj.username}@odense.dk"
            self.credential = UsernamePasswordCredential(
                client_id=credential_data.get('client_id', ''),
                username=username_with_domain,
                password=self.credential_obj.password,
                tenant_id=credential_data.get('tenant_id', '')
            )

            # Create Graph service client
            scope = credential_data.get('graph_scope', 'https://graph.microsoft.com/.default')
            self.graph_client = GraphServiceClient(
                credentials=self.credential,
                scopes=[scope]
            )

            # Test authentication by getting user info
            await self._test_authentication()

        except Exception as e:
            logger.error(f"Failed to initialize Graph client: {e}")
            raise

    async def _test_authentication(self) -> None:
        """Test authentication by making a simple Graph API call."""
        try:

            # Get current user to verify authentication works
            user = await self.graph_client.me.get()
            if user and user.display_name:
                logger.debug(
                    f"Authenticated successfully as: {user.display_name} ({user.user_principal_name})")
            else:
                raise Exception(
                    "Authentication test failed - no user data returned")

        except Exception as e:
            logger.error(f"Authentication test failed: {e}")
            raise Exception(
                f"Failed to authenticate with Microsoft Graph: {e}") from e

    def _is_personal_mailbox(self, mailbox_address: str) -> bool:
        """Check if the mailbox address is the personal mailbox."""
        username_with_domain = f"{self.credential_obj.username}@odense.dk"
        return mailbox_address.lower() == username_with_domain.lower()

    def _get_messages_request_builder(self, mailbox_address: str):
        """Get the appropriate messages request builder for personal or shared mailbox."""
        if self._is_personal_mailbox(mailbox_address):
            return self.graph_client.me.messages
        return self.graph_client.users.by_user_id(mailbox_address).messages

    def _extract_message_info(self, msg) -> Dict[str, Any]:
        """Extract message information into a standardized dictionary."""
        # Extract sender information
        from_address = "Unknown"
        from_name = "Unknown"
        if hasattr(msg, 'from_') and msg.from_ and msg.from_.email_address:
            from_address = msg.from_.email_address.address
            from_name = msg.from_.email_address.name or from_address

        return {
            'id': msg.id,
            'internet_message_id': getattr(msg, 'internet_message_id', None),
            'subject': msg.subject or "(No subject)",
            'from_address': from_address,
            'from_name': from_name,
            'received_date_time': msg.received_date_time,
            'is_read': msg.is_read,
            'importance': getattr(msg, 'importance', 'normal'),
            'has_attachments': getattr(msg, 'has_attachments', False),
            'body_preview': getattr(msg, 'body_preview', '')[:200] if getattr(msg, 'body_preview', '') else ''
        }

    async def _find_folder_by_name(self, mailbox_address: str, folder_name: str) -> Optional[Dict[str, Any]]:
        """Find a folder by name in a mailbox."""
        folders = await self.list_shared_mailbox_folders(mailbox_address)
        for folder in folders:
            if folder['display_name'].lower() == folder_name.lower():
                return folder
        return None

    async def check_inbox_messages(self, limit: int = 50) -> List[Dict[str, Any]]:
        """
        Check the user's inbox for messages.

        Args:
            limit: Maximum number of messages to retrieve

        Returns:
            List of message information dictionaries
        """
        # Delegate to get_shared_mailbox_messages with personal mailbox address
        return await self.get_shared_mailbox_messages(
            mailbox_address="hjaelpemidler@odense.dk",
            folder_name="Inbox",
            limit=limit,
            unread_only=False
        )

    async def get_inbox_subfolders(self) -> List[Dict[str, Any]]:
        """
        Get all subfolders under the Inbox (Indbakke) folder only.

        Returns:
            List of subfolder information dictionaries with nested subfolders
        """
        if not self.graph_client:
            raise Exception("Graph client not initialized")

        try:
            logger.info(f"Listing subfolders under Inbox for {self.credential_obj.username}")

            all_folders = await self.list_shared_mailbox_folders(self.credential_obj.username)
            
            # Find the Inbox folder (handles both English and Danish names)
            inbox_folder = None
            for folder in all_folders:
                if folder['display_name'].lower() in ['inbox', 'indbakke']:
                    inbox_folder = folder
                    break

            if not inbox_folder:
                logger.warning("Inbox folder not found")
                return []

            logger.info(f"Found {len(inbox_folder['subfolders'])} subfolders under Inbox")
            return inbox_folder['subfolders']

        except Exception as e:
            logger.error(f"Error getting inbox subfolders: {e}")
            raise

    async def _list_subfolders_recursive(self, mailbox_address: str, parent_folder_id: Optional[str] = None, depth: int = 0) -> List[Dict[str, Any]]:
        """
        Recursively list all subfolders within a folder.

        Args:
            mailbox_address: Email address of the mailbox
            parent_folder_id: ID of the parent folder (None for root)
            depth: Current recursion depth for logging

        Returns:
            List of folder information dictionaries with nested structure
        """
        folder_list = []
        
        try:
            if parent_folder_id is None:
                # Get root level folders
                if self._is_personal_mailbox(mailbox_address):
                    folders = await self.graph_client.me.mail_folders.get()
                else:
                    folders = await self.graph_client.users.by_user_id(mailbox_address).mail_folders.get()
            else:
                # Get subfolders of a specific folder
                if self._is_personal_mailbox(mailbox_address):
                    folders = await self.graph_client.me.mail_folders.by_mail_folder_id(parent_folder_id).child_folders.get()
                else:
                    folders = await self.graph_client.users.by_user_id(mailbox_address).mail_folders.by_mail_folder_id(parent_folder_id).child_folders.get()

            if folders and folders.value:
                for folder in folders.value:
                    folder_info = {
                        'id': folder.id,
                        'display_name': folder.display_name,
                        'total_item_count': folder.total_item_count,
                        'unread_item_count': folder.unread_item_count,
                        'child_folder_count': folder.child_folder_count,
                        'subfolders': []
                    }

                    # Recursively get subfolders if this folder has children
                    if folder.child_folder_count and folder.child_folder_count > 0:
                        subfolders = await self._list_subfolders_recursive(
                            mailbox_address=mailbox_address,
                            parent_folder_id=folder.id,
                            depth=depth + 1
                        )
                        folder_info['subfolders'] = subfolders

                    folder_list.append(folder_info)

            return folder_list

        except Exception as e:
            logger.error(f"Error listing subfolders at depth {depth} for {mailbox_address}: {e}")
            raise

    async def list_shared_mailbox_folders(self, mailbox_address: str) -> List[Dict[str, Any]]:
        """
        List all folders in a mailbox (personal or shared), including nested subfolders.

        Args:
            mailbox_address: Email address of the mailbox

        Returns:
            List of folder information dictionaries with nested subfolders
        """
        if not self.graph_client:
            raise Exception("Graph client not initialized")

        try:
            logger.info(
                f"Listing folders for mailbox: {mailbox_address}")

            folder_list = await self._list_subfolders_recursive(
                mailbox_address=mailbox_address,
                parent_folder_id=None
            )

            logger.info(
                f"Found {len(folder_list)} top-level folders in {mailbox_address}")

            return folder_list

        except Exception as e:
            logger.error(f"Error listing folders for {mailbox_address}: {e}")
            raise

    async def get_shared_mailbox_messages(self, mailbox_address: str, folder_name: str = "Inbox",
                                          limit: int = 50, unread_only: bool = False) -> List[Dict[str, Any]]:
        """
        Get messages from a shared mailbox folder with reasonable limits.

        Args:
            mailbox_address: Email address of the shared mailbox
            folder_name: Name of the folder (default: "Inbox")
            limit: Maximum number of messages to retrieve (default: 50, max: 100)
            unread_only: If True, only get unread messages

        Returns:
            List of message information dictionaries
        """
        if not self.graph_client:
            raise Exception("Graph client not initialized")

        # Enforce reasonable limits
        limit = min(limit, 100)  # Max 100 messages

        try:
            logger.debug(
                f"Getting messages from {mailbox_address} folder '{folder_name}' (limit: {limit})")

            messages = None

            # Simplified approach - get messages directly from mailbox for inbox
            if folder_name.lower() in ["inbox", "indbakke"]:
                from msgraph.generated.users.item.messages.messages_request_builder import MessagesRequestBuilder
                request_config = MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
                    query_parameters=MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
                        top=limit
                    )
                )
                # Get messages from the Inbox folder specifically
                if self._is_personal_mailbox(mailbox_address):
                    messages = await self.graph_client.me.mail_folders.by_mail_folder_id("inbox").messages.get(request_configuration=request_config)
                else:
                    messages = await self.graph_client.users.by_user_id(mailbox_address).mail_folders.by_mail_folder_id("inbox").messages.get(request_configuration=request_config)
            else:
                # Find the specific folder first
                target_folder = await self._find_folder_by_name(mailbox_address, folder_name)
                if target_folder:
                    # Use appropriate request builder for personal vs shared
                    from msgraph.generated.users.item.mail_folders.item.messages.messages_request_builder import MessagesRequestBuilder
                    request_config = MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
                        query_parameters=MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
                            top=limit
                        )
                    )
                    if self._is_personal_mailbox(mailbox_address):
                        messages = await self.graph_client.me.mail_folders.by_mail_folder_id(target_folder['id']).messages.get(request_configuration=request_config)
                    else:
                        messages = await self.graph_client.users.by_user_id(mailbox_address).mail_folders.by_mail_folder_id(target_folder['id']).messages.get(request_configuration=request_config)
                else:
                    logger.warning(
                        f"Folder '{folder_name}' not found in {mailbox_address}")
                    return []

            message_list = []
            if messages and messages.value:
                # Apply filtering and limiting
                filtered_messages = messages.value

                # Filter for unread if requested
                if unread_only:
                    filtered_messages = [
                        msg for msg in filtered_messages if not msg.is_read]

                # Apply limit
                filtered_messages = filtered_messages[:limit]

                for msg in filtered_messages:
                    message_list.append(self._extract_message_info(msg))

                logger.debug(
                    f"Retrieved {len(message_list)} messages from {mailbox_address}")
            else:
                logger.info(
                    f"No messages found in {mailbox_address} folder '{folder_name}'")

            return message_list

        except Exception as e:
            logger.error(f"Error getting messages from {mailbox_address}: {e}")
            raise

    async def list_attachments(self, mailbox_address: str, message_id: str) -> List[Tuple[str, str, Dict[str, Any]]]:
        """
        List attachments for a message and save each to a temporary file.

        Args:
            mailbox_address: Email address of the mailbox
            message_id: ID of the message to get attachments from

        Returns:
            List of tuples containing (attachment_name, temp_file_path, attachment_metadata)
        """
        if not self.graph_client:
            raise Exception("Graph client not initialized")

        try:
            logger.debug(f"Getting attachments for message {message_id}")

            # Get attachments based on mailbox type
            attachments = await self._get_messages_request_builder(mailbox_address).by_message_id(message_id).attachments.get()

            attachment_files = []

            if attachments and attachments.value:
                for attachment in attachments.value:
                    try:
                        # Get attachment details
                        attachment_detail = await self._get_messages_request_builder(mailbox_address).by_message_id(message_id).attachments.by_attachment_id(attachment.id).get()

                        if attachment_detail and hasattr(attachment_detail, 'content_bytes'):
                            # Create temporary file with shorter filename
                            # Get file extension if present
                            name_parts = attachment.name.rsplit('.', 1)
                            if len(name_parts) == 2:
                                ext = '.' + name_parts[1]
                            else:
                                ext = ''
                            # Create short filename using attachment ID and extension
                            _, temp_file_path = tempfile.mkstemp(
                                suffix=ext, prefix=f"att_{attachment.id[:8]}_")

                            # Decode and save attachment content
                            content_bytes = base64.b64decode(
                                attachment_detail.content_bytes)

                            async with aiofiles.open(temp_file_path, 'wb') as f:
                                await f.write(content_bytes)

                            # Create attachment metadata
                            attachment_metadata = {
                                'id': attachment.id,
                                'name': attachment.name,
                                'size': attachment.size,
                                'content_type': attachment.content_type,
                                'is_inline': getattr(attachment, 'is_inline', False),
                                'last_modified_date_time': getattr(attachment, 'last_modified_date_time', None)
                            }

                            attachment_files.append(
                                (attachment.name, temp_file_path, attachment_metadata))
                            logger.debug(
                                f"Saved attachment '{attachment.name}' to {temp_file_path}")

                    except Exception as e:
                        logger.error(
                            f"Error processing attachment {attachment.id}: {e}")
                        continue

            logger.info(
                f"Retrieved {len(attachment_files)} attachments for message {message_id}")
            return attachment_files

        except Exception as e:
            logger.error(
                f"Error listing attachments for message {message_id}: {e}")
            return []

    async def mark_message_as_read(self, mailbox_address: str, message_id: str) -> bool:
        """
        Mark a message as read in a mailbox.

        Args:
            mailbox_address: Email address of the mailbox
            message_id: ID of the message to mark as read

        Returns:
            True if successful, False otherwise
        """
        if not self.graph_client:
            raise Exception("Graph client not initialized")

        try:
            logger.debug(
                f"Marking message {message_id} as read in {mailbox_address}")

            await self._get_messages_request_builder(mailbox_address).by_message_id(message_id).patch(
                body={"isRead": True}
            )

            logger.debug(f"Successfully marked message {message_id} as read")
            return True

        except Exception as e:
            logger.error(f"Error marking message {message_id} as read: {e}")
            return False

    async def get_message_body(self, mailbox_address: str, message_id: str) -> Dict[str, Any]:
        """
        Get the full body content of a message.

        Args:
            mailbox_address: Email address of the mailbox
            message_id: ID of the message

        Returns:
            Dictionary with 'content_type' and 'content' keys, or empty dict if error
        """
        if not self.graph_client:
            raise Exception("Graph client not initialized")

        try:
            logger.debug(f"Getting body for message {message_id}")

            message = await self._get_messages_request_builder(mailbox_address).by_message_id(message_id).get()

            if message and message.body:
                return {
                    'content_type': message.body.content_type.value if message.body.content_type else 'text',
                    'content': message.body.content or ''
                }

            return {}

        except Exception as e:
            logger.error(f"Error getting message body: {e}")
            return {}

    async def move_message(self, mailbox_address: str, message_id: str, destination_folder_id: str) -> bool:
        """
        Move a message to a different folder.

        Args:
            mailbox_address: Email address of the mailbox
            message_id: ID of the message to move
            destination_folder_id: ID of the destination folder

        Returns:
            True if successful, False otherwise
        """
        if not self.graph_client:
            raise Exception("Graph client not initialized")

        try:
            logger.debug(f"Moving message {message_id} to folder {destination_folder_id}")

            from msgraph.generated.users.item.messages.item.move.move_post_request_body import MovePostRequestBody

            body = MovePostRequestBody()
            body.destination_id = destination_folder_id

            await self._get_messages_request_builder(mailbox_address).by_message_id(message_id).move.post(body)

            logger.info(f"Successfully moved message {message_id} to folder {destination_folder_id}")
            return True

        except Exception as e:
            logger.error(f"Error moving message {message_id}: {e}")
            return False
