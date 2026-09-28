from textwrap import shorten
from urllib.parse import quote


class OutlookClient:

    def __init__(self, graph_client):
        """
        Parameters
        ----------
        graph_client : GraphClient
            Instance of the GraphClient class.
        """
        self.graph = graph_client

    ###########################################################################
    # Send Email
    ###########################################################################

    def send_email(
        self,
        to_recipients,
        subject,
        body,
        body_type="HTML",
        cc_recipients=None,
        bcc_recipients=None,
        save_to_sent=True
    ):
        """
        Send an email.

        Parameters
        ----------
        to_recipients : str or list
        subject : str
        body : str
        body_type : str
            HTML or Text
        cc_recipients : list
        bcc_recipients : list
        """

        if isinstance(to_recipients, str):
            to_recipients = [to_recipients]

        cc_recipients = cc_recipients or []
        bcc_recipients = bcc_recipients or []

        payload = {
            "message": {
                "subject": subject,
                "body": {
                    "contentType": body_type,
                    "content": body
                },
                "toRecipients": [
                    {
                        "emailAddress": {
                            "address": email
                        }
                    }
                    for email in to_recipients
                ],
                "ccRecipients": [
                    {
                        "emailAddress": {
                            "address": email
                        }
                    }
                    for email in cc_recipients
                ],
                "bccRecipients": [
                    {
                        "emailAddress": {
                            "address": email
                        }
                    }
                    for email in bcc_recipients
                ]
            },
            "saveToSentItems": save_to_sent
        }

        self.graph.post(
            "me/sendMail",
            payload
        )

        return True

    ###########################################################################
    # Mail Folder Functions
    ###########################################################################

    def get_mail_folders(self):
        """
        Return all mail folders.
        """

        return self.graph.get_all(
            "me/mailFolders"
        )

    def get_folder_id(self, folder_path):
        """
        Folder path examples:

        Inbox
        Inbox/Processed
        Inbox/Watch Office/Processed
        """

        parts = folder_path.split("/")

        parent_id = None

        for part in parts:

            if parent_id is None:

                folders = self.graph.get_all(
                    "me/mailFolders"
                )

            else:

                folders = self.graph.get_all(
                    f"me/mailFolders/{parent_id}/childFolders"
                )

            match = next(
                (
                    f for f in folders
                    if f["displayName"].lower()
                    == part.lower()
                ),
                None
            )

            if not match:
                raise ValueError(
                    f"Folder path not found: {folder_path}"
                )

            parent_id = match["id"]

        return parent_id

    ###########################################################################
    # Search Messages
    ###########################################################################

    def search_messages(
        self,
        folder_name="Inbox",
        subject_contains=None,
        sender_contains=None,
        top=100
    ):
        """
        Search messages within a folder.
        """

        folder_id = self.get_folder_id(folder_name)

        endpoint = (
            f"me/mailFolders/{folder_id}/messages"
            f"?$top={top}"
        )

        messages = self.graph.get_all(endpoint)

        results = []

        for message in messages:

            subject = (
                message.get("subject", "")
            )

            sender = (
                message
                .get("from", {})
                .get("emailAddress", {})
                .get("address", "")
            )

            include = True

            if subject_contains:
                include = (
                    subject_contains.lower()
                    in subject.lower()
                )

            if sender_contains:
                include = include and (
                    sender_contains.lower()
                    in sender.lower()
                )

            if include:
                results.append(message)

        return results

    ###########################################################################
    # Move Email
    ###########################################################################

    def move_message(
        self,
        message_id,
        destination_folder
    ):
        """
        Move a message to another folder.
        """

        destination_folder_id = (
            self.get_folder_id(
                destination_folder
            )
        )

        payload = {
            "destinationId":
                destination_folder_id
        }

        return self.graph.post(
            f"me/messages/{message_id}/move",
            payload
        )

    ###########################################################################
    # Bulk Move
    ###########################################################################

    def move_messages(
        self,
        message_ids,
        destination_folder
    ):
        """
        Move multiple emails.
        """

        results = []

        for message_id in message_ids:

            result = self.move_message(
                message_id,
                destination_folder
            )

            results.append(result)

        return results

    ###########################################################################
    # Get Folder Messages
    ###########################################################################

    def get_folder_messages(
        self,
        folder_name="Inbox",
        top=100
    ):
        """
        Get messages from a folder.
        """

        folder_id = self.get_folder_id(
            folder_name
        )

        return self.graph.get_all(
            f"me/mailFolders/{folder_id}/messages?$top={top}"
        )

    ###########################################################################
    # Get Unread Messages
    ###########################################################################

    def get_unread_messages(
        self,
        folder_name="Inbox"
    ):
        """
        Get unread messages from folder.
        """

        folder_id = self.get_folder_id(
            folder_name
        )

        endpoint = (
            f"me/mailFolders/{folder_id}/messages"
            f"?$filter=isRead eq false"
        )

        return self.graph.get_all(endpoint)

    ###########################################################################
    # Mark Message Read
    ###########################################################################

    def mark_as_read(
        self,
        message_id
    ):

        return self.graph.patch(
            f"me/messages/{message_id}",
            {"isRead": True}
        )

    ###########################################################################
    # Mark Message Unread
    ###########################################################################

    def mark_as_unread(
        self,
        message_id
    ):

        return self.graph.patch(
            f"me/messages/{message_id}",
            {"isRead": False}
        )


    from textwrap import shorten

    ###########################################################################
    # Print Message
    ###########################################################################
    
    def print_message(self, message, body_preview_length=500):
        """
        Pretty-print a Microsoft Graph email message.

        Parameters
        ----------
        message : dict
            Message object returned by Graph API.

        body_preview_length : int
            Maximum length of body preview.
        """

        if not isinstance(message, dict):
            raise TypeError(
                f"Expected Graph message dict, got {type(message)}"
            )

        sender = (
            message.get("from", {})
                .get("emailAddress", {})
                .get("address", "")
        )

        sender_name = (
            message.get("from", {})
                .get("emailAddress", {})
                .get("name", "")
        )

        recipients = ", ".join(
            r.get("emailAddress", {}).get("address", "")
            for r in message.get("toRecipients", [])
        )

        cc = ", ".join(
            r.get("emailAddress", {}).get("address", "")
            for r in message.get("ccRecipients", [])
        )

        body = (
            message.get("body", {})
                .get("content", "")
                .replace("\r", "")
        )

        print("\n" + "=" * 100)
        print(f"Subject     : {message.get('subject', '')}")
        print(f"From        : {sender_name} <{sender}>")
        print(f"To          : {recipients}")
        print(f"CC          : {cc}")
        print(f"Received    : {message.get('receivedDateTime', '')}")
        print(f"Sent        : {message.get('sentDateTime', '')}")
        print(f"Message ID  : {message.get('internetMessageId', '')}")
        print(f"Read        : {message.get('isRead', False)}")
        print(f"Importance  : {message.get('importance', '')}")

        if message.get("hasAttachments"):
            print("Attachments : Yes")
        else:
            print("Attachments : No")

        print("-" * 100)
        print("BODY PREVIEW")
        print("-" * 100)

        print(
            shorten(
                body,
                width=body_preview_length,
                placeholder="..."
            )
        )

        print("=" * 100)