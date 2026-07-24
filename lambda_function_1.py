import base64
import json
import logging
import os
from urllib.error import HTTPError, URLError
from urllib.parse import quote, urlencode
from urllib.request import Request, urlopen

import boto3


logger = logging.getLogger()
logger.setLevel(logging.INFO)

s3 = boto3.client("s3")

# Lambda environment variables
TENANT_ID = os.environ["TENANT_ID"]
CLIENT_ID = os.environ["CLIENT_ID"]
CLIENT_SECRET = os.environ["CLIENT_SECRET"]
USER_EMAIL = os.environ["USER_EMAIL"]
BUCKET_NAME = os.environ["S3_BUCKET"]

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
HTTP_TIMEOUT_SECONDS = 30
MESSAGE_PAGE_SIZE = 50
ATTACHMENT_PAGE_SIZE = 100


def make_http_request(
    url,
    method="GET",
    headers=None,
    data=None,
):
    """
    Make a JSON HTTP request using Python's built-in urllib.

    Returns:
        tuple: (status_code, response_json)
    """
    request_headers = headers or {}

    if data is not None:
        if isinstance(data, dict):
            if request_headers.get("Content-Type") == "application/json":
                request_body = json.dumps(data).encode("utf-8")
            else:
                request_body = urlencode(data).encode("utf-8")
        elif isinstance(data, str):
            request_body = data.encode("utf-8")
        else:
            request_body = data
    else:
        request_body = None

    request = Request(
        url=url,
        data=request_body,
        headers=request_headers,
        method=method,
    )

    try:
        with urlopen(
            request,
            timeout=HTTP_TIMEOUT_SECONDS,
        ) as response:
            response_body = response.read().decode("utf-8")

            if response_body:
                return response.status, json.loads(response_body)

            return response.status, {}

    except HTTPError as error:
        error_body = error.read().decode("utf-8", errors="replace")

        logger.error(
            "HTTP request failed: method=%s status=%s url=%s body=%s",
            method,
            error.code,
            url,
            error_body,
        )

        raise RuntimeError(
            f"HTTP {error.code} calling {url}: {error_body}"
        ) from error

    except URLError as error:
        logger.error(
            "Network request failed: method=%s url=%s error=%s",
            method,
            url,
            error.reason,
        )

        raise RuntimeError(
            f"Network error calling {url}: {error.reason}"
        ) from error


def get_graph_token():
    token_url = (
        f"https://login.microsoftonline.com/"
        f"{quote(TENANT_ID, safe='')}/oauth2/v2.0/token"
    )

    token_payload = {
        "client_id": CLIENT_ID,
        "scope": "https://graph.microsoft.com/.default",
        "client_secret": CLIENT_SECRET,
        "grant_type": "client_credentials",
    }

    _, token_response = make_http_request(
        url=token_url,
        method="POST",
        headers={
            "Content-Type": "application/x-www-form-urlencoded",
            "Accept": "application/json",
        },
        data=token_payload,
    )

    access_token = token_response.get("access_token")

    if not access_token:
        error_description = token_response.get(
            "error_description",
            "Microsoft did not return an access token.",
        )
        raise RuntimeError(error_description)

    return access_token


def list_unread_inbox_messages(headers, encoded_user):
    """
    Retrieve all unread messages from the Inbox before any are marked read.

    Fetching all pages first avoids changing the filtered result set while
    Microsoft Graph pagination is still in progress.
    """
    endpoint = (
        f"{GRAPH_BASE_URL}/users/{encoded_user}/mailFolders/inbox/messages"
        "?$filter=isRead%20eq%20false"
        "&$select=id,internetMessageId,receivedDateTime,subject,"
        "bodyPreview,from,hasAttachments"
        f"&$top={MESSAGE_PAGE_SIZE}"
    )

    messages = []

    while endpoint:
        _, graph_response = make_http_request(
            url=endpoint,
            method="GET",
            headers=headers,
        )

        messages.extend(graph_response.get("value", []))
        endpoint = graph_response.get("@odata.nextLink")

    return messages


def list_message_attachments(
    message_id,
    headers,
    encoded_user,
):
    """Retrieve all attachments for one message."""
    encoded_message_id = quote(message_id, safe="")

    endpoint = (
        f"{GRAPH_BASE_URL}/users/{encoded_user}/messages/"
        f"{encoded_message_id}/attachments"
        f"?$top={ATTACHMENT_PAGE_SIZE}"
    )

    attachments = []

    while endpoint:
        _, graph_response = make_http_request(
            url=endpoint,
            method="GET",
            headers=headers,
        )

        attachments.extend(graph_response.get("value", []))
        endpoint = graph_response.get("@odata.nextLink")

    return attachments


def sanitize_filename(filename):
    """Prevent attachment filenames from creating unintended S3 paths."""
    cleaned = filename.replace("\\", "_").replace("/", "_").strip()

    if not cleaned:
        return "attachment.pdf"

    return cleaned


def upload_pdf(
    message_id,
    attachment,
    sender,
):
    filename = sanitize_filename(
        attachment.get("name") or "attachment.pdf"
    )

    content_bytes = attachment.get("contentBytes")

    if not content_bytes:
        raise RuntimeError(
            f"Attachment {filename} does not contain contentBytes."
        )

    try:
        pdf_bytes = base64.b64decode(
            content_bytes,
            validate=True,
        )
    except (ValueError, TypeError) as error:
        raise RuntimeError(
            f"Attachment {filename} contains invalid base64 data."
        ) from error

    pdf_key = f"incoming/{message_id}/attachments/{filename}"

    s3.put_object(
        Bucket=BUCKET_NAME,
        Key=pdf_key,
        Body=pdf_bytes,
        ContentType="application/pdf",
        Metadata={
            "message-id": message_id[:1024],
            "sender": sender[:1024],
        },
    )

    return {
        "attachment_id": attachment.get("id"),
        "filename": filename,
        "content_type": attachment.get("contentType"),
        "s3_key": pdf_key,
        "size_bytes": len(pdf_bytes),
    }


def upload_metadata(
    message,
    sender,
    sender_name,
    discovered_attachment_count,
    uploaded_attachments,
):
    """Store one JSON object for every successfully processed email."""
    message_id = message["id"]

    metadata = {
        "message_id": message_id,
        "internet_message_id": message.get("internetMessageId"),
        "received_date_time": message.get("receivedDateTime"),
        "subject": message.get("subject") or "",
        "body_preview": message.get("bodyPreview") or "",
        "sender": sender,
        "sender_name": sender_name,
        "has_attachments": bool(message.get("hasAttachments", False)),
        "discovered_attachment_count": discovered_attachment_count,
        "uploaded_pdf_count": len(uploaded_attachments),
        # This key is retained for compatibility with the existing JSON schema.
        "attachments": uploaded_attachments,
    }

    metadata_key = f"incoming/{message_id}/metadata.json"

    s3.put_object(
        Bucket=BUCKET_NAME,
        Key=metadata_key,
        Body=json.dumps(
            metadata,
            ensure_ascii=False,
            indent=2,
        ).encode("utf-8"),
        ContentType="application/json",
    )

    return metadata_key


def mark_message_as_read(
    message_id,
    headers,
    encoded_user,
):
    encoded_message_id = quote(message_id, safe="")

    patch_url = (
        f"{GRAPH_BASE_URL}/users/{encoded_user}/messages/"
        f"{encoded_message_id}"
    )

    make_http_request(
        url=patch_url,
        method="PATCH",
        headers={
            **headers,
            "Content-Type": "application/json",
        },
        data={
            "isRead": True,
        },
    )


def lambda_handler(event, context):
    processed_messages = 0
    messages_with_uploaded_pdfs = 0
    messages_without_uploaded_pdfs = 0
    uploaded_pdfs = 0
    skipped_messages = 0
    failed_messages = 0

    token = get_graph_token()

    graph_headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
    }

    encoded_user = quote(USER_EMAIL, safe="")

    # Load every unread Inbox message, including messages with no attachments.
    messages = list_unread_inbox_messages(
        headers=graph_headers,
        encoded_user=encoded_user,
    )

    logger.info("Found %s unread Inbox message(s).", len(messages))

    for message in messages:
        message_id = message.get("id")

        if not message_id:
            logger.warning("Skipping message with no ID.")
            skipped_messages += 1
            continue

        sender_email_address = (
            message.get("from", {}).get("emailAddress", {})
        )
        sender = sender_email_address.get("address", "")
        sender_name = sender_email_address.get("name", "")

        uploaded_attachments = []
        discovered_attachments = []

        try:
            # Only request the attachment collection when Graph reports that
            # the message has non-inline attachments.
            if message.get("hasAttachments", False):
                discovered_attachments = list_message_attachments(
                    message_id=message_id,
                    headers=graph_headers,
                    encoded_user=encoded_user,
                )

            for attachment in discovered_attachments:
                attachment_type = attachment.get("@odata.type")
                filename = attachment.get("name") or ""

                # Preserve the existing behavior: upload only normal PDF files.
                if attachment_type != "#microsoft.graph.fileAttachment":
                    continue

                # Ignore inline images such as email signatures.
                if attachment.get("isInline", False):
                    continue

                if not filename.lower().endswith(".pdf"):
                    continue

                uploaded_attachment = upload_pdf(
                    message_id=message_id,
                    attachment=attachment,
                    sender=sender,
                )

                uploaded_attachments.append(uploaded_attachment)
                uploaded_pdfs += 1

            # This is intentionally unconditional. Every email receives a
            # metadata.json object, even when attachments is an empty list.
            metadata_key = upload_metadata(
                message=message,
                sender=sender,
                sender_name=sender_name,
                discovered_attachment_count=len(discovered_attachments),
                uploaded_attachments=uploaded_attachments,
            )

            # Mark as read only after the JSON and all eligible PDFs succeed.
            mark_message_as_read(
                message_id=message_id,
                headers=graph_headers,
                encoded_user=encoded_user,
            )

            processed_messages += 1

            if uploaded_attachments:
                messages_with_uploaded_pdfs += 1
            else:
                messages_without_uploaded_pdfs += 1

            logger.info(
                "Processed message %s: metadata=%s uploaded_pdfs=%s.",
                message_id,
                metadata_key,
                len(uploaded_attachments),
            )

        except Exception:
            failed_messages += 1

            logger.exception(
                "Failed to process message %s. "
                "The message will remain unread.",
                message_id,
            )

    result = {
        "discovered_unread_messages": len(messages),
        "processed_messages": processed_messages,
        "messages_with_uploaded_pdfs": messages_with_uploaded_pdfs,
        "messages_without_uploaded_pdfs": messages_without_uploaded_pdfs,
        "uploaded_pdfs": uploaded_pdfs,
        "skipped_messages": skipped_messages,
        "failed_messages": failed_messages,
    }

    logger.info("Email ingestion completed: %s", result)

    return {
        "statusCode": 200 if failed_messages == 0 else 207,
        "headers": {
            "Content-Type": "application/json",
        },
        "body": json.dumps(result),
    }
