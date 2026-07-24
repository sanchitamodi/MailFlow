import json
import os
import urllib.parse
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

import boto3


# ---------------------------------------------------------------------------
# DESK EMAIL CONFIGURATION
# ---------------------------------------------------------------------------

DESK_EMAILS = {
    "Desk 1": ["Desk1@Team4DxHub.onmicrosoft.com"],
    "Desk 2": ["Desk2@Team4DxHub.onmicrosoft.com"],
    "Desk 3": ["Desk3@Team4DxHub.onmicrosoft.com"],
    "Desk 4": ["Desk4@Team4DxHub.onmicrosoft.com"],
    "Desk 5": ["Desk5@Team4DxHub.onmicrosoft.com"],
    "Desk 3 & 4 (Dual)": [
        "Desk3@Team4DxHub.onmicrosoft.com",
        "Desk4@Team4DxHub.onmicrosoft.com",
    ],
}


# ---------------------------------------------------------------------------
# AWS CLIENTS
# ---------------------------------------------------------------------------

AWS_REGION = os.environ.get("AWS_REGION", "us-west-2")

bedrock = boto3.client(
    "bedrock-runtime",
    region_name=AWS_REGION,
)

s3_client = boto3.client("s3")


# ---------------------------------------------------------------------------
# ENVIRONMENT VARIABLES
# ---------------------------------------------------------------------------

TENANT_ID = os.environ.get("TENANT_ID", "")
CLIENT_ID = os.environ.get("CLIENT_ID", "")
CLIENT_SECRET = os.environ.get("CLIENT_SECRET", "")
USER_EMAIL = os.environ.get("USER_EMAIL", "")

# This lets you change the model without editing the Lambda code.
BEDROCK_MODEL_ID = os.environ.get(
    "BEDROCK_MODEL_ID",
    "us.amazon.nova-pro-v1:0",
)

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"


# ---------------------------------------------------------------------------
# MICROSOFT GRAPH AUTHENTICATION
# ---------------------------------------------------------------------------

def get_graph_token():
    """Authenticate with Microsoft Entra ID and return a Graph access token."""

    missing_variables = [
        variable_name
        for variable_name, variable_value in {
            "TENANT_ID": TENANT_ID,
            "CLIENT_ID": CLIENT_ID,
            "CLIENT_SECRET": CLIENT_SECRET,
            "USER_EMAIL": USER_EMAIL,
        }.items()
        if not variable_value
    ]

    if missing_variables:
        raise RuntimeError(
            "Missing required environment variables: "
            + ", ".join(missing_variables)
        )

    token_url = (
        "https://login.microsoftonline.com/"
        f"{urllib.parse.quote(TENANT_ID, safe='')}"
        "/oauth2/v2.0/token"
    )

    token_payload = urllib.parse.urlencode(
        {
            "client_id": CLIENT_ID,
            "scope": "https://graph.microsoft.com/.default",
            "client_secret": CLIENT_SECRET,
            "grant_type": "client_credentials",
        }
    ).encode("utf-8")

    request = Request(
        token_url,
        data=token_payload,
        headers={
            "Content-Type": "application/x-www-form-urlencoded",
            "Accept": "application/json",
        },
        method="POST",
    )

    try:
        with urlopen(request, timeout=15) as response:
            response_body = response.read().decode("utf-8")
            response_data = json.loads(response_body)

    except HTTPError as error:
        error_body = error.read().decode(
            "utf-8",
            errors="replace",
        )

        raise RuntimeError(
            "Microsoft Graph authentication failed. "
            f"HTTP {error.code}: {error_body}"
        ) from error

    except (URLError, TimeoutError) as error:
        raise RuntimeError(
            f"Microsoft Graph authentication request failed: {error}"
        ) from error

    except json.JSONDecodeError as error:
        raise RuntimeError(
            "Microsoft authentication returned invalid JSON."
        ) from error

    access_token = response_data.get("access_token")

    if not access_token:
        raise RuntimeError(
            "Microsoft authentication response did not contain "
            "an access token."
        )

    return access_token


# ---------------------------------------------------------------------------
# MICROSOFT GRAPH FORWARDING
# ---------------------------------------------------------------------------

def forward_email_via_graph(
    message_id,
    to_addresses,
    token,
):
    """Forward one Outlook message using Microsoft Graph."""

    if not message_id:
        raise ValueError("A Microsoft Graph message ID is required.")

    if not to_addresses:
        raise ValueError("At least one forwarding recipient is required.")

    if not token:
        raise ValueError("A Microsoft Graph access token is required.")

    encoded_user = urllib.parse.quote(
        USER_EMAIL,
        safe="",
    )

    encoded_message_id = urllib.parse.quote(
        message_id,
        safe="",
    )

    url = (
        f"{GRAPH_BASE_URL}/users/{encoded_user}/messages/"
        f"{encoded_message_id}/forward"
    )

    recipients = [
        {
            "emailAddress": {
                "address": address,
            }
        }
        for address in to_addresses
    ]

    payload = json.dumps(
        {
            "comment": (
                "Auto-routed by Bedrock AI Accounts Payable Assistant"
            ),
            "toRecipients": recipients,
        }
    ).encode("utf-8")

    request = Request(
        url,
        data=payload,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        },
        method="POST",
    )

    try:
        with urlopen(request, timeout=15) as response:
            status_code = response.status

            if status_code not in (200, 202):
                raise RuntimeError(
                    "Unexpected Microsoft Graph forwarding response: "
                    f"HTTP {status_code}"
                )

        print(
            "Successfully forwarded Outlook message "
            f"{message_id[:20]}... to {to_addresses}"
        )

    except HTTPError as error:
        error_body = error.read().decode(
            "utf-8",
            errors="replace",
        )

        raise RuntimeError(
            "Microsoft Graph forwarding failed. "
            f"HTTP {error.code}: {error_body}"
        ) from error

    except (URLError, TimeoutError) as error:
        raise RuntimeError(
            f"Microsoft Graph forwarding request failed: {error}"
        ) from error


# ---------------------------------------------------------------------------
# BEDROCK CLASSIFICATION
# ---------------------------------------------------------------------------

def classify_email(email_text):
    """Classify an email and attach its configured desk recipients."""

    prompt = f"""
Role & Objective:
You are an automated Accounts Payable email routing assistant for Foothill
De-Anza CC. Analyze the incoming email and route it according to the rules
below.

ROUTING LOGIC — EVALUATE IN THIS EXACT ORDER

1. If the email contains "Foothill De-Anza Foundation",
   "travel reimbursement", "trip reimbursement", "trip voucher",
   or "travel voucher":
   Destination: Desk 5

2. If it contains both "lease" and "printer", or both "lease" and "copier":
   Destination: Desk 3

3. If it contains "OEI", "CVC", or an obvious variant:
   Destination: Desk 4

4. If the PO starts with "EE", "HH", or "SS":
   Extract the first character of the company or last name.
   - Special character or A-F: Desk 3
   - G-J: Desk 4
   - K-Z: Desk 1

5. If the PO starts with "AA":
   Destination: Desk 3

6. If it contains "cash receipts" or an obvious variant:
   Destination: Desk 3 & 4 (Dual)

7. If it contains "Measure G", "MG", "Measure C", "MC",
   "Utilities", "Phone Bills", "Library Bills", or "ProCard":
   Destination: Desk 2

8. If it contains "TIN matching" or an obvious variant:
   Destination: Desk 1

9. If it contains "Statement" or an obvious variant:
   Destination: Statement

10. If it contains "DPR", "Direct Pay", or "Direct Pay Request":
    - Service invoice: Desk 5
    - Otherwise, use the first character of the company or last name:
      - Special character or A-L: Desk 2
      - M-Z: Desk 4

FALLBACK

If no rule matches:
Destination: Uncategorized

OUTPUT INSTRUCTIONS

Return only valid raw JSON. Do not include Markdown, commentary, or code
fences.

Use exactly this schema:

{{
    "destination": "Desk 1" | "Desk 2" | "Desk 3" | "Desk 4" |
                   "Desk 5" | "Desk 3 & 4 (Dual)" | "Statement" |
                   "Uncategorized",
    "target_folder": "Desk 1" | "Desk 2" | "Desk 3" | "Desk 4" |
                     "Desk 5" | "Uncategorized - Statement" |
                     "Uncategorized",
    "copy_folder": "Desk 4" or null,
    "rule_matched": "Brief rule name or number"
}}

EMAIL CONTENT

{email_text}
"""

    response = bedrock.converse(
        modelId=BEDROCK_MODEL_ID,
        messages=[
            {
                "role": "user",
                "content": [
                    {
                        "text": prompt,
                    }
                ],
            }
        ],
        inferenceConfig={
            "temperature": 0.0,
            "maxTokens": 500,
        },
    )

    raw_output = (
        response["output"]["message"]["content"][0]["text"].strip()
    )

    # Remove accidental Markdown code fences.
    if raw_output.startswith("```"):
        raw_output = raw_output.removeprefix("```json")
        raw_output = raw_output.removeprefix("```")
        raw_output = raw_output.removesuffix("```")
        raw_output = raw_output.strip()

    try:
        result = json.loads(raw_output)

    except json.JSONDecodeError as error:
        raise RuntimeError(
            "Bedrock returned invalid routing JSON: "
            f"{raw_output}"
        ) from error

    destination = result.get("destination")

    allowed_destinations = {
        "Desk 1",
        "Desk 2",
        "Desk 3",
        "Desk 4",
        "Desk 5",
        "Desk 3 & 4 (Dual)",
        "Statement",
        "Uncategorized",
    }

    if destination not in allowed_destinations:
        raise RuntimeError(
            f"Bedrock returned an invalid destination: {destination!r}"
        )

    forward_to_addresses = DESK_EMAILS.get(
        destination,
        [],
    )

    # Forwarding is determined by the configured destination map,
    # not by a model-generated boolean.
    result["forward_to_addresses"] = forward_to_addresses
    result["should_forward"] = bool(forward_to_addresses)

    return result


# ---------------------------------------------------------------------------
# LAMBDA HANDLER
# ---------------------------------------------------------------------------

def lambda_handler(event, context):
    print(f"Received event: {json.dumps(event)}")

    records = event.get("Records", [])

    if not records:
        return {
            "statusCode": 400,
            "body": json.dumps(
                {
                    "error": (
                        "This Lambda expects an S3 event notification."
                    ),
                    "received_event": event,
                }
            ),
        }

    results = []

    for event_record in records:
        bucket = None
        key = None

        try:
            if event_record.get("eventSource") != "aws:s3":
                print(
                    f"Skipping non-S3 record: {event_record}"
                )
                continue

            s3_record = event_record.get("s3", {})

            bucket = (
                s3_record
                .get("bucket", {})
                .get("name")
            )

            encoded_key = (
                s3_record
                .get("object", {})
                .get("key")
            )

            if not bucket or not encoded_key:
                print(
                    f"Skipping invalid S3 record: {event_record}"
                )
                continue

            key = urllib.parse.unquote_plus(encoded_key)

            # Process metadata only. Ignore PDF attachment events.
            if not key.endswith("/metadata.json"):
                print(
                    "Skipping non-metadata object: "
                    f"s3://{bucket}/{key}"
                )

                results.append(
                    {
                        "bucket": bucket,
                        "key": key,
                        "status": "skipped_non_metadata",
                    }
                )

                continue

            obj = s3_client.get_object(
                Bucket=bucket,
                Key=key,
            )

            raw_content = obj["Body"].read()

            metadata = json.loads(
                raw_content.decode("utf-8")
            )

            # Use the complete Graph ID stored in the JSON.
            # Do not derive it from the S3 path.
            message_id = metadata.get("message_id")

            if not message_id:
                raise ValueError(
                    "metadata.json does not contain message_id: "
                    f"{key}"
                )

            email_text = json.dumps(
                {
                    "subject": metadata.get("subject", ""),
                    "body_preview": metadata.get(
                        "body_preview",
                        "",
                    ),
                    "sender": metadata.get("sender", ""),
                    "sender_name": metadata.get(
                        "sender_name",
                        "",
                    ),
                    "attachments": metadata.get(
                        "attachments",
                        [],
                    ),
                },
                ensure_ascii=False,
            )

            result = classify_email(email_text)

            print(
                "Bedrock routing result for "
                f"message_id={message_id}: "
                f"{json.dumps(result)}"
            )

            to_addresses = result.get(
                "forward_to_addresses",
                [],
            )

            if to_addresses:
                token = get_graph_token()

                forward_email_via_graph(
                    message_id=message_id,
                    to_addresses=to_addresses,
                    token=token,
                )

                routing_status = "forwarded"

            else:
                # Statement and Uncategorized currently have no forwarding
                # addresses in DESK_EMAILS.
                routing_status = "classified_without_forwarding"

                print(
                    "No forwarding recipient configured for "
                    f"destination={result.get('destination')!r}"
                )

            results.append(
                {
                    "bucket": bucket,
                    "key": key,
                    "message_id": message_id,
                    "status": routing_status,
                    "routing_result": result,
                }
            )

        except Exception as error:
            print(
                f"Error processing s3://{bucket}/{key}: "
                f"{type(error).__name__}: {error}"
            )

            results.append(
                {
                    "bucket": bucket,
                    "key": key,
                    "status": "error",
                    "error_type": type(error).__name__,
                    "error": str(error),
                }
            )

    failed_results = [
        result
        for result in results
        if result.get("status") == "error"
    ]

    return {
        "statusCode": 207 if failed_results else 200,
        "body": json.dumps(results),
    }
