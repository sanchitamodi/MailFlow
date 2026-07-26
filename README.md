# MailFlow_EmailRouter
# Accounts Payable Automated Email Ingestion & AI Routing Pipeline

An event-driven serverless system deployed on **AWS Lambda** that automates processing, AI classification, and auto-forwarding for incoming Accounts Payable (AP) emails via the **Microsoft Graph API** and **Amazon Bedrock**.

---

## Architecture Overview
                      ┌──────────────────────────┐
                      │    Microsoft Graph API   │
                      │   (Outlook Shared Inbox) │
                      └────────────┬─────────────┘
                                   │
                                   │ 1. Scheduled Poll / Trigger
                                   ▼
                 ┌──────────────────────────────────┐
                 │  Lambda 1: Ingestion & Extraction │
                 └─────────────────┬────────────────┘
                                   │
                                   │ 2. Save metadata.json & PDF attachments
                                   ▼
                      ┌──────────────────────────┐
                      │     Amazon S3 Bucket     │
                      │  s3://bucket/incoming/   │
                      └────────────┬─────────────┘
                                   │
                                   │ 3. S3 Event Trigger (metadata.json)
                                   ▼
                 ┌──────────────────────────────────┐
                 │   Lambda 2: AI Classification    │
                 │          & Forwarding            │
                 └───────┬──────────────────┬───────┘
                         │                  │
       4. Inference Call │                  │ 5. Forward Message
                         ▼                  ▼
               ┌──────────────────┐  ┌─────────────────────┐
               │  Amazon Bedrock  │  │ Microsoft Graph API │
               │  (Amazon Nova)   │  │  (Forward Endpoint) │
               └──────────────────┘  └─────────────────────┘


## Core Components

### 1. Ingestion Worker (`lambda_1_ingestion.py`)
* **Trigger:** Scheduled EventBridge Rule (Cron/Rate) or manual execution.
* **Responsibilities:**
  * Connects to Microsoft Entra ID (OAuth 2.0 `client_credentials`) to retrieve an access token.
  * Queries unread messages from the target Outlook inbox using Microsoft Graph API.
  * Extracts PDF attachments, sanitizes file paths, and uploads them to S3 under `incoming/<message_id>/attachments/`.
  * Generates an email summary payload (`metadata.json`) containing sender info, subject line, body preview, and attachment paths, saved to `incoming/<message_id>/metadata.json`.
  * Marks the email as **read** in Outlook *only after* successful S3 upload.

### 2. Router & Classifier (`lambda_2_routing.py`)
* **Trigger:** S3 Event Notification (`s3:ObjectCreated:*` matching suffix `/metadata.json`).
* **Responsibilities:**
  * Reads the `metadata.json` payload directly from S3.
  * Calls **Amazon Bedrock** (using Amazon Nova Pro by default) to classify the email based on custom AP business rules (such as matching PO prefixes, key terms like "travel reimbursement", "lease/copier", or "cash receipts").
  * Maps the classification output to designated desk email distribution addresses (e.g., Desk 1 through Desk 5, or dual routing).
  * Automatically forwards the original Outlook message to the target desk recipients via Microsoft Graph API.

## Tech Stack
  * AWS Lambda
  * Microsoft Graph API
  * AWS S3
  * Amazon EventBridge
  * Amazon Bedrock
  * IAM
  * Python
