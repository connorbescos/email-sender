import datetime
import logging
import os
import requests
from azure.functions import FuncExtensionException, TimerRequest
from jinja2 import Environment, FileSystemLoader, select_autoescape

# Azure Function Entry Point
def main(mytimer: TimerRequest) -> None:
    utc_timestamp = datetime.datetime.utcnow().replace(tzinfo=datetime.timezone.utc).isoformat()
    if mytimer.past_due:
        logging.info('The timer is past due!')

    logging.info('Python timer trigger function ran at %s', utc_timestamp)

    send_status = send_success_findmysurvey_email("123456", "http://example.com/survey/123456", "user@example.com")
    logging.info(f"Email send status: {send_status}")


def authenticate_ms_graph(client_id, client_secret, tenant_id, scope="https://graph.microsoft.com/.default") -> str:
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    data = {
        "client_id": client_id,
        "scope": scope,
        "client_secret": client_secret,
        "grant_type": "client_credentials",
    }
    response = requests.post(url, headers=headers, data=data)
    if response.status_code != 200 or "access_token" not in response.json():
        raise FuncExtensionException("Failed to authenticate with MS Graph. Emails will not be sent.")
    return response.json()["access_token"]


def generate_success_email_body(survey_id: str, survey_link: str) -> str:
    template_dir = os.path.abspath("./templates")  # Adjust path as necessary
    env = Environment(loader=FileSystemLoader(template_dir), autoescape=select_autoescape(["html"]))
    template = env.get_template("lookup_success_email.html")
    email_data = {"survey_id": survey_id, "survey_link": survey_link}
    return template.render(email_data)

def send_email(access_token, subject, sender_email, recipient_email, content):
    url = f"https://graph.microsoft.com/v1.0/users/{sender_email}/sendMail"
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
    data = {
        "message": {
            "subject": subject,
            "body": {"contentType": "HTML", "content": content},
            "toRecipients": [{"emailAddress": {"address": recipient_email}}],
        },
        "saveToSentItems": "true",
    }
    response = requests.post(url, headers=headers, json=data)
    response.raise_for_status() 
    return response.status_code

def send_success_findmysurvey_email(survey_id, survey_link, contact_email):
    client_ID = "<your_client_id>"
    client_secret = "<your_client_secret>"
    tenant_ID = "<your_tenant_id>"
    sender_email = "<your_sender_email>"
    token = authenticate_ms_graph(client_ID, client_secret, tenant_ID)
    email_body = generate_success_email_body(survey_id, survey_link)
    send_status = send_email(token, "Tenant Satisfaction Survey Request", sender_email, contact_email, email_body)
    return send_status
