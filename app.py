import os
from dotenv import load_dotenv
import gspread
from slack_bolt import App
from slack_bolt.adapter.socket_mode import SocketModeHandler
from google.oauth2.service_account import Credentials
import certifi
import ssl
import time
import schedule
import random
import pytz
from datetime import datetime
from google.api_core import retry
from google.api_core.exceptions import ResourceExhausted
from slack_sdk.errors import SlackApiError

ssl._create_default_https_context = ssl._create_unverified_context

load_dotenv()
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
SERVICE_ACCOUNT_FILE = "credentials.json"
SHEET_ID = os.environ.get("GOOGLE_SHEET_ID")
UTBK_COLS = {
    "STATUS_QC": "Status",
    "SOLUTION": "Solution including Concepts",
    "QC_COL": "Hasil QC",
    "REJECTION_NOTES": "Rejection Notes",
    "SOLUTION_LINK": "Solution Link",
    "STARTED_AT": "Started At",
    "APPROVED_AT": "Approved At",
    "REJECTED_AT": "Rejected At",
    "PIC": "PIC",
}

QC_CHANNEL = os.environ.get("SLACK_QC_CHANNEL")
SLACK_BOT_TOKEN = os.environ.get("SLACK_BOT_TOKEN")
SLACK_APP_TOKEN = os.environ.get("SLACK_APP_TOKEN")

app = App(token=SLACK_BOT_TOKEN)

creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
gc = gspread.authorize(creds)
sheet = gc.open_by_key(SHEET_ID).worksheet("UTBK")
headers = sheet.row_values(1)


def with_retry(func, max_retries=3):
    for attempt in range(max_retries):
        try:
            return func()
        except ResourceExhausted as e:
            if attempt == max_retries - 1:
                raise
            wait_time = (2**attempt) + random.random()
            time.sleep(wait_time)
        except Exception as e:
            raise


def convert_utc_to_jakarta(time):
    utc_time = time.replace(tzinfo=pytz.utc)
    jakarta_tz = pytz.timezone("Asia/Jakarta")
    changed_timezone = utc_time.astimezone(jakarta_tz)
    return changed_timezone.strftime("%Y-%m-%d %H:%M:%S")


header_cache = {}


def find_col_index(header_name):
    if header_name in header_cache:
        return header_cache[header_name]

    try:
        col_index = headers.index(header_name)
        header_cache[header_name] = col_index
        return col_index
    except ValueError:
        print(f"header with name {header_name} not found")
        return -1


def strip_html_tags(html):
    if not html:
        return ""
    import re

    clean = re.compile("<.*?>")
    return (
        re.sub(clean, "", html)
        .replace("&nbsp;", " ")
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&amp;", "&")
    )


def contains_image(content):
    if content is None:
        return False
    content = str(content).lower()
    return "<img" in content or "<image" in content


def check_for_new_questions():
    """Periksa spreadsheet untuk soal baru yang perlu di-QC"""
    try:
        all_values = sheet.get_all_values()

        for row_idx, row in enumerate(all_values[1:], start=2):
            qc_status = row[find_col_index(UTBK_COLS["STATUS_QC"])]
            solution = row[find_col_index(UTBK_COLS["SOLUTION"])]
            if qc_status == "Ready to QC" and solution:
                send_question_to_slack(row_idx)
                time.sleep(1)
    except Exception as e:
        print(f"Error checking for new questions: {e}")


def send_question_to_slack(row_number):
    question_id = sheet.cell(row_number, 1).value
    subject_name = sheet.cell(row_number, 3).value
    chapter_name = sheet.cell(row_number, 4).value
    topic_name = sheet.cell(row_number, 5).value
    question = sheet.cell(row_number, 8).value
    option_a = sheet.cell(row_number, 10).value
    option_b = sheet.cell(row_number, 11).value
    option_c = sheet.cell(row_number, 12).value
    option_d = sheet.cell(row_number, 13).value
    option_e = sheet.cell(row_number, 14).value
    correct_option = sheet.cell(row_number, 15).value
    solution_link = sheet.cell(row_number, 22).value
    pic = sheet.cell(row_number, 26).value

    blocks = [
        {
            "type": "header",
            "text": {
                "type": "plain_text",
                "text": f"Question #{question_id}",
                "emoji": True,
            },
        },
        {
            "type": "section",
            "fields": [
                {"type": "mrkdwn", "text": f"*Subject:*\n{subject_name}"},
                {"type": "mrkdwn", "text": f"*Chapter:*\n{chapter_name}"},
                {"type": "mrkdwn", "text": f"*Topic:*\n{topic_name}"},
            ],
        },
        {"type": "divider"},
        {
            "type": "section",
            "text": {
                "type": "mrkdwn",
                "text": "🚨 *The question failed to generate!* Please click *Lihat Soal* below for details.",
            },
        },
        {"type": "divider"},
        {
            "type": "actions",
            "elements": [
                {
                    "type": "button",
                    "text": {
                        "type": "plain_text",
                        "text": "Lihat Soal",
                        "emoji": True,
                    },
                    "url": solution_link,
                },
                {
                    "type": "button",
                    "text": {"type": "plain_text", "text": "Approve", "emoji": True},
                    "style": "primary",
                    "value": f"approve_{question_id}_{row_number}",
                    "action_id": "approve_question",
                },
                {
                    "type": "button",
                    "text": {"type": "plain_text", "text": "Reject", "emoji": True},
                    "style": "danger",
                    "value": f"reject_{question_id}_{row_number}",
                    "action_id": "reject_question",
                },
            ],
        },
    ]

    full_blocks = [
        {
            "type": "header",
            "text": {
                "type": "plain_text",
                "text": f"Question #{question_id}",
                "emoji": True,
            },
        },
        {
            "type": "section",
            "fields": [
                {"type": "mrkdwn", "text": f"*Subject:*\n{subject_name}"},
                {"type": "mrkdwn", "text": f"*Chapter:*\n{chapter_name}"},
                {"type": "mrkdwn", "text": f"*Topic:*\n{topic_name}"},
            ],
        },
        {"type": "divider"},
        {
            "type": "section",
            "text": {
                "type": "mrkdwn",
                "text": f"*Pertanyaan:*\n{strip_html_tags(question)}",
            },
        },
        {
            "type": "section",
            "text": {
                "type": "mrkdwn",
                "text": f"*Pilihan Jawaban:*\n"
                + f"A: {strip_html_tags(option_a)}\n"
                + f"B: {strip_html_tags(option_b)}\n"
                + f"C: {strip_html_tags(option_c)}\n"
                + f"D: {strip_html_tags(option_d)}\n"
                + f"E: {strip_html_tags(option_e)}\n\n"
                + f"*Jawaban Benar:* {correct_option}",
            },
        },
        {"type": "divider"},
        {
            "type": "actions",
            "elements": [
                {
                    "type": "button",
                    "text": {
                        "type": "plain_text",
                        "text": "Lihat Solusi",
                        "emoji": True,
                    },
                    "url": solution_link,
                },
                {
                    "type": "button",
                    "text": {"type": "plain_text", "text": "Approve", "emoji": True},
                    "style": "primary",
                    "value": f"approve_{question_id}_{row_number}",
                    "action_id": "approve_question",
                },
                {
                    "type": "button",
                    "text": {"type": "plain_text", "text": "Reject", "emoji": True},
                    "style": "danger",
                    "value": f"reject_{question_id}_{row_number}",
                    "action_id": "reject_question",
                },
            ],
        },
    ]
    options = [option_a, option_b, option_c, option_d, option_e]
    if (
        contains_image(question)
        or any(contains_image(opt) for opt in options)
        or len(question) > 2900
    ):
        final_blocks = blocks
    else:
        final_blocks = full_blocks

    try:
        result = app.client.chat_postMessage(
            channel=pic,
            text=f"Question #{question_id}",
            blocks=final_blocks,
        )
        sheet.update_cell(
            row_number, find_col_index(UTBK_COLS["STATUS_QC"]) + 1, "Assigned"
        )
        sheet.update_cell(
            row_number,
            find_col_index(UTBK_COLS["STARTED_AT"]) + 1,
            convert_utc_to_jakarta(datetime.utcnow()),
        )
        print(f"Sent question #{question_id} (row {row_number}) for QC")
        return True
    except SlackApiError as e:
        print(f"Error sending question to Slack: {e.response['error']}")
        return False
    except Exception as e:
        print(f"General error sending question to Slack: {e}")
        return False


@app.action("approve_question")
def handle_approve(ack, body, client):
    ack()
    try:
        value = body["actions"][0]["value"]
        _, question_id, row_number = value.split("_")
        row_number = int(row_number)

        sheet.update_cell(
            row_number, find_col_index(UTBK_COLS["QC_COL"]) + 1, "Approved"
        )
        sheet.update_cell(
            row_number, find_col_index(UTBK_COLS["STATUS_QC"]) + 1, "Checked"
        )
        sheet.update_cell(
            row_number,
            find_col_index(UTBK_COLS["APPROVED_AT"]) + 1,
            convert_utc_to_jakarta(datetime.utcnow()),
        )

        original_message = body["message"]
        blocks = original_message["blocks"]

        if len(blocks) > 0 and blocks[-1]["type"] == "actions":
            blocks[-1] = {
                "type": "section",
                "text": {
                    "type": "mrkdwn",
                    "text": f"✅ *Approved* oleh <@{body['user']['id']}> pada {convert_utc_to_jakarta(datetime.utcnow())}",
                },
            }

        try:
            client.chat_update(
                channel=body["channel"]["id"], ts=original_message["ts"], blocks=blocks
            )
        except SlackApiError as e:
            print(f"Error updating message: {e.response['error']}")
            client.chat_postMessage(
                channel=body["channel"]["id"],
                thread_ts=original_message["ts"],
                text=f"✅ Question #{question_id} telah diapprove oleh <@{body['user']['id']}>!",
            )
    except Exception as e:
        print(f"Error in approval handler: {e}")
        try:
            client.chat_postMessage(
                channel=body["channel"]["id"],
                text=f"✅ Approval berhasil dicatat, tapi terjadi error saat memperbarui pesan: {str(e)}",
            )
        except:
            pass


@app.action("reject_question")
def handle_reject(ack, body, client):
    ack()

    try:
        value = body["actions"][0]["value"]
        client.views_open(
            trigger_id=body["trigger_id"],
            view={
                "type": "modal",
                "callback_id": "reject_modal",
                "title": {"type": "plain_text", "text": "Alasan Reject"},
                "submit": {"type": "plain_text", "text": "Submit"},
                "blocks": [
                    {
                        "type": "input",
                        "block_id": "reject_reason",
                        "element": {
                            "type": "plain_text_input",
                            "action_id": "reason",
                            "multiline": True,
                        },
                        "label": {"type": "plain_text", "text": "Alasan Reject"},
                    }
                ],
                "private_metadata": value
                + "_"
                + body["channel"]["id"]
                + "_"
                + body["message"]["ts"]
                + "_",
            },
        )
    except Exception as e:
        print(f"Error opening modal: {e}")
        try:
            client.chat_postMessage(
                channel=body["channel"]["id"],
                text=f"❌ Error saat membuka form reject: {str(e)}",
            )
        except:
            pass


@app.view("reject_modal")
def handle_rejection_submission(ack, body, client, view):
    ack()

    try:
        reason = view["state"]["values"]["reject_reason"]["reason"]["value"]
        private_metadata = view["private_metadata"]
        metadata_parts = private_metadata.split("_")

        if len(metadata_parts) >= 5:
            _, question_id, row_number, channel_id, message_ts = metadata_parts[:5]
            row_number = int(row_number)

            sheet.update_cell(
                row_number, find_col_index(UTBK_COLS["QC_COL"]) + 1, "Rejected"
            )
            sheet.update_cell(
                row_number,
                find_col_index(UTBK_COLS["STATUS_QC"]) + 1,
                "Question Returned",
            )
            sheet.update_cell(
                row_number, find_col_index(UTBK_COLS["REJECTION_NOTES"]) + 1, reason
            )
            sheet.update_cell(
                row_number,
                find_col_index(UTBK_COLS["REJECTED_AT"]) + 1,
                convert_utc_to_jakarta(datetime.utcnow()),
            )

            try:
                if message_ts:
                    blocks = [
                        {
                            "type": "header",
                            "text": {
                                "type": "plain_text",
                                "text": f"Question #{question_id}",
                                "emoji": True,
                            },
                        },
                        {"type": "divider"},
                        {
                            "type": "section",
                            "text": {
                                "type": "mrkdwn",
                                "text": f"❌ *Rejected* oleh <@{body['user']['id']}> pada {convert_utc_to_jakarta(datetime.utcnow())}\n*Alasan:* {reason}",
                            },
                        },
                    ]

                    client.chat_update(channel=channel_id, ts=message_ts, blocks=blocks)
                else:
                    client.chat_postMessage(
                        channel=channel_id,
                        thread_ts=message_ts,
                        text=f"❌ Question #{question_id} telah direject oleh <@{body['user']['id']}>.\n*Alasan:* {reason}",
                    )
            except SlackApiError as e:
                print(f"Error updating message: {e.response['error']}")
                client.chat_postMessage(
                    channel=channel_id,
                    text=f"❌ Question #{question_id} telah direject oleh <@{body['user']['id']}>.\n*Alasan:* {reason}",
                )
        else:
            # Format metadata tidak sesuai, gunakan cara lama
            _, question_id, row_number = metadata_parts[:3]
            row_number = int(row_number)

            # Update spreadsheet

            sheet.update_cell(
                row_number,
                find_col_index(UTBK_COLS["QC_COL"]) + 1,
                f"Rejected: {reason}",
            )

            # Kirim notifikasi
            client.chat_postMessage(
                channel=body["user"]["id"],  # DM ke user yang reject
                text=f"❌ Question #{question_id} telah direject dengan alasan: {reason}",
            )
    except Exception as e:
        print(f"Error handling rejection: {e}")
        # Notify user of error
        try:
            client.chat_postMessage(
                channel=body["user"]["id"],
                text=f"❌ Terjadi error saat memproses rejection: {str(e)}",
            )
        except:
            pass


def run_scheduled_check():
    print(f"Checking for new questions at {time.strftime('%H:%M:%S')}")
    check_for_new_questions()


if __name__ == "__main__":
    ssl._create_default_https_context = ssl._create_unverified_context
    os.environ["SSL_CERT_FILE"] = certifi.where()

    print(f"Bot token defined: {'Yes' if SLACK_BOT_TOKEN else 'No'}")
    print(f"App token defined: {'Yes' if SLACK_APP_TOKEN else 'No'}")
    print(f"QC channel defined: {'Yes' if QC_CHANNEL else 'No'}")

    schedule.every(2).minutes.do(run_scheduled_check)

    print("Starting initial check for new questions...")
    try:
        with_retry(lambda: check_for_new_questions())
    except Exception as e:
        print(f"Initial check failed: {e}")

    handler = SocketModeHandler(app, SLACK_APP_TOKEN)

    from threading import Thread

    def run_scheduler():
        while True:
            try:
                schedule.run_pending()
                time.sleep(1)
            except Exception as e:
                print(f"Scheduler error: {e}")
                time.sleep(10)

    scheduler_thread = Thread(target=run_scheduler)
    scheduler_thread.daemon = True
    scheduler_thread.start()

    print("App is running! Press Ctrl+C to exit.")
    handler.start()
