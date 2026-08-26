import os
import json
import openpyxl
from dotenv import load_dotenv
from slack_bolt import App
from slack_bolt.adapter.socket_mode import SocketModeHandler
from slack_sdk.errors import SlackApiError
from io import BytesIO

# Importing database functions

import database

# Loading environment variables from .env
load_dotenv()

# Initialsing Slack with the Bot Token and Signing Secret

app = App(
    token=os.environ.get("SLACK_bot_token"),
    signing_secret=os.environ.get("SLACK_SIGNING_SECRET"),
    # Bolt runs the handlers on worker threads. This executor carries the
    # "which Slack delivery is this" note set by the middleware below onto
    # the worker thread, so a redelivered click is still recognised as the
    # same action there. See listener_executor() in database.py.
    listener_executor=database.listener_executor()
)

def packing_modal_view(metadata, summary_text):
    """
    The Packing form.

    Two paths reach it - the border phase finishing, and a job that turns out
    to have no border at all - and they differ only in the few lines of summary
    above the button. Defined here once so both say the same thing, and so a
    change to the form cannot land on one path and miss the other.
    """
    return {
        "type": "modal",
        "callback_id": "trk_packing_modal",
        "title": {"type": "plain_text", "text": "Packing (Phase 3)"},
        "submit": {"type": "plain_text", "text": "Start Packing Phase"},
        "close": {"type": "plain_text", "text": "Cancel"},
        "private_metadata": metadata,
        "blocks": [
            {
                "type": "section",
                "text": {"type": "mrkdwn", "text": summary_text}
            }
        ]
    }


def border_time_display(task):
    """
    What to show wherever a border time would go.

    A job with no border has no border time. Printing "0 h 0m 0s" would read as
    a border that was worked and happened to take no time, which is a different
    thing and the wrong thing. Every border time on a card, in the summary and
    in the export goes through here so they all say it the same way.
    """
    if task.get("border_skipped"):
        return "No Border"
    return database.format_elapsed(task["border_elapsed"] or 0)


@app.middleware
def track_slack_delivery(body, next):
    # Note which Slack delivery is being handled, so a redelivery of the same
    # click is recognised as the same action rather than applied twice. Scoped
    # to this handler and cleared afterwards, so nothing carries over into the
    # next request. Changes no behaviour a maker can see.
    with database.slack_request(body):
        return next()

# /hello command test in slack

@app.command ("/hello")
def hello_command(ack, body, say):
    
    ack()
    user_id = body["user_id"]
    say(f"Hi there, <@{user_id}>! I'm ready to track your projects.")

# /track command to start the project tracking process. Loading the form.    

@app.command("/track")
def track_command(ack, body, client):
    # Acknowledging the command
    ack()
    user_id = body ["user_id"]
    
    # Subcommand: /track export
    subcommand = body.get ("text", "").strip().lower()
    if subcommand == "export":
        handle_export(body,client)
        return
    
    
    # Checking if the task is already running
    active_task = database.get_active_task(user_id)
    if active_task:
        client.chat_postEphemeral(
            channel=body["channel_id"],
            user=user_id,
            text=f"You have already have an active task: * {active_task['task_description']}*. Please complete or stop it before starting a new one."
            
        )
        return
    # opening the Modal - When you click on start this is the dictionaries and lists which creates the look of the form
    client.views_open(
        trigger_id=body["trigger_id"],
        view={
            "type": "modal",
            "callback_id": "trk_track_step_1",  # ID used to catch the submission
            "title": {"type": "plain_text", "text": "Field Sheeting"},
            "private_metadata":body["channel_id"],
            "submit": {"type": "plain_text", "text": "Next"},
            "close": {"type": "plain_text", "text": "Cancel"},
            "blocks": [
                {
                    "type": "input",
                    "block_id": "customer_block",
                    "label": {"type": "plain_text", "text": "Customer Name"},
                    "element": {
                        "type": "plain_text_input",
                        "action_id": "customer_name"
                    }
                },
                {
                    "type": "input",
                    "block_id": "invoice_block",
                    "label": {"type": "plain_text", "text": "Invoice Number"},
                    "element": {
                        "type": "plain_text_input",
                        "action_id": "invoice_num"
                    }
                },
                {
                    "type": "input",
                    "block_id": "task_block",
                    "label": {"type": "plain_text", "text": "Task Description"},
                    "element": {
                        "type": "plain_text_input",
                        "multiline": True,
                        "action_id": "task_desc"
                    }
                },
                {
                    "type": "input",
                    "block_id": "date_block",
                    "optional": True,
                    "label": {"type": "plain_text", "text": "Due Date (DD/MM/YY)"},
                    "element": {
                        "type":"plain_text_input",
                        "action_id": "due_date"
                    }
                },
                {
                    "type": "input",
                    "block_id": "is_na_block",
                    "optional": True,
                    "label": {"type": "plain_text", "text": "No Set Date?"},
                    "element": {
                        "type": "checkboxes",
                        "action_id": "is_na",
                        "options": [
                            {
                                "text": {"type": "plain_text", "text": "N/A"},
                                "value": "is_na"
                            }
                        ]
                    }
                }
            ]
        }
    )

def resolve_existing_dm(client, user_id):
    # Find the bot's existing DM conversation with a user, paging through the
    # full list rather than trusting the first page.
    #
    # conversations_open would mint the conversation, but it needs the im:write
    # scope, which the LMSA Slack app does not hold. Listing the bot's own IM
    # conversations needs only im:read. Returns None when no DM exists yet --
    # files_upload_v2 rejects a raw user id, so the caller must handle that
    # rather than passing user_id through.
    cursor = None
    while True:
        resp = client.users_conversations(types="im", limit=200, cursor=cursor)
        for conversation in resp["channels"]:
            if conversation.get("user") == user_id:
                return conversation["id"]
        cursor = (resp.get("response_metadata") or {}).get("next_cursor") or None
        if not cursor:
            return None

def handle_export(body, client):
    user_id = body["user_id"]
    channel_id = body ["channel_id"]
    
    tasks = database.get_completed_tasks()
    
    user_name_cache = {}
    
    def get_user_name(slack_user_id):
        if slack_user_id in user_name_cache:
            return user_name_cache[slack_user_id]

        try:
            response = client.users_info(user=slack_user_id)
            profile = response.get("user", {}).get("profile", {})
            name = profile.get("display_name") or profile.get("real_name") or slack_user_id
        except SlackApiError:
            name = slack_user_id

        user_name_cache[slack_user_id] = name
        return name
    
    if not tasks:
        client.chat_postEphemeral(
            channel = channel_id,
            user = user_id,
            text = "No completed jobs found to export yet."
        )
        return
    
    # Building the Excel Spreadsheet
    
    wb = openpyxl.Workbook()
    ws = wb.active
    
    if ws is None:
        ws = wb.create_sheet(title="Completed Jobs")
    else:
        ws.title = "Completed Jobs"
    
    # Headers
    
    headers = [
        "Task ID",
        "Customer",
        "Invoice Number",
        "Task Description",
        "Due Date",
        "Field Design",
        "Field Difficulty",
        "Field Jig Size(s)",
        "Field Sheeting Time",
        "Border Design",
        "Border Difficulty",
        "Border Jig Size(s)",
        "Border Sheeting Time",
        "Packing Time",
        "Total Time",
        "General Notes",
        "Issues Encountered",
        "Completed By",
        "Date Created",
    ]
    ws.append(headers)
    
    # Styling the headers
    
    from openpyxl.styles import Font, PatternFill, Alignment
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor ="2C3E50")
    
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
        
    # Data Rows
    for task in tasks:
        field_elapsed = task["field_elapsed"] or 0
        border_elapsed = task["border_elapsed"] or 0
        packing_elapsed = task["packing_elapsed"] or 0
        total_elapsed = field_elapsed + border_elapsed + packing_elapsed

        ws.append([
            f"T-{task['task_id']}",
            task["customer_name"],
            task["invoice_number"],
            task["task_description"],
            task["due_date"] or "N/A",
            task["field_design"] or "-",
            task["difficulty"],
            # Several jigs show as one readable cell, e.g. "49.6 / 50"
            task["field_jigs"] or "-",
            database.format_elapsed(field_elapsed),
            "No Border" if task.get("border_skipped") else (task["border_design"] or "-"),
            "-" if task.get("border_skipped") else (task["border_difficulty"] or "-"),
            "-" if task.get("border_skipped") else (task["border_jigs"] or "-"),
            border_time_display(task),
            database.format_elapsed(packing_elapsed),
            database.format_elapsed(total_elapsed),
            task["general_notes"] or "None",
            task["issues_encountered"] or "None",
            get_user_name(task["user_id"]),
            task["created_at"],
        ])
        
    #Auto-sizing columns, so the text can fit
    for column_cells in ws.columns:
        max_length = max(
            len(str(cell.value)) if cell.value else 0
            for cell in column_cells
        )
        ws.column_dimensions[column_cells[0].column_letter].width = min(max_length + 4,50)
        
    # Writing to in-memory bytes
    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    
    dm_channel_id = resolve_existing_dm(client, user_id)
    if not dm_channel_id:
        client.chat_postEphemeral(
            channel=channel_id,
            user=user_id,
            text=("Unable to send the export because you do not have a direct message "
                  "conversation with the bot yet. Start a job with `/track` first, then "
                  "run the export again.")
        )
        return

    #Upload Files to user's DM
    try:
        client.files_upload_v2(
            channel = dm_channel_id,
            file=buffer.getvalue(),
            filename = "trackbot_jobs_export.xlsx",
            title = f"Trackbot Export - {len(tasks)} Completed Job(s)"
        )
    except SlackApiError as err:
        error_code = err.response.get("error") if err.response is not None else "unknown_error"
        if error_code == "missing_scope":
            client.chat_postEphemeral(
                channel=channel_id,
                user=user_id,
                text=("Unable to upload export because the Slack app is missing the required "
                      "`files:write` scope. Please update the app scopes and reinstall the app.")
            )
            return
        raise
        
        #Confirmation of in the channel of the command working
        
        client.chat_postEphemeral(
            channel=channel_id,
            user=user_id,
            text=f"Export ready! Check your DMs - {len(tasks)} completed job(s) exported to Excel."
        )
        
#Step 1 Submission - This stage collects data and pushes to the Step 2 Modal
@app.view("trk_track_step_1")
def handle_step_1(ack,body,client,):
    vals = body["view"]["state"]["values"]
    channel_id = body["view"]["private_metadata"]
    
    # Extract the user and channel
    
    # Pulling the values out of the submitted form
    customer_name = vals["customer_block"]["customer_name"]["value"]
    invoice_number = vals["invoice_block"]["invoice_num"]["value"]
    task_description = vals["task_block"]["task_desc"]["value"]
    due_date = vals["date_block"]["due_date"]["value"] or "N/A"
    is_na_options = vals["is_na_block"]["is_na"].get("selected_options", [])
    is_na = 1 if is_na_options else 0
    
    # If N/A is ticked, override the due date
    if is_na:
        due_date = "N/A"
    
    #Bundling Step 1 data to pass forward
    step1_data = {
        "channel_id": channel_id,
        "customer_name": customer_name,
        "invoice_number": invoice_number,
        "task_description": task_description,
        "due_date": due_date,
        "is_na": is_na
    }

    # Pushing the 2nd Step of the Modal
    
    ack(response_action="push", view={
        "type": "modal",
        "callback_id": "trk_track_step_2",
        "title": {"type": "plain_text", "text": "Field Design 2/2"},
        "submit": {"type": "plain_text", "text": "Create Task"},
        "close": {"type": "plain_text", "text": "Back"},
        "private_metadata": json.dumps(step1_data),
        "blocks": [
            {"type": "input", "block_id": "design", "label":
                {"type": "plain_text", "text": "Field Design Name"},
                "element":{"type": "plain_text_input", "action_id": "val"}
                },
            {"type": "input", "block_id": "diff", "label":
                {"type": "plain_text", "text": "Sheeting Difficulty"},
                "element":{"type": "plain_text_input", "action_id": "difficulty", "max_length": 2}
                },
            # Usually a millimetre size like 49.6, but not always a clean
            # number - "49.4/49.8" and "template" are real entries too, so the
            # box takes whatever the maker types. Optional: not every job has
            # the jig to hand when it is logged.
            {"type": "input", "block_id": "jig_block", "optional": True, "label":
                {"type": "plain_text", "text": "Jig Size (mm)"},
                "element":{"type": "plain_text_input", "action_id": "jig_size",
                    "placeholder": {"type": "plain_text", "text": "e.g. 49.6 or template"}}
                }
        ]
    }
    )

@app.view("trk_track_step_2")
def handle_step_2(ack, body, client):
    ack(response_action = "clear")
    user_id = body["user"]["id"]
    vals = body["view"]["state"]["values"]
    
    # Fetching the Step 1 data from the metadata
    prev_data = json.loads(body["view"]["private_metadata"])
    team_channel_id = prev_data["channel_id"]
    
    # Collecting Step 2 values
    design = vals["design"]["val"]["value"]
    difficulty = vals["diff"]["difficulty"]["value"]
    jig_size = (vals["jig_block"]["jig_size"]["value"] or "").strip()

    # Saving the task to the database
    task_id = database.create_task(
        user_id=user_id,
        channel_id=team_channel_id,
        customer_name=prev_data["customer_name"],
        invoice_number=prev_data["invoice_number"],
        task_description=prev_data["task_description"],
        due_date=prev_data["due_date"],
        is_na=prev_data["is_na"],
        design=design,
        difficulty=difficulty,
        jig_size=jig_size
    )
    
    # Displaying due date on card
    due_display = "N/A" if prev_data["is_na"] else prev_data["due_date"]

    # The jig line only appears once there is a jig to show
    jig_line = f"*Jig Size:*\n{jig_size}\n" if jig_size else ""

    # chat_postMessage accepts a user id and resolves the DM itself, returning
    # the real D... conversation id in result["channel"]. conversations_open
    # would need the im:write scope, which the LMSA Slack app does not hold.
    #Posting the task card to channel
    result = client.chat_postMessage(
        channel=user_id,
        text=f"New Task -{task_id} created.",
        blocks=[
            {
                "type": "section",
                "text": {
                    "type": "mrkdwn",
                    "text": (
                        f"*New Task Created - Phase 1/4: Field Sheeting*\n"
                        f"*ID:* T-{task_id}\n"
                        f"*Customer:*\n{prev_data['customer_name']}\n"
                        f"*Invoice:*\n{prev_data['invoice_number']}\n"
                        f"*Task:*\n{prev_data['task_description']}\n"
                        f"*Field Design:*\n{design}\n"
                        f"*Difficulty:*\n{difficulty}\n"
                        f"{jig_line}"
                        f"*Due:*\n{due_display}\n"
                        f"*Status:* Created"
                    )
                }
            },
            {
                "type": "actions",
                "block_id":f"task_actions_{task_id}",
                "elements": [
                    {
                        "type":"button",
                        "text": {"type": "plain_text", "text": "Start"},
                        "style": "primary",
                        "action_id": "trk_start_task",
                        "value": str(task_id)
                    },
                    {
                        "type": "button",
                        "text": {"type": "plain_text", "text": "Edit"},
                        "action_id": "trk_edit_task",
                        "value": str(task_id) 
                    },
                    {
                        "type": "button",
                        "text": {"type": "plain_text", "text": "Delete"},
                        "style": "danger",
                        "action_id": "trk_delete_task",
                        "value":str(task_id),
                        "confirm": {
                            "title": {"type": "plain_text", "text": "Delete Task?"},
                            "text": {"type": "plain_text", "text": "Are you sure you want to delete this task? This cannot be undone."},
                            "confirm": {"type": "plain_text", "text": "Yes, Delete"},
                            "deny": {"type": "plain_text", "text": "Cancel"},
                            "style": "danger"
                        }
                    }
                ]
            }
        ]
    )
    
    # Saving the timestamp
    database.update_message_ts(task_id, result["channel"], result["ts"])
    
    client.chat_postEphemeral(
        channel = team_channel_id,
        user=user_id,
        text=f" Task -T {task_id} created! Check your DMs with the bot to start tracking the job."
    )

@app.action("trk_start_task")
def handle_start(ack, body, client):
    ack()
    task_id = int(body["actions"][0]["value"])
    user_id = body["user"]["id"]
    task = database.get_task(task_id)
    channel_id = body["container"]["channel_id"]
    
    #If there is no task found
    if task is None:
        client.chat_postEphemeral(
            channel=channel_id, 
            user=user_id,
            text="Task not found. It may have been deleted."
        )
        return

    # Block if task belongs to someone else
    if task["user_id"] != user_id:
        client.chat_postEphemeral(
            channel=channel_id,
            user=user_id,
            text="You can only control your own tasks."
        )
        return

    # Block if already running
    if task["status"] == "in_progress":
        client.chat_postEphemeral(
            channel=channel_id,
            user=user_id,
            text="This task is already running!"
        )
        return

    database.start_task(task_id)
    phase = task["current_phase"]

    # Jig lines only appear once a jig has been recorded, so old jobs with no
    # jig look exactly as they always did
    field_jig_line = f"*Jig Size:* {task['field_jigs']}\n" if task["field_jigs"] else ""
    border_jig_line = f"*Jig Size:* {task['border_jigs']}\n" if task["border_jigs"] else ""

    if phase == "field_sheeting":
        card_text = (
            f"*Phase 1/4: Field Sheeting - In Progress*\n"
            f"*ID: T-{task_id}\n"
            f"*Customer:* {task['customer_name']}\n"
            f"*Invoice:* {task['invoice_number']}\n"
            f"*Task:* {task['task_description']}"
            f"*Field Design:* {task['field_design']}\n"
            f"*Difficulty:*{task['difficulty']}\n"
            f"{field_jig_line}"
            f"*Due:* {task['due_date']}\n"
            f"*Created by:* <@{task['user_id']}>\n"
            f"*Status:* In Progress"
        )
    elif phase == "border_sheeting":
        field_time = database.format_elapsed(task["field_elapsed"])
        card_text = (
            f"*Phase 2/4: Border Sheeting — In Progress*\n"
            f"*ID:* T-{task_id}\n"
            f"*Customer:* {task['customer_name']}\n"
            f"*Invoice:* {task['invoice_number']}\n"
            f"*Task:* {task['task_description']}\n"
            f"*Border Design:* {task['border_design']}\n"
            f"*Border Difficulty:* {task['border_difficulty']}\n"
            f"{border_jig_line}"
            f"*Created by:* <@{task['user_id']}>\n"
            f"*Field Sheeting Time:* {field_time}\n"
            f"*Status:* In Progress"
        )
        
    elif phase == "packing":
        field_time = database.format_elapsed(task["field_elapsed"])
        border_time = border_time_display(task)
        # Packing is not a jig phase itself, so its lines say which phase
        # each jig belongs to
        named_field_jig_line = f"*Field Jig Size:* {task['field_jigs']}\n" if task["field_jigs"] else ""
        named_border_jig_line = f"*Border Jig Size:* {task['border_jigs']}\n" if task["border_jigs"] else ""
        card_text = (
            f"*Phase 3/4: Packing - In Progress*\n"
            f"*ID:* T-{task_id}\n"
            f"*Customer:* {task['customer_name']}\n"
            f"*Invoice:* {task['invoice_number']}\n"
            f"*Task:* {task['task_description']}\n"
            f"{named_field_jig_line}"
            f"{named_border_jig_line}"
            f"*Created by:* <@{task['user_id']}>\n"
            f"*Field Sheeting Time:* {field_time}\n"
            f"*Border Sheeting Time:* {border_time}\n"
            f"*Status:* In Progress"
        )
    
    else:
        card_text = f"*Task T-{task_id} - In Progress*\n*Status:* In Progress"
        
    buttons = [
        {
            "type": "button",
            "text": {"type": "plain_text", "text": "Stop"},
            "style": "danger",
            "action_id": "trk_stop_task",
            "value": str(task_id)
        },
        {
            "type": "button",
            "text": {"type": "plain_text", "text": "Complete Phase"},
            "action_id": "trk_complete_task",
            "value": str(task_id)
        }
    ]

    # Packing can cut in on sheeting work - a maker may need to stop and pack
    # for a while, then come back. One press swaps the timers over; nothing
    # starts a timer on its own. Offered until packing has been finished for
    # good.
    if phase in ("field_sheeting", "border_sheeting") and not task.get("packing_finished"):
        buttons.append({
            "type": "button",
            "text": {"type": "plain_text", "text": "Start Packing"},
            "action_id": "trk_start_packing",
            "value": str(task_id)
        })

    # Every working card offers Add Jig - from Border onwards the modal asks
    # which phase used it, so a late Field jig has a supported route in
    if phase in ("field_sheeting", "border_sheeting", "packing"):
        buttons.append({
            "type": "button",
            "text": {"type": "plain_text", "text": "Add Jig"},
            "action_id": "trk_add_jig",
            "value": str(task_id)
        })

    client.chat_update(
        channel=channel_id,
        ts=task["message_ts"],
        text=f"Task T-{task_id} is now in progress.",
        blocks=[
            {
                "type": "section",
                "text": {"type": "mrkdwn", "text": card_text}
            },
            {
                "type": "actions",
                "block_id": f"task_actions_{task_id}",
                "elements": buttons
            }
        ]
    )

#Stop Task Button
@app.action("trk_stop_task")
def handle_stop(ack, body, client):
    ack()
    task_id = int(body["actions"][0]["value"])
    user_id = body["user"]["id"]
    task = database.get_task(task_id)
    channel_id = body["container"]["channel_id"]

    if task is None:
        client.chat_postEphemeral(channel=channel_id,user=user_id, text="Task not found. It may have been deleted")
        return
    
    if task["user_id"] != user_id:
        client.chat_postEphemeral(
            channel=channel_id,
            user=user_id,
            text="You can only control your own tasks."
        )
        return

    database.stop_task(task_id)
    updated_task = database.get_task(task_id)
    phase = updated_task["current_phase"]

    # Jig lines only appear once a jig has been recorded
    field_jig_line = f"*Jig Size:* {updated_task['field_jigs']}\n" if updated_task["field_jigs"] else ""
    border_jig_line = f"*Jig Size:* {updated_task['border_jigs']}\n" if updated_task["border_jigs"] else ""

    #This pause card is baased on the current phase

    # Packing worked as an interruption shows up here too: once the job holds
    # any packing time, the paused sheeting card says so, so the maker can see
    # both halves of their day on one card.
    packing_line = (
        f"*Packing Time So Far:* {database.format_elapsed(updated_task['packing_elapsed'])}\n"
        if updated_task["packing_elapsed"] else ""
    )

    if phase == "field_sheeting":
        elapsed = database.format_elapsed(updated_task["field_elapsed"])
        card_text = (
            f"*Phase 1/4: Field Sheeting — Paused*\n"
            f"*ID:* T-{task_id}\n"
            f"*Customer:* {task['customer_name']}\n"
            f"*Invoice:* {task['invoice_number']}\n"
            f"*Task:* {task['task_description']}\n"
            f"*Field Design:* {task['field_design']}\n"
            f"*Difficulty:* {task['difficulty']}\n"
            f"{field_jig_line}"
            f"*Due:* {task['due_date']}\n"
            f"*Created by:* <@{task['user_id']}>\n"
            f"*Status:* Paused\n"
            f"{packing_line}"
            f"*Field Time So Far:* {elapsed}"
        )

    elif phase == "border_sheeting":
        elapsed = database.format_elapsed(updated_task["border_elapsed"])
        card_text = (
            f"*Phase 2/4: Border Sheeting — Paused*\n"
            f"*ID:* T-{task_id}\n"
            f"*Customer:* {task['customer_name']}\n"
            f"*Invoice:* {task['invoice_number']}\n"
            f"*Task:* {task['task_description']}\n"
            f"*Border Design:* {task['border_design']}\n"
            f"*Border Difficulty:* {task['border_difficulty']}\n"
            f"{border_jig_line}"
            f"*Created by:* <@{task['user_id']}>\n"
            f"*Status:* Paused\n"
            f"{packing_line}"
            f"*Border Time So Far:* {elapsed}"
        )
    
    else:
        elapsed = database.format_elapsed(updated_task["packing_elapsed"])
        # Packing is not a jig phase itself, so its lines say which phase
        # each jig belongs to
        named_field_jig_line = f"*Field Jig Size:* {updated_task['field_jigs']}\n" if updated_task["field_jigs"] else ""
        named_border_jig_line = f"*Border Jig Size:* {updated_task['border_jigs']}\n" if updated_task["border_jigs"] else ""
        card_text = (
            f"*Phase 3/4: Packing — Paused*\n"
            f"*ID:* T-{task_id}\n"
            f"*Customer:* {task['customer_name']}\n"
            f"*Invoice:* {task['invoice_number']}\n"
            f"*Task:* {task['task_description']}\n"
            f"{named_field_jig_line}"
            f"{named_border_jig_line}"
            f"*Created by:* <@{task['user_id']}>\n"
            f"*Status:* Paused\n"
            f"*Packing Time So Far:* {elapsed}"
        )

    buttons = [
        {
            "type": "button",
            "text": {"type": "plain_text", "text": "Resume"},
            "style": "primary",
            "action_id": "trk_start_task",
            "value": str(task_id)
        },
        {
            "type":"button",
            "text":{"type": "plain_text", "text": "Edit"},
            "action_id": "trk_edit_task",
            "value": str(task_id)
        },
        {
            "type": "button",
            "text": {"type": "plain_text", "text": "Complete Phase"},
            "action_id": "trk_complete_task",
            "value": str(task_id)
        }
    ]

    # Same interruption route as the working card: packing can be picked up
    # from a paused sheeting card too.
    if phase in ("field_sheeting", "border_sheeting") and not updated_task.get("packing_finished"):
        buttons.append({
            "type": "button",
            "text": {"type": "plain_text", "text": "Start Packing"},
            "action_id": "trk_start_packing",
            "value": str(task_id)
        })

    # A job marked No Border can still turn out to need one, even after some
    # packing has been done - the way back stays open until packing is
    # finished. The button lives on the paused card, not the running one,
    # because the correction is refused while the timer is going: stopping is
    # what reopens it, and a button that always refuses is worse than none.
    if phase == "packing" and updated_task.get("border_skipped") and not updated_task.get("packing_finished"):
        buttons.append({
            "type": "button",
            "text": {"type": "plain_text", "text": "Border after all"},
            "action_id": "trk_undo_no_border",
            "value": str(task_id)
        })

    # Every working card offers Add Jig - from Border onwards the modal asks
    # which phase used it, so a late Field jig has a supported route in
    if phase in ("field_sheeting", "border_sheeting", "packing"):
        buttons.append({
            "type": "button",
            "text": {"type": "plain_text", "text": "Add Jig"},
            "action_id": "trk_add_jig",
            "value": str(task_id)
        })

    client.chat_update(
        channel=channel_id,
        ts=task["message_ts"],
        text=f"Task T-{task_id} has been paused.",
        blocks=[
            {
                "type": "section",
                "text": {
                    "type": "mrkdwn",
                    "text": card_text
                }
            },
            {
                "type": "actions",
                "block_id": f"task_actions_{task_id}",
                "elements": buttons
            }
        ]
    )

# Start Packing button - packing as an interruption of sheeting work
@app.action("trk_start_packing")
def handle_start_packing(ack, body, client):
    """
    The maker breaks off sheeting to pack for a while.

    One press: the sheeting timer stops, the packing timer starts, and the
    card shows both. The job itself stays on its sheeting phase - packing done
    this way is time against packing, not a decision that the sheeting is
    over - and "Back to ..." on the new card is how the maker returns.
    """
    ack()
    task_id = int(body["actions"][0]["value"])
    user_id = body["user"]["id"]
    task = database.get_task(task_id)
    channel_id = body["container"]["channel_id"]

    if task is None:
        client.chat_postEphemeral(
            channel=channel_id,
            user=user_id,
            text="Task not found. It may have been deleted."
        )
        return

    if task["user_id"] != user_id:
        client.chat_postEphemeral(
            channel=channel_id,
            user=user_id,
            text="You can only control your own tasks."
        )
        return

    outcome = database.start_packing(task_id)
    if outcome != "started":
        if outcome == "phase_already_complete":
            text = "Packing has already been finished on this job."
        elif outcome == "job_not_open":
            text = "This job is no longer open."
        else:
            text = "Packing could not be started, so nothing has been changed."
        client.chat_postEphemeral(channel=channel_id, user=user_id, text=text)
        return

    updated_task = database.get_task(task_id)
    phase = updated_task["current_phase"]
    packing_time = database.format_elapsed(updated_task["packing_elapsed"])

    if phase == "packing":
        # The job had already moved on to its packing phase on another surface,
        # so this press is an ordinary packing start - show the normal card.
        field_time = database.format_elapsed(updated_task["field_elapsed"])
        border_time = border_time_display(updated_task)
        named_field_jig_line = f"*Field Jig Size:* {updated_task['field_jigs']}\n" if updated_task["field_jigs"] else ""
        named_border_jig_line = f"*Border Jig Size:* {updated_task['border_jigs']}\n" if updated_task["border_jigs"] else ""
        client.chat_update(
            channel=channel_id,
            ts=task["message_ts"],
            text=f"Task T-{task_id} is now in progress.",
            blocks=[
                {
                    "type": "section",
                    "text": {
                        "type": "mrkdwn",
                        "text": (
                            f"*Phase 3/4: Packing - In Progress*\n"
                            f"*ID:* T-{task_id}\n"
                            f"*Customer:* {updated_task['customer_name']}\n"
                            f"*Invoice:* {updated_task['invoice_number']}\n"
                            f"*Task:* {updated_task['task_description']}\n"
                            f"{named_field_jig_line}"
                            f"{named_border_jig_line}"
                            f"*Created by:* <@{updated_task['user_id']}>\n"
                            f"*Field Sheeting Time:* {field_time}\n"
                            f"*Border Sheeting Time:* {border_time}\n"
                            f"*Status:* In Progress"
                        )
                    }
                },
                {
                    "type": "actions",
                    "block_id": f"task_actions_{task_id}",
                    "elements": [
                        {
                            "type": "button",
                            "text": {"type": "plain_text", "text": "Stop"},
                            "style": "danger",
                            "action_id": "trk_stop_task",
                            "value": str(task_id)
                        },
                        {
                            "type": "button",
                            "text": {"type": "plain_text", "text": "Complete Phase"},
                            "action_id": "trk_complete_task",
                            "value": str(task_id)
                        },
                        {
                            "type": "button",
                            "text": {"type": "plain_text", "text": "Add Jig"},
                            "action_id": "trk_add_jig",
                            "value": str(task_id)
                        }
                    ]
                }
            ]
        )
        return

    # The card the maker packs from. The waiting sheeting phase is named, and
    # both times sit side by side so it is clear which timer is going.
    if phase == "border_sheeting":
        sheet_name = "Border Sheeting"
        sheet_time = database.format_elapsed(updated_task["border_elapsed"])
    else:
        sheet_name = "Field Sheeting"
        sheet_time = database.format_elapsed(updated_task["field_elapsed"])

    client.chat_update(
        channel=channel_id,
        ts=task["message_ts"],
        text=f"Task T-{task_id}: packing now, {sheet_name.lower()} is waiting.",
        blocks=[
            {
                "type": "section",
                "text": {
                    "type": "mrkdwn",
                    "text": (
                        f"*Packing — In Progress*\n"
                        f"_{sheet_name} is paused while you pack. Its time is safe, "
                        f"and the job is still on that phase._\n"
                        f"*ID:* T-{task_id}\n"
                        f"*Customer:* {updated_task['customer_name']}\n"
                        f"*Invoice:* {updated_task['invoice_number']}\n"
                        f"*Task:* {updated_task['task_description']}\n"
                        f"*Created by:* <@{updated_task['user_id']}>\n"
                        f"*{sheet_name} Time So Far:* {sheet_time}\n"
                        f"*Packing Time So Far:* {packing_time}"
                    )
                }
            },
            {
                "type": "actions",
                "block_id": f"task_actions_{task_id}",
                "elements": [
                    {
                        "type": "button",
                        "text": {"type": "plain_text", "text": "Stop Packing"},
                        "style": "danger",
                        "action_id": "trk_stop_task",
                        "value": str(task_id)
                    },
                    {
                        "type": "button",
                        "text": {"type": "plain_text", "text": f"Back to {sheet_name}"},
                        "style": "primary",
                        "action_id": "trk_start_task",
                        "value": str(task_id)
                    }
                ]
            }
        ]
    )


# Complete Task Button
@app.action("trk_complete_task")
def handle_complete(ack, body, client):
    ack()
    task_id = int(body["actions"][0]["value"])
    user_id = body["user"]["id"]
    task = database.get_task(task_id)
    channel_id = body["container"]["channel_id"]
    
    if task is None:
        client.chat_postEphemeral(channel=channel_id, user=user_id, text="Task Not Found. It may have been deleted")
        return

    if task["user_id"] != user_id:
        client.chat_postEphemeral(
            channel=channel_id,
            user=user_id,
            text="You can only control your own tasks."
        )
        return

    outcome = database.complete_task(task_id)
    if outcome == "another_phase_running":
        # The packing timer is going - most likely this press came from an old
        # sheeting card while the maker is packing. Finishing the phase under
        # a running timer would strand it, so nothing has moved.
        client.chat_postEphemeral(
            channel=channel_id,
            user=user_id,
            text=("The packing timer is running on this job, so this phase has "
                  "not been completed. Stop the packing first, then press "
                  "Complete Phase again.")
        )
        return
    updated_task = database.get_task(task_id)
    phase = updated_task["current_phase"]
    
    if phase == "field_sheeting":
        field_time = database.format_elapsed(updated_task["field_elapsed"])
        metadata = json.dumps({"task_id": task_id, "dm_channel_id": channel_id, "team_channel_id": task["channel_id"]})
        
        client.chat_postMessage(
            channel=task["channel_id"],
            text=f" *T-{task_id} - Field Sheeting complete* | <@{user_id}> | Time: {field_time}"
        )
        
        client.views_open(
            trigger_id=body["trigger_id"],
            view={
                "type": "modal",
                "callback_id": "trk_border_modal",
                "title": {"type": "plain_text", "text": "Border Sheeting"},
                "submit": {"type": "plain_text", "text": "Start Border Phase"},
                "close": {"type": "plain_text", "text": "Cancel"},
                "private_metadata": metadata,
                "blocks": [
                    {
                        "type": "section",
                        "text": {"type": "mrkdwn", "text": f"*Field Sheeting complete!* Time logged: *{field_time}*\nNow enter the Border Sheeting details."}
                    },
                    {
                        "type": "input",
                        "block_id": "border_design_block",
                        "label": {"type": "plain_text", "text": "Border Design Name"},
                        
                        "element": {
                            "type": "plain_text_input",
                            "action_id": "border_design"
                        }
                    },
                    {
                        "type": "input",
                        "block_id": "border_diff_block",
                        "label": {"type": "plain_text", "text": "Border Difficulty"},
                        "element": {
                            "type": "plain_text_input",
                            "action_id": "border_difficulty",
                            "max_length": 2
                        }
                    },
                    # Same box as the Field one: usually a millimetre size,
                    # but "template" and split sizes are fine too. Optional.
                    {
                        "type": "input",
                        "block_id": "border_jig_block",
                        "optional": True,
                        "label": {"type": "plain_text", "text": "Jig Size (mm)"},
                        "element": {
                            "type": "plain_text_input",
                            "action_id": "border_jig",
                            "placeholder": {"type": "plain_text", "text": "e.g. 49.6 or template"}
                        }
                    },
                    # Some jobs genuinely have no border. This is the moment the
                    # maker knows that, so it is the moment they are asked - it
                    # is deliberately not on the intake form, where it would be
                    # one more thing to answer before the job can start.
                    {
                        "type": "actions",
                        "block_id": "no_border_block",
                        "elements": [
                            {
                                "type": "button",
                                "text": {"type": "plain_text", "text": "No Border"},
                                "action_id": "trk_no_border",
                                "value": str(task_id)
                            }
                        ]
                    }
                ]
            }
        )
    
    #To start packing phase automatically
    
    elif phase == "border_sheeting":
        updated_task = database.get_task(task_id)
        field_time = database.format_elapsed(updated_task["field_elapsed"])
        border_time = database.format_elapsed(updated_task["border_elapsed"])
        metadata = json.dumps({"task_id": task_id, "dm_channel_id": channel_id,"team_channel_id": task["channel_id"]})
        
        client.chat_postMessage(
            channel=task["channel_id"],
            text=f" *T -{task_id} - Border Sheeting complete* | <@{user_id}> | Time: {border_time}"
        )
        
        #Deleting old card (message) before posting new phase
        
        client.chat_delete(channel=channel_id, ts=task ["message_ts"])
        
        client.views_open(
            trigger_id=body["trigger_id"],
            view=packing_modal_view(
                metadata,
                (
                    f"*Border Sheeting Complete!*\n"
                    f"Field Sheeting Time: *{field_time}*\n"
                    f"Border Sheeting Time: *{border_time}*\n"
                    f"Click 'Start Packing Phase' when you're ready"
                )
            )
        )
    
    elif phase == "packing":
        packing_time = database.format_elapsed(updated_task["packing_elapsed"])
        metadata = json.dumps({"task_id":task_id, "dm_channel_id": channel_id, "team_channel_id": task["channel_id"]})
        
        client.views_open(
            trigger_id=body["trigger_id"],
            view={
                "type": "modal",
                "callback_id": "trk_notes_modal",
                "title": {"type": "plain_text", "text": "Job Notes (Phase 4)"},
                "submit": {"type": "plain_text", "text": "Complete Job"},
                "close": {"type": "plain_text", "text": "Cancel"},
                "private_metadata": metadata,
                "blocks": [
                    {
                        "type": "section",
                        "text": {
                            "type": "mrkdwn",
                            "text": f"*Packing complete!* Time logged: *{packing_time}*\nAdd any final notes before closing this job."
                        }
                    },
                    {
                        "type": "input",
                        "block_id": "notes_block",
                        "optional": True,
                        "label": {"type": "plain_text", "text": "General Notes"},
                        "element": {
                            "type": "plain_text_input",
                            "multiline": True,
                            "action_id": "general_notes"
                        }
                    },
                    {
                        "type": "input",
                        "block_id": "issues_block",
                        "optional": True,
                        "label": {"type": "plain_text", "text": "Issues Encountered"},
                        "element": {
                            "type": "plain_text_input",
                            "multiline": True,
                            "action_id": "issues"
                        }
                    }
                ]
            }
        )
        
@app.view("trk_border_modal")
def handle_border_submission(ack,body, client):
    ack()
    user_id = body["user"]["id"]
    vals = body["view"]["state"]["values"]
    metadata = json.loads(body["view"]["private_metadata"])
    task_id = metadata["task_id"]
    dm_channel_id = metadata["dm_channel_id"]
    team_channel_id = metadata["team_channel_id"]
    
# border details
    border_design = vals["border_design_block"]["border_design"]["value"]
    border_difficulty = vals["border_diff_block"]["border_difficulty"]["value"]
    border_jig = (vals["border_jig_block"]["border_jig"]["value"] or "").strip()

# Transitioning to border phase in the database
    try:
        database.move_to_border_phase(task_id, border_design, border_difficulty, border_jig)
    except database.TrackerRefused as refusal:
        if refusal.reason != "another_phase_running":
            raise
        # The packing timer is going, so the border cannot be put back
        # underneath it. Say so; the form can be submitted again once the
        # packing has been stopped.
        client.chat_postEphemeral(
            channel=dm_channel_id,
            user=user_id,
            text=("The packing timer is running on this job, so the border "
                  "details have not been saved. Stop the packing first, then "
                  "press Complete Phase to get this form back.")
        )
        return
    task = database.get_task(task_id)
    field_time = database.format_elapsed(task["field_elapsed"])
    border_jig_line = f"*Jig Size:* {border_jig}\n" if border_jig else ""
    # A job can already hold packing time here - packed during field work, or
    # packed before a missed border came to light. Shown so it is not "lost".
    packing_line = (
        f"*Packing Time So Far:* {database.format_elapsed(task['packing_elapsed'])}\n"
        if task["packing_elapsed"] else ""
    )
    
# Updating the DM card

    client.chat_delete(channel=dm_channel_id, ts=task["message_ts"])
    
# posting card to the channel
    result = client.chat_postMessage(
        channel=dm_channel_id,
        text=f"Task T -{task_id} has moved to Border Sheeting.",
        blocks=[
            {
                "type": "section",
                "text": {
                    "type": "mrkdwn",
                    "text": (
                        f"*Phase 2/4: Border Sheeting — Ready to Start*\n"
                        f"*ID:* T-{task_id}\n"
                        f"*Customer:* {task['customer_name']}\n"
                        f"*Invoice:* {task['invoice_number']}\n"
                        f"*Task:* {task['task_description']}\n"
                        f"*Border Design:* {border_design}\n"
                        f"*Border Difficulty:* {border_difficulty}\n"
                        f"{border_jig_line}"
                        f"*Created by:* <@{task['user_id']}>\n"
                        f"*Field Sheeting Time:* {field_time}\n"
                        f"{packing_line}"
                        f"*Status:* Created"
                    )
                }
            },
            {
                "type": "actions",
                "block_id": f"task_actions_{task_id}",
                "elements": [
                    {
                        "type": "button",
                        "text": {"type": "plain_text", "text": "Start"},
                        "style": "primary",
                        "action_id": "trk_start_task",
                        "value": str(task_id)
                    },
                    # The field sheets exist by now, so packing can genuinely
                    # cut in before border work has even started.
                    {
                        "type": "button",
                        "text": {"type": "plain_text", "text": "Start Packing"},
                        "action_id": "trk_start_packing",
                        "value": str(task_id)
                    }
                ]
            }
        ]
    )

    database.update_message_ts(task_id, result["channel"], result["ts"])

@app.action("trk_no_border")
def handle_no_border(ack, body, client):
    """
    The maker says this job has no border.

    Records the decision and turns the same modal into the packing one, so the
    job carries straight on. The DM card is deliberately NOT touched here: the
    cursor stays on field until the packing modal is submitted, so cancelling
    at this point leaves the maker exactly where they were, with the border
    decision still open on a live card.
    """
    ack()
    metadata = json.loads(body["view"]["private_metadata"])
    task_id = metadata["task_id"]
    dm_channel_id = metadata["dm_channel_id"]
    team_channel_id = metadata["team_channel_id"]
    user_id = body["user"]["id"]

    task = database.get_task(task_id)

    if task is None:
        client.chat_postEphemeral(
            channel=dm_channel_id,
            user=user_id,
            text="Task not found. It may have been deleted."
        )
        return

    if task["user_id"] != user_id:
        client.chat_postEphemeral(
            channel=dm_channel_id,
            user=user_id,
            text="You can only control your own tasks."
        )
        return

    database.skip_border_phase(task_id)

    updated_task = database.get_task(task_id)
    field_time = database.format_elapsed(updated_task["field_elapsed"])

    client.chat_postMessage(
        channel=team_channel_id,
        text=f" *T-{task_id} - No Border on this job* | <@{user_id}>"
    )

    client.views_update(
        view_id=body["view"]["id"],
        view=packing_modal_view(
            json.dumps(metadata),
            (
                f"*No Border on this job.*\n"
                f"Field Sheeting Time: *{field_time}*\n"
                f"Click 'Start Packing Phase' when you're ready.\n\n"
                f"_Chose this by mistake? Cancel, then press Complete Phase "
                f"again to get the border details form back._"
            )
        )
    )


@app.action("trk_undo_no_border")
def handle_undo_no_border(ack, body, client):
    """
    Take back a "No Border" from the packing card.

    The refusal is shown, never hidden. If the correction did not happen the
    card is left exactly as it was: rebuilding it as though it had worked would
    leave the maker believing they have a border phase back when the record
    still says the border was skipped.
    """
    ack()
    task_id = int(body["actions"][0]["value"])
    user_id = body["user"]["id"]
    channel_id = body["container"]["channel_id"]
    task = database.get_task(task_id)

    if task is None:
        client.chat_postEphemeral(
            channel=channel_id,
            user=user_id,
            text="Task not found. It may have been deleted."
        )
        return

    if task["user_id"] != user_id:
        client.chat_postEphemeral(
            channel=channel_id,
            user=user_id,
            text="You can only control your own tasks."
        )
        return

    outcome = database.revert_border_skip(task_id)

    if outcome != "reverted":
        if outcome == "border_skip_not_reversible":
            text = (
                "Packing has already been finished on this job, so the border "
                "cannot be reopened from here. Nothing has been changed. Add "
                "what happened to the job's notes and tell a supervisor."
            )
        elif outcome == "another_phase_running":
            text = (
                "The packing timer is running on this job. Stop it first, then "
                "press 'Border after all' again. Nothing has been changed."
            )
        elif outcome == "border_not_skipped":
            text = "This job is not marked 'No Border', so there is nothing to undo."
        else:
            text = "That could not be undone, so nothing has been changed."
        client.chat_postEphemeral(channel=channel_id, user=user_id, text=text)
        return

    updated_task = database.get_task(task_id)
    field_time = database.format_elapsed(updated_task["field_elapsed"])
    field_jig_line = f"*Jig Size:* {updated_task['field_jigs']}\n" if updated_task["field_jigs"] else ""
    # Packing may already have been worked before the border came to light.
    # That time is real and keeps counting toward the job - showing it here
    # says so, instead of leaving the maker wondering where it went.
    packing_line = (
        f"*Packing Time So Far:* {database.format_elapsed(updated_task['packing_elapsed'])}\n"
        if updated_task["packing_elapsed"] else ""
    )

    client.chat_delete(channel=channel_id, ts=task["message_ts"])

    result = client.chat_postMessage(
        channel=channel_id,
        text=f"Task T-{task_id} is back at the border details step.",
        blocks=[
            {
                "type": "section",
                "text": {
                    "type": "mrkdwn",
                    "text": (
                        f"*Phase 1/4: Field Sheeting — Complete*\n"
                        f"*ID:* T-{task_id}\n"
                        f"*Customer:* {updated_task['customer_name']}\n"
                        f"*Invoice:* {updated_task['invoice_number']}\n"
                        f"*Task:* {updated_task['task_description']}\n"
                        f"*Field Design:* {updated_task['field_design']}\n"
                        f"*Difficulty:* {updated_task['difficulty']}\n"
                        f"{field_jig_line}"
                        f"*Due:* {updated_task['due_date']}\n"
                        f"*Created by:* <@{updated_task['user_id']}>\n"
                        f"*Field Sheeting Time:* {field_time}\n"
                        f"{packing_line}"
                        f"*Status:* No Border undone — press Complete Phase for "
                        f"the border details"
                    )
                }
            },
            {
                "type": "actions",
                "block_id": f"task_actions_{task_id}",
                "elements": [
                    {
                        "type": "button",
                        "text": {"type": "plain_text", "text": "Complete Phase"},
                        "style": "primary",
                        "action_id": "trk_complete_task",
                        "value": str(task_id)
                    },
                    {
                        "type": "button",
                        "text": {"type": "plain_text", "text": "Add Jig"},
                        "action_id": "trk_add_jig",
                        "value": str(task_id)
                    }
                ]
            }
        ]
    )

    database.update_message_ts(task_id, result["channel"], result["ts"])


@app.view("trk_packing_modal")
def handle_packing_submission(ack, body, client):
    ack()
    user_id = body["user"]["id"]
    metadata = json.loads(body["view"]["private_metadata"])
    task_id = metadata["task_id"]
    dm_channel_id = metadata["dm_channel_id"]
    team_channel_id = metadata ["team_channel_id"]

    outcome = database.move_to_packing_phase(task_id)
    if outcome != "moved":
        # This form was opened before something else changed the job - most
        # likely the border was put back on another device. Say so and change
        # nothing; the job's card is still live and still correct.
        client.chat_postEphemeral(
            channel=dm_channel_id,
            user=user_id,
            text=("This job's border has changed since this form was opened, so it "
                  "has not been moved to Packing and nothing has been changed. "
                  "Go back to the job's card and carry on from there.")
        )
        return

    task = database.get_task(task_id)
    field_time = database.format_elapsed(task["field_elapsed"])
    border_time = border_time_display(task)
    # Packing worked earlier as an interruption arrives here with time already
    # on the clock. Shown, so the maker knows it counted.
    packing_line = (
        f"*Packing Time So Far:* {database.format_elapsed(task['packing_elapsed'])}\n"
        if task["packing_elapsed"] else ""
    )

    # The border route deleted the old card before opening this modal. The No
    # Border route deliberately did not, so that cancelling left the maker on a
    # live card with the decision still open - so clear it here, now that they
    # have committed to packing.
    if task.get("border_skipped") and task.get("message_ts"):
        try:
            client.chat_delete(channel=dm_channel_id, ts=task["message_ts"])
        except SlackApiError:
            # Already gone is the outcome this wanted anyway.
            pass

    # The way back, offered only while there is a way back. A phase can wait
    # paused while other work happens, so packing time on the clock does not
    # close the border question - only FINISHING the packing does. Until then
    # the button stays, and a button that always refuses is worse than no
    # button.
    packing_actions = [
        {
            "type": "button",
            "text": {"type": "plain_text", "text": "Start"},
            "style": "primary",
            "action_id": "trk_start_task",
            "value": str(task_id)
        }
    ]
    if task.get("border_skipped") and not task.get("packing_finished"):
        packing_actions.append({
            "type": "button",
            "text": {"type": "plain_text", "text": "Border after all"},
            "action_id": "trk_undo_no_border",
            "value": str(task_id)
        })

    result = client.chat_postMessage(
        channel=dm_channel_id,
        text = f"Task T-{task_id} has moved to Packing.",
        blocks=[
            {
                "type": "section",
                "text": {
                    "type": "mrkdwn",
                    "text": (
                        f"*Phase 3/4: Packing — Ready to Start*\n"
                        f"*ID:* T-{task_id}\n"
                        f"*Customer:* {task['customer_name']}\n"
                        f"*Invoice:* {task['invoice_number']}\n"
                        f"*Task:* {task['task_description']}\n"
                        f"*Field Sheeting Time:* {field_time}\n"
                        f"*Border Sheeting Time:* {border_time}\n"
                        f"{packing_line}"
                        f"*Created by:* <@{task['user_id']}>\n"
                        f"*Status:* Created"
                    )
                }
            },
            {
                "type": "actions",
                "block_id": f"task_actions_{task_id}",
                "elements": packing_actions
            }
        ]
    )
    
    database.update_message_ts(task_id, result["channel"], result["ts"])
    
@app.view("trk_notes_modal")
def handle_notes_submission(ack,body,client):
    ack()
    user_id = body["user"]["id"]
    vals = body["view"]["state"]["values"]
    metadata = json.loads(body["view"]["private_metadata"])
    task_id = metadata["task_id"]
    dm_channel_id = metadata["dm_channel_id"]
    team_channel_id = metadata["team_channel_id"]
    
    general_notes = vals["notes_block"]["general_notes"]["value"] or "None"
    issues = vals["issues_block"]["issues"]["value"] or "None"
    
    database.save_notes_and_complete(task_id, general_notes, issues)
    task = database.get_task(task_id)
    
# Calculating all phase times and overall time

    elapsed = database.get_phase_elapsed(task_id)
    field_time = database.format_elapsed(elapsed["field_elapsed"])
    border_time = border_time_display(task)
    packing_time = database.format_elapsed(elapsed["packing_elapsed"])
    total_time = database.format_elapsed(elapsed["total_elapsed"])

    # Jig lines for the summary, shown only when a jig was recorded
    field_jig_line = f"*Field Jig Size:* {task['field_jigs']}\n" if task["field_jigs"] else ""
    border_jig_line = f"*Border Jig Size:* {task['border_jigs']}\n" if task["border_jigs"] else ""
    
# Deleting packing card before posting final summary
    client.chat_update(
        channel = dm_channel_id,
        ts=task["message_ts"],
        text=f" Job T-{task_id} is complete. Summary posted to the team channel.",
        blocks =[
            {
                "type": "section",
                "text": {
                    "type": "mrkdwn",
                    "text": (
                        f"✅ *Job Complete — T-{task_id}*\n"
                        f"*Customer:* {task['customer_name']}\n"
                        f"Total Time: {total_time}\n"
                        f"The full summary has been posted to the team channel"
                    )
                }
            }
        ]
    )
    
    client.chat_postMessage(
        channel=team_channel_id,
        text=f" Job T-{task_id} fully completed by <@{user_id}>",
        blocks =[
            {
                "type": "section",
                "text": {
                    "type": "mrkdwn",
                    "text": (
                        f"✅ *Job Complete — T-{task_id}*\n"
                        f"*Customer:* {task['customer_name']}\n"
                        f"*Invoice:* {task['invoice_number']}\n"
                        f"*Task:* {task['task_description']}\n"
                        f"{field_jig_line}"
                        f"{border_jig_line}"
                        f"*Completed by:* <@{user_id}>\n\n"
                        f"*Phase Breakdown:*\n"
                        f"🟦 Field Sheeting: {field_time}\n"
                        f"🟨 Border Sheeting: {border_time}\n"
                        f"📦 Packing: {packing_time}\n\n"
                        f"⏱️ *Total Time: {total_time}*\n\n"
                        f"*General Notes:* {general_notes}\n"
                        f"*Issues Encountered:* {issues}"
                    )
                }
            }
        ]
    )
    

        
# Add Jig button - a phase sometimes needs another jig part way through.
# Maybe the first jig turned out wrong and was swapped, maybe two sizes are
# genuinely needed together. Either way the earlier jig really was used, so
# this ADDS a record next to it - it never overwrites one. Typing mistakes
# are fixed through Edit instead, which changes the value it names.

@app.action("trk_add_jig")
def handle_add_jig(ack, body, client):
    ack()
    task_id = int(body["actions"][0]["value"])
    user_id = body["user"]["id"]
    task = database.get_task(task_id)
    channel_id = body["container"]["channel_id"]

    if task is None:
        client.chat_postEphemeral(
            channel=channel_id,
            user=user_id,
            text="Task not found. It may have been deleted."
        )
        return

    if task["user_id"] != user_id:
        client.chat_postEphemeral(
            channel=channel_id,
            user=user_id,
            text="You can only control your own tasks."
        )
        return

    # Jigs can be added as long as the job is open. A finished job has no
    # card with this button, but a stale one could still be clicked.
    phase = task["current_phase"]
    if phase == "completed":
        client.chat_postEphemeral(
            channel=channel_id,
            user=user_id,
            text="This job is finished - jigs can no longer be added."
        )
        return

    # Remember which phase the card was on, so the right card comes back
    # after the modal
    metadata = json.dumps({"task_id": task_id, "channel_id": channel_id, "phase": phase})

    # During Field there is only one place a jig can go. From Border onwards
    # the maker chooses, because sometimes Field work genuinely has to
    # continue with another jig AFTER Field was completed - that is still an
    # added jig, not a correction, and the earlier one must stay. The Border
    # phase is offered once the maker has reached it.
    blocks = []
    if phase != "field_sheeting":
        field_option = {
            "text": {"type": "plain_text", "text": "Field"},
            "value": "field_sheeting"
        }
        border_option = {
            "text": {"type": "plain_text", "text": "Border"},
            "value": "border_sheeting"
        }
        # A border that did not happen used no jig, and storage refuses one,
        # so it is not offered - a choice that always errors is worse than no
        # choice.
        phase_element = {
            "type": "static_select",
            "action_id": "jig_phase",
            "placeholder": {"type": "plain_text", "text": "Field or Border?"},
            "options": [field_option] if task.get("border_skipped") else [field_option, border_option]
        }
        if task.get("border_skipped"):
            phase_element["initial_option"] = field_option
        # On the Border card the border is the usual answer, so it is
        # pre-picked; from Packing there is no obvious answer, so the maker
        # must choose
        if phase == "border_sheeting":
            phase_element["initial_option"] = border_option
        blocks.append({
            "type": "input",
            "block_id": "phase_block",
            "label": {"type": "plain_text", "text": "Which phase used it?"},
            "element": phase_element
        })
    blocks.append({
        "type": "input",
        "block_id": "jig_block",
        "label": {"type": "plain_text", "text": "Jig Size (mm)"},
        "element": {
            "type": "plain_text_input",
            "action_id": "jig_size",
            "placeholder": {"type": "plain_text", "text": "e.g. 49.6 or template"}
        }
    })

    client.views_open(
        trigger_id=body["trigger_id"],
        view={
            "type": "modal",
            "callback_id": "trk_add_jig_modal",
            "title": {"type": "plain_text", "text": "Add Jig"},
            "submit": {"type": "plain_text", "text": "Add"},
            "close": {"type": "plain_text", "text": "Cancel"},
            "private_metadata": metadata,
            "blocks": blocks
        }
    )

@app.view("trk_add_jig_modal")
def handle_add_jig_submission(ack, body, client):
    ack()
    user_id = body["user"]["id"]
    vals = body["view"]["state"]["values"]
    metadata = json.loads(body["view"]["private_metadata"])
    task_id = metadata["task_id"]
    channel_id = metadata["channel_id"]
    phase = metadata["phase"]

    jig_size = (vals["jig_block"]["jig_size"]["value"] or "").strip()
    if not jig_size:
        return

    # The dropdown says which phase used the jig; when the modal had no
    # dropdown the job was still on Field, so Field it is
    if "phase_block" in vals:
        target_phase = vals["phase_block"]["jig_phase"]["selected_option"]["value"]
    else:
        target_phase = "field_sheeting"

    database.add_jig(task_id, target_phase, jig_size)
    task = database.get_task(task_id)

    # The job can disappear between opening the modal and submitting it -
    # deleted, or finished. Tell the maker instead of closing the modal in
    # silence with their jig unrecorded, and leave the final card alone.
    if task is None or task["current_phase"] == "completed":
        client.chat_postEphemeral(
            channel=channel_id,
            user=user_id,
            text=f"Jig '{jig_size}' was not recorded - the task is no longer open."
        )
        return

    # Refresh the card so the maker sees every jig recorded so far
    field_jig_line = f"*Jig Size:* {task['field_jigs']}\n" if task["field_jigs"] else ""
    border_jig_line = f"*Jig Size:* {task['border_jigs']}\n" if task["border_jigs"] else ""
    running = task["status"] == "in_progress"
    status_text = "In Progress" if running else "Paused"

    if phase == "field_sheeting":
        card_text = (
            f"*Phase 1/4: Field Sheeting — {status_text}*\n"
            f"*ID:* T-{task_id}\n"
            f"*Customer:* {task['customer_name']}\n"
            f"*Invoice:* {task['invoice_number']}\n"
            f"*Task:* {task['task_description']}\n"
            f"*Field Design:* {task['field_design']}\n"
            f"*Difficulty:* {task['difficulty']}\n"
            f"{field_jig_line}"
            f"*Due:* {task['due_date']}\n"
            f"*Created by:* <@{task['user_id']}>\n"
            f"*Status:* {status_text}"
        )
    elif phase == "border_sheeting":
        field_time = database.format_elapsed(task["field_elapsed"])
        card_text = (
            f"*Phase 2/4: Border Sheeting — {status_text}*\n"
            f"*ID:* T-{task_id}\n"
            f"*Customer:* {task['customer_name']}\n"
            f"*Invoice:* {task['invoice_number']}\n"
            f"*Task:* {task['task_description']}\n"
            f"*Border Design:* {task['border_design']}\n"
            f"*Border Difficulty:* {task['border_difficulty']}\n"
            f"{border_jig_line}"
            f"*Created by:* <@{task['user_id']}>\n"
            f"*Field Sheeting Time:* {field_time}\n"
            f"*Status:* {status_text}"
        )
    else:
        # The packing card is not a jig phase itself, so its lines say which
        # phase each jig belongs to
        field_time = database.format_elapsed(task["field_elapsed"])
        border_time = border_time_display(task)
        named_field_line = f"*Field Jig Size:* {task['field_jigs']}\n" if task["field_jigs"] else ""
        named_border_line = f"*Border Jig Size:* {task['border_jigs']}\n" if task["border_jigs"] else ""
        card_text = (
            f"*Phase 3/4: Packing — {status_text}*\n"
            f"*ID:* T-{task_id}\n"
            f"*Customer:* {task['customer_name']}\n"
            f"*Invoice:* {task['invoice_number']}\n"
            f"*Task:* {task['task_description']}\n"
            f"{named_field_line}"
            f"{named_border_line}"
            f"*Created by:* <@{task['user_id']}>\n"
            f"*Field Sheeting Time:* {field_time}\n"
            f"*Border Sheeting Time:* {border_time}\n"
            f"*Status:* {status_text}"
        )

    # Same buttons the card had before the modal opened
    if running:
        buttons = [
            {
                "type": "button",
                "text": {"type": "plain_text", "text": "Stop"},
                "style": "danger",
                "action_id": "trk_stop_task",
                "value": str(task_id)
            },
            {
                "type": "button",
                "text": {"type": "plain_text", "text": "Complete Phase"},
                "action_id": "trk_complete_task",
                "value": str(task_id)
            }
        ]
    else:
        buttons = [
            {
                "type": "button",
                "text": {"type": "plain_text", "text": "Resume"},
                "style": "primary",
                "action_id": "trk_start_task",
                "value": str(task_id)
            },
            {
                "type": "button",
                "text": {"type": "plain_text", "text": "Edit"},
                "action_id": "trk_edit_task",
                "value": str(task_id)
            },
            {
                "type": "button",
                "text": {"type": "plain_text", "text": "Complete Phase"},
                "action_id": "trk_complete_task",
                "value": str(task_id)
            }
        ]
    # Keep Start Packing through a jig add, the same way Add Jig itself is
    # kept - a recorded jig must not cost the card a button it had.
    if phase in ("field_sheeting", "border_sheeting") and not task.get("packing_finished"):
        buttons.append({
            "type": "button",
            "text": {"type": "plain_text", "text": "Start Packing"},
            "action_id": "trk_start_packing",
            "value": str(task_id)
        })
    # And keep the way back on a paused No Border packing card, for the same
    # reason.
    if phase == "packing" and not running and task.get("border_skipped") and not task.get("packing_finished"):
        buttons.append({
            "type": "button",
            "text": {"type": "plain_text", "text": "Border after all"},
            "action_id": "trk_undo_no_border",
            "value": str(task_id)
        })
    buttons.append({
        "type": "button",
        "text": {"type": "plain_text", "text": "Add Jig"},
        "action_id": "trk_add_jig",
        "value": str(task_id)
    })

    client.chat_update(
        channel=channel_id,
        ts=task["message_ts"],
        text=f"Task T-{task_id}: jig {jig_size} recorded.",
        blocks=[
            {
                "type": "section",
                "text": {"type": "mrkdwn", "text": card_text}
            },
            {
                "type": "actions",
                "block_id": f"task_actions_{task_id}",
                "elements": buttons
            }
        ]
    )

#Delete Button
@app.action("trk_delete_task")
def handle_delete(ack, body, client):
    ack()
    task_id = int(body["actions"][0]["value"])
    user_id = body["user"]["id"]
    task = database.get_task(task_id)
    channel_id = body["container"]["channel_id"]

    if task is None:
        client.chat_postEphemeral(
            channel=channel_id,
            user=user_id,
            text="Task not found. It may have already been deleted."
        )
        return

    # Only the creator can delete
    if task["user_id"] != user_id:
        client.chat_postEphemeral(
            channel=channel_id,
            user=user_id,
            text="You can only delete your own tasks."
        )
        return

    # Delete from the database
    database.delete_task(task_id)

    # Replace the card with a simple deleted message
    client.chat_update(
        channel=channel_id,
        ts=task["message_ts"],
        text=f"Task T-{task_id} has been deleted by <@{user_id}>.",
        blocks=[
            {
                "type": "section",
                "text": {
                    "type": "mrkdwn",
                    "text": f"❌ *Task T-{task_id} has been deleted by <@{user_id}>.*"
                }
            }
        ]
    )


@app.action("trk_edit_task")
def handle_edit(ack, body, client):
    ack()
    task_id = int(body["actions"][0]["value"])
    user_id = body["user"]["id"]
    task = database.get_task(task_id)
    channel_id = body["container"]["channel_id"]

    if task is None:
        client.chat_postEphemeral(
            channel=channel_id,
            user=user_id,
            text="Task not found. It may have been deleted."
        )
        return

    if task["user_id"] != user_id:
        client.chat_postEphemeral(
            channel=channel_id,
            user=user_id,
            text="You can only edit your own tasks."
        )
        return

    # Bundle task_id and channel_id to pass through the modal
    edit_metadata = json.dumps({
        "task_id": task_id,
        "channel_id": channel_id
    })

    # One box per jig already recorded, pre-filled, so a mistyped value can
    # be corrected later - even after that phase has finished. These boxes
    # fix EXISTING jigs; a genuinely new jig goes through Add Jig instead.
    jig_blocks = []
    for i, rec in enumerate(task["field_jig_records"], start=1):
        jig_blocks.append({
            "type": "input",
            "block_id": f"jig_edit_{rec['id']}",
            "label": {"type": "plain_text", "text": f"Field Jig {i} (mm)"},
            "element": {
                "type": "plain_text_input",
                "action_id": "jig_value",
                "initial_value": rec["value"]
            }
        })
    for i, rec in enumerate(task["border_jig_records"], start=1):
        jig_blocks.append({
            "type": "input",
            "block_id": f"jig_edit_{rec['id']}",
            "label": {"type": "plain_text", "text": f"Border Jig {i} (mm)"},
            "element": {
                "type": "plain_text_input",
                "action_id": "jig_value",
                "initial_value": rec["value"]
            }
        })

    # Open pre-filled edit modal
    client.views_open(
        trigger_id=body["trigger_id"],
        view={
            "type": "modal",
            "callback_id": "trk_edit_task_modal",
            "title": {"type": "plain_text", "text": "Edit Task"},
            "submit": {"type": "plain_text", "text": "Save Changes"},
            "close": {"type": "plain_text", "text": "Cancel"},
            "private_metadata": edit_metadata,
            "blocks": [
                {
                    "type": "input",
                    "block_id": "customer_block",
                    "label": {"type": "plain_text", "text": "Customer Name"},
                    "element": {
                        "type": "plain_text_input",
                        "action_id": "customer_name",
                        "initial_value": task["customer_name"]
                    }
                },
                {
                    "type": "input",
                    "block_id": "invoice_block",
                    "label": {"type": "plain_text", "text": "Invoice Number"},
                    "element": {
                        "type": "plain_text_input",
                        "action_id": "invoice_num",
                        "initial_value": task["invoice_number"]
                    }
                },
                {
                    "type": "input",
                    "block_id": "task_block",
                    "label": {"type": "plain_text", "text": "Task Description"},
                    "element": {
                        "type": "plain_text_input",
                        "multiline": True,
                        "action_id": "task_desc",
                        "initial_value": task["task_description"]
                    }
                },
                {
                    "type": "input",
                    "block_id": "design_block",
                    "label": {"type": "plain_text", "text": "Field Design Name"},
                    "element": {
                        "type": "plain_text_input",
                        "action_id": "design",
                        "initial_value": task["field_design"] or ""
                    }
                },
                {
                    "type": "input",
                    "block_id": "difficulty_block",
                    "label": {"type": "plain_text", "text": "Sheeting Difficulty"},
                    "element": {
                        "type": "plain_text_input",
                        "action_id": "difficulty",
                        "max_length": 2,
                        "initial_value": task["difficulty"] or ""
                    }
                },
                *jig_blocks,
                {
                    "type": "input",
                    "block_id": "date_block",
                    "optional": True,
                    "label": {"type": "plain_text", "text": "Due Date (DD/MM/YYYY)"},
                    "element": {
                        "type": "plain_text_input",
                        "action_id": "due_date",
                        "initial_value": task["due_date"] if task["due_date"] != "N/A" else ""
                    }
                }
            ]
        }
    )


@app.view("trk_edit_task_modal")
def handle_edit_submission(ack, body, client):
    ack()
    user_id = body["user"]["id"]
    vals = body["view"]["state"]["values"]

    # Retrieve task_id and channel_id from metadata
    metadata = json.loads(body["view"]["private_metadata"])
    task_id = metadata["task_id"]
    channel_id = metadata["channel_id"]

    # Collect updated values
    customer_name = vals["customer_block"]["customer_name"]["value"]
    invoice_number = vals["invoice_block"]["invoice_num"]["value"]
    task_description = vals["task_block"]["task_desc"]["value"]
    design = vals["design_block"]["design"]["value"]
    difficulty = vals["difficulty_block"]["difficulty"]["value"]
    due_date = vals["date_block"]["due_date"]["value"] or "N/A"

    # What each jig said before the modal opened, so only boxes the maker
    # actually changed get corrected
    task_before = database.get_task(task_id)
    previous_jigs = {}
    if task_before is not None:
        for rec in task_before["field_jig_records"] + task_before["border_jig_records"]:
            previous_jigs[rec["id"]] = rec["value"]

    # Save to database
    database.update_task(
        task_id=task_id,
        customer=customer_name,
        invoice=invoice_number,
        task_desc=task_description,
        design=design,
        difficulty=difficulty,
        due_date=due_date
    )

    # Fix any jig boxes the maker changed. Each correction names its own
    # record, so fixing one jig never touches the others.
    for block_id, entry in vals.items():
        if not block_id.startswith("jig_edit_"):
            continue
        jig_id = block_id[len("jig_edit_"):]
        new_value = (entry["jig_value"]["value"] or "").strip()
        if new_value and new_value != previous_jigs.get(jig_id):
            database.correct_jig(task_id, jig_id, new_value)

    # Fetch updated task to refresh the card
    task = database.get_task(task_id)

    # Rebuild the card based on current status
    if task["status"] == "created":
        buttons = [
            {
                "type": "button",
                "text": {"type": "plain_text", "text": "Start"},
                "style": "primary",
                "action_id": "trk_start_task",
                "value": str(task_id)
            },
            {
                "type": "button",
                "text": {"type": "plain_text", "text": "Edit"},
                "action_id": "trk_edit_task",
                "value": str(task_id)
            },
            {
                "type": "button",
                "text": {"type": "plain_text", "text": "Delete"},
                "style": "danger",
                "action_id": "trk_delete_task",
                "value": str(task_id),
                "confirm": {
                        "title": {"type": "plain_text", "text": "Delete Task?"},
                        "text": {"type": "plain_text", "text": "Are you sure you want to delete this task? This cannot be undone."},
                        "confirm": {"type": "plain_text", "text": "Yes, Delete"},
                        "deny": {"type": "plain_text", "text": "Cancel"}
                        }
            }
        ]
        status_text = "Created"
    else:
        buttons = [
            {
                "type": "button",
                "text": {"type": "plain_text", "text": "Resume"},
                "style": "primary",
                "action_id": "trk_start_task",
                "value": str(task_id)
            },
            {
                "type": "button",
                "text": {"type": "plain_text", "text": "Edit"},
                "action_id": "trk_edit_task",
                "value": str(task_id)
            },
            {
                "type": "button",
                "text": {"type": "plain_text", "text": "Complete"},
                "action_id": "trk_complete_task",
                "value": str(task_id)
            }
        ]
        # Keep Start Packing through an edit for the same reason Add Jig is
        # kept below - saving an edit must not cost the card a button it had.
        if task["current_phase"] in ("field_sheeting", "border_sheeting") and not task.get("packing_finished"):
            buttons.append({
                "type": "button",
                "text": {"type": "plain_text", "text": "Start Packing"},
                "action_id": "trk_start_packing",
                "value": str(task_id)
            })
        # And the way back on a paused No Border packing card, likewise.
        if task["current_phase"] == "packing" and task.get("border_skipped") and not task.get("packing_finished"):
            buttons.append({
                "type": "button",
                "text": {"type": "plain_text", "text": "Border after all"},
                "action_id": "trk_undo_no_border",
                "value": str(task_id)
            })
        # Keep the Add Jig button through an edit - without this, saving an
        # edit on a paused sheeting card would silently drop it until the
        # next stop or resume
        if task["current_phase"] in ("field_sheeting", "border_sheeting", "packing"):
            buttons.append({
                "type": "button",
                "text": {"type": "plain_text", "text": "Add Jig"},
                "action_id": "trk_add_jig",
                "value": str(task_id)
            })
        status_text = "🟠 Paused"

    # Jig lines refreshed from the database, so they show any corrections.
    # This card is not phase-specific, so each line names its phase - a
    # border job's jig must not look wiped just because an edit was saved.
    field_jig_line = f"*Field Jig Size:* {task['field_jigs']}\n" if task["field_jigs"] else ""
    border_jig_line = f"*Border Jig Size:* {task['border_jigs']}\n" if task["border_jigs"] else ""

    client.chat_update(
        channel=channel_id,
        ts=task["message_ts"],
        text=f"Task T-{task_id} has been updated.",
        blocks=[
            {
                "type": "section",
                "text": {
                    "type": "mrkdwn",
                    "text": (
                        f"*Task Updated ✏️ - Phase 1/4: Field Sheeting*\n"
                        f"*ID:* T-{task_id}\n"
                        f"*Customer:* {customer_name}\n"
                        f"*Invoice:* {invoice_number}\n"
                        f"*Task:* {task_description}\n"
                        f"*Field Design:* {design}\n"
                        f"*Difficulty:* {difficulty}\n"
                        f"{field_jig_line}"
                        f"{border_jig_line}"
                        f"*Due:* {due_date}\n"
                        f"*Created by:* <@{task['user_id']}>\n"
                        f"*Status:* {status_text}"
                    )
                }
            },
            {
                "type": "actions",
                "block_id": f"task_actions_{task_id}",
                "elements": buttons
            }
        ]
    )   
if __name__ == "__main__":
    database.setup_database()
    handler = SocketModeHandler(app, os.environ.get("SLACK_APP_TOKEN"))
    print("Trackbot is running!")
    handler.start()
