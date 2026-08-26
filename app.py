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
            task["border_design"] or "-",
            task["border_difficulty"] or "-",
            task["border_jigs"] or "-",
            database.format_elapsed(border_elapsed),
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
        border_time = database.format_elapsed(task["border_elapsed"])
        card_text = (
            f"*Phase 3/4: Packing - In Progress*\n"
            f"*ID:* T-{task_id}\n"
            f"*Customer:* {task['customer_name']}\n"
            f"*Invoice:* {task['invoice_number']}\n"
            f"*Task:* {task['task_description']}\n"
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

    # Only the sheeting phases use a jig, so only their cards offer the button
    if phase in ("field_sheeting", "border_sheeting"):
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
            f"*Border Time So Far:* {elapsed}"
        )
    
    else:
        elapsed = database.format_elapsed(updated_task["packing_elapsed"])
        card_text = (
            f"*Phase 3/4: Packing — Paused*\n"
            f"*ID:* T-{task_id}\n"
            f"*Customer:* {task['customer_name']}\n"
            f"*Invoice:* {task['invoice_number']}\n"
            f"*Task:* {task['task_description']}\n"
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

    # Only the sheeting phases use a jig, so only their cards offer the button
    if phase in ("field_sheeting", "border_sheeting"):
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

    database.complete_task(task_id)
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
            view={
                    "type": "modal",
                    "callback_id": "trk_packing_modal",
                "title": {"type": "plain_text", "text": "Packing (Phase 3)"},
                "submit": {"type": "plain_text", "text": "Start Packing Phase"},
                "close": {"type": "plain_text", "text": "Cancel"},
                "private_metadata": metadata,
                "blocks":[
                    {
                        "type": "section",
                        "text": {
                            "type": "mrkdwn",
                            "text": (
                                f"*Border Sheeting Complete!*\n"
                                f"Field Sheeting Time: *{field_time}*\n"
                                f"Border Sheeting Time: *{border_time}*\n"
                                f"Click 'Start Packing Phase' when you're ready"
                            )
                        }
                    }
                ]
            }
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
    database.move_to_border_phase(task_id, border_design, border_difficulty, border_jig)
    task = database.get_task(task_id)
    field_time = database.format_elapsed(task["field_elapsed"])
    border_jig_line = f"*Jig Size:* {border_jig}\n" if border_jig else ""
    
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

    database.move_to_packing_phase(task_id)
    task = database.get_task(task_id)
    field_time = database.format_elapsed(task["field_elapsed"])
    border_time = database.format_elapsed(task["border_elapsed"])

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
                        f"*Created by:* <@{task['user_id']}>\n"
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
                    }
                ]
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
    border_time = database.format_elapsed(elapsed["border_elapsed"])
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

    # The button lives on the Field and Border cards, but a stale card could
    # still be clicked after the job moved on - packing has no jig.
    phase = task["current_phase"]
    if phase not in ("field_sheeting", "border_sheeting"):
        client.chat_postEphemeral(
            channel=channel_id,
            user=user_id,
            text="Jigs belong to the Field or Border phase - this job has moved past those."
        )
        return

    # Remember which phase the button was pressed from, so the new jig lands
    # on the right one
    metadata = json.dumps({"task_id": task_id, "channel_id": channel_id, "phase": phase})

    client.views_open(
        trigger_id=body["trigger_id"],
        view={
            "type": "modal",
            "callback_id": "trk_add_jig_modal",
            "title": {"type": "plain_text", "text": "Add Jig"},
            "submit": {"type": "plain_text", "text": "Add"},
            "close": {"type": "plain_text", "text": "Cancel"},
            "private_metadata": metadata,
            "blocks": [
                {
                    "type": "input",
                    "block_id": "jig_block",
                    "label": {"type": "plain_text", "text": "Jig Size (mm)"},
                    "element": {
                        "type": "plain_text_input",
                        "action_id": "jig_size",
                        "placeholder": {"type": "plain_text", "text": "e.g. 49.6 or template"}
                    }
                }
            ]
        }
    )

@app.view("trk_add_jig_modal")
def handle_add_jig_submission(ack, body, client):
    ack()
    vals = body["view"]["state"]["values"]
    metadata = json.loads(body["view"]["private_metadata"])
    task_id = metadata["task_id"]
    channel_id = metadata["channel_id"]
    phase = metadata["phase"]

    jig_size = (vals["jig_block"]["jig_size"]["value"] or "").strip()
    if not jig_size:
        return

    database.add_jig(task_id, phase, jig_size)
    task = database.get_task(task_id)
    if task is None:
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
    else:
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
        status_text = "🟠 Paused"

    # Jig line refreshed from the database, so it shows any corrections
    field_jig_line = f"*Jig Size:* {task['field_jigs']}\n" if task["field_jigs"] else ""

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
