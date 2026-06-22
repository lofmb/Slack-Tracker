import os
import json
from dotenv import load_dotenv
from slack_bolt import App
from slack_bolt.adapter.socket_mode import SocketModeHandler

# Importing database functions

import database

# Loading environment variables from .env
load_dotenv()

# Initialsing Slack with the Bot Token and Signing Secret

app = App(
    token=os.environ.get("SLACK_bot_token"),
    signing_secret=os.environ.get("SLACK_SIGNING_SECRET")
)

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
            "callback_id": "track_step_1",  # ID used to catch the submission
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

#Step 1 Submission - This stage collects data and pushes to the Step 2 Modal
@app.view("track_step_1")
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
        "callback_id": "track_step_2",
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
                }
        ]
    }
    )

@app.view("track_step_2")
def handle_step_2(ack, body, client):
    ack(response_action = "clear")
    user_id = body["user"]["id"]
    vals = body["view"]["state"]["values"]
    
    # Fetching the Step 1 data from the metadata
    prev_data = json.loads(body["view"]["private_metadata"])
    channel_id = prev_data["channel_id"]
    
    # Collecting Step 2 values
    design = vals["design"]["val"]["value"]
    difficulty = vals["diff"]["difficulty"]["value"]
    
    # Saving the task to the database
    task_id = database.create_task(
        user_id=user_id,
        channel_id=channel_id,
        customer_name=prev_data["customer_name"],
        invoice_number=prev_data["invoice_number"],
        task_description=prev_data["task_description"],
        due_date=prev_data["due_date"],
        is_na=prev_data["is_na"],
        design=design,
        difficulty=difficulty
    )
    
    # Displaying due date on card
    due_display = "N/A" if prev_data["is_na"] else prev_data["due_date"]
    
    #Posting the task card to channel
    result = client.chat_postMessage(
        channel=channel_id,
        text=f"New Task -{task_id} created by <@{user_id}>",
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
                        f"*Due:*\n{due_display}\n"
                        f"*Created by:*\n<@{user_id}>\n"
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
                        "action_id": "start_task",
                        "value": str(task_id)
                    },
                    {
                        "type": "button",
                        "text": {"type": "plain_text", "text": "Edit"},
                        "action_id": "edit_task",
                        "value": str(task_id) 
                    },
                    {
                        "type": "button",
                        "text": {"type": "plain_text", "text": "Delete"},
                        "style": "danger",
                        "action_id": "delete_task",
                        "value":str(task_id)
                    }
                ]
            }
        ]
    )
    
    # Saving the timestamp
    database.update_message_ts(task_id, result["ts"])

@app.action("start_task")
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
    
    if phase == "field_sheeting":
        card_text = (
            f"* Phase 1/4: Field Sheeting - In Progress*\n"
            f"*ID: T-{task_id}\n"
            f"*Customer:* {task['customer_name']}\n"
            f"Invoice:* {task['invoice_number']}\n"
            f"Task:* {task['task_description']}"
            f"*Field Design:* {task['field_design']}\n"
            f"*Difficulty:*{task['difficulty']}\n"
            f"Due:* {task['due_date']}\n"
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
            f"*Created by:* <@{task['user_id']}>\n"
            f"*Field Sheeting Time:* {field_time}\n"
            f"*Status:* In Progress"
        )
    else:
        card_text = f"*Task T-{task_id} - In Progress*\n*Status:* In Progress"
        
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
                "elements": [
                    {
                        "type": "button",
                        "text": {"type": "plain_text", "text": "Stop"},
                        "style": "danger",
                        "action_id": "stop_task",
                        "value": str(task_id)
                    },
                    {
                        "type": "button",
                        "text": {"type": "plain)text", "text": "Stop"},
                        "style": "danger",
                        "action_id": "stop_task",
                        "value": str(task_id)    
                    },
                    {
                        "type": "button",
                        "text": {"type": "plain_text", "text": "Complete Phase"},
                        "action_id": "complete_task",
                        "value": str(task_id)                        
                    }
                ]
            }
        ]
    )

#Stop Task Button
@app.action("stop_task")
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
            f"*Due:* {task['due_date']}\n"
            f"*Created by:* <@{task['user_id']}>\n"
            f"*Status:* Paused\n"
            f"*Field Time So Far:* {elapsed}"            
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

    client.chat_update(
        channel=body["channel"]["id"],
        ts=task["message_ts"],
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
                "elements": [
                    {
                        "type": "button",
                        "text": {"type": "plain_text", "text": "Resume"},
                        "style": "primary",
                        "action_id": "start_task",
                        "value": str(task_id)
                    },
                    {
                        "type":"button",
                        "text":{"type": "plain_text", "text": "Edit"},
                        "action_id": "edit_task",
                        "value": str(task_id)
                    },
                    {
                        "type": "button",
                        "text": {"type": "plain_text", "text": "Complete Phase"},
                        "action_id": "complete_task",
                        "value": str(task_id)
                    }
                ]
            }
        ]
    )

# Complete Task Button
@app.action("complete_task")
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
            channel=body["channel"]["id"],
            user=user_id,
            text="You can only control your own tasks."
        )
        return

    database.complete_task(task_id)
    updated_task = database.get_task(task_id)
    phase = updated_task["current_phase"]
    
    if phase == "field_sheeting":
        field_time = database.format_elapsed(updated_task["field_elapsed"])
        metadata = json.dumps({"task_id": task_id, "channel_id": channel_id})
        
        client.views_open(
            trigger_id=body["trigger_id"],
            view={
                "type": "modal",
                "callback_id": "border_modal",
                "title": {"type": "plain_text", "text": "Border Sheeting (Phase 2)"},
                "submit": {"type": "plain_text", "text": "Start Border Phase"},
                "close": {"type": "plain_text", "text": "Cancel"},
                "private_metadata": metadata,
                "blocks": [
                    {
                        "type": "section",
                        "text": f" Field Sheeting complete!* Time logged: *{field_time}*\nNow enter the Border Sheeting details."
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
                    }
                ]
            }
        )
    
    #To start packing phase automatically
    
    elif phase == "border_sheeting":
        database.move_to_packing_phase(task_id)
        updated_task = database.get_task(task_id)
        field_time = database.format_elapsed(updated_task["field_elapsed"])
        border_time = database.format_elapsed(updated_task["border_elapsed"])
        
        result = client.chat_postMessage(
            channel=channel_id,
            text=f"Task T-{task_id} has moved to Packing.",
            blocks=[
                {
                    "type": "section",
                    "text": {
                        "type": "mrkdwn",
                        "text": (
                            f"*Phase 3/4: Packing — In Progress* \n"
                            f"*ID:* T-{task_id}\n"
                            f"*Customer:* {updated_task['customer_name']}\n"
                            f"*Invoice:* {updated_task['invoice_number']}\n"
                            f"*Task:* {updated_task['task_description']}\n"
                            f"*Created by:* <@{updated_task['user_id']}>\n"
                            f"*Field Sheeting Time:* {field_time}\n"
                            f"*Border Sheeting Time:* {border_time}\n"
                            f"*Status:* Packing In Progress"
                        )
                    }
                },
                {
                    "type": "action",
                    "block_id": f"task_actions_{task_id}",
                    "elements": [
                        {
                            "type": "button",
                            "text": {"type": "plain_text", "text": "Stop"},
                            "style": "danger",
                            "action_id": "stop_task",
                            "value": str(task_id)
                        },
                        {
                            "type": "button",
                            "text": {"type": "plain_text", "text": "Complete Phase"},
                            "action_id": "complete_task",
                            "value": str(task_id)
                        }
                    ]
                }
            ]
        )
        
        database.update_message_ts(task_id, result["ts"])
    
    elif phase == "packing":
        packing_time = database.format_elapsed(updated_task["packing_elapsed"])
        metadata = json.dumps({"task_id":task_id, "channel_id": channel_id})
        
        client.views_open(
            trigger_id=body["trigger_id"],
            view={
                "type": "modal",
                "callback_id": "notes_modal",
                "title": {"type": "plain_text", "text": "Job Notes (Phase 4)"},
                "submit": {"type": "plain_text", "text": "Complete Job"},
                "close": {"type": "plain_text", "text": "Cancel"},
                "private_metadata": metadata,
                "blocks": [
                    {
                        "type": "section",
                        "text": {
                            "type": "mrkdwn",
                            "text": f" *Packing complete!* Time logged: *{packing_time}*\nAdd any final notes before closing this job."
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
        
@app.view("border_modal")
def handle_border_submission(ack,body, client):
    ack()
    user_id = body["user"]["id"]
    vals = body["view"]["state"]["values"]
    metadata = json.loads(body["view"]["private_metadata"])
    task_id = metadata["task_id"]
    channel_id = metadata["channel_id"]
    
# border details
    border_design = vals["border_design_block"]["border_design"]["value"]
    border_difficulty = vals["border_diff_block"]["border_difficulty"]["value"]
    
# Transitioning to border phase in the database
    database.move_to_border_phase(task_id, border_design, border_difficulty)
    task = database.get_task(task_id)
    field_time = database.format_elapsed(task["field_elapased"])
    
# posting card to the channel
    result = client.chat_postMessage(
        channel=channel_id,
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
                        "action_id": "start_task",
                        "value": str(task_id)
                    }
                ]
            }
        ]
    )
    
    database.update_message_ts(task_id, result["ts"])
    
@app.view("notes_modal")
def handle_notes_submission(ack,body,client):
    ack()
    user_id = body["user"]["id"]
    vals = body["view"]["state"]["values"]
    metadata = json.loads(body["view"]["private_metadata"])
    task_id = metadata["task_id"]
    channel_id = metadata["channel_id"]
    
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
    
    client.chat_postMessage(
        channel=channel_id,
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
    

        
#Delete Button
@app.action("delete_task")
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


@app.action("edit_task")
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

    # Open pre-filled edit modal
    client.views_open(
        trigger_id=body["trigger_id"],
        view={
            "type": "modal",
            "callback_id": "edit_task_modal",
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


@app.view("edit_task_modal")
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

    # Fetch updated task to refresh the card
    task = database.get_task(task_id)

    # Rebuild the card based on current status
    if task["status"] == "created":
        buttons = [
            {
                "type": "button",
                "text": {"type": "plain_text", "text": "Start"},
                "style": "primary",
                "action_id": "start_task",
                "value": str(task_id)
            },
            {
                "type": "button",
                "text": {"type": "plain_text", "text": "Edit"},
                "action_id": "edit_task",
                "value": str(task_id)
            },
            {
                "type": "button",
                "text": {"type": "plain_text", "text": "Delete"},
                "style": "danger",
                "action_id": "delete_task",
                "value": str(task_id)
            }
        ]
        status_text = "Created"
    else:
        buttons = [
            {
                "type": "button",
                "text": {"type": "plain_text", "text": "Resume"},
                "style": "primary",
                "action_id": "start_task",
                "value": str(task_id)
            },
            {
                "type": "button",
                "text": {"type": "plain_text", "text": "Edit"},
                "action_id": "edit_task",
                "value": str(task_id)
            },
            {
                "type": "button",
                "text": {"type": "plain_text", "text": "Complete"},
                "action_id": "complete_task",
                "value": str(task_id)
            }
        ]
        status_text = "🟠 Paused"

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
