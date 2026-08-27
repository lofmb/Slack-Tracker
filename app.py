"""
The Slack side of the workshop tracker: what a maker sees, and what happens
when they press it.

This file owns the conversation with the maker. It builds the working card and
the forms, decides which buttons a job should be offering, reads what comes
back, and says no in words when a job cannot do what was asked. It holds no
data of its own: every durable operation goes to database.py, which is the
other half of the tracker and talks to LMSA.

A handler here reads the same way throughout:

    the maker presses something
        -> work out which job, and whether it is theirs to touch
            -> check what the job actually allows
                -> ask database.py to record it
                    -> rebuild the card so it shows the new truth

The job runs Field -> Border -> Packing. Field and Border each carry setup and
sheeting underneath them, cutting is measured inside sheeting, and exactly one
piece of work accrues at a time. Those rules belong to the job, not to this
file; the sections below say where each part of the conversation lives.
"""

import os
import json
import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
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

# ---------------------------------------------------------------------------
# Forms and wording shared by more than one route
# ---------------------------------------------------------------------------
# Defined once here because two different paths reach each of them, and a
# change to one must not miss the other.

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
        "title": {"type": "plain_text", "text": "Packing"},
        "submit": {"type": "plain_text", "text": "Go to the packing"},
        "close": {"type": "plain_text", "text": "Cancel"},
        "private_metadata": metadata,
        "blocks": [
            {
                "type": "section",
                "text": {"type": "mrkdwn", "text": summary_text}
            }
        ]
    }


def notes_modal_view(metadata, summary_text):
    """
    The last form: what happened, before the job closes.

    Built here rather than inline so it reads beside the packing form it
    follows, and so the two forms that end a job are changed in one place.
    """
    return {
        "type": "modal",
        "callback_id": "trk_notes_modal",
        "title": {"type": "plain_text", "text": "Finish the job"},
        "submit": {"type": "plain_text", "text": "Finish the job"},
        "close": {"type": "plain_text", "text": "Cancel"},
        "private_metadata": metadata,
        "blocks": [
            {
                "type": "section",
                "text": {"type": "mrkdwn", "text": summary_text}
            },
            {
                "type": "input",
                "block_id": "notes_block",
                "optional": True,
                "label": {"type": "plain_text", "text": "Notes"},
                "element": {
                    "type": "plain_text_input",
                    "multiline": True,
                    "action_id": "general_notes",
                    "placeholder": {"type": "plain_text",
                                    "text": "Anything the next person should know"}
                }
            },
            {
                "type": "input",
                "block_id": "issues_block",
                "optional": True,
                "label": {"type": "plain_text", "text": "Anything go wrong?"},
                "element": {
                    "type": "plain_text_input",
                    "multiline": True,
                    "action_id": "issues",
                    "placeholder": {"type": "plain_text",
                                    "text": "Breakages, wrong material, missing pieces"}
                }
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
    return database.format_duration(task["border_elapsed"] or 0)


# ---------------------------------------------------------------------------
# The card the maker works from
# ---------------------------------------------------------------------------
#
# ONE card, built from what the job actually is, so every route through the
# workflow shows the same thing and a change to it cannot land on one path and
# miss another. It answers four questions, in the order a maker asks them:
#
#     Which job is this?          the header
#     What am I working on?       the status line
#     How long have I recorded?   the times
#     What can I do next?         the buttons
#
# Everything else the job knows is grouped underneath, because giving every
# field the same weight is what turns a card into a wall of text.

# What each piece of work is called wherever the maker sees it. Setup and
# sheeting share a lane, so a name has to say both which lane and which work.
WORK_NAMES = {
    ("field_sheeting", "setup"): "Field setup",
    ("field_sheeting", "production"): "Field sheeting",
    ("border_sheeting", "setup"): "Border setup",
    ("border_sheeting", "production"): "Border sheeting",
    ("packing", "production"): "Packing",
}

# What each move to that work is worth saying about it, on the Switch work
# form where the maker is choosing between them.
WORK_BLURBS = {
    ("field_sheeting", "setup"): "Getting the field ready - material, drawings, jig",
    ("field_sheeting", "production"): "Sheeting the field",
    ("border_sheeting", "setup"): "Getting the border ready",
    ("border_sheeting", "production"): "Sheeting the border",
    ("packing", "production"): "Packing this job; the sheeting waits, unfinished",
}

# The lane itself, as opposed to a piece of work inside it. Used where a figure
# covers the whole lane - its setup and its sheeting together - so that number
# is never labelled with the name of one of its halves.
LANE_NAMES = {
    "field_sheeting": "Field",
    "border_sheeting": "Border",
    "packing": "Packing",
}

# The press that says a lane is genuinely done.
FINISH_LABELS = {
    "field_sheeting": "Field sheeting finished",
    "border_sheeting": "Border finished",
    "packing": "Packing finished",
}

# What happens next once it is. A lane can be finished while the form that
# follows it is still owed - a cancelled modal, or a "no border" taken back -
# and then the button is not a finish at all, it is the way back to that form.
NEXT_STEP_LABELS = {
    "field_sheeting": "Enter border details",
    "border_sheeting": "Go to packing",
    "packing": "Finish the job",
}


def work_name(phase, activity):
    return WORK_NAMES.get((phase, activity), "Work")


def work_value(task_id, phase=None, activity=None):
    """
    What a button carries.

    A bare number still means "this job", which is what Edit, Delete and the
    rest need and what every card posted by an earlier version of the tracker
    sends. A button that moves the maker onto a particular piece of work names
    it, so the handler never has to guess which one was meant.
    """
    if phase is None:
        return str(task_id)
    return f"{task_id}|{phase}|{activity}"


def read_work_value(raw):
    """Read a button's value back. Returns (task_id, phase, activity)."""
    parts = str(raw).split("|")
    task_id = int(parts[0])
    if len(parts) == 3:
        return task_id, parts[1], parts[2]
    return task_id, None, None


def lane_state(task, phase):
    return (task.get("phase_states") or {}).get(phase)


def lane_open(task, phase):
    """A lane still to be worked: not finished, and not declared absent."""
    return lane_state(task, phase) not in ("complete", "skipped")


def work_elapsed(task, phase, activity):
    """
    Seconds recorded against one piece of work.

    Packing has no setup, so asking for its setup is nought - not the packing
    total over again. Anything that adds a lane's two activities together
    depends on that.
    """
    if phase == "packing":
        return (task["packing_elapsed"] or 0) if activity == "production" else 0
    lane = "field" if phase == "field_sheeting" else "border"
    return task[f"{lane}_{activity}_elapsed"] or 0


def switch_destinations(task):
    """
    The work the maker may move to from here.

    Derived from the job rather than listed per card. A lane that is finished,
    or that the maker declared did not happen, is not somewhere to go; a border
    nobody has described yet has not been reached; and packing has no setup.
    Whatever survives is offered, and the maker's press is what chooses.
    """
    here = task.get("working_on") or {}
    out = []
    for phase, activity in (
        ("field_sheeting", "setup"),
        ("field_sheeting", "production"),
        ("border_sheeting", "setup"),
        ("border_sheeting", "production"),
        ("packing", "production"),
    ):
        if phase == here.get("phase") and activity == here.get("activity"):
            continue
        if not lane_open(task, phase):
            continue
        # The border modal is where a border is described. Until it has been,
        # there is no border to go and work on.
        if phase == "border_sheeting" and not task.get("border_design"):
            continue
        out.append({
            "phase": phase,
            "activity": activity,
            "label": work_name(phase, activity),
            "blurb": WORK_BLURBS[(phase, activity)],
        })
    return out


def resume_target(task):
    """
    What Resume means on a paused card: the last thing the maker was doing.

    Read from the ledger rather than assumed, because the last thing they were
    doing is not always the lane the job is on - a maker who stopped packing
    mid-field is paused on the packing. If that work has since been finished or
    declared absent, the lane the job is on is the honest fallback.
    """
    last = task.get("last_work")
    if last and lane_open(task, last["phase"]):
        return last["phase"], last["activity"]
    cursor = task["current_phase"]
    if cursor != "completed" and lane_open(task, cursor):
        return cursor, "production"
    return None


# Slack's ceiling for a header block's text.
HEADER_LIMIT = 150


def header_text(task, suffix=""):
    """
    "T-12  Customer Name", trimmed to something Slack will accept.

    The job number and any suffix are never what gets cut: they are how a maker
    finds the card. A name too long to fit is shortened here and shown in full
    in the fields below, so nothing is lost.
    """
    prefix = "T-" + str(task["task_id"]) + "  "
    room = HEADER_LIMIT - len(prefix) - len(suffix)
    name = task["customer_name"] or ""
    if len(name) > room:
        name = name[: max(room - 1, 0)].rstrip() + "…"
    return prefix + name + suffix


def customer_was_trimmed(task):
    return not header_text(task).endswith(task["customer_name"] or "")


def _button(text, action_id, value, style=None, confirm=None):
    button = {
        "type": "button",
        "text": {"type": "plain_text", "text": text},
        "action_id": action_id,
        "value": value,
    }
    if style:
        button["style"] = style
    if confirm:
        button["confirm"] = confirm
    return button


def _start_button(text, task_id, phase, activity, style=None):
    """
    A button that opens a piece of work and starts timing it.

    Slack requires an action_id to be UNIQUE within a message, and a card
    legitimately offers a lane's setup beside its sheeting - "Start border
    setup" next to "Start border sheeting", or "Resume field setup" next to
    "Start field sheeting". So the id says WHICH of the two is being asked
    for, rather than leaving two identical ids to be told apart by a value
    Slack never looks at. Both reach the same handler; the value still carries
    the job, the lane and the activity.
    """
    action_id = "trk_start_setup" if activity == "setup" else "trk_start_production"
    return _button(text, action_id, work_value(task_id, phase, activity), style=style)


def _jig_button(task):
    """
    Set jig / Add jig.

    Same action either way - the record is always appended, never overwritten,
    because a jig that was genuinely used stays used. The wording changes
    because "Add" reads as a second one, and the first time there is nothing to
    add to.
    """
    has_jig = bool(task.get("field_jigs") or task.get("border_jigs"))
    return _button(
        "Add jig" if has_jig else "Set jig / template",
        "trk_add_jig",
        work_value(task["task_id"]),
    )


def _finish_button(task):
    """
    The one press that says a lane is done - or, once it is, the way on to the
    form that follows it.

    Nothing else on the card finishes anything, and this is deliberately not
    offered until the lane has some sheeting time on it: straight out of setup
    the forward move is to START the sheeting, not to declare it over.
    """
    cursor = task["current_phase"]
    if cursor == "completed":
        return None
    if lane_state(task, cursor) == "complete":
        # Finished already; what is owed is the form that comes next.
        return _button(NEXT_STEP_LABELS[cursor], "trk_complete_task", work_value(task["task_id"]),
                       style="primary")
    if not lane_open(task, cursor):
        return None
    here = task.get("working_on") or {}
    on_it_now = here.get("phase") == cursor and here.get("activity") == "production"
    # Straight out of setup there is nothing to declare finished, so the
    # forward move is to START the sheeting. Once it has started - this
    # instant, not once a minute has accrued - finishing it is a real choice.
    if not on_it_now and work_elapsed(task, cursor, "production") <= 0:
        return None
    label = FINISH_LABELS[cursor]
    # Whichever move this card actually offers for leaving the lane unfinished.
    # At packing both other lanes are done, so there is nowhere to switch to
    # and Switch work is not on the card; Pause is the one that leaves it.
    leave_unfinished = "Switch work" if switch_destinations(task) else "Pause"
    return _button(
        label,
        "trk_complete_task",
        work_value(task["task_id"]),
        confirm={
            "title": {"type": "plain_text", "text": label + "?"},
            # Slack renders a confirmation's text as PLAIN TEXT. Asterisks
            # meant as emphasis are printed, so the maker read "*field
            # sheeting*" with the asterisks in it.
            "text": {
                "type": "plain_text",
                "text": (
                    "This closes the " + work_name(cursor, "production").lower() +
                    " for good and moves the job on.\n\n"
                    "Still something to do on it? Use " + leave_unfinished + " instead - that "
                    "leaves it unfinished and you can come back."
                ),
            },
            "confirm": {"type": "plain_text", "text": "Yes, it is finished"},
            "deny": {"type": "plain_text", "text": "Not yet"},
        },
    )


def lane_finish_is_new(task):
    """
    Whether this press is actually finishing a lane, or only reopening a form.

    The finish button carries two jobs (see NEXT_STEP_LABELS): it closes a
    lane, and once the lane is closed it becomes the way back to the form that
    lane still owes - after a cancelled modal, or a "no border" taken back.
    Only the first is something that happened; the second is a maker returning
    to paperwork.

    The team channel hears about the first only. Announcing both put "Field
    sheeting finished" in the channel twice, two minutes apart, with identical
    figures, for a lane that was finished once - which reads as the field
    having been done again.

    Same test the button itself uses to decide its label, so what the card
    calls the press and whether the channel hears about it cannot drift apart.
    """
    return lane_state(task, task["current_phase"]) != "complete"


def delete_still_applies(task):
    """
    Whether this job could still be one that should never have been entered.

    Delete is that correction, and it is only honest while nothing has been
    made yet. "Setup" alone does not say so: the BORDER's setup is reached with
    the field worked and finished behind it, and offering to take the job off
    the list there put a red button beside several hours of real labour.

    So the test is the state of the job, not the name of the activity - the job
    is still on its first lane, and no lane has produced anything.
    """
    if task["current_phase"] != "field_sheeting":
        return False
    for phase in ("field_sheeting", "border_sheeting", "packing"):
        if work_elapsed(task, phase, "production"):
            return False
    return True


def _delete_button(task_id):
    return _button(
        "Delete job",
        "trk_delete_task",
        work_value(task_id),
        style="danger",
        confirm={
            "title": {"type": "plain_text", "text": "Delete this job?"},
            "text": {
                "type": "plain_text",
                "text": (
                    "It comes off your list and you can start another. Any time already "
                    "recorded stays on the record, but you cannot pick this job back up "
                    "from here - a supervisor would have to."
                ),
            },
            "confirm": {"type": "plain_text", "text": "Yes, take it off my list"},
            "deny": {"type": "plain_text", "text": "Keep it"},
        },
    )


def card_actions(task):
    """The buttons, chosen by what the maker is actually able to do next."""
    task_id = task["task_id"]
    here = task.get("working_on")
    cutting = task.get("cutting_now")
    cursor = task["current_phase"]
    buttons = []

    if here and here["activity"] == "setup":
        # Setting up. The forward move is the sheeting itself, and this is the
        # moment the jig becomes known, so both are on the card.
        buttons.append(_start_button(
            "Start " + work_name(here["phase"], "production").lower(),
            task_id, here["phase"], "production", style="primary",
        ))
        buttons.append(_button("Pause setup", "trk_stop_task", work_value(task_id)))
        buttons.append(_jig_button(task))
        buttons.append(_button("Edit details", "trk_edit_task", work_value(task_id)))
        if delete_still_applies(task):
            buttons.append(_delete_button(task_id))
        return buttons

    if here:
        # Working. Pause, measure cutting alongside it, move to other work, or
        # say the lane is finished.
        if cutting:
            buttons.append(_button("Stop cutting", "trk_stop_cutting", work_value(task_id),
                                   style="primary"))
        buttons.append(_button("Pause", "trk_stop_task", work_value(task_id)))
        if not cutting and here["phase"] != "packing" and here["activity"] == "production":
            buttons.append(_button("Start cutting", "trk_start_cutting", work_value(task_id)))
        # Packing worked as an interruption: the way back to the waiting
        # sheeting is one press, because that is the common one.
        if here["phase"] != cursor and cursor != "completed" and lane_open(task, cursor):
            buttons.append(_start_button(
                "Back to " + work_name(cursor, "production").lower(),
                task_id, cursor, "production", style="primary",
            ))
        if switch_destinations(task):
            buttons.append(_button("Switch work", "trk_switch_work", work_value(task_id)))
        buttons.append(_jig_button(task))
        if here["phase"] == cursor:
            finish = _finish_button(task)
            if finish:
                buttons.append(finish)
        return buttons

    # Nothing is being timed.
    resume = resume_target(task)
    if resume:
        untouched = (
            not work_elapsed(task, resume[0], "setup")
            and not work_elapsed(task, resume[0], "production")
        )
        if untouched and resume[0] != "packing":
            # A sheeting lane nobody has started yet. Setup comes first, the
            # same way it does on the field, and the sheeting is right beside
            # it for a job that needs no preparing.
            buttons.append(_start_button(
                "Start " + work_name(resume[0], "setup").lower(),
                task_id, resume[0], "setup", style="primary",
            ))
            buttons.append(_start_button(
                "Start " + work_name(resume[0], "production").lower(),
                task_id, resume[0], "production",
            ))
        else:
            # "Resume" only where there is something to resume; a lane with no
            # time on it is being started, whatever the ledger last recorded.
            verb = "Resume" if work_elapsed(task, resume[0], resume[1]) else "Start"
            buttons.append(_start_button(
                verb + " " + work_name(*resume).lower(),
                task_id, resume[0], resume[1], style="primary",
            ))
            if resume[1] == "setup":
                buttons.append(_start_button(
                    "Start " + work_name(resume[0], "production").lower(),
                    task_id, resume[0], "production",
                ))
    if switch_destinations(task):
        buttons.append(_button("Switch work", "trk_switch_work", work_value(task_id)))
    buttons.append(_jig_button(task))
    buttons.append(_button("Edit details", "trk_edit_task", work_value(task_id)))
    # A job entered by mistake is as likely to be spotted on the paused card as
    # on the running one - the maker stops, looks, and sees it is the wrong job.
    if delete_still_applies(task):
        buttons.append(_delete_button(task_id))
    finish = _finish_button(task)
    if finish:
        buttons.append(finish)
    # A job marked "no border" can still turn out to need one. The way back
    # stays open until packing is finished for good, and lives on the paused
    # card because the correction is refused while a timer runs - a button that
    # always refuses is worse than no button.
    if cursor == "packing" and task.get("border_skipped") and not task.get("packing_finished"):
        buttons.append(_button("Border after all", "trk_undo_no_border", work_value(task_id)))
    return buttons


def _status_lines(task):
    """
    What the maker is doing, said plainly and first.

    A working card names the work and how long it has been going. A paused card
    says nothing is being timed, and what it was. During a packing interruption
    it also says what is waiting, because "the field is paused, not finished"
    is the fact a maker most needs and the one a status word cannot carry.
    """
    here = task.get("working_on")
    cursor = task["current_phase"]
    if not here:
        last = task.get("last_work")
        lines = ["*Paused - nothing is being timed*"]
        if last:
            lines.append(
                "Last on " + work_name(last["phase"], last["activity"]).lower() + ", "
                + database.format_duration(work_elapsed(task, last["phase"], last["activity"]))
                + " recorded."
            )
        return lines

    name = work_name(here["phase"], here["activity"])
    lines = ["*Working on: " + name + "*"]
    if here["activity"] == "setup":
        lines.append("_Getting the job ready to sheet - material, drawings, and the jig._")
    if here["phase"] != cursor and cursor != "completed":
        waiting = work_name(cursor, "production").lower()
        lines.append(
            "_" + waiting.capitalize() + " is paused while you do this. Its time is safe, "
            "nothing about it is finished, and the job is still on it._"
        )
    lines.append(
        name + " so far: *"
        + database.format_duration(work_elapsed(task, here["phase"], here["activity"])) + "*"
    )
    cutting = task.get("cutting_now")
    if cutting:
        inside = work_name(cutting["parent_phase"], "production").lower()
        lines.append(
            "*Cutting now* - the " + inside + " timer is still running, so this time counts "
            "as " + inside + " either way."
        )
    return lines


def _fields(task):
    """
    The job itself. Two columns, so it reads as a small label rather than
    another paragraph, and only the things that have an answer.
    """
    pairs = []
    # Only when the header could not hold it - otherwise it would be on the
    # card twice.
    if customer_was_trimmed(task):
        pairs.append(("Customer", task["customer_name"]))
    pairs += [("Job", task["task_description"]), ("Invoice", task["invoice_number"])]
    if task.get("field_design"):
        pairs.append(("Field design", task["field_design"]))
    if task.get("difficulty"):
        pairs.append(("Field difficulty", task["difficulty"]))
    if task.get("border_design"):
        pairs.append(("Border design", task["border_design"]))
    if task.get("border_difficulty"):
        pairs.append(("Border difficulty", task["border_difficulty"]))
    if task.get("field_jigs"):
        pairs.append(("Field jig", task["field_jigs"]))
    if task.get("border_jigs"):
        pairs.append(("Border jig", task["border_jigs"]))
    pairs.append(("Due", due_date_display(task)))
    # Slack renders at most ten fields in a section.
    return [{"type": "mrkdwn", "text": "*" + label + "*\n" + str(value)}
            for label, value in pairs[:10]]


# ---------------------------------------------------------------------------
# The due date
# ---------------------------------------------------------------------------
# One field, meaning DD/MM/YY. Blank means nobody supplied a date - not that
# the job has no deadline, which is not a state this workshop has.

NO_DUE_DATE = "Not set"


DUE_DATE_LABEL = "Due date (DD/MM/YY)"
DUE_DATE_HINT = "e.g. 01/09/26 - or leave blank"
DUE_DATE_ERROR = ("Enter the date as DD/MM/YY, for example 01/09/26. "
                  "If nobody has given you one, leave it blank.")


def read_due_date(typed):
    """
    What the maker meant by what they typed in the due date box.

    Returns (text to store, error to show them). The text is None when no date
    has been supplied: a blank box, or the "N/A" the old form used to write,
    which is read as the same thing so retyping it cannot mint another one.

    The box says DD/MM/YY and means it. A label that accepts "Friday" is a
    suggestion rather than a promise, and stores a due date nothing can sort by
    or chase. A real date is stored the way the label reads, whatever separator
    was typed and whether the year was given as two digits or four, so every
    card says it the same way.

    The date has to exist: 31/02/26 is refused here rather than accepted and
    then rejected by the database, where the maker would see nothing useful.

    This governs TYPING ONLY. Rows already holding free text keep it, and are
    read back untouched.
    """
    raw = (typed or "").strip()
    if not raw or raw.upper() == "N/A":
        return None, None
    for separator in ("/", "-", "."):
        parts = raw.split(separator)
        if len(parts) != 3 or not all(p.isdigit() for p in parts):
            continue
        day, month, year = parts
        if len(year) == 2:
            year = "20" + year
        if len(year) != 4:
            break
        try:
            datetime.date(int(year), int(month), int(day))
        except ValueError:
            break
        return "%02d/%02d/%s" % (int(day), int(month), year[2:]), None
    return None, DUE_DATE_ERROR


def due_date_supplied(task):
    """
    The due date the maker was given, or None when nobody has given them one.

    Every job needs doing as soon as practicable, so there is no such thing as
    a job with no deadline. There are only jobs where a specific calendar date
    is known and jobs where it is not, and a blank box says the second one.

    Rows entered through the old form stored the word "N/A" for that, either as
    the typed text or as a ticked box. It is read as "none supplied" here, at
    the moment it is shown, so an old record reads correctly without anything
    being rewritten. Anything else the maker typed is theirs and comes back
    untouched - "Friday" and "when the paint arrives" are real answers.
    """
    text = (task.get("due_date") or "").strip()
    if not text or text.upper() == "N/A":
        return None
    return text


def due_date_display(task):
    """The due date as a person reads it, wherever one is shown."""
    return due_date_supplied(task) or NO_DUE_DATE


# ---------------------------------------------------------------------------
# Difficulty
# ---------------------------------------------------------------------------
# A number from 1 to 10, and nothing else. That is what it has always been -
# every difficulty ever recorded is a whole number in that range - but the box
# never said so, so a maker who answered it in words was told only that they
# had used too many characters. The label carries the range now, and a value
# that is not one is refused in those words.

DIFFICULTY_LABEL_FIELD = "Field difficulty (1-10)"
DIFFICULTY_LABEL_BORDER = "Border difficulty (1-10)"
DIFFICULTY_HINT = "e.g. 3"
DIFFICULTY_ERROR = "Give the difficulty as a whole number from 1 to 10, for example 3."


def read_difficulty(typed):
    """
    Returns (text to store, error to show them). None with no error means the
    box was left blank, which is allowed: not every job has both lanes, and a
    lane the diagram does not have has no difficulty either.
    """
    raw = (typed or "").strip()
    if not raw:
        return None, None
    if not all(character in "0123456789" for character in raw):
        return None, DIFFICULTY_ERROR
    value = int(raw)
    if value < 1 or value > 10:
        return None, DIFFICULTY_ERROR
    return str(value), None


def _lane_lines(task, phase):
    """
    One lane's time, with its parts underneath in the shape they really have.

    Setup and sheeting ADD UP to the lane's total. Cutting does not: it is time
    spent inside the sheeting, already counted in it. Listing both after one
    word - "includes 3m 42s setup, 50s cutting" - put two different
    relationships in one sum, and the arithmetic then did not work: a maker
    adding those two figures came up short of the lane total and had no way to
    see why. So the parts sit under the lane, and the cutting sits under the
    sheeting it happened in.

    Packing has no setup, so its total IS its one figure and repeating it as a
    part underneath would say nothing.
    """
    here = task.get("working_on") or {}
    on_this_lane = here.get("phase") == phase
    setup = work_elapsed(task, phase, "setup")
    production = work_elapsed(task, phase, "production")
    if not setup and not production and not on_this_lane:
        return []

    label = LANE_NAMES[phase]
    lines = ["*" + label + "*  " + database.format_duration(setup + production)]
    if phase == "packing":
        return lines

    if setup or (on_this_lane and here.get("activity") == "setup"):
        lines.append("•  Setup  " + database.format_duration(setup))
    if production or (on_this_lane and here.get("activity") == "production"):
        lines.append("•  Sheeting  " + database.format_duration(production))
        lane = "field" if phase == "field_sheeting" else "border"
        cutting = task.get(lane + "_cutting_elapsed") or 0
        if cutting:
            # Cutting is written under the sheeting it happened inside, never
            # as a line of its own, because it is time already counted. An en
            # dash rather than a hollow bullet: both read as a level below, and
            # this one still prints where the card text is rendered outside
            # Slack.
            lines.append("     –  of which cutting  " + database.format_duration(cutting))
    return lines


def _time_lines(task, total_label="Total job time"):
    """
    What has been recorded, lane by lane.

    A lane appears once there is something to say about it, so an early card is
    short and a late one is complete.
    """
    here = task.get("working_on") or {}
    rows = _lane_lines(task, "field_sheeting")
    if task.get("border_skipped"):
        rows.append("*Border*  no border on this job")
    else:
        rows += _lane_lines(task, "border_sheeting")
    rows += _lane_lines(task, "packing")

    # Nothing worked yet: the status line has already said what the maker is on
    # and how long for, and repeating it under a heading is three noughts and no
    # information.
    if not rows or not task["total_elapsed"]:
        return []
    return ["*Time recorded*"] + rows + [
        "*" + total_label + "*  " + database.format_duration(task["total_elapsed"])
    ]


def lane_report(task, phase):
    """
    One line saying how long a lane took and how that time was made up.

    Used wherever a lane is announced as finished - the team channel, the form
    that follows it, the closing summary - so those three can never disagree
    about the same lane, and so "5h 12m" is never quietly a different 5h 12m in
    two places.
    """
    if phase == "border_sheeting" and task.get("border_skipped"):
        return "*Border*  no border on this job"
    if phase == "field_sheeting" and task.get("field_skipped"):
        return "*Field*  no field on this job"
    return "\n".join(_lane_lines(task, phase))


def job_summary_blocks(task, finished_by, general_notes, issues):
    """
    The closing summary posted to the team channel.

    The same time lines the card showed all along, so a maker reading it
    recognises the job they just worked. Cutting appears inside the lane it
    happened in and never as a line of its own: it is time already counted, and
    a separate entry would read as time on top of the job.
    """
    facts = [f"Invoice {task['invoice_number']}", task["task_description"]]
    if task.get("field_jigs"):
        facts.append("Field jig " + task["field_jigs"])
    if task.get("border_jigs"):
        facts.append("Border jig " + task["border_jigs"])

    return [
        {
            "type": "header",
            "text": {"type": "plain_text", "text": header_text(task)},
        },
        {
            "type": "section",
            "text": {"type": "mrkdwn",
                     "text": "*Finished by <@" + finished_by + ">*\n" + "  ·  ".join(facts)},
        },
        {
            "type": "section",
            "text": {"type": "mrkdwn", "text": "\n".join(_time_lines(task))},
        },
        {
            "type": "section",
            "text": {"type": "mrkdwn",
                     "text": "*Notes*\n" + general_notes + "\n\n*Anything go wrong?*\n" + issues},
        },
    ]

def border_route(task):
    """
    How THIS card reaches the border, in the words of a button that is on it.

    Straight from a finished field the card offers the border's own starts.
    After Border after all it does not: the maker is paused on packing, so
    Resume packing is the forward move and the border is reached through Switch
    work. Saying "start the border" there names a button that is not there.
    """
    for button in card_actions(task):
        if button.get("action_id") not in ("trk_start_setup", "trk_start_production"):
            continue
        if "border_sheeting" in (button.get("value") or ""):
            return "Start the border when you are ready."
    return "Use Switch work and choose the border when you are ready."


def job_card(task, note=None):
    """
    The whole card. Returns (fallback text, blocks).

    `note` is a single line put at the top when something just happened that
    the card alone would not explain - a jig recorded, a border put back.
    """
    task_id = task["task_id"]
    blocks = [{
        "type": "header",
        "text": {"type": "plain_text", "text": header_text(task)},
    }]
    if note:
        blocks.append({"type": "section", "text": {"type": "mrkdwn", "text": note}})
    blocks.append({
        "type": "section",
        "text": {"type": "mrkdwn", "text": "\n".join(_status_lines(task))},
    })
    # The buttons come BEFORE the reading matter. On a phone the job's
    # details and its times run to about thirty lines, and a maker wanting
    # to pause or start the cutting had to scroll past all of it - while
    # standing at a bench, usually mid-task. Which job this is and what
    # they are on stay above, because those say whether it is even the
    # right card; the times and the rest of the job sit below, where
    # reading is what a maker is doing anyway.
    actions = card_actions(task)
    if actions:
        blocks.append({
            "type": "actions",
            "block_id": "task_actions_" + str(task_id),
            "elements": actions,
        })
    times = _time_lines(task)
    if times:
        blocks.append({"type": "section", "text": {"type": "mrkdwn", "text": "\n".join(times)}})
    blocks.append({"type": "section", "fields": _fields(task)})
    footer = "Logged by <@" + task["user_id"] + ">"
    on_setup = task.get("working_on") or {}
    if on_setup.get("activity") == "setup":
        # A job starts timing its setup the moment it exists, and nothing
        # closes a timer left running overnight. Pause is already on the card;
        # this says, where the risk actually is, what it is for.
        footer += "  ·  Pause setup if you stop working on it - you can pick it up again any time."
    if any(b.get("action_id") == "trk_switch_work" for b in actions):
        footer += "  ·  Switching work or pausing never finishes anything."
    blocks.append({"type": "context", "elements": [{"type": "mrkdwn", "text": footer}]})

    here = task.get("working_on")
    summary = ("T-" + str(task_id) + ": " + work_name(here["phase"], here["activity"])
               if here else "T-" + str(task_id) + ": paused")
    return summary, blocks


def update_card(client, task, channel_id, note=None):
    """
    Rewrite the job's card where it already is - or post one, if the job has
    somehow ended up without a card to rewrite.
    """
    if not task.get("message_ts"):
        repost_card(client, task, channel_id, note=note)
        return
    text, blocks = job_card(task, note=note)
    client.chat_update(
        channel=channel_id,
        ts=task["message_ts"],
        text=text,
        blocks=blocks,
    )


def repost_card(client, task, channel_id, note=None):
    """
    Replace the card with a fresh one at the bottom of the DM.

    Used at the points where the job genuinely moves on - a lane finished, the
    border decided - because by then the old card has usually scrolled away
    behind the modal that was just filled in, and a maker should not have to go
    looking for the job they are working on.
    """
    text, blocks = job_card(task, note=note)
    if task.get("message_ts"):
        try:
            client.chat_delete(channel=channel_id, ts=task["message_ts"])
        except SlackApiError:
            # Already gone is the outcome this wanted anyway.
            pass
    result = client.chat_postMessage(channel=channel_id, text=text, blocks=blocks)
    database.update_message_ts(task["task_id"], result["channel"], result["ts"])
    return result


def resolve_job(client, body, task_id, user_id, channel_id):
    """
    The three things every button checks before it does anything: the job is
    still there, it belongs to this maker, and it is not already finished.

    Returns the job, or None having already said why.
    """
    task = database.get_task(task_id)
    if task is None:
        client.chat_postEphemeral(
            channel=channel_id,
            user=user_id,
            text="That job is not there any more - it may have been deleted.",
        )
        return None
    if task["user_id"] != user_id:
        client.chat_postEphemeral(
            channel=channel_id,
            user=user_id,
            text="This is <@" + task["user_id"] + ">'s job, so only they can change it.",
        )
        return None
    if task["current_phase"] == "completed":
        client.chat_postEphemeral(
            channel=channel_id,
            user=user_id,
            text="T-" + str(task_id) + " is finished, so nothing more can be recorded on it.",
        )
        return None
    return task


def refusal_text(reason, task, phase=None):
    """
    A refusal, in workshop words: what happened, and what to do about it.

    Every one of these is something a real maker can cause by pressing a
    button, usually from a card that has gone stale in another window. So the
    answer says what the job is actually doing now and what to press instead -
    never a reason code, and never nothing at all.
    """
    lane = work_name(phase or task["current_phase"], "production").lower()
    here = task.get("working_on")
    doing = work_name(here["phase"], here["activity"]).lower() if here else None
    texts = {
        "already_running": "You are already on that, so nothing has changed.",
        "another_phase_running": (
            "The " + (doing or "other") + " timer is running on this job. Pause it first, "
            "then try again."
        ),
        "other_activity_running": (
            "You are on " + (doing or "other work") + " on this job. Use *Switch work* to "
            "move across - it pauses what you are doing rather than finishing it."
        ),
        "phase_already_complete": (
            "The " + lane + " is already finished on this job, so it cannot be started again."
        ),
        "phase_already_skipped": (
            "This job is marked as having no border, so there is no border to work on. If that "
            "was wrong, use *Border after all* on the packing card."
        ),
        "phase_has_no_setup": "Packing has no setup step - there is nothing to prepare.",
        "cutting_needs_production_work": (
            "Start the sheeting first. Cutting is recorded as part of the field or border work "
            "it happens during, so there has to be some running."
        ),
        "not_running": (
            "Nothing is being timed on this job at the moment - your card will show "
            "what this job is up to and what you can start."
        ),
        "already_cutting": "The cutting is already being timed.",
        "not_cutting": (
            "No cutting is being timed on this job at the moment - your card will show "
            "what this job is up to and what you can start."
        ),
        "job_not_open": "This job is no longer open, so nothing has been changed.",
    }
    return texts.get(reason, "That could not be done, so nothing has been changed.")


# ---------------------------------------------------------------------------
# Commands, and the handlers behind every button
# ---------------------------------------------------------------------------
# From here down, every function is registered with Slack. The middleware
# runs first on each delivery and notes which delivery it is, so a click that
# arrives twice is recognised as the same action.

@app.middleware
def track_slack_delivery(body, next):
    # Note which Slack delivery is being handled, so a redelivery of the same
    # click is recognised as the same action rather than applied twice. Scoped
    # to this handler and cleared afterwards, so nothing carries over into the
    # next request. Changes no behaviour a maker can see.
    with database.slack_request(body):
        return next()

# A liveness check: it answers, so the bot is connected and its commands are
# registered. Nothing in the workflow uses it.

@app.command ("/hello")
def hello_command(ack, body, say):

    ack()
    user_id = body["user_id"]
    say(f"Hi there, <@{user_id}>! I'm ready to track your projects.")

# /track opens the intake form. /track export sends the spreadsheet instead.

@app.command("/track")
def track_command(ack, body, client):
    ack()
    user_id = body ["user_id"]

    subcommand = body.get ("text", "").strip().lower()
    if subcommand == "export":
        handle_export(body,client)
        return

    # One job at a time per person. A maker with something already open is told
    # what it is, rather than quietly given a second job to lose track of.
    active_task = database.get_active_task(user_id)
    if active_task:
        client.chat_postEphemeral(
            channel=body["channel_id"],
            user=user_id,
            text=f"You already have an active job: {active_task['task_description']}. Finish or pause it before starting another."
            
        )
        return
    # The first of the two intake forms. private_metadata carries the channel
    # the command came from, so the job can be announced back to the same room.
    client.views_open(
        trigger_id=body["trigger_id"],
        view={
            "type": "modal",
            "callback_id": "trk_track_step_1",  # ID used to catch the submission
            "title": {"type": "plain_text", "text": "New job - 1 of 2"},
            "private_metadata":body["channel_id"],
            "submit": {"type": "plain_text", "text": "Next"},
            "close": {"type": "plain_text", "text": "Cancel"},
            "blocks": [
                {
                    "type": "input",
                    "block_id": "customer_block",
                    "label": {"type": "plain_text", "text": "Customer name"},
                    "element": {
                        "type": "plain_text_input",
                        "action_id": "customer_name"
                    }
                },
                {
                    "type": "input",
                    "block_id": "invoice_block",
                    "label": {"type": "plain_text", "text": "Invoice / Pro Forma number"},
                    "element": {
                        "type": "plain_text_input",
                        "action_id": "invoice_num"
                    }
                },
                {
                    "type": "input",
                    "block_id": "task_block",
                    "label": {"type": "plain_text", "text": "Job description"},
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
                    "label": {"type": "plain_text", "text": DUE_DATE_LABEL},
                    "element": {
                        "type": "plain_text_input",
                        "action_id": "due_date",
                        "placeholder": {"type": "plain_text", "text": DUE_DATE_HINT}
                    }
                }
            ]
        }
    )

# ---------------------------------------------------------------------------
# The spreadsheet export
# ---------------------------------------------------------------------------

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
        # The lane total, then how it was made up. Setup and sheeting ADD UP to
        # the lane total; cutting is time already inside the sheeting, so it is
        # a breakdown of it and must never be added on.
        "Field Time",
        "Field Setup Time",
        "Field Sheeting Time",
        "Field Cutting (within sheeting)",
        "Border Design",
        "Border Difficulty",
        "Border Jig Size(s)",
        "Border Time",
        "Border Setup Time",
        "Border Sheeting Time",
        "Border Cutting (within sheeting)",
        "Packing Time",
        "Total Time",
        "General Notes",
        "Issues Encountered",
        "Completed By",
        "Date Created",
    ]
    ws.append(headers)
    
    # Styling the headers

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
            due_date_display(task),
            task["field_design"] or "-",
            task["difficulty"],
            # Several jigs show as one readable cell, e.g. "49.6 / 50"
            task["field_jigs"] or "-",
            database.format_elapsed(field_elapsed),
            database.format_elapsed(task["field_setup_elapsed"] or 0),
            database.format_elapsed(task["field_production_elapsed"] or 0),
            database.format_elapsed(task["field_cutting_elapsed"] or 0),
            "No Border" if task.get("border_skipped") else (task["border_design"] or "-"),
            "-" if task.get("border_skipped") else (task["border_difficulty"] or "-"),
            "-" if task.get("border_skipped") else (task["border_jigs"] or "-"),
            border_time_display(task),
            "-" if task.get("border_skipped") else database.format_elapsed(task["border_setup_elapsed"] or 0),
            "-" if task.get("border_skipped") else database.format_elapsed(task["border_production_elapsed"] or 0),
            "-" if task.get("border_skipped") else database.format_elapsed(task["border_cutting_elapsed"] or 0),
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

# ---------------------------------------------------------------------------
# Job intake
# ---------------------------------------------------------------------------
# Two forms, then the job exists. Submitting the second one is the handover
# into workshop work: it creates the job and starts its setup in the same
# transaction, which is why the card that follows has no Start button.

@app.view("trk_track_step_1")
def handle_step_1(ack,body,client,):
    vals = body["view"]["state"]["values"]
    channel_id = body["view"]["private_metadata"]

    customer_name = vals["customer_block"]["customer_name"]["value"]
    invoice_number = vals["invoice_block"]["invoice_num"]["value"]
    task_description = vals["task_block"]["task_desc"]["value"]
    # An empty box is carried as nothing at all, not as the word "N/A". Nobody
    # has given this maker a date yet; that is not a job with no deadline.
    due_date, due_date_error = read_due_date(vals["date_block"]["due_date"]["value"])
    if due_date_error:
        # Sent back to the box it belongs to, so the maker reads the message
        # under the date rather than losing the whole form.
        ack(response_action="errors", errors={"date_block": due_date_error})
        return

    # private_metadata is the only way to carry these across to the pushed
    # form, which arrives as a separate submission.
    step1_data = {
        "channel_id": channel_id,
        "customer_name": customer_name,
        "invoice_number": invoice_number,
        "task_description": task_description,
        "due_date": due_date,
    }

    # The second form asks what the diagram has on it. A job may be a field, a
    # border, or both, so every box here is optional and the submission checks
    # that at least one of them was answered. Asking for a field on a
    # border-only job is how a lane nobody drew ends up on the record.
    #
    # No jig question here. A maker filling this in has just been handed the job
    # and normally does not know the jig yet - finding and testing it is the
    # setup. The card asks at the point it can actually be answered.
    ack(response_action="push", view={
        "type": "modal",
        "callback_id": "trk_track_step_2",
        "title": {"type": "plain_text", "text": "New job - 2 of 2"},
        "submit": {"type": "plain_text", "text": "Create the job"},
        "close": {"type": "plain_text", "text": "Back"},
        "private_metadata": json.dumps(step1_data),
        "blocks": [
            {
                "type": "section",
                "text": {"type": "mrkdwn",
                         "text": "*What is on the diagram?*\nFill in the field, the border, or "
                                 "both. Leave a part blank if the diagram does not have it."}
            },
            {
                "type": "input",
                "block_id": "design",
                "optional": True,
                "label": {"type": "plain_text", "text": "Field design"},
                "element": {
                    "type": "plain_text_input",
                    "action_id": "val",
                    "placeholder": {"type": "plain_text", "text": "e.g. Tivoli"},
                },
            },
            {
                "type": "input",
                "block_id": "diff",
                "optional": True,
                "label": {"type": "plain_text", "text": DIFFICULTY_LABEL_FIELD},
                "element": {
                    "type": "plain_text_input",
                    "action_id": "difficulty",
                    "max_length": 2,
                    "placeholder": {"type": "plain_text", "text": DIFFICULTY_HINT},
                },
            },
            {
                "type": "input",
                "block_id": "border_design",
                "optional": True,
                "label": {"type": "plain_text", "text": "Border design"},
                "element": {
                    "type": "plain_text_input",
                    "action_id": "val",
                    "placeholder": {"type": "plain_text", "text": "e.g. Greek Key"},
                },
            },
            {
                "type": "input",
                "block_id": "border_diff",
                "optional": True,
                "label": {"type": "plain_text", "text": DIFFICULTY_LABEL_BORDER},
                "element": {
                    "type": "plain_text_input",
                    "action_id": "difficulty",
                    "max_length": 2,
                    "placeholder": {"type": "plain_text", "text": DIFFICULTY_HINT},
                },
            },
            {
                "type": "context",
                "elements": [{"type": "mrkdwn",
                              "text": "Creating the job starts its setup - getting it ready to "
                                      "work. The card goes to your DMs."}]
            },
        ]
    }
    )

def _typed(values, block_id, action_id):
    """
    What the maker typed into one box, or None if that box was not on the form
    they submitted.

    A modal a maker already had open when the form changed still submits the
    blocks it was built with. Reading those defensively means an older form
    lands as the job it describes - a field job with no border on it - rather
    than raising on a block that was never there.
    """
    block = values.get(block_id) or {}
    return (block.get(action_id) or {}).get("value")


@app.view("trk_track_step_2")
def handle_step_2(ack, body, client):
    user_id = body["user"]["id"]
    vals = body["view"]["state"]["values"]

    prev_data = json.loads(body["view"]["private_metadata"])
    team_channel_id = prev_data["channel_id"]

    design = (_typed(vals, "design", "val") or "").strip()
    border_design = (_typed(vals, "border_design", "val") or "").strip()

    # A job is what the diagram has on it, and a diagram with neither a field
    # nor a border is not a job. Said against the first box, because that is
    # where a maker's eye is when they have filled in nothing.
    if not design and not border_design:
        ack(response_action="errors", errors={
            "design": "Fill in the field, the border, or both - a job needs at least one."
        })
        return

    difficulty, difficulty_error = read_difficulty(_typed(vals, "diff", "difficulty"))
    if difficulty_error:
        ack(response_action="errors", errors={"diff": difficulty_error})
        return
    border_difficulty, border_difficulty_error = read_difficulty(
        _typed(vals, "border_diff", "difficulty"))
    if border_difficulty_error:
        ack(response_action="errors", errors={"border_diff": border_difficulty_error})
        return

    # A difficulty for a lane the diagram does not have describes nothing.
    if difficulty and not design:
        ack(response_action="errors", errors={
            "design": "Name the field design, or clear the field difficulty."
        })
        return
    if border_difficulty and not border_design:
        ack(response_action="errors", errors={
            "border_design": "Name the border design, or clear the border difficulty."
        })
        return

    ack(response_action="clear")

    task_id = database.create_task(
        user_id=user_id,
        channel_id=team_channel_id,
        customer_name=prev_data["customer_name"],
        invoice_number=prev_data["invoice_number"],
        task_description=prev_data["task_description"],
        due_date=prev_data["due_date"],
        design=design or None,
        difficulty=difficulty,
        border_design=border_design or None,
        border_difficulty=border_difficulty,
    )

    # Submitting this form is the handover into the workshop: the maker has the
    # job and is already getting it ready. So the setup timer is running by the
    # time the card appears, and there is no "Start" button - there is nothing
    # left to start.
    task = database.get_task(task_id)

    # chat_postMessage accepts a user id and resolves the DM itself, returning
    # the real D... conversation id in result["channel"]. conversations_open
    # would need the im:write scope, which the LMSA Slack app does not hold.
    text, blocks = job_card(task)
    result = client.chat_postMessage(channel=user_id, text=text, blocks=blocks)

    # Saving the timestamp
    database.update_message_ts(task_id, result["channel"], result["ts"])

    client.chat_postEphemeral(
        channel=team_channel_id,
        user=user_id,
        text=(f"T-{task_id} is yours and the setup clock is running. The card is in your DMs "
              f"with the bot.")
    )

# ---------------------------------------------------------------------------
# Working the job: starting, pausing, cutting, switching
# ---------------------------------------------------------------------------
# Exactly one piece of work accrues at a time. Pause means "not working this
# job for now"; Switch work means "still on this job, on something else".
# Neither finishes anything. Cutting is measured inside sheeting and leaves
# the sheeting timer running.

@app.action("trk_start_setup")
@app.action("trk_start_production")
def handle_start(ack, body, client):
    """
    Move onto a piece of work and start timing it.

    One handler behind every button that does that: Start field sheeting,
    Resume, Start border setup and Back to field sheeting. The two action ids
    exist because Slack will not accept the same one twice in a message, not
    because the work differs - the value carries which lane and which activity,
    and a button posted by an earlier version of the tracker says only which
    job, so the lane the job is on is what resumes.
    """
    ack()
    task_id, phase, activity = read_work_value(body["actions"][0]["value"])
    user_id = body["user"]["id"]
    channel_id = body["container"]["channel_id"]

    task = resolve_job(client, body, task_id, user_id, channel_id)
    if task is None:
        return

    outcome = database.start_work(task_id, phase, activity or "production")
    if outcome != "started":
        client.chat_postEphemeral(
            channel=channel_id,
            user=user_id,
            text=refusal_text(outcome, task, phase or task["current_phase"]),
        )
        return

    update_card(client, database.get_task(task_id), channel_id)


@app.action("trk_stop_task")
def handle_stop(ack, body, client):
    """
    Pause. The maker is not working on this job for the moment.

    Whatever was being timed stops, including any cutting that was being
    measured inside it - there is nothing left for cutting to be inside. The
    job keeps everything it has recorded and nothing about it is finished.
    """
    ack()
    task_id, _phase, _activity = read_work_value(body["actions"][0]["value"])
    user_id = body["user"]["id"]
    channel_id = body["container"]["channel_id"]

    task = resolve_job(client, body, task_id, user_id, channel_id)
    if task is None:
        return

    database.stop_work(task_id)
    update_card(client, database.get_task(task_id), channel_id)


@app.action("trk_start_cutting")
def handle_start_cutting(ack, body, client):
    """
    The maker goes and cuts tiles for a while.

    The sheeting timer keeps running, because they are still working this job -
    they have gone downstairs to cut for it. This measures how much of that
    time was spent cutting; it never takes time away from the sheeting.
    """
    ack()
    task_id, _phase, _activity = read_work_value(body["actions"][0]["value"])
    user_id = body["user"]["id"]
    channel_id = body["container"]["channel_id"]

    task = resolve_job(client, body, task_id, user_id, channel_id)
    if task is None:
        return

    outcome = database.start_cutting(task_id)
    if outcome != "started":
        client.chat_postEphemeral(
            channel=channel_id,
            user=user_id,
            text=refusal_text(outcome, task, task["current_phase"]),
        )
        return

    update_card(client, database.get_task(task_id), channel_id)


@app.action("trk_stop_cutting")
def handle_stop_cutting(ack, body, client):
    """Back upstairs. The sheeting was running throughout and still is."""
    ack()
    task_id, _phase, _activity = read_work_value(body["actions"][0]["value"])
    user_id = body["user"]["id"]
    channel_id = body["container"]["channel_id"]

    task = resolve_job(client, body, task_id, user_id, channel_id)
    if task is None:
        return

    outcome = database.stop_cutting(task_id)
    if outcome != "stopped":
        client.chat_postEphemeral(
            channel=channel_id,
            user=user_id,
            text=refusal_text(outcome, task, task["current_phase"]),
        )
        return

    update_card(client, database.get_task(task_id), channel_id)


@app.action("trk_switch_work")
def handle_switch_work(ack, body, client):
    """
    Ask what the maker is moving on to.

    A form rather than a button per destination, because the important part is
    what it says before the press: this pauses what you are doing, it does not
    finish it, and you can come back. A row of buttons cannot say that, and
    "Start Packing" on a field card read to a maker as leaving the field
    behind for good.
    """
    ack()
    task_id, _phase, _activity = read_work_value(body["actions"][0]["value"])
    user_id = body["user"]["id"]
    channel_id = body["container"]["channel_id"]

    task = resolve_job(client, body, task_id, user_id, channel_id)
    if task is None:
        return

    options = switch_destinations(task)
    if not options:
        client.chat_postEphemeral(
            channel=channel_id,
            user=user_id,
            text="There is no other work to move to on T-" + str(task_id) + " just now.",
        )
        return

    here = task.get("working_on")
    if here:
        heading = (
            "You are on *" + work_name(here["phase"], here["activity"]) + "*. "
            "Moving pauses it - it is *not* marked finished, everything recorded on it "
            "stays, and you can come back to it."
        )
    else:
        heading = (
            "Nothing is being timed on T-" + str(task_id) + " at the moment. "
            "Pick what you are starting."
        )

    client.views_open(
        trigger_id=body["trigger_id"],
        view={
            "type": "modal",
            "callback_id": "trk_switch_work_modal",
            "title": {"type": "plain_text", "text": "Switch work"},
            "submit": {"type": "plain_text", "text": "Move to this"},
            "close": {"type": "plain_text", "text": "Cancel"},
            "private_metadata": json.dumps({"task_id": task_id, "channel_id": channel_id}),
            "blocks": [
                {"type": "section", "text": {"type": "mrkdwn", "text": heading}},
                {
                    "type": "input",
                    "block_id": "target_block",
                    "label": {"type": "plain_text", "text": "What are you moving to?"},
                    "element": {
                        "type": "radio_buttons",
                        "action_id": "target",
                        "options": [
                            {
                                "text": {"type": "plain_text", "text": option["label"]},
                                "description": {"type": "plain_text", "text": option["blurb"]},
                                "value": option["phase"] + "|" + option["activity"],
                            }
                            for option in options
                        ],
                    },
                },
            ],
        },
    )


@app.view("trk_switch_work_modal")
def handle_switch_work_submission(ack, body, client):
    ack()
    user_id = body["user"]["id"]
    metadata = json.loads(body["view"]["private_metadata"])
    task_id = metadata["task_id"]
    channel_id = metadata["channel_id"]
    chosen = body["view"]["state"]["values"]["target_block"]["target"]["selected_option"]["value"]
    phase, activity = chosen.split("|")

    task = resolve_job(client, body, task_id, user_id, channel_id)
    if task is None:
        return

    # The job can move while the form sits open - the other surface finished
    # the lane, or took a border back. Checked again here rather than trusted
    # from when the form was built.
    if not any(o["phase"] == phase and o["activity"] == activity
               for o in switch_destinations(task)):
        client.chat_postEphemeral(
            channel=channel_id,
            user=user_id,
            text=("T-" + str(task_id) + " has moved on since that form was opened, so nothing "
                  "has been changed. Go back to the job's card and carry on from there."),
        )
        return

    outcome = database.start_work(task_id, phase, activity)
    if outcome != "started":
        client.chat_postEphemeral(
            channel=channel_id,
            user=user_id,
            text=refusal_text(outcome, task, phase),
        )
        return

    update_card(client, database.get_task(task_id), channel_id)


@app.action("trk_start_packing")
def handle_start_packing(ack, body, client):
    """
    Go and pack for a while, leaving the sheeting where it is.

    New cards say Switch work and go through the form above. This stays
    registered because a card posted by an earlier version of the tracker is
    still live in somebody's DM, and it should keep working rather than fall
    silent the moment a deployment lands.
    """
    ack()
    task_id, _phase, _activity = read_work_value(body["actions"][0]["value"])
    user_id = body["user"]["id"]
    channel_id = body["container"]["channel_id"]

    task = resolve_job(client, body, task_id, user_id, channel_id)
    if task is None:
        return

    outcome = database.start_packing(task_id)
    if outcome != "started":
        client.chat_postEphemeral(
            channel=channel_id,
            user=user_id,
            text=refusal_text(outcome, task, "packing"),
        )
        return

    update_card(client, database.get_task(task_id), channel_id)


# ---------------------------------------------------------------------------
# Finishing a lane, and the decisions that follow it
# ---------------------------------------------------------------------------
# Finishing does not move the job on by itself. The form that follows commits
# the move - which is what lets a cancelled form be reopened by pressing the
# same button again.

@app.action("trk_complete_task")
def handle_complete(ack, body, client):
    """
    Say a lane is finished, and go on to what it owes.

    The only press on a card that finishes anything. The lane closes, the card
    is rewritten so it says what comes next, and then the form that collects it
    opens - in that order, so a maker who cancels the form is left on a card
    offering it again rather than one still saying "finished".
    """
    ack()
    task_id, _phase, _activity = read_work_value(body["actions"][0]["value"])
    user_id = body["user"]["id"]
    channel_id = body["container"]["channel_id"]

    task = resolve_job(client, body, task_id, user_id, channel_id)
    if task is None:
        return

    # Read before the lane closes: afterwards every lane looks finished, and
    # the difference between finishing one and returning to its form is gone.
    announce = lane_finish_is_new(task)

    outcome = database.complete_task(task_id)
    if outcome != "completed":
        client.chat_postEphemeral(
            channel=channel_id,
            user=user_id,
            text=refusal_text(outcome, task),
        )
        return

    updated_task = database.get_task(task_id)
    phase = updated_task["current_phase"]
    metadata = json.dumps({
        "task_id": task_id,
        "dm_channel_id": channel_id,
        "team_channel_id": task["channel_id"],
    })
    update_card(client, updated_task, channel_id)

    if phase == "field_sheeting":
        if announce:
            client.chat_postMessage(
                channel=task["channel_id"],
                text=f"T-{task_id} {updated_task['customer_name']}: field sheeting finished",
                blocks=[{
                    "type": "section",
                    "text": {
                        "type": "mrkdwn",
                        "text": (f"*T-{task_id}  {updated_task['customer_name']}*\n"
                                 f"Field sheeting finished by <@{user_id}>\n"
                                 f"{lane_report(updated_task, 'field_sheeting')}")
                    }
                }]
            )

        client.views_open(
            trigger_id=body["trigger_id"],
            view={
                "type": "modal",
                "callback_id": "trk_border_modal",
                "title": {"type": "plain_text", "text": "Border details"},
                "submit": {"type": "plain_text", "text": "Save the border details"},
                "close": {"type": "plain_text", "text": "Cancel"},
                "private_metadata": metadata,
                "blocks": [
                    {
                        "type": "section",
                        "text": {
                            "type": "mrkdwn",
                            "text": (f"*Field sheeting finished.* "
                                     f"{lane_report(updated_task, 'field_sheeting')}\n"
                                     f"Now the border - or say there isn't one.")
                        }
                    },
                    {
                        "type": "input",
                        "block_id": "border_design_block",
                        "label": {"type": "plain_text", "text": "Border design"},
                        "element": {
                            "type": "plain_text_input",
                            "action_id": "border_design",
                            # Prefilled when the diagram already said what the
                            # border is. Known at intake, worked now - and
                            # still correctable here, which is the point of
                            # asking again rather than assuming.
                            "initial_value": updated_task.get("border_design") or ""
                        }
                    },
                    {
                        "type": "input",
                        "block_id": "border_diff_block",
                        "label": {"type": "plain_text", "text": DIFFICULTY_LABEL_BORDER},
                        "element": {
                            "type": "plain_text_input",
                            "action_id": "border_difficulty",
                            "max_length": 2,
                            "placeholder": {"type": "plain_text", "text": DIFFICULTY_HINT},
                            "initial_value": updated_task.get("border_difficulty") or ""
                        }
                    },
                    # Usually a millimetre size, but "template" and split sizes
                    # are real entries too. Optional here because the border
                    # jig is often established during the border setup, and the
                    # card can take it then.
                    {
                        "type": "input",
                        "block_id": "border_jig_block",
                        "optional": True,
                        "label": {"type": "plain_text", "text": "Border jig or template"},
                        "element": {
                            "type": "plain_text_input",
                            "action_id": "border_jig",
                            "placeholder": {"type": "plain_text",
                                            "text": "e.g. 49.6 or template - leave blank if not known yet"}
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
                                "text": {"type": "plain_text", "text": "This job has no border"},
                                "action_id": "trk_no_border",
                                "value": str(task_id)
                            }
                        ]
                    }
                ]
            }
        )

    elif phase == "border_sheeting":
        if announce:
            client.chat_postMessage(
                channel=task["channel_id"],
                text=f"T-{task_id} {updated_task['customer_name']}: border finished",
                blocks=[{
                    "type": "section",
                    "text": {
                        "type": "mrkdwn",
                        "text": (f"*T-{task_id}  {updated_task['customer_name']}*\n"
                                 f"Border finished by <@{user_id}>\n"
                                 f"{lane_report(updated_task, 'border_sheeting')}")
                    }
                }]
            )

        client.views_open(
            trigger_id=body["trigger_id"],
            view=packing_modal_view(
                metadata,
                (f"*Border finished.*\n"
                 f"{lane_report(updated_task, 'field_sheeting')}\n"
                 f"{lane_report(updated_task, 'border_sheeting')}\n\n"
                 f"That leaves the packing.")
            )
        )

    elif phase == "packing":
        client.views_open(
            trigger_id=body["trigger_id"],
            view=notes_modal_view(
                metadata,
                (f"*Packing finished.* "
                 f"{lane_report(updated_task, 'packing')}\n"
                 f"Anything worth recording before this job closes?")
            )
        )


@app.view("trk_border_modal")
def handle_border_submission(ack, body, client):
    user_id = body["user"]["id"]
    vals = body["view"]["state"]["values"]
    metadata = json.loads(body["view"]["private_metadata"])
    task_id = metadata["task_id"]
    dm_channel_id = metadata["dm_channel_id"]

    border_design = vals["border_design_block"]["border_design"]["value"]
    border_difficulty, difficulty_error = read_difficulty(
        vals["border_diff_block"]["border_difficulty"]["value"])
    if difficulty_error:
        ack(response_action="errors", errors={"border_diff_block": difficulty_error})
        return
    ack()
    border_jig = (vals["border_jig_block"]["border_jig"]["value"] or "").strip()

    try:
        database.move_to_border_phase(task_id, border_design, border_difficulty, border_jig)
    except database.TrackerRefused as refusal:
        if refusal.reason != "another_phase_running":
            raise
        # A timer is running on the job, so the border cannot be put back
        # underneath it. Say so; the form can be filled in again once it stops.
        client.chat_postEphemeral(
            channel=dm_channel_id,
            user=user_id,
            text=("A timer is running on this job, so the border details have not been saved. "
                  "Pause it first, then press the button on the card to get this form back.")
        )
        return

    task = database.get_task(task_id)
    repost_card(
        client, task, dm_channel_id,
        note="*Border details saved.* " + border_route(task),
    )


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

    # The same three guards resolve_job() applies, checked by hand because this
    # arrives from a modal: a form submission carries no container.channel_id,
    # so the channel has to come from private_metadata instead.
    task = database.get_task(task_id)

    if task is None:
        client.chat_postEphemeral(
            channel=dm_channel_id,
            user=user_id,
            text="That job is not there any more - it may have been deleted."
        )
        return

    if task["user_id"] != user_id:
        client.chat_postEphemeral(
            channel=dm_channel_id,
            user=user_id,
            text="This is <@" + task["user_id"] + ">'s job, so only they can change it."
        )
        return

    database.skip_border_phase(task_id)

    updated_task = database.get_task(task_id)

    client.chat_postMessage(
        channel=team_channel_id,
        text=f"T-{task_id} {updated_task['customer_name']}: no border on this job",
        blocks=[{
            "type": "section",
            "text": {
                "type": "mrkdwn",
                "text": (f"*T-{task_id}  {updated_task['customer_name']}*\n"
                         f"No border on this job - <@{user_id}>")
            }
        }]
    )

    client.views_update(
        view_id=body["view"]["id"],
        view=packing_modal_view(
            json.dumps(metadata),
            (f"*No border on this job.*\n"
             f"{lane_report(updated_task, 'field_sheeting')}\n\n"
             f"That leaves the packing.\n\n"
             f"_Pressed this by mistake? Cancel, then press the button on the card to get the "
             f"border form back._")
        )
    )


@app.action("trk_undo_no_border")
def handle_undo_no_border(ack, body, client):
    """
    Take back a "no border" from the packing card.

    The refusal is shown, never hidden. If the correction did not happen the
    card is left exactly as it was: rebuilding it as though it had worked would
    leave the maker believing they have a border back when the record still
    says the border was skipped.
    """
    ack()
    task_id, _phase, _activity = read_work_value(body["actions"][0]["value"])
    user_id = body["user"]["id"]
    channel_id = body["container"]["channel_id"]

    task = resolve_job(client, body, task_id, user_id, channel_id)
    if task is None:
        return

    outcome = database.revert_border_skip(task_id)

    if outcome != "reverted":
        if outcome == "border_skip_not_reversible":
            text = ("The packing is already finished on this job, so the border cannot be "
                    "reopened from here. Nothing has been changed - put what happened in the "
                    "job's notes and tell a supervisor.")
        elif outcome == "another_phase_running":
            text = ("A timer is running on this job. Pause it first, then press "
                    "'Border after all' again. Nothing has been changed.")
        elif outcome == "border_not_skipped":
            text = "This job is not marked as having no border, so there is nothing to undo."
        else:
            text = "That could not be undone, so nothing has been changed."
        client.chat_postEphemeral(channel=channel_id, user=user_id, text=text)
        return

    updated_task = database.get_task(task_id)
    repost_card(
        client, updated_task, channel_id,
        note=("*Border back on this job.* Any packing time already recorded stays. "
              "Press the button below for the border details."),
    )


@app.view("trk_packing_modal")
def handle_packing_submission(ack, body, client):
    ack()
    user_id = body["user"]["id"]
    metadata = json.loads(body["view"]["private_metadata"])
    task_id = metadata["task_id"]
    dm_channel_id = metadata["dm_channel_id"]

    outcome = database.move_to_packing_phase(task_id)
    if outcome != "moved":
        # This form was opened before something else changed the job - most
        # likely the border was put back on another device. Say so and change
        # nothing; the job's card is still live and still correct.
        client.chat_postEphemeral(
            channel=dm_channel_id,
            user=user_id,
            text=("This job's border has changed since that form was opened, so it has not been "
                  "moved to packing and nothing has been changed. Go back to the job's card and "
                  "carry on from there.")
        )
        return

    task = database.get_task(task_id)
    repost_card(
        client, task, dm_channel_id,
        note="*On to the packing.* Start it when you are ready.",
    )


@app.view("trk_notes_modal")
def handle_notes_submission(ack, body, client):
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

    total_time = database.format_duration(task["total_elapsed"])

    client.chat_update(
        channel=dm_channel_id,
        ts=task["message_ts"],
        text=f"T-{task_id} is finished.",
        blocks=[
            {
                "type": "header",
                "text": {"type": "plain_text", "text": header_text(task, "  -  finished")},
            },
            {
                "type": "section",
                "text": {
                    "type": "mrkdwn",
                    "text": (f"*Total time: {total_time}*\n"
                             f"The full breakdown has gone to the team channel."),
                },
            },
        ]
    )

    client.chat_postMessage(
        channel=team_channel_id,
        text=f"T-{task_id} {task['customer_name']} finished by <@{user_id}> - {total_time}",
        blocks=job_summary_blocks(task, user_id, general_notes, issues),
    )


# The jig. A lane sometimes needs another one part way through: maybe the first
# turned out wrong and was swapped, maybe two sizes are genuinely needed
# together. Either way the earlier jig really was used, so this ADDS a record
# next to it - it never overwrites one. Typing mistakes are fixed through Edit
# instead, which changes the value it names.
#
# It lives on every working card because the jig is normally established during
# the setup, which is after the job was logged and can be well after the
# sheeting started.

# ---------------------------------------------------------------------------
# Jig and template
# ---------------------------------------------------------------------------
# A phase records any number of jigs, in the order they were used. Adding
# appends, because a jig that was genuinely used stays on the record; a
# typing mistake is corrected through Edit instead.

@app.action("trk_add_jig")
def handle_add_jig(ack, body, client):
    ack()
    task_id, _phase, _activity = read_work_value(body["actions"][0]["value"])
    user_id = body["user"]["id"]
    channel_id = body["container"]["channel_id"]

    task = resolve_job(client, body, task_id, user_id, channel_id)
    if task is None:
        return
    phase = task["current_phase"]

    # Remember which phase the card was on, so the right card comes back
    # after the modal
    metadata = json.dumps({"task_id": task_id, "channel_id": channel_id, "phase": phase})

    # On the field there is only one place a jig can go. From the border
    # onwards the maker chooses, because field work genuinely can continue with
    # another jig AFTER the field was finished - that is still an added jig,
    # not a correction, and the earlier one must stay.
    blocks = []
    if phase != "field_sheeting":
        field_option = {
            "text": {"type": "plain_text", "text": "Field sheeting"},
            "value": "field_sheeting"
        }
        border_option = {
            "text": {"type": "plain_text", "text": "Border sheeting"},
            "value": "border_sheeting"
        }
        # A border that did not happen used no jig, and storage refuses one,
        # so it is not offered - a choice that always errors is worse than no
        # choice.
        phase_element = {
            "type": "static_select",
            "action_id": "jig_phase",
            "placeholder": {"type": "plain_text", "text": "Field or border?"},
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
            "label": {"type": "plain_text", "text": "Which work used it?"},
            "element": phase_element
        })
    blocks.append({
        "type": "input",
        "block_id": "jig_block",
        "label": {"type": "plain_text", "text": "Jig or template"},
        "element": {
            "type": "plain_text_input",
            "action_id": "jig_size",
            "placeholder": {"type": "plain_text", "text": "e.g. 49.6, 49.4/49.8, or template"}
        }
    })

    client.views_open(
        trigger_id=body["trigger_id"],
        view={
            "type": "modal",
            "callback_id": "trk_add_jig_modal",
            "title": {"type": "plain_text", "text": "Jig or template"},
            "submit": {"type": "plain_text", "text": "Save"},
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
            text=f"'{jig_size}' was not recorded - that job is no longer open."
        )
        return

    update_card(client, task, channel_id, note=f"*Jig recorded: {jig_size}*")


#Delete Button
# ---------------------------------------------------------------------------
# Corrections: editing a job, and cancelling one
# ---------------------------------------------------------------------------
# Editing changes the values a job was given. Cancelling keeps the job and
# everything recorded on it - nothing is ever deleted outright.

@app.action("trk_delete_task")
def handle_delete(ack, body, client):
    ack()
    task_id, _phase, _activity = read_work_value(body["actions"][0]["value"])
    user_id = body["user"]["id"]
    task = database.get_task(task_id)
    channel_id = body["container"]["channel_id"]

    if task is None:
        client.chat_postEphemeral(
            channel=channel_id,
            user=user_id,
            text="That job is no longer on your list."
        )
        return

    # A job is only ever taken off the list by the maker whose job it is. There
    # is no supervisor path anywhere in the tracker yet, so this check is the
    # whole of the rule - it is not a friendly message in front of a second one
    # enforced further down.
    if task["user_id"] != user_id:
        client.chat_postEphemeral(
            channel=channel_id,
            user=user_id,
            text="You can only take your own jobs off the list."
        )
        return

    # Delete from the database
    database.delete_task(task_id)

    # The card is replaced by what actually happened. Nothing is destroyed -
    # LMSA cancels the job and keeps it, which is exactly why a maker may safely
    # use this on a job that turned out to be a mistake. Saying "deleted" here
    # contradicted the dialog they had just agreed to, one press earlier.
    client.chat_update(
        channel=channel_id,
        ts=task["message_ts"],
        text=f"Job T-{task_id} was cancelled by <@{user_id}>. Its recorded time was kept.",
        blocks=[
            {
                "type": "section",
                "text": {
                    "type": "mrkdwn",
                    "text": (f"*Job T-{task_id} was cancelled by <@{user_id}>.*"
                             "\nIts recorded time and history have been kept.")
                }
            }
        ]
    )


@app.action("trk_edit_task")
def handle_edit(ack, body, client):
    ack()
    task_id, _phase, _activity = read_work_value(body["actions"][0]["value"])
    user_id = body["user"]["id"]
    task = database.get_task(task_id)
    channel_id = body["container"]["channel_id"]

    if task is None:
        client.chat_postEphemeral(
            channel=channel_id,
            user=user_id,
            text="That job is no longer on your list."
        )
        return

    if task["user_id"] != user_id:
        client.chat_postEphemeral(
            channel=channel_id,
            user=user_id,
            text="You can only edit your own jobs."
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
            "label": {"type": "plain_text", "text": f"Field jig {i} (or template)"},
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
            "label": {"type": "plain_text", "text": f"Border jig {i} (or template)"},
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
            "title": {"type": "plain_text", "text": "Edit job"},
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
                    "label": {"type": "plain_text", "text": "Job description"},
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
                    "label": {"type": "plain_text", "text": DIFFICULTY_LABEL_FIELD},
                    "element": {
                        "type": "plain_text_input",
                        "action_id": "difficulty",
                        "max_length": 2,
                        "placeholder": {"type": "plain_text", "text": DIFFICULTY_HINT},
                        "initial_value": task["difficulty"] or ""
                    }
                },
                *jig_blocks,
                {
                    "type": "input",
                    "block_id": "date_block",
                    "optional": True,
                    # The same field as the intake form, so the same words, the
                    # same format and the same rules. Both build the box from
                    # DUE_DATE_LABEL and DUE_DATE_HINT so one box cannot end up
                    # promising two different things.
                    "label": {"type": "plain_text", "text": DUE_DATE_LABEL},
                    "element": {
                        "type": "plain_text_input",
                        "action_id": "due_date",
                        "placeholder": {"type": "plain_text", "text": DUE_DATE_HINT},
                        "initial_value": due_date_supplied(task) or ""
                    }
                }
            ]
        }
    )


@app.view("trk_edit_task_modal")
def handle_edit_submission(ack, body, client):
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
    difficulty, difficulty_error = read_difficulty(
        vals["difficulty_block"]["difficulty"]["value"])
    if difficulty_error:
        ack(response_action="errors", errors={"difficulty_block": difficulty_error})
        return
    # What the job said before the modal opened. Read first, because the due
    # date needs it too.
    task_before = database.get_task(task_id)

    # Same box, same rules as the intake form. One exception, and it is about
    # history rather than about dates: a row written before this screen asked
    # for a real date may hold free text, which the form pre-fills. Judging
    # that on submit would stop a maker fixing a customer's name until they had
    # also rewritten a due date they never touched. So a value that comes back
    # exactly as it was stored passes through untouched - that is not new
    # typing, and nothing about it is being changed.
    typed_due_date = vals["date_block"]["due_date"]["value"]
    stored_due_date = (task_before["due_date"] if task_before else None) or ""
    if stored_due_date.strip() and (typed_due_date or "").strip() == stored_due_date.strip():
        due_date = stored_due_date
    else:
        due_date, due_date_error = read_due_date(typed_due_date)
        if due_date_error:
            ack(response_action="errors", errors={"date_block": due_date_error})
            return
    ack()

    # What each jig said before the modal opened, so only boxes the maker
    # actually changed get corrected
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

    # Refresh the card so the corrections show. The job's own state decides
    # what it says and which buttons it keeps - an edit is not a change of
    # what the maker is doing.
    task = database.get_task(task_id)
    if task is not None and task["current_phase"] != "completed":
        update_card(client, task, channel_id, note="*Details updated.*")


# ---------------------------------------------------------------------------
# Running the tracker on its own
# ---------------------------------------------------------------------------
# Inside LMSA the relay delivers events instead, so nothing below runs there.
# This is the path for running the tracker standalone against Slack.

if __name__ == "__main__":
    database.setup_database()
    handler = SocketModeHandler(app, os.environ.get("SLACK_APP_TOKEN"))
    print("Trackbot is running!")
    handler.start()
