# Running this tracker inside LMSA

This file records how the tracker in this repository was integrated into LMSA, the
London Mosaic Slack assistant, and where the line sits between "the original
tracker, made to work inside LMSA" and "new features added afterwards".

It lives on `feature/lmsa-integration` only. It is not on `main`.

## Original baseline

The original tracker is `main` at commit `0b736c0`.

**It has not been modified and will not be.** No commit, merge, rebase, reset or
force-push has been made to `main`, and none will be. It stays as the reference
copy of the tracker as it was originally written.

All the integration work described below lives on `feature/lmsa-integration`,
which starts from that commit.

## LMSA-compatible baseline

**`845d8de` on `feature/lmsa-integration` is the point at which the tracker was
considered fully integrated and working.** It is five commits ahead of `main`.

Everything in those five commits exists for one reason: to let the same tracker,
with the same cards, the same modals, the same buttons and the same wording, run
inside LMSA instead of standing on its own. Nothing was added to the workflow and
nothing was taken away from it.

What changed, and why:

**The card is posted to the user id instead of opening a DM** (`476547c`). The
LMSA Slack app does not hold the `im:write` scope, so `conversations_open` is not
available to it. `chat_postMessage` accepts a user id, resolves the DM itself and
returns the real `D…` conversation id, which is then stored so the card can be
updated later.

**The export finds the existing DM instead of opening one** (`aa18778`). Same
reason. `users_conversations` needs only `im:read`. If no DM exists yet the maker
is told to start a job first, because the upload will not accept a raw user id.

**The database functions save through LMSA instead of the local SQLite file**
(`2fce308`). This is the largest change. `database.py` keeps every function name
and signature `app.py` already calls, and still returns dictionaries with the same
keys, so `app.py` did not have to change. Underneath, each function now makes a
small HTTP request to LMSA on the local machine, and LMSA owns the database. The
rows, the timing and the history live in Postgres rather than `trackbot.db`.

**The due date is kept exactly as it was typed** (`a074da1`). The due date field
is free text and people write things like `ASAP`, `Friday` and
`01/09/26 (if paint arrives)`. Turning those into real dates lost most of them, so
the text is stored as written and shown back unchanged. A parsed date is kept
alongside it only when the text happens to be one.

**The button and form ids carry a `trk_` prefix** (`bd0d9a3`). LMSA runs several
Slack features on one app, so it needs to tell which clicks belong to the tracker
and pass only those through. The prefix is how it tells.

**The phase the maker is on is read from LMSA rather than worked out** (`845d8de`).
This one is worth reading properly, because it was a real bug. LMSA had started
working the current phase out from which phases were finished. That sounds the
same as storing it, and it is not: completing a phase made it move, and
`handle_complete` completes the phase *before* it asks which phase it is on. So
every branch landed a step early — pressing Complete on Field opened the Packing
form instead of the Border form, the border design and difficulty were never
collected, and once all three phases were finished nothing matched at all, so
Complete quietly did nothing and the job could not be finished. LMSA now stores
the phase the maker is on, exactly as the original tracker does, and only a
submitted form moves it. That is also what makes cancelling a form and pressing
Complete again bring the same form back.

There is one more thing worth knowing about, which is not visible in the workflow.
Every request now runs inside a small wrapper (`slack_request` in `database.py`)
that notes which Slack delivery is being handled, so that if Slack sends the same
click twice it is recognised as the same action rather than applied twice. It
changes nothing a maker can see.

The deeper LMSA-side detail — the database schema, the internal API, the audit
trail and the deployment — is documented on the LMSA side, in
`docs/investigations/slack-tracker/HANDOVER.md`. There is no need to read it to
work on the Python here.

## What LMSA does differently on purpose

Three behaviours differ from the original, deliberately, and should not be
"fixed" back:

- **Delete keeps the row.** The original deletes the task and its time segments
  outright. LMSA marks the job cancelled instead, so a morning's timings are not
  destroyed. To the maker it behaves the same: the card goes, and the job stops
  blocking a new one.
- **Finishing a phase no longer marks the whole job completed.** The original sets
  the job's status to `completed` every time a phase finishes, which quietly drops
  its own one-job-at-a-time check part way through a job, and puts unfinished jobs
  into the Excel export. In LMSA the job stays open until the Notes form is
  submitted.
- **Elapsed time is added up from the timing segments when it is asked for**,
  rather than stored in a column. The original declares a `total_elapsed` column
  and never writes to it, so a stored total was already out of step with the
  segments it claimed to summarise.

## Development boundary

**Everything above was needed to get the original tracker working inside LMSA.
Changes listed after this section are features added after the original baseline
was successfully integrated.**

Anything appended below is new development, not baseline recovery.

Future Python work on the tracker branches from `feature/lmsa-integration`, never
from `main`, and comes back into that line once it is proven.

# Post-baseline feature history

Nothing yet. The first entry will be added when the first post-baseline feature
lands.

For each feature, record: the feature name, the date, what changed from a maker's
point of view, whether it was agreed with Luis beforehand, anything it needs on
the LMSA side, and the commit or commits it arrived in.
