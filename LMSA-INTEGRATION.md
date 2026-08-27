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

**The code describes the tracker, never how we happened to build it.**
Comments, function names, test wording and commit messages explain actual
workshop behaviour and call features by their real names — "Jig Size", "No
Border", "Packing interruption" — never internal delivery shorthand: no
F1/F2/F3-style numbers, no "slice 1", no "phase 2 of the work", no waves,
stages, milestones or rollout language. Those labels describe our development
order, which means nothing to anyone reading this code later. The word
"phase" stays, of course, where it means a real workflow phase — Field,
Border and Packing genuinely are phases in this tracker. Python comments
describe what the workshop workflow is doing, why a state transition exists,
what a maker can do, and what safety condition is being protected. Exact Git
commit ids are used only where something technical genuinely needs them (the
vendor pin on the LMSA side, deployment records), never as a feature's name.

Maintained Python commits use truthful human authorship and this repository's
normal commit style: a plain sentence saying what the change lets the tracker
do. Do not add AI co-author, session, model, tool or generation metadata to
commits, and never impersonate the original author.

# Post-baseline feature history

For each feature, record: the feature name, the date, what changed from a maker's
point of view, whether it was agreed with Luis beforehand, anything it needs on
the LMSA side, and the commit or commits it arrived in.

## Delivery identity on Bolt's worker threads (2026-08-26)

Integration hardening, not a new feature — nothing a maker sees changed.

The middleware that notes which Slack delivery is being handled runs on the
thread that receives the request, but Bolt runs the handler itself on a worker
thread, and the note was already cleared by the time the handler started. So in
real use the handlers never saw it, and the replay protection LMSA builds on it
never engaged. The fix hands Bolt an executor (`database.listener_executor()`)
that picks the note up on the receiving thread and carries it onto the worker
thread for exactly the length of the handler.

Not agreed with Luis beforehand because it changes no tracker behaviour; it
makes an LMSA-side guarantee real. LMSA side: none beyond re-vendoring —
the receiving column (`job_events.idempotency_key`) already existed.

Arrived in the commit that adds this entry, on `feature/listener-idempotency`,
merged into `feature/lmsa-integration`.

## Jig Size for Field and Border (2026-08-26)

The first agreed post-baseline feature. What changed for a maker:

- The Field details form and the Border details form each gained an optional
  "Jig Size (mm)" box. Usually a millimetre size like 49.6, but the box takes
  whatever is real - "49.4/49.8" and "template" are legitimate entries, so it
  is not restricted to numbers.
- Every working card (Field, Border and Packing) gained an "Add Jig" button
  for the times a phase genuinely uses another jig - one swapped mid-run
  after a problem, or two needed together. From the Border card onwards the
  box asks which phase used it (Border is pre-picked on the Border card;
  from Packing the maker chooses), so Field work that genuinely continued
  AFTER Field was completed still gets its jig recorded properly. Add Jig
  ADDS a record next to the existing one; the jig that was already used
  stays on the card, oldest first ("49.6 / 50").
- Edit now shows one pre-filled box per recorded jig, so a mistyped value can
  be corrected later - even after that phase has finished. A correction
  changes only the box it names, and LMSA keeps the old value in its audit
  trail. Fixing a typo is Edit; a genuinely different jig is Add Jig.
- The completed-job summary and the Excel export show the jig values for both
  phases.

Agreed with Luis beforehand: yes - jig size in millimetres for Field and
Border, editable in later phases, changes audited, was confirmed in the
creator review. Tom's workshop detail refined the shape afterwards: one phase
can use several jigs, and the value cannot be numeric-only.

LMSA side: a new tracker.phase_jig_sizes table (0..many ordered values per
phase), two API operations (add, correct) with the usual audit and replay
protection, and the completed/export projection. Arrived in the single
feature commit on feature/jig-size, merged into feature/lmsa-integration.
## No Border (2026-08-26)

Some jobs genuinely have no border. Until now the only way to get such a job
past the border was to complete the border phase, which recorded a phase that
was worked for zero seconds - indistinguishable, on every card and in the
export, from a border that really was worked and happened to be quick.

What changed for a maker:

- The Border details form gained a **No Border** button next to the design and
  difficulty boxes. It is offered at that moment and nowhere else: the job
  paper does not always say whether there is a border, but by the time field
  sheeting is finished the maker knows. Intake is deliberately left alone.
- Choosing it turns the same form into the Packing one, so the job carries
  straight on. The team channel gets a line saying the job has no border, the
  same way it gets one when a phase finishes.
- Everywhere a border time would be shown - the packing card, the finished-job
  summary, the Excel export - a job with no border now says **No Border**
  instead of "0 h 0m 0s".
- Add Jig stops offering Border on such a job. There was no border, so there
  was no jig.

Changed your mind:

- Cancelling the Packing form leaves everything as it was, on a live card. Press
  Complete Phase again and the Border details form comes back, exactly as it
  does when any modal is cancelled.
- Once the Packing card is showing, it carries a **Border after all** button
  while nothing has been packed yet. That puts the job back at the border
  details step.
- Once packing has genuinely started, the button is gone and the correction is
  refused with a message explaining why. Reopening a border while packing time
  is already running needs one phase paused while another is worked, which is
  the packing-interruption behaviour that has not been built yet. It refuses
  out loud rather than quietly doing nothing.

The record of what happened is kept either way: choosing no border and taking
it back are both written to the job's history, with who and when. Taking it
back does not erase the original choice.

Agreed with Luis beforehand: yes - "some jobs genuinely have no Border, add an
explicit No Border path, do not generalise it into skipping any phase" was
confirmed in the creator review. It is Border-only for exactly that reason.
Tom settled the two questions the review left open: where the maker is asked,
and how a mistake is corrected.

LMSA side: the border phase row is marked `skipped` rather than completed, so
the database itself refuses to give it a completion time and no reader can
mistake it for a zero-length border. Two API operations (skip, and revert)
with the usual audit and replay protection, a job can now finish with a skipped
border, and a skipped border cannot be started, completed, given details or
given a jig. Arrived in the single feature commit on `feature/no-border`.

## Packing interruption (2026-08-26)

Packing no longer has to wait for the sheeting to be finished. A maker doing
field or border work can break off, pack for a while, and come back - which is
how the workshop actually runs - and the times stay separate and truthful:
sheeting time stays sheeting time, packing time is packing time, and nothing
is ever added up into the wrong pot because work moved around.

What changed for a maker:

- The Field and Border cards - working or paused - gained a **Start Packing**
  button. One press: the sheeting timer stops, the packing timer starts.
  Nothing ever starts a timer on its own; the press is the start.
- While packing this way the card shows packing in progress, names the
  sheeting phase that is waiting, and shows both times side by side. Two
  buttons: **Stop Packing**, and **Back to Field Sheeting** (or Border) -
  one press each way, and there is never a moment with two timers going.
- A paused sheeting card on a job that holds packing time now shows
  "Packing Time So Far" alongside the sheeting time, so both halves of the
  day are visible on one card.
- The job itself does not move. Packing done this way is time against
  packing, not a decision that the sheeting is over - Complete Phase still
  walks the job through Field, Border and Packing in order, exactly as
  before. Pressing Complete Phase on a sheeting card while the packing timer
  is running is refused with a message saying to deal with the packing first.

This also loosened the "Border after all" rule from the No Border entry
above, which refused as soon as packing had any time - a limit that existed
only because one phase could not wait paused while another was worked. Now it
can, so the way back stays open until packing is **finished**: stop the
packing timer, press Border after all, and the border details step comes
back. The packing time already worked stays on the job - it was real labour -
and the job returns to packing afterwards through the normal steps. A
correction is still refused while the packing timer is actually running
(stop it first, the message says so), and once packing has been completed the
button is gone and the refusal explains why.

Agreed with Luis beforehand: yes - "packing is manually started and may
interrupt sheeting" and "nothing ever auto-starts a timer" are both from the
creator review, and this is built to exactly those words. The one automatic
thing is conservative: switching to packing closes the timer the maker is
walking away from, never opens one they did not ask for.

LMSA side: no schema change - the timing ledger and the phase lanes were
built for this. The segment-start operation gained an "interrupting" switch
that closes the running segment (recorded as interrupted, not stopped by
hand) in the same transaction the new one opens in, so a job can never hold
two live timers however clicks race or repeat. Completing a phase now refuses
while another phase's timer runs, and the border-skip revert refuses only for
a live timer or finished packing. Arrived in the single feature commit on
`feature/packing-interruption`.

## Setup, cutting, and the card rebuilt (2026-08-27)

The largest change to what a maker sees since the baseline. Three things, and
they belong together because they are all answers to the same complaint: the
card told the maker about the database rather than about the job.

**The job starts when it is handed over.** The details form no longer asks for
a jig size — a maker filling it in has just been given the job and normally
does not know the jig yet, and calling the box "optional" did not make asking
any less premature. Submitting the form is the handover into the workshop, so
the setup clock is running by the time the card appears and there is no
"Start" button: there is nothing left to start. The jig is asked for on the
card instead, under "Set jig / template", at the point it can be answered —
finding and testing it IS the setup.

**Setup is timed work, not pause time.** Getting the job organised, checking
what was supplied, fetching the tiles, reading the drawings, finding the jig.
It pauses and resumes like anything else, and "Start field sheeting" closes it
and opens the sheeting. It sits underneath the Field and Border work rather
than beside it, so the job still runs Field → Border → Packing and a lane's
time is its setup plus its sheeting, reported as two lines that add up.

**Cutting is measured inside the sheeting, never beside it.** A maker sheeting
a field goes downstairs and cuts tiles for twelve minutes; they never stopped
working the field, so the field timer keeps running and "Start cutting" simply
records how part of that hour was spent. Sixty minutes of field work
containing twelve minutes of cutting is sixty minutes of labour — never
seventy-two, and never forty-eight plus twelve. The cutting knows which work
it belonged to (field cutting and border cutting stay distinguishable), and it
closes when that work does, rather than running on against nothing.

**"Switch work" replaces "Start Packing".** The old button was correct and
read wrong: on a field card it looked like leaving the field behind for good.
The new one opens a short form that says what is about to happen — your work
pauses, nothing is marked finished, you can come back — and offers whatever
the job actually allows, so a finished lane and a border nobody has described
are not on the list. Packing interrupting the sheeting still works exactly as
it did; it is now one of several moves rather than a special case.

**And one card, built from what the job is.** It leads with the job, then what
the maker is working on, then the time recorded, then what they can do next,
with the rest of the job's details grouped underneath rather than given the
same weight. Finishing a lane is one press that names the lane and asks first;
once a lane is finished the same place says what is owed next ("Enter border
details"). Refusals say what the job is doing and what to press instead.

Agreed with Luis beforehand: no. This is workshop-originated, from watching
the tracker in real use — the jig question came too early, "Start" was on a
job already begun, and the cards read as a data dump. It changes no rule from
the creator review: nothing auto-starts a timer except the handover the maker
themselves submits, one person still times one thing at a time, and packing
still only ever moves the job on when the maker says so.

LMSA side: a `job_segments.activity` column (setup or production), a separate
`job_contained_segments` ledger for cutting so no total can pick it up twice,
and "one open segment per person" as a database rule rather than only a
handler check. Arrived in the single feature commit on `feature/workshop-flow`.
