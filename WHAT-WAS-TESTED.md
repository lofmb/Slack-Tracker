# What was tested, and what it proved

Luis asked to be able to see what had actually been tested rather than take it on
trust. This is that record. It is deliberately short: it says what was added,
what was proven, what kind of proof was used, and — the part that usually gets
left out — what is **not** proven.

The detailed assertions live in the proof scripts themselves, which are named
here so you can read any of them directly rather than have them summarised.

---

## 1. What has been added since the handover

Six pieces of behaviour, in the order they shipped.

| What | Whose idea | Where it runs |
| --- | --- | --- |
| **Jig size** for Field and Border | Yours — you asked for it | Live |
| **No Border** | Yours — you asked for it | Live |
| **Packing interrupting the sheeting** | Yours — you asked for it | Live |
| **Setup as timed work**, before the sheeting | Tom's, from watching real use | Deployed, under review |
| **Cutting**, measured inside the sheeting | Tom's, from watching real use | Deployed, under review |
| **One rebuilt card** and a general "Switch work" | Tom's | Deployed, under review |

Underneath all of it the data now lives in LMSA's Postgres rather than the local
SQLite file, reached over a loopback HTTP API. That was not a feature request
from anyone — it is what running your tracker inside LMSA requires. Your Slack
code is unchanged by it: same function names, same dictionary keys, same cards.

Jig size grew slightly beyond what you asked for. You asked for millimetres; a
phase now records **any number of jigs in order**, as free text, because a jig
swapped mid-run is real and `template` and `49.4/49.8` are legitimate answers.

---

## 2. What kind of proof was used

Three kinds, deliberately, because each catches what the others cannot.

**Real-boundary proofs.** A signed Slack payload goes into LMSA's actual Express
and Bolt server, through the relay, into `app.py` and `database.py` running as a
real subprocess, into the real internal API, into a real Postgres. Nothing is
mocked except Slack itself. If a button would not work in production, these fail.

**Integration suites.** The LMSA storage and API layers tested directly against a
disposable Postgres, so a timing or locking rule can be asserted precisely rather
than inferred from a card.

**A parity replay.** Your original `database.py`, imported from `main` and run
against real SQLite, driven through the same gestures as the LMSA version, so the
two can be compared rather than assumed equivalent. This is what caught a real
regression: the workflow cursor was answering one phase too far ahead.

---

## 3. What each proof covers

| Script | What it proves |
| --- | --- |
| `workshop_flow_proof.py` | The whole journey: intake, setup, sheeting, cutting inside it, switching between work, a finished job. Also checks every payload against Slack's own structural rules. |
| `jig_proof.py` | Jigs append rather than overwrite; a correction changes the value it names; a late jig lands on the phase the maker chose; both phases stay visible. |
| `no_border_proof.py` | A skipped border records that nothing happened, carries no completion time, and can never read as a border worked for zero seconds; the way back stays open until packing finishes. |
| `packing_interruption_proof.py` | Packing can cut in on field or border work, the job remembers where it was, and the maker can go back to the sheeting. |
| `local_ingress_proof.py` | The full Slack path end to end, including that a refusal travels back to the maker rather than disappearing. |
| `idempotency_proof.py` | A redelivered click does not write twice — across Bolt's real worker-thread boundary, which is where an earlier version of this check was wrong. |
| `integrated_run.py` | The vendored Python against a real LMSA and a real database, including that the adapter returns values word for word as they were stored. |
| `interface_check.py` | Every call `app.py` makes fits the adapter it is calling, and the due-date rules behave as the label promises. |
| `parity_replay.py` | Your original against the LMSA version, on the same gestures. |

There are also four LMSA-side integration suites (storage, API, workflow cursor,
schema migration), the TypeScript type check, the unit suite, a check that the
vendored copy of your files is byte-identical to your branch, and a boot probe of
the built artefact.

**Latest full local run: all green.** `workshop_flow_proof.py` alone is 109
checks. The exact counts move as checks are added, so read the run rather than
trust a number written here.

To run any of them, the commands are in `script/tracker/deploy/README.md` in the
LMSA repository.

---

## 4. What is proven by a machine, and what is not

This is the part worth reading twice.

**Proven.** Every transition above, on a real boundary, repeatedly. The timing
arithmetic — a lane's total is its setup plus its sheeting, and cutting inside
the sheeting is never added on top of it. That one person times one thing at a
time, now enforced by the database rather than only by a handler. That a
cancelled form leaves the job exactly where it was. That the same click arriving
twice writes once.

**Proven by a person, not a machine.** Whether the wording is right. A harness
drives a fake Slack that accepts any payload, so it cannot see that a form's
submit button was one character over Slack's limit and the form therefore never
opened on a real workspace. That happened, twice, on two different limits. Both
are now measured automatically — but the finding came from a person reading the
cards, not from a test.

**Not proven, because it is not decided.** Nothing protects a timer left running
overnight. There is no scheduler in the tracker at all: every timer moves when a
maker presses something. Your instinct to stop the clock at 18:00 was right about
the problem; the current thinking is to *ask* the worker rather than cut the
clock silently, because a hard stop deletes late work that really happened. It is
not designed yet.

**Known and unfixed.** The Excel export's "Export ready!" confirmation never
sends. It sits after an unconditional `raise`, so Python cannot reach it — the
same shape is in your original, so it has never fired for anyone. It is recorded
as its own small fix rather than folded into unrelated work.

**Still open questions, not omissions.** What the difficulty number means and
what it is for. What David actually needs from the export — he has never been
asked whether the seventeen columns are the right seventeen. Whether `template`
deserves to be its own thing rather than a jig value that reads as a word.

---

## 5. Things on the list that are not built

So they are not mistaken for oversights: structured problem and delay codes
alongside free-text issues (you agreed to these — they were lost from the plan
for a while and have been put back); `/track status`, `/track list` and
`/history`; a pause-too-long nudge; a morning resume reminder; stuck-job recovery
and any supervisor path; more than one person on a job; putting a job on hold;
and the dashboard.
