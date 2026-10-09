# Design

## Fact model

Each email and attachment becomes a source record keyed by mailbox path and attachment name. PDFs keep their page text (OCR text for image-only scans, using local Tesseract OCR); spreadsheets keep their rows. From these I derive invoices, contracts, amendments,
tracking events, AP ledger rows, waivers, rate notices, late-fee extensions, credit notes and feedback facts. Every
derived fact keeps its source document and page, so each finding cites exact evidence. Contract terms are keyed by
contract id, rates by lane, weight band and validity period, operational facts by AWB, and AP facts by invoice id.

## Conflicts

For each shipment I pick the contract in force on its pickup date, then apply signed amendments, authorized rate notices
and feedback in the order the precedence rules give. If two trusted sources still apply at once and the contract says
neither wins, I emit `conflict` instead of choosing. The amount range is the billed amount minus each of the two
possible correct amounts, floored at zero. A missing named annex gives `cannot_determine`, but only
when that document could change the answer.

## Confidence

Confidence is a fixed value set by the type of outcome, not a computed score. A finding with complete evidence gets 0.95,
a conflict 0.90, and a `cannot_determine` 0.88 to 0.90. I do not currently
lower it for OCR or ambiguous extraction, and agreement between sources does not raise it. I checked it on the practice
set: no reported item was wrong and the Brier score was 0.0031. That mostly shows the set is clean, since AUROC could not
be measured without a wrong item. These numbers are untested on messier data.

## People in the loop

Supervised mode only does internal actions itself. Outbound disputes, credit claims and document requests are drafts,
conflicts are escalated, and a finding of $5,000 or more on one key goes to `human_review`. Autonomous mode may send an
undoable outbound remedy only when the evidence is complete, the key has no conflict, the combined outbound amount is
under $5,000, the terms relied on are authoritative, and the recipient comes from Schedule 1. Expiry alerts and payment
reminders stay internal in both modes.

## Feedback

Feedback is stored as a separate record. A standing `rate_accepted` fact is scoped to its contract, lane and valid dates.
`sender_authority` grants only the named category. An exact-key decision can close, reject or override one finding.
Precedence feedback for a conflict applies only to its contract, lane and dates. Findings changed by a human decision
carry `learned_at`. One-off decisions, such as a waiver for one shipment, are never promoted to a standing rule.

## New documents

Each report re-reads the stored state and recomputes everything. A back-dated amendment therefore changes pricing for
earlier pickup dates. A new rate notice can create a conflict or replace an older authorized rate. A new scan, tracking
row, credit note or AP payment can remove an earlier finding. Findings keep a stable id, and a changed result links to
the one it replaces through `supersedes`.

## Cost

There are no model calls, so model cost is **$0 per 1,000 documents**. The work is local PDF, XLSX and email parsing plus
local OCR for scans. The practice set replays in seconds per mode; I have not timed a 1,000-document run, and scanned
pages are what would dominate it. If I added a vision or LLM fallback I would send only low-confidence pages or
ambiguous clauses to it.

## What I'd do next

With another week I would build real confidence scoring (lower for OCR, higher for agreeing sources) and test it on
messy data, add a margin step-down parser, group multi-page OCR, and add layout-independent table extraction. I would
also add property-based tests for currency, rounding, credit-note netting and overlapping amendment and notice windows.
I am least sure about unseen invoice layouts and about the scan id repair, which relies on the attachment filename.
