# Design

## Fact model

Each ingested email and attachment becomes a source record keyed by mailbox path and attachment name. PDFs retain page
text (or OCR text for image-only scans); spreadsheets retain sheet rows. Derived records are invoices, contracts,
amendments, tracking events, AP ledger rows, waivers, rate notices, extensions, credit notes and feedback facts. Every
derived fact keeps its source document and page so a report can cite the exact evidence. Contract terms are keyed by
contract id; rates by lane, weight band and validity period; operational facts by AWB; AP facts by invoice id.

## Conflicts

For a shipment I first select the contract in force on its pickup date, then apply signed amendments, authorized rate
notices and feedback according to the precedence rules. If two trusted sources remain simultaneously applicable and the
contract says neither wins, I emit `conflict` rather than choosing. The amount range is the billed amount minus the two
possible correct amounts, floored at zero. Missing named rate annexes produce `cannot_determine` only when the missing
document could change the answer.

## Confidence

Confidence starts from evidence completeness and source authority, then is reduced conceptually for OCR or ambiguous
field extraction. Signed contract/amendment evidence gets the highest score; authorized human correspondence and
tracking/AP records are next. Agreement between independent sources raises confidence. Current deterministic findings
with complete evidence use 0.95; conflicts/cannot-determine cases use lower values because the unresolved issue is
material. Practice was my calibration set: the final run had no wrong reported item and a Brier score of 0.0031.

## People in the loop

Supervised mode performs only internal actions itself. Outbound disputes, credit claims and document requests are drafts;
conflicts escalate; large findings (at least $5,000 on a key) use `human_review`. Autonomous mode may execute only an
undoable outbound remedy when evidence is complete, the key is conflict-free, the combined outbound amount is below
$5,000, the relied-on terms are authoritative, and the recipient comes from Schedule 1. Expiry and payment reminders
remain internal in either mode.

## Feedback

Feedback is persisted as a separate, auditable record. Standing `rate_accepted` facts are scoped by contract, lane and
valid dates; `sender_authority` can grant only the named category; exact-key decisions can close, reject or override one
finding. Conflict-precedence feedback is applied only to its contract/lane/date scope. Affected findings carry
`learned_at`, so later reports show which human decision changed the result. One-off decisions are never promoted to a
standing rule.

## New documents

Every batch is re-read from the persisted state and rules are recomputed. A backdated amendment therefore changes the
pricing calculation for earlier pickup dates instead of being treated as merely a future rule. A new rate notice can
create a conflict or replace an older authorized rate. A newly arriving scan, tracking row, credit note or AP payment can
remove a prior finding. The report keeps the stable finding id and can link a changed result through `supersedes`.

## Cost

The submitted path makes no external model calls, so its direct model cost is **$0 per 1,000 documents**. Work is local
PDF/XLSX/email parsing plus optional local OCR. On this container, the public practice set is small enough to replay in
seconds per mode; the main variable is the number of scanned pages. If a future version adds a vision/LLM fallback, I
would route only low-confidence pages or ambiguous clauses to it, rather than sending every document.

## What I'd do next

I would add an explicit fact-version layer for every learned term, implement a richer margin-step-down parser, and add
multi-page OCR grouping plus layout-independent table extraction. I would also add property-based tests around currency,
rounding, credit-note netting and overlapping amendment/notice windows.
