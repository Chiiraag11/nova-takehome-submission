# My solution

This submission uses a stateful, document-first harness. It keeps extracted facts and feedback on disk so later
batches can revise earlier conclusions without changing the command-line contract.

## How to run it

Use Python 3.11+. The repository works with the provided `make` targets. In this environment I used an existing
Python environment containing `pymupdf`, `openpyxl`, `jsonschema` and `pyyaml`; a normal checkout should use:

```text
make setup
make grade PY=.venv/bin/python MODE=both
make sheet PY=.venv/bin/python
make eval PY=.venv/bin/python
make validate PY=.venv/bin/python
```

No API key is required by the submitted implementation. The three grader-facing commands are implemented directly in
`harness/harness.py`. `make sheet` writes `out/exam_report.json` and `answer_sheet.json`; `make grade` writes the two
practice report-card files; `make eval` writes `evals/results.json`.

## What works

The public practice grader finishes at **99.9/100** with all five hard gates clear: 22/22 findings detected, no false
positives, exact recovery, 11/11 adaptation checks, 3/3 feedback checks, and all clean controls left alone.

The exam sheet is dated **2031-02-28** and reports **$8,327.09** total savings: **$3,847.70** overpayment,
**$4,469.10** cashback and **$10.29** late-fee risk. It also reports one trusted-source conflict and one
`cannot_determine` item for the missing `Annex A rev 2`.

The harness supports the supplied invoice layout and a second freight-bill layout, scans via a local OCR fallback,
tracking-driven pickup dates, INR/FX conversion, amendments and rate notices, authorized waivers, duplicate resends and
re-bills, late-fee extensions, quarter rebates, annual improvement credits, expiry alerts, autonomous-action gating,
and feedback-driven rate/authority changes.

## What doesn't

This is intentionally a dependency-light deterministic implementation rather than an LLM agent. Very novel document
layouts that cannot be recovered by the generic line parser may need another parser adapter. Margin step-down terms are
not as deeply generalized as the other cashback paths, and the OCR fallback is designed for single-page invoice scans
rather than complex multi-page image documents. Those are the first areas I would expand against the hidden mailbox.

## The eval case that caught a bug

`evals/cases/own_scan_invoice_id.yaml` caught a real OCR regression: the scan could extract only `CDI` from the printed
invoice number `CDI-3031-0306`, which would silently change the finding key. I fixed this by using an invoice-like
attachment filename to restore a clearly matching OCR prefix. The final case passes with the full key and the exact
$85.46 finding. The conflict cases also caught a rate-notice parser mismatch between `3.03 USD` and `USD 4.38`; the
rate parser now accepts both placements.
