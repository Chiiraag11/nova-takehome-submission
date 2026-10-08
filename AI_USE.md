# AI use

## Tools

I used **OpenAI ChatGPT (GPT-5.6 Luna)** as the coding and review assistant, together with the local container/Python
runtime for file inspection and verification. I used `PyMuPDF`/the repository extractor for PDF text and page evidence,
plus the supplied grading scripts. I did not use an external model API or place any API key in the repository.

## How I worked with them

I asked the assistant to inspect the take-home archive, identify the deliverables, reason about the contract/rate/waiver/
late-fee/cashback rules, and help implement and review the stateful parser. I then ran the repository's own grader after
each material change. The assistant initially missed two generalization details: OCR could truncate the scanned invoice id,
and rate notices used different currency placements. Both were found by regression evals and fixed. The public practice
card was re-run after those fixes and remained 99.9/100 with every hard gate clear.

## Written without AI

The final validation outputs, test execution, and numeric checks were produced by local scripts. The repository narrative
was assembled with AI assistance and then checked against the actual files and grader output before packaging.

## Prompt log

1. **Task prompt:** inspect the uploaded Nova take-home archive and produce the output it asks for.
2. **Brief-review request:** identify the submission requirements, grading gates, document rules and required artifacts.
3. **Implementation request:** inspect the existing harness and fix the parser so practice and exam outputs follow the
   glossary, including conflicts, missing annexes, waivers, duplicates, late fees, cashback and expiry.
4. **Debug request:** investigate why the scanned exam invoice produced a truncated id and why a rate notice parser failed
   across two currency/number formats.
5. **Verification request:** run the public grader, build the exam sheet, create at least ten independent eval cases,
   and run `make validate` before packaging.

Secrets and API keys were never included in this log or repository.
