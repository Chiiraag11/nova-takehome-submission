# AI use

## Tools

- **OpenAI ChatGPT (GPT-5.6 Luna):** coding and review assistant for the harness, the parser fixes and the eval cases.
- **Claude (Anthropic):** drafting `SOLUTION.md`, `DESIGN.md` and this file from my code and command output.
- Local Python, PyMuPDF and the supplied grader and check scripts for verification.

No model API is called by the submitted code and no API key is in the repository.

## How I worked with them

I had the assistant read the take-home archive, list the deliverables, and reason through the contract, rate, waiver,
late-fee and cashback rules. It helped implement and review the stateful parser. After every material change I ran the
repository's grader and my evals myself. The assistant got two things wrong at first: OCR truncated the scanned invoice
id `CDI-3031-0306` to `CDI`, and the rate-notice parser handled `3.03 USD` but not `USD 4.38`. My regression evals
caught both, and the practice card stayed at 99.9/100 after the fixes. I also corrected a confidence description in
`DESIGN.md` that overstated what the code does.

## Written without AI

The grader runs, eval runs, exam sheet and validation results come from local scripts, and I checked the figures in the
write-ups against that output. The write-ups themselves were drafted with AI help and then checked by me.

## Prompt log

1. **Task:** inspect the Nova take-home archive and produce what it asks for.
2. **Brief review:** identify the submission requirements, grading gates, document rules and required files.
3. **Implementation:** fix the parser so practice and exam outputs follow the glossary: conflicts, missing annexes,
   waivers, duplicates, late fees, cashback and expiry.
4. **Debugging:** find why the scanned exam invoice id was truncated and why a rate notice failed in one currency format.
5. **Verification:** run the public grader, build the exam sheet, write at least ten own eval cases, run `make validate`.
6. **Write-ups:** draft `SOLUTION.md` in the template's four sections from my code and outputs, then trim it to what is
   required and relevant.
7. **Design and AI use:** tighten `DESIGN.md` and `AI_USE.md` to their templates.

No secrets or API keys appear in this log.
