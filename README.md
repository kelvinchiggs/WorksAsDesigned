# WorksAsDesigned
If it passed validation, produced objects (not strings), and wrote a log file, then it works as designed... 

---

# ⚠️ Disclaimer

This repository contains scripts, tooling, and experiments that **do exactly what they are told to do** — which may or may not be what was *intended*.

All content is provided **as-is**, without warranty of any kind, express or implied, including (but not limited to) warranties of correctness, fitness for purpose, or compatibility with environments that were “just quickly tweaked in production.”

---

## Important Notes

* Scripts may:

  * Modify configuration
  * Query sensitive systems
  * Produce logs that are *uncomfortably honest*
* No guarantees are made that:

  * The script will fix the problem
  * The problem was the real problem
  * The environment was sane to begin with

If a script completes successfully and the result is unexpected, this is **not a bug**, it is evidence.

---

## Usage Guidance

* Review all scripts before execution
* Test in non-production environments first
* Ensure appropriate permissions, change approval, and caffeine intake
* Understand that “it worked on my tenant” is not a transferable property

Running anything from this repository implies acceptance that **automation enforces intent, not hope**.

---

## Responsibility

Responsibility for outcomes lies with:

* The individual executing the script
* The assumptions made before execution
* The environment in which “nothing changed recently”

Responsibility does **not** lie with:

* The author
* The repository
* PowerShell (it did exactly what it was asked)

---

## Final Reminder

If the script:

* Returned objects ✔
* Logged actions ✔
* Did not throw terminating errors ✔

Then it **WorksAsDesigned™**.

Proceed accordingly.
