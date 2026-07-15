# English Dictation HTML Template Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Add an HTML visual template and deterministic DOCX generator to the English dictation skill, then generate and visually verify a worksheet from the supplied image.

**Architecture:** Bundle the adapted HTML under `assets/` as the visual contract. A Python generator reads contract tokens from the HTML and creates an A4 five-column DOCX containing only Chinese prompts and uniform writing lines.

**Tech Stack:** HTML/CSS, Python 3.12, python-docx, unittest, LibreOffice render QA.

---

### Task 1: Add failing generator contract tests

**Files:**
- Create: `tests/test_create_dictation_doc.py`

**Steps:**
1. Test that the bundled HTML declares A4 portrait, five columns, and the English title.
2. Test input validation, stable order, and output rules.
3. Run the test file and verify it fails because the template and generator do not exist.

### Task 2: Add the English template and DOCX generator

**Files:**
- Create: `assets/template.html`
- Create: `scripts/create_dictation_doc.py`
- Modify: `SKILL.md`

**Steps:**
1. Adapt the approved Chinese template to the English dictation form.
2. Implement template-contract loading, prompt validation, and five-column DOCX generation.
3. Update the skill workflow to require the bundled template and script.
4. Run all tests and verify they pass.

### Task 3: Generate and verify the supplied worksheet

**Files:**
- Create: `fixtures/2026-07-15-weekdays-frequency.json`
- Create: `/Users/steve/School/英语/英语默写-星期与频率副词.docx`

**Steps:**
1. Encode the 12 source entries in image order, keeping English only in the answer field for validation.
2. Generate the DOCX with the bundled Python runtime.
3. Render it with the canonical document renderer.
4. Inspect every page and correct any layout defects.
5. Run structural checks confirming all Chinese prompts and no English answers are present.
