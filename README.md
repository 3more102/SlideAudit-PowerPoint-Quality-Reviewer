# SlideAudit — PowerPoint Quality Reviewer

A free, explainable **Streamlit web app** for automated `.pptx` presentation review.
Upload one or more PowerPoints to receive an **overall percentage**, six-category breakdown, **per-slide grades**, prioritized edit notes, and downloadable **JSON/Markdown reports** plus a **copy of the PPTX with editable feedback in speaker notes**.

## Run locally

```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
# macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

## Publish online with GitHub + Streamlit

1. Create a repository named `SlideAudit-PowerPoint-Quality-Reviewer` in your GitHub account.
2. Upload the **contents of this project directory** to the repository root (so `app.py` and `requirements.txt` are at top level).
3. Open [Streamlit Community Cloud](https://share.streamlit.io/), create an app from the repository, choose branch `main` and entry point `app.py`, and deploy.
4. Access it from your Streamlit app URL. GitHub repositories do not automatically host Streamlit apps.

[Official deployment documentation](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/deploy)

## Checks and point allocation

| Category | Maximum points | Examples |
|---|---:|---|
| Typography | 20 | explicitly set small fonts |
| Content density | 20 | word count, number of lines |
| Layout | 20 | off-slide objects, text-box overlaps |
| Visual hierarchy | 15 | missing title, title/body sizing, dense text-only slides |
| Consistency | 15 | too many fonts, inconsistent title sizes across slides |
| Contrast | 10 | contrast ratio when **explicit solid RGB** text and background are inspectable |
| **Total** | **100** | **mean score across slides** |

Each flagged issue has a published deduction. Per-category penalties are capped at the assigned maximum. Every slide starts with 100; the final deck score is the rounded arithmetic mean of slide scores. Source for all thresholds: [`reviewer.py`](reviewer.py). **This is a heuristic rubric authored for SlideAudit, not a validated industry standard or calibrated human rating.**

## Notes embedded into a PPTX copy

Use **Download PPTX with speaker notes** for an optional copy of the original with edit instructions appended to slide speaker notes. Keep the original: complex or unsupported PowerPoint features may not survive a `python-pptx` rewrite unchanged. The original uploaded bytes are never overwritten.

## What is and isn't checked

**Measurable**: location/size of slide shapes, text, declared fonts, text length, simple overlaps, explicitly known RGB contrast.

**Not reliably measurable without further work**: inherited theme font size/color; effective image/background contrast; animation; diagrams' conceptual correctness; artistic quality; technical accuracy; accessibility of chart data; storytelling; audience relevance. Complex slide design may trigger false positives. Review flagged areas visually before editing.

Font-size and text-density thresholds are practical editorial *defaults*, not universal specifications. A data-dense scientific presentation may intentionally violate them. Contrast is based on WCAG-type luminance ratios when explicitly available; the app doesn't claim complete accessibility conformance.

## Privacy and deployment

There is no login, persistent database, third-party AI API, or deliberate saving of uploaded presentations in the code. A third-party hosting provider **does process files on the server**; private/confidential material should instead be evaluated locally or on approved infrastructure. Upload safety limits are 50 MB compressed, 180 MB expanded, and 7,000 archive members per `.pptx`.

## Tests

```bash
pip install pytest
pytest -q
```

Tests generate disposable PPTX decks with `python-pptx` and verify predictable checks and score calculations.

## Roadmap

- Optional actual slide rendering (requires controlled LibreOffice service) for visual/overflow confirmation.
- Human-in-the-loop content/argument evaluation, clearly separated from deterministic rubric.
- More accurate layout-aware reading order and table-specific text rules.
- Improved format preservation and ability to insert genuine review comments in PPTX.

## Project structure

```text
app.py                  Streamlit user interface
reviewer.py             Heuristic evaluation engine and report exporters
requirements.txt        Deploy/runtime requirements
tests/test_reviewer.py  Regression tests
.streamlit/config.toml   User interface theme and upload limits
```

License: MIT. See `LICENSE`.
