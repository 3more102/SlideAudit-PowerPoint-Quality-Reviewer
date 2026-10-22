"""SlideAudit Streamlit UI. Run: streamlit run app.py"""
from __future__ import annotations
import pandas as pd
import streamlit as st
from reviewer import review_pptx, add_review_notes, WEIGHTS

st.set_page_config(page_title="SlideAudit | PowerPoint Quality Reviewer", page_icon="📊", layout="wide")
st.markdown("""
<style>
.block-container{max-width:1250px;padding-top:2rem}
.hero{padding:1.65rem 2rem;border-radius:17px;background:linear-gradient(115deg,#0c1b35,#243f65);color:white;margin-bottom:1.4rem}
.hero h1{font-size:2.35rem;font-weight:800;color:#fff;margin:0 0 .3rem}
.hero p{color:#d6e4f5;margin:0;font-size:1.08rem}
.smalltext{font-size:0.85rem;color:#6d7790}
</style>""", unsafe_allow_html=True)
st.markdown('<div class="hero"><h1>SlideAudit</h1><p>PowerPoint quality scoring & slide-by-slide edit recommendations.</p></div>', unsafe_allow_html=True)
st.caption("Transparent, rule-based heuristic review — not a substitute for human editorial or design judgment.")
files = st.file_uploader("Upload PowerPoint presentation(s)", type=["pptx"], accept_multiple_files=True,
                         help="PPTX only, up to 50 MB per file. Slides are processed in memory; no database is configured.")
with st.expander("How the percentage is calculated"):
    st.write("Each slide begins with 100 available points. Findings subtract published penalties in six categories, capped by category weight. The deck score is the rounded mean across slides. The checklist is configurable in reviewer.py.")
    st.table(pd.DataFrame([{"Category": name, "Maximum points": weight} for name, weight in WEIGHTS.items()]))
    st.info("Automatic inspection does not understand the audience, whether the content is true, or whether a design is beautiful. Unknown theme colors and inherited font sizes are not assumed.")
if not files:
    st.info("Upload a .pptx to view the quality score, category breakdown, and editing notes for every slide.")
    st.stop()

all_reports = []
for file in files:
    try:
        report = review_pptx(file.getvalue(), file.name)
        all_reports.append(report)
    except ValueError as exc:
        st.error(f"{file.name}: {exc}")
    except Exception:
        st.error(f"{file.name}: Could not evaluate this deck. Please re-export it as .pptx and try again.")
if not all_reports:
    st.stop()

if len(all_reports) > 1:
    st.subheader("Presentation comparison")
    st.dataframe(pd.DataFrame([{"File": r.name, "Quality score (%)": r.score, "Slides": r.slide_count,
                                "Flagged edits": sum(len(s.findings) for s in r.slides)} for r in all_reports]),
                 use_container_width=True, hide_index=True)

for report in all_reports:
    st.divider()
    st.subheader(report.name)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Overall quality", f"{report.score}%")
    c2.metric("Slides", report.slide_count)
    c3.metric("Flagged edits", sum(len(s.findings) for s in report.slides))
    c4.metric("Slides ≥80%", sum(s.score >= 80 for s in report.slides))
    st.caption("Percentages indicate how this deck performed against the published measurable checks, not a definitive design grade.")
    cat_frame = pd.DataFrame([{"Category": cat, "Earned": report.category_scores[cat], "Possible": weight,
                               "Percent": round(100*report.category_scores[cat]/weight)} for cat, weight in WEIGHTS.items()])
    left, right = st.columns([1.1, 1.6])
    with left:
        st.markdown("**Category breakdown**")
        st.dataframe(cat_frame[["Category", "Earned", "Possible"]], hide_index=True, use_container_width=True)
    with right:
        st.markdown("**Scores by slide**")
        chart_df = pd.DataFrame([{"Slide": f"{s.number}: {s.title[:24]}", "Score": s.score} for s in report.slides]).set_index("Slide")
        st.bar_chart(chart_df, y="Score", horizontal=True, color="#367ccc")
    st.markdown("#### Edit these first")
    for note in report.executive_notes:
        st.markdown(f"- {note}")
    st.markdown("#### Detailed feedback")
    only_issues = st.checkbox("Only slides with findings", value=False, key=f"issues_{report.name}")
    severity_filter = st.selectbox("Minimum severity", ["All", "Medium + High", "High only"], key=f"severity_{report.name}")
    filtered = []
    for slide in report.slides:
        findings = [f for f in slide.findings if severity_filter == "All" or
                    (severity_filter == "Medium + High" and f.severity in ("Medium", "High")) or
                    (severity_filter == "High only" and f.severity == "High")]
        if only_issues and not findings:
            continue
        filtered.append((slide, findings))
    for slide, findings in filtered:
        title = f"Slide {slide.number} · {slide.title} · {slide.score}% · {len(findings)} edits"
        with st.expander(title, expanded=(slide.score < 70)):
            st.caption(f"{slide.metrics['word_count']} words · {slide.metrics['text_objects']} text objects · {slide.metrics['visual_objects']} visual objects · {slide.metrics['contrast_runs_checked']} explicit contrast checks")
            if not findings:
                st.success("No findings at the selected severity level.")
            for f in findings:
                st.markdown(f"**{f.severity} — {f.headline}** · {f.category}")
                st.write(f"**Evidence:** {f.evidence}")
                st.write(f"**Edit:** {f.fix}")
                st.divider()
    st.markdown("#### Export results")
    d1, d2, d3 = st.columns(3)
    base = report.name.rsplit(".", 1)[0]
    d1.download_button("Download full JSON report", report.to_json(), f"{base}_slideaudit.json", "application/json", key=f"j_{report.name}")
    d2.download_button("Download editable Markdown notes", report.to_markdown(), f"{base}_edit_notes.md", "text/markdown", key=f"m_{report.name}")
    with d3:
        source = next((f.getvalue() for f in files if f.name == report.name), None)
        if source is not None:
            try:
                annotated = add_review_notes(source, report)
                st.download_button("Download PPTX with speaker notes", annotated, f"{base}_review_notes.pptx",
                                   "application/vnd.openxmlformats-officedocument.presentationml.presentation", key=f"p_{report.name}",
                                   help="Creates a new PowerPoint with edit recommendations in each slide's speaker notes. Keep the original.")
            except Exception:
                st.warning("Could not create the optional annotated PPTX; use the Markdown report instead.")
    with st.expander("Limitations & required human review"):
        for limitation in report.limitations:
            st.write("• " + limitation)
